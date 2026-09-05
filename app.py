import os
import io
import shutil
import secrets
from datetime import datetime, timedelta
from calendar import month_abbr
from functools import wraps

from flask import Flask, render_template, redirect, url_for, request, flash, jsonify, abort, send_from_directory, send_file
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import pandas as pd
import plotly
import plotly.graph_objects as go
import json

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from models import (db, User, School, Section, Student, Assessment, GroupScreeningTest,
                     InterventionMaterial, MaterialAssignment, InterventionRecord,
                     SystemSetting, PasswordResetToken, PassageTotalPreset, AudioLog, AuditLog)
import ml_utils

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(BASE_DIR, "instance", "reading_ready.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db.init_app(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Please log in to continue."


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ---------------------------------------------------------------------------
# Role-based access helper
# ---------------------------------------------------------------------------
def roles_required(*roles):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if not current_user.is_authenticated:
                return login_manager.unauthorized()
            if current_user.role not in roles:
                abort(403)
            return f(*args, **kwargs)
        return wrapped
    return decorator


READING_LEVEL_COLORS = {
    "Independent": "#2e7d32",
    "Instructional": "#f9a825",
    "Frustration": "#c62828",
}

READING_LEVEL_HEX_SOFT = {
    "Independent": "E4F0EC",
    "Instructional": "FBF1DE",
    "Frustration": "FBE9E6",
}

ALLOWED_MATERIAL_EXTENSIONS = {"pdf", "doc", "docx", "ppt", "pptx", "jpg", "jpeg", "png", "mp3", "wav", "zip", "txt"}
DEPED_EMAIL_DOMAIN = "@deped.gov.ph"

def _is_valid_deped_email(email):
    return email.lower().strip().endswith(DEPED_EMAIL_DOMAIN)
PASSWORD_RESET_TOKEN_EXPIRY_MINUTES = 60

# ---------------------------------------------------------------------------
# System Settings: basic key-value configuration
# ---------------------------------------------------------------------------
SETTINGS_DEFINITIONS = [
    {"key": "current_school_year", "label": "Current School Year", "default": "2025-2026",
     "help": "Used as the default school year when teachers encode a new GST or assessment."},
    {"key": "system_display_name", "label": "System Display Name", "default": "Reading-Ready",
     "help": "Cosmetic label shown to Developers on this settings page (does not change the header/logo)."},
    {"key": "support_contact_email", "label": "Support Contact Email", "default": "",
     "help": "Shown to users who need help with an account approval issue. Leave blank to hide."},
]


def _get_setting(key, default=""):
    row = SystemSetting.query.get(key)
    return row.value if row and row.value is not None else default


def _get_all_settings():
    return {d["key"]: _get_setting(d["key"], d["default"]) for d in SETTINGS_DEFINITIONS}


def _current_school_year():
    return _get_setting("current_school_year", "2025-2026")

DEFAULT_PASSAGE_TOTALS = {
    "word_recognition_total": 100, "literal_total": 5, "inferential_total": 5,
    "critical_total": 5, "gst_total_items": 20,
}
ASSESSMENT_TERMS = ["gst", "pre", "mid", "post"]
TERM_LABELS = {"gst": "Group Screening Test", "pre": "Pre-test", "mid": "Mid-year", "post": "Post-test"}


def _get_passage_preset(school_id, grade_level, term="gst"):
    """Returns the remembered totals for this school+grade+stage, falling
    back to system defaults for anything not set yet."""
    preset = PassageTotalPreset.query.filter_by(school_id=school_id, grade_level=grade_level, term=term).first()
    result = dict(DEFAULT_PASSAGE_TOTALS)
    if preset:
        for key in result:
            val = getattr(preset, key)
            if val is not None:
                result[key] = val
    return result


def _save_passage_preset(school_id, grade_level, term, totals, user_id):
    """Upserts totals for one school+grade+stage. This is the ONLY place
    totals are ever written -- encode forms never call this, by design."""
    preset = PassageTotalPreset.query.filter_by(school_id=school_id, grade_level=grade_level, term=term).first()
    if not preset:
        preset = PassageTotalPreset(school_id=school_id, grade_level=grade_level, term=term)
        db.session.add(preset)
    for key, val in totals.items():
        if val is not None:
            setattr(preset, key, val)
    preset.updated_by = user_id
    db.session.commit()


def _assessment_totals_for_teacher(school_id, grade_level):
    """All four stage presets (GST + Pre/Mid/Post) for the teacher's OWN
    section grade only -- deliberately not all grades 4-6, since a teacher
    shouldn't be managing totals for grades their own section doesn't use."""
    return {term: _get_passage_preset(school_id, grade_level, term) for term in ASSESSMENT_TERMS}


def _generate_student_account(student):
    """
    Auto-provisions a learner portal login the moment a student is added,
    per the school's standard convention:
      username = first name, lowercased (deduplicated with a numeric
                 suffix if another student already has it)
      password = last name (lowercased, spaces stripped) + the last three
                 digits of the LRN
    Returns (username, password) so the caller can show it to the teacher
    once, or None if this student already has a portal account.
    """
    if User.query.filter_by(student_id=student.id).first():
        return None

    base_username = "".join(student.first_name.lower().split()) or "student"
    username = base_username
    suffix = 1
    while User.query.filter_by(username=username).first():
        suffix += 1
        username = f"{base_username}{suffix}"

    digits_only = "".join(ch for ch in student.lrn if ch.isdigit())
    last_three = digits_only[-3:] if len(digits_only) >= 3 else digits_only.zfill(3)
    password = f"{student.last_name.lower().replace(' ', '')}{last_three}"

    account = User(
        username=username, full_name=student.full_name, role="student",
        school_id=student.section.school_id if student.section else None,
        student_id=student.id,
    )
    account.set_password(password)
    db.session.add(account)
    db.session.flush()
    return username, password

def _looks_like_sqlite_file(file_storage):
    """
    SQLite database files begin with a fixed 16-byte header. Checking it is
    a cheap, reliable way to reject non-database uploads before we ever
    consider swapping them in for the live app database.
    """
    header = file_storage.read(16)
    file_storage.seek(0)
    return header == b"SQLite format 3\x00"


# ---------------------------------------------------------------------------
# Auth + Landing
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    schools = School.query.order_by(School.name).all()
    return render_template("landing.html", schools=schools, login_username="")

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    login_identifier = ""
    selected_role = request.form.get("selected_role", "").strip() if request.method == "POST" else ""

    if request.method == "POST":
        identifier = request.form.get("identifier", "").strip()
        password = request.form.get("password", "")
        login_identifier = identifier
        user = User.query.filter(
            (User.username == identifier) | (User.email == identifier)
        ).first()

        role_display = {
            "developer": "an IT Administrator", "coordinator": "a Reading Coordinator",
            "teacher": "a Teacher", "student": "a Learner",
        }

        if user and selected_role and user.role != selected_role:
            flash(f"This account is registered as {role_display.get(user.role, user.role.title())}. "
                  "Please go back and select the correct role.", "error")
            _log_audit("login_failed",
                        f"Failed login attempt for '{identifier}' -- role mismatch "
                        f"(selected {selected_role}, account is {user.role}).")
        elif user and user.check_password(password):
            if user.status == "pending":
                flash("This account is still awaiting approval. You'll be able to log in "
                      "once your Reading Coordinator or system admin reviews it.", "error")
                _log_audit("login_blocked", f"Login blocked for '{user.username}' -- account still pending approval.", user=user)
            elif user.status == "rejected":
                flash("This account request was not approved. Contact your Reading "
                      "Coordinator or the system admin for details.", "error")
                _log_audit("login_blocked", f"Login blocked for '{user.username}' -- account was rejected.", user=user)
            elif user.status == "deactivated":
                flash("This account has been deactivated. Contact your Reading "
                      "Coordinator or the system admin if you believe this is a mistake.", "error")
                _log_audit("login_blocked", f"Login blocked for '{user.username}' -- account deactivated.", user=user)
            else:
                _log_audit("login_success", f"{user.full_name} logged in.", user=user)
                login_user(user)
                return redirect(url_for("dashboard"))
        else:
            flash("Invalid username/email or password.", "error")
            _log_audit("login_failed", f"Failed login attempt for '{identifier}'.")

    schools = School.query.order_by(School.name).all()
    return render_template("landing.html", schools=schools, login_username=login_identifier,
                            selected_role=selected_role)

@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    reset_link = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        user = User.query.filter_by(username=username).first()

        if email and not _is_valid_deped_email(email):
            errors.append("Email must be an official DepEd email address ending in @deped.gov.ph.")
        # Always show the same generic confirmation regardless of whether the
        # username/email actually matched -- this route must not reveal
        # whether a given username or email exists in the system.
        if user and user.email and user.email.lower() == email.lower():
            token = secrets.token_urlsafe(32)
            db.session.add(PasswordResetToken(
                user_id=user.id, token=token,
                expires_at=datetime.utcnow() + timedelta(minutes=PASSWORD_RESET_TOKEN_EXPIRY_MINUTES),
            ))
            db.session.commit()
            _log_audit("password_reset_requested", f"Password reset requested for '{user.username}'.", user=user)
            reset_link = url_for("reset_password", token=token, _external=True)

        flash("If that username and email match an account on file, a reset link has been generated below.", "success")

    return render_template("forgot_password.html", reset_link=reset_link)


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    reset_token = PasswordResetToken.query.filter_by(token=token).first()
    valid = bool(reset_token) and not reset_token.used and reset_token.expires_at > datetime.utcnow()

    if not valid:
        flash("This password reset link is invalid or has expired. Please request a new one.", "error")
        return redirect(url_for("forgot_password"))

    if request.method == "POST":
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if len(new_password) < 8:
            flash("New password must be at least 8 characters.", "error")
            return render_template("reset_password.html", token=token)
        if new_password != confirm_password:
            flash("Passwords do not match.", "error")
            return render_template("reset_password.html", token=token)

        user = User.query.get(reset_token.user_id)
        user.set_password(new_password)
        reset_token.used = True
        db.session.commit()
        _log_audit("password_reset_completed", f"{user.full_name} reset their password via a reset link.", user=user)
        flash("Password reset successfully. Please log in with your new password.", "success")
        return redirect(url_for("index"))

    return render_template("reset_password.html", token=token)

@app.route("/signup", methods=["GET", "POST"])
def signup():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    schools = School.query.order_by(School.name).all()

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        full_name = request.form.get("full_name", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        role = request.form.get("role", "")
        school_choice = request.form.get("school_id", "")
        new_school_name = request.form.get("new_school_name", "").strip()

        errors = []
        if not all([username, email, full_name, password, role, school_choice]):
            errors.append("All fields are required.")
        if request.form.get("agree_terms") != "on":
            errors.append("You must agree to the Terms and Conditions to create an account.")
        if role not in ("coordinator", "teacher"):
            errors.append("Please select a valid role.")
        if password != confirm_password:
            errors.append("Passwords do not match.")
        if len(password) < 8:
            errors.append("Password must be at least 8 characters.")
        if User.query.filter_by(username=username).first():
            errors.append("That username is already taken.")
        if email and User.query.filter_by(email=email).first():
            errors.append("An account with that email already exists.")

        school = None
        pending_school_name = None

        if school_choice == "other":
            # Requesting a brand-new school -- only a Reading Coordinator can
            # do this, since a Teacher account is meaningless without a
            # coordinator to approve it at that school.
            if role != "coordinator":
                errors.append("Only a Reading Coordinator can request a new school to be added. "
                               "Teachers must select an existing school that already has an "
                               "approved Reading Coordinator.")
            elif not new_school_name:
                errors.append("Please type your school's name so it can be reviewed by the system admin.")
            else:
                existing = School.query.filter(School.name.ilike(new_school_name)).first()
                if existing:
                    # Already exists (maybe under a slightly different typed
                    # name) -- link to it directly instead of creating a
                    # duplicate school entry.
                    school = existing
                else:
                    pending_school_name = new_school_name
        else:
            try:
                school = School.query.get(int(school_choice))
            except (TypeError, ValueError):
                school = None
            if not school:
                errors.append("Please select a valid school.")
            elif role == "teacher":
                has_coordinator = User.query.filter_by(
                    role="coordinator", school_id=school.id, status="approved"
                ).first()
                if not has_coordinator:
                    errors.append(
                        f"{school.name} doesn't have an approved Reading Coordinator yet, so teacher "
                        "accounts can't be created for it right now. Please check back later or "
                        "contact your school's Reading Coordinator."
                    )

        if errors:
            for e in errors:
                flash(e, "error")
            return render_template("signup.html", schools=schools, form=request.form)

        user = User(
            username=username, email=email, full_name=full_name, role=role,
            school_id=school.id if school else None,
            pending_school_name=pending_school_name,
            status="pending",
        )
        user.set_password(password)
        db.session.add(user)
        db.session.commit()

        school_desc = school.name if school else f"'{pending_school_name}' (new school, pending admin review)"
        _log_audit("signup_submitted",
                    f"{full_name} ({username}) requested a {role} account for {school_desc}.",
                    user=user, school_id=school.id if school else None)

        if pending_school_name:
            flash("Your account request has been submitted. Since your school isn't in our system "
                  "yet, a system admin will review your school and account together before you "
                  "can log in.", "success")
        else:
            flash("Your account request has been submitted. A system admin will review it "
                  "before you can log in.", "success")
        return redirect(url_for("index"))

    return render_template("signup.html", schools=schools, form={})


@app.route("/logout")
@login_required
def logout():
    _log_audit("logout", f"{current_user.full_name} logged out.", user=current_user)
    logout_user()
    flash("Logged out successfully.", "success")
    return redirect(url_for("index"))


# ---------------------------------------------------------------------------
# Profile & Password — shared across every role
# ---------------------------------------------------------------------------
@app.route("/profile")
@login_required
def my_profile():
    return render_template("profile.html", user=current_user)


@app.route("/profile/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not current_user.check_password(current_password):
            flash("Current password is incorrect.", "error")
            return redirect(url_for("change_password"))
        if len(new_password) < 8:
            flash("New password must be at least 8 characters.", "error")
            return redirect(url_for("change_password"))
        if new_password != confirm_password:
            flash("New passwords do not match.", "error")
            return redirect(url_for("change_password"))

        current_user.set_password(new_password)
        db.session.commit()
        _log_audit("password_changed", f"{current_user.full_name} changed their password.", user=current_user)
        flash("Password updated successfully.", "success")
        return redirect(url_for("my_profile"))

    return render_template("change_password.html")


# ---------------------------------------------------------------------------
# Dashboard router (per role)
# ---------------------------------------------------------------------------
@app.route("/dashboard")
@login_required
def dashboard():
    if current_user.role == "developer":
        return redirect(url_for("developer_dashboard"))
    elif current_user.role == "coordinator":
        return redirect(url_for("coordinator_dashboard"))
    elif current_user.role == "teacher":
        return redirect(url_for("teacher_dashboard"))
    elif current_user.role == "student":
        return redirect(url_for("student_dashboard"))
    abort(403)


# ---------------------------------------------------------------------------
# (a) IT Capstone Developer: technical admin
# ---------------------------------------------------------------------------
@app.route("/admin")
@roles_required("developer")
def developer_dashboard():
    schools = School.query.all()
    coordinators = (User.query.filter(User.role == "coordinator",
                                       User.status.in_(["approved", "deactivated"]))
                     .order_by(User.full_name).all())
    pending_coordinators = (User.query.filter_by(role="coordinator", status="pending")
                             .order_by(User.requested_at.asc()).all())
    pending_teachers = (User.query.filter_by(role="teacher", status="pending")
                         .order_by(User.requested_at.asc()).all())
    stats = {
        "schools": len(schools),
        "coordinators": User.query.filter_by(role="coordinator", status="approved").count(),
        "teachers": User.query.filter_by(role="teacher", status="approved").count(),
        "students": Student.query.count(),
        "assessments": Assessment.query.count(),
        "pending_coordinators": len(pending_coordinators),
        "pending_teachers": len(pending_teachers),
    }

    # Dashboard is view-only: latest 5 of each, with a "View All" link to
    # the dedicated management page where the actual create-account /
    # add-school forms now live.
    latest_coordinators = (User.query.filter_by(role="coordinator")
                            .order_by(User.created_at.desc()).limit(5).all())
    latest_schools = School.query.order_by(School.created_at.desc()).limit(5).all()
    recent_activity = AuditLog.query.order_by(AuditLog.created_at.desc()).limit(5).all()

    role_filter = request.args.get("chart_role", "")
    start_param = request.args.get("chart_start", "")
    end_param = request.args.get("chart_end", "")
    login_chart, chart_start, chart_end = _build_login_activity_chart(role_filter, start_param, end_param)

    return render_template("developer_dashboard.html", schools=schools, coordinators=coordinators,
                            pending_coordinators=pending_coordinators, pending_teachers=pending_teachers,
                            stats=stats, recent_activity=recent_activity,
                            latest_coordinators=latest_coordinators, latest_schools=latest_schools,
                            login_chart=login_chart, chart_start=chart_start, chart_end=chart_end,
                            chart_role=role_filter)


@app.route("/admin/coordinators")
@roles_required("developer")
def all_coordinators():
    coordinators = (User.query.filter_by(role="coordinator")
                     .order_by(User.created_at.desc()).all())
    schools = School.query.order_by(School.name).all()
    return render_template("all_coordinators.html", coordinators=coordinators, schools=schools)


@app.route("/admin/schools")
@roles_required("developer")
def all_schools():
    schools = School.query.order_by(School.created_at.desc()).all()
    return render_template("all_schools.html", schools=schools)

AUDIT_ACTION_LABELS = {
    "login_success": "Login", "login_failed": "Failed Login", "login_blocked": "Login Blocked",
    "logout": "Logout", "signup_submitted": "Signup Submitted",
    "account_approved": "Account Approved", "account_rejected": "Account Rejected",
    "account_status_toggled": "Status Changed", "account_edited": "Account Edited",
    "password_changed": "Password Changed", "report_exported": "Report Exported",
    "settings_updated": "Settings Updated", "database_backed_up": "Database Backed Up",
    "database_restored": "Database Restored",
    "password_reset_requested": "Password Reset Requested",
    "password_reset_completed": "Password Reset Completed",
        "student_added": "Learner Added",
    "gst_encoded": "GST Encoded",
    "assessment_encoded": "Assessment Encoded",
    "assessment_edited": "Assessment Edited",
    "assessment_deleted": "Assessment Deleted",
    "material_assigned": "Material Assigned",
    "external_resource_added": "External Resource Added",
    "external_resource_edited": "External Resource Edited",
    "external_resource_deleted": "External Resource Deleted",
}

@app.route("/admin/teachers")
@roles_required("developer")
def all_teachers():
    teachers = User.query.filter_by(role="teacher").order_by(User.created_at.desc()).all()
    # A teacher owns at most one section in this system, so a single pass
    # over sections gives us "which grade does this teacher teach" without
    # an N+1 query per row.
    sections_by_teacher = {sec.teacher_id: sec for sec in Section.query.filter(Section.teacher_id.isnot(None)).all()}
    schools = School.query.order_by(School.name).all()
    return render_template("all_teachers.html", teachers=teachers,
                            sections_by_teacher=sections_by_teacher, schools=schools)


@app.route("/admin/students")
@roles_required("developer")
def all_students_admin():
    students = Student.query.all()
    students = sorted(students, key=lambda s: (s.last_name.lower(), s.first_name.lower()))
    student_ids = [s.id for s in students]
    latest_by_student = {a.student_id: a for a in _latest_assessment_per_student(student_ids)}
    schools = School.query.order_by(School.name).all()
    return render_template("all_students_admin.html", students=students,
                            latest_by_student=latest_by_student, schools=schools)


@app.route("/admin/assessments")
@roles_required("developer")
def all_assessments_admin():
    term_filter = request.args.get("term", "")
    level_filter = request.args.get("level", "")
    query = Assessment.query
    if term_filter:
        query = query.filter_by(term=term_filter)
    assessments = query.order_by(Assessment.date_administered.desc()).limit(500).all()
    if level_filter:
        assessments = [a for a in assessments if a.reading_level == level_filter]
    return render_template("all_assessments_admin.html", assessments=assessments,
                            term_filter=term_filter, level_filter=level_filter)

@app.route("/admin/audit-logs")
@roles_required("developer")
def audit_logs():
    action_filter = request.args.get("action", "")
    search = request.args.get("q", "").strip()

    query = AuditLog.query
    if action_filter:
        query = query.filter_by(action=action_filter)
    if search:
        query = query.filter(AuditLog.description.ilike(f"%{search}%"))

    logs = query.order_by(AuditLog.created_at.desc()).limit(300).all()
    return render_template("audit_logs.html", logs=logs, action_labels=AUDIT_ACTION_LABELS,
                            available_actions=sorted(AUDIT_ACTION_LABELS.items(), key=lambda kv: kv[1]),
                            action_filter=action_filter, search=search)


@app.route("/admin/coordinator/<int:user_id>/edit", methods=["GET", "POST"])
@roles_required("developer")
def edit_coordinator(user_id):
    coordinator = User.query.get_or_404(user_id)
    if coordinator.role != "coordinator":
        abort(404)
    schools = School.query.order_by(School.name).all()

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()
        school_id = request.form.get("school_id")

        if not all([full_name, email, school_id]):
            flash("All fields are required.", "error")
            return render_template("edit_coordinator.html", coordinator=coordinator, schools=schools)
        existing = User.query.filter(User.email == email, User.id != coordinator.id).first()
        if existing:
            flash("Another account already uses that email.", "error")
            return render_template("edit_coordinator.html", coordinator=coordinator, schools=schools)
        if not _is_valid_deped_email(email):
            flash("Email must be an official DepEd email address ending in @deped.gov.ph.", "error")
            return render_template("edit_coordinator.html", coordinator=coordinator, schools=schools)

        coordinator.full_name = full_name
        coordinator.email = email
        coordinator.school_id = int(school_id)
        db.session.commit()
        _log_audit("account_edited", f"{current_user.full_name} edited coordinator account '{coordinator.username}'.",
                    user=current_user, school_id=coordinator.school_id)
        flash(f"Updated {coordinator.full_name}'s account.", "success")
        return redirect(url_for("developer_dashboard"))

    return render_template("edit_coordinator.html", coordinator=coordinator, schools=schools)


@app.route("/admin/coordinator/<int:user_id>/toggle-status", methods=["POST"])
@roles_required("developer")
def toggle_coordinator_status(user_id):
    coordinator = User.query.get_or_404(user_id)
    if coordinator.role != "coordinator":
        abort(404)
    if coordinator.status == "approved":
        coordinator.status = "deactivated"
        flash(f"Deactivated {coordinator.full_name}'s account. They can no longer log in.", "success")
        log_msg = f"{current_user.full_name} deactivated coordinator account '{coordinator.username}'."
    elif coordinator.status == "deactivated":
        coordinator.status = "approved"
        flash(f"Reactivated {coordinator.full_name}'s account.", "success")
        log_msg = f"{current_user.full_name} reactivated coordinator account '{coordinator.username}'."
    else:
        flash("This account isn't in an active or deactivated state yet.", "error")
        return redirect(url_for("developer_dashboard"))
    db.session.commit()
    _log_audit("account_status_toggled", log_msg, user=current_user, school_id=coordinator.school_id)
    return redirect(url_for("developer_dashboard"))


@app.route("/admin/approve/<int:user_id>", methods=["POST"])
@roles_required("developer")
def approve_account(user_id):
    user = User.query.get_or_404(user_id)
    if user.role != "coordinator":
        flash("Teacher account requests are reviewed by their school's Reading Coordinator, "
              "not the system admin.", "error")
        return redirect(url_for("developer_dashboard"))
    if user.status != "pending":
        flash("This request has already been reviewed.", "error")
        return redirect(url_for("developer_dashboard"))

    # New-school request: the School row doesn't exist yet -- create it now,
    # at approval time, so a request that gets rejected never leaves an
    # orphan school behind in the database.
    if not user.school_id and user.pending_school_name:
        new_school = School(name=user.pending_school_name)
        db.session.add(new_school)
        db.session.flush()
        user.school_id = new_school.id
        user.pending_school_name = None
        flash(f"Created new school '{new_school.name}' and linked it to this account.", "success")

    user.status = "approved"
    user.reviewed_at = datetime.utcnow()
    user.reviewed_by = current_user.id
    db.session.commit()
    _log_audit("account_approved", f"{current_user.full_name} approved coordinator account '{user.username}'.",
                user=current_user, school_id=user.school_id)
    flash(f"Approved {user.full_name} ({user.role}) — they can now log in.", "success")
    return redirect(url_for("developer_dashboard"))


@app.route("/admin/reject/<int:user_id>", methods=["POST"])
@roles_required("developer")
def reject_account(user_id):
    user = User.query.get_or_404(user_id)
    if user.role != "coordinator":
        flash("Teacher account requests are reviewed by their school's Reading Coordinator, "
              "not the system admin.", "error")
        return redirect(url_for("developer_dashboard"))
    if user.status != "pending":
        flash("This request has already been reviewed.", "error")
        return redirect(url_for("developer_dashboard"))
    user.status = "rejected"
    user.reviewed_at = datetime.utcnow()
    user.reviewed_by = current_user.id
    db.session.commit()
    _log_audit("account_rejected", f"{current_user.full_name} rejected coordinator account '{user.username}'.",
                user=current_user, school_id=user.school_id)
    flash(f"Rejected the request from {user.full_name}.", "success")
    return redirect(url_for("developer_dashboard"))


@app.route("/admin/create-coordinator", methods=["POST"])
@roles_required("developer")
def create_coordinator():
    username = request.form.get("username", "").strip()
    full_name = request.form.get("full_name", "").strip()
    email = request.form.get("email", "").strip()
    password = request.form.get("password", "")
    school_id = request.form.get("school_id")

    if not all([username, full_name, email, password, school_id]):
        flash("All fields are required.", "error")
        return redirect(url_for("developer_dashboard"))
    if User.query.filter_by(username=username).first():
        flash("Username already exists.", "error")
        return redirect(url_for("developer_dashboard"))
    if User.query.filter_by(email=email).first():
        flash("An account with that email already exists.", "error")
        return redirect(url_for("developer_dashboard"))
    if not _is_valid_deped_email(email):
        flash("Email must be an official DepEd email address ending in @deped.gov.ph.", "error")
        return redirect(url_for("all_coordinators"))

    u = User(username=username, full_name=full_name, email=email, role="coordinator", school_id=int(school_id))
    u.set_password(password)
    db.session.add(u)
    db.session.commit()
    flash(f"Reading Coordinator account '{username}' created. No email server is configured for this "
          f"deployment, so please share the login details with {full_name} directly — in production "
          f"this would be emailed automatically to {email}.", "success")
    return redirect(url_for("all_coordinators"))


@app.route("/admin/create-school", methods=["POST"])
@roles_required("developer")
def create_school():
    name = request.form.get("name", "").strip()
    if name:
        db.session.add(School(name=name))
        db.session.commit()
        flash(f"School '{name}' added.", "success")
    return redirect(url_for("all_schools"))


# ---------------------------------------------------------------------------
# System Settings: configuration + database backup/restore
# ---------------------------------------------------------------------------
@app.route("/admin/settings", methods=["GET", "POST"])
@roles_required("developer")
def system_settings():
    if request.method == "POST":
        for definition in SETTINGS_DEFINITIONS:
            key = definition["key"]
            value = request.form.get(key, "").strip()
            setting = SystemSetting.query.get(key)
            if setting:
                setting.value = value
                setting.updated_at = datetime.utcnow()
                setting.updated_by = current_user.id
            else:
                db.session.add(SystemSetting(key=key, value=value, updated_by=current_user.id))
        db.session.commit()
        _log_audit("settings_updated", f"{current_user.full_name} updated system settings.", user=current_user)
        flash("System settings updated.", "success")
        return redirect(url_for("system_settings"))

    values = _get_all_settings()
    db_path = os.path.join(BASE_DIR, "instance", "reading_ready.db")
    db_size_kb = round(os.path.getsize(db_path) / 1024, 1) if os.path.exists(db_path) else 0
    backups_dir = os.path.join(BASE_DIR, "instance", "backups")
    backups = sorted(os.listdir(backups_dir), reverse=True)[:10] if os.path.isdir(backups_dir) else []

    return render_template("system_settings.html", settings=SETTINGS_DEFINITIONS, values=values,
                            db_size_kb=db_size_kb, backups=backups)


@app.route("/admin/settings/backup")
@roles_required("developer")
def backup_database():
    db_path = os.path.join(BASE_DIR, "instance", "reading_ready.db")
    if not os.path.exists(db_path):
        flash("No database file found to back up.", "error")
        return redirect(url_for("system_settings"))

    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    filename = f"reading_ready_backup_{timestamp}.db"
    _log_audit("database_backed_up", f"{current_user.full_name} downloaded a database backup.", user=current_user)
    return send_file(db_path, mimetype="application/x-sqlite3", as_attachment=True, download_name=filename)


@app.route("/admin/settings/restore", methods=["POST"])
@roles_required("developer")
def restore_database():
    upload = request.files.get("backup_file")
    if not upload or not upload.filename:
        flash("Select a database file to restore.", "error")
        return redirect(url_for("system_settings"))
    if not upload.filename.lower().endswith((".db", ".sqlite", ".sqlite3")):
        flash("That doesn't look like a SQLite database file (expected .db / .sqlite / .sqlite3).", "error")
        return redirect(url_for("system_settings"))
    if not _looks_like_sqlite_file(upload):
        flash("That file doesn't have a valid SQLite header -- restore cancelled to protect the live database.", "error")
        return redirect(url_for("system_settings"))

    db_path = os.path.join(BASE_DIR, "instance", "reading_ready.db")
    backups_dir = os.path.join(BASE_DIR, "instance", "backups")
    os.makedirs(backups_dir, exist_ok=True)

    # Always back up the CURRENT database before overwriting it, so a bad
    # restore can be undone by re-uploading this backup afterward.
    if os.path.exists(db_path):
        timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(db_path, os.path.join(backups_dir, f"pre_restore_{timestamp}.db"))

    # Release SQLAlchemy's open connections before swapping the file out
    # from under it. Acceptable for a small school-scale SQLite deployment;
    # not something you'd do against a live production database.
    db.session.remove()
    db.engine.dispose()

    upload.save(db_path)

    # Best-effort log -- this now writes to the RESTORED database, and the
    # actor's own account may not exist in it if it's an older backup. The
    # _log_audit helper already swallows failures so this never blocks the
    # restore itself from completing.
    _log_audit("database_restored",
                f"{current_user.full_name} restored the database from an uploaded backup.",
                user=current_user)

    # The current session's user may no longer exist (or may exist with
    # different data) in the restored database -- force a clean re-login
    # rather than continue on a session that might not match reality.
    logout_user()
    flash("Database restored successfully. The previous database was saved automatically as a "
          "backup. Please log in again.", "success")
    return redirect(url_for("index"))


# ---------------------------------------------------------------------------
# (b) Reading Coordinator: school-wide monitoring
# ---------------------------------------------------------------------------
@app.route("/coordinator")
@roles_required("coordinator")
def coordinator_dashboard():
    school = current_user.school
    sections = Section.query.filter_by(school_id=school.id).order_by(Section.grade_level, Section.name).all()
    section_ids = [s.id for s in sections]
    students = Student.query.filter(Student.section_id.in_(section_ids)).all() if section_ids else []
    students = sorted(students, key=lambda s: (s.sex != "M", s.last_name.lower(), s.first_name.lower()))
    student_ids = [s.id for s in students]

    latest_assessments = _latest_assessment_per_student(student_ids)
    latest_assessment_by_student = {a.student_id: a for a in latest_assessments}

    level_counts = {"Independent": 0, "Instructional": 0, "Frustration": 0}
    for a in latest_assessments:
        if a.reading_level in level_counts:
            level_counts[a.reading_level] += 1

    fig = go.Figure(data=[go.Pie(
        labels=list(level_counts.keys()),
        values=list(level_counts.values()),
        marker=dict(colors=[READING_LEVEL_COLORS[k] for k in level_counts.keys()]),
        hole=0.45,
    )])
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=320)
    level_chart = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    heatmap_html = _build_heatmap(sections, latest_assessments)

    teachers = (User.query.filter(User.role == "teacher", User.school_id == school.id,
                                   User.status.in_(["approved", "deactivated"]))
                .order_by(User.created_at.desc()).all())
    pending_teachers_count = User.query.filter_by(role="teacher", school_id=school.id, status="pending").count()

    at_risk = []
    for stu in students:
        risk = _predict_student_risk(stu.assessments)
        if risk and risk.get("at_risk"):
            latest = stu.assessments[-1] if stu.assessments else None
            at_risk.append({
                "student": stu,
                "reading_level": latest.reading_level if latest else "Not assessed",
                "predicted_next_composite": risk["predicted_next_composite"],
                "trend": risk["trend"],
            })
    at_risk.sort(key=lambda r: r["predicted_next_composite"])
    at_risk = at_risk[:8]

    sex_counts = {"M": sum(1 for s in students if s.sex == "M"),
                  "F": sum(1 for s in students if s.sex == "F")}

    return render_template(
        "coordinator_dashboard.html",
        school=school, sections=sections, students=students,
        latest_assessment_by_student=latest_assessment_by_student, sex_counts=sex_counts,
        level_counts=level_counts, level_chart=level_chart,
        heatmap_html=heatmap_html, teachers=teachers, pending_teachers_count=pending_teachers_count,
        at_risk_students=at_risk,
        active_teacher_count=sum(1 for t in teachers if t.status == "approved"),
    )

@app.route("/coordinator/students")
@roles_required("coordinator")
def coordinator_students():
    school = current_user.school
    sections = Section.query.filter_by(school_id=school.id).all()
    section_ids = [s.id for s in sections]
    students = Student.query.filter(Student.section_id.in_(section_ids)).all() if section_ids else []
    students = sorted(students, key=lambda s: (s.sex != "M", s.last_name.lower(), s.first_name.lower()))
    student_ids = [s.id for s in students]
    latest_assessment_by_student = {a.student_id: a for a in _latest_assessment_per_student(student_ids)}
    sex_counts = {"M": sum(1 for s in students if s.sex == "M"),
                  "F": sum(1 for s in students if s.sex == "F")}
    return render_template("coordinator_students.html", school=school, sections=sections,
                            students=students, latest_assessment_by_student=latest_assessment_by_student,
                            sex_counts=sex_counts, level_param=request.args.get("level", ""))

@app.route("/coordinator/teacher/<int:user_id>/edit", methods=["GET", "POST"])
@roles_required("coordinator", "developer")
def edit_teacher(user_id):
    teacher = User.query.get_or_404(user_id)
    if teacher.role != "teacher":
        abort(404)
    if current_user.role == "coordinator" and teacher.school_id != current_user.school_id:
        abort(403)

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()

        if not all([full_name, email]):
            flash("All fields are required.", "error")
            return render_template("edit_teacher.html", teacher=teacher)
        existing = User.query.filter(User.email == email, User.id != teacher.id).first()
        if existing:
            flash("Another account already uses that email.", "error")
            return render_template("edit_teacher.html", teacher=teacher)
        if not _is_valid_deped_email(email):
            flash("Email must be an official DepEd email address ending in @deped.gov.ph.", "error")
            return render_template("edit_teacher.html", teacher=teacher)

        teacher.full_name = full_name
        teacher.email = email
        db.session.commit()
        _log_audit("account_edited", f"{current_user.full_name} edited teacher account '{teacher.username}'.",
                    user=current_user, school_id=teacher.school_id)
        flash(f"Updated {teacher.full_name}'s account.", "success")
        return redirect(url_for("all_teachers") if current_user.role == "developer" else url_for("coordinator_teachers"))

    return render_template("edit_teacher.html", teacher=teacher)


@app.route("/coordinator/teacher/<int:user_id>/toggle-status", methods=["POST"])
@roles_required("coordinator", "developer")
def toggle_teacher_status(user_id):
    teacher = User.query.get_or_404(user_id)
    if teacher.role != "teacher":
        abort(404)
    if current_user.role == "coordinator" and teacher.school_id != current_user.school_id:
        abort(403)

    redirect_target = url_for("all_teachers") if current_user.role == "developer" else url_for("coordinator_teachers")

    if teacher.status == "approved":
        teacher.status = "deactivated"
        flash(f"Deactivated {teacher.full_name}'s account. They can no longer log in.", "success")
        log_msg = f"{current_user.full_name} deactivated teacher account '{teacher.username}'."
    elif teacher.status == "deactivated":
        teacher.status = "approved"
        flash(f"Reactivated {teacher.full_name}'s account.", "success")
        log_msg = f"{current_user.full_name} reactivated teacher account '{teacher.username}'."
    else:
        flash("This account isn't in an active or deactivated state yet.", "error")
        return redirect(redirect_target)

    db.session.commit()
    _log_audit("account_status_toggled", log_msg, user=current_user, school_id=teacher.school_id)
    return redirect(redirect_target)

@app.route("/coordinator/approve-teacher/<int:user_id>", methods=["POST"])
@roles_required("coordinator")
def approve_teacher_account(user_id):
    user = User.query.get_or_404(user_id)
    if user.role != "teacher" or user.school_id != current_user.school_id:
        abort(403)
    if user.status != "pending":
        flash("This request has already been reviewed.", "error")
        return redirect(url_for("coordinator_dashboard"))
    user.status = "approved"
    user.reviewed_at = datetime.utcnow()
    user.reviewed_by = current_user.id
    db.session.commit()
    _log_audit("account_approved", f"{current_user.full_name} approved teacher account '{user.username}'.",
                user=current_user, school_id=user.school_id)
    flash(f"Approved {user.full_name} — they can now log in as a Teacher.", "success")
    return redirect(url_for("coordinator_dashboard"))


@app.route("/coordinator/reject-teacher/<int:user_id>", methods=["POST"])
@roles_required("coordinator")
def reject_teacher_account(user_id):
    user = User.query.get_or_404(user_id)
    if user.role != "teacher" or user.school_id != current_user.school_id:
        abort(403)
    if user.status != "pending":
        flash("This request has already been reviewed.", "error")
        return redirect(url_for("coordinator_dashboard"))
    user.status = "rejected"
    user.reviewed_at = datetime.utcnow()
    user.reviewed_by = current_user.id
    db.session.commit()
    _log_audit("account_rejected", f"{current_user.full_name} rejected teacher account '{user.username}'.",
                user=current_user, school_id=user.school_id)
    flash(f"Rejected the request from {user.full_name}.", "success")
    return redirect(url_for("coordinator_dashboard"))


@app.route("/coordinator/create-teacher", methods=["POST"])
@roles_required("coordinator")
def create_teacher():
    username = request.form.get("username", "").strip()
    full_name = request.form.get("full_name", "").strip()
    email = request.form.get("email", "").strip()
    password = request.form.get("password", "")
    section_id = request.form.get("section_id")

    if not all([username, full_name, email, password]):
        flash("All fields are required.", "error")
        return redirect(url_for("coordinator_teachers"))
    if User.query.filter_by(username=username).first():
        flash("Username already exists.", "error")
        return redirect(url_for("coordinator_teachers"))
    if User.query.filter_by(email=email).first():
        flash("An account with that email already exists.", "error")
        return redirect(url_for("coordinator_teachers"))
    if not _is_valid_deped_email(email):
        flash("Email must be an official DepEd email address ending in @deped.gov.ph.", "error")
        return redirect(url_for("coordinator_teachers"))

    u = User(username=username, full_name=full_name, email=email, role="teacher", school_id=current_user.school_id)
    u.set_password(password)
    db.session.add(u)
    db.session.flush()
    if section_id:
        section = Section.query.get(int(section_id))
        if section and section.school_id == current_user.school_id:
            section.teacher_id = u.id
    db.session.commit()
    flash(f"Teacher account '{username}' created. No email server is configured for this deployment, "
          f"so please share the login details with {full_name} directly — in production this would be "
          f"emailed automatically to {email}.", "success")
    return redirect(url_for("coordinator_dashboard"))

@app.route("/coordinator/sections")
@roles_required("coordinator")
def coordinator_sections():
    sections = (Section.query.filter_by(school_id=current_user.school_id)
                .order_by(Section.grade_level, Section.name).all())
    return render_template("coordinator_sections.html", sections=sections)


@app.route("/coordinator/teachers")
@roles_required("coordinator")
def coordinator_teachers():
    teachers = (User.query.filter(User.role == "teacher", User.school_id == current_user.school_id,
                                   User.status.in_(["approved", "deactivated"]))
                .order_by(User.created_at.desc()).all())
    pending_teachers = (User.query.filter_by(role="teacher", school_id=current_user.school_id, status="pending")
                         .order_by(User.requested_at.asc()).all())
    sections = (Section.query.filter_by(school_id=current_user.school_id)
                .order_by(Section.grade_level, Section.name).all())
    sections_by_teacher = {sec.teacher_id: sec for sec in sections if sec.teacher_id}
    return render_template("coordinator_teachers.html", teachers=teachers, pending_teachers=pending_teachers,
                            sections=sections, sections_by_teacher=sections_by_teacher)

@app.route("/coordinator/create-section", methods=["POST"])
@roles_required("coordinator")
def create_section():
    name = request.form.get("name", "").strip()
    grade_level = request.form.get("grade_level")
    if name and grade_level:
        db.session.add(Section(name=name, grade_level=int(grade_level), school_id=current_user.school_id))
        db.session.commit()
        flash(f"Section '{name}' added.", "success")
    return redirect(url_for("coordinator_sections"))


def _coordinator_report_rows(school):
    """
    Shared data assembly for the coordinator's Consolidated Literacy Report --
    used by the on-screen HTML view and both the PDF and Excel exports, so
    all three always show identical numbers.
    """
    sections = Section.query.filter_by(school_id=school.id).all()
    rows = []
    for sec in sections:
        for stu in sec.students:
            latest = stu.assessments[-1] if stu.assessments else None
            rows.append({
                "section": sec.name,
                "student": stu.full_name,
                "lrn": stu.lrn,
                "term": latest.term if latest else "-",
                "word_recognition": latest.word_recognition_accuracy if latest else None,
                "comprehension_avg": latest.comprehension_avg if latest else None,
                "reading_speed_wpm": latest.reading_speed_wpm if latest else None,
                "reading_level": latest.reading_level if latest else "Not yet assessed",
            })
    return rows


@app.route("/coordinator/report")
@roles_required("coordinator")
def coordinator_report():
    """Consolidated school-wide report suitable for division submission."""
    rows = _coordinator_report_rows(current_user.school)
    return render_template("coordinator_report.html", rows=rows, school=current_user.school)


@app.route("/coordinator/report/export.pdf")
@roles_required("coordinator")
def coordinator_report_export_pdf():
    school = current_user.school
    rows = _coordinator_report_rows(school)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter),
                             topMargin=40, bottomMargin=40, leftMargin=40, rightMargin=40)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(school.name, styles["Title"]))
    story.append(Paragraph("Consolidated Literacy Report &mdash; latest recorded assessment per learner", styles["Normal"]))
    story.append(Paragraph(f"Generated {datetime.utcnow().strftime('%B %d, %Y')}", styles["Normal"]))
    story.append(Spacer(1, 16))

    table_data = [["Section", "Student", "LRN", "Last Term", "Word Recog. (%)",
                    "Comprehension (%)", "Speed (WPM)", "Reading Level"]]
    for r in rows:
        table_data.append([
            r["section"],
            r["student"],
            r["lrn"],
            r["term"].title() if r["term"] != "-" else "-",
            f'{r["word_recognition"]}%' if r["word_recognition"] is not None else "-",
            f'{r["comprehension_avg"]}%' if r["comprehension_avg"] is not None else "-",
            f'{r["reading_speed_wpm"]}' if r["reading_speed_wpm"] is not None else "-",
            r["reading_level"],
        ])

    if len(table_data) == 1:
        table_data.append(["No students on file yet.", "", "", "", "", "", "", ""])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1b2430")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e4ddd0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#faf8f4")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(table)
    doc.build(story)
    buffer.seek(0)

    filename = f"{school.name.replace(' ', '_')}_summary_report.pdf"
    _log_audit("report_exported", f"{current_user.full_name} exported the school summary report as PDF.",
                user=current_user, school_id=current_user.school_id)
    return send_file(buffer, mimetype="application/pdf", as_attachment=True, download_name=filename)


@app.route("/coordinator/report/export.xlsx")
@roles_required("coordinator")
def coordinator_report_export_xlsx():
    school = current_user.school
    rows = _coordinator_report_rows(school)

    wb = Workbook()
    ws = wb.active
    ws.title = "School Summary"

    ws.merge_cells("A1:H1")
    ws["A1"] = school.name
    ws["A1"].font = Font(name="Arial", size=14, bold=True)

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Consolidated Literacy Report -- generated {datetime.utcnow().strftime('%B %d, %Y')}"
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color="52606D")

    headers = ["Section", "Student", "LRN", "Last Term", "Word Recognition (%)",
               "Comprehension (%)", "Reading Speed (WPM)", "Reading Level"]
    header_row = 4
    ws.append([])  # row 3 spacer
    ws.append(headers)

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B2430", end_color="1B2430", fill_type="solid")
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    level_fills = {level: PatternFill(start_color=hex_code, end_color=hex_code, fill_type="solid")
                   for level, hex_code in READING_LEVEL_HEX_SOFT.items()}

    if rows:
        for r in rows:
            row_idx = ws.max_row + 1
            ws.append([
                r["section"], r["student"], r["lrn"],
                r["term"].title() if r["term"] != "-" else "-",
                r["word_recognition"] if r["word_recognition"] is not None else "-",
                r["comprehension_avg"] if r["comprehension_avg"] is not None else "-",
                r["reading_speed_wpm"] if r["reading_speed_wpm"] is not None else "-",
                r["reading_level"],
            ])
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_idx).font = Font(name="Arial", size=10)
            level_fill = level_fills.get(r["reading_level"])
            if level_fill:
                ws.cell(row=row_idx, column=8).fill = level_fill
    else:
        ws.append(["No students on file yet.", "", "", "", "", "", "", ""])
        ws.cell(row=ws.max_row, column=1).font = Font(name="Arial", italic=True, color="52606D")

    for col_idx, header in enumerate(headers, start=1):
        col_letter = ws.cell(row=header_row, column=col_idx).column_letter
        cells_below = [ws.cell(row=r, column=col_idx) for r in range(header_row, ws.max_row + 1)]
        max_len = max((len(str(c.value)) for c in cells_below if c.value is not None), default=len(header))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 42)

    ws.freeze_panes = f"A{header_row + 1}"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"{school.name.replace(' ', '_')}_summary_report.xlsx"
    _log_audit("report_exported", f"{current_user.full_name} exported the school summary report as Excel.",
                user=current_user, school_id=current_user.school_id)
    return send_file(buffer, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      as_attachment=True, download_name=filename)


@app.route("/coordinator/analytics")
@roles_required("coordinator")
def coordinator_analytics():
    """
    School Analytics: grade-level breakdowns, trend-over-time, and
    intervention effectiveness -- all computed from Assessment/
    GroupScreeningTest data that already exists, no new models needed.
    """
    school = current_user.school
    sections = Section.query.filter_by(school_id=school.id).all()
    section_ids = [s.id for s in sections]
    students = Student.query.filter(Student.section_id.in_(section_ids)).all() if section_ids else []

    grade_level_chart = _build_grade_level_chart(students)
    trend_chart = _build_grade_trend_chart(students)
    intervention_chart, intervention_stats = _build_intervention_effectiveness_chart(students)

    return render_template(
        "school_analytics.html", school=school,
        grade_level_chart=grade_level_chart,
        trend_chart=trend_chart,
        intervention_chart=intervention_chart,
        intervention_stats=intervention_stats,
    )


# ---------------------------------------------------------------------------
# (c) Teacher: assessment cycle, encoding, recommendations
# ---------------------------------------------------------------------------
@app.route("/teacher")
@roles_required("teacher")
def teacher_dashboard():
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    students = sorted(section.students, key=lambda s: (s.sex != "M", s.last_name.lower(), s.first_name.lower())) if section else []

    student_ids = [s.id for s in students]
    latest_assessments = {a.student_id: a for a in _latest_assessment_per_student(student_ids)}

    level_counts = {"Independent": 0, "Instructional": 0, "Frustration": 0}
    assessed_count = 0
    for a in latest_assessments.values():
        if a.reading_level in level_counts:
            level_counts[a.reading_level] += 1
            assessed_count += 1

    assessments_by_term = {}
    overall_scores = {}
    for s in students:
        by_term = {a.term: a for a in s.assessments}
        assessments_by_term[s.id] = by_term
        overall_scores[s.id] = {
            term: ml_utils.composite_score(a.word_recognition_accuracy, a.comprehension_avg)
            for term, a in by_term.items()
        }

    # Class Reading Level donut, pre-computed for every (stage x sex)
    # combination so the in-page filter redraws instantly without a reload.
    reading_levels = ["Independent", "Instructional", "Frustration"]
    chart_data = {}
    for scope in ["latest", "pre", "mid", "post"]:
        for sex_key in ["all", "M", "F"]:
            counts = {lvl: 0 for lvl in reading_levels}
            for s in students:
                if sex_key != "all" and s.sex != sex_key:
                    continue
                a = latest_assessments.get(s.id) if scope == "latest" else assessments_by_term.get(s.id, {}).get(scope)
                if a and a.reading_level in counts:
                    counts[a.reading_level] += 1
            chart_data[f"{scope}|{sex_key}"] = counts

    trend_terms = ["pre", "mid", "post"]
    trend_values = []
    for term in trend_terms:
        scores = [overall_scores[s.id][term] for s in students if term in overall_scores.get(s.id, {})]
        trend_values.append(round(sum(scores) / len(scores), 1) if scores else None)

    at_risk = []
    for s in students:
        risk = _predict_student_risk(s.assessments)
        if risk and risk.get("at_risk"):
            latest = s.assessments[-1] if s.assessments else None
            at_risk.append({
                "student": s,
                "reading_level": latest.reading_level if latest else "Not assessed",
                "predicted_next_composite": risk["predicted_next_composite"],
                "trend": risk["trend"],
            })
    at_risk.sort(key=lambda r: r["predicted_next_composite"])

    assessment_totals = _assessment_totals_for_teacher(current_user.school_id, section.grade_level) if section else {}

    activity = (AuditLog.query.filter_by(user_id=current_user.id)
                .order_by(AuditLog.created_at.desc()).limit(10).all())

    return render_template("teacher_dashboard.html", section=section, students=students,
                            level_counts=level_counts, assessed_count=assessed_count,
                            chart_data=chart_data, reading_level_colors=READING_LEVEL_COLORS,
                            trend_terms=trend_terms, trend_values=trend_values,
                            at_risk_students=at_risk, assessment_totals=assessment_totals,
                            term_labels=TERM_LABELS, activity=activity,
                            action_labels=AUDIT_ACTION_LABELS)

@app.route("/teacher/assessment-totals")
@roles_required("teacher")
def manage_assessment_totals():
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section:
        flash("You are not yet assigned to a section.", "error")
        return redirect(url_for("teacher_dashboard"))
    totals = _assessment_totals_for_teacher(current_user.school_id, section.grade_level)
    return render_template("manage_assessment_totals.html", section=section, totals=totals, term_labels=TERM_LABELS)

@app.route("/teacher/assessment-totals/update", methods=["POST"])
@roles_required("teacher")
def update_assessment_totals():
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section:
        flash("You are not yet assigned to a section.", "error")
        return redirect(url_for("teacher_dashboard"))

    term = request.form.get("term")
    if term not in ASSESSMENT_TERMS:
        flash("Invalid assessment stage selected.", "error")
        return redirect(url_for("teacher_dashboard"))

    try:
        if term == "gst":
            totals = {"gst_total_items": int(request.form.get("gst_total_items"))}
        else:
            totals = {
                "word_recognition_total": int(request.form.get("word_recognition_total")),
                "literal_total": int(request.form.get("literal_total")),
                "inferential_total": int(request.form.get("inferential_total")),
                "critical_total": int(request.form.get("critical_total")),
            }
    except (TypeError, ValueError):
        flash("Total items must be whole numbers.", "error")
        return redirect(url_for("teacher_dashboard"))

    if any(v <= 0 for v in totals.values()):
        flash("Total items must be greater than zero.", "error")
        return redirect(url_for("teacher_dashboard"))

    _save_passage_preset(current_user.school_id, section.grade_level, term, totals, current_user.id)
    flash(f"{TERM_LABELS[term]} totals for Grade {section.grade_level} updated.", "success")
    return redirect(url_for("manage_assessment_totals"))

@app.route("/teacher/learners")
@roles_required("teacher")
def teacher_students():
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    students = sorted(section.students, key=lambda s: (s.sex != "M", s.last_name.lower(), s.first_name.lower())) if section else []

    student_ids = [s.id for s in students]
    latest_assessments = {a.student_id: a for a in _latest_assessment_per_student(student_ids)}
    latest_gst = {s.id: s.latest_gst for s in students}

    overall_scores = {}
    for s in students:
        by_term = {a.term: a for a in s.assessments}
        overall_scores[s.id] = {
            term: ml_utils.composite_score(a.word_recognition_accuracy, a.comprehension_avg)
            for term, a in by_term.items()
        }

    sex_counts = {"M": sum(1 for s in students if s.sex == "M"),
                  "F": sum(1 for s in students if s.sex == "F")}

    return render_template("teacher_students.html", section=section, students=students,
                            latest_assessments=latest_assessments, latest_gst=latest_gst,
                            overall_scores=overall_scores, sex_counts=sex_counts)

@app.route("/teacher/gst/<int:student_id>", methods=["GET", "POST"])
@roles_required("teacher")
def encode_gst(student_id):
    student = Student.query.get_or_404(student_id)
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    default_school_year = _current_school_year()
    preset = _get_passage_preset(current_user.school_id, student.grade_level, "gst")

    if request.method == "POST":
        try:
            score = int(request.form.get("score"))
        except (TypeError, ValueError):
            flash("Score must be a whole number.", "error")
            return render_template("encode_gst.html", student=student,
                                    default_school_year=default_school_year, preset=preset)

        total_items = preset["gst_total_items"]
        if score < 0 or score > total_items:
            flash(f"Score must be between 0 and {total_items}.", "error")
            return render_template("encode_gst.html", student=student,
                                    default_school_year=default_school_year, preset=preset)

        gst = GroupScreeningTest(
            student_id=student.id,
            school_year=request.form.get("school_year", "").strip() or default_school_year,
            grade_level_at_test=student.grade_level,
            score=score, total_items=total_items, encoded_by=current_user.id,
        )
        db.session.add(gst)
        db.session.commit()

        _log_audit("gst_encoded",
                    f"{current_user.full_name} recorded a Group Screening Test for {student.full_name}: {gst.placement_label}.",
                    user=current_user, school_id=current_user.school_id)

        flash(
            f"Group Screening Test recorded for {student.full_name}: "
            f"{gst.percentage}% \u2014 {gst.placement_label}. "
            f"Pre-test passage: Grade {gst.placement_grade_level}.",
            "success",
        )
        return redirect(url_for("student_profile", student_id=student.id))

    return render_template("encode_gst.html", student=student,
                            default_school_year=default_school_year, preset=preset)

@app.route("/teacher/encode/<int:student_id>", methods=["GET", "POST"])
@roles_required("teacher")
def encode_assessment(student_id):
    student = Student.query.get_or_404(student_id)
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    gst = student.latest_gst
    default_passage_grade = gst.placement_grade_level if gst else student.grade_level
    default_school_year = _current_school_year()

    # Reading-Ready allows exactly ONE assessment per term (Pre / Mid / Post)
    # per learner. Once a term is recorded, further corrections go through
    # Edit -- never a second "new" record for the same term -- so a
    # learner's assessment history always caps at 3 rows, ever.
    existing_terms = {a.term for a in student.assessments}
    available_terms = [t for t in ["pre", "mid", "post"] if t not in existing_terms]

    roster = sorted(section.students, key=lambda s: (s.last_name.lower(), s.first_name.lower()))
    roster_ids = [s.id for s in roster]
    idx = roster_ids.index(student.id) if student.id in roster_ids else None
    prev_student = roster[idx - 1] if idx is not None and idx > 0 else None
    next_student = roster[idx + 1] if idx is not None and idx < len(roster) - 1 else None

    presets_by_grade = {g: {t: _get_passage_preset(current_user.school_id, g, t) for t in ASSESSMENT_TERMS}
                         for g in range(1, 7)}

    if not available_terms:
        flash(f"{student.full_name} already has all three assessments (Pre, Mid, Post) on file. "
              "Edit or delete an existing one from their profile instead of adding a new one.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    if request.method == "POST":
        term = request.form.get("term")
        if term not in available_terms:
            flash("That assessment stage is already recorded for this learner -- edit the "
                  "existing one from their profile instead.", "error")
            return redirect(url_for("student_profile", student_id=student.id))

        passage_grade_used = int(request.form.get("passage_grade_level") or default_passage_grade)
        # Totals are read from the teacher's own configured preset -- never
        # trusted from the submitted form, so a learner's scoring can't be
        # accidentally (or deliberately) inflated by editing hidden fields.
        preset = _get_passage_preset(current_user.school_id, passage_grade_used, term)

        try:
            fields = {
                "word_recognition_correct": int(request.form.get("word_recognition_correct")),
                "literal_correct": int(request.form.get("literal_correct")),
                "inferential_correct": int(request.form.get("inferential_correct")),
                "critical_correct": int(request.form.get("critical_correct")),
            }
        except (TypeError, ValueError):
            flash("Checks-correct fields must be whole numbers.", "error")
            return render_template("encode_assessment.html", student=student,
                                    default_passage_grade=default_passage_grade, gst=gst,
                                    default_school_year=default_school_year, preset=preset,
                                    presets_json=json.dumps(presets_by_grade),
                                    prev_student=prev_student, next_student=next_student,
                                    available_terms=available_terms)

        totals = {
            "word_recognition_total": preset["word_recognition_total"],
            "literal_total": preset["literal_total"],
            "inferential_total": preset["inferential_total"],
            "critical_total": preset["critical_total"],
        }

        errors = []
        for label, correct_key, total_key in [
            ("Word Recognition", "word_recognition_correct", "word_recognition_total"),
            ("Literal", "literal_correct", "literal_total"),
            ("Inferential", "inferential_correct", "inferential_total"),
            ("Critical", "critical_correct", "critical_total"),
        ]:
            if fields[correct_key] < 0 or fields[correct_key] > totals[total_key]:
                errors.append(f"{label}: checks-correct must be between 0 and {totals[total_key]}.")
        if errors:
            for e in errors:
                flash(e, "error")
            return render_template("encode_assessment.html", student=student,
                                    default_passage_grade=default_passage_grade, gst=gst,
                                    default_school_year=default_school_year, preset=preset,
                                    presets_json=json.dumps(presets_by_grade),
                                    prev_student=prev_student, next_student=next_student,
                                    available_terms=available_terms)

        a = Assessment(
            student_id=student.id,
            school_year=request.form.get("school_year", "").strip() or default_school_year,
            term=term, passage_grade_level=passage_grade_used,
            reading_speed_wpm=float(request.form.get("reading_speed_wpm")),
            encoded_by=current_user.id,
            word_recognition_correct=fields["word_recognition_correct"], word_recognition_total=totals["word_recognition_total"],
            literal_correct=fields["literal_correct"], literal_total=totals["literal_total"],
            inferential_correct=fields["inferential_correct"], inferential_total=totals["inferential_total"],
            critical_correct=fields["critical_correct"], critical_total=totals["critical_total"],
        )
        db.session.add(a)
        db.session.commit()

        _log_audit("assessment_encoded",
                    f"{current_user.full_name} encoded the {a.term.title()}-test for {student.full_name}: {a.reading_level}.",
                    user=current_user, school_id=current_user.school_id)
        
        gap_note = f" Weakest skill: {a.weakest_skill}." if a.weakest_skill else ""
        flash(f"{a.term.title()}-test recorded for {student.full_name}: {a.reading_level} level.{gap_note}",
              "success")

        go_next = request.form.get("go_next") == "1"
        if go_next and next_student:
            return redirect(url_for("encode_assessment", student_id=next_student.id))
        return redirect(url_for("student_profile", student_id=student.id))

    preset = _get_passage_preset(current_user.school_id, default_passage_grade, available_terms[0])
    return render_template("encode_assessment.html", student=student,
                            default_passage_grade=default_passage_grade, gst=gst,
                            default_school_year=default_school_year, preset=preset,
                            presets_json=json.dumps(presets_by_grade),
                            prev_student=prev_student, next_student=next_student,
                            available_terms=available_terms)

@app.route("/teacher/assessment/<int:assessment_id>/edit", methods=["GET", "POST"])
@roles_required("teacher")
def edit_assessment(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    student = assessment.student
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    preset = _get_passage_preset(current_user.school_id,
                                  assessment.passage_grade_level or student.grade_level,
                                  assessment.term)

    if request.method == "POST":
        try:
            passage_grade_used = int(request.form.get("passage_grade_level") or
                                      assessment.passage_grade_level or student.grade_level)
            fields = {
                "word_recognition_correct": int(request.form.get("word_recognition_correct")),
                "literal_correct": int(request.form.get("literal_correct")),
                "inferential_correct": int(request.form.get("inferential_correct")),
                "critical_correct": int(request.form.get("critical_correct")),
            }
            reading_speed_wpm = float(request.form.get("reading_speed_wpm"))
        except (TypeError, ValueError):
            flash("Score fields must be numbers.", "error")
            return render_template("edit_assessment.html", assessment=assessment, student=student, preset=preset)

        preset = _get_passage_preset(current_user.school_id, passage_grade_used, assessment.term)
        totals = {
            "word_recognition_total": preset["word_recognition_total"],
            "literal_total": preset["literal_total"],
            "inferential_total": preset["inferential_total"],
            "critical_total": preset["critical_total"],
        }

        errors = []
        for label, correct_key, total_key in [
            ("Word Recognition", "word_recognition_correct", "word_recognition_total"),
            ("Literal", "literal_correct", "literal_total"),
            ("Inferential", "inferential_correct", "inferential_total"),
            ("Critical", "critical_correct", "critical_total"),
        ]:
            if fields[correct_key] < 0 or fields[correct_key] > totals[total_key]:
                errors.append(f"{label}: checks-correct must be between 0 and {totals[total_key]}.")
        if errors:
            for e in errors:
                flash(e, "error")
            return render_template("edit_assessment.html", assessment=assessment, student=student, preset=preset)

        assessment.passage_grade_level = passage_grade_used
        assessment.reading_speed_wpm = reading_speed_wpm
        assessment.word_recognition_correct = fields["word_recognition_correct"]
        assessment.word_recognition_total = totals["word_recognition_total"]
        assessment.literal_correct = fields["literal_correct"]
        assessment.literal_total = totals["literal_total"]
        assessment.inferential_correct = fields["inferential_correct"]
        assessment.inferential_total = totals["inferential_total"]
        assessment.critical_correct = fields["critical_correct"]
        assessment.critical_total = totals["critical_total"]
        db.session.commit()

        _log_audit("assessment_edited",
                    f"{current_user.full_name} edited the {assessment.term.title()}-test for {student.full_name}.",
                    user=current_user, school_id=current_user.school_id)

        flash(f"{assessment.term.title()}-test updated for {student.full_name}.", "success")
        return redirect(url_for("student_profile", student_id=student.id))

    return render_template("edit_assessment.html", assessment=assessment, student=student, preset=preset)


@app.route("/teacher/assessment/<int:assessment_id>/delete", methods=["POST"])
@roles_required("teacher")
def delete_assessment(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    student = assessment.student
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    term_label = assessment.term.title()
    db.session.delete(assessment)
    db.session.commit()

    _log_audit("assessment_deleted",
                f"{current_user.full_name} deleted the {term_label}-test for {student.full_name}.",
                user=current_user, school_id=current_user.school_id)

    flash(f"{term_label}-test deleted for {student.full_name}. You can re-encode it if needed.", "success")
    return redirect(url_for("student_profile", student_id=student.id))

@app.route("/teacher/wizard", methods=["GET"])
@roles_required("teacher")
def encode_wizard():
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    students = section.students if section else []

    students_data = []
    for s in students:
        gst = s.latest_gst
        latest_assessment = s.assessments[-1] if s.assessments else None
        students_data.append({
            "id": s.id, "lrn": s.lrn, "full_name": s.full_name, "grade_level": s.grade_level,
            "gst_placement_grade": gst.placement_grade_level if gst else s.grade_level,
            "gst_placement_label": gst.placement_label if gst else None,
            "latest_term": latest_assessment.term if latest_assessment else None,
            "latest_reading_level": latest_assessment.reading_level if latest_assessment else None,
        })

    presets_by_grade = {g: {t: _get_passage_preset(current_user.school_id, g, t) for t in ASSESSMENT_TERMS}
                         for g in range(1, 7)}

    return render_template("encode_wizard.html", section=section, students=students,
                            students_json=json.dumps(students_data),
                            presets_json=json.dumps(presets_by_grade),
                            default_school_year=_current_school_year())

@app.route("/teacher/wizard/submit", methods=["POST"])
@roles_required("teacher")
def encode_wizard_submit():
    student_id = request.form.get("student_id")
    if not student_id:
        flash("Select a learner to begin the wizard.", "error")
        return redirect(url_for("encode_wizard"))

    student = Student.query.get_or_404(int(student_id))
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    default_school_year = _current_school_year()
    school_year = request.form.get("school_year", "").strip() or default_school_year

    try:
        passage_grade_level = int(request.form.get("passage_grade_level"))
    except (TypeError, ValueError):
        flash("Passage grade level is required -- please check the Passage step.", "error")
        return redirect(url_for("encode_wizard"))

    record_gst = request.form.get("record_gst") == "on"
    if record_gst:
        gst_preset = _get_passage_preset(current_user.school_id, student.grade_level, "gst")
        gst_total = gst_preset["gst_total_items"]
        try:
            gst_score = int(request.form.get("gst_score"))
        except (TypeError, ValueError):
            flash("GST score must be a whole number.", "error")
            return redirect(url_for("encode_wizard"))
        if gst_score < 0 or gst_score > gst_total:
            flash(f"GST score must be between 0 and {gst_total}.", "error")
            return redirect(url_for("encode_wizard"))

        db.session.add(GroupScreeningTest(
            student_id=student.id, school_year=school_year,
            grade_level_at_test=student.grade_level,
            score=gst_score, total_items=gst_total, encoded_by=current_user.id,
        ))
        db.session.flush()

    term = request.form.get("term", "pre")
    preset = _get_passage_preset(current_user.school_id, passage_grade_level, term)

    try:
        correct = {
            "word_recognition_correct": int(request.form.get("word_recognition_correct")),
            "literal_correct": int(request.form.get("literal_correct")),
            "inferential_correct": int(request.form.get("inferential_correct")),
            "critical_correct": int(request.form.get("critical_correct")),
        }
        reading_speed_wpm = float(request.form.get("reading_speed_wpm"))
    except (TypeError, ValueError):
        flash("Score fields must be numbers -- please check the Score Entry step.", "error")
        return redirect(url_for("encode_wizard"))

    totals = {
        "word_recognition_total": preset["word_recognition_total"],
        "literal_total": preset["literal_total"],
        "inferential_total": preset["inferential_total"],
        "critical_total": preset["critical_total"],
    }

    errors = []
    for label, correct_key, total_key in [
        ("Word Recognition", "word_recognition_correct", "word_recognition_total"),
        ("Literal", "literal_correct", "literal_total"),
        ("Inferential", "inferential_correct", "inferential_total"),
        ("Critical", "critical_correct", "critical_total"),
    ]:
        if correct[correct_key] < 0 or correct[correct_key] > totals[total_key]:
            errors.append(f"{label}: checks-correct must be between 0 and {totals[total_key]}.")
    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("encode_wizard"))

    assessment = Assessment(
        student_id=student.id, school_year=school_year, term=term,
        passage_grade_level=passage_grade_level, reading_speed_wpm=reading_speed_wpm,
        encoded_by=current_user.id,
        word_recognition_correct=correct["word_recognition_correct"], word_recognition_total=totals["word_recognition_total"],
        literal_correct=correct["literal_correct"], literal_total=totals["literal_total"],
        inferential_correct=correct["inferential_correct"], inferential_total=totals["inferential_total"],
        critical_correct=correct["critical_correct"], critical_total=totals["critical_total"],
    )
    db.session.add(assessment)
    db.session.flush()

    audio_file = request.files.get("audio_blob")
    if audio_file and audio_file.filename:
        upload_dir = os.path.join(BASE_DIR, "instance", "audio_logs")
        os.makedirs(upload_dir, exist_ok=True)
        ext = audio_file.filename.rsplit(".", 1)[-1].lower() if "." in audio_file.filename else "webm"
        filename = f"{student.lrn}_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.{ext}"
        audio_file.save(os.path.join(upload_dir, filename))
        db.session.add(AudioLog(
            student_id=student.id, assessment_id=assessment.id,
            filename=filename, notes=request.form.get("audio_notes", "").strip(),
        ))

    db.session.commit()

    gap_note = f" Weakest skill: {assessment.weakest_skill}." if assessment.weakest_skill else ""
    flash(f"{assessment.term.title()}-test recorded for {student.full_name} via guided wizard: "
          f"{assessment.reading_level} level.{gap_note}", "success")
    return redirect(url_for("student_profile", student_id=student.id))

@app.route("/teacher/student/<int:student_id>")
@roles_required("teacher", "coordinator", "developer")
def student_profile(student_id):
    student = Student.query.get_or_404(student_id)
    assessments = student.assessments
    gst = student.latest_gst

    # progress evaluation routing
    routing = _evaluate_progress_routing(student, gst, assessments)

    # trend chart
    trend_chart = _build_trend_chart(assessments)

    # recommended materials from intervention library (live auto-match)
    recommendations = _recommend_materials(student, assessments)

    # actual tracked assignments (Materials Library)
    assignments = (MaterialAssignment.query.filter_by(student_id=student.id)
                   .order_by(MaterialAssignment.assigned_at.desc()).all())

    # risk prediction
    risk = _predict_student_risk(assessments)

    # learner's own portal login, if a teacher has created one
    learner_account = User.query.filter_by(student_id=student.id).first()

    return render_template("student_profile.html", student=student, assessments=assessments,
                            routing=routing, trend_chart=trend_chart,
                            recommendations=recommendations, assignments=assignments,
                            risk=risk, gst=gst, learner_account=learner_account)


@app.route("/teacher/mark-passed/<int:student_id>", methods=["POST"])
@roles_required("teacher")
def mark_passed_early(student_id):
    student = Student.query.get_or_404(student_id)
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    latest = student.assessments[-1] if student.assessments else None
    if not latest:
        flash("This learner has no assessment on file yet -- encode a Pre-test first.", "error")
        return redirect(url_for("student_profile", student_id=student.id))
    if latest.term == "post":
        flash("This learner already has a Post-test on file -- the cycle is already complete.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    latest.teacher_override_pass = True
    latest.override_note = request.form.get("note", "").strip() or None
    db.session.commit()
    flash(f"{student.full_name} marked as passed early at {latest.term.title()}-test.", "success")
    return redirect(url_for("student_profile", student_id=student.id))


@app.route("/teacher/add-student", methods=["POST"])
@roles_required("teacher", "coordinator")
def add_student():
    section = Section.query.filter_by(teacher_id=current_user.id).first() if current_user.role == "teacher" else None
    section_id = request.form.get("section_id") or (section.id if section else None)
    lrn = request.form.get("lrn", "").strip()
    first_name = request.form.get("first_name", "").strip()
    last_name = request.form.get("last_name", "").strip()
    grade_level = request.form.get("grade_level")
    sex = request.form.get("sex", "")

    if not all([lrn, first_name, last_name, grade_level, section_id]):
        flash("All fields are required to add a student.", "error")
        return redirect(request.referrer or url_for("dashboard"))
    if Student.query.filter_by(lrn=lrn).first():
        flash("A student with that LRN already exists.", "error")
        return redirect(request.referrer or url_for("dashboard"))

    stu = Student(lrn=lrn, first_name=first_name, last_name=last_name,
                  grade_level=int(grade_level), section_id=int(section_id), sex=sex)
    db.session.add(stu)
    db.session.flush()

    credentials = _generate_student_account(stu)
    db.session.commit()

    _log_audit("student_added",
                f"{current_user.full_name} added learner '{first_name} {last_name}' (LRN {lrn}).",
                user=current_user, school_id=current_user.school_id)

    if credentials:
        username, password = credentials
        flash(f"Student '{first_name} {last_name}' added. Learner portal account auto-created — "
              f"username: '{username}', password: '{password}'. Please share these with the learner.",
              "success")
    else:
        flash(f"Student '{first_name} {last_name}' added.", "success")
    return redirect(request.referrer or url_for("dashboard"))

@app.route("/teacher/student/<int:student_id>/create-account", methods=["POST"])
@roles_required("teacher")
def create_student_account(student_id):
    student = Student.query.get_or_404(student_id)
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    existing = User.query.filter_by(student_id=student.id).first()
    if existing:
        flash(f"{student.full_name} already has a learner account ('{existing.username}').", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    if not username or not password:
        flash("Username and password are required to create a learner account.", "error")
        return redirect(url_for("student_profile", student_id=student.id))
    if len(password) < 8:
        flash("Password must be at least 8 characters.", "error")
        return redirect(url_for("student_profile", student_id=student.id))
    if User.query.filter_by(username=username).first():
        flash("That username is already taken -- choose another.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    # Teacher-created accounts are auto-approved, same as coordinator- and
    # developer-created accounts elsewhere -- the approval workflow only
    # applies to public sign-ups.
    u = User(username=username, full_name=student.full_name, role="student",
             school_id=current_user.school_id, student_id=student.id)
    u.set_password(password)
    db.session.add(u)
    db.session.commit()
    _log_audit("account_edited",
                f"{current_user.full_name} created a learner portal account for {student.full_name} ('{username}').",
                user=current_user, school_id=current_user.school_id)
    flash(f"Learner account '{username}' created for {student.full_name}.", "success")
    return redirect(url_for("student_profile", student_id=student.id))


@app.route("/teacher/student/<int:student_id>/reset-account-password", methods=["POST"])
@roles_required("teacher")
def reset_student_account_password(student_id):
    student = Student.query.get_or_404(student_id)
    section = Section.query.filter_by(teacher_id=current_user.id).first()
    if not section or student.section_id != section.id:
        abort(403)

    account = User.query.filter_by(student_id=student.id).first()
    if not account:
        flash("This learner doesn't have a portal account yet.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    new_password = request.form.get("new_password", "")
    if len(new_password) < 8:
        flash("New password must be at least 8 characters.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    account.set_password(new_password)
    db.session.commit()
    _log_audit("password_changed",
                f"{current_user.full_name} reset the password for learner account '{account.username}'.",
                user=current_user, school_id=current_user.school_id)
    flash(f"Password reset for {student.full_name}'s learner account.", "success")
    return redirect(url_for("student_profile", student_id=student.id))

# ---------------------------------------------------------------------------
# Intervention Recommendations (dedicated page, teacher + coordinator)
# ---------------------------------------------------------------------------
def _check_intervention_scope(student):
    """Same section/school scoping used throughout the app -- teachers can
    only act on their own section's learners, coordinators on their own school."""
    if current_user.role == "teacher":
        section = Section.query.filter_by(teacher_id=current_user.id).first()
        if not section or student.section_id != section.id:
            abort(403)
    else:  # coordinator
        if not student.section or student.section.school_id != current_user.school_id:
            abort(403)


def _build_intervention_entry(student):
    """
    One row of the Intervention Recommendations worklist: the diagnostic
    confidence for this student's latest assessment, the auto-matched
    materials, and whether a decision has already been recorded for THIS
    specific assessment (so re-assessing a student surfaces them again).
    Returns None for students at Independent level or with no assessment.
    """
    assessments = student.assessments
    if not assessments:
        return None
    latest = assessments[-1]
    if latest.reading_level == "Independent":
        return None

    confidence = latest.diagnostic_confidence
    recommendations = _recommend_materials(student, assessments)
    latest_record = (InterventionRecord.query
                      .filter_by(student_id=student.id, assessment_id=latest.id)
                      .order_by(InterventionRecord.recorded_at.desc()).first())
    return {
        "student": student,
        "latest_assessment": latest,
        "confidence": confidence,
        "recommendations": recommendations,
        "reviewed": latest_record is not None,
    }


@app.route("/teacher/interventions")
@roles_required("teacher", "coordinator")
def intervention_recommendations():
    if current_user.role == "teacher":
        section = Section.query.filter_by(teacher_id=current_user.id).first()
        sections = [section] if section else []
    else:
        sections = Section.query.filter_by(school_id=current_user.school_id).all()

    worklist = []
    for sec in sections:
        for stu in sec.students:
            entry = _build_intervention_entry(stu)
            if entry:
                entry["section"] = sec
                worklist.append(entry)

    # Highest diagnostic confidence first -- these are the clearest calls to
    # act on; low-confidence, ambiguous cases sink toward the bottom.
    worklist.sort(key=lambda e: e["confidence"]["score"] if e["confidence"] else 0, reverse=True)

    return render_template("intervention_recommendations.html", worklist=worklist, sections=sections,
                            is_coordinator=(current_user.role == "coordinator"))


@app.route("/teacher/student/<int:student_id>/intervention")
@roles_required("teacher", "coordinator")
def intervention_detail(student_id):
    student = Student.query.get_or_404(student_id)
    _check_intervention_scope(student)

    assessments = student.assessments
    latest = assessments[-1] if assessments else None
    confidence = latest.diagnostic_confidence if latest else None
    recommendations = _recommend_materials(student, assessments) if latest else []

    modify_options = []
    if latest:
        target_grade = latest.passage_grade_level or student.grade_level
        modify_options = (InterventionMaterial.query.filter_by(grade_level=target_grade)
                           .order_by(InterventionMaterial.title).all())

    history = (InterventionRecord.query.filter_by(student_id=student.id)
               .order_by(InterventionRecord.recorded_at.desc()).all())

    return render_template("intervention_detail.html", student=student, latest=latest,
                            confidence=confidence, recommendations=recommendations,
                            modify_options=modify_options, history=history)


@app.route("/teacher/student/<int:student_id>/intervention/record", methods=["POST"])
@roles_required("teacher", "coordinator")
def record_intervention_decision(student_id):
    student = Student.query.get_or_404(student_id)
    _check_intervention_scope(student)

    decision = request.form.get("decision")
    if decision not in ("accepted", "modified", "declined"):
        flash("Invalid decision.", "error")
        return redirect(url_for("intervention_detail", student_id=student.id))

    assessments = student.assessments
    latest = assessments[-1] if assessments else None
    if not latest:
        flash("No assessment on file for this learner yet.", "error")
        return redirect(url_for("intervention_recommendations"))

    confidence = latest.diagnostic_confidence
    note = request.form.get("note", "").strip()
    material_id = request.form.get("material_id")
    material = None
    if decision in ("accepted", "modified") and material_id:
        material = InterventionMaterial.query.get_or_404(int(material_id))

    record = InterventionRecord(
        student_id=student.id, assessment_id=latest.id,
        weakest_skill=latest.weakest_skill,
        confidence_score=confidence["score"] if confidence else None,
        confidence_label=confidence["label"] if confidence else None,
        decision=decision,
        material_id=material.id if material else None,
        note=note or None,
        recorded_by=current_user.id,
    )
    db.session.add(record)

    # Accepting or modifying also creates the actual tracked assignment --
    # same MaterialAssignment model the Materials Library page uses, so
    # there's one source of truth for "what's currently assigned."
    if material:
        already_open = MaterialAssignment.query.filter_by(
            material_id=material.id, student_id=student.id, status="assigned"
        ).first()
        if not already_open:
            db.session.add(MaterialAssignment(
                material_id=material.id, student_id=student.id, assigned_by=current_user.id,
            ))

    db.session.commit()

    labels = {"accepted": "accepted", "modified": "modified and reassigned", "declined": "declined"}
    flash(f"Intervention recommendation {labels[decision]} for {student.full_name}.", "success")
    return redirect(url_for("intervention_detail", student_id=student.id))


# ---------------------------------------------------------------------------
# (d) Machine Learning profiling (school-wide, coordinator + teacher view)
# ---------------------------------------------------------------------------
@app.route("/analytics/clusters")
@login_required
def literacy_clusters():
    if current_user.role not in ("coordinator", "teacher", "developer"):
        abort(403)

    if current_user.role == "teacher":
        section = Section.query.filter_by(teacher_id=current_user.id).first()
        student_ids = [s.id for s in section.students] if section else []
    elif current_user.role == "coordinator":
        sections = Section.query.filter_by(school_id=current_user.school_id).all()
        student_ids = [s.id for sec in sections for s in sec.students]
    else:
        student_ids = [s.id for s in Student.query.all()]

    latest = _latest_assessment_per_student(student_ids)
    df = ml_utils.build_feature_frame(latest)
    clustered_df, meta = ml_utils.cluster_literacy_patterns(df)

    chart_html = None
    if not clustered_df.empty and "cluster" in clustered_df.columns and meta.get("model_used") != "none":
        id_to_name = {s.id: s.full_name for s in Student.query.filter(Student.id.in_(student_ids)).all()}
        clustered_df["student_name"] = clustered_df["student_id"].map(id_to_name)

        # Built as a plain go.Figure with one trace per cluster, using native
        # Python lists (not px.scatter on a DataFrame) -- Plotly's newer
        # PlotlyJSONEncoder base64-encodes numpy/pandas-backed trace data as
        # "typed arrays", which the pinned Plotly.js CDN version can't
        # decode, silently rendering an empty chart. Plain lists sidestep
        # that entirely and match how every other chart in this app is built.
        speeds = clustered_df["reading_speed_wpm"].tolist()
        min_speed, max_speed = (min(speeds), max(speeds)) if speeds else (0, 1)

        def scaled_size(v):
            if max_speed == min_speed:
                return 16
            return 8 + (v - min_speed) / (max_speed - min_speed) * 22

        fig = go.Figure()
        for cluster_id in sorted(clustered_df["cluster"].unique()):
            sub = clustered_df[clustered_df["cluster"] == cluster_id]
            fig.add_trace(go.Scatter(
                x=[float(v) for v in sub["word_recognition_accuracy"]],
                y=[float(v) for v in sub["comprehension_avg"]],
                mode="markers",
                name=f"Cluster {cluster_id}",
                marker=dict(size=[scaled_size(v) for v in sub["reading_speed_wpm"]]),
                text=[f"{name}<br>{level}<br>{speed} wpm" for name, level, speed in
                      zip(sub["student_name"], sub["reading_level"], sub["reading_speed_wpm"])],
                hoverinfo="text",
            ))
        fig.update_layout(
            margin=dict(t=30, b=10, l=10, r=10), height=420,
            xaxis_title="Word Recognition Accuracy (%)",
            yaxis_title="Comprehension Avg (%)",
            legend_title_text="Cluster",
        )
        chart_html = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    return render_template("clusters.html", meta=meta, chart_html=chart_html,
                            table=clustered_df.to_dict(orient="records") if not clustered_df.empty else [])


# ---------------------------------------------------------------------------
# (e) Academic risk prediction (single student, used inside student_profile too)
# ---------------------------------------------------------------------------
@app.route("/analytics/risk/<int:student_id>")
@login_required
def student_risk_api(student_id):
    student = Student.query.get_or_404(student_id)
    risk = _predict_student_risk(student.assessments)
    return jsonify(risk or {})


# ---------------------------------------------------------------------------
# (f) Audio documentation module
# ---------------------------------------------------------------------------
@app.route("/teacher/student/<int:student_id>/audio", methods=["POST"])
@roles_required("teacher")
def upload_audio_log(student_id):
    student = Student.query.get_or_404(student_id)
    file = request.files.get("audio_file")
    notes = request.form.get("notes", "")
    assessment_id = request.form.get("assessment_id") or None

    filename = None
    if file and file.filename:
        upload_dir = os.path.join(BASE_DIR, "instance", "audio_logs")
        os.makedirs(upload_dir, exist_ok=True)
        filename = f"{student.lrn}_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{file.filename}"
        file.save(os.path.join(upload_dir, filename))

    log = AudioLog(student_id=student.id, assessment_id=assessment_id, filename=filename, notes=notes)
    db.session.add(log)
    db.session.commit()
    flash("Oral reading audio log saved to student profile.", "success")
    return redirect(url_for("student_profile", student_id=student.id))


# ---------------------------------------------------------------------------
# Intervention Library
# ---------------------------------------------------------------------------
@app.route("/library")
@login_required
def intervention_library():
    if current_user.role == "student":
        # Students see their actual tracked assignments, not just the live
        # auto-matched suggestion list.
        student = Student.query.get_or_404(current_user.student_id)
        assignments = (MaterialAssignment.query.filter_by(student_id=student.id)
                       .order_by(MaterialAssignment.assigned_at.desc()).all())
        return render_template("student_library.html", assignments=assignments, student=student)

    skill = request.args.get("skill")
    level = request.args.get("level")
    grade = request.args.get("grade")
    query = InterventionMaterial.query
    if skill:
        query = query.filter_by(target_skill=skill)
    if level:
        query = query.filter_by(reading_level=level)
    if grade:
        query = query.filter_by(grade_level=int(grade))
    materials = query.order_by(InterventionMaterial.id.desc()).all()
    return render_template("intervention_library.html", materials=materials)

@app.route("/admin/resources")
@roles_required("developer")
def external_resources():
    resources = (InterventionMaterial.query.filter_by(is_external=True)
                 .order_by(InterventionMaterial.created_at.desc()).all())
    return render_template("external_resources.html", resources=resources)


@app.route("/admin/resources/add", methods=["POST"])
@roles_required("developer")
def add_external_resource():
    title = request.form.get("title", "").strip()
    target_skill = request.form.get("target_skill", "")
    reading_level = request.form.get("reading_level", "")
    grade_level = request.form.get("grade_level")
    description = request.form.get("description", "").strip()
    source_name = request.form.get("source_name", "").strip()
    source_url = request.form.get("source_url", "").strip()

    if not all([title, target_skill, reading_level, grade_level, source_name, source_url]):
        flash("Title, skill, level, grade, source name, and source URL are all required.", "error")
        return redirect(url_for("external_resources"))
    if not (source_url.startswith("http://") or source_url.startswith("https://")):
        flash("Source URL must be a valid http:// or https:// link.", "error")
        return redirect(url_for("external_resources"))

    resource = InterventionMaterial(
        title=title, target_skill=target_skill, reading_level=reading_level,
        grade_level=int(grade_level), description=description,
        is_external=True, source_name=source_name, source_url=source_url,
        uploaded_by=current_user.id,
    )
    db.session.add(resource)
    db.session.commit()
    _log_audit("external_resource_added",
                f"{current_user.full_name} added external resource '{title}' from {source_name}.",
                user=current_user)
    flash(f"External resource '{title}' added.", "success")
    return redirect(url_for("external_resources"))


@app.route("/admin/resources/<int:resource_id>/edit", methods=["GET", "POST"])
@roles_required("developer")
def edit_external_resource(resource_id):
    resource = InterventionMaterial.query.get_or_404(resource_id)
    if not resource.is_external:
        abort(404)

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        target_skill = request.form.get("target_skill", "")
        reading_level = request.form.get("reading_level", "")
        grade_level = request.form.get("grade_level")
        description = request.form.get("description", "").strip()
        source_name = request.form.get("source_name", "").strip()
        source_url = request.form.get("source_url", "").strip()

        if not all([title, target_skill, reading_level, grade_level, source_name, source_url]):
            flash("Title, skill, level, grade, source name, and source URL are all required.", "error")
            return render_template("edit_external_resource.html", resource=resource)
        if not (source_url.startswith("http://") or source_url.startswith("https://")):
            flash("Source URL must be a valid http:// or https:// link.", "error")
            return render_template("edit_external_resource.html", resource=resource)

        resource.title = title
        resource.target_skill = target_skill
        resource.reading_level = reading_level
        resource.grade_level = int(grade_level)
        resource.description = description
        resource.source_name = source_name
        resource.source_url = source_url
        db.session.commit()
        _log_audit("external_resource_edited",
                    f"{current_user.full_name} edited external resource '{title}'.",
                    user=current_user)
        flash(f"'{title}' updated.", "success")
        return redirect(url_for("external_resources"))

    return render_template("edit_external_resource.html", resource=resource)


@app.route("/admin/resources/<int:resource_id>/delete", methods=["POST"])
@roles_required("developer")
def delete_external_resource(resource_id):
    resource = InterventionMaterial.query.get_or_404(resource_id)
    if not resource.is_external:
        abort(404)

    in_use = MaterialAssignment.query.filter_by(material_id=resource.id).first()
    if in_use:
        flash(f"Can't delete '{resource.title}' -- it's currently assigned to one or more learners.", "error")
        return redirect(url_for("external_resources"))

    title = resource.title
    db.session.delete(resource)
    db.session.commit()
    _log_audit("external_resource_deleted",
                f"{current_user.full_name} deleted external resource '{title}'.",
                user=current_user)
    flash(f"'{title}' removed.", "success")
    return redirect(url_for("external_resources"))

def _save_material_file(file_storage):
    """
    Saves an uploaded material file into instance/materials/. Uses
    secure_filename() plus a random hex prefix to avoid collisions and to
    keep the on-disk name unguessable (this directory is never served
    directly -- everything goes through the auth-checked download route).
    Returns the stored filename, or None if no valid file was given.
    """
    if not file_storage or not file_storage.filename:
        return None
    ext = file_storage.filename.rsplit(".", 1)[-1].lower() if "." in file_storage.filename else ""
    if ext not in ALLOWED_MATERIAL_EXTENSIONS:
        return None
    safe_name = secure_filename(file_storage.filename)
    stored_name = f"{secrets.token_hex(8)}_{safe_name}"
    upload_dir = os.path.join(BASE_DIR, "instance", "materials")
    os.makedirs(upload_dir, exist_ok=True)
    file_storage.save(os.path.join(upload_dir, stored_name))
    return stored_name


@app.route("/library/upload", methods=["GET", "POST"])
@roles_required("teacher", "coordinator")
def upload_material():
    if request.method == "GET":
        # The upload form is embedded directly in intervention_library.html.
        return redirect(url_for("intervention_library"))

    title = request.form.get("title", "").strip()
    target_skill = request.form.get("target_skill", "")
    reading_level = request.form.get("reading_level", "")
    grade_level = request.form.get("grade_level")
    description = request.form.get("description", "").strip()

    if not all([title, target_skill, reading_level, grade_level]):
        flash("Title, target skill, reading level, and grade level are required.", "error")
        return redirect(url_for("intervention_library"))

    upload = request.files.get("material_file")
    stored_name = _save_material_file(upload)
    if upload and upload.filename and not stored_name:
        flash("That file type isn't supported. Allowed: PDF, Word, PowerPoint, images, audio, zip, txt.", "error")
        return redirect(url_for("intervention_library"))

    material = InterventionMaterial(
        title=title, target_skill=target_skill, reading_level=reading_level,
        grade_level=int(grade_level), description=description,
        file_path=stored_name,
        original_filename=upload.filename if stored_name else None,
        uploaded_by=current_user.id,
    )
    db.session.add(material)
    db.session.commit()
    flash(f"Material '{title}' added to the library.", "success")
    return redirect(url_for("intervention_library"))


@app.route("/library/download/<int:material_id>")
@login_required
def download_material(material_id):
    material = InterventionMaterial.query.get_or_404(material_id)

    if current_user.role == "student":
        assigned = MaterialAssignment.query.filter_by(
            material_id=material.id, student_id=current_user.student_id
        ).first()
        if not assigned:
            abort(403)

    if material.is_external:
        if not material.source_url:
            abort(404)
        return redirect(material.source_url)

    if not material.file_path:
        abort(404)

    materials_dir = os.path.join(BASE_DIR, "instance", "materials")
    download_name = material.original_filename or material.file_path
    return send_from_directory(materials_dir, material.file_path,
                                as_attachment=True, download_name=download_name)

@app.route("/teacher/student/<int:student_id>/assign-material", methods=["POST"])
@roles_required("teacher", "coordinator")
def assign_material(student_id):
    student = Student.query.get_or_404(student_id)

    # Same section/school scoping used everywhere else in the app.
    if current_user.role == "teacher":
        section = Section.query.filter_by(teacher_id=current_user.id).first()
        if not section or student.section_id != section.id:
            abort(403)
    else:  # coordinator
        if not student.section or student.section.school_id != current_user.school_id:
            abort(403)

    material_id = request.form.get("material_id")
    material = InterventionMaterial.query.get_or_404(int(material_id)) if material_id else None
    if not material:
        flash("Select a material to assign.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    already_open = MaterialAssignment.query.filter_by(
        material_id=material.id, student_id=student.id, status="assigned"
    ).first()
    if already_open:
        flash(f"'{material.title}' is already assigned to {student.full_name}.", "error")
        return redirect(url_for("student_profile", student_id=student.id))

    db.session.add(MaterialAssignment(
        material_id=material.id, student_id=student.id, assigned_by=current_user.id,
    ))
    db.session.commit()
    _log_audit("material_assigned",
                f"{current_user.full_name} assigned '{material.title}' to {student.full_name}.",
                user=current_user, school_id=current_user.school_id)
    flash(f"'{material.title}' assigned to {student.full_name}.", "success")
    return redirect(url_for("student_profile", student_id=student.id))


@app.route("/student/materials/<int:assignment_id>/complete", methods=["POST"])
@roles_required("student")
def complete_material_assignment(assignment_id):
    assignment = MaterialAssignment.query.get_or_404(assignment_id)
    if assignment.student_id != current_user.student_id:
        abort(403)

    assignment.status = "completed"
    assignment.completed_at = datetime.utcnow()
    db.session.commit()
    flash(f"Marked '{assignment.material.title}' as completed. Great job!", "success")
    return redirect(url_for("intervention_library"))


# ---------------------------------------------------------------------------
# (student role) restricted personal view
# ---------------------------------------------------------------------------
@app.route("/student")
@roles_required("student")
def student_dashboard():
    student = Student.query.get_or_404(current_user.student_id)

    # Materials-completion progress (e.g. "3 of 5 assigned materials done,
    # 60%") is intentionally NOT built yet -- planned as a future addition.
    # Reading level, assessment scores, and history are deliberately
    # withheld from the learner-facing view per the teacher's request;
    # that data still exists and is fully visible to Teacher/Coordinator
    # on the student profile page, just not surfaced here.
    assignments = (MaterialAssignment.query.filter_by(student_id=student.id)
                   .order_by(MaterialAssignment.assigned_at.desc()).all())

    return render_template("student_dashboard.html", student=student, assignments=assignments)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def _log_audit(action, description, user=None, school_id=None):
    """
    Self-contained -- adds AND commits its own entry, so the audit trail
    stays durable even if something in the caller's own transaction later
    rolls back. Never raises: a logging failure should never take down the
    actual feature it's describing.
    """
    try:
        entry = AuditLog(
            user_id=user.id if user else None,
            school_id=school_id if school_id is not None else (user.school_id if user else None),
            action=action,
            description=description,
            ip_address=request.remote_addr if request else None,
        )
        db.session.add(entry)
        db.session.commit()
    except Exception:
        db.session.rollback()


def _latest_assessment_per_student(student_ids):
    latest = []
    for sid in student_ids:
        a = (Assessment.query.filter_by(student_id=sid)
             .order_by(Assessment.date_administered.desc()).first())
        if a:
            latest.append(a)
    return latest


def _evaluate_progress_routing(student, gst, assessments):
    """
    Progress Evaluation Module -- the real Phil-IRI decision tree:

    1. EXEMPT (no intervention at all): only if the learner was on-grade-level
       at the GST (0 levels down) AND scored Independent on the Pre-test.
    2. MUST ENTER INTERVENTION: if levels-down > 0, even an Independent
       Pre-test score doesn't exempt them -- the goal is climbing back to
       their OWN grade level, not just passing the lower passage they were
       placed at.
    3. TEACHER OVERRIDE: a teacher can manually pass a learner early at any
       point after Pre-test, skipping the remaining cycle. This is never
       automatic -- only set via explicit teacher action.
    4. STILL BELOW GRADE LEVEL: even an Independent score doesn't fully
       close the loop until the passage grade used matches the learner's
       own enrolled grade -- tracked per-assessment via passage_grade_level.
    5. COMPLETED / FOR ARAL: outcome at Post-test -- passed at their own
       grade level, or flagged for the (separately-scoped) ARAL summer
       program if still struggling.
    """
    by_term = {a.term: a for a in assessments}
    pre = by_term.get("pre")
    mid = by_term.get("mid")
    post = by_term.get("post")
    levels_down = gst.levels_down if gst else 0
    own_grade = student.grade_level

    def at_own_grade(a):
        return (a.passage_grade_level or own_grade) >= own_grade

    # Teacher override always wins, wherever it was set.
    overridden = next((a for a in assessments if a.teacher_override_pass), None)
    if overridden:
        return {
            "status": "passed_override",
            "decision": f"Passed early (teacher override at {overridden.term.title()}-test)",
            "rationale": overridden.override_note or "Teacher judged the learner ready without finishing the full test cycle.",
            "levels_down": levels_down,
        }

    if not pre:
        return {"status": "pending", "message": "Pre-test not yet recorded. Routing unavailable.",
                "levels_down": levels_down}

    # --- Exemption check: only possible right after the Pre-test ---
    if levels_down == 0 and pre.reading_level == "Independent":
        return {
            "status": "exempt",
            "decision": "Passed / Independent -- no intervention needed",
            "rationale": "On grade level at GST and Independent on the Pre-test.",
            "levels_down": levels_down,
        }

    if levels_down > 0 and pre.reading_level == "Independent" and not at_own_grade(pre):
        exemption_note = (
            f"Independent at the Grade {pre.passage_grade_level} passage, but still "
            f"{levels_down} level{'s' if levels_down > 1 else ''} below Grade {own_grade}. "
            "Intervention continues -- the goal is passing at their own grade level, not just this one."
        )
    else:
        exemption_note = None

    if not mid:
        return {
            "status": "in_intervention",
            "decision": "In intervention -- awaiting Mid-year assessment",
            "rationale": exemption_note or f"Pre-test result: {pre.reading_level}.",
            "levels_down": levels_down,
        }

    # --- Mid-year decision point ---
    mid_closed_gap = mid.reading_level == "Independent" and at_own_grade(mid)
    eligible_for_early_pass = mid.reading_level == "Independent"

    if not post:
        if mid.reading_level == "Frustration":
            decision = "CONTINUE intervention"
            rationale = "Mid-year scores remain at Frustration level."
        elif mid.reading_level == "Instructional":
            decision = "CONTINUE intervention (monitoring)"
            rationale = "Mid-year scores show partial progress."
        elif mid_closed_gap:
            decision = "On track -- Independent at own grade level"
            rationale = "Consider Post-test to confirm, or teacher may pass early."
        else:
            decision = "CONTINUE intervention"
            rationale = (f"Independent at Grade {mid.passage_grade_level}, but not yet at "
                          f"Grade {own_grade}. Still working toward their own grade level.")

        return {
            "status": "in_intervention",
            "decision": decision,
            "rationale": rationale,
            "mid_level": mid.reading_level,
            "eligible_for_early_pass": eligible_for_early_pass,
            "levels_down": levels_down,
        }

    # --- Post-test / Final outcome ---
    if post.reading_level == "Independent" and at_own_grade(post):
        return {
            "status": "completed",
            "decision": "Completed -- Independent at own grade level",
            "rationale": "Post-test confirms the learner has closed the gap.",
            "levels_down": levels_down,
        }
    else:
        return {
            "status": "for_aral",
            "decision": "Still struggling at Post-test -- refer to ARAL",
            "rationale": ("Learner has not yet reached Independent at their own grade level "
                          "after the full intervention cycle. (ARAL summer program tracking is "
                          "handled outside this system.)"),
            "levels_down": levels_down,
        }


def _recommend_materials(student, assessments):
    """
    Instructional Intervention Library: recommends targeted reading materials
    matched to the student's documented literacy gaps (weakest skill area).

    Materials are matched to the PASSAGE grade level actually used for the
    assessment (which may be below the learner's enrolled grade if GST
    placed them levels-down) — a Grade 6 learner reading Grade 4 passages
    needs Grade 4 materials, not Grade 6 ones.
    """
    if not assessments:
        return []
    latest = assessments[-1]
    if latest.reading_level == "Independent":
        return []  # no remediation materials needed

    # Coarse two-way match for the material LIBRARY (word_recognition vs.
    # comprehension) — the library isn't yet split into literal/inferential/
    # critical buckets. The finer four-way diagnosis is shown separately via
    # Assessment.weakest_skill on the profile page.
    gaps = {
        "word_recognition": latest.word_recognition_accuracy,
        "comprehension": latest.comprehension_avg,
    }
    weakest_skill = min(gaps, key=gaps.get)
    target_grade = latest.passage_grade_level or student.grade_level

    materials = InterventionMaterial.query.filter_by(
        target_skill=weakest_skill, reading_level=latest.reading_level,
        grade_level=target_grade,
    ).all()

    if not materials:
        materials = InterventionMaterial.query.filter_by(
            target_skill=weakest_skill, grade_level=target_grade,
        ).all()

    return materials


def _predict_student_risk(assessments):
    if len(assessments) < 1:
        return None
    rows = []
    term_order_map = {"pre": 0, "mid": 1, "post": 2}
    base = 0
    for i, a in enumerate(assessments):
        order = base + term_order_map.get(a.term, i)
        rows.append({"student_id": a.student_id, "term_order": order,
                      "composite": ml_utils.composite_score(a.word_recognition_accuracy, a.comprehension_avg)})
        base += 3
    df = pd.DataFrame(rows)
    return ml_utils.predict_academic_risk(df)

def _build_login_activity_chart(role_filter, start_param, end_param):
    """
    Monthly count of successful logins, optionally filtered to one role.
    start_param/end_param are 'YYYY-MM' strings from an <input type="month">;
    defaults to the trailing 6 months ending this month when not given.
    """
    query = AuditLog.query.filter(AuditLog.action == "login_success")
    if role_filter:
        role_user_ids = [u.id for u in User.query.filter_by(role=role_filter).all()]
        query = query.filter(AuditLog.user_id.in_(role_user_ids)) if role_user_ids else query.filter(db.false())
    logs = query.all()

    now = datetime.utcnow()
    if end_param:
        end_year, end_month = map(int, end_param.split("-"))
    else:
        end_year, end_month = now.year, now.month
    if start_param:
        start_year, start_month = map(int, start_param.split("-"))
    else:
        m, y = end_month - 5, end_year
        while m <= 0:
            m += 12
            y -= 1
        start_year, start_month = y, m

    buckets = []
    y, m = start_year, start_month
    while (y, m) <= (end_year, end_month) and len(buckets) < 60:
        buckets.append((y, m))
        m += 1
        if m > 12:
            m, y = 1, y + 1

    counts = {b: 0 for b in buckets}
    for log in logs:
        if log.created_at:
            key = (log.created_at.year, log.created_at.month)
            if key in counts:
                counts[key] += 1

    labels = [f"{month_abbr[m]} {y}" for (y, m) in buckets]
    values = [counts[b] for b in buckets]

    fig = go.Figure(data=[go.Scatter(
        x=labels, y=values, mode="lines+markers",
        line=dict(color="#1f5c4d", width=2), marker=dict(size=6),
    )])
    fig.update_layout(
        margin=dict(t=20, b=10, l=10, r=10), height=280,
        yaxis=dict(title="Successful Logins", rangemode="tozero"),
    )
    chart_json = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    start_value = f"{start_year:04d}-{start_month:02d}"
    end_value = f"{end_year:04d}-{end_month:02d}"
    return chart_json, start_value, end_value

def _build_trend_chart(assessments):
    if not assessments:
        return None
    terms = [f"{a.term.title()} ({a.school_year})" for a in assessments]
    wr = [a.word_recognition_accuracy for a in assessments]
    comp = [a.comprehension_avg for a in assessments]
    speed = [a.reading_speed_wpm for a in assessments]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=terms, y=wr, mode="lines+markers", name="Word Recognition Accuracy (%)"))
    fig.add_trace(go.Scatter(x=terms, y=comp, mode="lines+markers", name="Comprehension Avg (%)"))
    fig.add_trace(go.Scatter(x=terms, y=speed, mode="lines+markers", name="Reading Speed (WPM)", yaxis="y2"))
    fig.update_layout(
        margin=dict(t=30, b=10, l=10, r=10), height=350,
        yaxis=dict(title="Percent (%)", range=[0, 100]),
        yaxis2=dict(title="WPM", overlaying="y", side="right"),
        legend=dict(orientation="h", y=-0.2),
    )
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)


def _build_heatmap(sections, latest_assessments):
    """Literacy heatmap: sections (rows) x reading-level distribution (columns)."""
    by_section = {}
    a_by_student = {a.student_id: a for a in latest_assessments}
    for sec in sections:
        counts = {"Independent": 0, "Instructional": 0, "Frustration": 0}
        for stu in sec.students:
            a = a_by_student.get(stu.id)
            if a and a.reading_level in counts:
                counts[a.reading_level] += 1
        by_section[sec.name] = counts

    if not by_section:
        return None

    levels = ["Independent", "Instructional", "Frustration"]
    z = [[by_section[sec][lvl] for lvl in levels] for sec in by_section]
    fig = go.Figure(data=go.Heatmap(
        z=z, x=levels, y=list(by_section.keys()),
        colorscale=[[0, "#f5f5f5"], [1, "#1565c0"]],
        showscale=True,
    ))
    fig.update_layout(margin=dict(t=20, b=10, l=10, r=10), height=280)
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)


def _build_grade_level_chart(students):
    """
    Reading Level by Grade: grouped bar chart, grade level (x) x reading-level
    distribution (grouped bars), using each student's LATEST assessment.
    Hand-built go.Figure with plain Python lists -- never px.bar on a raw
    DataFrame (see the clustering-page note: the pinned Plotly.js CDN can't
    decode the base64 'typed arrays' the newer PlotlyJSONEncoder emits for
    DataFrame-backed traces, which silently renders an empty chart).
    """
    grades = sorted(set(s.grade_level for s in students))
    if not grades:
        return None

    student_ids = [s.id for s in students]
    latest_by_student = {a.student_id: a for a in _latest_assessment_per_student(student_ids)}
    grade_of = {s.id: s.grade_level for s in students}

    levels = ["Independent", "Instructional", "Frustration"]
    grade_index = {g: i for i, g in enumerate(grades)}
    counts = {lvl: [0] * len(grades) for lvl in levels}

    for sid, a in latest_by_student.items():
        g = grade_of.get(sid)
        if g is None or a.reading_level not in levels:
            continue
        counts[a.reading_level][grade_index[g]] += 1

    fig = go.Figure()
    for lvl in levels:
        fig.add_trace(go.Bar(
            name=lvl, x=[f"Grade {g}" for g in grades], y=counts[lvl],
            marker=dict(color=READING_LEVEL_COLORS[lvl]),
        ))
    fig.update_layout(
        barmode="group", margin=dict(t=20, b=10, l=10, r=10), height=340,
        yaxis_title="Number of Students", legend_title_text="Reading Level",
    )
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)


def _build_grade_trend_chart(students):
    """
    Performance Trends: average comprehension score across Pre/Mid/Post, one
    line per grade level. Uses each student's ENROLLED grade (student.grade_level),
    not the passage grade used for a given assessment -- this tracks how the
    Grade 4/5/6 cohorts are trending as groups, not where individual learners
    were placed by the GST.
    """
    grades = sorted(set(s.grade_level for s in students))
    if not grades:
        return None

    grade_of = {s.id: s.grade_level for s in students}
    terms = ["pre", "mid", "post"]
    term_labels = ["Pre-test", "Mid-year", "Post-test"]

    scores_by_grade_term = {g: {t: [] for t in terms} for g in grades}
    for s in students:
        g = grade_of[s.id]
        for a in s.assessments:
            if a.term in terms:
                scores_by_grade_term[g][a.term].append(a.comprehension_avg)

    fig = go.Figure()
    for g in grades:
        y = []
        for t in terms:
            vals = scores_by_grade_term[g][t]
            y.append(round(sum(vals) / len(vals), 1) if vals else None)
        fig.add_trace(go.Scatter(x=term_labels, y=y, mode="lines+markers", name=f"Grade {g}"))
    fig.update_layout(
        margin=dict(t=20, b=10, l=10, r=10), height=340,
        yaxis=dict(title="Avg. Comprehension (%)", range=[0, 100]),
        legend_title_text="Grade",
    )
    return json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)


def _build_intervention_effectiveness_chart(students):
    """
    Intervention Effectiveness: before/after bar comparing Pre-test vs
    Post-test scores for students who actually went through intervention --
    i.e. excluding anyone whose routing outcome was 'exempt' (never needed
    remediation) or 'passed_override' (teacher skipped the remaining cycle
    early), since neither is a fair test of whether the intervention itself
    worked. Requires both a Pre-test and a Post-test on file.

    Returns (chart_json, stats) -- stats always has at least {"n": ...} so
    the template can show a clear "not enough data" message instead of a
    misleading empty chart.
    """
    pre_wr, pre_comp, post_wr, post_comp = [], [], [], []
    for s in students:
        assessments = s.assessments
        by_term = {a.term: a for a in assessments}
        pre, post = by_term.get("pre"), by_term.get("post")
        if not pre or not post:
            continue
        gst = s.latest_gst
        routing = _evaluate_progress_routing(s, gst, assessments)
        if routing.get("status") in ("exempt", "passed_override"):
            continue
        pre_wr.append(pre.word_recognition_accuracy)
        pre_comp.append(pre.comprehension_avg)
        post_wr.append(post.word_recognition_accuracy)
        post_comp.append(post.comprehension_avg)

    n = len(pre_comp)
    if n == 0:
        return None, {"n": 0}

    avg_pre_wr = round(sum(pre_wr) / n, 1)
    avg_post_wr = round(sum(post_wr) / n, 1)
    avg_pre_comp = round(sum(pre_comp) / n, 1)
    avg_post_comp = round(sum(post_comp) / n, 1)

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Pre-test", x=["Word Recognition", "Comprehension"],
                          y=[avg_pre_wr, avg_pre_comp], marker=dict(color="#a7761d")))
    fig.add_trace(go.Bar(name="Post-test", x=["Word Recognition", "Comprehension"],
                          y=[avg_post_wr, avg_post_comp], marker=dict(color="#1f5c4d")))
    fig.update_layout(
        barmode="group", margin=dict(t=20, b=10, l=10, r=10), height=340,
        yaxis=dict(title="Average Score (%)", range=[0, 100]),
    )
    chart_json = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    stats = {
        "n": n,
        "avg_pre_wr": avg_pre_wr, "avg_post_wr": avg_post_wr,
        "avg_pre_comp": avg_pre_comp, "avg_post_comp": avg_post_comp,
        "wr_gain": round(avg_post_wr - avg_pre_wr, 1),
        "comp_gain": round(avg_post_comp - avg_pre_comp, 1),
    }
    return chart_json, stats

# Error handlers
@app.errorhandler(403)
def forbidden(e):
    return render_template("error.html", code=403, message="You don't have permission to view this page."), 403


@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404, message="Page not found."), 404


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)