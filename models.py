from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

def _percentage_tier(pct):
    if pct < 40:
        return 0
    elif pct < 70:
        return 1
    return 2


class School(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    sections = db.relationship("Section", backref="school", lazy=True)


class Section(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    grade_level = db.Column(db.Integer, nullable=False)
    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    students = db.relationship("Student", backref="section", lazy=True)


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(150), unique=True, nullable=True)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(150), nullable=False)
    role = db.Column(db.String(20), nullable=False)
    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=True)
    pending_school_name = db.Column(db.String(150), nullable=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=True)

    status = db.Column(db.String(20), nullable=False, default="approved")
    requested_at = db.Column(db.DateTime, default=datetime.utcnow)
    reviewed_at = db.Column(db.DateTime, nullable=True)
    reviewed_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    school = db.relationship("School", foreign_keys=[school_id])
    student = db.relationship("Student", foreign_keys=[student_id])

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_active(self):
        return self.status == "approved"


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lrn = db.Column(db.String(20), unique=True, nullable=False)
    first_name = db.Column(db.String(80), nullable=False)
    last_name = db.Column(db.String(80), nullable=False)
    grade_level = db.Column(db.Integer, nullable=False)
    section_id = db.Column(db.Integer, db.ForeignKey("section.id"), nullable=False)
    sex = db.Column(db.String(10))

    assessments = db.relationship("Assessment", backref="student", lazy=True,
                                   order_by="Assessment.date_administered")
    gst_records = db.relationship("GroupScreeningTest", backref="student", lazy=True,
                                   order_by="GroupScreeningTest.date_administered")
    audio_logs = db.relationship("AudioLog", backref="student", lazy=True)

    @property
    def full_name(self):
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def latest_gst(self):
        return self.gst_records[-1] if self.gst_records else None


class GroupScreeningTest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    school_year = db.Column(db.String(9), nullable=False)
    grade_level_at_test = db.Column(db.Integer, nullable=False)
    score = db.Column(db.Integer, nullable=False)
    total_items = db.Column(db.Integer, nullable=False, default=20)
    date_administered = db.Column(db.DateTime, default=datetime.utcnow)
    encoded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)

    @property
    def percentage(self):
        if not self.total_items:
            return 0.0
        return round((self.score / self.total_items) * 100, 1)

    @property
    def levels_down(self):
        return {0: 2, 1: 1, 2: 0}[_percentage_tier(self.percentage)]

    @property
    def placement_label(self):
        ld = self.levels_down
        if ld == 0:
            return "On grade level"
        return f"{ld} level{'s' if ld > 1 else ''} down"

    @property
    def placement_grade_level(self):
        return max(1, self.grade_level_at_test - self.levels_down)

    @property
    def placement_in_library_range(self):
        return 4 <= self.placement_grade_level <= 6


class Assessment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    school_year = db.Column(db.String(9), nullable=False)
    term = db.Column(db.String(10), nullable=False)
    passage_grade_level = db.Column(db.Integer, nullable=True)

    word_recognition_correct = db.Column(db.Integer, nullable=False)
    word_recognition_total = db.Column(db.Integer, nullable=False)
    literal_correct = db.Column(db.Integer, nullable=False)
    literal_total = db.Column(db.Integer, nullable=False)
    inferential_correct = db.Column(db.Integer, nullable=False)
    inferential_total = db.Column(db.Integer, nullable=False)
    critical_correct = db.Column(db.Integer, nullable=False)
    critical_total = db.Column(db.Integer, nullable=False)

    reading_speed_wpm = db.Column(db.Float, nullable=False)
    teacher_override_pass = db.Column(db.Boolean, default=False, nullable=False)
    override_note = db.Column(db.Text, nullable=True)

    date_administered = db.Column(db.DateTime, default=datetime.utcnow)
    encoded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)

    @staticmethod
    def _pct(correct, total):
        if not total:
            return 0.0
        return round((correct / total) * 100, 1)

    @property
    def word_recognition_accuracy(self):
        return self._pct(self.word_recognition_correct, self.word_recognition_total)

    @property
    def comprehension_literal(self):
        return self._pct(self.literal_correct, self.literal_total)

    @property
    def comprehension_inferential(self):
        return self._pct(self.inferential_correct, self.inferential_total)

    @property
    def comprehension_critical(self):
        return self._pct(self.critical_correct, self.critical_total)

    @property
    def comprehension_avg(self):
        total_correct = self.literal_correct + self.inferential_correct + self.critical_correct
        total_items = self.literal_total + self.inferential_total + self.critical_total
        return self._pct(total_correct, total_items)

    @property
    def reading_level(self):
        tier = _percentage_tier(self.comprehension_avg)
        return {0: "Frustration", 1: "Instructional", 2: "Independent"}[tier]

    @property
    def weakest_skill(self):
        if self.reading_level == "Independent":
            return None
        scores = {
            "Word Recognition": self.word_recognition_accuracy,
            "Literal Comprehension": self.comprehension_literal,
            "Inferential Comprehension": self.comprehension_inferential,
            "Critical Comprehension": self.comprehension_critical,
        }
        return min(scores, key=scores.get)

    @property
    def diagnostic_confidence(self):
        if self.reading_level == "Independent":
            return None
        scores = sorted([
            self.word_recognition_accuracy,
            self.comprehension_literal,
            self.comprehension_inferential,
            self.comprehension_critical,
        ])
        gap = round(scores[1] - scores[0], 1)
        pct = min(100, round(gap / 40 * 100))
        if pct >= 66:
            label = "High"
        elif pct >= 33:
            label = "Medium"
        else:
            label = "Low"
        return {"score": pct, "label": label, "gap": gap}

class InterventionMaterial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    target_skill = db.Column(db.String(30), nullable=False)
    reading_level = db.Column(db.String(20), nullable=False)
    grade_level = db.Column(db.Integer, nullable=False)
    description = db.Column(db.Text)
    file_path = db.Column(db.String(255), nullable=True)
    original_filename = db.Column(db.String(255), nullable=True)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # External, research-supported resources curated by the System Admin --
    # distinct from a teacher's own file upload. No file lives on this
    # server for these; they link out to a vetted external source instead.
    is_external = db.Column(db.Boolean, default=False, nullable=False)
    source_name = db.Column(db.String(150), nullable=True)   # e.g. "Reading Rockets", "DepEd LRMDS"
    source_url = db.Column(db.String(500), nullable=True)

    assignments = db.relationship("MaterialAssignment", backref="material", lazy=True)

class MaterialAssignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    material_id = db.Column(db.Integer, db.ForeignKey("intervention_material.id"), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    assigned_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), nullable=False, default="assigned")
    completed_at = db.Column(db.DateTime, nullable=True)
    student = db.relationship("Student", backref="material_assignments")


class PassageTotalPreset(db.Model):
    """
    Remembers the standard total-item counts for a school + grade + specific
    assessment stage. 'term' is one of 'gst' | 'pre' | 'mid' | 'post' -- a
    Mid-year passage can legitimately have a different word count than the
    Pre-test passage for the same grade, so totals are now scoped per stage
    rather than shared across the whole grade.

    These are ONLY editable from the teacher's "Assessment Totals" panel on
    the dashboard -- encode forms read them but can never write them, so a
    teacher can't accidentally change an official total while scoring an
    individual learner.
    """
    id = db.Column(db.Integer, primary_key=True)
    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=False)
    grade_level = db.Column(db.Integer, nullable=False)
    term = db.Column(db.String(10), nullable=False, default="gst")

    word_recognition_total = db.Column(db.Integer, nullable=True)
    literal_total = db.Column(db.Integer, nullable=True)
    inferential_total = db.Column(db.Integer, nullable=True)
    critical_total = db.Column(db.Integer, nullable=True)
    gst_total_items = db.Column(db.Integer, nullable=True)

    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    updated_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)

    __table_args__ = (db.UniqueConstraint("school_id", "grade_level", "term", name="uq_preset_school_grade_term"),)


class InterventionRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    assessment_id = db.Column(db.Integer, db.ForeignKey("assessment.id"), nullable=True)
    weakest_skill = db.Column(db.String(50), nullable=True)
    confidence_score = db.Column(db.Integer, nullable=True)
    confidence_label = db.Column(db.String(10), nullable=True)
    decision = db.Column(db.String(20), nullable=False)
    material_id = db.Column(db.Integer, db.ForeignKey("intervention_material.id"), nullable=True)
    note = db.Column(db.Text, nullable=True)
    recorded_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    recorded_at = db.Column(db.DateTime, default=datetime.utcnow)

    student = db.relationship("Student", backref="intervention_records")
    assessment = db.relationship("Assessment")
    material = db.relationship("InterventionMaterial")


class SystemSetting(db.Model):
    key = db.Column(db.String(100), primary_key=True)
    value = db.Column(db.Text, nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    updated_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)


class PasswordResetToken(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    token = db.Column(db.String(64), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime, nullable=False)
    used = db.Column(db.Boolean, default=False, nullable=False)
    user = db.relationship("User")


class AudioLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    assessment_id = db.Column(db.Integer, db.ForeignKey("assessment.id"), nullable=True)
    filename = db.Column(db.String(255))
    notes = db.Column(db.Text)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=True)
    action = db.Column(db.String(40), nullable=False)
    description = db.Column(db.Text, nullable=False)
    ip_address = db.Column(db.String(64), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    actor = db.relationship("User", foreign_keys=[user_id])
    school = db.relationship("School", foreign_keys=[school_id])