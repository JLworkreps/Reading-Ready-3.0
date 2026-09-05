"""
Seeds the database with demo data across all four schools:
Villasis 1 Central School SPED Center, San Blas ES, Piaz ES, Puelay ES.
Run: python seed_data.py
"""
import random
from datetime import datetime, timedelta
from app import app
from models import (db, School, Section, User, Student, Assessment, GroupScreeningTest,
                     InterventionMaterial, MaterialAssignment, InterventionRecord,
                     SystemSetting, AuditLog)

random.seed(42)

FIRST_NAMES = ["Juan", "Maria", "Jose", "Ana", "Pedro", "Rosa", "Carlos", "Liza",
               "Mark", "Grace", "Paolo", "Angel", "Rico", "Bea", "Noel", "Cathy",
               "Kevin", "Joy", "Ryan", "Faith", "Aldrin", "Mica", "Jerome", "Nica",
               "Michael", "Sofia", "Daniel", "Ella", "Vince", "Kaye", "Miguel", "Trisha"]
LAST_NAMES = ["Santos", "Reyes", "Cruz", "Bautista", "Garcia", "Torres", "Flores",
              "Ramos", "Dela Cruz", "Mendoza", "Castillo", "Aquino", "Villanueva",
              "Domingo", "Salazar", "Navarro", "Pascual", "Gonzales", "Del Rosario",
              "Manalo", "Fernandez", "Ocampo"]

SCHOOL_CONFIGS = [
    {
        "coordinator_username": "coordinator1", "coordinator_name": "Ma. Teresa Villanueva",
        "coordinator_email": "coordinator1@example.com",
        "sections": [
            ("Grade 4 - Mabini", 4, 35, "teacher_g4", "Mr. Ramon Cruz", "ramon.cruz@example.com"),
            ("Grade 4 - Rizal", 4, 40, "teacher_g4b", "Ms. Divina Santos", "divina.santos@example.com"),
            ("Grade 5 - Bonifacio", 5, 40, "teacher_g5", "Ms. Liezel Ramos", "liezel.ramos@example.com"),
            ("Grade 6 - Aguinaldo", 6, 35, "teacher_g6", "Mr. Arnold Torres", "arnold.torres@example.com"),
        ],
    },
    {
        "coordinator_username": "coordinator2", "coordinator_name": "Mr. Danilo Fernandez",
        "coordinator_email": "coordinator2@example.com",
        "sections": [
            ("Grade 4 - Mabuhay", 4, 38, "sb_teacher_g4", "Ms. Corazon Lim", "corazon.lim@example.com"),
            ("Grade 5 - Malaya", 5, 36, "sb_teacher_g5", "Mr. Bayani Ocampo", "bayani.ocampo@example.com"),
            ("Grade 6 - Bayanihan", 6, 34, "sb_teacher_g6", "Ms. Perla Guevarra", "perla.guevarra@example.com"),
        ],
    },
    {
        "coordinator_username": "coordinator3", "coordinator_name": "Ms. Editha Ramirez",
        "coordinator_email": "coordinator3@example.com",
        "sections": [
            ("Grade 4 - Sampaguita", 4, 32, "pz_teacher_g4", "Mr. Ferdinand Bonoan", "ferdinand.bonoan@example.com"),
            ("Grade 5 - Ilang-Ilang", 5, 34, "pz_teacher_g5", "Ms. Marilou Espino", "marilou.espino@example.com"),
            ("Grade 6 - Rosal", 6, 30, "pz_teacher_g6", "Mr. Wilfredo Manalo", "wilfredo.manalo@example.com"),
        ],
    },
    {
        "coordinator_username": "coordinator4", "coordinator_name": "Ms. Josefina Del Rosario",
        "coordinator_email": "coordinator4@example.com",
        "sections": [
            ("Grade 4 - Malikhain", 4, 30, "pl_teacher_g4", "Ms. Remedios Tolentino", "remedios.tolentino@example.com"),
            ("Grade 5 - Matatag", 5, 33, "pl_teacher_g5", "Mr. Gregorio Panganiban", "gregorio.panganiban@example.com"),
            ("Grade 6 - Matalino", 6, 31, "pl_teacher_g6", "Ms. Nenita Aledia", "nenita.aledia@example.com"),
        ],
    },
]


def make_score(base, spread=8):
    return round(max(20, min(100, random.gauss(base, spread))), 1)


def make_gst_score(profile):
    if profile == "strong":
        return random.randint(15, 20)
    elif profile == "developing":
        return random.randint(9, 13)
    else:
        return random.randint(2, 7)


LEVELS_DOWN_BY_PROFILE = {"strong": 0, "developing": 1, "struggling": 2}


def create_assessment(student, term, school_year, profile, encoder_id):
    if profile == "strong":
        wr_pct, comp_base, speed = make_score(97, 2), 88, random.gauss(130, 10)
    elif profile == "developing":
        wr_pct, comp_base, speed = make_score(92, 4), 68, random.gauss(95, 12)
    else:
        wr_pct, comp_base, speed = make_score(80, 8), 45, random.gauss(60, 10)

    lit_pct = round(max(10, min(100, random.gauss(comp_base + 5, 8))), 1)
    inf_pct = round(max(10, min(100, random.gauss(comp_base, 9))), 1)
    crit_pct = round(max(10, min(100, random.gauss(comp_base - 8, 9))), 1)

    wr_total = 100
    lit_total, inf_total, crit_total = 5, 5, 5
    passage_grade = max(1, student.grade_level - LEVELS_DOWN_BY_PROFILE[profile])

    return Assessment(
        student_id=student.id, school_year=school_year, term=term,
        passage_grade_level=passage_grade,
        word_recognition_correct=round(wr_pct / 100 * wr_total), word_recognition_total=wr_total,
        literal_correct=round(lit_pct / 100 * lit_total), literal_total=lit_total,
        inferential_correct=round(inf_pct / 100 * inf_total), inferential_total=inf_total,
        critical_correct=round(crit_pct / 100 * crit_total), critical_total=crit_total,
        reading_speed_wpm=round(max(20, speed), 1),
        encoded_by=encoder_id,
    )


def seed():
    with app.app_context():
        db.drop_all()
        db.create_all()

        school_names = ["Villasis 1 Central School SPED Center", "San Blas Elementary School",
                         "Piaz Elementary School", "Puelay Elementary School"]
        schools = [School(name=n) for n in school_names]
        db.session.add_all(schools)
        db.session.commit()

        dev = User(username="devadmin", full_name="IT Capstone Developer", role="developer")
        dev.set_password("devadmin123")
        db.session.add(dev)
        db.session.commit()

        profiles_pool = (["strong"] * 2 + ["developing"] * 4 + ["struggling"] * 4)

        all_coordinators = []
        all_teachers = []
        all_students = []

        for school, config in zip(schools, SCHOOL_CONFIGS):
            coord = User(username=config["coordinator_username"], full_name=config["coordinator_name"],
                         role="coordinator", school_id=school.id, email=config["coordinator_email"])
            coord.set_password("coord123")
            db.session.add(coord)
            db.session.commit()
            all_coordinators.append(coord)

            for name, grade, count, t_username, t_name, t_email in config["sections"]:
                sec = Section(name=name, grade_level=grade, school_id=school.id)
                db.session.add(sec)
                db.session.commit()

                t = User(username=t_username, full_name=t_name, role="teacher",
                         school_id=school.id, email=t_email)
                t.set_password("teach123")
                db.session.add(t)
                db.session.flush()
                sec.teacher_id = t.id
                db.session.commit()
                all_teachers.append(t)

                for i in range(count):
                    lrn = f"{grade}{sec.id:02d}{i:03d}{random.randint(10,99)}"
                    first_name = random.choice(FIRST_NAMES)
                    last_name = random.choice(LAST_NAMES)
                    stu = Student(lrn=lrn, first_name=first_name, last_name=last_name,
                                   grade_level=grade, section_id=sec.id,
                                   sex=random.choice(["M", "F"]))
                    db.session.add(stu)
                    db.session.flush()
                    all_students.append((stu, t, random.choice(profiles_pool)))

        db.session.commit()
        print(f"Created {len(all_coordinators)} coordinators, {len(all_teachers)} teachers, "
              f"{len(all_students)} students across {len(schools)} schools.")

        # --- Assessments: GST + Pre + Mid + (Post, unless passed early) ---
        for idx, (stu, teacher, profile) in enumerate(all_students):
            gst = GroupScreeningTest(
                student_id=stu.id, school_year="2025-2026",
                grade_level_at_test=stu.grade_level,
                score=make_gst_score(profile), total_items=20,
                encoded_by=teacher.id,
            )
            db.session.add(gst)

            pre = create_assessment(stu, "pre", "2025-2026", profile, teacher.id)
            db.session.add(pre)
            db.session.flush()

            mid_profile = profile
            if profile == "struggling" and random.random() < 0.4:
                mid_profile = "developing"
            elif profile == "developing" and random.random() < 0.3:
                mid_profile = "strong"
            mid = create_assessment(stu, "mid", "2025-2026", mid_profile, teacher.id)

            overridden_early = (profile == "developing" and mid_profile == "strong" and random.random() < 0.15)
            if overridden_early:
                mid.teacher_override_pass = True
                mid.override_note = ("Reading comfortably above grade level texts during "
                                      "independent reading time; teacher judged intervention "
                                      "no longer necessary.")
                db.session.add(mid)
            else:
                db.session.add(mid)
                db.session.flush()

                post_profile = mid_profile
                if mid_profile == "struggling" and random.random() < 0.35:
                    post_profile = "developing"
                elif mid_profile == "developing" and random.random() < 0.45:
                    post_profile = "strong"
                post = create_assessment(stu, "post", "2025-2026", post_profile, teacher.id)
                db.session.add(post)

            if idx % 50 == 0:
                db.session.commit()
        db.session.commit()

        # --- Sample student login accounts (a few, from the first school) ---
        for idx, (stu, teacher, profile) in enumerate(all_students[:3]):
            su = User(username=f"student{idx+1}", full_name=stu.full_name, role="student",
                      school_id=stu.section.school_id, student_id=stu.id)
            su.set_password("student123")
            db.session.add(su)
        db.session.commit()

        # --- Intervention Library ---
        materials = [
            ("Phonics Power Pack: Consonant Blends", "word_recognition", "Frustration", 4,
             "Structured phonics drills targeting consonant blends and digraphs for learners with low word recognition accuracy."),
            ("Sight Word Bingo Set", "word_recognition", "Frustration", 5,
             "Game-based repeated exposure to high-frequency sight words to build automatic word recognition."),
            ("Decoding Ladder Worksheets", "word_recognition", "Frustration", 6,
             "Scaffolded decoding exercises moving from syllables to multisyllabic words."),
            ("Guided Oral Reading Passages (Level C)", "word_recognition", "Instructional", 4,
             "Leveled passages with teacher-guided repeated reading to build accuracy and automaticity."),
            ("Vocabulary Word Wall Builder", "word_recognition", "Instructional", 5,
             "Weekly vocabulary set with word-recognition games to reinforce grade-level word lists."),
            ("Fluency Timed Reading Cards", "word_recognition", "Instructional", 6,
             "One-minute timed reading cards to strengthen word recognition speed and accuracy."),
            ("Story Map Graphic Organizers", "comprehension", "Frustration", 4,
             "Visual story-mapping tool to help learners identify characters, setting, and basic plot before answering comprehension questions."),
            ("Picture-Supported Short Stories", "comprehension", "Frustration", 5,
             "Illustrated short stories paired with literal-level comprehension questions to rebuild basic understanding."),
            ("Guided Retelling Cards", "comprehension", "Frustration", 6,
             "Sequencing cards that scaffold retelling of a passage to support literal comprehension."),
            ("Inferential Thinking Question Stems", "comprehension", "Instructional", 4,
             "Question stem cards that prompt learners to make inferences from grade-level texts."),
            ("Compare & Contrast Passage Set", "comprehension", "Instructional", 5,
             "Paired passages with graphic organizers to build inferential and critical comprehension."),
            ("Critical Reading Discussion Guide", "comprehension", "Instructional", 6,
             "Discussion-based activities that push learners toward critical/evaluative comprehension of texts."),
            ("Reader's Theater Scripts", "fluency", "Instructional", 4,
             "Short scripts for partner reading to build fluency, expression, and pacing."),
            ("Echo Reading Passages", "fluency", "Frustration", 5,
             "Teacher-led echo reading passages to model pacing and prosody for struggling readers."),
            ("Poetry Fluency Cards", "fluency", "Instructional", 6,
             "Short poems for repeated reading practice to build rate and expression."),
        ]
        for title, skill, level, grade, desc in materials:
            db.session.add(InterventionMaterial(
                title=title, target_skill=skill, reading_level=level,
                grade_level=grade, description=desc))
        db.session.commit()
        # --- External, research-supported resources (Option A: curated,
        # not live-searched) -- all URLs verified against real, currently
        # live pages before seeding. Each is duplicated across grades 4-6
        # since these are general teacher strategy guides, not grade-locked
        # content, so they can surface in any grade's recommendations. ---
        EXTERNAL_RESOURCES = [
            ("Phonics Instruction: The Basics", "word_recognition", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/phonics-and-decoding/articles/phonics-instruction-basics",
             "Summary of National Reading Panel findings on why systematic, explicit phonics instruction "
             "outperforms non-systematic approaches for building word recognition."),
            ("Multisensory Phonics Instruction", "word_recognition", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/topics/curriculum-and-instruction/articles/phonics-instruction-value-multi-sensory-approach",
             "Multisensory phonics techniques (visual, auditory, kinesthetic, tactile) that reinforce "
             "letter-sound correspondence for learners needing extra reinforcement."),
            ("Strategies That Promote Comprehension", "comprehension", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/topics/background-knowledge/articles/strategies-promote-comprehension",
             "Research-based before/during/after reading strategies that help learners coordinate key "
             "comprehension techniques across a lesson."),
            ("Seven Strategies to Teach Text Comprehension", "comprehension", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/comprehension/articles/seven-strategies-teach-students-text-comprehension",
             "Seven evidence-based comprehension strategies -- including monitoring, summarizing, and "
             "questioning -- for learners who need more structured comprehension support."),
            ("Developing Fluent Readers Through Repeated Reading", "fluency", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/fluency/articles/developing-fluent-readers",
             "Explains repeated oral reading with teacher feedback -- the most research-supported "
             "approach for building fluency in struggling readers."),
            ("Timed Repeated Readings for Fluency", "fluency", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/classroom/classroom-strategies/timed-repeated-readings",
             "Step-by-step timed repeated-reading protocol for building reading rate, accuracy, and "
             "expression using familiar instructional-level text."),
            ("DepEd Learning Resource Portal (LRMDS)", "word_recognition", "Instructional", "DepEd LRMDS",
             "https://lrmds.deped.gov.ph/",
             "Official DepEd repository of vetted textbooks, teacher's manuals, and supplementary "
             "reading materials, searchable by grade level and subject."),
        ]
        for title, skill, level, source_name, source_url, desc in EXTERNAL_RESOURCES:
            for grade in (4, 5, 6):
                db.session.add(InterventionMaterial(
                    title=title, target_skill=skill, reading_level=level,
                    grade_level=grade, description=desc,
                    is_external=True, source_name=source_name, source_url=source_url,
                    uploaded_by=dev.id,
                ))
        db.session.commit()
                # --- External, research-supported resources (Option A: curated,
        # not live-searched) -- all URLs verified against real, currently
        # live pages before seeding. Each is duplicated across grades 4-6
        # since these are general teacher strategy guides, not grade-locked
        # content, so they can surface in any grade's recommendations. ---
        EXTERNAL_RESOURCES = [
            ("Phonics Instruction: The Basics", "word_recognition", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/phonics-and-decoding/articles/phonics-instruction-basics",
             "Summary of National Reading Panel findings on why systematic, explicit phonics instruction "
             "outperforms non-systematic approaches for building word recognition."),
            ("Multisensory Phonics Instruction", "word_recognition", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/topics/curriculum-and-instruction/articles/phonics-instruction-value-multi-sensory-approach",
             "Multisensory phonics techniques (visual, auditory, kinesthetic, tactile) that reinforce "
             "letter-sound correspondence for learners needing extra reinforcement."),
            ("Strategies That Promote Comprehension", "comprehension", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/topics/background-knowledge/articles/strategies-promote-comprehension",
             "Research-based before/during/after reading strategies that help learners coordinate key "
             "comprehension techniques across a lesson."),
            ("Seven Strategies to Teach Text Comprehension", "comprehension", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/comprehension/articles/seven-strategies-teach-students-text-comprehension",
             "Seven evidence-based comprehension strategies -- including monitoring, summarizing, and "
             "questioning -- for learners who need more structured comprehension support."),
            ("Developing Fluent Readers Through Repeated Reading", "fluency", "Frustration", "Reading Rockets",
             "https://www.readingrockets.org/topics/fluency/articles/developing-fluent-readers",
             "Explains repeated oral reading with teacher feedback -- the most research-supported "
             "approach for building fluency in struggling readers."),
            ("Timed Repeated Readings for Fluency", "fluency", "Instructional", "Reading Rockets",
             "https://www.readingrockets.org/classroom/classroom-strategies/timed-repeated-readings",
             "Step-by-step timed repeated-reading protocol for building reading rate, accuracy, and "
             "expression using familiar instructional-level text."),
            ("DepEd Learning Resource Portal (LRMDS)", "word_recognition", "Instructional", "DepEd LRMDS",
             "https://lrmds.deped.gov.ph/",
             "Official DepEd repository of vetted textbooks, teacher's manuals, and supplementary "
             "reading materials, searchable by grade level and subject."),
        ]
        for title, skill, level, source_name, source_url, desc in EXTERNAL_RESOURCES:
            for grade in (4, 5, 6):
                db.session.add(InterventionMaterial(
                    title=title, target_skill=skill, reading_level=level,
                    grade_level=grade, description=desc,
                    is_external=True, source_name=source_name, source_url=source_url,
                    uploaded_by=dev.id,
                ))
        db.session.commit()
        seeded_materials = InterventionMaterial.query.all()
        for idx, (stu, teacher, profile) in enumerate(all_students[:3]):
            if profile != "strong" and seeded_materials:
                candidate = next((m for m in seeded_materials if m.grade_level == stu.grade_level),
                                  seeded_materials[0])
                db.session.add(MaterialAssignment(
                    material_id=candidate.id, student_id=stu.id, assigned_by=teacher.id,
                ))
        db.session.commit()

        for idx, (stu, teacher, profile) in enumerate(all_students[:3]):
            if profile != "strong" and stu.assessments:
                latest = stu.assessments[-1]
                confidence = latest.diagnostic_confidence
                existing_assignment = MaterialAssignment.query.filter_by(student_id=stu.id).first()
                db.session.add(InterventionRecord(
                    student_id=stu.id, assessment_id=latest.id,
                    weakest_skill=latest.weakest_skill,
                    confidence_score=confidence["score"] if confidence else None,
                    confidence_label=confidence["label"] if confidence else None,
                    decision="accepted",
                    material_id=existing_assignment.material_id if existing_assignment else None,
                    note="Seeded demo record.",
                    recorded_by=teacher.id,
                ))
        db.session.commit()

        db.session.add_all([
            SystemSetting(key="current_school_year", value="2025-2026", updated_by=dev.id),
            SystemSetting(key="system_display_name", value="Reading-Ready", updated_by=dev.id),
            SystemSetting(key="support_contact_email", value="", updated_by=dev.id),
        ])
        db.session.commit()

        # --- Audit logs: signup/approval history for every coordinator ---
        now = datetime.utcnow()
        for coord in all_coordinators:
            req_ts = now - timedelta(days=random.randint(200, 260))
            appr_ts = req_ts + timedelta(hours=random.randint(2, 48))
            db.session.add(AuditLog(
                user_id=None, school_id=coord.school_id, action="signup_submitted",
                description=f"{coord.full_name} ({coord.username}) requested a coordinator account for {coord.school.name}.",
                ip_address="127.0.0.1", created_at=req_ts,
            ))
            db.session.add(AuditLog(
                user_id=dev.id, school_id=coord.school_id, action="account_approved",
                description=f"{dev.full_name} approved coordinator account '{coord.username}'.",
                ip_address="127.0.0.1", created_at=appr_ts,
            ))
        db.session.commit()

        # --- Audit logs: a spread of logins across ~8 months for every staff
        # member (developer, coordinators, teachers) -- gives the System
        # Admin login-activity chart real month-to-month shape to display. ---
        login_logs = []
        all_staff = [dev] + all_coordinators + all_teachers
        for person in all_staff:
            num_logins = random.randint(15, 40)
            for _ in range(num_logins):
                days_ago = random.randint(0, 240)
                ts = now - timedelta(days=days_ago, hours=random.randint(0, 23), minutes=random.randint(0, 59))
                login_logs.append(AuditLog(
                    user_id=person.id, school_id=person.school_id,
                    action="login_success", description=f"{person.full_name} logged in.",
                    ip_address="127.0.0.1", created_at=ts,
                ))

        for _ in range(60):
            days_ago = random.randint(0, 240)
            ts = now - timedelta(days=days_ago, hours=random.randint(0, 23))
            login_logs.append(AuditLog(
                user_id=None, school_id=None, action="login_failed",
                description=f"Failed login attempt for username 'user{random.randint(100,999)}'.",
                ip_address="127.0.0.1", created_at=ts,
            ))

        db.session.add_all(login_logs)
        db.session.commit()

        # --- A handful of password-changed and student-added activity logs,
        # spread across staff and time, to give the teacher activity log and
        # audit log some realistic variety beyond just logins. ---
        extra_logs = []
        for person in random.sample(all_teachers, min(8, len(all_teachers))):
            ts = now - timedelta(days=random.randint(1, 200))
            extra_logs.append(AuditLog(
                user_id=person.id, school_id=person.school_id, action="password_changed",
                description=f"{person.full_name} changed their password.",
                ip_address="127.0.0.1", created_at=ts,
            ))
        for _ in range(30):
            stu, teacher, _profile = random.choice(all_students)
            ts = now - timedelta(days=random.randint(0, 230), hours=random.randint(0, 23))
            extra_logs.append(AuditLog(
                user_id=teacher.id, school_id=teacher.school_id, action="student_added",
                description=f"{teacher.full_name} added learner '{stu.full_name}' (LRN {stu.lrn}).",
                ip_address="127.0.0.1", created_at=ts,
            ))
        for _ in range(40):
            stu, teacher, _profile = random.choice(all_students)
            term = random.choice(["pre", "mid", "post"])
            ts = now - timedelta(days=random.randint(0, 230), hours=random.randint(0, 23))
            extra_logs.append(AuditLog(
                user_id=teacher.id, school_id=teacher.school_id, action="assessment_encoded",
                description=f"{teacher.full_name} encoded the {term.title()}-test for {stu.full_name}.",
                ip_address="127.0.0.1", created_at=ts,
            ))
        db.session.add_all(extra_logs)
        db.session.commit()

        print("Seed complete.")
        print(f"Schools: {School.query.count()}")
        print(f"Sections: {Section.query.count()}")
        print(f"Users: {User.query.count()}")
        print(f"Students: {Student.query.count()}")
        print(f"Assessments: {Assessment.query.count()}")
        print(f"Group Screening Tests: {GroupScreeningTest.query.count()}")
        print(f"Intervention materials: {InterventionMaterial.query.count()}")
        print(f"Material assignments: {MaterialAssignment.query.count()}")
        print(f"Intervention records: {InterventionRecord.query.count()}")
        print(f"System settings: {SystemSetting.query.count()}")
        print(f"Audit log entries: {AuditLog.query.count()}")


if __name__ == "__main__":
    seed()