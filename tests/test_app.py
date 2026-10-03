import io
import os
import re
import hashlib
import sys
import sqlite3
import tempfile
import unittest
import openpyxl
from flask import render_template
from werkzeug.security import generate_password_hash

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app import create_app, calculate_validity_minutes, utcnow_iso
from db import init_db
from config import Config


class LabPortalTestCase(unittest.TestCase):
    def setUp(self):
        self.db_fd, self.db_path = tempfile.mkstemp()

        class TestConfig(Config):
            SECRET_KEY = "test"
            UPLOAD_FOLDER = tempfile.mkdtemp()
            ALLOWED_EXTENSIONS = {"pdf"}
            LAB_KEY_DEFAULT_MINUTES = 5
            CSRF_ENABLED = False

        self.app = create_app(TestConfig, db_path=self.db_path)
        self.app.config["TESTING"] = True

        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO course (name, year) VALUES (?, ?)", ("Data Structures", "1Year"))
        conn.commit()
        conn.close()

        self.client = self.app.test_client()

    def tearDown(self):
        os.close(self.db_fd)
        os.unlink(self.db_path)

    def _register(
        self, name, email, password, course_id="1", student_id="S-1000",
        mobile_no="9999999999", section="A", year="1Year",
    ):
        return self.client.post(
            "/register",
            data={
                "name": name,
                "email": email,
                "password": password,
                "course_id": course_id,
                "student_id": student_id,
                "mobile_no": mobile_no,
                "section": section,
                "year": year,
            },
            follow_redirects=True,
        )

    def _login(self, email, password):
        return self.client.post(
            "/login", data={"email": email, "password": password}, follow_redirects=True
        )

    def test_calculate_validity_minutes_returns_duration_in_minutes(self):
        start = "2024-01-01 09:00:00"
        end = "2024-01-01 11:30:00"
        self.assertEqual(calculate_validity_minutes(start, end), 150)
        self.assertEqual(calculate_validity_minutes(end, start), 0)
        self.assertEqual(calculate_validity_minutes(None, end), 0)

    def test_register_and_login(self):
        registration_page = self.client.get("/register")
        self.assertIn(b'class="brand-logo"', registration_page.data)
        self.assertIn(b"/static/logo.jpg", registration_page.data)
        background_image = self.client.get("/static/campus-background.jpg")
        self.assertEqual(background_image.status_code, 200)
        self.assertEqual(background_image.mimetype, "image/jpeg")
        background_image.close()
        self.assertIn(b'<select id="section" name="section" required>', registration_page.data)
        self.assertNotIn(b'name="year"', registration_page.data)
        self.assertIn(b"Data Structures 1 YEAR", registration_page.data)
        resp = self._register("Test Student", "student@test.com", "pass1234", course_id="1")
        self.assertEqual(resp.status_code, 200)
        conn = sqlite3.connect(self.db_path)
        saved_section, saved_year = conn.execute(
            "SELECT section, year FROM user WHERE email = ?", ("student@test.com",)
        ).fetchone()
        conn.close()
        self.assertEqual((saved_section, saved_year), ("A", "1Year"))

        resp = self.client.post(
            "/login",
            data={"email": "student@test.com", "password": "pass1234", "course_id": "1"},
            follow_redirects=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Welcome", resp.data)
        self.assertIn(b"Enter Lab Key", resp.data)
        self.assertIn(b"dashboard-lab-code", resp.data)
        self.assertIn(b"Course Documents", resp.data)

    def test_inactive_courses_are_hidden_from_auth_views(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("UPDATE course SET status = 'inactive' WHERE id = 1")
        conn.execute(
            "INSERT INTO course (name, year, status) VALUES (?, ?, ?)",
            ("Visible Course", "2Year", "active"),
        )
        conn.commit()
        conn.close()

        login_page = self.client.get("/login")
        register_page = self.client.get("/register")

        self.assertNotIn(b"Data Structures", login_page.data)
        self.assertNotIn(b"Data Structures", register_page.data)
        self.assertIn(b"Visible Course", login_page.data)
        self.assertIn(b"Visible Course", register_page.data)

    def test_email_column_is_varchar_255_and_unique(self):
        conn = sqlite3.connect(self.db_path)
        email_column = next(row for row in conn.execute("PRAGMA table_info(user)") if row[1] == "email")
        mobile_column = next(row for row in conn.execute("PRAGMA table_info(user)") if row[1] == "mobile_no")
        email_indexes = conn.execute("PRAGMA index_list(user)").fetchall()
        conn.close()

        self.assertEqual(email_column[2].upper(), "VARCHAR(255)")
        self.assertEqual(mobile_column[2].upper(), "VARCHAR(10)")
        self.assertTrue(any(index[2] for index in email_indexes))

    def test_course_migration_removes_section_and_preserves_references(self):
        legacy_fd, legacy_path = tempfile.mkstemp()
        os.close(legacy_fd)
        try:
            conn = sqlite3.connect(legacy_path)
            conn.executescript(
                "CREATE TABLE course (id INTEGER PRIMARY KEY, name TEXT NOT NULL, "
                "section TEXT NOT NULL UNIQUE, subject TEXT NOT NULL DEFAULT '', year TEXT NOT NULL DEFAULT '2Year'); "
                "CREATE TABLE course_reference (course_id INTEGER REFERENCES course(id)); "
                "INSERT INTO course (id, name, section, subject, year) VALUES (9, 'BCA', 'A', 'Math', '2Year'); "
                "INSERT INTO course_reference (course_id) VALUES (9);"
            )
            conn.close()

            init_db(legacy_path, {"DB_ENGINE": "sqlite"})

            conn = sqlite3.connect(legacy_path)
            columns = {row[1] for row in conn.execute("PRAGMA table_info(course)")}
            self.assertNotIn("section", columns)
            self.assertTrue(
                {"name", "year", "course_section", "status", "created_at", "updated_at"}.issubset(columns)
            )
            self.assertNotIn("duration_years", columns)
            self.assertNotIn("is_active", columns)
            self.assertEqual(
                conn.execute("SELECT id, name, subject, year FROM course").fetchone(),
                (9, "BCA", "Math", "2Year"),
            )
            self.assertEqual(conn.execute("SELECT course_section FROM course").fetchone(), ("A",))
            self.assertEqual(conn.execute("SELECT course_id FROM course_reference").fetchone(), (9,))
            self.assertEqual(conn.execute("PRAGMA foreign_key_check").fetchall(), [])
            conn.close()
        finally:
            os.unlink(legacy_path)

    def test_course_section_migration_normalizes_existing_values(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("UPDATE course SET course_section = '7' WHERE id = 1")
        conn.execute(
            "INSERT INTO course (name, year, course_section) VALUES (?, ?, ?)",
            ("Second Course", "2Year", "b"),
        )
        conn.execute(
            "INSERT INTO course (name, year, course_section) VALUES (?, ?, ?)",
            ("Third Course", "3Year", ""),
        )
        conn.commit()
        conn.close()

        init_db(self.db_path, self.app.config)

        conn = sqlite3.connect(self.db_path)
        sections = conn.execute(
            "SELECT course_section FROM course ORDER BY id"
        ).fetchall()
        conn.close()
        self.assertEqual(sections, [("G",), ("B",), ("A",)])

    def test_existing_course_section_column_migrates_to_varchar_five(self):
        legacy_fd, legacy_path = tempfile.mkstemp()
        os.close(legacy_fd)
        try:
            conn = sqlite3.connect(legacy_path)
            conn.execute(
                "CREATE TABLE course (id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "name TEXT NOT NULL, subject TEXT NOT NULL DEFAULT '', "
                "year TEXT NOT NULL DEFAULT '', duration_years INTEGER NOT NULL DEFAULT 4, "
                "course_section TEXT NOT NULL DEFAULT 'A', is_active INTEGER NOT NULL DEFAULT 1)"
            )
            conn.execute(
                "INSERT INTO course (id, name, course_section) VALUES (?, ?, ?)",
                (5, "Existing Course", "C"),
            )
            conn.commit()
            conn.close()

            init_db(legacy_path, {"DB_ENGINE": "sqlite"})

            conn = sqlite3.connect(legacy_path)
            section_column = next(
                row for row in conn.execute("PRAGMA table_info(course)")
                if row[1] == "course_section"
            )
            name_column = next(
                row for row in conn.execute("PRAGMA table_info(course)")
                if row[1] == "name"
            )
            self.assertEqual(section_column[2].upper(), "VARCHAR(10)")
            self.assertEqual(name_column[2].upper(), "VARCHAR(100)")
            self.assertEqual(
                conn.execute("SELECT id, course_section FROM course").fetchone(),
                (5, "C"),
            )
            self.assertNotIn(
                "duration_years",
                {row[1] for row in conn.execute("PRAGMA table_info(course)")},
            )
            self.assertNotIn(
                "is_active",
                {row[1] for row in conn.execute("PRAGMA table_info(course)")},
            )
            self.assertEqual(conn.execute("PRAGMA foreign_key_check").fetchall(), [])
            conn.close()
        finally:
            os.unlink(legacy_path)

    def test_legacy_duration_value_migrates_to_course_section(self):
        legacy_fd, legacy_path = tempfile.mkstemp()
        os.close(legacy_fd)
        try:
            conn = sqlite3.connect(legacy_path)
            conn.execute(
                "CREATE TABLE course (id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "name TEXT NOT NULL, duration_years INTEGER NOT NULL DEFAULT 4, "
                "is_active INTEGER NOT NULL DEFAULT 1)"
            )
            conn.execute(
                "INSERT INTO course (id, name, duration_years, is_active) VALUES (?, ?, ?, ?)",
                (8, "Legacy Course", 6, 0),
            )
            conn.commit()
            conn.close()

            init_db(legacy_path, {"DB_ENGINE": "sqlite"})

            conn = sqlite3.connect(legacy_path)
            columns = {row[1] for row in conn.execute("PRAGMA table_info(course)")}
            self.assertNotIn("duration_years", columns)
            self.assertEqual(
                conn.execute("SELECT id, course_section, status FROM course").fetchone(),
                (8, "F", "inactive"),
            )
            self.assertNotIn("is_active", columns)
            self.assertEqual(conn.execute("PRAGMA foreign_key_check").fetchall(), [])
            conn.close()
        finally:
            os.unlink(legacy_path)

    def test_student_login_rejects_wrong_course(self):
        self._register("Course Student", "course@test.com", "pass1234", course_id="1")
        resp = self.client.post(
            "/login",
            data={"email": "course@test.com", "password": "pass1234", "course_id": "999"},
            follow_redirects=True,
        )
        self.assertIn(b"Select your registered course", resp.data)

    def test_login_replaces_previous_device_session(self):
        self._register("Single Device Student", "single-device@test.com", "pass1234")
        first_client = self.client
        second_client = self.app.test_client()

        first_login = first_client.post(
            "/login", data={"email": "single-device@test.com", "password": "pass1234"}
        )
        self.assertEqual(first_login.status_code, 302)
        with first_client.session_transaction() as browser_session:
            first_token = browser_session["auth_token"]

        replacement = second_client.post(
            "/login",
            data={"email": "single-device@test.com", "password": "pass1234"},
        )
        self.assertEqual(replacement.status_code, 302)
        with second_client.session_transaction() as browser_session:
            replacement_token = browser_session["auth_token"]
        self.assertNotEqual(first_token, replacement_token)
        self.assertEqual(first_client.get("/dashboard").status_code, 302)
        self.assertIn(b"Welcome", second_client.get("/dashboard").data)

        conn = sqlite3.connect(self.db_path)
        session_statuses = conn.execute(
            "SELECT status, token_hash FROM auth_session WHERE user_id = (SELECT id FROM user WHERE email = ?) ORDER BY id",
            ("single-device@test.com",),
        ).fetchall()
        conn.close()
        self.assertEqual(session_statuses[0][0], "LOGGED_OUT")
        self.assertEqual(session_statuses[0][1], hashlib.sha256(first_token.encode()).hexdigest())
        self.assertEqual(session_statuses[1][0], "ACTIVE")
        self.assertEqual(session_statuses[1][1], hashlib.sha256(replacement_token.encode()).hexdigest())

    def test_admin_login_replaces_previous_device_session(self):
        first_admin = self.app.test_client()
        second_admin = self.app.test_client()

        first_login = first_admin.post(
            "/login", data={"email": "admin@ayu.com", "password": "admin123"}
        )
        self.assertEqual(first_login.status_code, 302)
        with first_admin.session_transaction() as browser_session:
            first_token = browser_session["auth_token"]

        second_login = second_admin.post(
            "/login", data={"email": "admin@ayu.com", "password": "admin123"}
        )
        self.assertEqual(second_login.status_code, 302)
        with second_admin.session_transaction() as browser_session:
            second_token = browser_session["auth_token"]

        self.assertNotEqual(first_token, second_token)
        self.assertEqual(first_admin.get("/admin").status_code, 302)
        self.assertIn(b"Admin Panel", second_admin.get("/admin").data)

    def test_admin_login_activity_tracks_device_and_logout(self):
        mobile_ua = (
            "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36"
        )
        self.client.post(
            "/login",
            data={"email": "admin@ayu.com", "password": "wrong-password"},
            headers={"User-Agent": mobile_ua},
        )
        login = self.client.post(
            "/login",
            data={"email": "admin@ayu.com", "password": "admin123"},
            headers={"User-Agent": mobile_ua},
        )
        self.assertEqual(login.status_code, 302)
        with self.client.session_transaction() as browser_session:
            raw_token = browser_session["auth_token"]

        self._register(
            "Login Activity Student",
            "login-activity-student@test.com",
            "pass1234",
            course_id="1",
            student_id="S-LOGIN-1",
            section="B",
            year="2Year",
        )
        student_client = self.app.test_client()
        student_login = student_client.post(
            "/login",
            data={"email": "login-activity-student@test.com", "password": "pass1234"},
        )
        self.assertEqual(student_login.status_code, 302)

        dashboard = self.client.get("/admin")
        self.assertNotIn(b"Active Sessions", dashboard.data)
        activity_page = self.client.get("/admin/login-activity")
        self.assertEqual(activity_page.status_code, 200)
        self.assertIn(b"Login Activity Student", activity_page.data)
        self.assertIn(b"S-LOGIN-1", activity_page.data)
        self.assertIn(b"Data Structures", activity_page.data)
        self.assertNotIn(b"Section", activity_page.data)
        self.assertNotIn(b"Year", activity_page.data)
        self.assertIn(b"ACTIVE", activity_page.data)
        self.assertNotIn(b"admin@ayu.com", activity_page.data)
        self.assertIn(b'<option value="lock">Lock</option>', activity_page.data)
        self.assertIn(b'<option value="unlock">Unlock</option>', activity_page.data)
        self.assertIn(b'<option value="delete">Delete</option>', activity_page.data)

        student_client.post("/logout")
        activity_page = self.client.get("/admin/login-activity")
        self.assertIn(b"LOGGED_OUT", activity_page.data)

        conn = sqlite3.connect(self.db_path)
        records = conn.execute(
            "SELECT device_type, browser, operating_system, ip_address, session_id, login_status "
            "FROM login_activity ORDER BY id"
        ).fetchall()
        conn.close()
        self.assertEqual(records[0][5], "Failed")
        self.assertEqual(records[1][:4], ("Mobile", "Chrome", "Android", "127.0.0.1"))
        self.assertEqual(records[1][4], hashlib.sha256(raw_token.encode()).hexdigest())
        self.assertNotEqual(records[1][4], raw_token)

        self.client.post("/logout")
        conn = sqlite3.connect(self.db_path)
        logout = conn.execute(
            "SELECT login_status, logout_time FROM login_activity WHERE session_id = ?",
            (records[1][4],),
        ).fetchone()
        conn.close()
        self.assertEqual(logout[0], "Logged out")
        self.assertIsNotNone(logout[1])

    def test_admin_login_allows_mobile_tablet_and_desktop(self):
        user_agents = (
            (
                "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36",
                "Mobile",
            ),
            (
                "Mozilla/5.0 (iPad; CPU OS 17_0 like Mac OS X) AppleWebKit/605.1.15 "
                "(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
                "Tablet",
            ),
            (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Desktop/Laptop",
            ),
        )
        for user_agent, expected_type in user_agents:
            client = self.app.test_client()
            response = client.post(
                "/login",
                data={"email": "admin@ayu.com", "password": "admin123"},
                headers={"User-Agent": user_agent},
            )
            self.assertEqual(response.status_code, 302)

        conn = sqlite3.connect(self.db_path)
        device_types = [row[0] for row in conn.execute(
            "SELECT device_type FROM login_activity ORDER BY id"
        ).fetchall()]
        conn.close()
        self.assertEqual(device_types, [expected for _, expected in user_agents])

    def test_admin_can_force_logout_another_admin_session(self):
        admin_client = self.app.test_client()
        super_admin_client = self.app.test_client()
        admin_client.post("/login", data={"email": "admin@ayu.com", "password": "admin123"})
        super_admin_client.post(
            "/login", data={"email": "ayush5075mi@gmail.com", "password": "india123"}
        )

        conn = sqlite3.connect(self.db_path)
        session_id = conn.execute(
            "SELECT auth_session.id FROM auth_session JOIN user ON user.id = auth_session.user_id "
            "WHERE user.email = ? AND auth_session.status = 'ACTIVE'",
            ("admin@ayu.com",),
        ).fetchone()[0]
        conn.close()
        forced = super_admin_client.post(f"/admin/sessions/{session_id}/logout")
        self.assertEqual(forced.status_code, 302)
        self.assertEqual(admin_client.get("/admin").status_code, 302)

        conn = sqlite3.connect(self.db_path)
        status = conn.execute(
            "SELECT login_status, logout_time FROM login_activity WHERE admin_id = "
            "(SELECT id FROM user WHERE email = 'admin@ayu.com')"
        ).fetchone()
        conn.close()
        self.assertEqual(status[0], "Force logged out")
        self.assertIsNotNone(status[1])

    def test_csrf_is_required_for_login_and_logout(self):
        self.app.config["CSRF_ENABLED"] = True
        client = self.app.test_client()
        rejected = client.post("/login", data={"email": "admin@ayu.com", "password": "admin123"})
        self.assertEqual(rejected.status_code, 400)

        client.get("/login")
        with client.session_transaction() as browser_session:
            csrf_token = browser_session["_csrf_token"]
        login = client.post(
            "/login",
            data={"email": "admin@ayu.com", "password": "admin123", "csrf_token": csrf_token},
        )
        self.assertEqual(login.status_code, 302)
        self.assertEqual(client.get("/logout").status_code, 405)
        client.get("/admin")
        with client.session_transaction() as browser_session:
            csrf_token = browser_session["_csrf_token"]
        logout = client.post("/logout", data={"csrf_token": csrf_token})
        self.assertEqual(logout.status_code, 302)

    def test_missing_page_renders_custom_404_template(self):
        response = self.client.get("/does-not-exist")
        self.assertEqual(response.status_code, 404)
        self.assertIn(b"Page Not Found", response.data)

    def test_admin_attendance_filters_and_exports_student_id(self):
        self._register("Selected Student", "selected@test.com", "pass1234", student_id="S-1001", section="B")
        self._register("Other Student", "other@test.com", "pass1234", student_id="S-1002")
        self._login("admin@ayu.com", "admin123")

        overview = self.client.get("/admin/attendance?section=B&student_id=S-1001&month=2026-09")
        self.assertIn(b"value=\"S-1001\"", overview.data)
        self.assertIn(b'<option value="B" selected>B</option>', overview.data)
        self.assertIn(b"Selected Student", overview.data)
        self.assertNotIn(b"Other Student", overview.data)
        self.assertIn(b"student_id=S-1001", overview.data)
        self.assertIn(b"section=B", overview.data)

        exported = self.client.get(
            "/admin/attendance/export?view=attendance&section=B&student_id=S-1001&month=2026-09"
        )
        workbook = openpyxl.load_workbook(io.BytesIO(exported.data), read_only=True)
        exported_names = [row[2] for row in list(workbook.active.values)[1:]]
        self.assertEqual(exported_names, ["Selected Student"])

    def test_admin_can_force_logout_student_device(self):
        self._register("Force Logout Student", "force-logout@test.com", "pass1234")
        student_client = self.app.test_client()
        student_client.post(
            "/login", data={"email": "force-logout@test.com", "password": "pass1234"}
        )

        admin_login = self._login("admin@ayu.com", "admin123")
        self.assertIn(b"Admin Panel", admin_login.data)
        conn = sqlite3.connect(self.db_path)
        student_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("force-logout@test.com",)
        ).fetchone()[0]
        conn.close()

        forced = self.client.post(
            f"/admin/students/{student_id}/logout-device", follow_redirects=True
        )
        self.assertIn(b"Student device session logged out", forced.data)
        student_dashboard = student_client.get("/dashboard")
        self.assertEqual(student_dashboard.status_code, 302)
        self.assertIn("/login", student_dashboard.headers["Location"])

        replacement = self.app.test_client().post(
            "/login",
            data={"email": "force-logout@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        self.assertIn(b"Welcome", replacement.data)

    def test_login_course_is_hidden_for_admin(self):
        resp = self.client.get("/login")
        self.assertIn(b'name="course_id"', resp.data)

        resp = self.client.post(
            "/login",
            data={"email": "admin@ayu.com", "password": "admin123"},
            follow_redirects=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Admin Panel", resp.data)

        resp = self.client.post(
            "/login",
            data={"email": "admin@ayu.com", "password": "wrong"},
        )
        self.assertIn(b'value="admin@ayu.com"', resp.data)
        self.assertIn(b'id="course-field" hidden', resp.data)

        super_admin_resp = self.client.post(
            "/login",
            data={"email": "ayush5075mi@gmail.com", "password": "wrong"},
        )
        self.assertIn(b'value="ayush5075mi@gmail.com"', super_admin_resp.data)
        self.assertIn(b'id="course-field" hidden', super_admin_resp.data)

        for admin_number in range(1, 6):
            email = f"admin{admin_number}@ayu.com"
            admin_resp = self.client.post(
                "/login", data={"email": email, "password": "wrong"}
            )
            self.assertIn(f'value="{email}"'.encode(), admin_resp.data)
            self.assertIn(b'id="course-field" hidden', admin_resp.data)

    def test_wrong_password_rejected(self):
        self._register("Student", "wrongpw@test.com", "pass1234")
        resp = self._login("wrongpw@test.com", "wrongpass")
        self.assertIn(b"Invalid email or password", resp.data)

    def test_student_is_deactivated_after_four_wrong_passwords(self):
        self._register("Locked Student", "locked@test.com", "pass1234")

        for attempt in range(3):
            response = self._login("locked@test.com", "wrongpass")
            self.assertIn(b"Invalid email or password", response.data)

        conn = sqlite3.connect(self.db_path)
        status = conn.execute(
            "SELECT is_active, failed_login_attempts FROM user WHERE email = ?",
            ("locked@test.com",),
        ).fetchone()
        conn.close()
        self.assertEqual(status, (1, 3))

        response = self._login("locked@test.com", "wrongpass")
        self.assertIn(b"deactivated after 4 incorrect password attempts", response.data)

        conn = sqlite3.connect(self.db_path)
        status = conn.execute(
            "SELECT is_active, failed_login_attempts FROM user WHERE email = ?",
            ("locked@test.com",),
        ).fetchone()
        conn.close()
        self.assertEqual(status, (0, 4))

        response = self._login("locked@test.com", "pass1234")
        self.assertIn(b"Your account is deactivated", response.data)

    def test_student_forgot_password_updates_password(self):
        self._register("Reset Student", "reset@test.com", "pass1234")

        resp = self.client.post(
            "/forgot-password",
            data={
                "email": "reset@test.com",
                "old_password": "pass1234",
                "new_password": "newpass123",
                "confirm_password": "newpass123",
            },
            follow_redirects=True,
        )

        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Password updated successfully", resp.data)

        login = self._login("reset@test.com", "newpass123")
        self.assertEqual(login.status_code, 200)
        self.assertIn(b"Welcome", login.data)

    def test_admin_route_blocked_for_student(self):
        self._register("Test Student2", "student2@test.com", "pass1234")
        self._login("student2@test.com", "pass1234")
        resp = self.client.get("/admin")
        self.assertEqual(resp.status_code, 403)

    def test_student_attendance_page_is_available(self):
        self._register("Attendance Student", "attendance@test.com", "pass1234", course_id="1")
        self._login("attendance@test.com", "pass1234")

        resp = self.client.get("/attendance")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"STUDENT COMPUTER LAB", resp.data)
        self.assertIn(b"Data Structures", resp.data)

        dashboard = self.client.get("/dashboard")
        self.assertIn(b'href="/attendance"', dashboard.data)

    def test_default_admin_login_and_panel(self):
        resp = self._login("admin@ayu.com", "admin123")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Admin Panel", resp.data)
        self.assertIn(b'name="minutes" value="5"', resp.data)
        self.assertIn(b"Students", resp.data)
        self.assertIn(b"Password", resp.data)

    def test_existing_default_admin_email_migrates(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "UPDATE user SET email = ? WHERE email = ?",
            ("admin@example.com", "admin@ayu.com"),
        )
        conn.commit()
        conn.close()

        init_db(self.db_path, self.app.config)

        conn = sqlite3.connect(self.db_path)
        email = conn.execute(
            "SELECT email FROM user WHERE role = 'admin' AND name = 'Admin'"
        ).fetchone()[0]
        conn.close()
        self.assertEqual(email, "admin@ayu.com")
        self.assertIn(b"Admin Panel", self._login("admin@ayu.com", "admin123").data)

    def test_super_admin_login_and_password_cannot_be_reset(self):
        resp = self._login("ayush5075mi@gmail.com", "india123")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Admin Panel", resp.data)

        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "UPDATE user SET role = 'student' WHERE email = ?",
            ("ayush5075mi@gmail.com",),
        )
        conn.commit()
        conn.close()
        self.client.post("/logout")
        self.assertIn(b"Admin Panel", self._login("ayush5075mi@gmail.com", "india123").data)

        reset = self.client.post(
            "/admin/forgot-password",
            data={
                "email": "ayush5075mi@gmail.com",
                "new_password": "newpassword123",
                "confirm_password": "newpassword123",
            },
            follow_redirects=True,
        )
        self.assertIn(b"super admin password cannot be reset", reset.data)
        self.assertIn(b"Admin Panel", self._login("ayush5075mi@gmail.com", "india123").data)

    def test_super_admin_can_manage_admin_permissions_and_block_inactive_admins(self):
        self._login("ayush5075mi@gmail.com", "india123")

        created = self.client.post(
            "/admin/permissions",
            data={
                "action": "create_admin",
                "name": "Support Admin",
                "email": "support-admin@test.com",
                "password": "pass1234",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Admin account created", created.data)

        page = self.client.get("/admin/permissions")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Admin Permission Management", page.data)
        self.assertIn(b"Support Admin", page.data)
        self.assertNotIn(b"Activity / Audit Log", page.data)
        self.assertIn(b'href="/admin/permissions?view_id=1#permission-editor"', page.data)
        self.assertIn(b'href="/admin/permissions?edit_id=1#permission-editor"', page.data)

        normal_admin = self.app.test_client()
        login_resp = normal_admin.post(
            "/login",
            data={"email": "support-admin@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        self.assertIn(b"Admin Panel", login_resp.data)
        self.assertEqual(normal_admin.get("/admin/permissions").status_code, 403)

        conn = sqlite3.connect(self.db_path)
        support_admin_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("support-admin@test.com",)
        ).fetchone()[0]
        conn.close()
        self.client.post(
            "/admin/permissions",
            data={
                "action": "save_permissions",
                "admin_id": str(support_admin_id),
                "perm_dashboard_view": "on",
                "perm_course_management_view": "on",
            },
        )
        permitted_admin_page = normal_admin.get("/admin")
        self.assertEqual(permitted_admin_page.status_code, 200)
        nav_links = permitted_admin_page.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertIn(b'href="/admin/courses"', nav_links)
        self.assertNotIn(b'href="/admin/subjects"', nav_links)
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 200)
        self.assertEqual(normal_admin.get("/admin/subjects").status_code, 403)
        self.assertEqual(normal_admin.get("/admin/students").status_code, 403)

        conn = sqlite3.connect(self.db_path)
        audit_count = conn.execute(
            "SELECT COUNT(*) FROM admin_permission_audit WHERE target_admin_id = ?",
            (support_admin_id,),
        ).fetchone()[0]
        conn.close()
        self.assertGreater(audit_count, 0)

        self.client.post(
            "/admin/permissions",
            data={
                "action": "set_status",
                "admin_id": str(support_admin_id),
                "account_status": "inactive",
            },
        )
        conn = sqlite3.connect(self.db_path)
        status, = conn.execute(
            "SELECT account_status FROM user WHERE id = ?", (support_admin_id,)
        ).fetchone()
        session_status = conn.execute(
            "SELECT status FROM auth_session WHERE user_id = ? ORDER BY id DESC LIMIT 1",
            (support_admin_id,),
        ).fetchone()[0]
        status_audit_count = conn.execute(
            "SELECT COUNT(*) FROM admin_permission_audit WHERE target_admin_id = ? AND permission_name = 'account_status'",
            (support_admin_id,),
        ).fetchone()[0]
        conn.close()
        self.assertEqual(status, "inactive")
        self.assertEqual(session_status, "FORCE_LOGOUT")
        self.assertEqual(status_audit_count, 1)

        blocked = normal_admin.post(
            "/login",
            data={"email": "support-admin@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        self.assertIn(b"inactive", blocked.data.lower())

        deleted = self.client.post(
            "/admin/permissions",
            data={"action": "delete_admin", "admin_id": str(support_admin_id)},
            follow_redirects=True,
        )
        self.assertIn(b"Admin account deleted.", deleted.data)
        conn = sqlite3.connect(self.db_path)
        self.assertIsNone(
            conn.execute("SELECT id FROM user WHERE id = ?", (support_admin_id,)).fetchone()
        )
        self.assertEqual(
            conn.execute(
                "SELECT COUNT(*) FROM admin_permission_audit WHERE target_admin_id = ?",
                (support_admin_id,),
            ).fetchone()[0],
            0,
        )
        conn.close()

    def test_super_admin_page_visibility_blocks_and_restores_admin_access(self):
        self._register("Visibility Student", "visibility-student@test.com", "pass1234")
        self._login("ayush5075mi@gmail.com", "india123")
        self.client.post(
            "/admin/permissions",
            data={
                "action": "create_admin",
                "name": "Visibility Admin",
                "email": "visibility-admin@test.com",
                "password": "pass1234",
            },
        )

        page = self.client.get("/admin/permissions")
        self.assertNotIn(b"Page Visibility Management", page.data)
        self.assertNotIn(b"Are you sure you want to disable this page?", page.data)
        conn = sqlite3.connect(self.db_path)
        course_page_id = conn.execute(
            "SELECT id FROM page_visibility WHERE page_name = 'Courses'"
        ).fetchone()[0]
        default_pages = conn.execute(
            "SELECT COUNT(*) FROM page_visibility WHERE is_visible = 1"
        ).fetchone()[0]
        conn.close()
        self.assertEqual(default_pages, 25)

        normal_admin = self.app.test_client()
        normal_admin.post(
            "/login",
            data={"email": "visibility-admin@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        self.assertEqual(
            normal_admin.post(
                "/admin/permissions",
                data={
                    "action": "update_page_visibility",
                    "page_id": str(course_page_id),
                    "is_visible": "0",
                },
            ).status_code,
            403,
        )

        hidden = self.client.post(
            "/admin/permissions",
            data={
                "action": "update_page_visibility",
                "page_id": str(course_page_id),
                "is_visible": "0",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Page visibility updated successfully.", hidden.data)
        self.assertNotIn(b"Page Visibility Management", hidden.data)
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 403)
        self.assertEqual(normal_admin.post("/admin/courses").status_code, 403)
        normal_admin_page = normal_admin.get("/admin")
        nav_links = normal_admin_page.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertNotIn(b'href="/admin/courses"', nav_links)
        self.assertEqual(self.client.get("/admin/courses").status_code, 200)

        init_db(self.db_path, self.app.config)
        conn = sqlite3.connect(self.db_path)
        saved_visibility, = conn.execute(
            "SELECT is_visible FROM page_visibility WHERE page_name = 'Courses'"
        ).fetchone()
        conn.close()
        self.assertEqual(saved_visibility, 0)

        restored = self.client.post(
            "/admin/permissions",
            data={
                "action": "update_page_visibility",
                "page_id": str(course_page_id),
                "is_visible": "1",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Page visibility updated successfully.", restored.data)
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 200)
        restored_admin_page = normal_admin.get("/admin")
        restored_nav_links = restored_admin_page.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertIn(b'href="/admin/courses"', restored_nav_links)

        conn = sqlite3.connect(self.db_path)
        admin_page_id = conn.execute(
            "SELECT id FROM page_visibility WHERE page_name = 'Admin'"
        ).fetchone()[0]
        conn.close()
        self.client.post(
            "/admin/permissions",
            data={
                "action": "update_page_visibility",
                "page_id": str(admin_page_id),
                "is_visible": "0",
            },
        )
        self.assertEqual(normal_admin.get("/admin").status_code, 403)
        self.assertEqual(normal_admin.post("/admin/lab/generate").status_code, 403)
        normal_admin_course_page = normal_admin.get("/admin/courses")
        course_nav_links = normal_admin_course_page.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertNotIn(b'href="/admin"', course_nav_links)

        student_client = self.app.test_client()
        student_client.post(
            "/login",
            data={"email": "visibility-student@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        conn = sqlite3.connect(self.db_path)
        dashboard_page_id = conn.execute(
            "SELECT id FROM page_visibility WHERE page_name = 'Student Dashboard'"
        ).fetchone()[0]
        conn.close()
        self.client.post(
            "/admin/permissions",
            data={
                "action": "update_page_visibility",
                "page_id": str(dashboard_page_id),
                "is_visible": "0",
            },
        )
        self.assertEqual(student_client.get("/dashboard").status_code, 403)
        visible_student_page = student_client.get("/attendance")
        student_nav_links = visible_student_page.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertNotIn(b'href="/dashboard"', student_nav_links)
        conn = sqlite3.connect(self.db_path)
        documents_page_id = conn.execute(
            "SELECT id FROM page_visibility WHERE page_name = 'Student Documents'"
        ).fetchone()[0]
        conn.close()
        self.client.post(
            "/admin/permissions",
            data={
                "action": "update_page_visibility",
                "page_id": str(documents_page_id),
                "is_visible": "0",
            },
        )
        self.assertEqual(student_client.get("/documents/download/1").status_code, 403)

    def test_super_admin_can_manage_per_admin_page_access(self):
        self._login("ayush5075mi@gmail.com", "india123")
        self.client.post(
            "/admin/permissions",
            data={
                "action": "create_admin",
                "name": "Page Access Admin",
                "email": "page-access-admin@test.com",
                "password": "pass1234",
            },
        )
        conn = sqlite3.connect(self.db_path)
        admin_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("page-access-admin@test.com",)
        ).fetchone()[0]
        conn.close()

        self.client.post(
            "/admin/permissions",
            data={
                "action": "save_permissions",
                "admin_id": str(admin_id),
                "perm_dashboard_view": "on",
                "perm_course_management_view": "on",
            },
        )
        page = self.client.get(f"/admin/permissions?pages_id={admin_id}")
        self.assertIn(b"Page Permissions - Page Access Admin", page.data)
        self.assertIn(b"Enable All", page.data)
        self.assertIn(b"Disable All", page.data)
        self.assertIn(b"Save Permissions", page.data)
        self.assertIn(b"page_perm_courses", page.data)
        self.assertIn(f"pages_id={admin_id}".encode(), page.data)
        self.assertIn(b"#page-permissions", page.data)

        normal_admin = self.app.test_client()
        normal_admin.post(
            "/login",
            data={"email": "page-access-admin@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        denied = normal_admin.post(
            "/admin/permissions",
            data={
                "action": "save_admin_page_permissions",
                "admin_id": str(admin_id),
                "page_perm_dashboard": "on",
            },
        )
        self.assertEqual(denied.status_code, 403)

        saved = self.client.post(
            "/admin/permissions",
            data={
                "action": "save_admin_page_permissions",
                "admin_id": str(admin_id),
                "page_perm_dashboard": "on",
                "page_perm_courses": "on",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Page permissions saved.", saved.data)
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 200)

        disabled = self.client.post(
            "/admin/permissions",
            data={
                "action": "save_admin_page_permissions",
                "admin_id": str(admin_id),
                "page_perm_dashboard": "on",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Page permissions saved.", disabled.data)
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 403)
        admin_home = normal_admin.get("/admin")
        nav_links = admin_home.data.split(b'<div class="nav-links">', 1)[1].split(b"</div>", 1)[0]
        self.assertNotIn(b'href="/admin/courses"', nav_links)
        conn = sqlite3.connect(self.db_path)
        enabled, updater_id = conn.execute(
            "SELECT is_enabled, updated_by FROM admin_page_permission WHERE admin_id = ? AND page_key = 'courses'",
            (admin_id,),
        ).fetchone()
        conn.close()
        self.assertEqual(enabled, 0)
        self.assertIsNotNone(updater_id)

        self.client.post(
            "/admin/permissions",
            data={
                "action": "save_admin_page_permissions",
                "admin_id": str(admin_id),
                "page_perm_dashboard": "on",
                "page_perm_courses": "on",
            },
        )
        self.assertEqual(normal_admin.get("/admin/courses").status_code, 200)

    def test_admin_permissions_template_defaults_missing_stats(self):
        with self.app.test_request_context("/admin/permissions"):
            page = render_template(
                "admin_permissions.html",
                admins=[],
                permission_modal=None,
                selected_status="",
                selected_role="",
                search_term="",
                default_permissions={},
                permission_audit=[],
            )

        self.assertIn("<strong>0</strong>", page)
        self.assertIn("Admin Permission Management", page)

    def test_admin_can_view_students_page(self):
        self._register(
            "Listed Student",
            "listed@test.com",
            "pass1234",
            course_id="1",
            student_id="S-2001",
            mobile_no="8888888888",
        )
        self._login("admin@ayu.com", "admin123")

        resp = self.client.get("/admin/students")

        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Registration", resp.data)
        self.assertIn(b"Listed Student", resp.data)
        self.assertIn(b"S-2001", resp.data)
        self.assertIn(b"Data Structures 1 YEAR", resp.data)

        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO course (name, year) VALUES (?, ?)", ("Algorithms", "2Year"))
        conn.commit()
        conn.close()
        self._register(
            "Other Course Student",
            "other@test.com",
            "pass1234",
            course_id="2",
            student_id="S-2002",
            mobile_no="7777777777",
        )

        filtered = self.client.get("/admin/students?course_id=1")
        self.assertIn(b"Listed Student", filtered.data)
        self.assertNotIn(b"Other Course Student", filtered.data)

        conn = sqlite3.connect(self.db_path)
        student_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("listed@test.com",)
        ).fetchone()[0]
        conn.close()
        update = self.client.post(
            f"/admin/students/{student_id}/edit",
            data={
                "name": "Updated Student",
                "student_id": "S-2009",
                "email": "updated@test.com",
                "mobile_no": "9999999999",
                "course_id": "1",
                "section": "A",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Student updated.", update.data)
        self.assertIn(b"Updated Student", update.data)

        delete = self.client.post(
            f"/admin/students/{student_id}/delete", follow_redirects=True
        )
        self.assertIn(b"Student deleted.", delete.data)
        self.assertNotIn(b"Updated Student", delete.data)

    def test_admin_can_manage_courses(self):
        self._login("admin@ayu.com", "admin123")

        page = self.client.get("/admin/courses")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Add Course", page.data)
        self.assertIn(b"Data Structures", page.data)
        self.assertIn(b'for="course-section"', page.data)
        self.assertIn(b'name="course_section"', page.data)
        self.assertIn(b'name="status"', page.data)
        for heading in (
            b"<th>ID</th>", b"<th>Course Name</th>", b"<th>Year</th>",
            b"<th>Section</th>", b"<th>Status</th>", b"<th>Actions</th>",
        ):
            self.assertIn(heading, page.data)
        self.assertIn(b"<th>Section</th>", page.data)
        self.assertNotIn(b"Duration", page.data)
        self.assertIn(b'<option value="A" selected>A</option>', page.data)
        self.assertNotIn(b"Add Subject", page.data)
        self.assertNotIn(b'id="course-subject-filter"', page.data)

        conn = sqlite3.connect(self.db_path)
        conn.execute("UPDATE course SET course_section = '' WHERE id = 1")
        conn.commit()
        conn.close()
        edit_page = self.client.get("/admin/courses/1/edit")
        self.assertIn(b'<option value="A" selected>A</option>', edit_page.data)

        add = self.client.post(
            "/admin/courses",
            data={"name": "Networks", "subject": "Computer Science", "year": "2Year", "course_section": "D", "status": "active"},
            follow_redirects=True,
        )
        self.assertIn(b"Course added.", add.data)
        conn = sqlite3.connect(self.db_path)
        saved_course = conn.execute(
            "SELECT course_section, status, created_at, updated_at "
            "FROM course WHERE name = ? AND year = ?", ("Networks", "2Year")
        ).fetchone()
        conn.close()
        self.assertEqual(saved_course[:2], ("D", "active"))
        self.assertTrue(saved_course[2])
        self.assertTrue(saved_course[3])

        missing_section = self.client.post(
            "/admin/courses",
            data={
                "name": "Course Without Section",
                "year": "2Year",
                "duration_years": "3",
                "status": "active",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Please enter Course Name, Year, Section and Status.", missing_section.data)
        self.assertNotIn(b"duration", missing_section.data.lower())

        missing_status = self.client.post(
            "/admin/courses",
            data={"name": "Course Without Status", "year": "2Year", "course_section": "A"},
            follow_redirects=True,
        )
        self.assertIn(b"Please enter Course Name, Year, Section and Status.", missing_status.data)

        same_name = self.client.post(
            "/admin/courses",
            data={"name": "Networks", "year": "3Year", "course_section": "B", "status": "active"},
            follow_redirects=True,
        )
        self.assertIn(b"Course added.", same_name.data)

        duplicate = self.client.post(
            "/admin/courses",
            data={"name": "Networks", "year": "2Year", "course_section": "A", "status": "active"},
            follow_redirects=True,
        )
        self.assertIn(b"already exists", duplicate.data)

        other_course = self.client.post(
            "/admin/courses",
            data={"name": "Operating Systems", "year": "1Year", "course_section": "C", "status": "active"},
            follow_redirects=True,
        )
        self.assertIn(b"Course added.", other_course.data)

        search = self.client.get("/admin/courses?subject=Computer+Science")
        self.assertIn(b"Networks", search.data)
        self.assertNotIn(b"Computer Science</td>", search.data)
        self.assertNotIn(b"Data Structures", search.data)

        conn = sqlite3.connect(self.db_path)
        course_id = conn.execute(
            "SELECT id FROM course WHERE name = ? AND year = ?", ("Networks", "2Year")
        ).fetchone()[0]
        conn.close()
        update = self.client.post(
            f"/admin/courses/{course_id}/edit",
            data={"name": "Computer Networks", "year": "4Year", "course_section": "G", "status": "inactive"},
            follow_redirects=True,
        )
        self.assertIn(b"Course updated.", update.data)
        self.assertIn(b"Computer Networks", update.data)
        self.assertIn(b"4 YEAR", update.data)
        conn = sqlite3.connect(self.db_path)
        updated_course = conn.execute(
            "SELECT course_section, status, created_at, updated_at, subject "
            "FROM course WHERE id = ?", (course_id,)
        ).fetchone()
        conn.close()
        self.assertEqual(updated_course[:2], ("G", "inactive"))
        self.assertTrue(updated_course[2])
        self.assertTrue(updated_course[3])
        self.assertEqual(updated_course[4], "Computer Science")

        requested_update = self.client.post(
            "/admin/courses/1/edit",
            data={"name": "B.TECH", "year": "1Year", "course_section": "A", "status": "active"},
            follow_redirects=True,
        )
        self.assertIn(b"Course updated.", requested_update.data)
        conn = sqlite3.connect(self.db_path)
        requested_course = conn.execute(
            "SELECT name, year, course_section, status FROM course WHERE id = 1"
        ).fetchone()
        conn.close()
        self.assertEqual(requested_course, ("B.TECH", "1Year", "A", "active"))

        self._register("Course Student", "course-owner@test.com", "pass1234", course_id="1")
        blocked = self.client.post("/admin/courses/1/delete", follow_redirects=True)
        self.assertIn(b"Course deleted", blocked.data)

        conn = sqlite3.connect(self.db_path)
        student_course = conn.execute(
            "SELECT course_id FROM user WHERE email = ?", ("course-owner@test.com",)
        ).fetchone()[0]
        conn.close()
        self.assertIsNone(student_course)

        delete = self.client.post(
            f"/admin/courses/{course_id}/delete", follow_redirects=True
        )
        self.assertIn(b"Course deleted.", delete.data)

    def test_admin_can_manage_devices(self):
        self._login("admin@ayu.com", "admin123")

        page = self.client.get("/admin/devices")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Device Management", page.data)
        self.assertIn(b"Add Device", page.data)

        settings = self.client.post(
            "/admin/devices/settings",
            data={"lock_minutes": "20"},
            follow_redirects=True,
        )
        self.assertIn(b"PC lock duration saved: 20 minutes", settings.data)
        conn = sqlite3.connect(self.db_path)
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM device WHERE release_minutes = 20").fetchone()[0], 0)
        conn.close()

        added = self.client.post(
            "/admin/devices",
            data={
                "name": "Computer 01",
                "asset_tag": "LAB-PC-001",
                "ip_address": "192.168.1.101",
                "location": "Lab Room A",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Device added.", added.data)
        self.assertIn(b"Computer 01", added.data)
        self.assertIn(b"LAB-PC-001", added.data)

        duplicate = self.client.post(
            "/admin/devices",
            data={"name": "Another Computer", "asset_tag": "lab-pc-001"},
            follow_redirects=True,
        )
        self.assertIn(b"asset tag is already in use", duplicate.data)

        conn = sqlite3.connect(self.db_path)
        device_id = conn.execute(
            "SELECT id FROM device WHERE asset_tag = ?", ("LAB-PC-001",)
        ).fetchone()[0]
        conn.close()
        toggled = self.client.post(
            f"/admin/devices/{device_id}/toggle", follow_redirects=True
        )
        self.assertIn(b"Device status updated.", toggled.data)

        conn = sqlite3.connect(self.db_path)
        is_active = conn.execute(
            "SELECT is_active FROM device WHERE id = ?", (device_id,)
        ).fetchone()[0]
        conn.close()
        self.assertEqual(is_active, 0)

    def test_legacy_auto_device_displays_pc_name(self):
        self._login("admin@ayu.com", "admin123")
        self.client.post(
            "/admin/devices",
            data={"name": "Attendance PC 001", "asset_tag": "AUTO-PC-001"},
        )

        page = self.client.get("/admin/devices")
        self.assertIn(b"<strong>PC_001</strong>", page.data)

    def test_deactivated_device_cannot_mark_attendance(self):
        self._register("Inactive Device Student", "inactive-device@test.com", "pass1234", course_id="1")
        self._login("admin@ayu.com", "admin123")
        generated = self.client.post(
            "/admin/lab/generate",
            data={"course_id": "1", "minutes": "15"},
            follow_redirects=True,
        )
        match = re.search(rb"Lab key generated: (\d{6})", generated.data)
        self.assertIsNotNone(match)

        self.client.post(
            "/admin/devices",
            data={
                "name": "Inactive PC",
                "asset_tag": "INACTIVE-PC-001",
                "ip_address": "127.0.0.1",
                "location": "Home",
            },
        )
        conn = sqlite3.connect(self.db_path)
        device_id = conn.execute(
            "SELECT id FROM device WHERE asset_tag = ?", ("INACTIVE-PC-001",)
        ).fetchone()[0]
        conn.close()
        self.client.post(f"/admin/devices/{device_id}/toggle")
        self.client.post("/logout")

        self._login("inactive-device@test.com", "pass1234")
        blocked = self.client.post(
            "/lab/enter", data={"code": match.group(1).decode()}, follow_redirects=True
        )
        self.assertIn(b"deactivated", blocked.data.lower())

        conn = sqlite3.connect(self.db_path)
        is_active = conn.execute(
            "SELECT is_active FROM device WHERE id = ?", (device_id,)
        ).fetchone()[0]
        attendance_count = conn.execute(
            "SELECT COUNT(*) FROM attendance WHERE student_id = (SELECT id FROM user WHERE email = ?)",
            ("inactive-device@test.com",),
        ).fetchone()[0]
        conn.close()
        self.assertEqual(is_active, 0)
        self.assertEqual(attendance_count, 0)

    def test_unregistered_pc_cannot_mark_attendance(self):
        self._register("Unregistered PC Student", "unregistered-pc@test.com", "pass1234", course_id="1")
        self._login("admin@ayu.com", "admin123")
        generated = self.client.post(
            "/admin/lab/generate",
            data={"course_id": "1", "minutes": "15"},
            follow_redirects=True,
        )
        match = re.search(rb"Lab key generated: (\d{6})", generated.data)
        self.assertIsNotNone(match)
        self.client.post("/logout")
        self._login("unregistered-pc@test.com", "pass1234")

        response = self.client.post(
            "/lab/enter",
            data={"code": match.group(1).decode()},
            headers={"X-Forwarded-For": "127.0.0.1"},
            environ_base={"REMOTE_ADDR": "192.0.2.8"},
            follow_redirects=True,
        )
        self.assertIn(b"registered lab PC", response.data)
        conn = sqlite3.connect(self.db_path)
        attendance_count = conn.execute("SELECT COUNT(*) FROM attendance").fetchone()[0]
        device_count = conn.execute("SELECT COUNT(*) FROM device").fetchone()[0]
        conn.close()
        self.assertEqual(attendance_count, 0)
        self.assertEqual(device_count, 0)

    def test_subject_attendance_and_reports(self):
        self._login("admin@ayu.com", "admin123")

        added = self.client.post(
            "/admin/subjects", data={"name": "DBMS"}, follow_redirects=True
        )
        self.assertIn(b"Subject added.", added.data)
        self.assertIn(b"DBMS", added.data)

        conn = sqlite3.connect(self.db_path)
        subject_id = conn.execute(
            "SELECT id FROM subject WHERE name = ?", ("DBMS",)
        ).fetchone()[0]
        conn.close()
        assigned = self.client.post(
            f"/admin/subjects/{subject_id}/courses",
            data={"course_id": "1"},
            follow_redirects=True,
        )
        self.assertIn(b"Subject assigned to course.", assigned.data)

        dashboard = self.client.get("/admin")
        self.assertIn(b"Subject Attendance", dashboard.data)
        self.assertIn(b"DBMS", dashboard.data)

        self._register("Subject Student", "subject@test.com", "pass1234", course_id="1")
        generated = self.client.post(
            "/admin/lab/generate",
            data={"course_id": "1", "subject_id": str(subject_id), "minutes": "15"},
            follow_redirects=True,
        )
        match = re.search(rb"Lab key generated: (\d{6})", generated.data)
        self.assertIsNotNone(match)
        code = match.group(1).decode()

        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "INSERT INTO device (name, asset_tag, ip_address, location, created_at) VALUES (?, ?, ?, ?, ?)",
            ("Lab PC 1", "LAB-PC-1", "127.0.0.1", "Lab", "2026-09-28 10:00:00"),
        )
        conn.commit()
        conn.close()
        self.client.post("/logout")
        self._login("subject@test.com", "pass1234")
        marked = self.client.post("/lab/enter", data={"code": code}, follow_redirects=True)
        self.assertIn(b"Attendance marked successfully", marked.data)

        conn = sqlite3.connect(self.db_path)
        device = conn.execute(
            "SELECT name, asset_tag, ip_address FROM device WHERE ip_address = ?",
            ("127.0.0.1",),
        ).fetchone()
        self.assertEqual(device, ("Lab PC 1", "LAB-PC-1", "127.0.0.1"))
        locked_until = conn.execute(
            "SELECT lock_until FROM device WHERE ip_address = ?", ("127.0.0.1",)
        ).fetchone()
        self.assertIsNotNone(locked_until[0])

        blocked = self.client.post("/lab/enter", data={"code": code}, follow_redirects=True)
        self.assertIn(b"This PC is locked until", blocked.data)

        self.client.post("/logout")
        self._login("admin@ayu.com", "admin123")
        devices_page = self.client.get("/admin/devices")
        self.assertIn(b"Subject Student", devices_page.data)
        device_id = conn.execute(
            "SELECT id FROM device WHERE ip_address = ?", ("127.0.0.1",)
        ).fetchone()[0]
        conn.close()
        released = self.client.post(
            f"/admin/devices/{device_id}/release", follow_redirects=True
        )
        self.assertIn(b"Device lock released by admin", released.data)

        conn = sqlite3.connect(self.db_path)
        student_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("subject@test.com",)
        ).fetchone()[0]
        conn.execute(
            "UPDATE device SET lock_until = ?, current_student_id = ? WHERE id = ?",
            ("2000-01-01 00:00:00", student_id, device_id),
        )
        conn.commit()
        conn.close()
        self.client.get("/admin/devices")
        activity_page = self.client.get("/admin/login-activity")
        self.assertIn(b"Subject Student", activity_page.data)
        self.assertIn(b"subject@test.com", activity_page.data)
        self.assertIn(b"9999999999", activity_page.data)
        self.assertIn(b"LAB-PC-1", activity_page.data)
        self.assertIn(b"LOCK", activity_page.data)
        self.assertIn(b"UNLOCK", activity_page.data)
        self.assertIn(b"configured device lock duration elapsed", activity_page.data)

        deleted = self.client.post(
            f"/admin/devices/{device_id}/delete", follow_redirects=True
        )
        self.assertIn(b"Device deleted.", deleted.data)
        activity_page = self.client.get("/admin/login-activity")
        self.assertIn(b"Device Lock / Unlock / Delete Activity", activity_page.data)
        self.assertIn(b"DELETE", activity_page.data)
        self.assertIn(b"Device deleted: Lab PC 1 (asset tag LAB-PC-1).", activity_page.data)

        overview = self.client.get(f"/admin/attendance?subject_id={subject_id}")
        self.assertIn(b"Subject Student", overview.data)
        self.assertIn(b"DBMS", overview.data)

        report = self.client.get(f"/admin/reports?subject_id={subject_id}")
        self.assertEqual(report.status_code, 200)
        self.assertIn(b"Subject Student", report.data)
        self.assertIn(b"DBMS", report.data)

        attendance_export = self.client.get(
            f"/admin/attendance/export?view=report&subject_id={subject_id}"
        )
        self.assertEqual(attendance_export.status_code, 200)
        self.assertEqual(
            attendance_export.headers["Content-Disposition"],
            "attachment; filename=reports.xlsx",
        )
        workbook = openpyxl.load_workbook(io.BytesIO(attendance_export.data), read_only=True)
        self.assertEqual(workbook.sheetnames, ["Attendance"])
        exported_rows = list(workbook.active.values)
        workbook.close()
        self.assertEqual(exported_rows[0], (
            "Date & Time", "Student Name", "Email", "Roll No.", "Course",
            "Section", "Subject", "Lab Code",
        ))
        self.assertTrue(any(row[1] == "Subject Student" and row[6] == "DBMS" for row in exported_rows[1:]))

    def test_admin_can_activate_and_deactivate_students(self):
        self._register("Inactive Student", "inactive@test.com", "pass1234", course_id="1")
        self._login("admin@ayu.com", "admin123")

        conn = sqlite3.connect(self.db_path)
        student = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("inactive@test.com",)
        ).fetchone()[0]
        conn.close()

        deactivate = self.client.post(
            f"/admin/students/{student}/deactivate",
            follow_redirects=True,
        )
        self.assertIn(b"Student deactivated.", deactivate.data)

        login_blocked = self.client.post(
            "/login",
            data={"email": "inactive@test.com", "password": "pass1234"},
            follow_redirects=True,
        )
        self.assertIn(b"deactivated", login_blocked.data.lower())

        activate = self.client.post(
            f"/admin/students/{student}/activate",
            follow_redirects=True,
        )
        self.assertIn(b"Student activated.", activate.data)

    def test_admin_forgot_password_does_not_require_current_password(self):
        resp = self.client.post(
            "/admin/forgot-password",
            data={
                "email": "admin@ayu.com",
                "new_password": "newadmin123",
                "confirm_password": "newadmin123",
            },
            follow_redirects=True,
        )

        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Admin password updated successfully", resp.data)

        login = self._login("admin@ayu.com", "newadmin123")
        self.assertEqual(login.status_code, 200)
        self.assertIn(b"Admin Panel", login.data)

    def test_admin_can_reset_user_password_without_current_password(self):
        self._register("Managed Student", "managed@test.com", "pass1234")
        self._login("admin@ayu.com", "admin123")

        page = self.client.get("/admin/forgot-user-password")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Reset Student Password", page.data)
        self.assertNotIn(b"Current Password", page.data)

        resp = self.client.post(
            "/admin/forgot-user-password",
            data={
                "email": "managed@test.com",
                "new_password": "managed123",
                "confirm_password": "managed123",
            },
            follow_redirects=True,
        )

        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"User password updated successfully", resp.data)

        self.client.post("/logout")
        login = self._login("managed@test.com", "managed123")
        self.assertEqual(login.status_code, 200)
        self.assertIn(b"Welcome", login.data)

    def test_invalid_lab_key_rejected(self):
        self._register("Test Student3", "student3@test.com", "pass1234")
        self._login("student3@test.com", "pass1234")
        resp = self.client.post("/lab/enter", data={"code": "000000"}, follow_redirects=True)
        self.assertIn(b"Invalid or expired", resp.data)

    def test_duplicate_email_rejected(self):
        self._register("Dup", "dup@test.com", "pass1234")
        resp = self._register("Dup2", "dup@test.com", "pass1234")
        self.assertIn(b"already exists", resp.data)

    def test_registration_rejects_invalid_email_formats(self):
        invalid_emails = (
            "studentgmail.com",
            "student@",
            "@gmail.com",
            "student gmail@gmail.com",
        )
        for index, email in enumerate(invalid_emails):
            with self.subTest(email=email):
                response = self._register(
                    "Invalid Email Student",
                    email,
                    "pass1234",
                    student_id=f"INVALID-{index}",
                )
                self.assertIn(b"Enter a valid email address", response.data)
                conn = sqlite3.connect(self.db_path)
                saved = conn.execute("SELECT id FROM user WHERE student_id = ?", (f"INVALID-{index}",)).fetchone()
                conn.close()
                self.assertIsNone(saved)

    def test_registration_accepts_multi_part_email_domain(self):
        response = self._register("College Student", "student@college.ac.in", "pass1234")
        self.assertIn(b"Registration successful", response.data)

    def test_registration_rejects_invalid_mobile_numbers(self):
        invalid_mobile_numbers = (
            "987654321",
            "98765432101",
            "+919876543210",
            "98765 43210",
            "98765ABCDE",
            " 9876543210 ",
        )
        for index, mobile_no in enumerate(invalid_mobile_numbers):
            with self.subTest(mobile_no=mobile_no):
                response = self._register(
                    "Invalid Mobile Student",
                    f"invalid-mobile-{index}@test.com",
                    "pass1234",
                    student_id=f"INVALID-MOBILE-{index}",
                    mobile_no=mobile_no,
                )
                self.assertIn(b"Mobile number must be exactly 10 digits", response.data)
                conn = sqlite3.connect(self.db_path)
                saved = conn.execute(
                    "SELECT id FROM user WHERE student_id = ?", (f"INVALID-MOBILE-{index}",)
                ).fetchone()
                conn.close()
                self.assertIsNone(saved)

    def test_student_edit_rejects_invalid_mobile_number(self):
        self._register("Edit Mobile Student", "edit-mobile@test.com", "pass1234", student_id="S-EDIT")
        self._login("admin@ayu.com", "admin123")
        conn = sqlite3.connect(self.db_path)
        student_id = conn.execute(
            "SELECT id FROM user WHERE email = ?", ("edit-mobile@test.com",)
        ).fetchone()[0]
        conn.close()

        response = self.client.post(
            f"/admin/students/{student_id}/edit",
            data={
                "student_id": "S-EDIT",
                "name": "Edit Mobile Student",
                "email": "edit-mobile@test.com",
                "mobile_no": "+919876543210",
                "course_id": "1",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Mobile number must be exactly 10 digits", response.data)
        conn = sqlite3.connect(self.db_path)
        mobile_no = conn.execute("SELECT mobile_no FROM user WHERE id = ?", (student_id,)).fetchone()[0]
        conn.close()
        self.assertEqual(mobile_no, "9999999999")

    def test_short_password_rejected(self):
        resp = self._register("Shorty", "short@test.com", "abc")
        self.assertIn(b"at least 6 characters", resp.data)

    def test_student_registration_includes_student_id_and_mobile(self):
        resp = self._register(
            "Enroll Student",
            "enroll@test.com",
            "pass1234",
            course_id="1",
            student_id="S-1001",
            mobile_no="9876543210",
        )
        self.assertEqual(resp.status_code, 200)

        conn = sqlite3.connect(self.db_path)
        row = conn.execute(
            "SELECT student_id, mobile_no FROM user WHERE email = ?",
            ("enroll@test.com",),
        ).fetchone()
        conn.close()

        self.assertIsNotNone(row)
        self.assertEqual(row[0], "S-1001")
        self.assertEqual(row[1], "9876543210")

    def test_admin_students_page_has_excel_tools(self):
        self._login("admin@ayu.com", "admin123")
        response = self.client.get("/admin/students")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Import Excel", response.data)
        self.assertIn(b"Export Excel", response.data)
        self.assertIn(b"Upload Excel File", response.data)
        self.assertIn(b".xlsx,.xls,.csv", response.data)

    def test_admin_can_import_students_and_skip_duplicate_ids(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["Student ID", "Student Name", "Course", "Mobile", "Email"])
        worksheet.append(["S-2001", "Imported Student", "Data Structures", "9876543210", "imported@test.com"])
        worksheet.append(["S-2001", "Duplicate Student", "Data Structures", "9876543211", "duplicate@test.com"])
        worksheet.append(["S-2002", "Invalid Email Student", "Data Structures", "9876543212", "studentgmail.com"])
        worksheet.append(["S-2003", "Invalid Mobile Student", "Data Structures", "98765 43210", "invalid-mobile@test.com"])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        self._login("admin@ayu.com", "admin123")
        response = self.client.post(
            "/admin/students/import",
            data={"student_file": (output, "students.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"1 student(s) imported", response.data)
        self.assertIn(b"Duplicate Student IDs skipped", response.data)
        self.assertIn(b"email studentgmail.com is invalid", response.data)
        self.assertIn(b"mobile number 98765 43210 must be exactly 10 digits", response.data)
        conn = sqlite3.connect(self.db_path)
        rows = conn.execute(
            "SELECT name, email, course_id FROM user WHERE student_id = ?", ("S-2001",)
        ).fetchall()
        conn.close()
        self.assertEqual(rows, [("Imported Student", "imported@test.com", 1)])

    def test_admin_can_export_students_as_xlsx(self):
        self._register(
            "Export Student",
            "export@test.com",
            "pass1234",
            course_id="1",
            student_id="S-2002",
            mobile_no="9999999999",
        )
        self._login("admin@ayu.com", "admin123")
        response = self.client.get("/admin/students/export")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["Content-Disposition"], 'attachment; filename=students.xlsx')
        workbook = openpyxl.load_workbook(io.BytesIO(response.data), read_only=True)
        self.assertEqual(workbook.sheetnames, ["Students"])
        self.assertEqual(
            list(workbook.active.values)[0],
            ("Student ID", "Student Name", "Course", "Section", "Mobile", "Email", "Registration Date", "Status"),
        )
        workbook.close()

    def test_registration_migrates_legacy_user_table_columns(self):
        legacy_db_fd, legacy_db_path = tempfile.mkstemp()
        os.close(legacy_db_fd)
        try:
            conn = sqlite3.connect(legacy_db_path)
            conn.execute(
                "CREATE TABLE user (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, "
                "email TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL, "
                "created_at TEXT NOT NULL)"
            )
            conn.commit()
            conn.close()

            legacy_app = create_app(type("LegacyConfig", (Config,), {
                "SECRET_KEY": "legacy-test",
                "UPLOAD_FOLDER": tempfile.mkdtemp(),
                "DB_ENGINE": "sqlite",
                "CSRF_ENABLED": False,
            }), db_path=legacy_db_path)
            conn = sqlite3.connect(legacy_db_path)
            conn.execute("INSERT INTO course (name, year) VALUES (?, ?)", ("Migrated Course", "2Year"))
            conn.commit()
            conn.close()
            client = legacy_app.test_client()
            response = client.post(
                "/register",
                data={
                    "name": "Migrated Student",
                    "email": "migrated@test.com",
                    "password": "pass1234",
                    "student_id": "S-3001",
                    "mobile_no": "6666666666",
                    "section": "B",
                    "course_id": "1",
                },
                follow_redirects=True,
            )
            self.assertIn(b"Registration successful", response.data)
            conn = sqlite3.connect(legacy_db_path)
            migrated_section_year = conn.execute(
                "SELECT section, year FROM user WHERE email = ?", ("migrated@test.com",)
            ).fetchone()
            conn.close()
            self.assertEqual(migrated_section_year, ("B", "2Year"))
        finally:
            os.unlink(legacy_db_path)

    def test_admin_document_lifecycle(self):
        self._login("admin@ayu.com", "admin123")
        upload = self.client.post(
            "/admin/documents/upload",
            data={
                "title": "Intro Notes",
                "course_id": "1",
                "file": (io.BytesIO(b"hello world"), "notes.pdf"),
            },
            follow_redirects=True,
        )
        self.assertIn(b"Document uploaded.", upload.data)

        admin_page = self.client.get("/admin")
        self.assertIn(b"Uploaded Documents", admin_page.data)
        self.assertNotIn(b"Intro Notes", admin_page.data)

        documents_page = self.client.get("/admin/documents")
        self.assertEqual(documents_page.status_code, 200)
        self.assertIn(b"Uploaded Documents", documents_page.data)
        self.assertIn(b"Intro Notes", documents_page.data)
        self.assertIn(b"Data Structures", documents_page.data)
        self.assertIn(b"Data Structures 1 YEAR", documents_page.data)

        update = self.client.post(
            "/admin/documents/1/update",
            data={"title": "Updated Notes", "course_id": "1"},
            follow_redirects=True,
        )
        self.assertIn(b"Document updated.", update.data)

        show = self.client.get("/admin/documents/1/show")
        self.assertEqual(show.status_code, 200)

        delete = self.client.post(
            "/admin/documents/1/delete",
            follow_redirects=True,
        )
        self.assertIn(b"Document deleted.", delete.data)

    def test_admin_panel_handles_no_courses_state(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM course")
        conn.commit()
        conn.close()

        self._login("admin@ayu.com", "admin123")
        resp = self.client.get("/admin")
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Add Course", resp.data)
        self.assertIn(b"No courses available yet", resp.data)

        documents = self.client.get("/admin/documents")
        self.assertEqual(documents.status_code, 200)
        self.assertIn(b"Upload Document", documents.data)
        self.assertIn(b"No courses available yet", documents.data)

    def test_admin_courses_page_renders_subject_options(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO subject (name) VALUES (?)", ("Math",))
        conn.execute("INSERT INTO subject (name) VALUES (?)", ("Physics",))
        conn.commit()
        conn.close()

        self._login("admin@ayu.com", "admin123")
        response = self.client.get("/admin/courses")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Math", response.data)
        self.assertIn(b"Physics", response.data)
        self.assertNotIn(b'id="course-subject"', response.data)
        self.assertIn(b'id="course-subject-search"', response.data)
        self.assertIn(b'name="subject"', response.data)

    def test_invalid_course_selection_rejected(self):
        resp = self._register(
            "Bad Student",
            "bad@test.com",
            "pass1234",
            course_id="999",
            student_id="S-999",
            mobile_no="8888888888",
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Select a valid course", resp.data)

    def test_admin_lab_generation_rejects_invalid_course(self):
        self._login("admin@ayu.com", "admin123")
        resp = self.client.post(
            "/admin/lab/generate",
            data={"course_id": "999", "minutes": "15"},
            follow_redirects=True,
        )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(b"Select a valid course", resp.data)

    def test_full_lab_key_attendance_flow(self):
        self._login("admin@ayu.com", "admin123")

        reg_page = self.client.get("/register")
        self.assertIn(b"Data Structures 1 YEAR", reg_page.data)

        self._register("Attendee", "attendee@test.com", "pass1234", course_id="1")
        self.client.post("/logout")
        self._login("attendee@test.com", "pass1234")

        dash = self.client.get("/dashboard")
        self.assertIn(b"Data Structures", dash.data)

        self.client.post("/logout")
        self._login("admin@ayu.com", "admin123")
        gen_resp = self.client.post(
            "/admin/lab/generate", data={"course_id": "1", "minutes": "15"}, follow_redirects=True
        )
        self.assertIn(b"Lab key generated", gen_resp.data)
        self.assertIn(b"Time Left", gen_resp.data)
        self.assertIn(b"countdown", gen_resp.data)
        self.assertIn(b"class=\"action-btn\"", gen_resp.data)
        self.assertIn(b"danger-btn", gen_resp.data)

        match = re.search(rb"Lab key generated: (\d{6})", gen_resp.data)
        self.assertIsNotNone(match)
        code = match.group(1).decode()

        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "INSERT INTO device (name, asset_tag, ip_address, location, created_at) VALUES (?, ?, ?, ?, ?)",
            ("Lab PC 1", "LAB-PC-1", "127.0.0.1", "Lab", "2026-09-28 10:00:00"),
        )
        conn.commit()
        conn.close()
        self.client.post("/logout")
        self._login("attendee@test.com", "pass1234")
        active_dashboard = self.client.get("/dashboard")
        self.assertIn(b"Active Lab Sessions For Your Course", active_dashboard.data)
        self.assertIn(b"Time Left", active_dashboard.data)
        self.assertIn(b"countdown", active_dashboard.data)
        attend_resp = self.client.post("/lab/enter", data={"code": code}, follow_redirects=True)
        self.assertIn(b"Attendance marked successfully", attend_resp.data)

        dup_resp = self.client.post("/lab/enter", data={"code": code}, follow_redirects=True)
        self.assertIn(b"This PC is locked until", dup_resp.data)

        self.client.post("/logout")
        self._login("admin@ayu.com", "admin123")
        attendance_page = self.client.get("/admin/attendance")
        self.assertEqual(attendance_page.status_code, 200)
        self.assertIn(b"View Attendance", attendance_page.data)
        self.assertIn(b"Attendee", attendance_page.data)
        self.assertIn(b"attendee@test.com", attendance_page.data)
        self.assertNotIn(code.encode(), attendance_page.data)

    def test_lab_key_wrong_course_rejected(self):
        # Second course, student in course 1 shouldn't be able to use course 2's key
        conn = sqlite3.connect(self.db_path)
        conn.execute("INSERT INTO course (name, year) VALUES (?, ?)", ("Algorithms", "2Year"))
        conn.commit()
        conn.close()

        self._register("Cross Student", "cross@test.com", "pass1234", course_id="1")
        self._login("admin@ayu.com", "admin123")
        gen_resp = self.client.post(
            "/admin/lab/generate", data={"course_id": "2", "minutes": "15"}, follow_redirects=True
        )
        match = re.search(rb"Lab key generated: (\d{6})", gen_resp.data)
        code = match.group(1).decode()

        self.client.post("/logout")
        self._login("cross@test.com", "pass1234")
        resp = self.client.post("/lab/enter", data={"code": code}, follow_redirects=True)
        self.assertIn(b"Invalid or expired", resp.data)

    def test_student_import_accepts_year_variants_in_course_value(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["Student ID", "Student Name", "Course", "Mobile", "Email"])
        worksheet.append(["S-3001", "Year Variant Student", "Data Structures 1Year", "9876543210", "year-variant@test.com"])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        self._login("admin@ayu.com", "admin123")
        response = self.client.post(
            "/admin/students/import",
            data={"student_file": (output, "students.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"1 student(s) imported", response.data)
        conn = sqlite3.connect(self.db_path)
        row = conn.execute(
            "SELECT course_id, name, email FROM user WHERE student_id = ?",
            ("S-3001",),
        ).fetchone()
        conn.close()
        self.assertIsNotNone(row)
        self.assertEqual(row[0], 1)


if __name__ == "__main__":
    unittest.main()
