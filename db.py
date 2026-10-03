import sqlite3
from datetime import datetime, timezone

try:
    import pymysql
    import pymysql.cursors
except ImportError:  # pragma: no cover - only used when MySQL is not installed
    pymysql = None
from werkzeug.security import generate_password_hash
PAGE_VISIBILITY_DEFAULTS = (
    ("Admin", "/admin"),
    ("Admin Permissions", "/admin/permissions"),
    ("Courses", "/admin/courses"),
    ("Subjects", "/admin/subjects"),
    ("Device Management", "/admin/devices"),
    ("Login Activity", "/admin/login-activity"),
    ("Registration", "/admin/students"),
    ("Uploaded Documents", "/admin/documents"),
    ("View Attendance", "/admin/attendance"),
    ("Reports", "/admin/reports"),
    ("Password", "/admin/password"),
    ("Student Dashboard", "/dashboard"),
    ("Student Attendance", "/attendance"),
    ("Student Documents", "/documents"),
    ("Lab Entry", "/lab/enter"),
    ("Student Registration", "/register"),
    ("Forgot Password", "/forgot-password"),
    ("Admin Forgot Password", "/admin/forgot-password"),
    ("Forgot User Password", "/admin/forgot-user-password"),
    ("Course Editor", "/admin/courses/<course_id>/edit"),
    ("Subject Editor", "/admin/subjects/<subject_id>/edit"),
    ("Student Editor", "/admin/students/<student_id>/edit"),
    ("Document Viewer", "/admin/documents/<doc_id>/show"),
    ("Document Editor", "/admin/documents/<doc_id>/edit"),
    ("Session Attendance", "/admin/lab/<session_id>/attendance"),
)

SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS course (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name VARCHAR(100) NOT NULL,
    subject TEXT NOT NULL DEFAULT '',
    year TEXT NOT NULL DEFAULT '',
    course_section VARCHAR(10) NOT NULL DEFAULT 'A',
    status TEXT NOT NULL DEFAULT 'active',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS subject (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS course_subject (
    course_id INTEGER NOT NULL,
    subject_id INTEGER NOT NULL,
    PRIMARY KEY (course_id, subject_id),
    FOREIGN KEY (course_id) REFERENCES course(id) ON DELETE CASCADE,
    FOREIGN KEY (subject_id) REFERENCES subject(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS user (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    student_id TEXT,
    section TEXT,
    year TEXT,
    email VARCHAR(255) UNIQUE NOT NULL,
    mobile_no VARCHAR(10),
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'student',
    super_admin INTEGER NOT NULL DEFAULT 0,
    course_id INTEGER,
    is_active INTEGER NOT NULL DEFAULT 1,
    account_status TEXT NOT NULL DEFAULT 'active',
    failed_login_attempts INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    FOREIGN KEY (course_id) REFERENCES course(id)
);

CREATE TABLE IF NOT EXISTS document (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    course_id INTEGER NOT NULL,
    title TEXT NOT NULL,
    filename TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    FOREIGN KEY (course_id) REFERENCES course(id)
);

CREATE TABLE IF NOT EXISTS lab_session (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    course_id INTEGER NOT NULL,
    subject_id INTEGER,
    code TEXT NOT NULL,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    FOREIGN KEY (course_id) REFERENCES course(id)
    ,FOREIGN KEY (subject_id) REFERENCES subject(id)
);

CREATE TABLE IF NOT EXISTS attendance (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL,
    lab_session_id INTEGER NOT NULL,
    timestamp TEXT NOT NULL,
    UNIQUE (student_id, lab_session_id),
    FOREIGN KEY (student_id) REFERENCES user(id),
    FOREIGN KEY (lab_session_id) REFERENCES lab_session(id)
);

CREATE TABLE IF NOT EXISTS device (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    asset_tag TEXT NOT NULL UNIQUE,
    ip_address TEXT,
    location TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'APPROVED',
    device_type TEXT NOT NULL DEFAULT 'Desktop/Laptop',
    registered_by INTEGER,
    blocked_reason TEXT,
    lock_until TEXT,
    current_student_id INTEGER,
    release_minutes INTEGER NOT NULL DEFAULT 15,
    created_at TEXT NOT NULL,
    FOREIGN KEY (current_student_id) REFERENCES user(id)
);
"""

SCHEMA_MYSQL = """
CREATE TABLE IF NOT EXISTS course (
    id INT NOT NULL AUTO_INCREMENT,
    name VARCHAR(100) NOT NULL,
    subject VARCHAR(255) NOT NULL DEFAULT '',
    year VARCHAR(20) NOT NULL DEFAULT '',
    course_section VARCHAR(10) NOT NULL DEFAULT 'A',
    status ENUM('active', 'inactive') NOT NULL DEFAULT 'active',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS subject (
    id INT NOT NULL AUTO_INCREMENT,
    name VARCHAR(255) NOT NULL UNIQUE,
    PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS course_subject (
    course_id INT NOT NULL,
    subject_id INT NOT NULL,
    PRIMARY KEY (course_id, subject_id),
    FOREIGN KEY (course_id) REFERENCES course(id) ON DELETE CASCADE,
    FOREIGN KEY (subject_id) REFERENCES subject(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS user (
    id INT NOT NULL AUTO_INCREMENT,
    name VARCHAR(255) NOT NULL,
    student_id VARCHAR(100),
    section VARCHAR(50),
    year VARCHAR(20),
    email VARCHAR(255) NOT NULL UNIQUE,
    mobile_no VARCHAR(10),
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(50) NOT NULL DEFAULT 'student',
    super_admin TINYINT(1) NOT NULL DEFAULT 0,
    course_id INT,
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    account_status VARCHAR(20) NOT NULL DEFAULT 'active',
    failed_login_attempts INT NOT NULL DEFAULT 0,
    created_at DATETIME NOT NULL,
    PRIMARY KEY (id),
    FOREIGN KEY (course_id) REFERENCES course(id)
);

CREATE TABLE IF NOT EXISTS document (
    id INT NOT NULL AUTO_INCREMENT,
    course_id INT NOT NULL,
    title VARCHAR(255) NOT NULL,
    filename VARCHAR(255) NOT NULL,
    uploaded_at DATETIME NOT NULL,
    PRIMARY KEY (id),
    FOREIGN KEY (course_id) REFERENCES course(id)
);

CREATE TABLE IF NOT EXISTS lab_session (
    id INT NOT NULL AUTO_INCREMENT,
    course_id INT NOT NULL,
    subject_id INT,
    code VARCHAR(20) NOT NULL,
    created_at DATETIME NOT NULL,
    expires_at DATETIME NOT NULL,
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    PRIMARY KEY (id),
    FOREIGN KEY (course_id) REFERENCES course(id),
    FOREIGN KEY (subject_id) REFERENCES subject(id)
);

CREATE TABLE IF NOT EXISTS attendance (
    id INT NOT NULL AUTO_INCREMENT,
    student_id INT NOT NULL,
    lab_session_id INT NOT NULL,
    timestamp DATETIME NOT NULL,
    PRIMARY KEY (id),
    UNIQUE (student_id, lab_session_id),
    FOREIGN KEY (student_id) REFERENCES user(id),
    FOREIGN KEY (lab_session_id) REFERENCES lab_session(id)
);

CREATE TABLE IF NOT EXISTS device (
    id INT NOT NULL AUTO_INCREMENT,
    name VARCHAR(255) NOT NULL,
    asset_tag VARCHAR(100) NOT NULL UNIQUE,
    ip_address VARCHAR(45),
    location VARCHAR(255),
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    status VARCHAR(20) NOT NULL DEFAULT 'APPROVED',
    device_type VARCHAR(40) NOT NULL DEFAULT 'Desktop/Laptop',
    registered_by INT NULL,
    blocked_reason TEXT,
    lock_until DATETIME NULL,
    current_student_id INT NULL,
    release_minutes INT NOT NULL DEFAULT 15,
    created_at DATETIME NOT NULL,
    PRIMARY KEY (id),
    FOREIGN KEY (current_student_id) REFERENCES user(id)
);
"""


class MySQLConnection:
    def __init__(self, connection):
        self._connection = connection
        self._cursor = None

    def execute(self, query, params=()):
        mysql_query = query
        if "?" in query:
            mysql_query = query.replace("?", "%s")
        self._cursor = self._connection.cursor(cursor=pymysql.cursors.DictCursor)
        self._cursor.execute(mysql_query, params)
        return self._cursor

    def commit(self):
        self._connection.commit()

    def begin(self):
        self._connection.begin()

    def rollback(self):
        self._connection.rollback()

    def close(self):
        self._connection.close()


def get_mysql_conn(config, database=None):
    if pymysql is None:
        raise RuntimeError("PyMySQL is required for MySQL mode. Install it with: pip install PyMySQL")

    connection_args = {
        "host": config.get("DB_HOST", "localhost"),
        "port": int(config.get("DB_PORT", 3306)),
        "user": config.get("DB_USER", "root"),
        "password": config.get("DB_PASSWORD", ""),
        "autocommit": True,
        "charset": "utf8mb4",
        "cursorclass": pymysql.cursors.DictCursor,
    }
    if database:
        connection_args["database"] = database
    return pymysql.connect(**connection_args)


def get_conn(db_path, config=None):
    config = config or {}
    engine = (config.get("DB_ENGINE") or "sqlite").lower()

    if engine == "mysql":
        return MySQLConnection(get_mysql_conn(config, config.get("DB_NAME", "lab_portal")))

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def is_integrity_error(error):
    if isinstance(error, sqlite3.IntegrityError):
        return True
    return pymysql is not None and isinstance(error, pymysql.err.IntegrityError)


def ensure_user_columns(conn, engine):
    if engine == "mysql":
        existing_columns = {
            row["Field"] for row in conn.execute("SHOW COLUMNS FROM user").fetchall()
        }
    else:
        existing_columns = {
            row[1] for row in conn.execute("PRAGMA table_info(user)").fetchall()
        }

    missing_columns = {
        "student_id": "VARCHAR(100)" if engine == "mysql" else "TEXT",
        "section": "VARCHAR(50)" if engine == "mysql" else "TEXT",
        "year": "VARCHAR(20)" if engine == "mysql" else "TEXT",
        "mobile_no": "VARCHAR(30)" if engine == "mysql" else "TEXT",
        "course_id": "INT" if engine == "mysql" else "INTEGER",
        "is_active": "TINYINT(1)" if engine == "mysql" else "INTEGER",
        "account_status": "VARCHAR(20)" if engine == "mysql" else "TEXT",
        "failed_login_attempts": "INT" if engine == "mysql" else "INTEGER",
        "super_admin": "TINYINT(1)" if engine == "mysql" else "INTEGER NOT NULL DEFAULT 0",
    }
    for column, column_type in missing_columns.items():
        if column not in existing_columns:
            conn.execute(f"ALTER TABLE user ADD COLUMN {column} {column_type}")
            if column == "is_active":
                conn.execute("UPDATE user SET is_active = 1 WHERE is_active IS NULL")
            if column == "account_status":
                conn.execute("UPDATE user SET account_status = 'active' WHERE account_status IS NULL OR account_status = ''")

    conn.execute("UPDATE user SET account_status = 'active' WHERE account_status IS NULL OR account_status = ''")


def ensure_device_columns(conn, engine):
    if engine == "mysql":
        existing_columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM device").fetchall()}
    else:
        existing_columns = {row[1] for row in conn.execute("PRAGMA table_info(device)").fetchall()}

    if "lock_until" not in existing_columns:
        column_type = "DATETIME NULL" if engine == "mysql" else "TEXT"
        conn.execute(f"ALTER TABLE device ADD COLUMN lock_until {column_type}")
    if "release_minutes" not in existing_columns:
        column_type = "INT NOT NULL DEFAULT 15" if engine == "mysql" else "INTEGER NOT NULL DEFAULT 15"
        conn.execute(f"ALTER TABLE device ADD COLUMN release_minutes {column_type}")
    if "current_student_id" not in existing_columns:
        column_type = "INT NULL" if engine == "mysql" else "INTEGER"
        conn.execute(f"ALTER TABLE device ADD COLUMN current_student_id {column_type}")
    if "status" not in existing_columns:
        column_type = "VARCHAR(20) NOT NULL DEFAULT 'APPROVED'" if engine == "mysql" else "TEXT NOT NULL DEFAULT 'APPROVED'"
        conn.execute(f"ALTER TABLE device ADD COLUMN status {column_type}")
        conn.execute("UPDATE device SET status = 'BLOCKED' WHERE is_active = 0")
    if "device_type" not in existing_columns:
        column_type = "VARCHAR(40) NOT NULL DEFAULT 'Desktop/Laptop'" if engine == "mysql" else "TEXT NOT NULL DEFAULT 'Desktop/Laptop'"
        conn.execute(f"ALTER TABLE device ADD COLUMN device_type {column_type}")
    if "registered_by" not in existing_columns:
        column_type = "INT NULL" if engine == "mysql" else "INTEGER"
        conn.execute(f"ALTER TABLE device ADD COLUMN registered_by {column_type}")
    if "blocked_reason" not in existing_columns:
        column_type = "TEXT NULL" if engine == "mysql" else "TEXT"
        conn.execute(f"ALTER TABLE device ADD COLUMN blocked_reason {column_type}")


def ensure_course_columns(conn, engine):
    if engine == "mysql":
        existing_columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM course").fetchall()}
        name_column = next(row for row in conn.execute("SHOW COLUMNS FROM course").fetchall() if row["Field"] == "name")
        if str(name_column["Type"]).lower() != "varchar(100)":
            long_names = conn.execute(
                "SELECT COUNT(*) AS count FROM course WHERE CHAR_LENGTH(name) > 100"
            ).fetchone()["count"]
            if not long_names:
                conn.execute("ALTER TABLE course MODIFY COLUMN name VARCHAR(100) NOT NULL")
    else:
        existing_columns = {row[1] for row in conn.execute("PRAGMA table_info(course)").fetchall()}

    if "section" not in existing_columns and "code" in existing_columns:
        if engine == "mysql":
            conn.execute("ALTER TABLE course CHANGE COLUMN code section VARCHAR(50) NOT NULL")
        else:
            conn.execute("ALTER TABLE course RENAME COLUMN code TO section")
        existing_columns.add("section")

    if "subject" not in existing_columns:
        column_type = "VARCHAR(255) NOT NULL DEFAULT ''" if engine == "mysql" else "TEXT NOT NULL DEFAULT ''"
        conn.execute(f"ALTER TABLE course ADD COLUMN subject {column_type}")
    if "year" not in existing_columns:
        column_type = "VARCHAR(20) NOT NULL DEFAULT ''" if engine == "mysql" else "TEXT NOT NULL DEFAULT ''"
        conn.execute(f"ALTER TABLE course ADD COLUMN year {column_type}")
    has_legacy_duration = "duration_years" in existing_columns
    if "course_section" not in existing_columns:
        column_type = "VARCHAR(5) NOT NULL DEFAULT 'A'"
        conn.execute(f"ALTER TABLE course ADD COLUMN course_section {column_type}")
        if has_legacy_duration:
            conn.execute(
                "UPDATE course SET course_section = CASE COALESCE(duration_years, 4) "
                "WHEN 1 THEN 'A' WHEN 2 THEN 'B' WHEN 3 THEN 'C' WHEN 4 THEN 'D' "
                "WHEN 5 THEN 'E' WHEN 6 THEN 'F' WHEN 7 THEN 'G' ELSE 'A' END"
            )
    conn.execute(
        "UPDATE course SET course_section = CASE UPPER(TRIM(COALESCE(course_section, ''))) "
        "WHEN 'A' THEN 'A' WHEN 'B' THEN 'B' WHEN 'C' THEN 'C' WHEN 'D' THEN 'D' "
        "WHEN 'E' THEN 'E' WHEN 'F' THEN 'F' WHEN 'G' THEN 'G' "
        "WHEN '1' THEN 'A' WHEN '2' THEN 'B' WHEN '3' THEN 'C' WHEN '4' THEN 'D' "
        "WHEN '5' THEN 'E' WHEN '6' THEN 'F' WHEN '7' THEN 'G' ELSE 'A' END "
        "WHERE course_section IS NULL OR course_section <> UPPER(TRIM(course_section)) "
        "OR UPPER(TRIM(COALESCE(course_section, ''))) NOT IN "
        "('A', 'B', 'C', 'D', 'E', 'F', 'G')"
    )
    if "status" not in existing_columns:
        column_type = (
            "ENUM('active', 'inactive') NOT NULL DEFAULT 'active'"
            if engine == "mysql" else "TEXT NOT NULL DEFAULT 'active'"
        )
        conn.execute(f"ALTER TABLE course ADD COLUMN status {column_type}")
        if "is_active" in existing_columns:
            conn.execute(
                "UPDATE course SET status = CASE WHEN COALESCE(is_active, 1) = 0 "
                "THEN 'inactive' ELSE 'active' END"
            )
    if "created_at" not in existing_columns:
        if engine == "mysql":
            conn.execute(
                "ALTER TABLE course ADD COLUMN created_at "
                "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP"
            )
        else:
            conn.execute("ALTER TABLE course ADD COLUMN created_at TEXT NOT NULL DEFAULT ''")
            conn.execute("UPDATE course SET created_at = CURRENT_TIMESTAMP WHERE created_at = ''")
    if "updated_at" not in existing_columns:
        if engine == "mysql":
            conn.execute(
                "ALTER TABLE course ADD COLUMN updated_at "
                "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP"
            )
        else:
            conn.execute("ALTER TABLE course ADD COLUMN updated_at TEXT NOT NULL DEFAULT ''")
            conn.execute("UPDATE course SET updated_at = CURRENT_TIMESTAMP WHERE updated_at = ''")


def remove_legacy_course_section(conn, engine):
    if engine == "mysql":
        indexes = conn.execute(
            "SHOW INDEX FROM course WHERE Column_name = 'section'"
        ).fetchall()
        for index in indexes:
            index_name = index["Key_name"].replace("`", "``")
            if index_name != "PRIMARY":
                conn.execute(f"ALTER TABLE course DROP INDEX `{index_name}`")
        columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM course").fetchall()}
        if "section" in columns:
            conn.execute(
                "UPDATE course SET course_section = CASE UPPER(TRIM(section)) "
                "WHEN 'A' THEN 'A' WHEN 'B' THEN 'B' WHEN 'C' THEN 'C' "
                "WHEN 'D' THEN 'D' WHEN 'E' THEN 'E' WHEN 'F' THEN 'F' "
                "WHEN 'G' THEN 'G' ELSE course_section END"
            )
            conn.execute("ALTER TABLE course DROP COLUMN section")
        return

    existing_columns = {row[1] for row in conn.execute("PRAGMA table_info(course)").fetchall()}
    if "section" not in existing_columns:
        return

    conn.commit()
    conn.execute("PRAGMA foreign_keys = OFF")
    conn.execute(
        "CREATE TABLE course_without_section ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR(100) NOT NULL, "
        "subject TEXT NOT NULL DEFAULT '', year TEXT NOT NULL DEFAULT '', "
        "course_section VARCHAR(10) NOT NULL DEFAULT 'A', "
        "status TEXT NOT NULL DEFAULT 'active', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, "
        "updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)"
    )
    conn.execute(
        "INSERT INTO course_without_section "
        "(id, name, subject, year, course_section, status, created_at, updated_at) "
        "SELECT id, name, COALESCE(subject, ''), COALESCE(year, ''), CASE UPPER(TRIM(section)) "
        "WHEN 'A' THEN 'A' WHEN 'B' THEN 'B' WHEN 'C' THEN 'C' "
        "WHEN 'D' THEN 'D' WHEN 'E' THEN 'E' WHEN 'F' THEN 'F' "
        "WHEN 'G' THEN 'G' ELSE COALESCE(course_section, 'A') END, "
        "COALESCE(status, 'active'), COALESCE(created_at, CURRENT_TIMESTAMP), "
        "COALESCE(updated_at, CURRENT_TIMESTAMP) FROM course"
    )
    conn.execute("DROP TABLE course")
    conn.execute("ALTER TABLE course_without_section RENAME TO course")
    conn.commit()
    conn.execute("PRAGMA foreign_keys = ON")


def ensure_course_section_width(conn, engine):
    if engine == "mysql":
        section_column = next(
            row for row in conn.execute("SHOW COLUMNS FROM course").fetchall()
            if row["Field"] == "course_section"
        )
        if str(section_column["Type"]).lower() != "varchar(10)":
            conn.execute(
                "ALTER TABLE course MODIFY COLUMN course_section "
                "VARCHAR(10) NOT NULL DEFAULT 'A'"
            )
        return

    columns = {row[1]: row[2].upper() for row in conn.execute("PRAGMA table_info(course)").fetchall()}
    if columns.get("course_section") == "VARCHAR(10)" and columns.get("name") == "VARCHAR(100)":
        return

    conn.commit()
    conn.execute("PRAGMA foreign_keys = OFF")
    conn.execute(
        "CREATE TABLE course_with_section_width ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR(100) NOT NULL, "
        "subject TEXT NOT NULL DEFAULT '', year TEXT NOT NULL DEFAULT '', "
        "course_section VARCHAR(10) NOT NULL DEFAULT 'A', "
        "status TEXT NOT NULL DEFAULT 'active', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, "
        "updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)"
    )
    conn.execute(
        "INSERT INTO course_with_section_width "
        "(id, name, subject, year, course_section, status, created_at, updated_at) "
        "SELECT id, name, COALESCE(subject, ''), COALESCE(year, ''), "
        "COALESCE(course_section, 'A'), COALESCE(status, 'active'), "
        "COALESCE(created_at, CURRENT_TIMESTAMP), COALESCE(updated_at, CURRENT_TIMESTAMP) "
        "FROM course"
    )
    conn.execute("DROP TABLE course")
    conn.execute("ALTER TABLE course_with_section_width RENAME TO course")
    conn.commit()
    conn.execute("PRAGMA foreign_keys = ON")


def remove_legacy_course_duration(conn, engine):
    if engine == "mysql":
        columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM course").fetchall()}
        if "duration_years" in columns:
            conn.execute("ALTER TABLE course DROP COLUMN duration_years")
        return

    columns = {row[1] for row in conn.execute("PRAGMA table_info(course)").fetchall()}
    if "duration_years" in columns:
        conn.execute("ALTER TABLE course DROP COLUMN duration_years")


def remove_legacy_course_is_active(conn, engine):
    if engine == "mysql":
        columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM course").fetchall()}
        if "is_active" in columns:
            conn.execute("ALTER TABLE course DROP COLUMN is_active")
        return

    columns = {row[1] for row in conn.execute("PRAGMA table_info(course)").fetchall()}
    if "is_active" in columns:
        conn.execute("ALTER TABLE course DROP COLUMN is_active")


def ensure_subject_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS subject ("
            "id INT NOT NULL AUTO_INCREMENT, name VARCHAR(255) NOT NULL UNIQUE, "
            "PRIMARY KEY (id))"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS course_subject ("
            "course_id INT NOT NULL, subject_id INT NOT NULL, "
            "PRIMARY KEY (course_id, subject_id), "
            "FOREIGN KEY (course_id) REFERENCES course(id) ON DELETE CASCADE, "
            "FOREIGN KEY (subject_id) REFERENCES subject(id) ON DELETE CASCADE)"
        )
        existing_columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM lab_session").fetchall()}
        if "subject_id" not in existing_columns:
            conn.execute("ALTER TABLE lab_session ADD COLUMN subject_id INT NULL")
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS subject ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE)"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS course_subject ("
            "course_id INTEGER NOT NULL, subject_id INTEGER NOT NULL, "
            "PRIMARY KEY (course_id, subject_id), "
            "FOREIGN KEY (course_id) REFERENCES course(id) ON DELETE CASCADE, "
            "FOREIGN KEY (subject_id) REFERENCES subject(id) ON DELETE CASCADE)"
        )
        existing_columns = {row[1] for row in conn.execute("PRAGMA table_info(lab_session)").fetchall()}
        if "subject_id" not in existing_columns:
            conn.execute("ALTER TABLE lab_session ADD COLUMN subject_id INTEGER")

    courses = conn.execute(
        "SELECT id, subject FROM course WHERE subject IS NOT NULL AND subject != ''"
    ).fetchall()
    for course in courses:
        for subject_name in str(course["subject"] if engine == "mysql" else course[1]).split("|"):
            subject_name = subject_name.strip()
            if not subject_name:
                continue
            if engine == "mysql":
                conn.execute("INSERT IGNORE INTO subject (name) VALUES (?)", (subject_name,))
            else:
                conn.execute("INSERT OR IGNORE INTO subject (name) VALUES (?)", (subject_name,))
            subject_row = conn.execute("SELECT id FROM subject WHERE name = ?", (subject_name,)).fetchone()
            course_id = course["id"] if engine == "mysql" else course[0]
            subject_id = subject_row["id"] if engine == "mysql" else subject_row[0]
            if engine == "mysql":
                conn.execute(
                    "INSERT IGNORE INTO course_subject (course_id, subject_id) VALUES (?, ?)",
                    (course_id, subject_id),
                )
            else:
                conn.execute(
                    "INSERT OR IGNORE INTO course_subject (course_id, subject_id) VALUES (?, ?)",
                    (course_id, subject_id),
                )

def ensure_auth_session_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS auth_session ("
            "id INT NOT NULL AUTO_INCREMENT, user_id INT NOT NULL, "
            "token_hash CHAR(64) NOT NULL UNIQUE, device_identifier CHAR(64) NOT NULL, "
            "login_time DATETIME NOT NULL, last_activity DATETIME NOT NULL, "
            "status VARCHAR(20) NOT NULL DEFAULT 'ACTIVE', logout_time DATETIME NULL, "
            "expires_at DATETIME NOT NULL, PRIMARY KEY (id), "
            "FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE)"
        )
        session_columns = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM auth_session").fetchall()}
        if "device_id" not in session_columns:
            conn.execute("ALTER TABLE auth_session ADD COLUMN device_id INT NULL")
        conn.execute(
            "CREATE TABLE IF NOT EXISTS active_user_session ("
            "user_id INT NOT NULL, auth_session_id INT NOT NULL UNIQUE, "
            "PRIMARY KEY (user_id), FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (auth_session_id) REFERENCES auth_session(id) ON DELETE CASCADE)"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS auth_session ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, "
            "token_hash TEXT NOT NULL UNIQUE, device_identifier TEXT NOT NULL, "
            "login_time TEXT NOT NULL, last_activity TEXT NOT NULL, "
            "status TEXT NOT NULL DEFAULT 'ACTIVE', logout_time TEXT, "
            "expires_at TEXT NOT NULL, FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE)"
        )
        session_columns = {row[1] for row in conn.execute("PRAGMA table_info(auth_session)").fetchall()}
        if "device_id" not in session_columns:
            conn.execute("ALTER TABLE auth_session ADD COLUMN device_id INTEGER")
        conn.execute(
            "CREATE TABLE IF NOT EXISTS active_user_session ("
            "user_id INTEGER PRIMARY KEY, auth_session_id INTEGER NOT NULL UNIQUE, "
            "FOREIGN KEY (user_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (auth_session_id) REFERENCES auth_session(id) ON DELETE CASCADE)"
        )


def ensure_device_status_audit_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS device_status_audit ("
            "id INT NOT NULL AUTO_INCREMENT, device_id INT NOT NULL, student_id INT NULL, "
            "previous_status VARCHAR(20) NOT NULL, new_status VARCHAR(20) NOT NULL, "
            "action VARCHAR(20) NOT NULL, admin_id INT NOT NULL, admin_name VARCHAR(255) NOT NULL, "
            "reason TEXT, action_time DATETIME NOT NULL, ip_address VARCHAR(45) NOT NULL, "
            "PRIMARY KEY (id), INDEX idx_device_audit_device (device_id), "
            "FOREIGN KEY (admin_id) REFERENCES user(id))"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS device_status_audit ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, device_id INTEGER NOT NULL, student_id INTEGER, "
            "previous_status TEXT NOT NULL, new_status TEXT NOT NULL, action TEXT NOT NULL, "
            "admin_id INTEGER NOT NULL, admin_name TEXT NOT NULL, reason TEXT, "
            "action_time TEXT NOT NULL, ip_address TEXT NOT NULL, "
            "FOREIGN KEY (admin_id) REFERENCES user(id))"
        )


def ensure_admin_permission_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_permission ("
            "id INT NOT NULL AUTO_INCREMENT, admin_id INT NOT NULL UNIQUE, permissions JSON NOT NULL, "
            "updated_by INT NOT NULL, updated_at DATETIME NOT NULL, PRIMARY KEY (id), "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_permission_audit ("
            "id INT NOT NULL AUTO_INCREMENT, super_admin_id INT NOT NULL, target_admin_id INT NOT NULL, "
            "module_name VARCHAR(80) NOT NULL, permission_name VARCHAR(40) NOT NULL, "
            "old_value TINYINT(1) NOT NULL DEFAULT 0, new_value TINYINT(1) NOT NULL DEFAULT 0, "
            "changed_at DATETIME NOT NULL, PRIMARY KEY (id), "
            "FOREIGN KEY (super_admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (target_admin_id) REFERENCES user(id) ON DELETE CASCADE)"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_permission ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, admin_id INTEGER NOT NULL UNIQUE, "
            "permissions TEXT NOT NULL DEFAULT '{}', updated_by INTEGER NOT NULL, updated_at TEXT NOT NULL, "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_permission_audit ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, super_admin_id INTEGER NOT NULL, "
            "target_admin_id INTEGER NOT NULL, module_name TEXT NOT NULL, permission_name TEXT NOT NULL, "
            "old_value INTEGER NOT NULL DEFAULT 0, new_value INTEGER NOT NULL DEFAULT 0, "
            "changed_at TEXT NOT NULL, "
            "FOREIGN KEY (super_admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (target_admin_id) REFERENCES user(id) ON DELETE CASCADE)"
        )


def ensure_admin_page_permission_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_page_permission ("
            "id INT NOT NULL AUTO_INCREMENT, admin_id INT NOT NULL, page_key VARCHAR(80) NOT NULL, "
            "is_enabled TINYINT(1) NOT NULL DEFAULT 1, updated_by INT NULL, updated_at DATETIME NOT NULL, "
            "PRIMARY KEY (id), UNIQUE KEY uq_admin_page_permission (admin_id, page_key), "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS admin_page_permission ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, admin_id INTEGER NOT NULL, page_key TEXT NOT NULL, "
            "is_enabled INTEGER NOT NULL DEFAULT 1, updated_by INTEGER, updated_at TEXT NOT NULL, "
            "UNIQUE (admin_id, page_key), "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE, "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )


def ensure_page_visibility_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS page_visibility ("
            "id INT NOT NULL AUTO_INCREMENT, page_name VARCHAR(100) NOT NULL UNIQUE, "
            "page_url VARCHAR(255) NOT NULL, is_visible TINYINT(1) NOT NULL DEFAULT 1, "
            "updated_by INT NULL, updated_at DATETIME NOT NULL, PRIMARY KEY (id), "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS page_visibility ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, page_name TEXT NOT NULL UNIQUE, "
            "page_url TEXT NOT NULL, is_visible INTEGER NOT NULL DEFAULT 1, "
            "updated_by INTEGER, updated_at TEXT NOT NULL, "
            "FOREIGN KEY (updated_by) REFERENCES user(id) ON DELETE SET NULL)"
        )

    for page_name, page_url in PAGE_VISIBILITY_DEFAULTS:
        if not conn.execute(
            "SELECT id FROM page_visibility WHERE page_name = ?", (page_name,)
        ).fetchone():
            conn.execute(
                "INSERT INTO page_visibility (page_name, page_url, is_visible, updated_at) "
                "VALUES (?, ?, 1, ?)",
                (page_name, page_url, utcnow_iso()),
            )


def ensure_login_activity_schema(conn, engine):
    if engine == "mysql":
        conn.execute(
            "CREATE TABLE IF NOT EXISTS login_activity ("
            "id INT NOT NULL AUTO_INCREMENT, admin_id INT NOT NULL, "
            "device_type VARCHAR(30) NOT NULL, browser VARCHAR(100) NOT NULL, "
            "operating_system VARCHAR(100) NOT NULL, ip_address VARCHAR(45) NOT NULL, "
            "session_id CHAR(64) NULL, login_status VARCHAR(30) NOT NULL, "
            "login_time DATETIME NOT NULL, logout_time DATETIME NULL, "
            "PRIMARY KEY (id), INDEX idx_login_activity_session (session_id), "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE)"
        )
    else:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS login_activity ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, admin_id INTEGER NOT NULL, "
            "device_type TEXT NOT NULL, browser TEXT NOT NULL, "
            "operating_system TEXT NOT NULL, ip_address TEXT NOT NULL, "
            "session_id TEXT, login_status TEXT NOT NULL, "
            "login_time TEXT NOT NULL, logout_time TEXT, "
            "FOREIGN KEY (admin_id) REFERENCES user(id) ON DELETE CASCADE)"
        )


def init_db(db_path, config=None):
    config = config or {}
    engine = (config.get("DB_ENGINE") or "sqlite").lower()

    if engine == "mysql":
        database = config.get("DB_NAME", "lab_portal")
        server_conn = get_mysql_conn(config)
        server_conn.cursor().execute(
            f"CREATE DATABASE IF NOT EXISTS `{database.replace('`', '``')}` "
            "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
        server_conn.close()
        conn = get_conn(db_path, config)
        for statement in [s.strip() for s in SCHEMA_MYSQL.split(";") if s.strip()]:
            conn.execute(statement)
        conn.commit()
    else:
        conn = get_conn(db_path, config)
        conn.executescript(SCHEMA_SQLITE)
        conn.commit()

    ensure_course_columns(conn, engine)
    remove_legacy_course_section(conn, engine)
    ensure_course_section_width(conn, engine)
    remove_legacy_course_duration(conn, engine)
    remove_legacy_course_is_active(conn, engine)
    ensure_user_columns(conn, engine)
    ensure_device_columns(conn, engine)
    ensure_subject_schema(conn, engine)
    ensure_auth_session_schema(conn, engine)
    ensure_login_activity_schema(conn, engine)
    ensure_device_status_audit_schema(conn, engine)
    ensure_admin_permission_schema(conn, engine)
    ensure_admin_page_permission_schema(conn, engine)
    ensure_page_visibility_schema(conn, engine)
    conn.commit()
    legacy_admin = conn.execute(
        "SELECT id FROM user WHERE email = ? AND role = 'admin' ORDER BY id LIMIT 1",
        ("admin@example.com",),
    ).fetchone()
    new_admin_email = "admin@ayu.com"
    email_in_use = conn.execute(
        "SELECT id FROM user WHERE email = ?", (new_admin_email,)
    ).fetchone()
    if legacy_admin and not email_in_use:
        conn.execute(
            "UPDATE user SET email = ? WHERE id = ?",
            (new_admin_email, legacy_admin["id"]),
        )
        conn.commit()

    existing_admin = conn.execute("SELECT id FROM user WHERE role = 'admin'").fetchone()
    if not existing_admin:
        conn.execute(
            "INSERT INTO user (name, email, password_hash, role, course_id, created_at) "
            "VALUES (%s, %s, %s, 'admin', NULL, %s)" if engine == "mysql" else
            "INSERT INTO user (name, email, password_hash, role, course_id, created_at) "
            "VALUES (?, ?, ?, 'admin', NULL, ?)",
            ("Admin", new_admin_email, generate_password_hash("admin123"), utcnow_iso()),
        )
        conn.commit()
        print(f"Seeded default admin -> email: {new_admin_email}  password: admin123 (CHANGE THIS)")

    super_admin_email = "ayush5075mi@gmail.com"
    super_admin = conn.execute(
        "SELECT id FROM user WHERE email = ?", (super_admin_email,)
    ).fetchone()
    if not super_admin:
        conn.execute(
            "INSERT INTO user (name, email, password_hash, role, super_admin, course_id, created_at) "
            "VALUES (?, ?, ?, 'admin', 1, NULL, ?)",
            ("Ayush", super_admin_email, generate_password_hash("india123"), utcnow_iso()),
        )
        conn.commit()
        print("Seeded super admin -> email: ayush5075mi@gmail.com")
    else:
        conn.execute(
            "UPDATE user SET role = 'admin', super_admin = 1 WHERE id = ?",
            (super_admin["id"] if engine == "mysql" else super_admin[0],),
        )
        conn.commit()

    conn.close()


def utcnow_iso():
    # MySQL DATETIME and SQLite text comparisons both support this UTC format.
    return datetime.now(timezone.utc).replace(tzinfo=None).strftime("%Y-%m-%d %H:%M:%S")
