import os
import re
from uuid import uuid4

import mysql.connector
from docx import Document
from dotenv import load_dotenv
from flask import (
    Flask,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from mysql.connector import Error
from pypdf import PdfReader
from werkzeug.utils import secure_filename


# Load values from .env
load_dotenv()

app = Flask(__name__)

app.secret_key = os.getenv(
    "FLASK_SECRET_KEY",
    "scholarship-project-secret-key"
)


# --------------------------------------------------
# Upload configuration
# --------------------------------------------------

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

ALLOWED_EXTENSIONS = {
    "pdf",
    "docx",
}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Maximum file size: 10 MB
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# --------------------------------------------------
# Database connection
# --------------------------------------------------

def get_db_connection():
    try:
        connection = mysql.connector.connect(
            host=os.getenv("MYSQL_HOST", "localhost"),
            user=os.getenv("MYSQL_USER", "root"),
            password=os.getenv("MYSQL_PASSWORD"),
            database=os.getenv(
                "MYSQL_DATABASE",
                "scholarship_db"
            ),
            port=int(os.getenv("MYSQL_PORT", "3306")),
        )

        return connection

    except Error as error:
        print("Database connection error:", error)
        return None


# --------------------------------------------------
# Helper functions
# --------------------------------------------------

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


def extract_text_from_pdf(file_path):
    text = ""

    reader = PdfReader(file_path)

    for page in reader.pages:
        page_text = page.extract_text()

        if page_text:
            text += page_text + "\n"

    return text


def extract_text_from_docx(file_path):
    document = Document(file_path)

    paragraphs = [
        paragraph.text
        for paragraph in document.paragraphs
        if paragraph.text.strip()
    ]

    return "\n".join(paragraphs)


def extract_document_text(file_path):
    extension = file_path.rsplit(".", 1)[1].lower()

    if extension == "pdf":
        return extract_text_from_pdf(file_path)

    if extension == "docx":
        return extract_text_from_docx(file_path)

    return ""


def extract_income(text):
    patterns = [
        r"annual\s+income\s*[:\-]?\s*(?:rs\.?|₹)?\s*([\d,]+)",
        r"family\s+income\s*[:\-]?\s*(?:rs\.?|₹)?\s*([\d,]+)",
        r"income\s*[:\-]?\s*(?:rs\.?|₹)?\s*([\d,]+)",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:
            return float(
                match.group(1).replace(",", "")
            )

    return None


def extract_category(text):
    match = re.search(
        r"\b(SC|ST|OBC|GENERAL|GEN)\b",
        text,
        re.IGNORECASE,
    )

    if match:
        category = match.group(1).upper()

        if category == "GEN":
            return "GENERAL"

        return category

    return None


def extract_ews_status(text):
    yes_patterns = [
        r"\bEWS\s*[:\-]?\s*YES\b",
        r"\bECONOMICALLY WEAKER SECTION\b",
        r"\bEWS CERTIFICATE\b",
    ]

    no_patterns = [
        r"\bEWS\s*[:\-]?\s*NO\b",
        r"\bNOT EWS\b",
    ]

    for pattern in yes_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return "Yes"

    for pattern in no_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return "No"

    return None


def extract_percentage(text):
    patterns = [
        r"percentage\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*%",
        r"marks\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*%",
        r"(\d+(?:\.\d+)?)\s*%",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:
            return float(match.group(1))

    return None


# --------------------------------------------------
# Home
# --------------------------------------------------

@app.route("/")
def home():
    return render_template("index.html")


# --------------------------------------------------
# Registration
# --------------------------------------------------

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        full_name = request.form.get(
            "name",
            ""
        ).strip()

        email = request.form.get(
            "email",
            ""
        ).strip()

        mobile = request.form.get(
            "phone",
            ""
        ).strip()

        student_password = request.form.get(
            "password",
            ""
        )

        gender = request.form.get(
            "gender",
            ""
        )

        category = request.form.get(
            "category",
            ""
        )

        ews_status = request.form.get(
            "ews_status",
            ""
        )

        family_income = request.form.get(
            "annual_income"
        )

        marks_10 = request.form.get(
            "tenth_percentage"
        )

        marks_12 = request.form.get(
            "twelfth_percentage"
        )

        state = request.form.get(
            "state",
            ""
        ).strip()

        district = request.form.get(
            "district",
            ""
        ).strip()

        connection = get_db_connection()

        if connection is None:
            return "Database connection failed."

        cursor = connection.cursor()

        try:
            sql = """
                INSERT INTO students (
                    full_name,
                    email,
                    mobile,
                    password,
                    gender,
                    category,
                    EWS_STATUS,
                    family_income,
                    marks_10,
                    marks_12,
                    state,
                    district
                )
                VALUES (
                    %s, %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s, %s, %s
                )
            """

            values = (
                full_name,
                email,
                mobile,
                student_password,
                gender,
                category,
                ews_status,
                family_income,
                marks_10,
                marks_12,
                state,
                district,
            )

            cursor.execute(sql, values)
            connection.commit()

            return redirect(url_for("login"))

        except Error as error:
            connection.rollback()
            print("Registration error:", error)

            return f"Registration failed: {error}"

        finally:
            cursor.close()
            connection.close()

    return render_template("register.html")


# --------------------------------------------------
# Login
# --------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        student_password = request.form.get("password", "")

        connection = get_db_connection()
        cursor = None

        if connection is None:
            return "Database connection failed."

        try:
            cursor = connection.cursor(
                dictionary=True,
                buffered=True
            )

            sql = """
                SELECT
                    id,
                    full_name,
                    email,
                    password
                FROM students
                WHERE email = %s
                  AND password = %s
                LIMIT 1
            """

            cursor.execute(
                sql,
                (email, student_password)
            )

            student = cursor.fetchone()

            if student:
                session["student_id"] = student["id"]
                session["student_name"] = student["full_name"]

                return redirect(url_for("dashboard"))

            return render_template(
                "login.html",
                message="Invalid email or password."
            )

        except Error as error:
            print("Login error:", error)

            return render_template(
                "login.html",
                message="An error occurred during login."
            )

        finally:
            if cursor is not None:
                cursor.close()

            if connection is not None and connection.is_connected():
                connection.close()

    return render_template("login.html")


# --------------------------------------------------
# Dashboard
# --------------------------------------------------

@app.route("/dashboard")
def dashboard():
    if "student_id" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()

    if connection is None:
        return "Database connection failed."

    cursor = None

    try:
        cursor = connection.cursor(
            dictionary=True,
            buffered=True
        )

        sql = """
            SELECT
                id,
                full_name,
                email,
                mobile,
                gender,
                category,
                EWS_STATUS AS ews_status,
                family_income,
                marks_10,
                marks_12,
                state,
                district
            FROM students
            WHERE id = %s
        """

        cursor.execute(
            sql,
            (session["student_id"],)
        )

        student = cursor.fetchone()

        if student is None:
            session.clear()
            return redirect(url_for("login"))

        return render_template(
            "dashboard.html",
            student=student
        )

    except Error as error:
        print("Dashboard error:", error)
        return "Unable to load dashboard."

    finally:
        if cursor is not None:
            cursor.close()

        if connection.is_connected():
            connection.close()

# --------------------------------------------------
# Eligibility using registered information
# --------------------------------------------------

@app.route("/check-eligibility")
def check_eligibility():
    if "student_id" not in session:
        return redirect(url_for("login"))

    connection = get_db_connection()

    if connection is None:
        return "Database connection failed."

    cursor = connection.cursor(
        dictionary=True
    )

    try:
        sql = """
            SELECT
                family_income,
                marks_10,
                marks_12,
                category,
                EWS_STATUS AS ews_status
            FROM students
            WHERE id = %s
        """

        cursor.execute(
            sql,
            (session["student_id"],)
        )

        student = cursor.fetchone()

        if student is None:
            session.clear()
            return redirect(url_for("login"))

        family_income = float(
            student["family_income"]
        )

        marks_10 = float(
            student["marks_10"]
        )

        marks_12 = float(
            student["marks_12"]
        )

        category = (
            student["category"]
            or ""
        ).strip().upper()

        ews_status = (
            student["ews_status"]
            or ""
        ).strip().lower()

        eligible_category = (
            category in ["SC", "ST", "OBC"]
            or ews_status == "yes"
        )

        eligible = (
            family_income <= 250000
            and marks_10 >= 60
            and marks_12 >= 60
            and eligible_category
        )

        if eligible:
            status = "✅ Eligible for Scholarship"

            reason = (
                "You satisfy the income, marks, "
                "category or EWS requirements."
            )

            scholarship_name = (
                "Merit-cum-Means Scholarship"
            )

        else:
            failed_conditions = []

            if family_income > 250000:
                failed_conditions.append(
                    "family income is above ₹2,50,000"
                )

            if marks_10 < 60:
                failed_conditions.append(
                    "10th percentage is below 60%"
                )

            if marks_12 < 60:
                failed_conditions.append(
                    "12th percentage is below 60%"
                )

            if not eligible_category:
                failed_conditions.append(
                    "category or EWS requirement "
                    "is not satisfied"
                )

            status = "❌ Not Eligible"

            reason = (
                "You are not eligible because "
                + ", ".join(failed_conditions)
                + "."
            )

            scholarship_name = None

        recommended_colleges = []

        if eligible:
            cursor.execute(
                """
                SELECT college_name
                FROM colleges
                WHERE min_percentage <= %s
                  AND max_income >= %s
                  AND (
                      category = %s
                      OR category = 'GENERAL'
                  )
                """,
                (
                    marks_12,
                    family_income,
                    category,
                ),
            )

            recommended_colleges = cursor.fetchall()

        return render_template(
            "eligibility_result.html",
            status=status,
            reason=reason,
            scholarship_name=scholarship_name,
            colleges=recommended_colleges,
        )


    except (Error, TypeError, ValueError) as error:
        print("Eligibility error:", error)
        return "Error checking eligibility."

    finally:
        cursor.close()
        connection.close()


# --------------------------------------------------
# Upload and process four documents
# --------------------------------------------------

@app.route(
    "/upload-document",
    methods=["GET", "POST"]
)
def upload_document():
    if "student_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        required_documents = {
            "income_certificate": "Income Certificate",
            "community_certificate": "Community Certificate",
            "mark10": "10th Marksheet",
            "mark12": "12th Marksheet",
        }

        saved_files = {}

        student_upload_folder = os.path.join(
            app.config["UPLOAD_FOLDER"],
            str(session["student_id"]),
        )

        os.makedirs(
            student_upload_folder,
            exist_ok=True,
        )

        for field_name, document_name in required_documents.items():
            if field_name not in request.files:
                return render_template(
                    "upload_document.html",
                    message=f"{document_name} is missing."
                )

            uploaded_file = request.files[field_name]

            if uploaded_file.filename == "":
                return render_template(
                    "upload_document.html",
                    message=(
                        f"Please select {document_name}."
                    )
                )

            if not allowed_file(uploaded_file.filename):
                return render_template(
                    "upload_document.html",
                    message=(
                        f"{document_name} must be "
                        "a PDF or DOCX file."
                    )
                )

            original_filename = secure_filename(
                uploaded_file.filename
            )

            extension = original_filename.rsplit(
                ".",
                1
            )[1].lower()

            unique_filename = (
                f"{field_name}_{uuid4().hex}."
                f"{extension}"
            )

            file_path = os.path.join(
                student_upload_folder,
                unique_filename
            )

            uploaded_file.save(file_path)

            saved_files[field_name] = file_path

        try:
            income_text = extract_document_text(
                saved_files["income_certificate"]
            )

            community_text = extract_document_text(
                saved_files["community_certificate"]
            )

            mark10_text = extract_document_text(
                saved_files["mark10"]
            )

            mark12_text = extract_document_text(
                saved_files["mark12"]
            )

            extracted_income = extract_income(
                income_text
            )

            extracted_category = extract_category(
                community_text
            )

            extracted_ews = extract_ews_status(
                community_text
            )

            extracted_marks_10 = extract_percentage(
                mark10_text
            )

            extracted_marks_12 = extract_percentage(
                mark12_text
            )

            missing_details = []

            if extracted_income is None:
                missing_details.append(
                    "family income"
                )

            if (
                extracted_category is None
                and extracted_ews is None
            ):
                missing_details.append(
                    "category or EWS status"
                )

            if extracted_marks_10 is None:
                missing_details.append(
                    "10th percentage"
                )

            if extracted_marks_12 is None:
                missing_details.append(
                    "12th percentage"
                )

            if missing_details:
                return render_template(
                    "upload_document.html",
                    message=(
                        "Documents were uploaded, but "
                        "these details could not be "
                        "extracted: "
                        + ", ".join(missing_details)
                        + ". Use clear text-based PDF "
                        "or DOCX documents."
                    )
                )

            category_eligible = (
                extracted_category
                in ["SC", "ST", "OBC"]
                or extracted_ews == "Yes"
            )

            eligible = (
                extracted_income <= 250000
                and extracted_marks_10 >= 60
                and extracted_marks_12 >= 60
                and category_eligible
            )

            if eligible:
                status = (
                    "✅ Eligible for Scholarship"
                )

                reason = (
                    "The uploaded documents satisfy "
                    "the income, academic marks, "
                    "category or EWS requirements."
                )

                scholarship_name = (
                    "Merit-cum-Means Scholarship"
                )

            else:
                failed_conditions = []

                if extracted_income > 250000:
                    failed_conditions.append(
                        "family income is above "
                        "₹2,50,000"
                    )

                if extracted_marks_10 < 60:
                    failed_conditions.append(
                        "10th percentage is below 60%"
                    )

                if extracted_marks_12 < 60:
                    failed_conditions.append(
                        "12th percentage is below 60%"
                    )
                if not category_eligible:
                    failed_conditions.append(
                        "category or EWS requirement "
                        "is not satisfied"
                    )

                status = "❌ Not Eligible"

                reason = (
                    "Not eligible because "
                    + ", ".join(failed_conditions)
                    + "."
                )

                scholarship_name = None


            # Get recommended colleges
            recommended_colleges = []

            if eligible:
                connection = get_db_connection()
                college_cursor = None

                try:
                    college_cursor = connection.cursor(
                        dictionary=True,
                        buffered=True
                    )

                    college_cursor.execute(
                        """
                        SELECT
                            college_name,
                            state,
                            min_percentage,
                            max_income,
                            category
                        FROM colleges
                        WHERE min_percentage <= %s
                          AND max_income >= %s
                          AND (
                              category = %s
                              OR category = 'GENERAL'
                          )
                        """,
                        (
                            extracted_marks_12,
                            extracted_income,
                            extracted_category,
                        ),
                    )

                    recommended_colleges = (
                        college_cursor.fetchall()
                    )

                except Error as error:
                    print(
                        "College recommendation error:",
                        error
                    )

                finally:
                    if college_cursor is not None:
                        college_cursor.close()

                    if (
                        connection is not None
                        and connection.is_connected()
                    ):
                        connection.close()


            return render_template(
                "eligibility_result.html",
                status=status,
                reason=reason,
                scholarship_name=scholarship_name,
                colleges=recommended_colleges,
            )

        except Exception as error:
            print("Document processing error:", error)

            return render_template(
                "upload_document.html",
                message=(
                    "The documents were uploaded, "
                    "but they could not be processed. "
                    "Use text-based PDF or DOCX files."
                ),
            )

    return render_template("upload_document.html")

@app.route("/apply-scholarship", methods=["POST"])
def apply_scholarship():
    if "student_id" not in session:
        return redirect(url_for("login"))

    college_name = request.form.get(
        "college_name",
        ""
    ).strip()

    scholarship_name = request.form.get(
        "scholarship_name",
        ""
    ).strip()

    if not college_name:
        return "Please select a college."

    connection = get_db_connection()

    if connection is None:
        return "Database connection failed."

    cursor = connection.cursor()

    try:
        cursor.execute(
            """
            INSERT INTO applications (
                student_id,
                college_name,
                scholarship_name,
                status
            )
            VALUES (%s, %s, %s, %s)
            """,
            (
                session["student_id"],
                college_name,
                scholarship_name,
                "Pending",
            ),
        )

        connection.commit()

        application_id = cursor.lastrowid

        return render_template(
    "application_success.html",
    scholarship_name=scholarship_name,
    college_name=college_name,
    application_id=application_id,
)

    except Error as error:
        connection.rollback()
        print("Application error:", error)
        return f"Application failed: {error}"

    finally:
        cursor.close()
        connection.close()
 

# --------------------------------------------------
# Logout
# --------------------------------------------------

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# --------------------------------------------------
# Run application
# --------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True)
