from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, jsonify
import os
import psycopg2
from psycopg2.extras import DictCursor
import pandas as pd
import io
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from werkzeug.security import generate_password_hash, check_password_hash
import threading
import queue
from werkzeug.utils import secure_filename
import time

app = Flask(__name__)
app.secret_key = "resultportal_secret_key_2026_secure"

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ==================== POSTGRESQL DATABASE SETUP ====================
DATABASE_URL = os.environ.get("DATABASE_URL")

# Fallback URL (remove this after setting environment variable on Render)
if not DATABASE_URL:
    DATABASE_URL = "postgresql://result_portal_db_user:3TI0q22twgpg53Dff9hjz8SRLPTCTHtO@dpg-d7s4b9ho3t8c73didi00-a.oregon-postgres.render.com/result_portal_db"

# Fix for Render (sometimes requires sslmode)
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

def get_db_connection():
    """Get database connection"""
    if not DATABASE_URL:
        raise Exception("DATABASE_URL environment variable not set!")
    return psycopg2.connect(DATABASE_URL, sslmode="require")

# ==================== BACKGROUND JOB QUEUE ====================
upload_queue = {}
job_counter = 0

def process_excel_in_background(job_id, filepath, year, semester, upload_type, username):
    """Process Excel file in background"""
    try:
        print(f"🔵 Job {job_id}: Started processing")
        
        # Read entire file at once
        if filepath.endswith('.xlsx'):
            df = pd.read_excel(filepath, engine='openpyxl')
        else:
            df = pd.read_excel(filepath, engine='xlrd')
        
        total_rows = len(df)
        print(f"📊 Job {job_id}: Loaded {total_rows} rows")
        print(f"📊 Job {job_id}: Columns: {list(df.columns)}")
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        inserted_count = 0
        updated_count = 0
        skipped_count = 0
        full_semester = f"{year} {semester}"
        
        # Process all rows
        for index, row in df.iterrows():
            try:
                reg_no = str(row["Reg_No"]).strip()
                if not reg_no or reg_no == "nan":
                    skipped_count += 1
                    continue
                
                name = str(row["Name"]).strip()
                subject_code = str(row["Subject_Code"]).strip()
                subject_name = str(row["Subject_Name"]).strip()
                credits = str(row["Credits"]).strip()
                grade = str(row["Grade"]).strip().upper()
                
                credit_value = credit_to_float(credits)
                grade_point = GRADE_POINTS.get(grade, 0)
                
                # Check if exists
                cursor.execute("""
                    SELECT id FROM results
                    WHERE semester=%s AND reg_no=%s AND subject_code=%s
                """, (full_semester, reg_no, subject_code))
                
                if cursor.fetchone():
                    cursor.execute("""
                        UPDATE results
                        SET name=%s, subject_name=%s, credits=%s, credit_value=%s, grade=%s, grade_point=%s
                        WHERE semester=%s AND reg_no=%s AND subject_code=%s
                    """, (name, subject_name, credits, credit_value, grade, grade_point, full_semester, reg_no, subject_code))
                    updated_count += 1
                else:
                    cursor.execute("""
                        INSERT INTO results (semester, reg_no, name, subject_code, subject_name, credits, credit_value, grade, grade_point)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (full_semester, reg_no, name, subject_code, subject_name, credits, credit_value, grade, grade_point))
                    inserted_count += 1
                
                # Commit every 100 rows
                if (inserted_count + updated_count) % 100 == 0:
                    conn.commit()
                    print(f"🟢 Job {job_id}: Processed {inserted_count + updated_count}/{total_rows} rows")
                    
            except Exception as e:
                print(f"❌ Job {job_id} Row {index} error: {e}")
                skipped_count += 1
                continue
        
        conn.commit()
        
        # Save upload history
        upload_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        cursor.execute("""
            INSERT INTO upload_history (semester, filename, upload_type, inserted, updated, skipped, upload_time, uploaded_by)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (full_semester, os.path.basename(filepath), upload_type, inserted_count, updated_count, skipped_count, upload_time, username))
        
        conn.commit()
        cursor.close()
        conn.close()
        
        # Clean up temp file
        os.remove(filepath)
        
        print(f"✅ Job {job_id}: COMPLETED! Inserted={inserted_count}, Updated={updated_count}, Skipped={skipped_count}")
        
        # Update job status
        upload_queue[job_id] = {
            'status': 'completed',
            'inserted': inserted_count,
            'updated': updated_count,
            'skipped': skipped_count,
            'total': total_rows,
            'semester': full_semester,
            'filename': os.path.basename(filepath)
        }
        
    except Exception as e:
        print(f"❌ Job {job_id} Failed: {e}")
        import traceback
        traceback.print_exc()
        upload_queue[job_id] = {
            'status': 'failed',
            'error': str(e)
        }
        # Clean up temp file if exists
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass

# ------------------ CUTM GRADE POINTS SYSTEM ------------------
GRADE_POINTS = {
    "O": 10,
    "E": 9,
    "A": 8,
    "B": 7,
    "C": 6,
    "D": 5,
    "F": 2,
    "M": 0,
    "S": 0,
    "P": 8,
    "AB": 0,
    "I": 0
}

# ------------------ CREDIT STRING TO FLOAT ------------------
def credit_to_float(credit_str):
    try:
        if pd.isna(credit_str):
            return 0
        parts = str(credit_str).split("+")
        total = 0
        for p in parts:
            total += float(p.strip())
        return total
    except:
        return 0

# ------------------ DATABASE SETUP ------------------
def init_db():
    """Initialize database with all required tables"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Results table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS results (
                id SERIAL PRIMARY KEY,
                semester TEXT NOT NULL,
                reg_no TEXT NOT NULL,
                name TEXT NOT NULL,
                subject_code TEXT NOT NULL,
                subject_name TEXT NOT NULL,
                credits TEXT,
                credit_value REAL,
                grade TEXT,
                grade_point REAL,
                uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Add unique constraint for upsert operations
        cursor.execute("""
            DO $$ 
            BEGIN 
                IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'unique_result') THEN
                    ALTER TABLE results ADD CONSTRAINT unique_result UNIQUE (semester, reg_no, subject_code);
                END IF;
            END $$;
        """)

        # Admins table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS admins (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_login TIMESTAMP
            )
        """)

        # Upload history table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS upload_history (
                id SERIAL PRIMARY KEY,
                semester TEXT,
                filename TEXT,
                upload_type TEXT,
                inserted INTEGER,
                updated INTEGER,
                skipped INTEGER,
                upload_time TEXT,
                uploaded_by TEXT
            )
        """)

        # Notices table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS notices (
                id SERIAL PRIMARY KEY,
                title TEXT NOT NULL,
                message TEXT NOT NULL,
                posted_on TEXT,
                posted_by TEXT
            )
        """)

        # Students table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id SERIAL PRIMARY KEY,
                reg_no TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                batch INTEGER,
                branch TEXT,
                photo TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Check if admins table is empty
        cursor.execute("SELECT COUNT(*) FROM admins")
        count = cursor.fetchone()[0]
        
        if count == 0:
            # Create default admin
            hashed_pass = generate_password_hash("Aaham@8990")
            cursor.execute("INSERT INTO admins (username, password) VALUES (%s, %s)", ("aaham_18", hashed_pass))
            print("✅ Default admin created: aaham_18 / Aaham@8990")
        else:
            print(f"✅ Admins table has {count} record(s)")

        conn.commit()
        cursor.close()
        conn.close()
        print("✅ Database initialized successfully")
        
    except Exception as e:
        print(f"❌ Database initialization error: {e}")

# Initialize database
init_db()

# ------------------ HOME PAGE ------------------
@app.route("/")
def home():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT title, message, posted_on FROM notices ORDER BY id DESC LIMIT 5")
    notices = cursor.fetchall()

    cursor.close()
    conn.close()
    return render_template("index.html", notices=notices)

# ------------------ ABOUT PAGE ------------------
@app.route("/about")
def about():
    return render_template("about.html")

# ------------------ CONTACT PAGE ------------------
@app.route("/contact")
def contact():
    return render_template("contact.html")

# ------------------ SEARCH RESULT ------------------
@app.route("/search", methods=["POST"])
def search_result():
    reg_no = request.form.get("regd_no", "").strip()
    year = request.form.get("year", "").strip()
    semester = request.form.get("semester", "").strip()

    if not reg_no or not year or not semester:
        flash("❌ Please fill all fields!", "error")
        return redirect(url_for("home"))

    full_semester = f"{year} {semester}"

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT subject_code, subject_name, credits, grade, credit_value, grade_point, name
        FROM results 
        WHERE reg_no=%s AND semester=%s
        ORDER BY subject_code
    """, (reg_no, full_semester))

    rows = cursor.fetchall()

    if not rows:
        cursor.close()
        conn.close()
        flash("❌ Result not found! Please check Registration Number, Year or Semester.", "error")
        return render_template("error.html", message="Result not found! Please check Registration Number, Year or Semester.")

    student_name = rows[0][6]

    # Calculate SGPA
    sem_total_credits = 0
    sem_credit_index = 0

    for r in rows:
        sem_total_credits += float(r[4])
        sem_credit_index += float(r[4]) * float(r[5])

    sgpa = round(sem_credit_index / sem_total_credits, 2) if sem_total_credits > 0 else 0

    # Calculate CGPA
    cursor.execute("""
        SELECT credit_value, grade_point
        FROM results
        WHERE reg_no=%s
    """, (reg_no,))
    all_data = cursor.fetchall()
    cursor.close()
    conn.close()

    total_all_credits = 0
    total_credit_index = 0

    for r in all_data:
        total_all_credits += float(r[0])
        total_credit_index += float(r[0]) * float(r[1])

    cgpa = round(total_credit_index / total_all_credits, 2) if total_all_credits > 0 else 0

    # Check backlog
    backlog_count = 0
    for r in rows:
        if str(r[3]).strip().upper() == "F":
            backlog_count += 1

    status = "PASS" if backlog_count == 0 else f"FAIL ({backlog_count} Backlogs)"

    return render_template(
        "result.html",
        reg_no=reg_no,
        semester=full_semester,
        student_name=student_name,
        rows=rows,
        subject_count=len(rows),
        sem_credits=round(sem_total_credits, 2),
        total_credits=round(total_all_credits, 2),
        sgpa=sgpa,
        cgpa=cgpa,
        status=status,
        backlog_count=backlog_count
    )

# ------------------ DOWNLOAD PDF ------------------
@app.route("/download_pdf/<reg_no>/<semester>")
def download_pdf(reg_no, semester):
    semester = semester.replace("_", " ")

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT subject_code, subject_name, credits, grade, credit_value, grade_point, name
        FROM results
        WHERE reg_no=%s AND semester=%s
        ORDER BY subject_code
    """, (reg_no, semester))

    rows = cursor.fetchall()

    if not rows:
        cursor.close()
        conn.close()
        flash("❌ Result not found for PDF generation!", "error")
        return redirect(url_for("home"))

    student_name = rows[0][6]

    # Calculate SGPA
    sem_total_credits = 0
    sem_credit_index = 0

    for r in rows:
        sem_total_credits += float(r[4])
        sem_credit_index += float(r[4]) * float(r[5])

    sgpa = round(sem_credit_index / sem_total_credits, 2) if sem_total_credits > 0 else 0

    # Calculate CGPA
    cursor.execute("""
        SELECT credit_value, grade_point
        FROM results
        WHERE reg_no=%s
    """, (reg_no,))
    all_data = cursor.fetchall()
    cursor.close()
    conn.close()

    total_all_credits = 0
    total_credit_index = 0

    for r in all_data:
        total_all_credits += float(r[0])
        total_credit_index += float(r[0]) * float(r[1])

    cgpa = round(total_credit_index / total_all_credits, 2) if total_all_credits > 0 else 0

    generated_on = datetime.now().strftime("%d-%m-%Y %I:%M %p")

    # Create PDF
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # Header
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(200, height - 50, "CUTM Result Portal")

    pdf.setFont("Helvetica", 12)
    pdf.drawString(230, height - 70, "Official Marksheet")

    pdf.line(40, height - 85, width - 40, height - 85)

    # Student Info
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, height - 120, "Student Name:")
    pdf.setFont("Helvetica", 12)
    pdf.drawString(160, height - 120, student_name)

    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, height - 140, "Registration No:")
    pdf.setFont("Helvetica", 12)
    pdf.drawString(160, height - 140, reg_no)

    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, height - 160, "Semester:")
    pdf.setFont("Helvetica", 12)
    pdf.drawString(160, height - 160, semester)

    # SGPA & CGPA
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, height - 190, "SGPA:")
    pdf.setFont("Helvetica", 12)
    pdf.drawString(110, height - 190, str(sgpa))

    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(200, height - 190, "CGPA:")
    pdf.setFont("Helvetica", 12)
    pdf.drawString(260, height - 190, str(cgpa))

    # Table Header
    y = height - 240
    pdf.setFont("Helvetica-Bold", 11)

    pdf.setFillColorRGB(0.2, 0.2, 0.2)
    pdf.rect(40, y, width - 80, 20, fill=True, stroke=False)

    pdf.setFillColorRGB(1, 1, 1)
    pdf.drawString(50, y + 5, "Subject Code")
    pdf.drawString(150, y + 5, "Subject Name")
    pdf.drawString(420, y + 5, "Credits")
    pdf.drawString(500, y + 5, "Grade")

    # Table Data
    y -= 25
    pdf.setFillColorRGB(0, 0, 0)
    pdf.setFont("Helvetica", 10)

    for r in rows:
        pdf.drawString(50, y, str(r[0]))
        pdf.drawString(150, y, str(r[1])[:45])
        pdf.drawString(430, y, str(r[2]))
        pdf.drawString(505, y, str(r[3]))
        y -= 20

        if y < 80:
            pdf.showPage()
            y = height - 60

    # Footer
    pdf.line(40, 70, width - 40, 70)
    pdf.setFont("Helvetica", 10)
    pdf.drawString(50, 55, f"Generated on: {generated_on}")
    pdf.drawString(50, 40, "This is a computer generated marksheet.")

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(420, 40, "Controller of Examination")

    pdf.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{reg_no}_{semester.replace(' ', '_')}_marksheet.pdf",
        mimetype="application/pdf"
    )

# ------------------ ADMIN LOGIN ------------------
@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    """Admin login page"""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if not username or not password:
            flash("❌ Please enter username and password!", "error")
            return render_template("admin_login.html")

        try:
            conn = get_db_connection()
            cursor = conn.cursor()

            # Get admin from database
            cursor.execute("SELECT id, username, password FROM admins WHERE username=%s", (username,))
            admin = cursor.fetchone()

            if admin:
                # Verify password
                if check_password_hash(admin[2], password):
                    # Update last login
                    now = datetime.now().strftime("%d-%m-%Y %I:%M %p")
                    cursor.execute("UPDATE admins SET last_login=%s WHERE id=%s", (now, admin[0]))
                    conn.commit()

                    # Set session
                    session["admin"] = True
                    session["admin_id"] = admin[0]
                    session["admin_user"] = admin[1]
                    session["last_login"] = now

                    flash(f"✅ Welcome back, {admin[1]}!", "success")
                    cursor.close()
                    conn.close()
                    return redirect(url_for("admin_dashboard"))
                else:
                    flash("❌ Invalid password!", "error")
            else:
                flash("❌ Username not found!", "error")

            cursor.close()
            conn.close()

        except Exception as e:
            flash(f"❌ Login error: {str(e)}", "error")

    return render_template("admin_login.html")

# ------------------ ADMIN DASHBOARD (WITH BACKGROUND PROCESSING) ------------------
@app.route("/admin/dashboard", methods=["GET", "POST"])
def admin_dashboard():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    if request.method == "POST":
        try:
            year = request.form.get("year")
            semester = request.form.get("semester")
            upload_type = request.form.get("upload_type")
            file = request.files.get("file")

            if not year or not semester or not upload_type:
                flash("⚠️ Please select all fields!", "warning")
                return redirect(url_for("admin_dashboard"))

            if not file or file.filename == '':
                flash("⚠️ Please select a file!", "warning")
                return redirect(url_for("admin_dashboard"))

            if not (file.filename.endswith(".xls") or file.filename.endswith(".xlsx")):
                flash("⚠️ Please upload only .xls or .xlsx file!", "warning")
                return redirect(url_for("admin_dashboard"))

            # Save file temporarily
            filename = secure_filename(f"{datetime.now().timestamp()}_{file.filename}")
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)
            
            # Quick validation - check columns
            try:
                if filepath.endswith('.xlsx'):
                    df_check = pd.read_excel(filepath, nrows=1, engine='openpyxl')
                else:
                    df_check = pd.read_excel(filepath, nrows=1, engine='xlrd')
                
                required_cols = ["Reg_No", "Name", "Subject_Code", "Subject_Name", "Credits", "Grade"]
                missing_cols = [col for col in required_cols if col not in df_check.columns]
                
                if missing_cols:
                    os.remove(filepath)
                    flash(f"❌ Missing columns: {', '.join(missing_cols)}. Found: {list(df_check.columns)}", "error")
                    return redirect(url_for("admin_dashboard"))
                    
            except Exception as e:
                os.remove(filepath)
                flash(f"❌ Error reading file: {str(e)}", "error")
                return redirect(url_for("admin_dashboard"))
            
            # Start background job
            global job_counter
            job_counter += 1
            job_id = job_counter
            
            username = session.get("admin_user", "Unknown")
            
            upload_queue[job_id] = {'status': 'processing', 'started_at': datetime.now().strftime("%H:%M:%S")}
            
            thread = threading.Thread(
                target=process_excel_in_background,
                args=(job_id, filepath, year, semester, upload_type, username)
            )
            thread.daemon = True
            thread.start()
            
            flash(f"✅ File uploaded! Processing started (Job #{job_id}). Total {len(df_check)} rows will be processed in background. Check 'Upload History' later for status.", "success")
            
        except Exception as e:
            flash(f"❌ Error: {str(e)}", "error")
        
        return redirect(url_for("admin_dashboard"))
    
    return render_template("admin_dashboard.html")

# ------------------ JOB STATUS API ------------------
@app.route("/admin/job_status/<int:job_id>")
def job_status(job_id):
    if "admin" not in session:
        return jsonify({"error": "Unauthorized"}), 401
    
    if job_id in upload_queue:
        return jsonify(upload_queue[job_id])
    else:
        return jsonify({"status": "not_found"})

# ------------------ DELETE SEMESTER ------------------
@app.route("/admin/delete_semester", methods=["POST"])
def delete_semester():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    year = request.form.get("delete_year")
    semester = request.form.get("semester")

    if not year or not semester:
        flash("⚠️ Please select Year and Semester!", "warning")
        return redirect(url_for("admin_dashboard"))

    full_semester = f"{year} {semester}"

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM results WHERE semester=%s", (full_semester,))
    deleted_rows = cursor.rowcount

    conn.commit()
    cursor.close()
    conn.close()

    flash(f"✅ {full_semester} deleted successfully! Deleted Records: {deleted_rows}", "success")
    return redirect(url_for("admin_dashboard"))

# ------------------ UPLOAD HISTORY ------------------
@app.route("/admin/upload_history")
def upload_history():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT semester, filename, upload_type, inserted, updated, skipped, upload_time, uploaded_by
        FROM upload_history
        ORDER BY id DESC
        LIMIT 100
    """)
    history = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template("upload_history.html", history=history)

# ------------------ ADD NOTICE ------------------
@app.route("/admin/add_notice", methods=["GET", "POST"])
def add_notice():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        message = request.form.get("message", "").strip()
        posted_on = datetime.now().strftime("%d-%m-%Y %I:%M %p")
        posted_by = session.get("admin_user", "Admin")

        if not title or not message:
            flash("❌ Please fill all fields!", "error")
            return redirect(url_for("add_notice"))

        if len(title) < 5:
            flash("❌ Title must be at least 5 characters!", "error")
            return redirect(url_for("add_notice"))

        if len(message) < 10:
            flash("❌ Message must be at least 10 characters!", "error")
            return redirect(url_for("add_notice"))

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("INSERT INTO notices (title, message, posted_on, posted_by) VALUES (%s, %s, %s, %s)",
                      (title, message, posted_on, posted_by))

        conn.commit()
        cursor.close()
        conn.close()

        flash("✅ Notice added successfully!", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template("add_notice.html")

# ------------------ VIEW NOTICES ------------------
@app.route("/admin/view_notices")
def view_notices():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, title, message, posted_on, posted_by FROM notices ORDER BY id DESC")
    notices = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template("view_notices.html", notices=notices)

# ------------------ DELETE NOTICE ------------------
@app.route("/admin/delete_notice/<int:notice_id>")
def delete_notice(notice_id):
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM notices WHERE id=%s", (notice_id,))
    deleted = cursor.rowcount

    conn.commit()
    cursor.close()
    conn.close()

    if deleted > 0:
        flash("✅ Notice deleted successfully!", "success")
    else:
        flash("❌ Notice not found!", "error")

    return redirect(url_for("view_notices"))

# ------------------ ADD ADMIN ------------------
@app.route("/admin/add_admin", methods=["GET", "POST"])
def add_admin():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    # Only main admin can add new admins
    if session.get("admin_user") != "aaham_18":
        flash("❌ Only main admin can add new admins!", "error")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if not username or not password:
            flash("❌ Please fill all fields!", "error")
            return redirect(url_for("add_admin"))

        if len(username) < 3:
            flash("❌ Username must be at least 3 characters!", "error")
            return redirect(url_for("add_admin"))

        if len(password) < 6:
            flash("❌ Password must be at least 6 characters!", "error")
            return redirect(url_for("add_admin"))

        hashed_pass = generate_password_hash(password)

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            cursor.execute("INSERT INTO admins (username, password) VALUES (%s, %s)",
                          (username, hashed_pass))
            conn.commit()
            flash(f"✅ Admin '{username}' created successfully!", "success")
        except psycopg2.IntegrityError:
            flash(f"❌ Username '{username}' already exists!", "error")
        finally:
            cursor.close()
            conn.close()

        return redirect(url_for("admin_dashboard"))

    return render_template("add_admin.html")

# ------------------ CHANGE PASSWORD ------------------
@app.route("/admin/change_password", methods=["GET", "POST"])
def change_password():
    if "admin" not in session:
        flash("⚠️ Please login first!", "warning")
        return redirect(url_for("admin_login"))

    if request.method == "POST":
        try:
            old_password = request.form.get("old_password", "").strip()
            new_password = request.form.get("new_password", "").strip()
            confirm_password = request.form.get("confirm_password", "").strip()

            if not old_password or not new_password or not confirm_password:
                flash("❌ All fields are required!", "error")
                return redirect(url_for("change_password"))

            if new_password != confirm_password:
                flash("❌ New passwords do not match!", "error")
                return redirect(url_for("change_password"))

            if len(new_password) < 6:
                flash("❌ Password must be at least 6 characters long!", "error")
                return redirect(url_for("change_password"))

            if new_password == old_password:
                flash("❌ New password cannot be same as old password!", "error")
                return redirect(url_for("change_password"))

            username = session.get("admin_user")

            if not username:
                session.pop("admin", None)
                flash("⚠️ Session expired. Please login again.", "warning")
                return redirect(url_for("admin_login"))

            conn = get_db_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT password FROM admins WHERE username=%s", (username,))
            admin = cursor.fetchone()

            if not admin:
                cursor.close()
                conn.close()
                session.pop("admin", None)
                session.pop("admin_user", None)
                flash("❌ Admin not found! Please login again.", "error")
                return redirect(url_for("admin_login"))

            if not check_password_hash(admin[0], old_password):
                cursor.close()
                conn.close()
                flash("❌ Current password is incorrect!", "error")
                return redirect(url_for("change_password"))

            hashed_new = generate_password_hash(new_password)
            cursor.execute("UPDATE admins SET password=%s WHERE username=%s", (hashed_new, username))
            conn.commit()

            if cursor.rowcount > 0:
                flash("✅ Password changed successfully! Please login with new password.", "success")
                session.pop("admin", None)
                session.pop("admin_user", None)
                cursor.close()
                conn.close()
                return redirect(url_for("admin_login"))
            else:
                flash("❌ Failed to update password. Please try again.", "error")

            cursor.close()
            conn.close()

        except Exception as e:
            flash(f"❌ Error: {str(e)}", "error")

        return redirect(url_for("change_password"))

    return render_template("change_password.html")

# ------------------ ADMIN LOGOUT ------------------
@app.route("/admin/logout")
def admin_logout():
    username = session.get("admin_user", "Admin")
    session.clear()
    flash(f"👋 Goodbye, {username}! Logged out successfully.", "success")
    return redirect(url_for("admin_login"))

# ------------------ ERROR HANDLERS ------------------
@app.errorhandler(404)
def page_not_found(e):
    flash("❌ Page not found!", "error")
    return render_template("error.html", message="The page you are looking for does not exist."), 404

@app.errorhandler(500)
def internal_server_error(e):
    flash("❌ Internal server error!", "error")
    return render_template("error.html", message="Something went wrong. Please try again later."), 500

# ------------------ RUN APPLICATION ------------------
if __name__ == "__main__":
    print("\n" + "="*60)
    print("🚀 CUTM RESULT PORTAL STARTING...")
    print("="*60)
    print("🐘 Using PostgreSQL database")
    print("👤 Default Admin: aaham_18 / Aaham@8990")
    print("="*60 + "\n")
    
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)