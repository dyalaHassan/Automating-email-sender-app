import os
import tempfile
import uuid
from flask_session import Session
import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
)
from utils import (
    extract_docx_html,
    extract_docx_placeholders,
    extract_placeholders,
    render_template_text,
    validate_email,
)

from outlook_sender import (
    get_outlook_application,
    send_email,
)


app = Flask(__name__)

app.secret_key = "one-world-ip-email-sender-secret"

# Store Flask session data on the computer
# instead of trying to put everything inside a browser cookie.
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_USE_SIGNER"] = True

app.config["SESSION_FILE_DIR"] = os.path.join(
    tempfile.gettempdir(),
    "oneworld_email_sessions"
)

os.makedirs(
    app.config["SESSION_FILE_DIR"],
    exist_ok=True
)

Session(app)

UPLOAD_FOLDER = os.path.join(
    tempfile.gettempdir(),
    "oneworld_email_sender"
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================================================
# Helpers
# =========================================================

def get_session_folder():
    if "session_id" not in session:
        session["session_id"] = str(uuid.uuid4())

    folder = os.path.join(
        UPLOAD_FOLDER,
        session["session_id"]
    )

    os.makedirs(folder, exist_ok=True)

    return folder


def load_dataframe():
    excel_path = session.get("excel_path")

    if not excel_path:
        return None

    if not os.path.exists(excel_path):
        return None

    return pd.read_excel(excel_path)


def detect_email_column(df):
    keywords = [
        "email",
        "e-mail",
        "mail",
        "email address",
        "البريد",
        "البريد الإلكتروني",
        "ايميل",
        "إيميل",
    ]

    for col in df.columns:
        name = str(col).strip().lower()

        if any(keyword.lower() in name for keyword in keywords):
            return col

    return None


def get_valid_rows(df, email_col):

    valid_rows = []
    invalid_rows = []

    for index, row in df.iterrows():

        email = str(row.get(email_col, "")).strip()

        if validate_email(email):
            valid_rows.append(index)

        else:
            invalid_rows.append({
                "row": index + 1,
                "email": email,
                "reason": "البريد الإلكتروني غير موجود أو غير صحيح"
            })

    return valid_rows, invalid_rows


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():
    return redirect(url_for("step1"))


# =========================================================
# STEP 1
# Excel
# =========================================================

@app.route("/step1", methods=["GET", "POST"])
def step1():

    if request.method == "POST":

        uploaded_file = request.files.get("excel_file")

        if not uploaded_file or uploaded_file.filename == "":
            flash("يرجى اختيار ملف Excel.", "error")
            return redirect(url_for("step1"))

        try:

            folder = get_session_folder()

            excel_path = os.path.join(
                folder,
                "recipients.xlsx"
            )

            uploaded_file.save(excel_path)

            df = pd.read_excel(excel_path)

            if df.empty:
                flash(
                    "ملف Excel لا يحتوي على بيانات.",
                    "error"
                )
                return redirect(url_for("step1"))

            session["excel_path"] = excel_path

            email_col = detect_email_column(df)

            if email_col:
                session["email_col"] = str(email_col)
            else:
                session.pop("email_col", None)

            session["columns"] = [
                str(c) for c in df.columns
            ]

            session["total_rows"] = len(df)

            return redirect(url_for("step2"))

        except Exception as e:

            flash(
                f"تعذر قراءة ملف Excel: {e}",
                "error"
            )

    return render_template(
        "step1_excel.html",
        current_step=1
    )


# =========================================================
# STEP 2
# Word template
# =========================================================

@app.route("/step2", methods=["GET", "POST"])
def step2():

    df = load_dataframe()

    if df is None:
        return redirect(url_for("step1"))

    columns = [str(column) for column in df.columns]

    # ---------------------------------------------------------
    # Check whether a Word file was already successfully saved
    # during this session.
    # ---------------------------------------------------------

    saved_word_path = session.get("word_path")

    word_already_uploaded = (
        saved_word_path is not None
        and os.path.exists(saved_word_path)
    )

    uploaded_word_name = session.get(
        "word_file_name",
        ""
    )

    # ---------------------------------------------------------
    # User pressed NEXT
    # ---------------------------------------------------------

    if request.method == "POST":

        new_word_file = request.files.get(
            "word_file"
        )

        # =====================================================
        # CASE 1:
        # User selected a NEW Word file
        # =====================================================

        if (
            new_word_file
            and new_word_file.filename
        ):

            if not new_word_file.filename.lower().endswith(
                ".docx"
            ):
                flash(
                    "يرجى رفع ملف Word بصيغة DOCX.",
                    "error"
                )

                return redirect(
                    url_for("step2")
                )

            try:

                folder = get_session_folder()

                word_path = os.path.join(
                    folder,
                    "email_template.docx"
                )

                # Save the uploaded Word document.
                new_word_file.save(
                    word_path
                )

                # Store its information in the Flask session.
                session["word_path"] = word_path

                session["word_file_name"] = (
                    new_word_file.filename
                )

            except Exception as error:

                flash(
                    f"تعذر حفظ ملف Word: {error}",
                    "error"
                )

                return redirect(
                    url_for("step2")
                )

        # =====================================================
        # CASE 2:
        # No new file selected, but we already have one saved.
        # =====================================================

        elif word_already_uploaded:

            word_path = saved_word_path

        # =====================================================
        # CASE 3:
        # No new file and no previously saved file.
        # =====================================================

        else:

            flash(
                "يرجى اختيار ملف Word أولاً.",
                "error"
            )

            return redirect(
                url_for("step2")
            )

        # =====================================================
        # Process the Word file
        # =====================================================

        try:

            html_body = extract_docx_html(
                word_path
            )

            placeholders = extract_docx_placeholders(
                word_path
            )

            # -------------------------------------------------
            # Compare Word placeholders with Excel columns.
            # Ignore spaces and capitalization.
            # -------------------------------------------------

            invalid_placeholders = [
                placeholder
                for placeholder in placeholders
                if not any(
                    str(placeholder).strip().lower()
                    ==
                    str(column).strip().lower()
                    for column in columns
                )
            ]

            if invalid_placeholders:

                flash(
                    "توجد حقول داخل ملف Word "
                    "غير موجودة في ملف Excel: "
                    +
                    ", ".join(
                        f"<<{placeholder}>>"
                        for placeholder
                        in invalid_placeholders
                    ),
                    "error"
                )

                # IMPORTANT:
                # We DON'T delete the saved Word file.
                # User can correct another issue without
                # needing to upload the file repeatedly.
                return redirect(
                    url_for("step2")
                )

            # -------------------------------------------------
            # Everything is valid.
            # -------------------------------------------------

            session["html_body"] = html_body

            session["placeholders"] = list(
                placeholders
            )

            return redirect(
                url_for("step3")
            )

        except Exception as error:

            flash(
                f"تعذر قراءة ملف Word: {error}",
                "error"
            )

            return redirect(
                url_for("step2")
            )

    # ---------------------------------------------------------
    # Display Step 2
    # ---------------------------------------------------------

    return render_template(
        "step2_word.html",
        current_step=2,
        columns=columns,
        word_already_uploaded=word_already_uploaded,
        uploaded_word_name=uploaded_word_name
    )
# =========================================================
# STEP 3
# Settings
# =========================================================

@app.route("/step3", methods=["GET", "POST"])
def step3():

    df = load_dataframe()

    if df is None:
        return redirect(url_for("step1"))

    if not session.get("html_body"):
        return redirect(url_for("step2"))

    columns = [str(c) for c in df.columns]

    detected_email_col = session.get(
        "email_col"
    )

    if request.method == "POST":

        email_col = request.form.get(
            "email_col"
        )

        subject = request.form.get(
            "subject",
            ""
        ).strip()

        cc_address = request.form.get(
            "cc_address",
            ""
        ).strip()

        if not email_col:
            flash(
                "اختر عمود البريد الإلكتروني.",
                "error"
            )
            return redirect(url_for("step3"))

        if not subject:
            flash(
                "يرجى كتابة عنوان الرسالة.",
                "error"
            )
            return redirect(url_for("step3"))

        subject_fields = extract_placeholders(
            subject
        )

        missing_fields = [
    field
    for field in subject_fields
    if not any(
        str(field).strip().lower()
        == str(column).strip().lower()
        for column in columns
    )
]

        if missing_fields:

            flash(
                "يوجد حقل في عنوان الرسالة "
                "غير موجود في Excel: "
                + ", ".join(
                    f"<<{field}>>"
                    for field in missing_fields
                ),
                "error"
            )

            return redirect(url_for("step3"))

        attachment_files = request.files.getlist(
            "attachments"
        )

        attachment_paths = []

        folder = get_session_folder()

        attachments_folder = os.path.join(
            folder,
            "attachments"
        )

        os.makedirs(
            attachments_folder,
            exist_ok=True
        )

        for uploaded_file in attachment_files:

            if uploaded_file and uploaded_file.filename:

                path = os.path.join(
                    attachments_folder,
                    uploaded_file.filename
                )

                uploaded_file.save(path)

                attachment_paths.append(path)

        session["email_col"] = email_col
        session["subject"] = subject
        session["cc_address"] = cc_address
        session["attachment_paths"] = attachment_paths

        return redirect(url_for("step4"))

    return render_template(
        "step3_settings.html",
        current_step=3,
        columns=columns,
        detected_email_col=detected_email_col
    )


# =========================================================
# STEP 4
# Preview
# =========================================================

@app.route("/step4")
def step4():

    df = load_dataframe()

    if df is None:
        return redirect(url_for("step1"))

    email_col = session.get("email_col")

    if not email_col:
        return redirect(url_for("step3"))

    valid_rows, invalid_rows = get_valid_rows(
        df,
        email_col
    )

    if not valid_rows:

        flash(
            "لا توجد عناوين بريد إلكتروني صحيحة.",
            "error"
        )

        return redirect(url_for("step3"))

    preview_position = request.args.get(
        "row",
        0,
        type=int
    )

    if preview_position < 0:
        preview_position = 0

    if preview_position >= len(valid_rows):
        preview_position = len(valid_rows) - 1

    row_index = valid_rows[
        preview_position
    ]

    row = df.loc[row_index]

    subject = render_template_text(
        session["subject"],
        row
    )

    html_body = render_template_text(
        session["html_body"],
        row
    )

    recipient = str(
        row[email_col]
    ).strip()

    return render_template(
        "step4_preview.html",
        current_step=4,
        recipient=recipient,
        cc_address=session.get(
            "cc_address",
            ""
        ),
        subject=subject,
        html_body=html_body,
        position=preview_position,
        total=len(valid_rows),
        previous_position=max(
            0,
            preview_position - 1
        ),
        next_position=min(
            len(valid_rows) - 1,
            preview_position + 1
        )
    )


# =========================================================
# STEP 5
# Send selection
# =========================================================

@app.route("/step5", methods=["GET", "POST"])
def step5():

    df = load_dataframe()

    if df is None:
        return redirect(url_for("step1"))

    email_col = session.get("email_col")

    valid_rows, invalid_rows = get_valid_rows(
        df,
        email_col
    )

    if request.method == "POST":

        send_type = request.form.get(
            "send_type"
        )

        rows_to_send = valid_rows.copy()

        if send_type == "first":

            try:
                number = int(
                    request.form.get(
                        "first_n",
                        1
                    )
                )

                rows_to_send = valid_rows[
                    :number
                ]

            except ValueError:
                pass

        elif send_type == "range":

            try:

                start = int(
                    request.form.get(
                        "start_row",
                        1
                    )
                )

                end = int(
                    request.form.get(
                        "end_row",
                        len(df)
                    )
                )

                rows_to_send = [
                    index
                    for index in valid_rows
                    if start - 1 <= index <= end - 1
                ]

            except ValueError:
                pass

        if not rows_to_send:

            flash(
                "لم يتم اختيار أي مستلم.",
                "error"
            )

            return redirect(url_for("step5"))

        session["rows_to_send"] = [
            int(x)
            for x in rows_to_send
        ]

        return render_template(
    "confirm_send.html",
    current_step=5,
    selected_count=len(rows_to_send)
)

    return render_template(
        "step5_send.html",
        current_step=5,
        valid_count=len(valid_rows),
        invalid_count=len(invalid_rows),
        invalid_rows=invalid_rows,
        total_rows=len(df)
    )


# =========================================================
# SEND
# =========================================================

@app.route("/send", methods=["POST"])
def send_emails():

    df = load_dataframe()

    if df is None:
        return redirect(url_for("step1"))

    rows = session.get(
        "rows_to_send",
        []
    )

    if not rows:

        flash(
            "لا توجد رسائل جاهزة للإرسال.",
            "error"
        )

        return redirect(url_for("step5"))

    email_col = session["email_col"]

    try:

        outlook = get_outlook_application()

    except Exception as e:

        flash(
            "تعذر الاتصال ببرنامج Outlook. "
            f"{e}",
            "error"
        )

        return redirect(url_for("step5"))

    sent_count = 0
    failed_count = 0

    log = []

    for index in rows:

        row = df.loc[index]

        recipient = str(
            row[email_col]
        ).strip()

        subject = render_template_text(
            session["subject"],
            row
        )

        html_body = render_template_text(
            session["html_body"],
            row
        )

        success, error = send_email(
            outlook_app=outlook,
            to_address=recipient,
            cc_address=session.get(
                "cc_address"
            ),
            subject=subject,
            html_body=html_body,
            attachment_paths=session.get(
                "attachment_paths",
                []
            )
        )

        if success:

            sent_count += 1

            status = "تم الإرسال"
            reason = ""

        else:

            failed_count += 1

            status = "فشل"
            reason = error

        log.append({
            "row": index + 1,
            "email": recipient,
            "status": status,
            "reason": reason
        })

    session["send_log"] = log

    return render_template(
        "result.html",
        current_step=5,
        sent_count=sent_count,
        failed_count=failed_count,
        total=len(rows),
        log=log
    )


# =========================================================
# Reset
# =========================================================

@app.route("/reset")
def reset():

    session.clear()

    return redirect(
        url_for("step1")
    )


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )