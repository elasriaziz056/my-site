from flask import (
    Flask,
    request,
    jsonify,
    send_file,
    after_this_request,
)

from flask_cors import CORS

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

import logging
import os
import re
import time
import uuid
import zipfile


# ============================================================
# APP CONFIGURATION
# ============================================================

app = Flask(__name__)

# Keep this if your frontend is hosted separately.
CORS(app)

# Maximum incoming request size.
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO
)

logger = logging.getLogger(__name__)


# ============================================================
# DIRECTORIES
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

TEMPLATES_DIR = os.path.join(
    BASE_DIR,
    "templates"
)

STAMPS_DIR = os.path.join(
    BASE_DIR,
    "stamps"
)

GENERATED_DIR = os.path.join(
    BASE_DIR,
    "generated"
)

os.makedirs(
    GENERATED_DIR,
    exist_ok=True
)


# ============================================================
# CONSTANTS
# ============================================================

TEMPLATE_FILES = [
    "vttemplate.docx",
    "uttemplate.docx",
]

DEFAULT_OUTPUT_NAMES = [
    "vt.docx",
    "ut.docx",
]

PLACEHOLDER_PATTERN = r"(\{\{.*?\}\})"

# Delete temporary files older than one hour.
TEMP_FILE_MAX_AGE = 60 * 60


# ============================================================
# CLEAN OLD GENERATED FILES
# ============================================================

def cleanup_old_files(
    directory,
    max_age_seconds=TEMP_FILE_MAX_AGE
):
    """
    Remove old temporary files.

    This handles files left behind because of:
        - server restart
        - interrupted request
        - browser disconnect
        - unexpected exception
    """

    now = time.time()

    try:

        for filename in os.listdir(directory):

            file_path = os.path.join(
                directory,
                filename
            )

            # Only remove files.
            if not os.path.isfile(file_path):
                continue

            try:

                file_age = (
                    now
                    - os.path.getmtime(file_path)
                )

                if file_age > max_age_seconds:

                    os.remove(
                        file_path
                    )

                    logger.info(
                        "Deleted old temporary file: %s",
                        file_path
                    )

            except OSError:

                logger.exception(
                    "Could not delete old file: %s",
                    file_path
                )

    except OSError:

        logger.exception(
            "Could not scan generated directory: %s",
            directory
        )


# ============================================================
# SAFE FILE DELETION
# ============================================================

def delete_file_safely(
    file_path
):
    """
    Delete a temporary file only if it is inside
    GENERATED_DIR.
    """

    if not file_path:
        return

    try:

        generated_directory = os.path.abspath(
            GENERATED_DIR
        )

        absolute_file_path = os.path.abspath(
            file_path
        )

        # Make sure the file is inside generated/.
        if not absolute_file_path.startswith(
            generated_directory + os.sep
        ):

            logger.warning(
                "Refused to delete file outside "
                "generated directory: %s",
                file_path
            )

            return

        if os.path.isfile(
            absolute_file_path
        ):

            os.remove(
                absolute_file_path
            )

            logger.info(
                "Deleted temporary file: %s",
                absolute_file_path
            )

    except OSError:

        logger.exception(
            "Could not delete temporary file: %s",
            file_path
        )


# ============================================================
# PLACEHOLDER PROCESSING
# ============================================================

def process_placeholders(
    paragraphs: list,
    data_dict: dict,
    replaced_keys: set,
    re_pattern=PLACEHOLDER_PATTERN,
    progress_callback=None,
):
    """
    Replace {{placeholders}} in Word documents.

    Supports placeholders split across multiple
    Microsoft Word runs.
    """

    total = len(data_dict)

    def count_replacement(key):

        if key in replaced_keys:
            return

        replaced_keys.add(key)

        completed = len(replaced_keys)

        progress = (
            int(completed * 100 / total)
            if total
            else 100
        )

        if progress_callback:
            progress_callback(progress)

    compiled_pattern = re.compile(
        re_pattern
    )

    for paragraph in paragraphs:

        runs = paragraph.runs

        if not runs:
            continue

        run_texts = [
            run.text or ""
            for run in runs
        ]

        cumulative_lengths = []

        total_length = 0

        for text in run_texts:

            cumulative_lengths.append(
                total_length
            )

            total_length += len(text)

        cumulative_lengths.append(
            total_length
        )

        full_text = "".join(
            run_texts
        )

        matches = list(
            compiled_pattern.finditer(
                full_text
            )
        )

        if not matches:
            continue

        # Process backwards so indexes stay valid.
        for match in reversed(matches):

            key = match.group(1).strip()

            # Unknown placeholder.
            if key not in data_dict:
                continue

            # None value.
            if data_dict[key] is None:
                continue

            # =================================================
            # STAMP
            # =================================================

            if key == "{{stamp}}":

                try:

                    stamp_name = str(
                        data_dict[key]
                    ).strip().lower()

                    # Remove directory components.
                    stamp_name = os.path.basename(
                        stamp_name
                    )

                    if not stamp_name:
                        raise ValueError(
                            "Empty stamp name"
                        )

                    # Only PNG stamps.
                    if not stamp_name.endswith(
                        ".png"
                    ):
                        stamp_name += ".png"

                    stamp_path = os.path.join(
                        STAMPS_DIR,
                        stamp_name
                    )

                    stamp_directory = os.path.abspath(
                        STAMPS_DIR
                    )

                    absolute_stamp_path = (
                        os.path.abspath(
                            stamp_path
                        )
                    )

                    # Prevent path traversal.
                    if not absolute_stamp_path.startswith(
                        stamp_directory + os.sep
                    ):
                        raise ValueError(
                            "Invalid stamp path"
                        )

                    if not os.path.isfile(
                        stamp_path
                    ):
                        raise FileNotFoundError(
                            f"Stamp not found: {stamp_name}"
                        )

                    # Remove placeholder.
                    paragraph.text = ""

                    run = paragraph.add_run()

                    run.add_picture(
                        stamp_path,
                        width=Inches(2),
                        height=Inches(0.5),
                    )

                    count_replacement(
                        key
                    )

                except Exception:

                    logger.exception(
                        "Failed to insert stamp"
                    )

                    # Preserve original behavior.
                    paragraph.text = ""

                continue

            # =================================================
            # NORMAL PLACEHOLDER
            # =================================================

            replacement = str(
                data_dict[key]
            )

            start = match.start()
            end = match.end()

            # -------------------------------------------------
            # Find starting run
            # -------------------------------------------------

            start_run_idx = None

            for i in range(
                len(run_texts)
            ):

                if (
                    cumulative_lengths[i]
                    <= start
                    < cumulative_lengths[i + 1]
                ):

                    start_run_idx = i
                    break

            if start_run_idx is None:
                continue

            # -------------------------------------------------
            # Find ending run
            # -------------------------------------------------

            end_run_idx = None

            for i in range(
                len(run_texts)
            ):

                if (
                    cumulative_lengths[i]
                    < end
                    <= cumulative_lengths[i + 1]
                ):

                    end_run_idx = i
                    break

            if end_run_idx is None:

                end_run_idx = (
                    len(run_texts) - 1
                )

            # -------------------------------------------------
            # Calculate offsets
            # -------------------------------------------------

            start_offset = (
                start
                - cumulative_lengths[
                    start_run_idx
                ]
            )

            end_offset = (
                end
                - cumulative_lengths[
                    end_run_idx
                ]
            )

            # -------------------------------------------------
            # Same run
            # -------------------------------------------------

            if start_run_idx == end_run_idx:

                run = runs[
                    start_run_idx
                ]

                original = (
                    run.text or ""
                )

                new_text = (
                    original[:start_offset]
                    + replacement
                    + original[end_offset:]
                )

                if new_text != original:

                    run.text = new_text

                    count_replacement(
                        key
                    )

            # -------------------------------------------------
            # Multiple runs
            # -------------------------------------------------

            else:

                start_run = runs[
                    start_run_idx
                ]

                end_run = runs[
                    end_run_idx
                ]

                prefix = (
                    start_run.text[
                        :start_offset
                    ]
                )

                suffix = (
                    end_run.text[
                        end_offset:
                    ]
                )

                start_run.text = (
                    prefix + replacement
                )

                for i in range(
                    start_run_idx + 1,
                    end_run_idx
                ):

                    runs[i].text = ""

                end_run.text = suffix

                count_replacement(
                    key
                )


# ============================================================
# TABLE CLEANING
# ============================================================

def clean_tables_row(
    doc,
    count,
    identifier
):

    to_replace = [
        "<d2>",
        "<d1>",
        "<d3>",
        "<d4>",
    ]

    to_remove = []

    # ========================================================
    # DETERMINE TABLES TO REMOVE
    # ========================================================

    if identifier == 1:

        if count < 5:

            to_remove.extend([
                "<d2>",
                "<d1>",
                "<d3>",
                "<d4>",
            ])

        elif count < 17:

            to_remove.extend([
                "<d2>",
                "<d3>",
                "<d4>",
            ])

        if count < 29:

            to_remove.extend([
                "<d3>",
                "<d4>",
            ])

        elif count < 41:

            to_remove.append(
                "<d4>"
            )

    elif identifier == 2:

        if count < 5:

            to_remove.extend([
                "<d2>",
                "<d1>",
            ])

        elif count < 17:

            to_remove.append(
                "<d2>"
            )

    # ========================================================
    # REMOVE / MODIFY TABLES
    # ========================================================

    for table in list(doc.tables):

        if not table.rows:
            continue

        cell = table.rows[0].cells[0]

        cell_text = cell.text

        # ----------------------------------------------------
        # Remove table
        # ----------------------------------------------------

        should_remove = any(
            placeholder in cell_text
            for placeholder in to_remove
        )

        if should_remove:

            table_element = table._element

            table_element.getparent().remove(
                table_element
            )

            continue

        # ----------------------------------------------------
        # Replace table placeholders
        # ----------------------------------------------------

        for paragraph in cell.paragraphs:

            full_text = "".join(
                run.text or ""
                for run in paragraph.runs
            )

            replaced = False

            for placeholder in to_replace:

                if placeholder in full_text:

                    full_text = full_text.replace(
                        placeholder,
                        "repere"
                    )

                    replaced = True

            if replaced:

                for run in paragraph.runs:
                    run.text = ""

                if paragraph.runs:

                    new_run = (
                        paragraph.runs[0]
                    )

                else:

                    new_run = (
                        paragraph.add_run()
                    )

                new_run.text = full_text
                new_run.bold = True

                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

    # ========================================================
    # REMOVE ROWS WITH PLACEHOLDERS
    # ========================================================

    rows_to_clear = [
        "{{",
        "}}",
        "{",
        "}",
    ]

    for table in doc.tables:

        for i in range(
            len(table.rows) - 1,
            -1,
            -1
        ):

            row = table.rows[i]

            first_cell_text = (
                row.cells[0].text
            )

            if any(
                trigger in first_cell_text
                for trigger in rows_to_clear
            ):

                row_element = row._element

                row_element.getparent().remove(
                    row_element
                )

    # ========================================================
    # FORMAT LAST PARAGRAPHS
    # ========================================================

    last_paragraphs = doc.paragraphs[-2:]

    for paragraph in last_paragraphs:

        paragraph_format = (
            paragraph.paragraph_format
        )

        paragraph_format.line_spacing = 1
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)

        if not paragraph.runs:
            paragraph.add_run()

        for run in paragraph.runs:

            run.font.size = Pt(1)

            if run.text:
                run.text = run.text.strip()


# ============================================================
# GET ALL DOCUMENT PARAGRAPHS
# ============================================================

def get_doc_paragraphs(doc):

    document_paragraphs = []

    def scan(container):

        for paragraph in container.paragraphs:

            document_paragraphs.append(
                paragraph
            )

        for table in container.tables:

            for row in table.rows:

                for cell in row.cells:

                    scan(cell)

    scan(doc)

    # Headers and footers.
    for section in doc.sections:

        scan(section.header)

        scan(section.footer)

    return document_paragraphs


# ============================================================
# SAFE OUTPUT FILENAME
# ============================================================

def safe_output_filename(
    filename
):
    """
    Convert user filename into a safe .docx filename.

    IMPORTANT:
    This function does NOT add a UUID or request ID.

    Example:

        ../../my report.docx

    becomes:

        my_report.docx
    """

    filename = str(
        filename
    ).strip()

    # Remove directory components.
    filename = os.path.basename(
        filename
    )

    # Keep only safe filename characters.
    filename = re.sub(
        r"[^A-Za-z0-9._-]",
        "_",
        filename
    )

    # Add .docx if necessary.
    if not filename.lower().endswith(
        ".docx"
    ):

        filename += ".docx"

    # Prevent empty filename.
    if filename == ".docx":

        filename = "document.docx"

    return filename


# ============================================================
# PROCESS DOCUMENTS
# ============================================================

def process_doc(
    output_names,
    data,
    count,
    request_id
):
    """
    Generate DOCX files.

    There are TWO names for every generated document:

    1. server_path
       Private temporary filename.

       Example:
           generated/
           83f91ab_vt.docx

    2. download_name
       Clean filename used inside ZIP.

       Example:
           vt.docx

    The request ID is NEVER used as the ZIP filename.
    """

    replaced_keys = set()

    generated_files = []

    for i, (
        template_name,
        download_name
    ) in enumerate(
        zip(
            TEMPLATE_FILES,
            output_names
        ),
        start=1
    ):

        # ----------------------------------------------------
        # Template
        # ----------------------------------------------------

        template_path = os.path.join(
            TEMPLATES_DIR,
            template_name
        )

        if not os.path.isfile(
            template_path
        ):

            raise FileNotFoundError(
                f"Template not found: "
                f"{template_name}"
            )

        # ----------------------------------------------------
        # PRIVATE SERVER FILENAME
        # ----------------------------------------------------
        #
        # UUID is ONLY used here.
        #
        # The user will never see this filename.
        #

        server_filename = (
            f"{request_id}_{download_name}"
        )

        server_path = os.path.join(
            GENERATED_DIR,
            server_filename
        )

        # ----------------------------------------------------
        # Open template
        # ----------------------------------------------------

        doc = Document(
            template_path
        )

        # ----------------------------------------------------
        # Replace placeholders
        # ----------------------------------------------------

        process_placeholders(
            get_doc_paragraphs(doc),
            data,
            replaced_keys
        )

        # ----------------------------------------------------
        # Clean tables
        # ----------------------------------------------------

        clean_tables_row(
            doc,
            count,
            i
        )

        # ----------------------------------------------------
        # Save private server file
        # ----------------------------------------------------

        doc.save(
            server_path
        )

        # ----------------------------------------------------
        # Keep both names
        # ----------------------------------------------------

        generated_files.append({
            "server_path": server_path,
            "download_name": download_name,
        })

    return generated_files


# ============================================================
# API - GENERATE REPORTS
# ============================================================

@app.route(
    "/generate",
    methods=["POST"]
)
def generate():

    generated_files = []

    zip_path = None

    # Unique ID used ONLY for server-side temporary files.
    request_id = uuid.uuid4().hex

    try:

        # ====================================================
        # FALLBACK CLEANUP
        # ====================================================

        cleanup_old_files(
            GENERATED_DIR,
            TEMP_FILE_MAX_AGE
        )

        # ====================================================
        # READ JSON
        # ====================================================

        request_data = request.get_json(
            silent=True
        )

        if not request_data:

            return jsonify({
                "error": (
                    "Invalid or missing JSON data."
                )
            }), 400

        # ====================================================
        # GET DATA
        # ====================================================

        data = request_data.get(
            "data",
            {}
        )

        count = request_data.get(
            "count"
        )

        output_names = request_data.get(
            "outputnames",
            DEFAULT_OUTPUT_NAMES
        )

        # ====================================================
        # VALIDATE DATA
        # ====================================================

        if not isinstance(
            data,
            dict
        ):

            return jsonify({
                "error": (
                    "'data' must be an object."
                )
            }), 400

        # ====================================================
        # VALIDATE COUNT
        # ====================================================

        try:

            count = int(
                count
            )

        except (
            TypeError,
            ValueError
        ):

            return jsonify({
                "error": (
                    "'count' must be a valid integer."
                )
            }), 400

        if count < 0:

            return jsonify({
                "error": (
                    "'count' cannot be negative."
                )
            }), 400

        if count > 1000:

            return jsonify({
                "error": (
                    "'count' is too large."
                )
            }), 400

        # ====================================================
        # VALIDATE OUTPUT NAMES
        # ====================================================

        if not isinstance(
            output_names,
            list
        ):

            return jsonify({
                "error": (
                    "'outputnames' must be a list."
                )
            }), 400

        if len(output_names) != 2:

            return jsonify({
                "error": (
                    "Exactly two output filenames "
                    "are required."
                )
            }), 400

        # ----------------------------------------------------
        # Sanitize filenames.
        #
        # IMPORTANT:
        # safe_output_filename() does NOT add UUID.
        # ----------------------------------------------------

        safe_names = [
            safe_output_filename(name)
            for name in output_names
        ]

        # ====================================================
        # GENERATE DOCUMENTS
        # ====================================================

        generated_files = process_doc(
            safe_names,
            data,
            count,
            request_id
        )

        # ====================================================
        # CREATE PRIVATE ZIP
        # ====================================================

        # UUID exists ONLY in server filename.
        zip_name = (
            f"{request_id}_reports.zip"
        )

        zip_path = os.path.join(
            GENERATED_DIR,
            zip_name
        )

        with zipfile.ZipFile(
            zip_path,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zip_file:

            for file_info in generated_files:

                # --------------------------------------------
                # Actual private file on server.
                # --------------------------------------------

                server_path = (
                    file_info["server_path"]
                )

                # --------------------------------------------
                # Clean filename inside ZIP.
                # --------------------------------------------

                download_name = (
                    file_info["download_name"]
                )

                zip_file.write(
                    server_path,
                    arcname=download_name
                )

        # ====================================================
        # CLEANUP AFTER RESPONSE
        # ====================================================

        @after_this_request
        def cleanup_after_response(
            response
        ):

            # Delete temporary DOCX files.
            for file_info in generated_files:

                delete_file_safely(
                    file_info["server_path"]
                )

            # Delete private ZIP.
            delete_file_safely(
                zip_path
            )

            return response

        # ====================================================
        # SEND ZIP
        # ====================================================
        #
        # IMPORTANT:
        #
        # download_name is explicitly "reports.zip".
        #
        # Therefore the browser sees:
        #
        #     reports.zip
        #
        # NOT:
        #
        #     UUID_reports.zip
        #

        return send_file(
            zip_path,
            as_attachment=True,
            download_name="reports.zip",
            mimetype="application/zip"
        )

    except Exception:

        # ====================================================
        # LOG REAL ERROR
        # ====================================================

        logger.exception(
            "Error generating reports. "
            "Request ID: %s",
            request_id
        )

        # ====================================================
        # CLEANUP AFTER ERROR
        # ====================================================

        for file_info in generated_files:

            delete_file_safely(
                file_info.get(
                    "server_path"
                )
            )

        delete_file_safely(
            zip_path
        )

        # ====================================================
        # SAFE ERROR RESPONSE
        # ====================================================

        return jsonify({
            "error": (
                "An internal server error occurred "
                "while generating the reports."
            ),
            "request_id": request_id
        }), 500


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route(
    "/",
    methods=["GET"]
)
def home():

    return app.send_static_file("index.html")


# ============================================================
# LOCAL DEVELOPMENT
# ============================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )
