from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import os
from dateutil.relativedelta import relativedelta
import traceback
from openpyxl.utils import get_column_letter
from report_generator import ReportGenerator

app = Flask(__name__)

# Database configuration
DB_CONFIG = {
    "host": "localhost",
    "database": "hikvision",
    "user": "postgres",
    "password": "postgres",
    "port": "5432",
}

report_generator = ReportGenerator(DB_CONFIG)


@app.route("/")
def index():
    return render_template("index.html")


def validate_dates(start_str, end_str):
    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date()

        if start_date > end_date:
            raise ValueError("Start date cannot be after end date")

        return start_date, end_date
    except ValueError as e:
        raise ValueError(f"Invalid date format: {str(e)}")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        app.logger.info(f"Received form data: {request.form}")

        if "report_type" not in request.form:
            return jsonify({"error": "Report type not specified"}), 400

        report_type = request.form["report_type"]
        today = datetime.now().date()

        # Handle date range selection
        if report_type == "custom":
            if "start_date" not in request.form or "end_date" not in request.form:
                return (
                    jsonify(
                        {
                            "error": "Start date and end date are required for custom reports"
                        }
                    ),
                    400,
                )

            try:
                start_date, end_date = validate_dates(
                    request.form["start_date"], request.form["end_date"]
                )
            except ValueError as e:
                return jsonify({"error": str(e)}), 400
        else:
            if report_type == "daily":
                start_date = end_date = today
            elif report_type == "weekly":
                start_date = today - timedelta(days=today.weekday())
                end_date = start_date + timedelta(days=6)
            elif report_type == "monthly":
                start_date = today.replace(day=1)
                end_date = today.replace(day=1) + relativedelta(months=1, days=-1)
            elif report_type == "quarterly":
                current_quarter = (today.month - 1) // 3
                quarter_month = current_quarter * 3 + 1
                start_date = today.replace(month=quarter_month, day=1)
                end_date = start_date + relativedelta(months=3, days=-1)
            else:
                return jsonify({"error": f"Invalid report type: {report_type}"}), 400

        app.logger.info(f"Date range: {start_date} to {end_date}")

        additional_params = request.form.to_dict()
        output_file = report_generator.generate_report(
            report_type, start_date, end_date, additional_params
        )

        # Handle work timetable file if provided
        if "use_timetable" in request.form and "work_timetable" in request.files:
            timetable_file = request.files["work_timetable"]
            if timetable_file:
                # Save uploaded file temporarily
                temp_path = os.path.join(os.getcwd(), "temp", timetable_file.filename)
                os.makedirs(os.path.dirname(temp_path), exist_ok=True)
                timetable_file.save(temp_path)
                additional_params["work_timetable"] = temp_path

                try:
                    output_file = report_generator.generate_report(
                        report_type, start_date, end_date, additional_params
                    )
                finally:
                    # Clean up temporary file
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
            else:
                return jsonify({"error": "Work timetable file is required"}), 400
        else:
            output_file = report_generator.generate_report(
                report_type, start_date, end_date, additional_params
            )

        if not os.path.exists(output_file):
            app.logger.error(f"Generated file not found: {output_file}")
            return jsonify({"error": "Report generation failed"}), 500

        return send_file(output_file, as_attachment=True)

    except Exception as e:
        app.logger.error(f"Error generating report: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == "__main__":
    app.debug = True
    app.run()
