import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime
import os
from openpyxl.utils import get_column_letter
import traceback
from filter import Filter
import threading, time

class ReportGenerator:
    def __init__(self, db_config):
        self.db_config = db_config

    def get_db_connection(self):
        """Create database connection"""
        connection_string = f"postgresql://{self.db_config['user']}:{self.db_config['password']}@{self.db_config['host']}:{self.db_config['port']}/{self.db_config['database']}"
        return create_engine(connection_string)

    def fetch_data(self, start_date, end_date):
        """Fetch data from the database"""
        try:
            engine = self.get_db_connection()

            # Convert dates to strings for SQL query
            start_date_str = start_date.strftime("%Y-%m-%d")
            end_date_str = end_date.strftime("%Y-%m-%d")

            query = text(
                """
                SELECT 
                    id,
                    date_and_time,
                    date,
                    time,
                    device_name,
                    reader_name,
                    person_name,
                    person_group
                FROM users
                WHERE date BETWEEN :start_date AND :end_date
                ORDER BY date_and_time
            """
            )

            # Create parameters dictionary
            params = {"start_date": start_date_str, "end_date": end_date_str}

            df = pd.read_sql_query(query, engine, params=params)
            return df

        except Exception as e:
            print(f"Error fetching data: {str(e)}")
            print(traceback.format_exc())
            raise

    def delete_file_after_delay(self, file_path, delay):
        """Delete the file after a specified delay"""
        def delete_file():
            time.sleep(delay)
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"File {file_path} deleted after {delay} seconds")

        threading.Thread(target=delete_file).start()

    def generate_excel(self, df, report_type, start_date, end_date, additional_params):
        """Generate Excel file from DataFrame"""
        try:
            if additional_params:
                filter_obj = Filter(df)
                df = filter_obj.apply_filters(additional_params)

            start_date_str = start_date.strftime("%Y-%m-%d")
            end_date_str = end_date.strftime("%Y-%m-%d")

            # Generate Excel file
            output_file = os.path.join(
                os.getcwd(),
                "reports",
                f"{report_type}_report_{start_date_str}_{end_date_str}.xlsx",
            )
            os.makedirs(os.path.dirname(output_file), exist_ok=True)

            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Detailed Data", index=False)

                summary = pd.DataFrame(
                    {
                        "Total Records": [len(df)],
                        "Unique Devices": [df["device_name"].nunique()],
                        "Unique Groups": [df["person_group"].nunique()],
                        "Date Range": [f"{start_date_str} to {end_date_str}"],
                        "Report Type": [report_type.capitalize()],
                    }
                )
                summary.to_excel(writer, sheet_name="Summary", index=False)

                # Adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter  # Get the column name
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = max_length + 2
                        if column in [
                            "A",
                            "B",
                            "C",
                        ]:  # Adjust the first three columns more
                            adjusted_width += 5
                        worksheet.column_dimensions[column].width = adjusted_width

            self.delete_file_after_delay(output_file, 5)
            return output_file

        except Exception as e:
            print(f"Error generating Excel file: {str(e)}")
            print(traceback.format_exc())
            raise

    def process_person_group(self, person_group):
        """Process person group by taking last value after '>' and stripping whitespace"""
        # Handle NaN/None values
        if pd.isna(person_group):
            return ""

        # Convert to string to handle any numeric values
        person_group = str(person_group)

        if ">" in person_group:
            return person_group.split(">")[-1].strip()
        return person_group.strip()

    def generate_custom_excel(
        self, df, work_timetable_path, report_type, start_date, end_date
    ):
        """Generate custom Excel file based on work timetable matching"""
        try:
            # Read work timetable
            timetable_df = pd.read_excel(work_timetable_path)

            # Ensure person_group column exists in timetable
            if "person_group" not in timetable_df.columns:
                raise ValueError("Work timetable must contain a 'person_group' column")

            # Process person groups in both dataframes
            df["processed_group"] = df["person_group"].apply(self.process_person_group)
            timetable_df["processed_group"] = timetable_df["person_group"].apply(
                self.process_person_group
            )

            # Remove empty groups from timetable
            timetable_df = timetable_df[timetable_df["processed_group"] != ""]

            # Get unique processed groups from timetable
            timetable_groups = set(timetable_df["processed_group"].unique())

            if not timetable_groups:
                raise ValueError("No valid person groups found in work timetable")

            # Filter main dataframe to only include matching groups
            matched_df = df[df["processed_group"].isin(timetable_groups)]

            if matched_df.empty:
                print(
                    "Warning: No matching records found with the provided work timetable"
                )

            # Generate Excel file
            start_date_str = start_date.strftime("%Y-%m-%d")
            end_date_str = end_date.strftime("%Y-%m-%d")

            output_file = os.path.join(
                os.getcwd(),
                "reports",
                f"custom_{report_type}_report_{start_date_str}_{end_date_str}.xlsx",
            )
            os.makedirs(os.path.dirname(output_file), exist_ok=True)

            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                # Detailed Data sheet - drop the temporary processed_group column
                matched_df.drop("processed_group", axis=1).to_excel(
                    writer, sheet_name="Detailed Data", index=False
                )

                # Summary sheet
                summary = pd.DataFrame(
                    {
                        "Total Records": [len(matched_df)],
                        "Matched Groups": [len(matched_df["processed_group"].unique())],
                        "Total Groups in Timetable": [len(timetable_groups)],
                        "Date Range": [f"{start_date_str} to {end_date_str}"],
                        "Report Type": [f"Custom {report_type.capitalize()}"],
                    }
                )
                summary.to_excel(writer, sheet_name="Summary", index=False)

                # Group Summary sheet
                if not matched_df.empty:
                    group_summary = (
                        matched_df.groupby("person_group")
                        .agg(
                            {
                                "id": "count",
                                "device_name": "nunique",
                                "person_name": "nunique",
                            }
                        )
                        .rename(
                            columns={
                                "id": "Total Records",
                                "device_name": "Unique Devices",
                                "person_name": "Unique Persons",
                            }
                        )
                        .reset_index()
                    )
                    group_summary.to_excel(
                        writer, sheet_name="Group Summary", index=False
                    )
                else:
                    pd.DataFrame(
                        columns=[
                            "person_group",
                            "Total Records",
                            "Unique Devices",
                            "Unique Persons",
                        ]
                    ).to_excel(writer, sheet_name="Group Summary", index=False)

                # Unmatched Groups sheet
                unmatched_groups = (
                    set(df["processed_group"].unique()) - timetable_groups
                )
                if unmatched_groups:
                    unmatched_df = pd.DataFrame(
                        {
                            "Unmatched Group": sorted(list(unmatched_groups)),
                            "Records Count": [
                                len(df[df["processed_group"] == group])
                                for group in sorted(list(unmatched_groups))
                            ],
                        }
                    )
                    unmatched_df.to_excel(
                        writer, sheet_name="Unmatched Groups", index=False
                    )

                # Adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = max_length + 2
                        worksheet.column_dimensions[column].width = adjusted_width

            print(
                f"Custom Excel file generated successfully with {len(matched_df)} matching records"
            )
            self.delete_file_after_delay(output_file, 5)
            return output_file

        except Exception as e:
            print(f"Error generating custom Excel file: {str(e)}")
            print(traceback.format_exc())
            raise

    def generate_report(
        self, report_type, start_date, end_date, additional_params=None
    ):
        """Generate report based on type and additional parameters"""
        df = self.fetch_data(start_date, end_date)

        # Check if custom timetable report is requested
        if additional_params and "work_timetable" in additional_params:
            return self.generate_custom_excel(
                df,
                additional_params["work_timetable"],
                report_type,
                start_date,
                end_date,
            )

        # Default Excel generation
        return self.generate_excel(
            df, report_type, start_date, end_date, additional_params
        )
