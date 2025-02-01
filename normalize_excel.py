import pandas as pd
import os
import re
from pathlib import Path
from datetime import datetime


class TimetableNormalizer:
    def __init__(self):
        self.group_id_counter = 1
        self.shift_mapping = {}

    def normalize_time(self, time_str):
        """Convert various time formats to HH:MM"""
        if pd.isna(time_str) or not time_str:
            return None

        # Handle numeric types first
        if isinstance(time_str, (float, int)):
            hours = int(time_str)
            minutes = int((time_str % 1) * 60)
            return f"{hours:02d}:{minutes:02d}"

        # Process string inputs
        time_str = str(time_str).strip().lower()

        # Handle AM/PM times
        am_pm_match = re.search(r"(am|pm)", time_str)
        if am_pm_match:
            time_str = time_str.replace(am_pm_match.group(), "").strip()
            try:
                time_obj = datetime.strptime(time_str, "%I:%M")
                if am_pm_match.group() == "pm" and time_obj.hour != 12:
                    time_obj = time_obj.replace(hour=time_obj.hour + 12)
                elif am_pm_match.group() == "am" and time_obj.hour == 12:
                    time_obj = time_obj.replace(hour=0)
                return time_obj.strftime("%H:%M")
            except ValueError:
                return None

        time_str = time_str.replace("-", ":").replace(" ", "")

        # Handle numeric strings (e.g., "9.5")
        if re.fullmatch(r"^\d+\.?\d*$", time_str):
            try:
                time_float = float(time_str)
                hours = int(time_float)
                minutes = int((time_float % 1) * 60)
                return f"{hours:02d}:{minutes:02d}"
            except:
                pass

        # Handle compact formats like "0900" or "930"
        if re.fullmatch(r"^\d{3,4}$", time_str):
            time_str = time_str.zfill(4)
            try:
                hours = int(time_str[:2])
                minutes = int(time_str[2:])
                if 0 <= hours < 24 and 0 <= minutes < 60:
                    return f"{hours:02d}:{minutes:02d}"
            except:
                pass

        # Standard HH:MM parsing
        try:
            return datetime.strptime(time_str, "%H:%M").strftime("%H:%M")
        except ValueError:
            return None

    def normalize_timetable(self, input_file, output_file):
        """Process timetable and create normalized output"""
        try:
            # Read and preserve original data
            original_df = pd.read_excel(input_file)
            df = original_df.copy()

            # Normalize column names
            df.columns = df.columns.str.strip().str.lower()
            required_columns = ["person_group", "start_time", "end_time"]

            # Validate columns
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

            # Enhanced data cleaning
            # 1. Normalize person_group values
            df["person_group"] = df["person_group"].astype(str).str.strip().str.lower()

            # 2. Convert empty strings to NaN in required columns
            df[required_columns] = df[required_columns].replace(
                r"^\s*$", pd.NA, regex=True
            )

            # 3. Remove rows with missing values in required columns except person_group
            df = df.dropna(subset=["start_time", "end_time"]).copy()

            print(
                "DataFrame after removing rows with missing values in required columns except person_group:"
            )
            print(df)

            try:
                df["start_time"] = df["start_time"].apply(self.normalize_time)
            except Exception as e:
                print(f"Error normalizing start_time: {e}")
                df = df.dropna(subset=["start_time"]).copy()

            try:
                df["end_time"] = df["end_time"].apply(self.normalize_time)
            except Exception as e:
                print(f"Error normalizing end_time: {e}")
                df = df.dropna(subset=["end_time"]).copy()

            print("DataFrame after normalizing times:")
            print(df)

            # Sort for proper shift ordering
            df = df.sort_values(by=["person_group", "start_time"])

            # Create group mappings with normalized values
            unique_groups = df["person_group"].unique()
            for group in unique_groups:
                if group not in self.shift_mapping:
                    self.shift_mapping[group] = self.group_id_counter
                    self.group_id_counter += 1

            # Assign IDs
            df["person_group_id"] = df["person_group"].map(self.shift_mapping)
            df["shift_id"] = df.groupby("person_group").cumcount() + 1

            print("DataFrame after assigning IDs:")
            print(df)

            # Split shifts
            shift1_df = df[df["shift_id"] == 1].sort_values("person_group_id")
            shift2_df = df[df["shift_id"] == 2].sort_values("person_group_id")

            # Ensure person_group and person_group_id are copied over to shift 2
            shift2_df = shift2_df[
                ["person_group", "person_group_id", "start_time", "end_time"]
            ]

            # Move rows with missing person_group to shift2_df
            missing_person_group_df = df[df["person_group"].isna()]
            shift2_df = pd.concat([shift2_df, missing_person_group_df])

            # Validate shifts
            if len(shift1_df) == 0:
                raise ValueError("No valid shift 1 data found")

            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_file), exist_ok=True)

            # Save results
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                shift1_df.to_excel(
                    writer, index=False, sheet_name="Shift 1", freeze_panes=(1, 0)
                )
                if not shift2_df.empty:
                    shift2_df.to_excel(
                        writer, index=False, sheet_name="Shift 2", freeze_panes=(1, 0)
                    )
                original_df.to_excel(
                    writer, index=False, sheet_name="Original Data", freeze_panes=(1, 0)
                )

            print(f"Successfully normalized timetable. Output saved to: {output_file}")
            print(f"Total groups processed: {len(self.shift_mapping)}")
            print(
                f"Shift 1 records: {len(shift1_df)}, Shift 2 records: {len(shift2_df)}"
            )
            return shift1_df, shift2_df

        except Exception as e:
            print(f"Error processing timetable: {str(e)}")
            raise


if __name__ == "__main__":
    normalizer = TimetableNormalizer()
    current_dir = Path(__file__).parent
    input_file = current_dir / "data" / "work_timetable.xlsx"
    output_file = current_dir / "data" / "normalized_work_timetable.xlsx"
    normalizer.normalize_timetable(str(input_file), str(output_file))
