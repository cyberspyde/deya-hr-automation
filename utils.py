import pandas as pd
import re
from datetime import datetime


def load_work_timetable(filepath: str):
    """Load work timetable from Excel and ensure correct time parsing."""
    try:
        df = pd.read_excel(
            filepath, dtype=str
        )  # Read as strings to process times correctly

        # Ensure correct column names
        required_columns = {"person_group", "start_time", "end_time"}
        missing_columns = required_columns - set(df.columns)

        if missing_columns:
            raise ValueError(
                f"Missing required columns in timetable: {missing_columns}"
            )

        # Helper function to clean and convert time formats
        def parse_time(value):
            if pd.isna(value) or not isinstance(value, str):
                return None  # Keep as None if missing or not a string

            value = value.strip()  # Remove leading/trailing spaces

            # Convert common formats like '9-00' → '9:00'
            value = re.sub(r"(\d+)[\s\-]+(\d+)", r"\1:\2", value)

            # Remove non-numeric characters except ':'
            value = re.sub(r"[^\d:]", "", value)

            try:
                return datetime.strptime(value, "%H:%M").time()
            except ValueError:
                return None  # Invalid format, return None

        df["start_time"] = df["start_time"].apply(parse_time)
        df["end_time"] = df["end_time"].apply(parse_time)

        # Ensure 'person_group' is unique; remove duplicates if necessary
        if df["person_group"].duplicated().any():
            print(
                "⚠️ Warning: Duplicate 'person_group' values found! Keeping the first occurrence."
            )
            df = df.drop_duplicates(subset=["person_group"], keep="first")

        return df.set_index("person_group").to_dict(orient="index")

    except Exception as e:
        print(f"❌ Error loading timetable: {e}")
        raise


test = {"s", "b", "t"}
