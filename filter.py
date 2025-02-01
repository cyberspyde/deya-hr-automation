import pandas as pd

class Filter:
    def __init__(self, df):
        self.df = df

    def filter_by_device(self, device_name):
        """Filter data by device name"""
        return self.df[self.df['device_name'] == device_name]

    def filter_by_person_group(self, person_group):
        """Filter data by person group"""
        return self.df[self.df['person_group'] == person_group]

    def filter_by_date_range(self, start_date, end_date):
        """Filter data by date range"""
        return self.df[(self.df['date'] >= start_date) & (self.df['date'] <= end_date)]

    def filter_by_person_group_from_excel(self, excel_file_path):
        """Filter data by person group from an Excel file"""
        try:
            person_groups_df = pd.read_excel(excel_file_path, usecols=['person_group'])
            person_groups = person_groups_df['person_group'].dropna().unique()
            return self.df[self.df['person_group'].isin(person_groups)]
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            raise

    def apply_filters(self, filters):
        """Apply multiple filters to the data"""
        filtered_df = self.df
        if 'device_name' in filters:
            filtered_df = self.filter_by_device(filters['device_name'])
        if 'person_group' in filters:
            filtered_df = self.filter_by_person_group(filters['person_group'])
        if 'person_group_excel' in filters:
            filtered_df = self.filter_by_person_group_from_excel(filters['person_group_excel'])
        return filtered_df