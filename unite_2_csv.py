import csv

from datetime import datetime
from typing import Any, Dict, List

from openpyxl import Workbook


class Unite2CSV:
    def __init__(self, csv_1: str, csv_2: str, csv_1_primary_field: str, csv_2_primary_field: str, time_series_by_column: List[str]) -> None:
        self.date_formats = ['%Y-%m-%d', '%m/%d/%y']
        self.csv_1 = self.read_data(csv_1, csv_1_primary_field)
        self.csv_2 = self.read_data(csv_2, csv_2_primary_field)
        self.primary_field_1 = csv_1_primary_field
        self.primary_field_2 = csv_2_primary_field
        self.time_series_by_column = time_series_by_column

    def read_data(self, csv_file: str, field: str) -> List[dict]:
        return_data = []
        with open(csv_file) as f:
            data = csv.DictReader(f, delimiter=',', quotechar='"')

            for d in data:
                if isinstance(d[field], datetime):
                    continue

                for date_format in self.date_formats:
                    try:
                        d[field] = datetime.strptime(d[field], date_format)
                        return_data.append(d)
                        break
                    except ValueError:
                        pass

                if not isinstance(d[field], datetime):
                    raise ValueError('Field was unable to convert to date!')

            return return_data

    def return_matching_fields(self, data: List[dict], primary_field: str, by_value: Any) -> List[str]:
        return_data = []
        for d in data:
            if d[primary_field] == by_value:
                return_data.append(d)
        return return_data

    def group_together(self) -> List[dict]:
        return_data = []

        for csv_1_data in self.csv_1:

            primary_value = csv_1_data[self.primary_field_1]
            csv_1_data.pop(self.primary_field_1, None)

            matching_data = self.return_matching_fields(
                self.csv_2, self.primary_field_2, primary_value)

            if not matching_data:
                continue

            for original_csv_2_data in matching_data:
                csv_2_data = original_csv_2_data.copy()
                csv_2_data.pop(self.primary_field_2, None)

                group_dict = {
                    'primary_value': primary_value.strftime('%Y-%m-%d')}
                updated_csv_1_data = {f'csv_1_{k}' if k in csv_2_data.keys(
                ) else k: v for k, v in csv_1_data.items()}
                updated_csv_2_data = {f'csv_2_{k}' if k in csv_1_data.keys(
                ) else k: v for k, v in csv_2_data.items()}

                group_dict.update(updated_csv_1_data)
                group_dict.update(updated_csv_2_data)
                return_data.append(group_dict)

        return return_data

    def to_csv(self, data: List[dict], fname: str) -> None:
        if not data:
            raise ValueError('No data matched primary field!')

        fieldnames = data[0].keys()
        with open(fname, 'w') as f:
            csv_file = csv.DictWriter(f, fieldnames=fieldnames)
            csv_file.writeheader()
            csv_file.writerows(data)

    def time_series(self, data: List[dict]) -> Dict[str, Dict[str, dict]]:
        return_dict = {}
        for d in data.copy():
            primary_column = d['primary_value']
            return_dict.setdefault(primary_column, {})
            attribute = ', '.join(
                [v for k, v in d.items() if k in self.time_series_by_column])
            return_dict[primary_column].setdefault(attribute, {})

            for column in self.time_series_by_column + ['primary_value']:
                del d[column]

            return_dict[primary_column][attribute] = d
        return return_dict

    def time_series_to_xlsx(self, data: Dict[str, Dict[str, dict]], fname: str) -> None:
        wb = Workbook()
        ws = wb.active

        for date_index, date in enumerate(data.keys()):
            ws.cell(column=1, row=date_index+3, value=date)

        gather_attr = []
        gather_subtitles = []
        nr_of_subtitles_attr = 0
        for items in data.values():
            for key, values in items.items():
                if key in gather_attr:
                    continue
                gather_attr.append(key)
                gather_subtitles.extend(values.keys())
                if nr_of_subtitles_attr < len(values.keys()):
                    nr_of_subtitles_attr = len(values.keys()) + 1

        start_column = 2
        end_column = nr_of_subtitles_attr
        for attr in gather_attr:
            ws.merge_cells(start_row=1, start_column=start_column, end_row=1, end_column=end_column)
            ws.cell(column=start_column, row=1, value=attr)
            start_column = end_column + 1
            end_column = start_column + (nr_of_subtitles_attr - 2)

        for subtitle_index, subtitle in enumerate(gather_subtitles):
            ws.cell(column=subtitle_index+2, row=2, value=subtitle)

        use_attribute = None
        for row_index, date in enumerate(ws['A:A'][2:]):
            for coll_index, attribute in enumerate(ws[1]):
                if attribute.value:
                    use_attribute = attribute.value
                
                if not use_attribute:
                    continue

                subtitle_value = ws[2][coll_index].value
                if subtitle_value != None and data[date.value].get(use_attribute):
                    ws.cell(column=coll_index+1, row=row_index + 3, value=data[date.value][use_attribute][subtitle_value])

        wb.save(filename=fname)


if __name__ == "__main__":
    import time
    start = time.time()
    unite2csv = Unite2CSV(
        csv_1='WASDE (input X).csv',
        csv_2='WheatPriceData (input Y).csv',
        csv_1_primary_field='Report date',
        csv_2_primary_field='Date',
        time_series_by_column=['Category', 'Units', 'Attribute']
    )

    united_data = unite2csv.group_together()
    # unite2csv.to_csv(united_data, 'united_data.csv')
    time_series_data = unite2csv.time_series(united_data)
    unite2csv.time_series_to_xlsx(time_series_data, 'united_data.xlsx')

    print("End:", time.time() - start)
