from RPA.Excel.Application import Application, catch_com_error
import win32com
from win32com.client import constants
import gc
from enum import Enum
from dataclasses import dataclass
from utils import create_test_excel_file
import logging


def log_type_and_dir(object):
    logging.warning(f"{type(object)} {dir(object)}")


@dataclass
class PivotField:
    data_column: str
    pivot_column_label: str
    pivot_operation: str
    pivot_numberformat: str


def to_pivot_operation(operation_name: str):
    if operation_name == "SUM":
        return constants.xlSum
    elif operation_name == "AVERAGE":
        return constants.xlAverage
    elif operation_name == "MAX":
        return constants.xlMax
    else:
        return None


class ExtendedExcel(Application):
    def __init__(self, autoexit: bool = True) -> None:
        super().__init__(autoexit)

    def quit_application(self, save_changes: bool = False) -> None:
        """Quit the application."""
        if not self.app:
            return

        self.close_document(save_changes)
        with catch_com_error():
            self.app.Quit()

        self.app = None
        gc.collect()

    def create_pivot_table(
        self,
        source_worksheet: str,
        pivot_worksheet: str,
        pivot_name: str,
        pt_rows: list,
        pt_cols: list,
        pt_filters: list,
        pt_fields: list,
    ):
        """
        wb = workbook1 reference
        ws1 = worksheet1
        pt_ws = pivot table worksheet number
        ws_name = pivot table worksheet name
        pt_name = name given to pivot table
        pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
        """
        # pivot table location
        pt_loc = len(pt_filters) + 2

        self.set_active_worksheet(source_worksheet)
        # grab the pivot table source data
        pc = self.workbook.PivotCaches().Create(
            SourceType=constants.xlDatabase, SourceData=self.worksheet.UsedRange
        )

        # create the pivot table object
        pc.CreatePivotTable(
            TableDestination=f"{pivot_worksheet}!R{pt_loc}C1", TableName=pivot_name
        )

        self.set_active_worksheet(pivot_worksheet)
        # select the pivot table work sheet and location to create the pivot table
        self.worksheet.Select()
        self.worksheet.Cells(pt_loc, 1).Select()

        # Sets the rows, columns and filters of the pivot table

        for field_list, field_r in (
            (pt_filters, constants.xlPageField),
            (pt_rows, constants.xlRowField),
            # (pt_cols, constants.xlColumnField),
        ):
            for i, value in enumerate(field_list):
                self.worksheet.PivotTables(pivot_name).PivotFields(
                    value
                ).Orientation = field_r
                self.worksheet.PivotTables(pivot_name).PivotFields(value).Position = (
                    i + 1
                )

        for field in pt_fields:
            pivot_operation = to_pivot_operation(field.pivot_operation)
            self.worksheet.PivotTables(pivot_name).AddDataField(
                self.worksheet.PivotTables(pivot_name).PivotFields(field.data_column),
                field.pivot_column_label,
                pivot_operation,
            ).NumberFormat = field.pivot_numberformat

        # Visiblity True or Valse
        self.worksheet.PivotTables(pivot_name).ShowValuesRow = True
        self.worksheet.PivotTables(pivot_name).ColumnGrand = True
        return pc

    def get_pivot_tables(self):
        pivot_tables = {}

        # TODO. Get tables for all worksheet (in the workbook)
        tables = self.worksheet.PivotTables()
        log_type_and_dir(tables)
        self.logger.warning(tables.Count)
        # for index, t in enumerate(tables):
        #    self.logger.warning(f"{index}: {t}")


if __name__ == "__main__":
    ee = ExtendedExcel(autoexit=False)
    create_test_excel_file("pivoting.xlsx", "data")
    ee.open_application(visible=True, display_alerts=True)
    ee.open_workbook("pivoting.xlsx")
    ee.add_new_sheet("test")
    ee.add_new_sheet("test2")
    filters = []
    fields = [
        PivotField("price", "Price Sum", "SUM", "0"),
        PivotField("price", "Price Average", "AVERAGE", "0"),
        PivotField(
            "price",
            "Price Max",
            "MAX",
            "0",
        ),
    ]
    pt = ee.create_pivot_table(
        "data",
        "test",
        "pivoting",
        ["products", "expense"],
        ["products"],
        filters,
        fields,
    )
    print(dir(pt))
    pt = ee.create_pivot_table(
        "data", "test2", "pivoting", ["date"], ["price"], [], fields
    )
    ee.save_excel()
    ee.quit_application()
