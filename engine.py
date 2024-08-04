# uncompyle6 version 3.5.0
# Python bytecode 3.7 (3394)
# Decompiled from: Python 3.7.2 (default, Dec 29 2018, 06:19:36)
# [GCC 7.3.0]
# Embedded file name: engine.py
import datetime as dt
import os
import pathlib as pl
import uuid

from mailmerge import MailMerge
import openpyxl
from openpyxl.styles.numbers import (
    FORMAT_PERCENTAGE_00,
    FORMAT_DATE_YYYYMMDD2,
    FORMAT_DATE_DDMMYY,
)


class TemplateGenerator(object):

    def __init__(
        self,
        path_to_data=None,
        path_to_folder=None,
        where_to_save=None,
        top_level_saving: bool = False,
        named_header=None,
    ):
        self.data_file = path_to_data
        self.template_folder = path_to_folder
        self.excel = None
        self.save_folder = where_to_save
        self.named_header = named_header
        self.top_level_saving = top_level_saving
        self._headers = None

    @property
    def headers(self):
        if self._headers is None:
            self._headers = [i.value for i in list(self.excel.iter_rows())[0]]
        return self._headers

    @staticmethod
    def format_cell_value(cell):
        """Convert cell value to a string formatted according to the cell's number format."""
        format_code = cell.number_format
        value = cell.value

        if not value:
            return value

        if format_code == FORMAT_PERCENTAGE_00:
            # Handle percentage
            return f"{value * 100:.2f}%"
        elif format_code in [FORMAT_DATE_YYYYMMDD2, FORMAT_DATE_DDMMYY] or cell.is_date:
            # Handle date
            return value.strftime("%d/%m/%Y")
        elif isinstance(value, float):
            # Handle general number format as 000 000,00
            return f"{value:,.2f}".replace(",", " ").replace(".", ",")
        else:
            # Handle all other cases
            return str(value)

    def read_data(self):
        extension = pl.Path(self.data_file).suffix
        if extension == ".xlsx":
            self.excel = openpyxl.load_workbook(self.data_file).active
        elif extension == ".xls":
            raise ValueError(
                "Ця версія програми не підритмує старі формати Екселю будь ласка збережіть у форматі .xlsx"
            )
        elif not self.excel:
            raise ValueError("Додано файл не підтримуваного формату")

    def generate_templates(self):
        init_row = 2
        for row in self.excel.iter_rows(min_row=init_row):
            try:
                self._make_template(row, init_row)
            except Exception as e:
                print(str(e))
                print(str(init_row))
                raise e
            init_row += 1

    def _make_template(self, row, row_number):
        row = dict(zip(self.headers, row))
        index_name = pl.Path(f"{str(row_number)}_{str(row[self.named_header].value)}")
        if not self.top_level_saving:
            write_folder = pl.Path(self.save_folder) / index_name
            if write_folder.exists():
                write_folder = self.save_folder / pl.Path(
                    index_name.name + str(uuid.uuid4())[:8]
                )
            write_folder.mkdir(parents=True, exist_ok=True)
        else:
            write_folder = pl.Path(self.save_folder)
        for template in self.list_templates():
            template_docx = MailMerge(template)
            fields = {
                variable: self.format_cell_value(row[variable])
                for variable in template_docx.get_merge_fields()
            }
            (template_docx.merge)(**fields)
            template_basename = f"{index_name.name}_{template.name}"
            template_docx.write(os.path.join(write_folder, template_basename))

    def list_templates(self):
        return pl.Path(self.template_folder).glob("*.docx")
