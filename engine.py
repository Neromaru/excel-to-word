# uncompyle6 version 3.5.0
# Python bytecode 3.7 (3394)
# Decompiled from: Python 3.7.2 (default, Dec 29 2018, 06:19:36)
# [GCC 7.3.0]
# Embedded file name: engine.py
import datetime as dt
import os
import pathlib as pl
import uuid

import pandas
from mailmerge import MailMerge


class TemplateGenerator(object):

    def __init__(
            self,
            path_to_data=None,
            path_to_folder=None,
            where_to_save=None,
            top_level_saving:bool=False,
            named_header=None,
    ):
        self.data_file = path_to_data
        self.template_folder = path_to_folder
        self.excel = None
        self.save_folder = where_to_save
        self.named_header = named_header
        self.top_level_saving = top_level_saving

    def _serialize_datetimes(self, dict_fields: dict) -> dict:
        """1970-10-18 00:00:00"""
        for key, value in dict_fields.items():
            try:
                dict_fields[key] = dt.datetime.strptime(
                    value, "%Y-%m-%d %H:%M:%S"
                ).strftime("%d.%m.%Y")
            except:
                continue

        return dict_fields

    def read_data(self):
        self.excel = pandas.read_excel(self.data_file).fillna("")

    def generate_templates(self):
        for idx, row in self.excel.iterrows():
            index_name = pl.Path(f"{str(idx)}_{str(row[self.named_header])}")
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
                    variable: str(row[variable])
                    for variable in template_docx.get_merge_fields()
                }
                serialized_fields = self._serialize_datetimes(fields)
                (template_docx.merge)(**serialized_fields)
                template_basename = f"{index_name.name}_{template.name}"
                template_docx.write(os.path.join(write_folder, template_basename))

    def list_templates(self):
        return pl.Path(self.template_folder).glob("*.docx")
