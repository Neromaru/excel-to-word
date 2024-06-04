import tkinter
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter.ttk import Combobox

from engine import TemplateGenerator


class App(object):
    window = None
    engine = None
    combox = None
    save_directory = Path.home()
    templates_directory = Path.home()
    top_level_saving = False
    data_file = None
    named_header = None
    save_label = None
    template_label = None

    def _get_save_directory(self):
        self.save_directory = filedialog.askdirectory()
        self.engine.save_folder = self.save_directory
        self.save_label.config(
            text=f"Усі сгенеровані документі будуть за шляхом : {self.save_directory}"
        )

    def _get_templates_directory(self):
        self.templates_directory = filedialog.askdirectory()
        self.engine.template_folder = self.templates_directory
        self.template_label.config(
            text=f"Усі шаблони знаходяться за шляхом : {self.templates_directory}"
        )

    def _get_data_file(self):
        self.data_file = filedialog.askopenfilename(
            filetypes=(("Excel", "*.xlsx"), ("Excel", "*.xls"))
        )
        self.engine.data_file = self.data_file
        self.engine.read_data()
        try:
            self.combox["values"] = list(self.engine.excel.columns)
            self.combox.current(0)
        except AttributeError:
            pass


    def _set_top_level_saving(self):
        self.engine.top_level_saving = self.top_level_saving.get()

    def _set_buttons(self):
        select_save_folder = Button(
            (self.window),
            text="Зберегти в ...",
            command=(self._get_save_directory),
            width=20,
            pady=5,
            bd=1,
        )
        select_load_templates = Button(
            (self.window),
            text="Папка з шаблонами",
            command=(self._get_templates_directory),
            width=20,
            pady=5,
            bd=1,
        )
        select_data_file = Button(
            (self.window),
            text="Виберіть файл даних",
            command=(self._get_data_file),
            width=20,
            pady=5,
            bd=1,
        )
        submit_button_form = Button(
            (self.window),
            text="Згенерувати файл",
            command=(self.submit_from),
            width=20,
            pady=5,
            bd=1,
        )
        self.top_level_saving = tkinter.BooleanVar(self.window, value=False)
        top_level_saving = tkinter.Checkbutton(
            (self.window),
            text="Зберегти в 1 папку усі результати ?",
            variable=self.top_level_saving,
            onvalue=True,
            offvalue=False,
            command=(self._set_top_level_saving)
        )
        self.save_label = Label(
            (self.window),
            text=f"Усі сгенеровані документі будуть за шляхом : {self.templates_directory}",
        )
        self.save_label.grid(row=0, column=1, pady=3)
        self.template_label = Label(
            (self.window),
            text=f"Усі шаблони знаходяться за шляхом : {self.templates_directory}",
        )
        self.template_label.grid(row=1, column=1, pady=3)
        self.combox = Combobox()
        select_save_folder.grid(row=0, column=0, pady=3, padx=3)
        top_level_saving.grid(row=0, column=2, pady=4, padx=4)
        select_load_templates.grid(row=1, column=0, pady=3, padx=3)
        select_data_file.grid(row=2, column=0, pady=3, padx=3)
        self.combox.grid(row=2, column=1, pady=3)
        submit_button_form.grid(row=3, column=1, pady=15)

    def submit_from(self):
        self.engine.named_header = self.combox.get()
        try:
            self.engine.generate_templates()
            messagebox.showinfo("Success", "Усе зроблено вірно")
        except Exception as e:
            try:
                try:
                    messagebox.showerror(
                        "ERROR",
                        f"Винкла помилка зробіть скриншот цієї помилки та надішліть розробнику: \n {e}",
                    )
                finally:
                    e = None
                    del e

            finally:
                e = None
                del e

    def _list_templates(self):
        txt = scrolledtext.ScrolledText((self.window), width=40, height=10)
        txt.grid(column=0, row=2)

    def run(self):
        self.window, self.engine = Tk(), TemplateGenerator()
        self.window.title("E2W")
        self.window.geometry("750x200")
        self._set_buttons()
        self.window.mainloop()
