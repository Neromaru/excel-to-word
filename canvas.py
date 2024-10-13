import tkinter
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import messagebox

from engine import TemplateGenerator


class App(object):
    window = None
    engine = None
    combox = None

    save_button = None
    save_directory = Path.home()

    template_button = None
    templates_directory = Path.home()
    template_listbox = None

    data_file_button = None
    data_file = None

    named_header = None

    def create_template_widgets(self):
        # Button to choose template directory
        self.template_button = ttk.Button(
            self.window,
            text="Виберіть папку з шаблонами",
            command=self.choose_template_directory,
        )
        self.template_button.pack(pady=20)

        # Frame for template selection
        template_frame = ttk.LabelFrame(self.window, text="Вибір шаблонів")
        template_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Listbox for template selection
        self.template_listbox = tk.Listbox(
            template_frame, selectmode=tk.MULTIPLE, width=25
        )
        self.template_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar for the listbox
        scrollbar = ttk.Scrollbar(
            template_frame, orient=tk.VERTICAL, command=self.template_listbox.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.template_listbox.config(yscrollcommand=scrollbar.set)

        # Select All checkbox
        self.select_all_var = tk.BooleanVar()
        self.select_all_checkbox = ttk.Checkbutton(
            self.window,
            text="Вибрати усі шаблони",
            variable=self.select_all_var,
            command=self.toggle_all,
        )
        self.select_all_checkbox.pack(pady=(0, 10))

    def get_save_directory(self):
        self.save_directory = filedialog.askdirectory(initialdir=self.save_directory)
        if self.save_directory:
            self.save_button.configure(text=str(self.save_directory))
        self.engine.save_folder = self.save_directory

    def choose_template_directory(self):
        self.templates_directory = filedialog.askdirectory(
            initialdir=self.templates_directory
        )
        if self.templates_directory:
            self.populate_template_listbox()
        self.template_button.configure(text=str(self.templates_directory))

    def populate_template_listbox(self):
        self.template_listbox.delete(0, tk.END)  # Clear existing items
        templates = self.list_templates()
        for template in templates:
            self.template_listbox.insert(tk.END, template)

    def list_templates(self):
        template_files = Path(self.templates_directory).glob("*.docx")
        return [template.name for template in template_files]

    def toggle_all(self):
        if self.select_all_var.get():
            self.template_listbox.select_set(0, tk.END)
        else:
            self.template_listbox.selection_clear(0, tk.END)

    def populate_headers_selector(self):
        self.engine.read_data()
        try:
            self.combox["values"] = [
                i.value for i in list(self.engine.excel.iter_rows())[0]
            ]
            self.combox.current(0)
        except AttributeError:
            pass

    def get_data_file(self):
        self.data_file = filedialog.askopenfilename(
            filetypes=(("Excel", "*.xlsx"), ("Excel", "*.xls"))
        )
        self.engine.data_file = self.data_file
        self.populate_headers_selector()
        self.data_file_button.config(text=str(self.data_file))

    def set_buttons(self):
        BUTTON_PUDDING = 10
        self.save_button = ttk.Button(
            self.window,
            text="Виберіть папку куди зберігти результати",
            command=self.get_save_directory,
        )
        self.data_file_button = ttk.Button(
            self.window,
            text="Виберіть файл данних",
            command=self.get_data_file,
        )
        submit_button_form = ttk.Button(
            self.window,
            text="Згенерувати файл",
            command=self.submit_from,
        )
        self.combox = ttk.Combobox()
        save_label = ttk.Label(text="Виберіть шлях зберігання")
        save_label.pack(pady=(BUTTON_PUDDING, 0))
        self.save_button.pack(pady=(0, BUTTON_PUDDING))
        data_label = ttk.Label(text="Виберіть файл даних")
        data_label.pack(pady=(BUTTON_PUDDING, 0))
        self.data_file_button.pack(pady=(0, BUTTON_PUDDING))
        combox_label = ttk.Label(text="Виберіть параметр зберігання назви")
        combox_label.pack(pady=(BUTTON_PUDDING, 0))
        self.combox.pack(pady=(0, BUTTON_PUDDING))
        submit_button_form.pack(pady=BUTTON_PUDDING)

    def submit_from(self):
        self.engine.named_header = self.combox.get()
        self.engine.templates = [
            self.templates_directory / Path(i)
            for i in self.template_listbox.selection_get().split("\n")
        ]
        try:
            self.engine.generate_templates()
            messagebox.showinfo("Успіх", "Усе зроблено вірно")
        except Exception as e:
            try:
                messagebox.showerror(
                    "ERROR",
                    f""
                    f"Перевірте правильність введених в програму даних якщо ж помилка не відносться до даниз див. далі"
                    f"Помилка: \n {e}"
                    f"У разі помилки програми, але не  даних зробіть скриншот цієї помилки та надішліть розробнику",
                )
            finally:
                e = None
                del e
                return
        messagebox.showinfo("Success", "Усе зроблено вірно")

    def pack_widgets(self):
        self.create_template_widgets()
        self.set_buttons()

    def run(self):
        self.window, self.engine = tk.Tk(), TemplateGenerator()
        self.window.title("E2W")
        self.pack_widgets()
        self.window.mainloop()
