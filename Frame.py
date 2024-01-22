from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from openpyxl.utils.exceptions import InvalidFileException
import openpyxl
from loguru import logger
from Mail import Mail
import constant as c


class Frame:
    def __init__(self, root):
        self.attachments = []
        self.list_email = []
        self.root = root
        self.content = ''
        self.content_field = ''
        self.subject = ''
        self.subject_field = Text()
        self.button_go = Button()
        self.total_addresses = 0
        self.wrong_addresses = 0

        self.mail = Mail()
        self.frame_init(self)

    @staticmethod
    def create_button(text, command, row=2, column=0):
        button = Button(text=text, command=command)
        button.grid(row=row, column=column)
        button.bind("<Return>", command)
        return button

    def click_button_list_email(self, e=None):
        sheet = self.read_sheet_excel()
        if sheet:
            self.list_email = self.create_list_email(sheet)
            self.address_verification(self, self.list_email)
            self.create_labels_column_0(self)
            self.output_list(self.list_email, 4, 0)

    def click_button_text_email(self, e=None):
        self.subject_field = self.create_text_field(self, 4, 1)
        self.content_field = self.create_scroll_text_field(self, 7, 1)
        self.create_label_column_1(self)

    def click_button_attachment(self, e=None):
        self.attachments = self.create_attachments()
        self.output_list(self.attachments, 4, 2)
        self.create_label_column_2(self)

    def click_button_go(self, e=None):
        self.button_go["state"] = "disabled"
        self.root.update()
        msgs = self.mail.create_messages(c.FROM, self.list_email, self.subject, self.content, self.attachments)
        self.mail.put_messages(msgs)
        self.create_label_column_3(self)

    @staticmethod
    def frame_init(self):
        self.root.title("отправка e-mail")
        self.root.geometry("1100x500")
        self.root.columnconfigure(index=0, weight=2)
        self.root.columnconfigure(index=1, weight=5)
        self.root.columnconfigure(index=2, weight=2)
        self.root.columnconfigure(index=3, weight=1)

        label_title = Label(text="Групповая отправка e-mail",
                            foreground=c.COLOR_TITLE,
                            justify=CENTER,
                            font=", 25")
        label_title.grid(row=0, column=0, columnspan=4, pady=50)
        label_empty = Label(text=' ')
        label_empty.grid(row=1, column=0)

        self.create_button(text="Импорт адресов (MS EXCEL)", command=self.click_button_list_email)
        self.create_button(text="Ввод текста e-mail", command=self.click_button_text_email, column=1)
        self.create_button(text="Ввод списка вложений", command=self.click_button_attachment, column=2)
        self.button_go = self.create_button(text="Отправить письма", command=self.click_button_go, column=3)

    @staticmethod
    def read_sheet_excel():
        filename = filedialog.askopenfilename(filetypes=[('Файлы Excel', 'xlsx')])
        try:
            workbook = openpyxl.load_workbook(filename)
            return workbook.active
        except InvalidFileException:
            logger.warning('Отказ от ввода списка адресов')
        return None

    @staticmethod
    def create_list_email(sheet):
        result = []
        for row in sheet.iter_rows(values_only=True):
            if row[0] != c.HEAD_EXCEL:
                result.append(row[0])
        return result

    @staticmethod
    def output_list(list_, row, column):
        lst = Variable(value=list_)
        listbox_list = Listbox(listvariable=lst, height=2)
        listbox_list.grid(row=row, column=column, sticky=N)

    @staticmethod
    def create_text_field(self, row, column):
        st = Text(height=1)
        st.grid(row=row, column=column, sticky=EW)
        st.bind("<FocusOut>", lambda e: self.focus_out_text(e, self))
        st.bind("<Tab>", self.focus_next)
        st.bind("<Shift-Tab>", self.focus_previous)

        return st

    @staticmethod
    def create_scroll_text_field(self, row, column):
        st = ScrolledText()
        st.grid(row=row, column=column, sticky=N, padx=2)
        st.bind("<FocusOut>", lambda e: self.focus_out_scroll_text(e, self))
        st.bind("<Tab>", self.focus_next)
        st.bind("<Shift-Tab>", self.focus_previous)

        return st

    @staticmethod
    def focus_out_text(event, self):
        self.subject = self.subject_field.get("1.0", "end")

    @staticmethod
    def focus_out_scroll_text(event, self):
        self.content = self.content_field.get("1.0", "end")

    @staticmethod
    def create_attachments():
        return filedialog.askopenfilenames(filetypes=[('Все файлы', '*.*')])

    @staticmethod
    def create_label(row, column, text):
        lbl = Label(text=text)
        lbl.grid(row=row, column=column)

    @staticmethod
    def create_labels_column_0(self):
        self.create_label(3, 0, 'Список адресов')
        self.create_label(5, 0, f'Всего адресов: {self.total_addresses}')
        self.create_label(6, 0, f'Ошибочных адресов: {self.wrong_addresses}')

    @staticmethod
    def create_label_column_1(self):
        self.create_label(3, 1, 'Заголовок сообщения')
        self.create_label(5, 1, 'Текст сообщения')

    @staticmethod
    def create_label_column_2(self):
        self.create_label(3, 2, 'Список вложений')
        self.create_label(5, 2, f'Всего вложений: {len(self.attachments)}')

    @staticmethod
    def create_label_column_3(self):
        self.create_label(3, 3, f'Отправлено писем - {self.mail.count_mails} \n'
                                f'не отправлено писем - {self.mail.count_errors} ')

    @staticmethod
    def address_verification(self, list_email):
        self.total_addresses = 0
        self.wrong_addresses = 0
        for email in list_email:
            self.total_addresses += 1
            if '@' not in email:
                self.wrong_addresses += 1

    @staticmethod
    def focus_next(event):
        event.widget.tk_focusNext().focus()
        return "break"

    @staticmethod
    def focus_previous(event):
        event.widget.tk_focusPrev().focus()
        return "break"
