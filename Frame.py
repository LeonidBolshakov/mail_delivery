from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from openpyxl.utils.exceptions import InvalidFileException
import openpyxl
from loguru import logger
from Mail import CreateMails, PutMessages
import constant as c


class WidgetOutput(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.click_button = ClickButton()
        self.output_init(self)

    @staticmethod
    def output_init(self):
        self.title("отправка e-mail")
        self.geometry("1100x500")
        self.columnconfigure(index=0, weight=2)
        self.columnconfigure(index=1, weight=5)
        self.columnconfigure(index=2, weight=2)
        self.columnconfigure(index=3, weight=1)

        label_title = Label(text="Групповая отправка e-mail",
                            foreground=c.COLOR_TITLE,
                            justify=CENTER,
                            font=", 25")
        label_title.grid(row=0, column=0, columnspan=4, pady=50)
        label_empty = Label(text=' ')
        label_empty.grid(row=1, column=0)

        self.create_button(text="Импорт адресов (MS EXCEL)",
                           command=lambda: self.click_button.click_button_list_email(self))
        self.create_button(text="Ввод текста e-mail",
                           command=lambda: self.click_button.click_button_text_email(self), column=1)
        self.create_button(text="Ввод списка вложений",
                           command=lambda: self.click_button.click_button_attachment(self), column=2)
        self.button_go = self.create_button(text="Отправить письма",
                                            command=lambda: self.click_button.click_button_go(self), column=3)

    @staticmethod
    def output_listbox(list_, row, column):
        lst = Variable(value=list_)
        listbox_list = Listbox(listvariable=lst, height=2)
        listbox_list.grid(row=row, column=column, sticky=N)

    @staticmethod
    def create_button(text, command, row=2, column=0):
        button = Button(text=text, command=command)
        button.bind("<Return>", command)
        button.grid(row=row, column=column)
        return button

    @staticmethod
    def create_label(row, column, text):
        lbl = Label(text=text)
        lbl.grid(row=row, column=column)

    def output_listbox_email(self, list_):
        self.output_listbox(list_, 4, 0)

    def output_listbox_attachment(self, list_):
        self.output_listbox(list_, 4, 2)

    def create_subject_field(self):
        st = Text(height=1)
        st.grid(row=4, column=1, sticky=EW)
        st.bind("<FocusOut>", lambda e: self.click_button.focus_out_subject(e))
        st.bind("<Tab>", self.focus_next)
        st.bind("<Shift-Tab>", self.focus_previous)

        return st

    def create_content_field(self):
        st = ScrolledText()
        st.grid(row=7, column=1, sticky=N, padx=2)
        st.bind("<FocusOut>", lambda e: self.click_button.focus_out_content(e))
        st.bind("<Tab>", self.focus_next)
        st.bind("<Shift-Tab>", self.focus_previous)

        return st

    @staticmethod
    def create_labels_column_0(self, total_addresses, wrong_addresses):
        self.create_label(3, 0, 'Список адресов')
        self.create_label(5, 0, f'Всего адресов: {total_addresses}')
        self.create_label(6, 0, f'Ошибочных адресов: {wrong_addresses}')

    def create_label_column_1(self):
        self.create_label(3, 1, 'Заголовок сообщения')
        self.create_label(5, 1, 'Текст сообщения')

    def create_label_column_2(self):
        self.create_label(3, 2, 'Список вложений')
        self.create_label(5, 2, f'Всего вложений: {len(self.click_button.attachments)}')

    def create_label_column_3(self, put_messages):
        self.create_label(3, 3, f'Отправлено писем - {put_messages.count_sends} \n'
                                f'Не отправлено писем - {put_messages.count_errors} ')


    @staticmethod
    def focus_next(event):
        event.widget.tk_focusNext().focus()
        return "break"

    @staticmethod
    def focus_previous(event):
        event.widget.tk_focusPrev().focus()
        return "break"


class ClickButton():
    def __init__(self):
        self.attachments = []
        self.list_email = []
        self.content = ''
        self.content_field = Text()
        self.subject = ''
        self.subject_field = Text()
        self.button_go = Button()
        self.total_addresses = 0
        self.wrong_addresses = 0

    def click_button_list_email(self, widget_output):
        sheet = self.read_sheet_excel()
        if sheet:
            self.list_email = self.create_list_email(sheet)
            self.address_verification(self, self.list_email)
            widget_output.create_labels_column_0(widget_output, self.total_addresses, self. wrong_addresses)
            widget_output.output_listbox_email(self.list_email)

    def click_button_text_email(self, widget_output):
        self.subject_field = widget_output.create_subject_field()
        self.content_field = widget_output.create_content_field()
        widget_output.create_label_column_1()

    def click_button_attachment(self, widget_output):
        self.attachments = self.create_attachments()
        widget_output.output_listbox_attachment(self.attachments)
        widget_output.create_label_column_2()

    def click_button_go(self, widget_output):
        widget_output.button_go["state"] = "disabled"
        widget_output.update()
        create_mails = CreateMails()
        msgs = create_mails.create(c.FROM, self.list_email, self.subject, self.content, self.attachments)
        put_messages = PutMessages()
        put_messages.put(msgs)
        widget_output.create_label_column_3(put_messages)

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
    def create_attachments():
        return filedialog.askopenfilenames(filetypes=[('Все файлы', '*.*')])

    @staticmethod
    def address_verification(self, list_email):
        self.total_addresses = 0
        self.wrong_addresses = 0
        for email in list_email:
            self.total_addresses += 1
            if '@' not in email:
                self.wrong_addresses += 1

    def focus_out_subject(self, event):
        self.subject = self.subject_field.get("1.0", "end")

    def focus_out_content(self, event):
        self.content = self.content_field.get("1.0", "end")


