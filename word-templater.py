import os
import sys
import time

import docx

import wt

import pythoncom
from threading import Thread
from win32com.client import Dispatch
from PyQt5 import QtCore, QtGui, QtWidgets


def thread(function: callable) -> callable:
    def wrapper(*args, **kwargs) -> Thread:
        t = Thread(target=function, args=args, kwargs=kwargs, daemon=True)
        t.start(); return t
    return wrapper


class Signals(QtCore.QObject):
    status = QtCore.pyqtSignal(str, int)
    inProgress = QtCore.pyqtSignal(bool)

    def flush(self):
        pass


class WordTemplateApp(QtWidgets.QMainWindow):
    files = []
    directory = os.getcwd()
    signals = Signals()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ui = wt.Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.btnAdd.clicked.connect(lambda: self.chooseFiles())
        self.ui.btnRemove.clicked.connect(lambda: self.removeWordFiles())
        self.ui.btnStart.clicked.connect(lambda: self.start())
        self.ui.btnOutput.clicked.connect(lambda: self.changeOutputDir())

        self.ui.output.setText(self.directory)
        self.ui.sex.addItems(["Муж", "Жен"])

        self.signals.status.connect(
            lambda s, v: (self.ui.lStatus.setText(s), self.ui.sbStatus.setProperty("value", v))
        )

        self.signals.inProgress.connect(
            lambda b: (self.ui.listWidget.setEnabled(not b),
                       self.ui.conteinerStatus.setEnabled(b),
                       self.ui.templateConteiner.setEnabled(not b)
                       )
        )

    @thread
    def _start(self):
        pythoncom.CoInitializeEx(0)

        self.signals.status.emit("Открытие Word", 0)
        self.signals.inProgress.emit(True)

        word = Dispatch("Word.Application")

        for fs in self.files:
            document = word.Documents.Open(os.path.abspath(fs))

            self.signals.status.emit(f"Открытие документа: {document.Name}", self.files.index(fs))

            for field in document.Fields:
                field.Unlink()

            new_path = os.path.join(self.directory, document.Name)
            document.SaveAs2(new_path)
            document.Close()

            document = docx.Document(new_path)

            keywords = self.fillInKeywords()
            for paragraph in document.paragraphs:
                for keyword in keywords:
                    paragraph.text = paragraph.text.replace(keyword, keywords[keyword])

            document.save(new_path)

        word.Quit()

        self.signals.status.emit("Закрытие Word", 100)
        self.signals.inProgress.emit(False)

        time.sleep(1)
        self.signals.status.emit("Статус", 0)

    def start(self):
        keywords = self.fillInKeywords()

        for keyword in keywords:
            if keywords[keyword] is None:
                QtWidgets.QMessageBox.critical(self, "Неполные данные!", f"Проверьте введенные данные!")
                return

        self._start()

    def changeOutputDir(self) -> None:
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения фалов")
        if directory != "":
            self.directory = directory
            self.ui.output.setText(self.directory)

    def chooseFiles(self) -> None:
        new_files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Выберите один или несколько фалов word",
            filter="Word Files|*.doc;*.docx;*.docm")

        for fs in new_files:
            if fs not in self.files:
                _, name = os.path.split(fs)
                self.ui.listWidget.addItem(name)
                self.files.append(fs)

    def removeWordFiles(self) -> None:
        for idx in self.ui.listWidget.selectedIndexes():
            self.ui.listWidget.takeItem(idx.row())
            self.files.pop(idx.row())

    def fillInKeywords(self) -> dict:

        organization = self.ui.organization.text().strip()
        position = self.ui.position.text().strip()
        getter_iof = self.ui.iof.text().strip()
        getter_sex = self.ui.sex.currentText().strip()
        short_organ = self.ui.shortOrganization.text().strip()
        sender_iof = self.ui.diof.text().strip()
        sender_surname = self.ui.surname.text().strip()
        doc_year = self.ui.year.text().strip()

        def create_iof(iof: str = ""):
            try:
                i, o, f = list(map(str.capitalize, iof.split()))
                return f"{i[0]}.{o[0]}.{f}"
            except ValueError as e:
                QtWidgets.QMessageBox.critical(self, "Неверные данные!", "Поле И.О.Ф заполненно не полностью!")
                return None

        def create_io(iof: str = ""):
            try:
                i, o, f = list(map(str.capitalize, iof.split()))
                return f"{i} {o}"
            except ValueError as e:
                QtWidgets.QMessageBox.critical(self, "Неверные данные!", "Поле И.О.Ф заполненно не полностью!")
                return None

        def create_surname(iof: str = "", surname: str = ""):
            try:
                i, o, f = list(map(str.capitalize, iof.split()))
                return f"{surname.capitalize()} {i[0]}.{o[0]}."
            except ValueError as e:
                QtWidgets.QMessageBox.critical(self, "Неверные данные!", "Поле И.О.Ф заполненно не полностью!")
                return None

        def sex(change: str):
            if change == "Жен":
                return "Уважаемая"
            else:
                return "Уважаемый"

        return {
            "[Организация]": organization,
            "[Должность получателя]": position,
            "[И.О.Фамилия]": create_iof(getter_iof),
            "[Имя Отчество]": create_io(getter_iof),
            "[сокращенное наименование проверяемой организации]": short_organ,
            "[И.О. Фамилия]": create_iof(sender_iof),
            "(Фамилия И.О. руководителя задания по аудиту)": create_surname(sender_iof, sender_surname),
            "20ХХ": doc_year,
            "Уважаемый": sex(getter_sex)
        }


if __name__ == '__main__':
    application = QtWidgets.QApplication(sys.argv)
    word_template = WordTemplateApp()
    word_template.show()
    sys.exit(application.exec_())
