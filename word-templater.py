import os
import sys

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

            for paragraph in document.paragraphs:
                print(paragraph.text)

        word.Quit()

        self.signals.status.emit("Закрытие Word", 100)
        self.signals.inProgress.emit(False)

    def start(self):
        self._start()

    def changeOutputDir(self) -> None:
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения фалов")
        if directory != "":
            self.directory = directory
            self.ui.output.setText(self.directory)

    def chooseFiles(self) -> None:
        new_files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Выберите один или несколько фалов эксель",
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

    def getIOFamily(self) -> str:
        return ""

    def getIOFull(self) -> str:
        return ""

    def fillInKeywords(self) -> dict:
        return {
            "[Организация]": self.ui.organization.text().strip(),
            "[Должность получателя]": self.ui.position.text().strip(),
            "[И.О.Фамилия]": self.getIOFamily().strip(),
            "[Имя Отчество]": self.getIOFull().strip(),
            "[сокращенное наименование проверяемой организации]": self.ui.shortOrganization.text().strip(),
            "[И.О. Фамилия]": None,
            "(Фамилия И.О. руководителя задания по аудиту)": None,
            "20ХХ": self.ui.year.text().strip()
        }


if __name__ == '__main__':
    application = QtWidgets.QApplication(sys.argv)
    word_template = WordTemplateApp()
    word_template.show()
    sys.exit(application.exec_())
