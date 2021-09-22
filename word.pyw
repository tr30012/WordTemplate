import os
import sys
import time

import docx

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


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(654, 545)
        font = QtGui.QFont()
        font.setFamily("Golos UI")
        font.setPointSize(12)
        MainWindow.setFont(font)
        MainWindow.setWindowTitle("Шаблонизатор")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.widget_2 = QtWidgets.QWidget(self.centralwidget)
        self.widget_2.setObjectName("widget_2")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget_2)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.filesConteiner = QtWidgets.QWidget(self.widget_2)
        self.filesConteiner.setObjectName("filesConteiner")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.filesConteiner)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label = QtWidgets.QLabel(self.filesConteiner)
        self.label.setText("Список файлов")
        self.label.setObjectName("label")
        self.verticalLayout_3.addWidget(self.label)
        self.listWidget = QtWidgets.QListWidget(self.filesConteiner)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout_3.addWidget(self.listWidget)
        self.widget = QtWidgets.QWidget(self.filesConteiner)
        self.widget.setMinimumSize(QtCore.QSize(0, 30))
        self.widget.setObjectName("widget")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.btnAdd = QtWidgets.QPushButton(self.widget)
        self.btnAdd.setText("Добавить")
        self.btnAdd.setObjectName("btnAdd")
        self.horizontalLayout_2.addWidget(self.btnAdd)
        self.btnRemove = QtWidgets.QPushButton(self.widget)
        self.btnRemove.setText("Удалить")
        self.btnRemove.setObjectName("btnRemove")
        self.horizontalLayout_2.addWidget(self.btnRemove)
        self.btnStart = QtWidgets.QPushButton(self.widget)
        self.btnStart.setText("Начать")
        self.btnStart.setObjectName("btnStart")
        self.horizontalLayout_2.addWidget(self.btnStart)
        self.verticalLayout_3.addWidget(self.widget)
        self.horizontalLayout.addWidget(self.filesConteiner)
        self.templateConteiner = QtWidgets.QWidget(self.widget_2)
        self.templateConteiner.setMinimumSize(QtCore.QSize(300, 0))
        self.templateConteiner.setObjectName("templateConteiner")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.templateConteiner)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.groupBox = QtWidgets.QGroupBox(self.templateConteiner)
        self.groupBox.setMaximumSize(QtCore.QSize(16777215, 200))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox.setFont(font)
        self.groupBox.setTitle("Информация о получателе")
        self.groupBox.setObjectName("groupBox")
        self.formLayout = QtWidgets.QFormLayout(self.groupBox)
        self.formLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.formLayout.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
        self.formLayout.setRowWrapPolicy(QtWidgets.QFormLayout.DontWrapRows)
        self.formLayout.setLabelAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.formLayout.setFormAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.formLayout.setObjectName("formLayout")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_2.setText("Организация")
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_2)
        self.organization = QtWidgets.QLineEdit(self.groupBox)
        self.organization.setText("")
        self.organization.setObjectName("organization")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.organization)
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setText("Должность")
        self.label_3.setObjectName("label_3")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_3)
        self.position = QtWidgets.QLineEdit(self.groupBox)
        self.position.setText("")
        self.position.setObjectName("position")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.position)
        self.label_4 = QtWidgets.QLabel(self.groupBox)
        self.label_4.setText("И.О.Ф")
        self.label_4.setObjectName("label_4")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_4)
        self.iof = QtWidgets.QLineEdit(self.groupBox)
        self.iof.setText("")
        self.iof.setObjectName("iof")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.iof)
        self.label_6 = QtWidgets.QLabel(self.groupBox)
        self.label_6.setText("Пол")
        self.label_6.setObjectName("label_6")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.label_6)
        self.sex = QtWidgets.QComboBox(self.groupBox)
        self.sex.setCurrentText("")
        self.sex.setObjectName("sex")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.sex)
        self.label_7 = QtWidgets.QLabel(self.groupBox)
        self.label_7.setText("СН Организации")
        self.label_7.setObjectName("label_7")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.label_7)
        self.shortOrganization = QtWidgets.QLineEdit(self.groupBox)
        self.shortOrganization.setText("")
        self.shortOrganization.setObjectName("shortOrganization")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.shortOrganization)
        self.verticalLayout_4.addWidget(self.groupBox)
        self.groupBox_2 = QtWidgets.QGroupBox(self.templateConteiner)
        self.groupBox_2.setMaximumSize(QtCore.QSize(16777215, 140))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setTitle("Руководитель задания по аудиту")
        self.groupBox_2.setObjectName("groupBox_2")
        self.formLayout_2 = QtWidgets.QFormLayout(self.groupBox_2)
        self.formLayout_2.setObjectName("formLayout_2")
        self.label_10 = QtWidgets.QLabel(self.groupBox_2)
        self.label_10.setText("И.О.Ф")
        self.label_10.setObjectName("label_10")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_10)
        self.diof = QtWidgets.QLineEdit(self.groupBox_2)
        self.diof.setText("")
        self.diof.setObjectName("diof")
        self.formLayout_2.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.diof)
        self.label_11 = QtWidgets.QLabel(self.groupBox_2)
        self.label_11.setText("Фамилия (в Р.п)")
        self.label_11.setObjectName("label_11")
        self.formLayout_2.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.label_11)
        self.surname = QtWidgets.QLineEdit(self.groupBox_2)
        self.surname.setText("")
        self.surname.setObjectName("surname")
        self.formLayout_2.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.surname)
        self.label_12 = QtWidgets.QLabel(self.groupBox_2)
        self.label_12.setText("Год")
        self.label_12.setObjectName("label_12")
        self.formLayout_2.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_12)
        self.year = QtWidgets.QLineEdit(self.groupBox_2)
        self.year.setText("")
        self.year.setObjectName("year")
        self.formLayout_2.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.year)
        self.verticalLayout_4.addWidget(self.groupBox_2)
        self.groupBox_3 = QtWidgets.QGroupBox(self.templateConteiner)
        self.groupBox_3.setMaximumSize(QtCore.QSize(16777215, 100))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setTitle("Папка с выходными данными")
        self.groupBox_3.setObjectName("groupBox_3")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.groupBox_3)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.output = QtWidgets.QLineEdit(self.groupBox_3)
        self.output.setObjectName("output")
        self.horizontalLayout_3.addWidget(self.output)
        self.btnOutput = QtWidgets.QPushButton(self.groupBox_3)
        self.btnOutput.setMaximumSize(QtCore.QSize(50, 16777215))
        self.btnOutput.setObjectName("btnOutput")
        self.horizontalLayout_3.addWidget(self.btnOutput)
        self.verticalLayout_4.addWidget(self.groupBox_3)
        self.horizontalLayout.addWidget(self.templateConteiner)
        self.verticalLayout.addWidget(self.widget_2)
        self.conteinerStatus = QtWidgets.QWidget(self.centralwidget)
        self.conteinerStatus.setMaximumSize(QtCore.QSize(16777215, 60))
        self.conteinerStatus.setObjectName("conteinerStatus")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.conteinerStatus)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.lStatus = QtWidgets.QLabel(self.conteinerStatus)
        self.lStatus.setText("Статус")
        self.lStatus.setObjectName("lStatus")
        self.verticalLayout_2.addWidget(self.lStatus)
        self.sbStatus = QtWidgets.QProgressBar(self.conteinerStatus)
        self.sbStatus.setMaximum(100)
        self.sbStatus.setProperty("value", 0)
        self.sbStatus.setObjectName("sbStatus")
        self.verticalLayout_2.addWidget(self.sbStatus)
        self.verticalLayout.addWidget(self.conteinerStatus)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.label_7.setToolTip(_translate("MainWindow", "<html><head/><body><p>Сокращенное название организации</p></body></html>"))
        self.label_11.setToolTip(_translate("MainWindow", "<html><head/><body><p>Фамилия руководителя в Родительном падеже. (Иванов - Иванова)</p></body></html>"))
        self.btnOutput.setText(_translate("MainWindow", "..."))


class WordTemplateApp(QtWidgets.QMainWindow):
    files = []
    directory = os.getcwd()
    signals = Signals()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ui = Ui_MainWindow()
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


def py_main(argc: int, argv: list):
    application = QtWidgets.QApplication(sys.argv)
    word_template = WordTemplateApp()
    word_template.show()
    sys.exit(application.exec_())


if __name__ == '__main__':
    py_main(sys.argv.__len__(), sys.argv)
