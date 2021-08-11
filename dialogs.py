# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 14:37:04 2017

@author: Alberto
"""

import sys
import os
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

import indice


class AddDocDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(AddDocDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index

        # Creation of the widgets
        label_number = QLabel("Número:")
        self.ledit_number = QLineEdit()

        # Initial value passed from the main window
        c_prot = self.indice.current_protocol

        if c_prot != "":
            self.ledit_number.setText(str(self.indice.next_number(c_prot)))

        self.ledit_number.setMaximumWidth(40)

        # Regular expression only allow numbers in range 1-999
        number_reg_exp = QRegExp(r"^[1-9]\d{0,2}$")
        self.ledit_number.setValidator(QRegExpValidator(number_reg_exp, self))
        self.label_number_status = QLabel()

        label_protocol = QLabel("Protocolo:")
        self.combo_protocol = QComboBox()

        # The combobox is populated with the existing protocols
        protocols = self.indice.view_protocols()
        for p in protocols:
            self.combo_protocol.addItem(p)

            if p == self.indice.current_protocol:
                self.combo_protocol.setCurrentIndex(
                    self.combo_protocol.count() - 1)

        # Shows document's number at beginning
        self.number_validation()

        label_tramit = QLabel("Escritura:")
        self.ledit_tramit = QLineEdit()

        group = QGroupBox("Comparecientes:")
        label_name = QLabel("Nombre(s):")
        label_last_name = QLabel("Apellidos:")
        self.ledit_name = QLineEdit()
        self.ledit_last_name = QLineEdit()
        self.list_users = QListWidget()

        button_add_user = QPushButton("Agregar")
        button_mod_user = QPushButton("Modificar")
        button_del_user = QPushButton("Eliminar")

        self.button_accept = QPushButton("Aceptar")
        self.button_accept.setEnabled(False)
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        group_layout = QGridLayout()
        group_layout.addWidget(label_name, 0, 0)
        group_layout.addWidget(self.ledit_name, 0, 1)
        group_layout.addWidget(label_last_name, 1, 0)
        group_layout.addWidget(self.ledit_last_name, 1, 1)
        group_layout.addWidget(self.list_users, 2, 0, 7, 2)

        group.setLayout(group_layout)

        layout = QGridLayout()

        layout.addWidget(label_number, 0, 0)
        layout.addWidget(self.ledit_number, 0, 1)
        layout.addWidget(self.label_number_status, 0, 2)

        layout.addWidget(label_protocol, 0, 3)
        layout.addWidget(self.combo_protocol, 0, 4, 1, 2)

        layout.addWidget(label_tramit, 1, 0)
        layout.addWidget(self.ledit_tramit, 1, 1, 1, 5)

        layout.addWidget(group, 2, 0, 9, 4)

        layout.addWidget(button_add_user, 2, 4, 1, 2)
        layout.addWidget(button_mod_user, 3, 4, 1, 2)
        layout.addWidget(button_del_user, 4, 4, 1, 2)

        layout.addWidget(self.button_accept, 9, 4, 1, 2)
        layout.addWidget(button_cancel, 10, 4, 1, 2)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Escritura")

        # Signals connections
        self.ledit_number.textChanged.connect(self.activate_button_accept)
        self.combo_protocol.currentIndexChanged.connect(
            self.activate_button_accept)
        self.ledit_tramit.textChanged.connect(self.activate_button_accept)

        button_add_user.clicked.connect(self.add_user)
        button_mod_user.clicked.connect(self.modify_user)
        button_del_user.clicked.connect(self.delete_user)

        self.button_accept.clicked.connect(self.check_dialog)
        button_cancel.clicked.connect(self.reject)

    def add_user(self):
        """ Add a user to the list.

        If name and last name line edits are not empty, a new user
        is added to the user's list.

        If the user is the first to be added, the format is last name, name.
        For all the other users, the format is: name last name

        All users added are made editable.

        After this operation the name and last name line edits are cleaned,
        and the focus return to the name line edit.
        """
        if self.ledit_name.text() != "" and self.ledit_last_name.text() != "":
            if self.list_users.count() == 0:
                self.list_users.addItem(self.ledit_last_name.text() + ", "
                                        + self.ledit_name.text())
            else:
                self.list_users.addItem(self.ledit_name.text() + " "
                                        + self.ledit_last_name.text())

            self.list_users.item(self.list_users.count() - 1).setFlags(
                Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            self.ledit_name.setText("")
            self.ledit_last_name.setText("")
            self.ledit_name.setFocus()
            self.activate_button_accept()

    def modify_user(self):
        """ Modify a user from the list.

        The user that is going to be modified is the one that is selected
        at the time. If no user is selected the function does nothing.
        """
        actual = self.list_users.currentItem()
        if actual is not None:
            self.list_users.editItem(actual)

    def delete_user(self):
        """ Delete a user from the list.

        The user that is going to be deleted is the one that is selected
        at the time. If no user is selected the function does nothing.
        """
        actual = self.list_users.currentItem()
        if actual is not None:
            self.list_users.takeItem(self.list_users.row(actual))
            self.activate_button_accept()

    def check_dialog(self):
        """ Shows a dialog of confirmation

        The new dialog contains all the information of this dialog.

        If the user accepts, this dialogs accept too and return to the
        caller. If the user does not accept, the focus return to this dialog.
        """
        dialog = ShowDataDlg(self)

        dialog.label_number.setText(dialog.label_number.text()
                                    + self.ledit_number.text())
        dialog.label_protocol.setText(self.combo_protocol.currentText())
        dialog.label_tramit.setText(self.ledit_tramit.text())

        for i in range(0, self.list_users.count()):
            dialog.tbrowser_users.append(str(self.list_users.item(i).text()))

        if dialog.exec_():
            self.accept()

    def number_validation(self):
        """Validates the document number.

        If the number is valid, a label next to it will show OK.
        If the number is no valid, a label next to it will show EXISTE
        If the number is empty, no message appear.
        """
        n = str(self.ledit_number.text())
        p = str(self.combo_protocol.currentText())
        if n == "":
            self.label_number_status.setText("")
            return False

        if p == "":
            self.label_number_status.setText(
                "<font color = red>No Protocolo<\font>")
            return False

        if self.indice.is_valid_number(p, n):
            self.label_number_status.setText("<font color = green>OK<\font>")
            return True
        else:
            self.label_number_status.setText("<font color = red>EXISTE<\font>")
            return False

    def activate_button_accept(self):
        """Activate o deactivate the accept button.

        If all the fields are valid it activates the accept button.
        If any of the fields has a non valid value, the accept button is
        disabled.
        """
        if self.number_validation():
            number_ok = True
        else:
            number_ok = False

        if str(self.ledit_tramit.text()) != "":
            tramit_ok = True
        else:
            tramit_ok = False

        if self.list_users.count() != 0:
            users_ok = True
        else:
            users_ok = False

        if number_ok and tramit_ok and users_ok:
            self.button_accept.setEnabled(True)
        else:
            self.button_accept.setEnabled(False)


class ShowDataDlg(QDialog):
    def __init__(self, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(ShowDataDlg, self).__init__(parent)

        # Creation of the widgets
        self.label_number = QLabel("Escritura No.- ")
        self.label_protocol = QLabel()
        self.label_tramit = QLabel("")
        label_users = QLabel("Comparecientes: ")
        self.tbrowser_users = QTextBrowser()

        button_accept = QPushButton("Aceptar")
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        layout1 = QHBoxLayout()
        layout1.addWidget(self.label_number)
        layout1.addWidget(self.label_protocol)

        layout2 = QHBoxLayout()
        layout2.addWidget(button_accept)
        layout2.addWidget(button_cancel)

        layout = QVBoxLayout()
        layout.addLayout(layout1)
        layout.addWidget(self.label_tramit)
        layout.addWidget(label_users)
        layout.addWidget(self.tbrowser_users)
        layout.addLayout(layout2)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Datos a insertar")

        # Signals connections
        button_accept.clicked.connect(self.accept)
        button_cancel.clicked.connect(self.reject)


class NewProtocolDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(NewProtocolDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index

        # Creation of the widgets
        self.combo_protocol = QComboBox()

        # The combobox is populated with the available protocols
        protocols = self.indice.view_protocols()
        for year in range(1990, 2100):
            p = "Protocolo" + str(year)
            if p not in protocols:
                self.combo_protocol.addItem(p)

        button_accept = QPushButton("Aceptar")
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        layout = QGridLayout()
        layout.addWidget(self.combo_protocol, 0, 0, 1, 2)

        layout.addWidget(button_accept, 1, 0)
        layout.addWidget(button_cancel, 1, 1)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Nuevo Protocolo")
        self.resize(QSize(200, 80))

        # Signals connections
        button_accept.clicked.connect(self.accept)
        button_cancel.clicked.connect(self.reject)


class DelProtocolDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(DelProtocolDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index

        # Creation of the widgets
        self.combo_protocol = QComboBox()

        # The combobox is populated with the existing protocols
        self.combo_protocol.addItems(self.indice.view_protocols())

        button_accept = QPushButton("Aceptar")
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        layout = QGridLayout()
        layout.addWidget(self.combo_protocol, 0, 0, 1, 2)

        layout.addWidget(button_accept, 1, 0)
        layout.addWidget(button_cancel, 1, 1)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Eliminar Protocolo")
        self.resize(QSize(220, 80))

        # Signals connections
        button_accept.clicked.connect(self.confirm_delete)
        button_cancel.clicked.connect(self.reject)

    def confirm_delete(self):
        title = "Eliminar Protocolo"
        msg = "¿Está seguro de eliminar "\
            + str(self.combo_protocol.currentText()) + "?\n"\
            + "Todos los datos de este protocolo se borarrán para SIEMPRE"

        reply = QMessageBox.question(self, title, msg, QMessageBox.Yes
                                     | QMessageBox.No | QMessageBox.Cancel)

        if reply == QMessageBox.Yes:
            self.accept()
        elif reply == QMessageBox.No:
            self.reject()


class MakeIndexDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(MakeIndexDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index
        self.doc_name = ""

        # Creation of the widgets
        self.combo_protocol = QComboBox()

        label_number_ini = QLabel("Número Inicial:")
        label_number_fin = QLabel("Número Final:")

        self.ledit_number_ini = QLineEdit()
        self.ledit_number_fin = QLineEdit()

        # Regular expression only allow numbers in range 1-999
        number_reg_exp = QRegExp(r"^[1-9]\d{0,2}$")
        self.ledit_number_ini.setValidator(
            QRegExpValidator(number_reg_exp, self))
        self.ledit_number_fin.setValidator(
            QRegExpValidator(number_reg_exp, self))

        self.label_number_ini_status = QLabel()
        self.label_number_fin_status = QLabel()

        # The combobox is populated with the available protocols
        protocols = self.indice.view_protocols()
        for p in protocols:
            self.combo_protocol.addItem(p)

            if p == self.indice.current_protocol:
                self.combo_protocol.setCurrentIndex(
                    self.combo_protocol.count() - 1)

        label_missing_docs = QLabel("Faltan:")
        label_missing_docs.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.combo_missing_docs = QComboBox()

        self.button_accept = QPushButton("Aceptar")
        self.button_accept.setEnabled(False)
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        layout = QGridLayout()
        layout.addWidget(self.combo_protocol, 0, 0)
        layout.addWidget(label_missing_docs, 0, 1)
        layout.addWidget(self.combo_missing_docs, 0, 2)

        layout.addWidget(label_number_ini, 1, 0)
        layout.addWidget(self.ledit_number_ini, 1, 1)
        layout.addWidget(self.label_number_ini_status, 1, 2)

        layout.addWidget(label_number_fin, 2, 0)
        layout.addWidget(self.ledit_number_fin, 2, 1)
        layout.addWidget(self.label_number_fin_status, 2, 2)

        layout.addWidget(self.button_accept, 3, 1)
        layout.addWidget(button_cancel, 3, 2)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Crear Índice")
        self.resize(QSize(200, 80))

        # Signals connections
        self.ledit_number_ini.textChanged.connect(self.activate_button_accept)
        self.ledit_number_fin.textChanged.connect(self.activate_button_accept)

        self.combo_protocol.currentIndexChanged.connect(
            self.activate_button_accept)

        self.button_accept.clicked.connect(self.select_file_name)
        button_cancel.clicked.connect(self.reject)

    def activate_button_accept(self):
        """Activate o deactivate the accept button.

        If all the fields are valid it activates the accept button.
        If any of the fields has a non valid value, the accept button is
        disabled.
        """
        if self.number_ini_validation():
            number_ini_ok = True
        else:
            number_ini_ok = False

        if self.number_fin_validation():
            number_fin_ok = True
        else:
            number_fin_ok = False

        if number_ini_ok and number_fin_ok:
            self.button_accept.setEnabled(True)
        else:
            self.button_accept.setEnabled(False)

    def number_ini_validation(self):
        """Validates the initial document number.

        If the number is valid, a label next to it will show OK.
        If the number is no valid, a label next to it will show EXISTE
        If the number is empty, no message appear.
        """
        n = str(self.ledit_number_ini.text())
        p = str(self.combo_protocol.currentText())
        if n == "":
            self.label_number_ini_status.setText("")
            return False

        if p == "":
            self.label_number_ini_status.setText(
                "<font color = red>No Protocolo<\font>")
            return False

        if self.indice.is_valid_number(p, n):
            msg = "<font color = red>NO EXISTE<\font>"
            self.label_number_ini_status.setText(msg)
            return False
        else:
            msg = "<font color = green><b>OK</b><\font>"
            self.label_number_ini_status.setText(msg)
            return True

    def number_fin_validation(self):
        """Validates the final document number.

        If the number is valid, a label next to it will show OK.
        If the number is no valid, a label next to it will show NO EXISTE
        If the number is empty, no message appear.
        If the interval is not valid (empty spaces), a label next to it will
        show MAL INTERVALO
        """
        n = str(self.ledit_number_fin.text())
        p = str(self.combo_protocol.currentText())

        self.combo_missing_docs.clear()
        self.combo_missing_docs.setCurrentText("NO")
        if n == "":
            self.label_number_fin_status.setText("")
            return False

        if p == "":
            self.label_number_fin_status.setText(
                "<font color = red>No Protocolo<\font>")
            return False

        if self.indice.is_valid_number(p, n):
            msg = "<font color = red>NO EXISTE<\font>"
            self.label_number_fin_status.setText(msg)
            return False
        else:
            if not self.number_ini_validation():
                msg = "<font color = red><b>MAL INTERVALO</b><\font>"
                self.label_number_fin_status.setText(msg)
                return False

            ini = int(self.ledit_number_ini.text()) + 1
            fin = int(self.ledit_number_fin.text())

            if ini > fin:
                msg = "<font color = red><b>MAL INTERVALO</b><\font>"
                self.label_number_fin_status.setText(msg)
                return False

            valid = True
            for i in range(ini, fin):
                if self.indice.is_valid_number(p, i):
                    valid = False
                    self.combo_missing_docs.addItem(str(i))
            if valid:
                msg = "<font color = green><b>OK</b><\font>"
                self.label_number_fin_status.setText(msg)
                return True
            else:
                self.combo_missing_docs.setCurrentText("NO")

            msg = "<font color = red><b>MAL INTERVALO</b><\font>"
            self.label_number_fin_status.setText(msg)
            return False

    def select_file_name(self):
        """Select the name of the file.

        Shows a QFileDialog to choose the name of the file in which
        we are going to save the data.
        """
        dir = "."
        formats = "Archivo de Word (*.docx)"

        fname = QFileDialog.getSaveFileName(self,
                                            "Guardar Índice",
                                            dir,
                                            formats)
        if fname[0]:
            self.doc_name = fname[0]
            if not self.doc_name.endswith(".docx"):
                self.doc_name += ".docx"
            self.accept()


class SelDocDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(SelDocDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index

        # Creation of the widgets
        label_number = QLabel("Número:")
        self.ledit_number = QLineEdit()

        # Initial value passed from the main window
        c_prot = self.indice.current_protocol

        self.ledit_number.setMaximumWidth(40)
        self.ledit_number.setMinimumWidth(40)

        # Regular expression only allow numbers in range 1-999
        number_reg_exp = QRegExp(r"^[1-9]\d{0,2}$")
        self.ledit_number.setValidator(QRegExpValidator(number_reg_exp, self))
        self.label_number_status = QLabel()

        label_protocol = QLabel("Protocolo:")
        self.combo_protocol = QComboBox()

        # The combobox is populated with the existing protocols
        protocols = self.indice.view_protocols()
        for p in protocols:
            self.combo_protocol.addItem(p)

            if p == self.indice.current_protocol:
                self.combo_protocol.setCurrentIndex(
                    self.combo_protocol.count() - 1)

        self.button_accept = QPushButton("Aceptar")
        self.button_accept.setEnabled(False)
        button_cancel = QPushButton("Cancelar")

        # Setting widgets into Layouts
        layout = QGridLayout()
        layout.addWidget(label_number, 0, 0)
        layout.addWidget(self.ledit_number, 0, 1)
        layout.addWidget(self.label_number_status, 1, 0, 1, 2)

        layout.addWidget(label_protocol, 0, 2)
        layout.addWidget(self.combo_protocol, 0, 3, 1, 2)

        layout.addWidget(self.button_accept, 1, 3)
        layout.addWidget(button_cancel, 1, 4)

        self.setLayout(layout)

        # Window's settings
        self.setWindowTitle("Seleccionar Documento")
        self.resize(QSize(240, 80))

        # Signals connections
        self.ledit_number.textChanged.connect(self.activate_button_accept)
        self.combo_protocol.currentIndexChanged.connect(
            self.activate_button_accept)
        self.button_accept.clicked.connect(self.accept)
        button_cancel.clicked.connect(self.reject)

    def number_validation(self):
        """Validates the document number.

        If the number is valid, a label next to it will show OK.
        If the number is no valid, a label next to it will show EXISTE
        If the number is empty, no message appear.
        """
        n = str(self.ledit_number.text())
        p = str(self.combo_protocol.currentText())
        if n == "":
            self.label_number_status.setText("")
            return False

        if p == "":
            self.label_number_status.setText(
                "<font color = red>No Protocolo<\font>")
            return False

        if not self.indice.is_valid_number(p, n):
            self.label_number_status.setText("<font color = green>EXISTE<\font>")
            return True
        else:
            self.label_number_status.setText("<font color = red>NO EXISTE<\font>")
            return False

    def activate_button_accept(self):
        """Activate o deactivate the accept button.

        If all the fields are valid it activates the accept button.
        If any of the fields has a non valid value, the accept button is
        disabled.
        """
        if self.number_validation():
            self.button_accept.setEnabled(True)
        else:
            self.button_accept.setEnabled(False)


class SearchDlg(QDialog):
    def __init__(self, index=None, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(SearchDlg, self).__init__(parent)

        if index is None:
            self.indice = indice.Indice()
        else:
            self.indice = index

        label = QLabel("Texto a buscar:")
        self.ledit_search = QLineEdit()
        self.tbrowser_results = QTextBrowser()

        buttonSearch = QPushButton("Buscar")

        layout = QGridLayout()
        layout.addWidget(label, 0, 0)

        layout.addWidget(self.ledit_search, 1, 0, 1, 4)
        layout.addWidget(buttonSearch, 1, 4)

        layout.addWidget(self.tbrowser_results, 2, 0, 10, 5)

        self.setLayout(layout)
        self.setWindowTitle("Buscar")
        self.setFixedSize(400, 500)

        buttonSearch.clicked.connect(self.search)

    def search(self):
        self.tbrowser_results.clear()

        search_string = str(self.ledit_search.text())
        if search_string == "":
            return

        protocols = self.indice.view_protocols()

        is_empty = True

        for protocol in protocols:
            result = self.indice.search_user(protocol, search_string)

            msg = "<font color = green><b><u>En "
            msg += protocol
            msg += ": </u></b><\font><font color = black>"

            if result:
                is_empty = False
                for line in result:
                    self.tbrowser_results.append(msg + line)

        if is_empty:
            self.tbrowser_results.append("No se encontraron resultados.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = SearchDlg()
    form.exec_()
