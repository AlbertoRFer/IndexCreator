# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 14:37:04 2017

@author: Alberto
"""

import os
import platform
import sys

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

import indice
import dialogs

__version__ = "1.0.0"


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        """ Creates and organizes the widgets.
        Connect all the relevant signals.
        """
        super(MainWindow, self).__init__(parent)

        # Creation of the widgets
        label_protocols = QLabel("<font size=6>Protocolos</font>")
        label_protocols.setAlignment(Qt.AlignCenter)

        self.list_protocols = QListWidget()

        self.label_last_docs = QLabel(
            "<font size=6>Últimos documentos agregados</font>")
        self.label_last_docs.setAlignment(Qt.AlignCenter)

        self.list_last_documents = QListWidget()

        # Setting widgets into Layouts
        layout = QGridLayout()

        layout.addWidget(label_protocols, 0, 0, 1, 2)
        layout.addWidget(self.list_protocols, 1, 0, 10, 2)

        layout.addWidget(self.label_last_docs, 0, 2, 1, 5)
        layout.addWidget(self.list_last_documents, 1, 2, 10, 5)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Creating Menus and actions
        # Menu Documents
        menu_documents = self.menuBar().addMenu("&Documentos")

        new_doc_action = self.create_action("Agregar Documento",
                                            self.new_document, "Ctrl+N")
        menu_documents.addAction(new_doc_action)

        mod_doc_action = self.create_action("Modificar Documento",
                                            self.mod_document, "Ctrl+M")
        menu_documents.addAction(mod_doc_action)

        view_doc_action = self.create_action("Ver Documento",
                                             self.view_document, "Ctrl+D")
        menu_documents.addAction(view_doc_action)

        menu_documents.addSeparator()
        search_action = self.create_action("Buscar",
                                             self.search, "Ctrl+S")
        menu_documents.addAction(search_action)

        # Menu Protocols
        menu_protocols = self.menuBar().addMenu("&Protocolos")

        new_prot_action = self.create_action("Nuevo Protocolo",
                                             self.new_protocol, "Ctrl+P")
        menu_protocols.addAction(new_prot_action)

        del_prot_action = self.create_action("Eliminar Protocolo",
                                             self.del_protocol, "Ctrl+E")
        menu_protocols.addAction(del_prot_action)

        menu_protocols.addSeparator()
        make_index_action = self.create_action("Crear Índice",
                                               self.make_index, "Ctrl+I")
        menu_protocols.addAction(make_index_action)

        # Restore settings
        settings = QSettings()
        last_docs = settings.value("LastDocs", [])
        current_protocol = settings.value("CurrentProtocol", "")

        # Creating indice
        self.indice = indice.Indice(last_docs, current_protocol)

        # Populating lists
        self.update_last_docs()
        self.update_protocols()

        # Signals connections
        self.list_protocols.currentRowChanged.connect(self.update_last_docs)

        # Windows settings
        self.resize(QSize(1000, 700))
        self.setWindowTitle("Índice")

    def create_action(self, text, slot=None, shortcut=None, icon=None,
                      tip=None, checkable=False):
        """Creates a new action.

        All the relevant data are passed as paremeters"""
        action = QAction(text, self)
        if icon is not None:
            action.setIcon(QIcon(":/%s.png" % icon))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            action.triggered.connect(slot)
        if checkable:
            action.setCheckable(True)
        return action

    def update_last_docs(self):
        """Updates the last docs list.

        If no item is selected, the last docs list shows the last
        documents added. Otherwise, the contents of the selected protocol
        are showed.
        """
        pos = self.list_protocols.currentRow()

        if pos == -1 or pos == 0:
            msg = "<font size=6>Últimos documentos agregados</font>"
            self.label_last_docs.setText(msg)
            docs = self.indice.recent_docs

        else:
            protocol = str(self.list_protocols.item(pos).text())
            msg = "<font size=6>" + protocol + "</font>"
            self.label_last_docs.setText(msg)
            docs = self.indice.view_protocol(protocol)

        if docs is not None:
            self.list_last_documents.clear()
            if docs:
                self.list_last_documents.addItems(docs)
            else:
                msg = "Todavía no se han agregado documentos a este protocolo"
                self.list_last_documents.addItem(msg)

    def update_protocols(self):
        """Updates the protocols list."""
        protocols = self.indice.view_protocols()
        self.list_protocols.clear()
        self.list_protocols.addItem("Documentos Recientes")
        self.list_protocols.addItems(protocols)

    def new_document(self):
        """Creates a new document.

        The AddDocDlg dialog is called for enter the new document's data.
        """
        if not self.indice.view_protocols():
            title = "Error al crear nuevo documento"
            msg = """Debe crear al menos un protocolo antes de insertar 
algún documento.
            """
            QMessageBox.warning(self, title, msg)
            return

        dialog = dialogs.AddDocDlg(self.indice, self)

        if dialog.exec_():
            number = str(dialog.ledit_number.text())
            p_name = str(dialog.list_users.item(0).text())
            o_name = ""
            size = dialog.list_users.count()
            for i in range(1, size):
                o_name += str(dialog.list_users.item(i).text())
                if i != size - 1:
                    o_name += ", "

            tramit = str(dialog.ledit_tramit.text())
            protocol = str(dialog.combo_protocol.currentText())

            if self.indice.add_doc(protocol, number, p_name.upper(),
                                   o_name.upper(), tramit.upper()):
                self.update_last_docs()

                title = "Nuevo Documento"
                msg = "El documento se ha agregado correctamente."
                QMessageBox.information(self, title, msg)
            else:
                title = "Nuevo Documento"
                msg = "Error al agregar el documento."
                QMessageBox.critical(self, title, msg)

            self.indice.current_protocol = protocol

    def mod_document(self):
        """Allows to modify an existing document.

        The SelDocDlg dialog is called for enter document's number.
        The selected document is removed from the database and a
        AddDocDlg dialog is displayed with it's data.
        """
        if not self.indice.view_protocols():
            title = "Error al modificar documento"
            msg = """No hay documentos que modificar"""
            QMessageBox.warning(self, title, msg)
            return

        sel_dialog = dialogs.SelDocDlg(self.indice, self)

        if sel_dialog.exec_():
            # Obtain number and protocol from mod_dialog
            number = int(sel_dialog.ledit_number.text())
            protocol = str(sel_dialog.combo_protocol.currentText())

            # Obtain all the document data from the database
            doc = self.indice.get_doc_by_number(protocol, number)

            # Delete the document
            self.indice.delete_document(protocol, number)

            # Create new AddDocDlg to modify data
            dialog = dialogs.AddDocDlg(self.indice, self)

            # Initialize fields with the old document data
            dialog.ledit_number.setText(str(number))
            dialog.ledit_tramit.setText(doc[0][3])

            for i in range(dialog.combo_protocol.count()):
                if str(dialog.combo_protocol.itemText(i)) == protocol:
                    dialog.combo_protocol.setCurrentIndex(i)
                    break

            dialog.list_users.addItem(doc[0][1])
            other_users = doc[0][2].split(", ")
            dialog.list_users.addItems(other_users)

            for i in range(dialog.list_users.count()):
                dialog.list_users.item(i).setFlags(Qt.ItemIsEditable
                                                   | Qt.ItemIsEnabled
                                                   | Qt.ItemIsSelectable)

            # From now on it works as in the new_doc function
            if dialog.exec_():
                number = str(dialog.ledit_number.text())
                p_name = str(dialog.list_users.item(0).text())
                o_name = ""
                size = dialog.list_users.count()
                for i in range(1, size):
                    o_name += str(dialog.list_users.item(i).text())
                    if i != size - 1:
                        o_name += ", "

                tramit = str(dialog.ledit_tramit.text())
                protocol = str(dialog.combo_protocol.currentText())

                if self.indice.add_doc(protocol, number, p_name.upper(),
                                       o_name.upper(), tramit.upper()):
                    self.update_last_docs()

                    title = "Modificar Documento"
                    msg = "El documento se ha modificado correctamente."
                    QMessageBox.information(self, title, msg)
                else:
                    title = "Modificar Documento"
                    msg = "Error al modificar el documento."
                    QMessageBox.critical(self, title, msg)

                self.indice.current_protocol = protocol

    def view_document(self):
        """Allows to view the details of an existing document.

        The SelDocDlg dialog is called for enter document's number.
        The details of the selected document are displayed in a
        ShowDataDlg dialog.
        """
        if not self.indice.view_protocols():
            title = "Error al buscar el documento"
            msg = """No hay documentos en los protocolos."""
            QMessageBox.warning(self, title, msg)
            return

        sel_dialog = dialogs.SelDocDlg(self.indice, self)

        if sel_dialog.exec_():
            # Obtain number and protocol from mod_dialog
            number = int(sel_dialog.ledit_number.text())
            protocol = str(sel_dialog.combo_protocol.currentText())

            # Obtain all the document data from the database
            doc = self.indice.get_doc_by_number(protocol, number)

            dialog = dialogs.ShowDataDlg(self)

            dialog.label_number.setText(dialog.label_number.text()
                                        + str(number))
            dialog.label_protocol.setText(protocol)
            dialog.label_tramit.setText(doc[0][3])

            dialog.tbrowser_users.append(doc[0][1])

            other_users = doc[0][2].split(", ")
            for user in other_users:
                dialog.tbrowser_users.append(user)

            dialog.exec_()

    def new_protocol(self):
        """Creates a new protocol.

        The NewProtocolDlg dialog is called for selecting the new protocol.
        """
        dialog = dialogs.NewProtocolDlg(self.indice, self)

        if dialog.exec():
            if self.indice.new_protocol(
                    str(dialog.combo_protocol.currentText())):
                self.update_protocols()

                title = "Nuevo Protocolo"
                msg = "El protocolo se ha creado correctamente."
                QMessageBox.information(self, title, msg)
            else:
                title = "Nuevo Protocolo"
                msg = "Error al crear el protocolo."
                QMessageBox.critical(self, title, msg)

    def del_protocol(self):
        """Deletes a protocol.

        The DelProtocolDlg dialog is called for selecting protocol that is
        going to be deleted.
        """
        if not self.indice.view_protocols():
            title = "Error al eliminar protocolo"
            msg = """No existen protocolos para eliminar."""
            QMessageBox.warning(self, title, msg)
            return
        dialog = dialogs.DelProtocolDlg(self.indice, self)

        if dialog.exec():
            if self.indice.del_protocol(
                    str(dialog.combo_protocol.currentText())):
                self.update_protocols()

                title = "Eliminar Protocolo"
                msg = "El protocolo se ha eliminado correctamente."
                QMessageBox.information(self, title, msg)
            else:
                title = "Eliminar Protocolo"
                msg = "Error al eliminar el protocolo."
                QMessageBox.critical(self, title, msg)
            self.update_protocols()
            self.update_last_docs()

    def make_index(self):
        """Creates a new index in docx format.

        The MakeIndexDlg dialog is called for selecting the protocol and the
        initial and final document numbers.
        """
        if not self.indice.view_protocols():
            title = "Error al crear índice"
            msg = """Debe crear al menos un protocolo antes de crear un índice.
                    """
            QMessageBox.warning(self, title, msg)
            return
        dialog = dialogs.MakeIndexDlg(self.indice, self)

        if dialog.exec():
            doc_name = dialog.doc_name
            protocol = str(dialog.combo_protocol.currentText())
            number_ini = int(dialog.ledit_number_ini.text())
            number_fin = int(dialog.ledit_number_fin.text())

            self.indice.make_index(doc_name, protocol, number_ini, number_fin)

    def search(self):
        dialog = dialogs.SearchDlg(self.indice)

        dialog.exec_()

    def closeEvent(self, event):
        """Save the settings."""
        settings = QSettings()
        settings.setValue("LastDocs", self.indice.recent_docs)
        settings.setValue("CurrentProtocol", self.indice.current_protocol)


def main():
    app = QApplication(sys.argv)

    app.setApplicationName("Índice")
    app.setOrganizationName("House")
    app.setOrganizationDomain("house.cu")

    form = MainWindow()
    form.show()
    app.exec_()

# TODO Ver el tema de las excepciones y los loggings.
# TODO Agregar funcionalidad de busqueda.
# TODO Hacer GUI más bonita


main()
