# -*- coding: utf-8 -*-
"""
Created on Thu Nov  2 15:52:54 2017

@author: Alberto
"""

import db
import string
import word


class Indice:
    def __init__(self, recent_docs=[], current_protocol="",
                 max_recent_docs=20):
        """Loads the database."""
        self.recent_docs = recent_docs
        self.current_protocol = current_protocol
        self.max_recent_docs = max_recent_docs
        self.database = db.Archivo('Protocolo.db')

    def new_protocol(self, protocol):
        """Creates a new protocol.

        The name of the new protocol is passed as an argument.
        Returns true is the protocol is created correctly, false otherwise.
        """
        if self.database.create_table(protocol):
            if self.current_protocol == "":
                self.current_protocol = protocol
            return True
        else:
            return False

    def del_protocol(self, protocol):
        """Deletes a protocol.

        The name of the protocol is passed as an argument.
        Returns true is the protocol is deleted correctly, false otherwise.
        """
        if self.database.delete_table(protocol):
            # Deletes this protocol's documents from recent_docs
            if self.recent_docs is not None:
                reverse_list = self.recent_docs[-1::-1]
                for line in reverse_list:
                    if str(line).count(protocol) > 0:
                        self.recent_docs.remove(line)

            # Stipulates new current_protocol if the old one is removed
            if protocol == self.current_protocol:
                p = self.database.view_tables()

                if p is None or p == []:
                    self.current_protocol = ""
                else:
                    self.current_protocol, = p[0]

            return True
        else:
            return False

    def view_protocol(self, protocol):
        """Returns the contents of a protocol.

        The name of the protocol is passed as an argument.
        Returns a list where every element corresponds to a document of the
        protocol. If an error occurs returns None.
        """
        raw = self.database.view_table(protocol)
        protocol_docs = []
        if raw is not None:
            for line in raw:
                number, m_name, o_name, tramit = line

                s = str(number) + ".- "
                s += tramit + ". "
                s += m_name
                if o_name != "":
                    s += ", " + o_name
                s += "."

                protocol_docs.append(s)

            protocol_docs.sort()
            return protocol_docs

    def add_doc(self, protocol, number, m_name, o_name, tramit):
        """Inserts a new document into a protocol.

        The name of the protocol and all the necessary data are passed
        as parameters.
        Returns true if the document is inserted correctly, false otherwise.
        """
        if self.database.insert_data(protocol, number, m_name, o_name, tramit):
            self.add_to_recent_docs(protocol, number, m_name, o_name, tramit)
            return True
        else:
            return False

    def is_valid_number(self, protocol, number):
        """Checks if a document's number is valid.

        The name of the protocol and the number are passed as parameter.
        Returns true if there is no document in the protocol with the passed
        number, false otherwise.
        """
        doc = self.database.doc_by_number(protocol, number)
        if doc is not None:
            if not doc:
                return True
            else:
                return False
        else:
            return "No pudo ejecutarse la orden"

    def next_number(self, protocol):
        """Returns the next available number in a protocol.

        The name of the protocol is passed as a parameter.
        Finds the maximum number assigned to a document and returns it's
        successor."""
        number = self.database.max_number(protocol)

        if number is not None:
            n, = number
            if n is None:
                return 1
            return n + 1
        else:
            return "No pudo ejecutarse la orden"

    def add_to_recent_docs(self, protocol, number, m_name, o_name, tramit):
        """Add a document to the list of recent docs.

        The protocol and the document data are passed as parameters.
        If the amount of recent documents exceeds the maximum number permited,
        the older one (first one) is removed."""
        s = "En " + protocol + ": "
        s += str(number) + ".- "
        s += tramit + ". "
        s += m_name
        if o_name != "":
            s += ", " + o_name
        s += "."

        if self.recent_docs is None:
            self.recent_docs = [s]
        else:
            self.recent_docs.append(s)

        if len(self.recent_docs) > self.max_recent_docs:
            self.recent_docs.pop(0)

    def view_protocols(self):
        """Returns a list of the protocols in the database.

        If an error occur returns an error message.
        """
        p = self.database.view_tables()

        if p is not None:
            protocols = []
            for item in p:
                prot, = item
                protocols.append(prot)
            protocols.sort()
            return protocols
        else:
            return "No pudo ejecutarse la orden"

    def make_index(self, doc_name, protocol, number_ini, number_fin):
        """Makes the index and print it in a docx document.

        The name of the file is passed as a parameter.
        The protocol and number range is passed too as a parameter.
        """
        temp = string.ascii_uppercase
        upper = temp.replace("N", "NÑ")

        document = word.Documento(doc_name)

        for letter in upper:
            data = self.database.get_table_by_letter(protocol, letter,
                                                     number_ini, number_fin)
            if data:
                page = []
                for doc in data:
                    number = doc[0]
                    body = doc[1]
                    if doc[2] == "":
                        body += ". "
                    else:
                        body += "; "
                        body += doc[2]
                        body += ". "
                    body += doc[3]
                    body += "."
                    page.append([number, body])
                document.write_page(letter, page)

    def get_doc_by_number(self, protocol, number):
        doc = self.database.doc_by_number(protocol, number)

        if doc is None:
            return "Error al acceder a la base de datos."
        else:
            return doc

    def delete_document(self, protocol, number):
        status = self.database.delete_data(protocol, number)

        if status is None:
            return "Error al acceder a la base de datos."
        else:
            return True

    def search_user(self, protocol, string):
        p1 = self.database.search(protocol, "princCompareciente", string)
        p2 = self.database.search(protocol, "otrosComparecientes", string)

        set_results = set()
        if p1 is not None and p2 is not None:
            for item in p1:
                number, m_name, o_name, tramit = item

                s = str(number) + ".- "
                s += tramit + ". "
                s += m_name
                if o_name != "":
                    s += ", " + o_name
                s += "."
                set_results.add(s)
            for item in p2:
                number, m_name, o_name, tramit = item

                s = str(number) + ".- "
                s += tramit + ". "
                s += m_name
                if o_name != "":
                    s += ", " + o_name
                s += "."
                set_results.add(s)

            results = list(set_results)
            results.sort()
            return results
        else:
            return "No pudo ejecutarse la orden"


if __name__ == "__main__":
    a = Indice(max_recent_docs=2)
    print(a.search_user("Protocolo2017", "M"))
    # a.make_index("Prueba.docx", "Protocolo2017", 1, 6)
    p = "Protocolo2018"
    # print(a.del_protocol(p))
    # print(a.new_protocol(p))
    # print(a.view_protocols())
    # print(a.view_protocol(p))
    # print(a.recent_docs)
    # print(a.add_doc(p, 1, "Alberto", "","Permuta"))
    # print(a.recent_docs)
    # print(a.add_doc(p, 2, "Al", "Alexis","Compraventa"))
    # print(a.recent_docs)
    # print(a.add_doc(p, 3, "Daysi", "Felipe","Compraventa"))
    # print(a.recent_docs)
    # print(a.view_protocol(p))
    # print(a.is_valid_number(p, 4))
    # print(a.next_number(p))
