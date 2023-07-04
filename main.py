import sqlite3
import sys

import pandas as pd
from PySide6 import QtCore
from PySide6.QtCore import QCoreApplication
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidgetItem

from cadastro_1 import Ui_MainWindow
from database import Data_base
from functions import consulta_cnpj


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        super().setupUi(self)
        self.setWindowTitle("ASATEC - Sistema de Cadastro de Empresas")
        appIcon = QIcon("icon_projeto.ico")
        self.setWindowIcon(appIcon)

        # Toggle button
        self.btn_toggle.clicked.connect(self.left_menu)

        # Páginas do sistema
        self.btn_home.clicked.connect(lambda: self.pages.setCurrentWidget(self.pg_home))
        self.btn_cadastrar.clicked.connect(
            lambda: self.pages.setCurrentWidget(self.pg_cadastrar)
        )
        self.btn_sobre.clicked.connect(
            lambda: self.pages.setCurrentWidget(self.pg_sobre)
        )
        self.btn_contatos.clicked.connect(
            lambda: self.pages.setCurrentWidget(self.pg_contatos)
        )

        # Preencher automaticamente os dados do CNPJ
        self.txt_cnpj.editingFinished.connect(self.consult_api)

        # Cadastrar empresa no banco de dados
        self.btn_cadastrar_empresa.clicked.connect(self.cadastrar_empresas)

        # Alterar empresas
        self.btn_alterar.clicked.connect(self.atualizar_empresas)

        # Listar Empresas
        self.buscar_empresas()

        # Excluir Empresa
        self.btn_excluir.clicked.connect(self.deletar_empresa)

        # Gerar Excel
        self.btn_excel.clicked.connect(self.gerar_excel_banco)

        # Atualizar Página
        self.tabWidget.tabBarClicked.connect(self.atualizar_pagina)

    def left_menu(self):
        width = self.left_container.width()

        if width == 2:
            new_width = 200
        else:
            new_width = 2

        self.animation = QtCore.QPropertyAnimation(self.left_container, b"maximumWidth")
        self.animation.setDuration(500)
        self.animation.setStartValue(width)
        self.animation.setEndValue(new_width)
        self.animation.setEasingCurve(QtCore.QEasingCurve.InOutQuart)  # type: ignore
        self.animation.start()

    def consult_api(self):
        campos = consulta_cnpj(
            self.txt_cnpj.text().replace(".", "").replace("/", "").replace("-", "")
        )

        self.txt_nome.setText(campos[0])
        self.txt_logradouro.setText(campos[1])
        self.txt_numero.setText(campos[2])
        self.txt_complemento.setText(campos[3])
        self.txt_bairro.setText(campos[4])
        self.txt_municipio.setText(campos[5])
        self.txt_uf.setText(campos[6])
        self.txt_cep.setText(campos[7].replace(".", "").replace("-", ""))
        self.txt_telefone.setText(
            campos[8].replace("(", "").replace("-", "").replace(")", "")
        )
        self.txt_email.setText(campos[9])

    def cadastrar_empresas(self):
        db = Data_base()
        db.connect()

        fullDataSet = [
            self.txt_cnpj.text().replace(".", "").replace("/", "").replace("-", ""),
            self.txt_nome.text(),
            self.txt_logradouro.text(),
            self.txt_numero.text(),
            self.txt_complemento.text(),
            self.txt_bairro.text(),
            self.txt_municipio.text(),
            self.txt_uf.text(),
            self.txt_cep.text(),
            self.txt_telefone.text(),
            self.txt_email.text(),
        ]
        # Verifica se a empresa já está cadastrada
        if not db.select_company(fullDataSet[0]):
            # Cadastrar no banco de dados
            resp = db.register_company(fullDataSet)
        else:
            resp = "JÁ CADASTRADO"

        if resp == "OK":
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Cadastro Realizado")
            msg.setText("Cadastro realizado com sucesso")
            msg.exec()
            db.close_connection()
            return
        elif resp == "JÁ CADASTRADO":
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Cadastro já foi realizado")
            msg.setText("Empresa já se encontra cadastrada")
            msg.exec()
            db.close_connection()
            return
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            msg.setText(
                "Erro ao cadastrar, verifique se as informações foram preenchidas corretamente"
            )
            msg.exec()
            db.close_connection()
            return

    def buscar_empresas(self):
        db = Data_base()
        db.connect()
        result = db.select_all_companies()
        self.tb_company.clearContents()

        if result:
            self.tb_company.setRowCount(len(result))

            for row, text in enumerate(result):
                for column, data in enumerate(text):
                    self.tb_company.setItem(row, column, QTableWidgetItem(str(data)))

            db.close_connection()

        else:
            return

    def atualizar_empresas(self):
        dados = []
        update_dados = []

        for row in range(self.tb_company.rowCount()):
            for column in range(self.tb_company.columnCount()):
                dados.append(self.tb_company.item(row, column).text())
            update_dados.append(dados)
            dados = []

        # Atualizar dados no Banco
        db = Data_base()
        db.connect()

        for emp in update_dados:
            db.update_company(tuple(emp))

        db.close_connection()

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)  # type:ignore
        msg.setWindowTitle("Atualização de Dados")
        msg.setText("Dados atualizados com sucesso")
        msg.exec()

        self.tb_company.reset()
        self.buscar_empresas()

    def deletar_empresa(self):
        db = Data_base()
        db.connect()

        msg = QMessageBox()
        msg.setWindowTitle("Excluir")
        msg.setText("Este registro será excluído.")
        msg.setInformativeText("Você tem certeza que deseja excluir?")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        resp = msg.exec()

        if resp == QMessageBox.Yes:
            cnpj = (
                self.tb_company.selectionModel()
                .currentIndex()
                .siblingAtColumn(0)
                .data()
            )
            result = db.delete_company(cnpj)
            self.buscar_empresas()

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)  # type:ignore
            msg.setWindowTitle("Empresas")
            msg.setText(result)
            msg.exec()

        db.close_connection()

    def gerar_excel(self):
        dados = []
        all_dados = []

        for row in range(self.tb_company.rowCount()):
            for column in range(self.tb_company.columnCount()):
                dados.append(self.tb_company.item(row, column).text())

            all_dados.append(dados)
            dados = []

        columns = (
            "CNPJ",
            "NOME",
            "LOGRADOURO",
            "NUMERO",
            "COMPLEMENTO",
            "BAIRRO",
            "MUNICIPIO",
            "UF",
            "CEP",
            "TELEFONE",
            "EMAIL",
        )

        empresas = pd.DataFrame(all_dados, columns=columns)
        empresas.to_excel("Empresas.xlsx", sheet_name="empresas", index=False)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Excel")
        msg.setText("Relatório Excel gerado com sucesso!")
        msg.exec()

    def gerar_excel_banco(self):
        cnx = sqlite3.connect("db_cadastro")
        empresas = pd.read_sql_query("""SELECT * FROM Empresas""", cnx)
        empresas.to_excel("Empresas.xlsx", sheet_name="empresas", index=False)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Excel")
        msg.setText("Relatório Excel gerado com sucesso!")
        msg.exec()

    def atualizar_pagina(self):
        self.tb_company.reset()
        self.buscar_empresas()


if __name__ == "__main__":
    db = Data_base("db_cadastro")
    db.connect()
    db.create_table_company()
    db.close_connection()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()
