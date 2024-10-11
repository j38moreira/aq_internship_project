from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtGui, QtWidgets, QtCore
from PyQt5.QtGui import *
import pyodbc
from db_config import DB_CONFIG, TABLE_CONFIG
import re
import string
import logging
import egoi
import egoi_transac
import egoi_auto_tag
import sys
import json

logging.basicConfig(level=logging.INFO, filename='py_log.log', filemode='w', format='%(asctime)s %(levelname)s %(message)s')


class App(QMainWindow):
    global conn
    conn = pyodbc.connect(
        f"DRIVER={DB_CONFIG['DRIVER']};"
        f"SERVER={DB_CONFIG['SERVER']};"
        f"DATABASE={DB_CONFIG['DATABASE']};"
        f"TRUSTED_CONNECTION={DB_CONFIG['TRUSTED_CONNECTION']};"
    )
    global cursor
    cursor = conn.cursor()
    dark_mode = False

    def __init__(self):
        super().__init__()
        # self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))
        self.central_widget = QWidget()
        self.layout = QVBoxLayout(self.central_widget)
        self.cor = None
        self.tabela = QTableWidget()
        self.PaginaInicial()

        self.group_boxes_to_hide = ['procGroupBox', 'addGroupBox', 'atuGroupBox', 'listarGroupBox',
                                    'esquecerGroupBox', 'criarTGroupBox', 'obterTGroupBox',
                                    'tagclienteGroupBox', 'enviaremailGroupBox', 'obtercampGroupBox',
                                    'criarcampGroupBox', 'remtagclienteGroupBox', 'obtertempGroupBox',
                                    'enviaremailtGroupBox', 'criartempGroupBox']

        self.dark_mode = False
        self.load_dark_mode_state()
        
        self.toggle_action = QAction("Toggle Mode", self)
        self.toggle_action.triggered.connect(self.toggleMode)
        
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        self.updateMode()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.offset = event.pos()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.offset is not None and event.buttons() == QtCore.Qt.LeftButton:
            self.move(self.pos() + event.pos() - self.offset)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.offset = None
        super().mouseReleaseEvent(event)

    def hideGroupBox(self):
        for group_box in self.group_boxes_to_hide:
            if hasattr(self, group_box):
                getattr(self, group_box).hide()

    def PaginaInicial(self):
        # self.setWindowTitle('Automatização E-goi')
        self.setWindowFlag(Qt.FramelessWindowHint)
        egoi_auto_tag.AutoTag()

        menubar = self.menuBar()
        font = QFont('Segoe UI Semibold', 11)
        menubar.setFont(font)

        acoescliente = menubar.addMenu('Ações Cliente')

        adicionar_action = QAction('Adicionar | Editar', self)
        adicionar_action.triggered.connect(self.AddCliente)
        acoescliente.addAction(adicionar_action)

        listarclientes_action = QAction('Listar', self)
        listarclientes_action.triggered.connect(self.ListarCliente)
        acoescliente.addAction(listarclientes_action)

        acoesegoi = menubar.addMenu('Ações E-goi')

        clientesubmenu = acoesegoi.addMenu('Cliente')

        clientestagssubmenu = clientesubmenu.addMenu('Clientes Tags')

        tagcliente_action = QAction('Atribuir | Remover tags a cliente', self)
        tagcliente_action.triggered.connect(self.TagsCliente)
        clientestagssubmenu.addAction(tagcliente_action)

        addegoi_action = QAction('Esquecer Clientes E-goi', self)
        addegoi_action.triggered.connect(self.AtualizarEgoi)
        clientesubmenu.addAction(addegoi_action)

        campsubmenu = acoesegoi.addMenu('Campanhas')

        camp_action = QAction('Criar | Obter campanha', self)
        camp_action.triggered.connect(self.CriarCampanhaEmail)
        campsubmenu.addAction(camp_action)

        emailsubmenu = acoesegoi.addMenu('Email')

        enviarcemail_action = QAction('Enviar Email c/Template', self)
        enviarcemail_action.triggered.connect(self.EnviarEmailT)
        emailsubmenu.addAction(enviarcemail_action)

        enviarsemail_action = QAction('Enviar Email s/Template', self)
        enviarsemail_action.triggered.connect(self.EnviarEmail)
        emailsubmenu.addAction(enviarsemail_action)

        tagsubmenu = acoesegoi.addMenu('Tags')

        criartag_action = QAction('Atualizar | Criar tags', self)
        criartag_action.triggered.connect(self.CriarTags)
        tagsubmenu.addAction(criartag_action)

        templatesmenu = acoesegoi.addMenu('Templates')

        criartemplate_action = QAction('Criar | Obter Templates', self)
        criartemplate_action.triggered.connect(self.CriarTemplate)
        templatesmenu.addAction(criartemplate_action)

        definicoesbar = menubar.addMenu('Definições')
        
        self.darkmode_action = QAction('Dark Mode', self)
        self.darkmode_action.triggered.connect(self.toggleMode)
        definicoesbar.addAction(self.darkmode_action)

        sair_action = QAction('Sair', self)
        sair_action.triggered.connect(qApp.quit)
        definicoesbar.addAction(sair_action)

        rightMenu = QMenuBar()
        rightMenuLayout = QHBoxLayout()
        rightMenuLayout.setContentsMargins(0, 0, 0, 0)
        rightMenu.setLayout(rightMenuLayout)

        self.mini_icon_path = 'images/mini_p.png'
        self.close_icon_path = 'images/close_p.png'

        self.miniaction = QAction(QIcon(self.mini_icon_path), "", self)
        self.miniaction.triggered.connect(self.showMinimized)
        rightMenu.addAction(self.miniaction)

        self.miniaction1 = QAction(QIcon(self.close_icon_path), "", self)
        self.miniaction1.triggered.connect(qApp.quit)
        rightMenu.addAction(self.miniaction1)

        menubar.setCornerWidget(rightMenu)

    def toggleMode(self):
        self.dark_mode = not self.dark_mode
        self.save_dark_mode_state()
        self.updateMode()

        if self.dark_mode:
            self.darkmode_action.setText("Light Mode")
        else:
            self.darkmode_action.setText("Dark Mode")

    def updateMode(self):
        if self.dark_mode:
            self.setStyleSheet("background-color: #333; color: white;")
            self.mini_icon_path = 'images/mini_wh.png'
            self.close_icon_path = 'images/close_wh.png'
        else:
            self.setStyleSheet("")
            self.mini_icon_path = 'images/mini_p.png'
            self.close_icon_path = 'images/close_p.png'
        
        self.miniaction.setIcon(QIcon(self.mini_icon_path))
        self.miniaction1.setIcon(QIcon(self.close_icon_path))

        

        self.corTabela()
        
    def save_dark_mode_state(self):
        with open("dark_mode_state.json", "w") as file:
            json.dump({"dark_mode": self.dark_mode}, file)

    def load_dark_mode_state(self):
        try:
            with open("dark_mode_state.json", "r") as file:
                data = json.load(file)
                self.dark_mode = data.get("dark_mode", False)
        except FileNotFoundError:
            self.dark_mode = False
    
    
    def AddCliente(self):
        self.hideGroupBox()
        
        self.addGroupBox = QGroupBox('Adicionar Cliente')
        self.layout.addWidget(self.addGroupBox)
        self.central_widget.setLayout(self.layout)

        self.nomeLineEdit = QLineEdit()
        reg_nomeLineEdit = QRegExp('[A-Za-zÀ-ú]{0,255}')
        self.nomeLineEdit.setValidator(QRegExpValidator(reg_nomeLineEdit))
        self.nomeLineEdit.returnPressed.connect(self.onAdicionarClicked)
        
        self.apelidoLineEdit = QLineEdit()
        reg_apelidoLineEdit = QRegExp('[A-Za-zÀ-ú]{0,255}')
        self.apelidoLineEdit.setValidator(QRegExpValidator(reg_apelidoLineEdit))
        self.apelidoLineEdit.returnPressed.connect(self.onAdicionarClicked)
        
        self.emailLineEdit = QLineEdit()
        self.emailLineEdit.returnPressed.connect(self.onAdicionarClicked)
        
        self.teleLineEdit = QLineEdit()
        reg_teleLineEdit = QRegExp('[0-9]{0,9}')
        self.teleLineEdit.setValidator(QRegExpValidator(reg_teleLineEdit))
        self.teleLineEdit.returnPressed.connect(self.onAdicionarClicked)
        
        self.idiomaComboBox = QComboBox()
        self.idiomaComboBox.addItems(['Português'])
        
        self.datanascE = QtWidgets.QDateEdit(calendarPopup=True)
        self.datanascE.setDateTime(QtCore.QDateTime.currentDateTime())
        calendar = self.datanascE.calendarWidget()
        calendar.setFirstDayOfWeek(QtCore.Qt.Monday)
        calendar.setGridVisible(False)

        self.rgpd = QCheckBox('Consente com o RGPD')
        
        self.FormularioAdicionar()

        addBtn = QPushButton('Adicionar', self)
        addBtn.clicked.connect(self.onAdicionarClicked)
        layout = self.addGroupBox.layout()
        layout.addRow(addBtn)
        
        self.layout.setSpacing(20)
        
        self.procGroupBox = QGroupBox('Editar Cliente')

        self.layout.addWidget(self.procGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.emailprocLineEdit = QLineEdit()
        self.emailprocLineEdit.returnPressed.connect(self.onProcurarClicked)
        
        self.FormularioProcurar()
        
        procBtn = QPushButton('Editar', self)
        procBtn.clicked.connect(self.onProcurarClicked)
        layout = self.procGroupBox.layout()
        layout.addRow(procBtn) 
        
    def FormularioAdicionar(self):
        layout = QFormLayout()
        layout.addRow(QLabel('Nome'), self.nomeLineEdit)
        layout.addRow(QLabel('Apelido'), self.apelidoLineEdit)
        layout.addRow(QLabel('Email'), self.emailLineEdit)
        layout.addRow(QLabel('Telemóvel'), self.teleLineEdit)
        layout.addRow(QLabel('Idioma'), self.idiomaComboBox)
        layout.addRow(QLabel('Data de Nascimento'), self.datanascE)
        layout.addRow(QLabel('RGPD'), self.rgpd)
        self.addGroupBox.setLayout(layout)      
 
    def onAdicionarClicked(self):
        nome = self.nomeLineEdit.text()
        apelido = self.apelidoLineEdit.text()
        email = self.emailLineEdit.text()
        telemovel = self.teleLineEdit.text()
        idioma = self.idiomaComboBox.currentText()
        datanasc = self.datanascE.date().toString('yyyy-MM-dd')
        rgpd_aceite = 1 if self.rgpd.isChecked() else 0
        
        if idioma == 'Português':
            idioma = 'pt'
            
        nome_cap = string.capwords(nome)
        apelido_cap = string.capwords(apelido)
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'

        if nome_cap and apelido_cap and email and telemovel and idioma and datanasc:
            if (re.fullmatch(regex_email, email)):
                if egoi_transac.ValidarEmail(email):
                    if telemovel.strip():
                        try:
                            telemovel = int(telemovel)
                            if (210000000 <= telemovel < 297000000) or (910000000 <= telemovel < 970000000):
                                try:                                
                                    cursor.execute(
                                        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE email_cliente = ?",
                                        (email,)
                                    )
                                    resEmail = cursor.fetchall()
                                        
                                    if not resEmail:
                                    
                                        msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja adicionar o cliente?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                                        
                                        if msg_confirmacao == QMessageBox.Yes:
                                            cursor.execute(
                                                f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] (nome_cliente, apelido_cliente, email_cliente, telemovel_cliente, idioma_cliente, data_nasc, rgpd_cliente) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                                (nome_cap, apelido_cap, email, telemovel, idioma, datanasc, rgpd_aceite, )
                                            )
                                            conn.commit()
                                            
                                            if rgpd_aceite == 1:
                                                egoi.CriarContactosEgoi()

                                            msg_info = 'Cliente adicionado'
                                            QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                                            logging.info('Cliente adicionado: Email: '+email, exc_info=False)
                                            
                                            
                                            self.nomeLineEdit.clear()
                                            self.apelidoLineEdit.clear()
                                            self.emailLineEdit.clear()
                                            self.teleLineEdit.clear()
                                        
                                    else:
                                        msg_erro = 'Email já existente'
                                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)

                                except Exception as e:
                                    logging.error('Erro', exc_info=True) 
                                    print('Ver py_log.log')
                            else:
                                msg_erro = 'Número de telemóvel incorreto'
                                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                        except ValueError:
                            msg_erro = 'Número de telemóvel incorreto'
                            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                else:
                    msg_erro = 'Email invalido'
                    QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)    
            else:
                msg_erro = 'Email incorreto'
                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
    
    def ListarCliente(self):
        self.hideGroupBox()
        self.listarGroupBox = QGroupBox('')
        self.layout.addWidget(self.listarGroupBox)
        self.central_widget.setLayout(self.layout)
        self.ListarListaCliente()
        self.corTabela()

    def ListarListaCliente(self):
        layout = QVBoxLayout()
        self.listarGroupBox.setLayout(layout)

        hbox = QHBoxLayout()

        self.column_combobox = QComboBox()
        hbox.addWidget(self.column_combobox)

        self.filter_edit = QLineEdit()
        hbox.addWidget(self.filter_edit)
        
        layout.addLayout(hbox)

        self.column_combobox.addItems(["Nome", "Email", "Telemóvel"])
        self.filter_edit.textChanged.connect(self.filter_table)

        scroll_area = QScrollArea(self.listarGroupBox)
        scroll_area.setWidgetResizable(True)
        
        layout.addWidget(scroll_area)

        self.tabela = QTableWidget(self)
        self.tabela.setStyleSheet("border: 1px;")
        self.tabela.setColumnCount(4)
        self.tabela.setHorizontalHeaderLabels(['Nome', 'Email', 'Telemóvel', 'RGPD'])

        header_font = self.tabela.horizontalHeader().font()
        header_font.setBold(True)
        self.tabela.horizontalHeader().setFont(header_font)

        scroll_area.setWidget(self.tabela)
        
        try:
            cursor.execute(
                f"SELECT nome_cliente, apelido_cliente, email_cliente, telemovel_cliente, rgpd_cliente FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] ORDER BY nome_cliente ASC, apelido_cliente"
            )
            resCliente = cursor.fetchall()              
            
            for row_index, row in enumerate(resCliente):
                self.tabela.insertRow(row_index)
                
                nomecomp = f"{row[0]} {row[1]}"
                item = QTableWidgetItem(nomecomp)
                self.tabela.setItem(row_index, 0, item)

                item = QTableWidgetItem(str(row[2]))
                self.tabela.setItem(row_index, 1, item)

                item = QTableWidgetItem(str(row[3]))
                self.tabela.setItem(row_index, 2, item)

                rgpd = "Aceite" if row[4] == 1 else "Rejeitado"    
                label = QLabel()
                if rgpd == "Aceite":
                    pixmap = QPixmap("images/accept.png").scaled(25, 25)
                    label.setPixmap(pixmap)
                else:
                    pixmap = QPixmap("images/reject.png").scaled(25, 25)
                    label.setPixmap(pixmap)
                label.setAlignment(Qt.AlignCenter)
                self.tabela.setCellWidget(row_index, 3, label)

            self.tabela.resizeColumnsToContents()
            self.tabela.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
            self.tabela.horizontalHeaderItem(3).setToolTip("Para filtrar : Aceite ou Rejeitado")
            self.tabela.verticalHeader().setVisible(False)

            scroll_area.setWidget(self.tabela)

        except Exception as e:
            logging.error('Erro', exc_info=True)
            print('Ver py_log.log')

    def filter_table(self):
        column_index = self.column_combobox.currentIndex()
        filter_text = self.filter_edit.text().strip().lower()

        for row_index in range(self.tabela.rowCount()):
            item = self.tabela.item(row_index, column_index)
            if item is not None:
                cell_text = item.text().strip().lower()
                self.tabela.setRowHidden(row_index, filter_text not in cell_text)

    def corTabela(self):
        light_colors = [QColor(255, 255, 255), QColor(195, 195, 195)]
        dark_colors = [QColor(40, 40, 40), QColor(70, 70, 70)]

        colors = dark_colors if self.dark_mode else light_colors

        for i in range(self.tabela.rowCount()):
            color_index = i % len(colors)
            color = colors[color_index]

            for j in range(self.tabela.columnCount()):
                item = self.tabela.item(i, j)
                if item is not None:
                    item.setBackground(color)
            
            rgpd_widget = self.tabela.cellWidget(i, 3)
            if isinstance(rgpd_widget, QLabel):
                rgpd_widget.setStyleSheet(f"background-color: {color.name()};")

    def FormularioProcurar(self):
        layout_p = QFormLayout()
        layout_p.addRow(QLabel('Email'), self.emailprocLineEdit)
        self.procGroupBox.setLayout(layout_p)
    
    def onProcurarClicked(self):
        global email_p
        email_p = self.emailprocLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        if email_p:
            if (re.fullmatch(regex_email, email_p)):
                try:
                    cursor.execute(
                        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE email_cliente = ?",
                        (email_p,)
                    )
                    resEmail_p = cursor.fetchall()
                    
                    if resEmail_p:
                        editarcliente_window = EditarWindow(resEmail_p)
                        editarcliente_window.setStyleSheet(self.styleSheet()) 
                        editarcliente_window.exec_() 
                    else:
                        msg_erro = 'Email não existente'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                        
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')
            else:
                msg_erro = 'Email incorreto'
                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)        
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
        
    def AtualizarEgoi(self):
        self.hideGroupBox()
            
        # self.atuGroupBox = QGroupBox('Adicionar Clientes na lista E-goi')
        # self.layout.addWidget(self.atuGroupBox)
        # self.central_widget.setLayout(self.layout)

        # self.FormularioAtualizar()

        # atBtn = QPushButton('Adicionar', self)
        # atBtn.clicked.connect(self.onAdicionarEgoiClicked)
        # layout = self.atuGroupBox.layout()
        # layout.addRow(atBtn)
        
        self.esquecerGroupBox = QGroupBox('Esquecer Cliente na lista E-goi')    
        self.layout.addWidget(self.esquecerGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.emailesqLineEdit = QLineEdit()
        self.emailesqLineEdit.returnPressed.connect(self.onEsquecerClicked)
        
        self.FormularioEsquecer()
        
        esqBtn = QPushButton('Esquecer', self)
        esqBtn.clicked.connect(self.onEsquecerClicked)
        layout = self.esquecerGroupBox.layout()
        layout.addRow(esqBtn)
        
    # def FormularioAtualizar(self):
    #     layout = QFormLayout(self.atuGroupBox)
    
    # def onAdicionarEgoiClicked(self):
    #     s = egoi.CriarContactosEgoi()
    #     if s:
    #         QMessageBox.information(self, 'Sucesso', 'Clientes foram adicionados ao sistema E-goi com sucesso!', QMessageBox.Ok)
    #     else:
    #         QMessageBox.information(self, 'Informação', 'Todos os clientes já estavam no sistema E-goi.', QMessageBox.Ok)
              
    def FormularioEsquecer(self):
        layout_e = QFormLayout()
        layout_e.addRow(QLabel('Email'), self.emailesqLineEdit)
        self.esquecerGroupBox.setLayout(layout_e)
    
    def onEsquecerClicked(self):
        email_e = self.emailesqLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        
        if email_e:
            if(re.fullmatch(regex_email, email_e)):
                try:
                    cursor.execute(
                        f"SELECT id_cliente FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE email_cliente = ?",
                        (email_e,)
                    )
                    resEmail_e = cursor.fetchall()
                    
                    if resEmail_e:
                        msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja continuar?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                        if msg_confirmacao == QMessageBox.Yes:
                            egoi.EsquecerContactosEgoi(resEmail_e[0][0])
                           
                            msg_info = 'Cliente esquecido com sucesso!'
                            QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                    else:
                        msg_erro = 'Email incorreto'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)     
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)      
    
    def CriarTags(self):
        self.hideGroupBox()
                
        self.criarTGroupBox = QGroupBox('Criar tag')
        self.layout.addWidget(self.criarTGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.nometagEdit = QLineEdit()
        self.nometagEdit.returnPressed.connect(self.onCriartagClicked)
        
        self.corbtn = QPushButton('Escolher cor')
        self.corbtn.clicked.connect(self.onCorClicked)
        
        self.prevcor = QLabel()
        
        self.FormularioCriartag()
        
        criarFBtn = QPushButton('Criar', self)
        criarFBtn.clicked.connect(self.onCriartagClicked)
        layout = self.criarTGroupBox.layout()
        layout.addRow(criarFBtn)
            
        self.obterTGroupBox = QGroupBox('Atualizar tags')
        self.layout.addWidget(self.obterTGroupBox)
        self.central_widget.setLayout(self.layout)

        self.FormularioAtualizarTags()

        atBtn = QPushButton('Atualizar tags no sistema', self)
        atBtn.clicked.connect(self.onAtualizarTagsClicked)
        layout = self.obterTGroupBox.layout()
        layout.addRow(atBtn)
        
        ateBtn = QPushButton('Atualizar tags no sistema e-goi', self)
        # ateBtn.clicked.connect(self.onAdicionarTagsClicked)
        layout = self.obterTGroupBox.layout()
        layout.addRow(ateBtn)
    
    def onCorClicked(self):
        self.openCorDialog()
        
    def openCorDialog(self):
        color_dialog = QColorDialog(self)

        cor = color_dialog.getColor()

        if cor.isValid():
            self.cor = cor
            self.prevcor.setStyleSheet(f"background-color: {cor.name()}; border :1px solid black;")
                           
    def FormularioCriartag(self):
        layout = QFormLayout()
        layout.addRow(QLabel('Nome'), self.nometagEdit)
        layout.addRow(QLabel('Cor'), self.corbtn)
        layout.addRow(QLabel('Preview'), self.prevcor)
        self.criarTGroupBox.setLayout(layout)
        
    def onCriartagClicked(self):
        nome_tag = self.nometagEdit.text()

        if nome_tag and self.cor:
            try:
                cursor.execute(
                    f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] WHERE nome_tag = ?",
                    (nome_tag,)
                )
                resNome = cursor.fetchall()

                if not resNome:
                    msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja adicionar a tag?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

                    if msg_confirmacao == QMessageBox.Yes:
                        cursor.execute(
                            f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] (nome_tag, cor_tag) VALUES (?, ?)",
                            (nome_tag, self.cor.name().upper())
                        )
                        conn.commit()

                        msg_info = 'Tag adicionada'
                        QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                        logging.info('Tag adicionada: ' + nome_tag, exc_info=False)

                        self.nometagEdit.clear()
                        self.cor = None

                else:
                    msg_erro = 'Tag já existente'
                    QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)

            except Exception as e:
                logging.error('Erro', exc_info=True)
                print('Ver py_log.log')

        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
            
    def FormularioAtualizarTags(self):
        layout = QFormLayout(self.obterTGroupBox)
    
    def onAtualizarTagsClicked(self):
        f = egoi.AtualizartagsDB()
        if f:
            QMessageBox.information(self, 'Sucesso', 'Tags atualizadas com sucesso!', QMessageBox.Ok)
        else:
            QMessageBox.information(self, 'Informação', 'Tags já se encontravam atualizadas.', QMessageBox.Ok)
            
    def onAdicionarTagsClicked(self):
        a = egoi.AtualizarTagsEgoi()
        if a:
            QMessageBox.information(self, 'Sucesso', 'Tags adicionadas ao sistema e-goi com sucesso!', QMessageBox.Ok)
        else:
            QMessageBox.information(self, 'Informação', 'Todas as tags encontradas no sistema já estavam no sistema e-goi.', QMessageBox.Ok)
            
    def TagsCliente(self):
        self.hideGroupBox()
        
        self.tagclienteGroupBox = QGroupBox('Atribuir ou Remover Tag a cliente')
        self.layout.addWidget(self.tagclienteGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.emailtclienteLineEdit = QLineEdit()
        
        self.FormularioProcurarTCliente()
        
        proctcBtn = QPushButton('Atribuir', self)
        proctcBtn.setStyleSheet('QPushButton {color: #11bf29}')
        proctcBtn.clicked.connect(self.onProcurarTClienteClicked)  
        layout = self.tagclienteGroupBox.layout()
        layout.addRow(proctcBtn)
        
        remtcBtn = QPushButton('Remover', self)
        remtcBtn.setStyleSheet('QPushButton {color: red}')
        remtcBtn.clicked.connect(self.onRemoverTClienteClicked)  
        layout = self.tagclienteGroupBox.layout()
        layout.addRow(remtcBtn)
    
    def FormularioProcurarTCliente(self):
        layout_tc = QFormLayout()
        layout_tc.addRow(QLabel('Email'), self.emailtclienteLineEdit)
        self.tagclienteGroupBox.setLayout(layout_tc)
    
    def onProcurarTClienteClicked(self):
        global email_tc
        email_tc = self.emailtclienteLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        
        if email_tc:
            if(re.fullmatch(regex_email, email_tc)):
                try:
                    cursor.execute(
                        f"SELECT ec.id_egoi_cliente "
                        f"FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] as ec "
                        f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] as c ON c.id_cliente = ec.id_cliente "
                        f"WHERE c.email_cliente = ?",
                        (email_tc,)
                    )
                    resEmail_tc = cursor.fetchall()
                                       
                    if resEmail_tc:
                        id_egoi_cliente = resEmail_tc[0][0]

                        cursor.execute(
                            f"SELECT id_egoi_tag, nome_tag FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] WHERE tag_id IS NOT NULL;"
                        )

                        tags = cursor.fetchall()

                        tagcliente_window = TagClienteWindow(id_egoi_cliente, tags)
                        tagcliente_window.setStyleSheet(self.styleSheet()) 
                        tagcliente_window.exec_()
                    else:
                        msg_erro = 'Email não existente'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                        
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')     
            else:
                msg_erro = 'Email incorreto'
                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
               
    def onRemoverTClienteClicked(self):
        global email_tc
        email_tc = self.emailtclienteLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        
        if email_tc:
            if(re.fullmatch(regex_email, email_tc)):
                try:
                    cursor.execute(
                        f"SELECT ec.id_egoi_cliente "
                        f"FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] as ec "
                        f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] as c ON c.id_cliente = ec.id_cliente "
                        f"WHERE c.email_cliente = ?",
                        (email_tc,)
                    )
                    resEmail_tc = cursor.fetchall()
                                       
                    if resEmail_tc:
                        id_egoi_cliente = resEmail_tc[0][0]

                        cursor.execute(
                            f"SELECT id_egoi_tag, nome_tag FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] WHERE tag_id IS NOT NULL;"
                        )

                        tags = cursor.fetchall()

                        remtagcliente_window = RemoverTagClienteWindow(id_egoi_cliente, tags)
                        remtagcliente_window.setStyleSheet(self.styleSheet()) 
                        remtagcliente_window.exec_()
                        
                    else:
                        msg_erro = 'Email não existente'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                        
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')     
            else:
                msg_erro = 'Email incorreto'
                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
    
    def EnviarEmail(self):
        self.hideGroupBox()
        
        self.enviaremailGroupBox = QGroupBox('Enviar Email ao Cliente')
        self.layout.addWidget(self.enviaremailGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.emailtclienteLineEdit = QLineEdit()
        self.emailtclienteLineEdit.returnPressed.connect(self.onEnviarEmailClicked)
        
        self.FormularioEnviarEmail()
        
        proctcBtn = QPushButton('Enviar', self)
        proctcBtn.clicked.connect(self.onEnviarEmailClicked)  
        layout = self.enviaremailGroupBox.layout()
        layout.addRow(proctcBtn)
    
    def FormularioEnviarEmail(self):
        layout_ee = QFormLayout()
        layout_ee.addRow(QLabel('Email'), self.emailtclienteLineEdit)
        self.enviaremailGroupBox.setLayout(layout_ee)   
    
    def onEnviarEmailClicked(self):
        global email_ee
        email_ee = self.emailtclienteLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        
        if email_ee:
            if(re.fullmatch(regex_email, email_ee)):
                try: 
                    cursor.execute(
                        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE email_cliente = ?",
                        (email_ee,)
                    )
                    
                    resEmail = cursor.fetchall()
                    
                    if resEmail:
                        cursor.execute(
                            f"SELECT id_cliente FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE rgpd_cliente = 1 and email_cliente = ?",
                            (email_ee,)
                        )
                        resEmail_ee = cursor.fetchall()
                        if resEmail_ee:
                            if egoi.EnviarEmail(resEmail_ee):
                                QMessageBox.information(self, 'Sucesso', 'Email enviado com sucesso!', QMessageBox.Ok)
                            else:
                                QMessageBox.critical(self, 'Erro', 'Erro ao enviar o email.', QMessageBox.Ok)
                        else:
                            msg_erro = 'Cliente não consentiu o RGPD'
                            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)            
                    else:
                        msg_erro = 'Email incorreto'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)      
    
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')
    
    def EnviarEmailT(self):
        self.hideGroupBox()
        
        self.enviaremailtGroupBox = QGroupBox('Enviar Email ao Cliente')
        self.layout.addWidget(self.enviaremailtGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.emailtclienteLineEdit = QLineEdit()
        self.emailtclienteLineEdit.returnPressed.connect(self.onEnviarEmailTClicked)
        
        self.FormularioEnviarEmailT()
        
        proctcBtn = QPushButton('Enviar', self)
        proctcBtn.clicked.connect(self.onEnviarEmailTClicked)  
        layout = self.enviaremailtGroupBox.layout()
        layout.addRow(proctcBtn)
    
    def FormularioEnviarEmailT(self):
        layout_ee = QFormLayout()
        layout_ee.addRow(QLabel('Email'), self.emailtclienteLineEdit)
        self.enviaremailtGroupBox.setLayout(layout_ee)   
    
    def onEnviarEmailTClicked(self):
        global email_ee
        email_ee = self.emailtclienteLineEdit.text()
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'
        
        if email_ee:
            if(re.fullmatch(regex_email, email_ee)):
                try: 
                    cursor.execute(
                        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE email_cliente = ?",
                        (email_ee,)
                    )
                    
                    resEmail = cursor.fetchall()
                    
                    if resEmail:
                        cursor.execute(
                            f"SELECT email_cliente FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE rgpd_cliente = 1 and email_cliente = ?",
                            (email_ee,)
                        )
                        resEmail_ee = cursor.fetchall()
                        if resEmail_ee:
                            if egoi_transac.EnviarEmailTemplate(resEmail_ee):
                                QMessageBox.information(self, 'Sucesso', 'Email enviado com sucesso!', QMessageBox.Ok)
                            else:
                                QMessageBox.critical(self, 'Erro', 'Erro ao enviar o email.', QMessageBox.Ok)
                        else:
                            msg_erro = 'Cliente não consentiu o RGPD'
                            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)            
                    else:
                        msg_erro = 'Email incorreto'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)      
    
                except Exception as e:
                    logging.error('Erro', exc_info=True)
                    print('Ver py_log.log')
           
    def FormularioObterCampanha(self):
        layout = QFormLayout(self.obtercampGroupBox)
    
    def onAtualizarCampanhasClicked(self):
        f = egoi.ObterCampanhaBD()
        if f:
            QMessageBox.information(self, 'Sucesso', 'Campanhas atualizadas com sucesso!', QMessageBox.Ok)
        else:
            QMessageBox.information(self, 'Informação', 'Campanhas já se encontravam atualizadas.', QMessageBox.Ok)
            
    def CriarCampanhaEmail(self):
        self.hideGroupBox()
            
        self.criarcampGroupBox = QGroupBox('Criar campanha')
        self.layout.addWidget(self.criarcampGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.nomecampEdit = QLineEdit()
        self.nomecampEdit.returnPressed.connect(self.onCriarCampanhaClicked)

        self.FormularioCamapanhaObter()

        ccBtn = QPushButton('Criar', self)
        ccBtn.clicked.connect(self.onCriarCampanhaClicked)
        layout = self.criarcampGroupBox.layout()
        layout.addRow(ccBtn)
        
        self.obtercampGroupBox = QGroupBox('Atualizar campanhas')
        self.layout.addWidget(self.obtercampGroupBox)
        self.central_widget.setLayout(self.layout)

        self.FormularioObterCampanha()

        atBtn = QPushButton('Atualizar campanhas no sistema', self)
        atBtn.clicked.connect(self.onAtualizarCampanhasClicked)
        layout = self.obtercampGroupBox.layout()
        layout.addRow(atBtn)

    def FormularioCamapanhaObter(self):
        layout = QFormLayout(self.criarcampGroupBox)
        layout.addRow(QLabel('Nome'), self.nomecampEdit)
    
    def onCriarCampanhaClicked(self):
        nome_camp = self.nomecampEdit.text()
        if nome_camp:
            try:
                cursor.execute(
                    f"SELECT campaign_hash FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}] WHERE internal_name = ?",
                    (nome_camp,)
                )
                resNome = cursor.fetchall()

                if not resNome:
                    msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja criar a campanha?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

                    if msg_confirmacao == QMessageBox.Yes:
                        cursor.execute(
                            f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}] (internal_name) VALUES (?)",
                            (nome_camp,)
                        )
                        conn.commit()
                        
                        egoi.CriarCampanhaEgoi()
                        
                        msg_info = 'Campanha adicionada'
                        QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                        logging.info('Campanha adicionada: ' + nome_camp, exc_info=False)

                        self.nomecampEdit.clear()

                else:
                    msg_erro = 'Campanha já existente'
                    QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)

            except Exception as e:
                logging.error('Erro', exc_info=True)
                print('Ver py_log.log')
        
    def FormularioObterTemplate(self):
        layout = QFormLayout(self.obtertempGroupBox)
    
    def onObterTemplateClicked(self):
        t = egoi_transac.ObterTemplate()
        if t:
            QMessageBox.information(self, 'Sucesso', 'Templates atualizados com sucesso!', QMessageBox.Ok)
        else:
            QMessageBox.information(self, 'Informação', 'Templates já se encontravam atualizadas.', QMessageBox.Ok)
    
    def CriarTemplate(self):
        self.hideGroupBox()
        
        self.criartempGroupBox = QGroupBox('Criar templates')
        self.layout.addWidget(self.criartempGroupBox)
        self.central_widget.setLayout(self.layout)
        
        self.subjectLineEdit = QLineEdit()
        self.subjectLineEdit.returnPressed.connect(self.onCriarTemplateClicked)
        
        self.templateNameLineEdit = QLineEdit()
        self.templateNameLineEdit.returnPressed.connect(self.onCriarTemplateClicked)
        
        self.FormularioCriarTemplate()
        
        crBtn = QPushButton('Criar', self)
        crBtn.clicked.connect(self.onCriarTemplateClicked)
        layout = self.criartempGroupBox.layout()
        layout.addRow(crBtn)
        
        self.obtertempGroupBox = QGroupBox('Obter templates')
        self.layout.addWidget(self.obtertempGroupBox)
        self.central_widget.setLayout(self.layout)

        self.FormularioObterTemplate()

        atBtn = QPushButton('Atualizar templates no sistema', self)
        atBtn.clicked.connect(self.onObterTemplateClicked)
        layout = self.obtertempGroupBox.layout()
        layout.addRow(atBtn)

    def FormularioCriarTemplate(self):
        layout = QFormLayout()
        layout.addRow(QLabel('Nome do template'), self.templateNameLineEdit)
        layout.addRow(QLabel('Assunto do mail'), self.subjectLineEdit)
        self.criartempGroupBox.setLayout(layout)
    
    def onCriarTemplateClicked(self):
        nome_template = self.templateNameLineEdit.text()
        assunto = self.subjectLineEdit.text()
        
        if nome_template and assunto:
            try:                                
                cursor.execute(
                    f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] WHERE templateName = ?",
                    (nome_template,)
                )
                resEmail = cursor.fetchall()
                    
                if not resEmail:
                    
                    msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja criar o template?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    
                    if msg_confirmacao == QMessageBox.Yes:
                        cursor.execute(
                            f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] (templateName, subjectEmail) VALUES (?,?)",
                            (nome_template, assunto,)
                        )
                        conn.commit()
                        
                        egoi_transac.CriarTemplate()
                        
                        msg_info = 'Template criado'
                        QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                        logging.info('Template criado: '+nome_template, exc_info=False)
                        
                        
                        self.subjectLineEdit.clear()
                        self.templateNameLineEdit.clear()
                        
                    else:
                        msg_erro = 'Template já existente'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)

            except Exception as e:
                logging.error('Erro', exc_info=True)
                print('Ver py_log.log')
        
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                               
class EditarWindow(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        # self.setWindowTitle('Editar Cliente')
        self.setWindowFlag(Qt.FramelessWindowHint)
        # self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))
        self.setFixedSize(400, 220)
        self.setGeometry(900, 500, 0, 0)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        self.data = data

        self.layout = QFormLayout()

        self.editGroupBox = QGroupBox()
        self.layout.addWidget(self.editGroupBox)

        self.nome_editar = QLineEdit()
        self.nome_editar.setText(data[0][1])
        reg_nome_editar = QRegExp('[A-Za-zÀ-ú]{0,255}')
        self.nome_editar.setValidator(QRegExpValidator(reg_nome_editar))
        
        self.apelido_editar = QLineEdit()
        self.apelido_editar.setText(data[0][2])
        reg_apelido_editar = QRegExp('[A-Za-zÀ-ú]{0,255}')
        self.apelido_editar.setValidator(QRegExpValidator(reg_apelido_editar))
        
        self.email_editar = QLineEdit()
        self.email_editar.setText(data[0][3])
        
        self.telemovel_editar = QLineEdit()
        reg_telemovel = QRegExp('[0-9]{0,9}')
        self.telemovel_editar.setValidator(QRegExpValidator(reg_telemovel))
        self.telemovel_editar.setText(str(data[0][4]))
        
        self.rgpd_editar = QCheckBox('Consente com o RGPD')
        self.rgpd_comp = data[0][7]
        if data[0][7] == 1:
            self.rgpd_editar.setChecked(True) 
        
        self.FormularioEditar()
            
        guardarBtn = QPushButton('Guardar')
        guardarBtn.clicked.connect(self.onEditClicked)
        layout = self.editGroupBox.layout()
        layout.addRow(guardarBtn)
        
        sair_button = QPushButton('Fechar', self)
        sair_button.clicked.connect(self.close)
        layout.addRow(sair_button)

    def FormularioEditar(self):
        layout = QFormLayout(self.editGroupBox)
        layout.addRow(QLabel('Nome:'), self.nome_editar)
        layout.addRow(QLabel('Apelido:'), self.apelido_editar)
        layout.addRow(QLabel('Email:'), self.email_editar)
        layout.addRow(QLabel('Telemóvel:'), self.telemovel_editar)
        layout.addRow(QLabel('RGPD'), self.rgpd_editar)
        self.setLayout(self.layout)
        
    def onEditClicked(self):
        nomeEdit = self.nome_editar.text()
        apelidoEdit = self.apelido_editar.text()
        emailEdit = self.email_editar.text()
        telemovelEdit = self.telemovel_editar.text()
        rgpd_aceite = 1 if self.rgpd_editar.isChecked() else 0
        
        nomeEdit_cap = string.capwords(nomeEdit)
        apelidoEdit_cap = string.capwords(apelidoEdit)
        regex_email = r'\b[A-Za-z0-9._%+-]{6,64}@[A-Za-z0-9.-]{2,255}\.[A-Z|a-z]{2,7}\b'

        if nomeEdit_cap and apelidoEdit_cap and emailEdit and telemovelEdit:
            if re.fullmatch(regex_email, emailEdit):
                if telemovelEdit.strip():
                    try:
                        telemovelEdit = int(telemovelEdit)
                        if (210000000 <= telemovelEdit < 297000000) or (910000000 <= telemovelEdit < 970000000):
                            try:
                                cursor = conn.cursor()
                                
                                msg_confirmacao = QMessageBox.question(self, 'Confirmação', 'Deseja confirmar as alterações?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                        
                                if msg_confirmacao == QMessageBox.Yes:
                                    
                                    if rgpd_aceite == 0 and rgpd_aceite != self.rgpd_comp:
                                        egoi.EsquecerContactosEgoi(self.data[0][0])    
                                        
                                        msg_info = 'Cliente atualizado e esquecido com sucesso na e-goi'
                                        QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                                        logging.info('Cliente atualizado.', exc_info=False)
                                        
                                    else:
                                        cursor.execute(
                                            f"UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] SET nome_cliente = ?, apelido_cliente = ?, email_cliente = ?, telemovel_cliente = ?, rgpd_cliente = ? WHERE email_cliente = ?", 
                                            (nomeEdit_cap, apelidoEdit_cap, emailEdit, telemovelEdit, rgpd_aceite, email_p)
                                        )

                                        conn.commit()
                                        
                                        cursor.execute(
                                            f"""
                                            SELECT
                                                ec.contact_id,
                                                c.id_cliente
                                            FROM
                                                [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] AS ec
                                                JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] AS c
                                                    ON ec.id_cliente = c.id_cliente
                                            WHERE
                                                c.email_cliente = ? and c.rgpd_cliente = '1';
                                            """,
                                            (email_p,)
                                        )
                                        res = cursor.fetchall()
                                        
                                        if res:
                                            contact_id = res[0][0]
                                            id_cliente = res[0][1]
                        
                                            egoi.AtualizarClienteEgoi(contact_id, id_cliente, nomeEdit, apelidoEdit, emailEdit, telemovelEdit)
                                           
                                        msg_info = 'Cliente atualizado'
                                        QMessageBox.information(self, 'Sucesso', msg_info, QMessageBox.Ok)
                                        logging.info('Cliente atualizado. ' + email_p, exc_info=False)

                                    self.close()
                                    
                            except Exception as e:
                                logging.error('Erro', exc_info=True)
                                print('Ver py_log.log')
                        else:
                            msg_erro = 'Número de telemóvel incorreto'
                            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                    except ValueError:
                        msg_erro = 'Número de telemóvel incorreto'
                        QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
                else:
                    msg_erro = 'Telemóvel vazio'
                    QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
            else:
                msg_erro = 'Email incorreto'
                QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
        else:
            msg_erro = 'É necessário preencher todos os campos'
            QMessageBox.critical(self, 'Erro', msg_erro, QMessageBox.Ok)
            
class TagClienteWindow(QDialog):
    def __init__(self, id_cliente_egoi, tag, parent=None):
        super().__init__(parent)
        # self.setWindowTitle('Atribuir tags')
        self.setWindowFlag(Qt.FramelessWindowHint)
        # self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))
        self.setFixedSize(1000, 500)
        self.setGeometry(500, 200, 0, 0)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.tag = tag
        self.id_cliente_egoi = id_cliente_egoi
        
        self.layout = QVBoxLayout()

        self.attributionGroupBox = QGroupBox('Escolha as tags:')
        self.layout.addWidget(self.attributionGroupBox)

        self.FormularioTagCliente()

        atribuirBtn = QPushButton('Atribuir')
        atribuirBtn.clicked.connect(self.onAtribuirClicked)
        self.layout.addWidget(atribuirBtn)
       
        closeBtn = QPushButton('Fechar')
        closeBtn.clicked.connect(self.close)
        self.layout.addWidget(closeBtn)
       
        self.setLayout(self.layout)

    def FormularioTagCliente(self):
        layout = QGridLayout(self.attributionGroupBox)
        font = QFont()
        font.setPointSize(10)

        abc_tags = sorted(self.tag, key=lambda x: x[1])

        num_colunas = 3
        for index, (id_cliente_egoi, nome_tag) in enumerate(abc_tags):
            row = index // num_colunas
            col = index % num_colunas

            checkbox = QCheckBox(nome_tag)
            checkbox.id_cliente_egoi = id_cliente_egoi
            checkbox.setFont(font)
            layout.addWidget(checkbox, row, col)

            if self.tag_attached(id_cliente_egoi):
                checkbox.setChecked(True)
                checkbox.setDisabled(True)
            
            checkbox.setStyleSheet(
                "QCheckBox::indicator:checked { background-color: green; border: 0.5px solid green; }"
            )

    def tag_attached(self, id_egoi_tag):
        try:
            cursor.execute(
                f"SELECT id_egoi_tag FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                (self.id_cliente_egoi, id_egoi_tag)
            )
            tag_existente = cursor.fetchone()
            return tag_existente is not None
        
        except Exception as e:
            logging.error('Erro', exc_info=True)
            print('Ver py_log.log')
            return False

    def onAtribuirClicked(self):
        selected_checkboxes = [checkbox for checkbox in self.findChildren(QCheckBox) if checkbox.isChecked()]

        if selected_checkboxes:
            id_egoi_tag = [checkbox.id_cliente_egoi for checkbox in selected_checkboxes]
            try:
                for id_tag in id_egoi_tag:
                    cursor.execute(
                        f"SELECT id_egoi_tag FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                        (self.id_cliente_egoi, id_tag)
                    )
                    tag_existente = cursor.fetchone()

                    if not tag_existente:
                        cursor.execute(
                            f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] (id_egoi_cliente, id_egoi_tag) VALUES (?, ?)",
                            (self.id_cliente_egoi, id_tag)
                        )
                        conn.commit()
                        
                        if egoi.AttachTagContacto():
                            QMessageBox.information(self, 'Sucesso', 'Tag(s) atribuída(s) ao cliente!', QMessageBox.Ok)
                        else:
                            QMessageBox.critical(self, 'Erro', 'Erro ao atribuir a(s) tag(s).', QMessageBox.Ok)
                    self.close()
                        
            except Exception as e:
                logging.error('Erro', exc_info=True)
                print('Ver py_log.log')
                
class RemoverTagClienteWindow(QDialog):
    def __init__(self, id_cliente_egoi, tag, parent=None):
        super().__init__(parent)
        # self.setWindowTitle('Remover tags')
        self.setWindowFlag(Qt.FramelessWindowHint)
        # self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))
        self.setFixedSize(1000, 500)
        self.setGeometry(500, 200, 0, 0)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.tag = tag
        self.id_cliente_egoi = id_cliente_egoi
        
        self.layout = QVBoxLayout()

        self.attributionGroupBox = QGroupBox('Escolha as tags:')
        self.layout.addWidget(self.attributionGroupBox)

        self.FormularioTagCliente()

        rmvBtn = QPushButton('Remover')
        rmvBtn.clicked.connect(self.onRemoverClicked)
        self.layout.addWidget(rmvBtn)

        closeBtn = QPushButton('Fechar')
        closeBtn.clicked.connect(self.close)
        self.layout.addWidget(closeBtn)
        
        self.setLayout(self.layout)

    def FormularioTagCliente(self):
        layout = QGridLayout(self.attributionGroupBox)
        font = QFont()
        font.setPointSize(10)

        abc_tags = sorted(self.tag, key=lambda x: x[1])

        num_colunas = 3
        for index, (id_cliente_egoi, nome_tag) in enumerate(abc_tags):
            row = index // num_colunas
            col = index % num_colunas

            checkbox = QCheckBox(nome_tag)
            checkbox.id_cliente_egoi = id_cliente_egoi
            checkbox.setFont(font)
            layout.addWidget(checkbox, row, col)

            if self.tag_detached(id_cliente_egoi):
                checkbox.setChecked(False)
                checkbox.setDisabled(True)
            else:
                checkbox.setChecked(False)
                
            checkbox.setStyleSheet(
                "QCheckBox::indicator:checked { background-color: red; border: 0.5px solid red; }"
            )

    def tag_detached(self, id_egoi_tag):
        try:
            cursor.execute(
                f"SELECT id_egoi_tag FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                (self.id_cliente_egoi, id_egoi_tag)
            )
            tag_existente = cursor.fetchone()
            return tag_existente is None
        
        except Exception as e:
            logging.error('Erro', exc_info=True)
            print('Ver py_log.log')
            return False

    def onRemoverClicked(self):
        selected_checkboxes = [checkbox for checkbox in self.findChildren(QCheckBox) if checkbox.isChecked()]

        if selected_checkboxes:
            try:
                for checkbox in selected_checkboxes:
                    id_egoi_tag = checkbox.id_cliente_egoi
                    cursor.execute(
                        f"SELECT et.tag_id FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] AS ct JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] AS et ON et.id_egoi_tag = ct.id_egoi_tag WHERE id_egoi_cliente = ? AND et.id_egoi_tag = ?",
                        (self.id_cliente_egoi, id_egoi_tag)
                    )
                    tag_id = cursor.fetchone()

                    if tag_id:
                       
                        if egoi.DetachTagContacto(self.id_cliente_egoi, tag_id[0]):
                            QMessageBox.information(self, 'Sucesso', 'Tag(s) removida(s) ao cliente!', QMessageBox.Ok)
                        else:
                            QMessageBox.critical(self, 'Erro', 'Erro ao atribuir a(s) tag(s).', QMessageBox.Ok)
                        
                        cursor.execute(
                            f"DELETE FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                            (self.id_cliente_egoi, id_egoi_tag)
                        )
                        conn.commit()
                self.close()

            except Exception as e:
                logging.error('Erro', exc_info=True)
                print('Ver py_log.log')

def main():
    app = QApplication([])
    app.setStyle('fusion')
    font = QFont('Bahnschrift', 11)
    QApplication.setFont(font)
    window = App()
    window.setFixedSize(635, 480)
    window.setCentralWidget(window.central_widget)
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()