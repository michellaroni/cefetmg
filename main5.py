from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QDialog,
    QMessageBox,
    QComboBox,
    QMenuBar,
    QHBoxLayout,
    QMenu,
    QTextEdit,
    QSplitter,
    QFileDialog
)
from PySide6.QtGui import QFont, QIcon, QAction, QGuiApplication
from PySide6.QtCore import Qt
import sys
import os
import json
from datetime import datetime
from fpdf import FPDF
from PyPDF2 import PdfMerger
from pdf2image import convert_from_path
import uuid


class Formulario(QWidget):
    def __init__(self):
        try:
            super().__init__()

            self.setWindowTitle("CEFET-MG - Atividades Complementares")
            self.resize(1000, 200)
            self.setWindowIcon(QIcon("cefetmg.ico"))

            # === Dicionário de percentuais por atividade ===
            # Agora baseado no “Máximo aproveitamento relativo à carga horária de AC”
            # (coluna da direita no Anexo III).
            # Obs.: Para o item 17, poderia ser 0.80 (certificado final) ou 0.50 (demais condições).
            # Escolhemos 0.80 como padrão.
            self.percentuais_atividade = {
                "01. Produção Científica e Tecnológica": 0.70,
                "02. Apresentação de Trabalhos em Eventos": 0.70,
                "03. Participação em congresso e encontro científico": 0.50,
                "04. Patente/software": 0.70,
                "05. Livro/Capítulo de livro": 0.70,
                "06. Participação na Organização de Eventos": 0.30,
                "07. Participação em Programas de Intercâmbio cultural/estudantil": 0.60,
                "08. Premiação em concurso técnico, científico e artístico": 0.40,
                "09. Visita Técnica": 0.40,
                "10. Representação estudantil em Órgão Colegiado da Instituição": 0.30,
                "11. Gestão de Órgãos de Representação Estudantil": 0.30,
                "12. Curso de Línguas Estrangeiras": 0.30,
                "13. Curso extracurricular na área de concentração do curso": 0.40,
                "14. Curso extracurricular em área diferenciada da área de concentração do curso": 0.20,
                "15. Palestra na área de concentração do curso": 0.30,
                "16. Participação em programas de intercâmbio de línguas estrangeiras": 0.40,
                "17. Programa de Educação Tutorial": 0.80,
                "18. Liga Universitária": 0.80,
                "19. Projetos de Ensino": 0.30,
                "20. Participação em empresa júnior": 0.30,
                "21. Outras Atividades": 0.30,
            }

            # ==========================================================================================
            # Agora, para cada Atividade, definimos (descrição, K, calcHoras)
            # Em "calcHoras" guardamos o valor fixo da coluna "Cálculo de horas a ser atribuída (horas)"
            # da Resolução. Onde houver "K x 30", "K x 10", etc., guardamos só "30", "10" etc.
            # Se a tabela pedia também "número de dias", "número de meses", etc., aqui usamos 1 como simplificação.
            self.detalhes_atividades = {
                # 01 - Produção Científica e Tecnológica => K x 30 (horas)
                "01. Produção Científica e Tecnológica": [
                    ("Periódicos (K=2, calcHoras=30)", 2.0, 30),
                    ("Congresso Internacional (K=1.5, calcHoras=30)", 1.5, 30),
                    ("Congresso Nacional (K=1, calcHoras=30)", 1.0, 30),
                    ("Outros eventos (K=0.5, calcHoras=30)", 0.5, 30),
                    ("Resumo (K=0.1, calcHoras=30)", 0.1, 30),
                ],
                # 02 - Apresentação de Trabalhos em Eventos => K x 10
                "02. Apresentação de Trabalhos em Eventos": [
                    ("Evento Internacional (K=1.5, calcHoras=10)", 1.5, 10),
                    ("Evento Nacional (K=1.0, calcHoras=10)", 1.0, 10),
                    ("Outros eventos (K=0.5, calcHoras=10)", 0.5, 10),
                ],
                # 03 - Participação em congresso => K x 8 x Número de dias => aqui assumimos 1 dia => 8
                "03. Participação em congresso e encontro científico": [
                    ("Congresso Internacional (K=1.5, calcHoras=8)", 1.5, 8),
                    ("Congresso Nacional (K=1.0, calcHoras=8)", 1.0, 8),
                    ("Outros eventos (K=0.5, calcHoras=8)", 0.5, 8),
                ],
                # 04 - Patente/software => K x 100
                "04. Patente/software": [
                    ("Autor único (K=1, calcHoras=100)", 1.0, 100),
                    ("Co-autor (K=0.5, calcHoras=100)", 0.5, 100),
                ],
                # 05 - Livro/Capítulo de livro => K x 100
                "05. Livro/Capítulo de livro": [
                    ("Livro (K=1.0, calcHoras=100)", 1.0, 100),
                    ("Capítulo (K=0.5, calcHoras=100)", 0.5, 100),
                ],
                # 06 - Participação na Organização de Eventos => K x 15
                "06. Participação na Organização de Eventos": [
                    ("K=1, calcHoras=15", 1.0, 15),
                ],
                # 07 - Participação em Programas de Intercâmbio => K x 10 x número de meses => assumindo 1 => 10
                "07. Participação em Programas de Intercâmbio cultural/estudantil": [
                    ("K=1, calcHoras=10", 1.0, 10),
                ],
                # 08 - Premiação em concurso => K x 30
                "08. Premiação em concurso técnico, científico e artístico": [
                    ("Três primeiras colocações (K=1.0, calcHoras=30)", 1.0, 30),
                    ("Demais colocações (K=0.5, calcHoras=30)", 0.5, 30),
                ],
                # 09 - Visita Técnica => K=1.5 => K x (horas das visitas). Aqui simplificamos p/ 1 => 8, por ex.
                # Mas no original diz "K x (horas das visitas)". Ajustar conforme a sua entrada real.
                "09. Visita Técnica": [
                    ("K=1.5, calcHoras=8 (exemplo)", 1.5, 8),
                ],
                # 10 - Representação estudantil => K x 15 x número de semestres => assumindo 1 => 15
                "10. Representação estudantil em Órgão Colegiado da Instituição": [
                    ("Titular (K=1, calcHoras=15)", 1.0, 15),
                    ("Suplente (K=0.5, calcHoras=15)", 0.5, 15),
                ],
                # 11 - Gestão de Órgãos de Representação => K x 30 x número de anos => assumindo 1 => 30
                "11. Gestão de Órgãos de Representação Estudantil": [
                    ("Presidente (K=1, calcHoras=30)", 1.0, 30),
                    ("Titulares (K=0.5, calcHoras=30)", 0.5, 30),
                    ("Suplentes (K=0.25, calcHoras=30)", 0.25, 30),
                ],
                # 12 - Curso de Línguas Estrangeiras => K x (número de horas). Aqui simplificado p/ "calcHoras=10"
                "12. Curso de Línguas Estrangeiras": [
                    ("K=0.5, calcHoras=10 (exemplo)", 0.5, 10),
                ],
                # 13 - Curso extracurricular na área => K x (número de horas). Simplificado p/ 10
                "13. Curso extracurricular na área de concentração do curso": [
                    ("K=1.0, calcHoras=10 (exemplo)", 1.0, 10),
                ],
                # 14 - Curso extracurricular em área diferenciada => K=0.5 x (horas). Simplificado p/ 10
                "14. Curso extracurricular em área diferenciada da área de concentração do curso": [
                    ("K=0.5, calcHoras=10 (exemplo)", 0.5, 10),
                ],
                # 15 - Palestra na área => K x (horas). Simplificado p/ 10
                "15. Palestra na área de concentração do curso": [
                    ("K=1.0, calcHoras=10 (exemplo)", 1.0, 10),
                ],
                # 16 - Participação em programas de intercâmbio de línguas => K=1 x (horas). Simplificado
                "16. Participação em programas de intercâmbio de línguas estrangeiras": [
                    ("K=1.0, calcHoras=10 (exemplo)", 1.0, 10),
                ],
                # 17 - Programa de Educação Tutorial => K x 120 x número de semestres => assumindo 1 => 120
                "17. Programa de Educação Tutorial": [
                    ("K=1.0, calcHoras=120 (1 semestre)", 1.0, 120),
                ],
                # 18 - Liga Universitária => K=1 x 120 x número de semestres => assumindo 1 => 120
                "18. Liga Universitária": [
                    ("K=1.0, calcHoras=120 (1 semestre)", 1.0, 120),
                ],
                # 19 - Projetos de Ensino => K=1 x (número de horas). Simplificado p/ 10
                "19. Projetos de Ensino": [
                    ("K=1.0, calcHoras=10 (exemplo)", 1.0, 10),
                ],
                # 20 - Participação em empresa júnior => K=1 x 80 x número de semestres => assumindo 1 => 80
                "20. Participação em empresa júnior": [
                    ("K=1.0, calcHoras=80 (1 semestre)", 1.0, 80),
                ],
                # 21 - Outras Atividades => "A definir". Vamos usar K=1, calcHoras=10 p/ exemplo.
                "21. Outras Atividades": [
                    ("A definir (K=1.0, calcHoras=10)", 1.0, 10),
                ],
            }
            # ==========================================================================================

            # Arquivo JSON para armazenar os dados
            self.arquivo_json = "dados_atividades.json"
            self.dados_atividades = self.carregar_arquivo_json()

            # Cria a barra de menu
            self.barra_menu = QMenuBar(self)
            self.menu_ajuda = QMenu("Ajuda", self)
            self.acao_sobre = QAction("Sobre", self)
            self.acao_sobre.triggered.connect(self.exibir_sobre)
            self.menu_ajuda.addAction(self.acao_sobre)
            self.barra_menu.addMenu(self.menu_ajuda)

            self.layout_principal = QVBoxLayout()
            self.layout_principal.setMenuBar(self.barra_menu)
            self.barra_menu.setLayoutDirection(Qt.RightToLeft)

            fonte = QFont("Helvetica Neue", 10)
            self.setStyleSheet(f"""
                QWidget {{
                    background-color: #ECECEC; 
                    font-family: "Helvetica Neue", sans-serif;
                    font-size: 10pt;
                }}
                QLabel {{
                    font-weight: bold;
                    color: #444;
                }}
                QLineEdit, QComboBox {{
                    border: 1px solid #BEBEBE;
                    border-radius: 4px;
                    padding: 6px;
                    background-color: white;
                    margin-bottom: 10px;
                }}
                QPushButton {{
                    background-color: #007AFF;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 12px;
                    font-size: 10pt;
                }}
                QPushButton:hover {{
                    background-color: #005EB8;
                }}
                QPushButton:pressed {{
                    background-color: #004A91;
                }}
                QTextEdit {{
                    border: 1px solid #BEBEBE;
                    border-radius: 4px;
                    padding: 8px;
                    background-color: white;
                    font-size: 10pt;
                    margin-top: 10px;
                }}
                QMenuBar {{
                    background-color: #F0F0F5;
                    color: #333;
                }}
                QMenuBar::item {{
                    background: transparent;
                    padding: 5px 10px;
                }}
                QMenuBar::item:selected {{
                    background: #007AFF;
                    color: white;
                }}
                QMenu {{
                    background-color: #FFFFFF;
                    border: 1px solid #C8C8C8;
                }}
                QMenu::item {{
                    padding: 5px 20px;
                    background: transparent;
                    color: #333;
                }}
                QMenu::item:selected {{
                    background: #007AFF;
                    color: white;
                }}
            """)

            self.divisor = QSplitter(Qt.Horizontal)
            self.layout_formulario = QVBoxLayout()

            # Campo Matrícula
            self.rotulo_matricula = QLabel("Matrícula:")
            self.campo_matricula = QLineEdit()
            self.layout_formulario.addWidget(self.rotulo_matricula)
            self.layout_formulario.addWidget(self.campo_matricula)

            self.campo_matricula.textChanged.connect(self.limpar_resultado_texto)

            # Lista de atividades
            self.lista_atividades = [
                "(Selecione)",
                "TODAS",
                "01. Produção Científica e Tecnológica",
                "02. Apresentação de Trabalhos em Eventos",
                "03. Participação em congresso e encontro científico",
                "04. Patente/software",
                "05. Livro/Capítulo de livro",
                "06. Participação na Organização de Eventos",
                "07. Participação em Programas de Intercâmbio cultural/estudantil",
                "08. Premiação em concurso técnico, científico e artístico",
                "09. Visita Técnica",
                "10. Representação estudantil em Órgão Colegiado da Instituição",
                "11. Gestão de Órgãos de Representação Estudantil",
                "12. Curso de Línguas Estrangeiras",
                "13. Curso extracurricular na área de concentração do curso",
                "14. Curso extracurricular em área diferenciada da área de concentração do curso",
                "15. Palestra na área de concentração do curso",
                "16. Participação em programas de intercâmbio de línguas estrangeiras",
                "17. Programa de Educação Tutorial",
                "18. Liga Universitária",
                "19. Projetos de Ensino",
                "20. Participação em empresa júnior",
                "21. Outras Atividades"
            ]

            self.rotulo_atividade = QLabel("Atividade:")
            self.caixa_atividade = QComboBox()
            self.caixa_atividade.addItems(self.lista_atividades)
            # Chamado para atualizar a exibição do JSON e também para atualizar os detalhes
            self.caixa_atividade.currentIndexChanged.connect(self.carregar_dados_filtrados)
            self.caixa_atividade.currentIndexChanged.connect(self.atualizar_detalhes)
            self.layout_formulario.addWidget(self.rotulo_atividade)
            self.layout_formulario.addWidget(self.caixa_atividade)

            # ComboBox Detalhe da Atividade
            self.rotulo_detalhe = QLabel("Detalhe da Atividade:")
            self.caixa_detalhe = QComboBox()
            self.layout_formulario.addWidget(self.rotulo_detalhe)
            self.layout_formulario.addWidget(self.caixa_detalhe)

            self.rotulo_horas_certificado = QLabel("Horas do Certificado:")
            self.campo_horas_certificado = QLineEdit()
            self.layout_formulario.addWidget(self.rotulo_horas_certificado)
            self.layout_formulario.addWidget(self.campo_horas_certificado)

            self.botao_selecionar_pdf = QPushButton("Selecionar PDF")
            self.botao_selecionar_pdf.clicked.connect(self.selecionar_pdf)
            self.layout_formulario.addWidget(self.botao_selecionar_pdf)

            self.rotulo_caminho_pdf = QLabel("Caminho do PDF:")
            self.campo_caminho_pdf = QLineEdit()
            self.campo_caminho_pdf.setReadOnly(True)
            self.layout_formulario.addWidget(self.rotulo_caminho_pdf)
            self.layout_formulario.addWidget(self.campo_caminho_pdf)

            self.botao_incluir_certificado = QPushButton("Incluir Certificado")
            self.botao_incluir_certificado.clicked.connect(self.incluir_certificado)
            self.layout_formulario.addWidget(self.botao_incluir_certificado)

            self.botao_relatorio = QPushButton("Gerar Relatório")
            self.botao_relatorio.clicked.connect(self.gerar_relatorio)
            self.layout_formulario.addWidget(self.botao_relatorio)

            self.widget_formulario = QWidget()
            self.widget_formulario.setLayout(self.layout_formulario)
            self.widget_formulario.setMinimumWidth(200)

            self.resultado_texto = QTextEdit()
            self.resultado_texto.setReadOnly(True)

            self.divisor.addWidget(self.resultado_texto)
            self.divisor.addWidget(self.widget_formulario)

            self.divisor.setStretchFactor(0, 3)
            self.divisor.setStretchFactor(1, 1)
            self.divisor.setSizes([500, 200])

            self.layout_principal.addWidget(self.divisor)
            self.setLayout(self.layout_principal)

            self.center_window()
            
        except Exception as e:
            print(f"Erro em Formulario.__init__: {e}")

    def center_window(self):
        """Centraliza a janela na tela."""
        qtRect = self.frameGeometry()
        centerPoint = QGuiApplication.primaryScreen().availableGeometry().center()
        qtRect.moveCenter(centerPoint)
        self.move(qtRect.topLeft())

    def exibir_sobre(self):
        """Exibe a janela 'Sobre'."""
        try:
            alerta = Sobre(self)
            alerta.exec()
        except Exception as e:
            print(f"Erro em exibir_sobre: {e}")

    def carregar_arquivo_json(self):
        """Carrega os dados do arquivo JSON."""
        try:
            if os.path.exists(self.arquivo_json):
                with open(self.arquivo_json, "r") as file:
                    return json.load(file)
            return {}
        except Exception as e:
            print(f"Erro em carregar_arquivo_json: {e}")
            return {}

    def salvar_arquivo_json(self):
        """Salva os dados no arquivo JSON."""
        try:
            with open(self.arquivo_json, "w") as file:
                json.dump(self.dados_atividades, file, indent=4)
        except Exception as e:
            print(f"Erro em salvar_arquivo_json: {e}")

    def carregar_dados_filtrados(self):
        """Carrega os dados filtrados com base na atividade e na matrícula informada."""
        try:
            atividade_selecionada = self.caixa_atividade.currentText()
            matricula_filtro = self.campo_matricula.text().strip()

            if not matricula_filtro:
                self.resultado_texto.clear()
                return

            if atividade_selecionada == "(Selecione)":
                self.resultado_texto.clear()
                return

            if atividade_selecionada == "TODAS":
                texto_exibicao = ""
                for atividade, itens in self.dados_atividades.items():
                    itens_filtrados = [item for item in itens if item['matricula'] == matricula_filtro]
                    if not itens_filtrados:
                        continue

                    texto_exibicao += f"{atividade}:\n"
                    converted_hours_sum = sum(item['horas_convertidas'] for item in itens_filtrados)
                    texto_exibicao += f"(HORAS CONVERTIDAS: {converted_hours_sum:.2f})\n"

                    for item in itens_filtrados:
                        texto_exibicao += (
                            f"  Matrícula: {item['matricula']}\n"
                            f"  Horas do Certificado: {item['horas']}\n"
                            f"  PDF: {item['pdf']}\n"
                            f"  Horas Convertidas: {item['horas_convertidas']:.2f}\n"
                            f"  Data Inclusão: {item['data_inclusao']}\n\n"
                        )
                self.resultado_texto.setPlainText(texto_exibicao)

            elif atividade_selecionada in self.dados_atividades:
                itens_atividade = self.dados_atividades[atividade_selecionada]
                itens_filtrados = [item for item in itens_atividade if item['matricula'] == matricula_filtro]
                if not itens_filtrados:
                    self.resultado_texto.clear()
                    return

                texto_exibicao = ""
                converted_hours_sum = sum(item['horas_convertidas'] for item in itens_filtrados)
                texto_exibicao += f"{atividade_selecionada}:\n"
                texto_exibicao += f"(HORAS CONVERTIDAS: {converted_hours_sum:.2f})\n\n"

                for item in itens_filtrados:
                    texto_exibicao += (
                        f"Matrícula: {item['matricula']}\n"
                        f"Horas do Certificado: {item['horas']}\n"
                        f"PDF: {item['pdf']}\n"
                        f"Horas Convertidas: {item['horas_convertidas']:.2f}\n"
                        f"Data Inclusão: {item['data_inclusao']}\n\n"
                    )
                self.resultado_texto.setPlainText(texto_exibicao)

            else:
                self.resultado_texto.clear()

        except Exception as e:
            print(f"Erro em carregar_dados_filtrados: {e}")

    def atualizar_detalhes(self):
        """Atualiza o comboBox de 'Detalhe' baseado na atividade selecionada."""
        atividade_selecionada = self.caixa_atividade.currentText().strip()
        self.caixa_detalhe.clear()

        if atividade_selecionada in ("(Selecione)", "TODAS"):
            return

        # Preenche o combo de detalhe com (descrição, K, calcHoras)
        if atividade_selecionada in self.detalhes_atividades:
            opcoes = self.detalhes_atividades[atividade_selecionada]
            for label, k_valor, calcHoras_valor in opcoes:
                # Guardamos (k_valor, calcHoras_valor) como "UserData"
                # e usamos 'label' para exibição
                self.caixa_detalhe.addItem(label, (k_valor, calcHoras_valor))

    def selecionar_pdf(self):
        """Abre um diálogo para selecionar o arquivo PDF."""
        try:
            caminho_pdf, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
            if caminho_pdf:
                self.campo_caminho_pdf.setText(caminho_pdf)
        except Exception as e:
            print(f"Erro em selecionar_pdf: {e}")

    def incluir_certificado(self):
        """
        Inclui um novo certificado no sistema:
        - Verifica a atividade selecionada.
        - Aplica (K * calcHoras) de acordo com o detalhe escolhido.
        - Salva no JSON.
        """
        try:
            matricula = self.campo_matricula.text()
            atividade = self.caixa_atividade.currentText()
            horas_digitadas = self.campo_horas_certificado.text()
            caminho_pdf = self.campo_caminho_pdf.text()

            if atividade == "Todas":
                self.mostrar_mensagem("Erro", "Selecione uma atividade específica para incluir o certificado.", "critical")
                return

            if not matricula or not horas_digitadas or not caminho_pdf:
                self.mostrar_mensagem("Erro", "Todos os campos (Matrícula, Horas, PDF) são obrigatórios.", "critical")
                return

            try:
                # O campo "Horas do Certificado" pode ou não influenciar no cálculo,
                # dependendo da sua necessidade. Se não quiser usá-lo, pode ignorar.
                horas_certificado = float(horas_digitadas)
            except ValueError:
                self.mostrar_mensagem("Erro", "Insira um número válido em 'Horas do Certificado'.", "critical")
                return

            # Pegamos do combo de detalhe:
            indice_detalhe = self.caixa_detalhe.currentIndex()
            if indice_detalhe < 0:
                self.mostrar_mensagem("Erro", "Selecione o detalhe da atividade!", "critical")
                return

            # Recupera (k_valor, calcHoras_valor):
            data_tuple = self.caixa_detalhe.itemData(indice_detalhe)
            if not data_tuple:
                self.mostrar_mensagem("Erro", "Detalhe de atividade inválido!", "critical")
                return

            k_valor, calcHoras_valor = data_tuple

            # =============== Cálculo: K * calcHoras ===============
            horas_convertidas = k_valor * calcHoras_valor
            # Se você quiser também multiplicar pelas horas_digitadas, faça algo como:
            # horas_convertidas = (k_valor * calcHoras_valor) * horas_certificado
            # Mas aqui estamos só fazendo K * calcHoras, conforme pedido.

            item = {
                "matricula": matricula,
                "horas": horas_certificado,  # Armazenamos só para histórico, se quiser
                "pdf": caminho_pdf,
                "horas_convertidas": horas_convertidas,
                "data_inclusao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "detalhe_escolhido": self.caixa_detalhe.currentText(),
            }

            if atividade not in self.dados_atividades:
                self.dados_atividades[atividade] = []
            self.dados_atividades[atividade].append(item)

            self.salvar_arquivo_json()
            self.carregar_dados_filtrados()
            self.limpar_campos()
        except Exception as e:
            print(f"Erro em incluir_certificado: {e}")

    def limpar_campos(self):
        """Limpa os campos de entrada."""
        try:
            self.campo_horas_certificado.clear()
            self.campo_caminho_pdf.clear()
        except Exception as e:
            print(f"Erro em limpar_campos: {e}")

    def mostrar_mensagem(self, titulo, mensagem, tipo="information"):
        """Exibe uma mensagem ao usuário."""
        try:
            if tipo == "information":
                QMessageBox.information(self, titulo, mensagem)
            elif tipo == "warning":
                QMessageBox.warning(self, titulo, mensagem)
            elif tipo == "critical":
                QMessageBox.critical(self, titulo, mensagem)
            else:
                QMessageBox.information(self, titulo, mensagem)
        except Exception as e:
            print(f"Erro em mostrar_mensagem: {e}")

    def limpar_resultado_texto(self):
        """
        Limpa o conteúdo do QTextEdit ao digitar nova matrícula
        e reseta os combos para o estado inicial.
        """
        self.resultado_texto.clear()
        self.caixa_atividade.setCurrentIndex(0)
        self.caixa_detalhe.clear()

    def gerar_relatorio(self):
        """
        Gera um relatório em PDF com os dados do JSON,
        convertendo as páginas dos PDFs associados em imagens,
        mas apenas para a matrícula informada no campo.
        """
        try:
            matricula_filtro = self.campo_matricula.text().strip()
            if not matricula_filtro:
                self.mostrar_mensagem("Erro", "Informe a matrícula para gerar o relatório.", "critical")
                return

            poppler_path = r"C:\Poppler\Library\bin"
            data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo_relatorio = f"{data_hora}_MATRICULA_{matricula_filtro}.pdf"

            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            with open("dados_atividades.json", "r") as file:
                dados_temp = json.load(file)

            for atividade, itens in dados_temp.items():
                itens_filtrados = [it for it in itens if it['matricula'] == matricula_filtro]
                if not itens_filtrados:
                    continue

                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")
                converted_hours_sum = sum(item['horas_convertidas'] for item in itens_filtrados)
                pdf.cell(0, 10, f"(HORAS CONVERTIDAS: {converted_hours_sum:.2f})", ln=True, align="L")

                for item in itens_filtrados:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas do Certificado: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f}", ln=True)
                    # Se quiser mostrar o detalhe
                    if 'detalhe_escolhido' in item:
                        pdf.cell(0, 10, f"Detalhe: {item['detalhe_escolhido']}", ln=True)

                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)

                    pdf_path = os.path.normpath(item['pdf'])
                    pdf.ln(5)
                    if os.path.exists(pdf_path):
                        try:
                            images = convert_from_path(pdf_path, dpi=100, poppler_path=poppler_path)
                            for image in images:
                                unique_name = str(uuid.uuid4())
                                image_path = f"temp_image_{unique_name}.jpg"
                                image.save(image_path, "JPEG")
                                pdf.image(image_path, x=10, w=150)
                                os.remove(image_path)
                        except Exception as e2:
                            pdf.cell(0, 10, f"Erro ao processar PDF ({pdf_path}): {e2}", ln=True)
                    else:
                        pdf.cell(0, 10, f"Arquivo PDF não encontrado: {pdf_path}", ln=True)

                    pdf.ln(10)

            pdf.output(nome_arquivo_relatorio)
            self.mostrar_mensagem("Sucesso", f"Relatório gerado com sucesso:\n{nome_arquivo_relatorio}", "information")
        except Exception as e:
            print(f"Erro em gerar_relatorio: {e}")
            self.mostrar_mensagem("Erro", f"Erro ao gerar relatório: {e}", "critical")

    def gerar_relatorio_(self):
        """Exemplo de relatório unindo PDFs (versão alternativa)."""
        try:
            relatorio_path = "relatorio_atividades.pdf"
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            with open("dados_atividades.json", "r") as file:
                dados_temp = json.load(file)

            for atividade, itens in dados_temp.items():
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")
                for item in itens:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas do Certificado: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)
                    pdf.ln(10)

            pdf.output(relatorio_path)

            merger = PdfMerger()
            merger.append(relatorio_path)

            for atividade, itens in dados_temp.items():
                for item in itens:
                    pdf_path = os.path.normpath(item["pdf"])
                    if os.path.exists(pdf_path):
                        merger.append(pdf_path)

            merger.write("relatorio_completo.pdf")
            merger.close()

            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            print(f"Erro em gerar_relatorio_: {e}")
            QMessageBox.critical(self, "Erro", f"Erro ao gerar relatório: {e}")


class Sobre(QDialog):
    """Classe para a janela 'Sobre'."""
    def __init__(self, parent=None):
        try:
            super().__init__(parent)
            self.setWindowTitle("Sobre")
            self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
            self.setModal(True)

            layout = QVBoxLayout()
            self.versao = "1.0.0"
            info_label = QLabel(
                f"Atividades Complementares - Extrato do Aluno\n"
                f"Versão: {self.versao}\n"
                f"Desenvolvido por: 20223006399 - Michelson Breno Pereira Silva"
            )
            layout.addWidget(info_label)
            
            instrucoes = """
Quadrilha
(Carlos Drummond de Andrade)

João amava Teresa que amava Raimundo
que amava Maria que amava Joaquim que amava Lili,
que não amava ninguém.
João foi para os Estados Unidos, Teresa para o convento,
Raimundo morreu de desastre, Maria ficou para tia,
Joaquim suicidou-se e Lili casou com J. Pinto Fernandes
que não tinha entrado na história.
""".strip()
            instrucoes_textbox = QTextEdit()
            instrucoes_textbox.setReadOnly(True)
            instrucoes_textbox.setPlainText(instrucoes)
            instrucoes_textbox.setFixedSize(400, 200)
            
            layout.addWidget(instrucoes_textbox)
            self.setLayout(layout)
        except Exception as e:
            print(f"Erro em Sobre.__init__: {e}")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        formulario = Formulario()
        formulario.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Erro na execução principal: {e}")
