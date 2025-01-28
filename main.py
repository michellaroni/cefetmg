from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QDialog, QMessageBox, QComboBox, QMenuBar, QHBoxLayout, QMenu, QTextEdit
from PySide6.QtGui import QFont, QIcon, QAction
from PySide6.QtCore import Qt
from PySide6.QtCore import QTimer, QTime, QUrl, QLoggingCategory
from PySide6.QtMultimedia import QMediaPlayer, QAudioOutput, QMediaDevices
import sys
import os
import json
from datetime import datetime
from fpdf import FPDF
from PyPDF2 import PdfReader
from PyPDF2 import PdfMerger
from pdf2image import convert_from_path
import uuid  # Para gerar nomes únicos para arquivos temporários

class Formulario(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("CEFET-MG - Atividades Complementares")
        self.setFixedSize(700, 500)
        self.setWindowIcon(QIcon("cefetmg.ico"))

        # Arquivo JSON para armazenar os dados
        self.json_file = "dados_atividades.json"
        self.dados = self.carregar_dados_json()

        # ----------------------

          # Cria a barra de menu
        menu_bar = QMenuBar(self)
        ajuda_menu = QMenu("Ajuda", self)

        sobre_action = QAction("Sobre", self)
        sobre_action.triggered.connect(self.mostrar_sobre)

        ajuda_menu.addAction(sobre_action)
        
        # baixar_action = QAction("Baixar nova versão", self)
        # baixar_action.triggered.connect(self.baixar_instalador)
        # ajuda_menu.addAction(baixar_action)

        menu_bar.addMenu(ajuda_menu)

       
        # Adiciona o layout principal
        layout_principal = QVBoxLayout()
        layout_principal.setMenuBar(menu_bar)

        # Espaçador para empurrar o menu para a direita
        menu_bar.setLayoutDirection(Qt.RightToLeft)


        # Fonte personalizada
        fonte = QFont("Segoe UI", 9)
        self.setStyleSheet("""
            QWidget {
                background-color: #f7f7f7;
                font-family: "Segoe UI", sans-serif;
                font-size: 9pt;
            }
            QLabel {
                font-weight: bold;
                color: #333;
            }
            QLineEdit, QComboBox {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 8px;
                background-color: white;
                margin-bottom: 10px;
            }
            QPushButton {
                background-color: #5a99d4;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: #4a7db5;
            }
            QPushButton:pressed {
                background-color: #3a5e8a;
            }
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 8px;
                background-color: white;
                font-size: 9pt;
                margin-top: 10px;
            }
            QMenuBar {
                background-color: #F0F0F5;
                color: #333;
            }
            QMenuBar::item {
                background: transparent;
                padding: 5px 10px;
            }
            QMenuBar::item:selected {
                background: #007AFF;
                color: white;
            }
            QMenu {
                background-color: #FFFFFF;
                border: 1px solid #C8C8C8;
            }
            QMenu::item {
                padding: 5px 20px;
                background: transparent;
                color: #333;
            }
            QMenu::item:selected {
                background: #007AFF;
                color: white;
            }
        """)

        # Campo Matrícula
        self.label_matricula = QLabel("Matrícula:")
        self.input_matricula = QLineEdit()
        layout_principal.addWidget(self.label_matricula)
        layout_principal.addWidget(self.input_matricula)

        # Lista de atividades
        self.atividades = ["Todos"] + [
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

        # ComboBox para selecionar a atividade
        self.label_atividade = QLabel("Selecione a Atividade:")
        self.combo_atividade = QComboBox()
        self.combo_atividade.addItems(self.atividades)
        self.combo_atividade.currentIndexChanged.connect(self.carregar_dados_filtrados)
        layout_principal.addWidget(self.label_atividade)
        layout_principal.addWidget(self.combo_atividade)

        # Campo Horas Pleiteadas
        self.label_horas = QLabel("Horas Pleiteadas:")
        self.input_horas = QLineEdit()
        layout_principal.addWidget(self.label_horas)
        layout_principal.addWidget(self.input_horas)

        # Botão para selecionar arquivo PDF
        self.botao_selecionar_pdf = QPushButton("Selecionar PDF")
        self.botao_selecionar_pdf.clicked.connect(self.selecionar_pdf)
        layout_principal.addWidget(self.botao_selecionar_pdf)

        # Campo para exibir o caminho do PDF
        self.label_caminho_pdf = QLabel("Caminho do PDF:")
        self.texto_caminho_pdf = QLineEdit()
        self.texto_caminho_pdf.setReadOnly(True)
        layout_principal.addWidget(self.label_caminho_pdf)
        layout_principal.addWidget(self.texto_caminho_pdf)

        # Botão de Calcular
        self.botao_calcular = QPushButton("Calcular")
        self.botao_calcular.clicked.connect(self.calcular_resultado)
        layout_principal.addWidget(self.botao_calcular)

        # Caixa de texto para exibir o resultado
        self.resultado_texto = QTextEdit()
        self.resultado_texto.setReadOnly(True)
        layout_principal.addWidget(self.resultado_texto)

        # Botão de Gerar Relatório
        self.botao_relatorio = QPushButton("Gerar Relatório")
        self.botao_relatorio.clicked.connect(self.gerar_relatorio)
        layout_principal.addWidget(self.botao_relatorio)

        # Configurar layout principal
        self.setLayout(layout_principal)

    def mostrar_sobre(self):
        alerta = Sobre(self)
        alerta.exec()  # Exibe como diálogo modal

    def carregar_dados_json(self):
        """Carrega os dados do arquivo JSON."""
        if os.path.exists(self.json_file):
            with open(self.json_file, "r") as file:
                return json.load(file)
        return {}

    def salvar_dados_json(self):
        """Salva os dados no arquivo JSON."""
        with open(self.json_file, "w") as file:
            json.dump(self.dados, file, indent=4)

    def carregar_dados_filtrados(self):
        """Carrega os dados filtrados com base na atividade selecionada."""
        atividade = self.combo_atividade.currentText()
        if atividade == "Todos":
            texto = ""
            for atividade, itens in self.dados.items():
                texto += f"{atividade}:\n"
                for item in itens:
                    texto += f"  Matrícula: {item['matricula']}\n"
                    texto += f"  Horas Pleiteadas: {item['horas']}\n"
                    texto += f"  PDF: {item['pdf']}\n"
                    texto += f"  Horas Aprovadas: {item['horas_aprovadas']:.2f}\n"
                    texto += f"  Data Inclusão: {item['data_inclusao']}\n\n"
            self.resultado_texto.setPlainText(texto)
        elif atividade in self.dados:
            resultados = self.dados[atividade]
            texto = ""
            for item in resultados:
                texto += f"Matrícula: {item['matricula']}\n"
                texto += f"Horas Pleiteadas: {item['horas']}\n"
                texto += f"PDF: {item['pdf']}\n"
                texto += f"Horas Aprovadas: {item['horas_aprovadas']:.2f}\n"
                texto += f"Data Inclusão: {item['data_inclusao']}\n\n"
            self.resultado_texto.setPlainText(texto)
        else:
            self.resultado_texto.clear()

    def selecionar_pdf(self):
        """Abre um diálogo para selecionar o arquivo PDF."""
        caminho_pdf, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if caminho_pdf:
            self.texto_caminho_pdf.setText(caminho_pdf)

    def calcular_resultado(self):
        """Calcula 30% das horas pleiteadas e salva no JSON."""
        matricula = self.input_matricula.text()
        atividade = self.combo_atividade.currentText()
        horas = self.input_horas.text()
        caminho_pdf = self.texto_caminho_pdf.text()

        if atividade == "Todos":
            self.mostrar_mensagem("Erro", "Selecione uma atividade específica para calcular.")
            return

        if not matricula or not horas or not caminho_pdf:
            self.mostrar_mensagem("Erro", "Todos os campos (Matrícula, Horas Pleiteadas e PDF) são obrigatórios.")
            return

        try:
            horas_pleiteadas = float(horas)
            horas_aprovadas = horas_pleiteadas * 0.30
        except ValueError:
            self.mostrar_mensagem("Erro", "Insira um número válido em 'Horas Pleiteadas'.")
            return

        item = {
            "matricula": matricula,
            "horas": horas_pleiteadas,
            "pdf": caminho_pdf,
            "horas_aprovadas": horas_aprovadas,
            "data_inclusao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        if atividade not in self.dados:
            self.dados[atividade] = []
        self.dados[atividade].append(item)

        self.salvar_dados_json()
        self.carregar_dados_filtrados()
        self.limpar_campos()

    def limpar_campos(self):
        """Limpa os campos de entrada."""
        self.input_horas.clear()
        self.texto_caminho_pdf.clear()

    def mostrar_mensagem(self, titulo, mensagem):
        """Exibe uma mensagem informativa."""
        QMessageBox.information(self, titulo, mensagem)


    def gerar_relatorio_(self):
        """Gera um relatório em PDF com os dados do JSON e imagens convertidas dos PDFs."""
        try:
            # Caminho do Poppler para a conversão
            poppler_path = r"C:\Poppler\Library\bin"  # Certifique-se de que o caminho está correto no seu sistema

            # Inicializa o PDF
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Abre os dados do JSON
            with open("dados_atividades.json", "r") as file:
                dados = json.load(file)

            for atividade, itens in dados.items():
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")

                for item in itens:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas Pleiteadas: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Aprovadas: {item['horas_aprovadas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)

                    # Caminho do PDF associado
                    pdf_path = os.path.normpath(item['pdf'])  # Garante o formato correto do caminho
                    if os.path.exists(pdf_path):
                        try:
                            # Converte o PDF em imagens
                            images = convert_from_path(pdf_path, dpi=100, poppler_path=poppler_path)
                            for image in images:
                                # Cria um nome de arquivo único usando UUID
                                unique_name = str(uuid.uuid4())
                                image_path = f"temp_image_{unique_name}.jpg"
                                image.save(image_path, "JPEG")
                                pdf.image(image_path, x=10, w=150)
                                os.remove(image_path)  # Remove imagem temporária
                        except Exception as e:
                            pdf.cell(0, 10, f"Erro ao processar PDF ({pdf_path}): {e}", ln=True)
                    else:
                        pdf.cell(0, 10, f"Arquivo PDF não encontrado: {pdf_path}", ln=True)

                    pdf.ln(10)  # Adiciona espaço entre os registros

            # Salva o relatório
            pdf.output("relatorio_atividades.pdf")
            print("Relatório gerado com sucesso: relatorio_atividades.pdf")
        except Exception as e:
            print(f"Erro ao gerar relatório: {e}")

    def gerar_relatorio(self):
        """Gera um relatório em PDF e combina com PDFs associados."""
        try:
            # Cria o relatório principal
            relatorio_path = "relatorio_atividades.pdf"
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            with open("dados_atividades.json", "r") as file:
                dados = json.load(file)

            for atividade, itens in dados.items():
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")
                for item in itens:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas Pleiteadas: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Aprovadas: {item['horas_aprovadas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)
                    pdf.ln(10)

            pdf.output(relatorio_path)

            # Combina os PDFs associados
            merger = PdfMerger()
            merger.append(relatorio_path)

            for atividade, itens in dados.items():
                for item in itens:
                    pdf_path = os.path.normpath(item["pdf"])
                    if os.path.exists(pdf_path):
                        merger.append(pdf_path)

            # Salva o PDF final combinado
            merger.write("relatorio_completo.pdf")
            merger.close()

            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar relatório: {e}")

            
    def obter_data_hora_atual(self):
        from datetime import datetime
        return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

class Sobre(QDialog):
    """Classe para a janela 'Sobre'."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sobre")
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)  # Define janela sempre no topo
        self.setModal(True)  # Janela modal, bloqueia app principal

        layout = QVBoxLayout()
        self.versao = "1.0.0"
        info_label = QLabel(f"Atividades Complementares - Extrato do Aluno\nVersão: {self.versao}\nDesenvolvido por: 20223006399 - Michelson Breno Pereira Silva")
        layout.addWidget(info_label)
        
        # Caixa de texto somente leitura para instruções
        instrucoes = """
Cota Zero
(Carlos Drummond de Andrade)

Stop.
A vida parou
ou foi o automóvel?
        """.strip()
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
        instrucoes_textbox.setFixedSize(400, 200)  # Tamanho médio
        
        layout.addWidget(instrucoes_textbox)
        
        self.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    formulario = Formulario()
    formulario.show()

    sys.exit(app.exec())
