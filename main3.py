from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QDialog, QMessageBox, QComboBox, QMenuBar, QHBoxLayout, QMenu, QTextEdit, QSplitter, QFileDialog
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
        self.resize(100, 200)
        self.setWindowIcon(QIcon("cefetmg.ico"))

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

        # Adiciona o layout principal
        self.layout_principal = QVBoxLayout()
        self.layout_principal.setMenuBar(self.barra_menu)

        # Espaçador para empurrar o menu para a direita
        self.barra_menu.setLayoutDirection(Qt.RightToLeft)

        # Fonte e estilo personalizados
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

        # Layout horizontal principal com QSplitter
        self.divisor = QSplitter(Qt.Horizontal)

        # Layout vertical para os itens do formulário
        self.layout_formulario = QVBoxLayout()

        # Campo Matrícula
        self.rotulo_matricula = QLabel("Matrícula:")
        self.campo_matricula = QLineEdit()
        self.layout_formulario.addWidget(self.rotulo_matricula)
        self.layout_formulario.addWidget(self.campo_matricula)

        # Lista de atividades
        self.lista_atividades = ["Todos"] + [
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
        self.rotulo_atividade = QLabel("Selecione a Atividade:")
        self.caixa_atividade = QComboBox()
        self.caixa_atividade.addItems(self.lista_atividades)
        self.caixa_atividade.currentIndexChanged.connect(self.carregar_dados_filtrados)
        self.layout_formulario.addWidget(self.rotulo_atividade)
        self.layout_formulario.addWidget(self.caixa_atividade)

        # Campo Horas Pleiteadas
        self.rotulo_horas = QLabel("Horas Pleiteadas:")
        self.campo_horas = QLineEdit()
        self.layout_formulario.addWidget(self.rotulo_horas)
        self.layout_formulario.addWidget(self.campo_horas)

        # Botão para selecionar arquivo PDF
        self.botao_selecionar_pdf = QPushButton("Selecionar PDF")
        self.botao_selecionar_pdf.clicked.connect(self.selecionar_pdf)
        self.layout_formulario.addWidget(self.botao_selecionar_pdf)

        # Campo para exibir o caminho do PDF
        self.rotulo_caminho_pdf = QLabel("Caminho do PDF:")
        self.campo_caminho_pdf = QLineEdit()
        self.campo_caminho_pdf.setReadOnly(True)
        self.layout_formulario.addWidget(self.rotulo_caminho_pdf)
        self.layout_formulario.addWidget(self.campo_caminho_pdf)

        # Botão de Calcular
        self.botao_calcular = QPushButton("Calcular")
        self.botao_calcular.clicked.connect(self.calcular_resultado)
        self.layout_formulario.addWidget(self.botao_calcular)

        # Botão de Gerar Relatório
        self.botao_relatorio = QPushButton("Gerar Relatório")
        self.botao_relatorio.clicked.connect(self.gerar_relatorio)
        self.layout_formulario.addWidget(self.botao_relatorio)

        # Widget para o layout do formulário
        self.widget_formulario = QWidget()
        self.widget_formulario.setLayout(self.layout_formulario)

        # Adiciona o widget do formulário ao splitter (divisor)
        self.divisor.addWidget(self.widget_formulario)

        # Caixa de texto para exibir o resultado (maior no splitter)
        self.resultado_texto = QTextEdit()
        self.resultado_texto.setReadOnly(True)

        # Adiciona a caixa de texto ao splitter
        self.divisor.addWidget(self.resultado_texto)

        # Ajuste para deixar o TextEdit maior que o formulário
        self.divisor.setStretchFactor(0, 1)
        self.divisor.setStretchFactor(1, 3)  # Aumenta o fator de crescimento do texto

        self.divisor.setSizes([100, 500])  # <-- ALTERAÇÃO AQUI

        # Adiciona o splitter ao layout principal
        self.layout_principal.addWidget(self.divisor)

        # Configurar layout principal
        self.setLayout(self.layout_principal)

    # =================== Métodos Renomeados para Português ===================

    def exibir_sobre(self):
        """Exibe a janela 'Sobre'."""
        alerta = Sobre(self)
        alerta.exec()  # Exibe como diálogo modal

    def carregar_arquivo_json(self):
        """Carrega os dados do arquivo JSON."""
        if os.path.exists(self.arquivo_json):
            with open(self.arquivo_json, "r") as file:
                return json.load(file)
        return {}

    def salvar_arquivo_json(self):
        """Salva os dados no arquivo JSON."""
        with open(self.arquivo_json, "w") as file:
            json.dump(self.dados_atividades, file, indent=4)

    # =================== Métodos da Lógica ===================

    def carregar_dados_filtrados(self):
        """Carrega os dados filtrados com base na atividade selecionada."""
        atividade_selecionada = self.caixa_atividade.currentText()

        if atividade_selecionada == "Todos":
            texto_exibicao = ""
            for atividade, itens in self.dados_atividades.items():
                texto_exibicao += f"{atividade}:\n"
                for item in itens:
                    texto_exibicao += f"  Matrícula: {item['matricula']}\n"
                    texto_exibicao += f"  Horas Pleiteadas: {item['horas']}\n"
                    texto_exibicao += f"  PDF: {item['pdf']}\n"
                    texto_exibicao += f"  Horas Aprovadas: {item['horas_aprovadas']:.2f}\n"
                    texto_exibicao += f"  Data Inclusão: {item['data_inclusao']}\n\n"
            self.resultado_texto.setPlainText(texto_exibicao)
        elif atividade_selecionada in self.dados_atividades:
            resultados = self.dados_atividades[atividade_selecionada]
            texto_exibicao = ""
            for item in resultados:
                texto_exibicao += f"Matrícula: {item['matricula']}\n"
                texto_exibicao += f"Horas Pleiteadas: {item['horas']}\n"
                texto_exibicao += f"PDF: {item['pdf']}\n"
                texto_exibicao += f"Horas Aprovadas: {item['horas_aprovadas']:.2f}\n"
                texto_exibicao += f"Data Inclusão: {item['data_inclusao']}\n\n"
            self.resultado_texto.setPlainText(texto_exibicao)
        else:
            self.resultado_texto.clear()

    def selecionar_pdf(self):
        """Abre um diálogo para selecionar o arquivo PDF."""
        caminho_pdf, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if caminho_pdf:
            self.campo_caminho_pdf.setText(caminho_pdf)

    def calcular_resultado(self):
        """Calcula 30% das horas pleiteadas e salva no JSON."""
        matricula = self.campo_matricula.text()
        atividade = self.caixa_atividade.currentText()
        horas_digitadas = self.campo_horas.text()
        caminho_pdf = self.campo_caminho_pdf.text()

        if atividade == "Todos":
            self.mostrar_mensagem("Erro", "Selecione uma atividade específica para calcular.")
            return

        if not matricula or not horas_digitadas or not caminho_pdf:
            self.mostrar_mensagem("Erro", "Todos os campos (Matrícula, Horas Pleiteadas e PDF) são obrigatórios.")
            return

        try:
            horas_pleiteadas = float(horas_digitadas)
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

        if atividade not in self.dados_atividades:
            self.dados_atividades[atividade] = []
        self.dados_atividades[atividade].append(item)

        self.salvar_arquivo_json()
        self.carregar_dados_filtrados()
        self.limpar_campos()

    def limpar_campos(self):
        """Limpa os campos de entrada."""
        self.campo_horas.clear()
        self.campo_caminho_pdf.clear()

    def mostrar_mensagem(self, titulo, mensagem):
        """Exibe uma mensagem informativa."""
        QMessageBox.information(self, titulo, mensagem)

    # =================== Relatórios ===================

    def gerar_relatorio_(self):
        """
        Gera um relatório em PDF com os dados do JSON e
        imagens convertidas dos PDFs (função alternativa).
        """
        try:
            # Caminho do Poppler para a conversão
            poppler_path = r"C:\Poppler\Library\bin"  # Ajuste de acordo com seu ambiente

            # Inicializa o PDF
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Abre os dados do JSON
            with open("dados_atividades.json", "r") as file:
                dados_temp = json.load(file)

            for atividade, itens in dados_temp.items():
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")

                for item in itens:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas Pleiteadas: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Aprovadas: {item['horas_aprovadas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)

                    # Caminho do PDF associado
                    pdf_path = os.path.normpath(item['pdf'])
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

                    pdf.ln(10)  # Espaço entre registros

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
                dados_temp = json.load(file)

            for atividade, itens in dados_temp.items():
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

            for atividade, itens in dados_temp.items():
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
