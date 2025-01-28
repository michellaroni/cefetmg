from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox,
    QPushButton, QComboBox, QTextEdit, QFileDialog
)
from PySide6.QtGui import QFont, QIcon
import sys
import os
import json
from datetime import datetime
from fpdf import FPDF
from pdf2image import convert_from_path
import uuid  # Para gerar nomes únicos para arquivos temporários


class Formulario(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DE LETRA")
        self.setFixedSize(800, 600)
        self.setWindowIcon(QIcon("cefetmg.png"))

        # Diretório para armazenar os arquivos JSON
        self.json_dir = "dados_json"
        os.makedirs(self.json_dir, exist_ok=True)

        self.json_file = None  # Arquivo JSON carregado
        self.dados = {}  # Dados atuais do JSON

        # Layouts principais
        layout_principal = QHBoxLayout()
        layout_esquerda = QVBoxLayout()
        layout_direita = QVBoxLayout()

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
            QLineEdit, QPushButton, QComboBox {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 8px;
                background-color: white;
                margin-bottom: 10px;
            }
            QPushButton {
                background-color: #5a99d4;
                color: white;
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
        """)

        # Campo Matrícula
        self.label_matricula = QLabel("Matrícula:")
        self.input_matricula = QLineEdit()
        self.input_matricula.returnPressed.connect(self.carregar_arquivo_json)
        layout_esquerda.addWidget(self.label_matricula)
        layout_esquerda.addWidget(self.input_matricula)

        # Campo Nome Completo
        self.label_nome = QLabel("Nome Completo:")
        self.input_nome = QLineEdit()
        layout_esquerda.addWidget(self.label_nome)
        layout_esquerda.addWidget(self.input_nome)

        # ComboBox para selecionar a atividade
        self.label_atividade = QLabel("Selecione a Atividade:")
        self.combo_atividade = QComboBox()
        self.combo_atividade.addItems([
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
        ])
        layout_esquerda.addWidget(self.label_atividade)
        layout_esquerda.addWidget(self.combo_atividade)

        # Campo Horas Pleiteadas
        self.label_horas = QLabel("Horas Pleiteadas:")
        self.input_horas = QLineEdit()
        layout_esquerda.addWidget(self.label_horas)
        layout_esquerda.addWidget(self.input_horas)

        # Botão para selecionar arquivo PDF
        self.botao_selecionar_pdf = QPushButton("Selecionar PDF")
        self.botao_selecionar_pdf.clicked.connect(self.selecionar_pdf)
        layout_esquerda.addWidget(self.botao_selecionar_pdf)

        # Campo para exibir o caminho do PDF
        self.texto_caminho_pdf = QLineEdit()
        self.texto_caminho_pdf.setReadOnly(True)
        layout_esquerda.addWidget(self.texto_caminho_pdf)

        # Botão de Adicionar
        self.botao_adicionar = QPushButton("Adicionar")
        self.botao_adicionar.clicked.connect(self.adicionar_dados)
        layout_esquerda.addWidget(self.botao_adicionar)

        # Botão de Gerar Relatório
        self.botao_relatorio = QPushButton("Gerar Relatório")
        self.botao_relatorio.clicked.connect(self.gerar_relatorio)
        layout_esquerda.addWidget(self.botao_relatorio)

        # Campo para exibir os dados do JSON
        self.resultado_texto = QTextEdit()
        self.resultado_texto.setReadOnly(True)
        layout_direita.addWidget(self.resultado_texto)

        # Configurar layouts
        layout_principal.addLayout(layout_esquerda, 1)
        layout_principal.addLayout(layout_direita, 2)
        self.setLayout(layout_principal)

    def carregar_arquivo_json(self):
        """Carrega o arquivo JSON com base na matrícula."""
        matricula = self.input_matricula.text().strip()
        if not matricula:
            self.mostrar_mensagem("Erro", "Por favor, insira a matrícula.")
            return

        self.json_file = os.path.join(self.json_dir, f"{matricula}.json")
        if os.path.exists(self.json_file):
            with open(self.json_file, "r") as file:
                self.dados = json.load(file)
                self.atualizar_texto_resultado()
        else:
            self.dados = {}
            self.resultado_texto.clear()
            self.mostrar_mensagem("Info", "Novo arquivo criado para esta matrícula.")

    def salvar_dados_json(self):
        """Salva os dados no arquivo JSON."""
        if not self.json_file:
            self.mostrar_mensagem("Erro", "Matrícula não foi carregada.")
            return
        with open(self.json_file, "w") as file:
            json.dump(self.dados, file, indent=4)

    def selecionar_pdf(self):
        """Abre um diálogo para selecionar o arquivo PDF."""
        caminho_pdf, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if caminho_pdf:
            self.texto_caminho_pdf.setText(caminho_pdf)

    def adicionar_dados(self):
        """Adiciona novos dados ao JSON."""
        matricula = self.input_matricula.text().strip()
        nome = self.input_nome.text().strip()
        atividade = self.combo_atividade.currentText()
        horas = self.input_horas.text().strip()
        caminho_pdf = self.texto_caminho_pdf.text().strip()

        if not (matricula and nome and horas and caminho_pdf):
            self.mostrar_mensagem("Erro", "Todos os campos devem ser preenchidos.")
            return

        try:
            horas_pleiteadas = float(horas)
        except ValueError:
            self.mostrar_mensagem("Erro", "Horas pleiteadas devem ser um número.")
            return

        item = {
            "nome_completo": nome,
            "atividade": atividade,
            "horas": horas_pleiteadas,
            "pdf": caminho_pdf,
            "data_inclusao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        if "dados" not in self.dados:
            self.dados["dados"] = []
        self.dados["dados"].append(item)

        self.salvar_dados_json()
        self.atualizar_texto_resultado()
        self.limpar_campos()

    def atualizar_texto_resultado(self):
        """Atualiza o texto exibido no resultado."""
        texto = json.dumps(self.dados, indent=4)
        self.resultado_texto.setPlainText(texto)

    def gerar_relatorio(self):
        """Gera um relatório em PDF com os dados do JSON."""
        try:
            if not self.dados.get("dados"):
                self.mostrar_mensagem("Erro", "Nenhum dado disponível para gerar o relatório.")
                return

            # Configuração do relatório PDF
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            for item in self.dados["dados"]:
                pdf.cell(0, 10, f"Nome Completo: {item['nome_completo']}", ln=True)
                pdf.cell(0, 10, f"Atividade: {item['atividade']}", ln=True)
                pdf.cell(0, 10, f"Horas Pleiteadas: {item['horas']}", ln=True)
                pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)
                pdf.cell(0, 10, f"PDF: {item['pdf']}", ln=True)
                pdf.ln(10)

            pdf.output("relatorio_atividades.pdf")
            self.mostrar_mensagem("Sucesso", "Relatório gerado com sucesso.")
        except Exception as e:
            self.mostrar_mensagem("Erro", f"Erro ao gerar relatório: {e}")

    def limpar_campos(self):
        """Limpa os campos de entrada."""
        self.input_nome.clear()
        self.input_horas.clear()
        self.texto_caminho_pdf.clear()

    def mostrar_mensagem(self, titulo, mensagem):
        """Exibe uma mensagem informativa."""
        QMessageBox.information(self, titulo, mensagem)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    formulario = Formulario()
    formulario.show()
    sys.exit(app.exec())
