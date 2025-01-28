from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QCheckBox, QFileDialog, QHBoxLayout
from PySide6.QtGui import QFont
import sys

class Formulario(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("CEFET-MG / Horas Complementares")
        self.setMinimumSize(800, 600)

        # Layout principal
        layout_principal = QVBoxLayout()

        # Fonte personalizada
        fonte = QFont("Arial", 10)

        # Lista de atividades
        atividades = [
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

        self.campos = []  # Lista para armazenar os widgets das atividades

        for atividade in atividades:
            layout_linha = QHBoxLayout()

            # Checkbox para a atividade
            checkbox = QCheckBox(atividade)
            checkbox.setFont(fonte)
            layout_linha.addWidget(checkbox)

            # TextBox para Horas Pleiteadas
            txt_horas_pleiteadas = QLineEdit()
            txt_horas_pleiteadas.setPlaceholderText("Horas Pleiteadas")
            layout_linha.addWidget(txt_horas_pleiteadas)

            # TextBox para Horas Aprovadas (desabilitado)
            txt_horas_aprovadas = QLineEdit()
            txt_horas_aprovadas.setPlaceholderText("Horas Aprovadas (30%)")
            txt_horas_aprovadas.setEnabled(False)
            layout_linha.addWidget(txt_horas_aprovadas)

            # Botão para selecionar arquivo
            botao_selecionar_arquivo = QPushButton("Selecionar Arquivo")
            botao_selecionar_arquivo.clicked.connect(lambda _, t=atividade: self.selecionar_arquivo(t))
            layout_linha.addWidget(botao_selecionar_arquivo)

            # Label para exibir o caminho do arquivo
            label_caminho_arquivo = QLineEdit()
            label_caminho_arquivo.setReadOnly(True)
            layout_linha.addWidget(label_caminho_arquivo)

            # Adiciona os widgets à lista de campos
            self.campos.append({
                "checkbox": checkbox,
                "txt_horas_pleiteadas": txt_horas_pleiteadas,
                "txt_horas_aprovadas": txt_horas_aprovadas,
                "label_caminho_arquivo": label_caminho_arquivo
            })

            layout_principal.addLayout(layout_linha)

        # Botão para calcular Horas Aprovadas
        botao_calcular = QPushButton("Calcular Horas Aprovadas")
        botao_calcular.clicked.connect(self.calcular_horas_aprovadas)
        layout_principal.addWidget(botao_calcular)

        # Configura o layout principal
        self.setLayout(layout_principal)

    def calcular_horas_aprovadas(self):
        for campo in self.campos:
            if campo["checkbox"].isChecked():
                try:
                    horas_pleiteadas = float(campo["txt_horas_pleiteadas"].text())
                    horas_aprovadas = horas_pleiteadas * 0.3
                    campo["txt_horas_aprovadas"].setText(f"{horas_aprovadas:.2f}")
                except ValueError:
                    campo["txt_horas_aprovadas"].setText("Erro")

    def selecionar_arquivo(self, atividade):
        caminho_arquivo, _ = QFileDialog.getOpenFileName(self, f"Selecione o arquivo para {atividade}", "", "PDF Files (*.pdf);;All Files (*)")
        if caminho_arquivo:
            for campo in self.campos:
                if campo["checkbox"].text() == atividade:
                    campo["label_caminho_arquivo"].setText(caminho_arquivo)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    formulario = Formulario()
    formulario.show()

    sys.exit(app.exec())