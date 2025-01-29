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
        try:
            super().__init__()

            self.setWindowTitle("CEFET-MG - Atividades Complementares")
            self.resize(1000, 200)
            self.setWindowIcon(QIcon("cefetmg.ico"))

            # === Dicionário de percentuais por atividade (exemplo de valores) ===
            self.percentuais_atividade = {
                "01. Produção Científica e Tecnológica": 0.30,
                "02. Apresentação de Trabalhos em Eventos": 0.40,
                "03. Participação em congresso e encontro científico": 0.25,
                "04. Patente/software": 0.50,
                "05. Livro/Capítulo de livro": 0.50,
                "06. Participação na Organização de Eventos": 0.20,
                "07. Participação em Programas de Intercâmbio cultural/estudantil": 0.40,
                "08. Premiação em concurso técnico, científico e artístico": 0.50,
                "09. Visita Técnica": 0.20,
                "10. Representação estudantil em Órgão Colegiado da Instituição": 0.30,
                "11. Gestão de Órgãos de Representação Estudantil": 0.25,
                "12. Curso de Línguas Estrangeiras": 0.25,
                "13. Curso extracurricular na área de concentração do curso": 0.30,
                "14. Curso extracurricular em área diferenciada da área de concentração do curso": 0.20,
                "15. Palestra na área de concentração do curso": 0.30,
                "16. Participação em programas de intercâmbio de línguas estrangeiras": 0.40,
                "17. Programa de Educação Tutorial": 0.50,
                "18. Liga Universitária": 0.20,
                "19. Projetos de Ensino": 0.40,
                "20. Participação em empresa júnior": 0.30,
                "21. Outras Atividades": 0.15,
            }

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

            # QUANDO ALTERAR O TEXTO DA MATRÍCULA, LIMPA O TEXTEDIT E VOLTA O COMBOBOX PARA (Selecione)
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

            # ComboBox para selecionar a atividade
            self.rotulo_atividade = QLabel("Atividade:")
            self.caixa_atividade = QComboBox()
            self.caixa_atividade.addItems(self.lista_atividades)
            self.caixa_atividade.currentIndexChanged.connect(self.carregar_dados_filtrados)
            self.layout_formulario.addWidget(self.rotulo_atividade)
            self.layout_formulario.addWidget(self.caixa_atividade)

            # Campo Horas do Certificado (antes Horas Pleiteadas)
            self.rotulo_horas_certificado = QLabel("Horas do Certificado:")
            self.campo_horas_certificado = QLineEdit()
            self.layout_formulario.addWidget(self.rotulo_horas_certificado)
            self.layout_formulario.addWidget(self.campo_horas_certificado)

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

            # Botão "Incluir Certificado" (antes era "Calcular")
            self.botao_incluir_certificado = QPushButton("Incluir Certificado")
            self.botao_incluir_certificado.clicked.connect(self.incluir_certificado)
            self.layout_formulario.addWidget(self.botao_incluir_certificado)

            # Botão de Gerar Relatório
            self.botao_relatorio = QPushButton("Gerar Relatório")
            self.botao_relatorio.clicked.connect(self.gerar_relatorio)
            self.layout_formulario.addWidget(self.botao_relatorio)

            # Widget para o layout do formulário
            self.widget_formulario = QWidget()
            self.widget_formulario.setLayout(self.layout_formulario)
            self.widget_formulario.setMinimumWidth(200)  # Define um tamanho mínimo para o formulário

            # =============== Inverte a ordem dos widgets no splitter ===============
            self.resultado_texto = QTextEdit()
            self.resultado_texto.setReadOnly(True)

            self.divisor.addWidget(self.resultado_texto)
            self.divisor.addWidget(self.widget_formulario)

            # Ajuste para deixar o lado esquerdo maior
            self.divisor.setStretchFactor(0, 3)  # Texto
            self.divisor.setStretchFactor(1, 1)  # Formulário

            # Define tamanhos iniciais (opcional)
            self.divisor.setSizes([500, 200])

            # Adiciona o splitter ao layout principal
            self.layout_principal.addWidget(self.divisor)

            # Configurar layout principal
            self.setLayout(self.layout_principal)

            # Centraliza a janela após definir o tamanho
            self.center_window()
        except Exception as e:
            print(f"Erro em Formulario.__init__: {e}")

    # =================== Métodos Renomeados para Português ===================
    def exibir_sobre(self):
        """Exibe a janela 'Sobre'."""
        try:
            alerta = Sobre(self)
            alerta.exec()  # Exibe como diálogo modal
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
            matricula_filtro = self.campo_matricula.text().strip()  # Lê a matrícula do campo

            # Se não houver matrícula, pode-se limpar ou exibir mensagem
            if not matricula_filtro:
                self.resultado_texto.clear()
                return

            # 1) Se for "(Selecione)", limpa o texto
            if atividade_selecionada == "(Selecione)":
                self.resultado_texto.clear()
                return

            # 2) Se for "TODAS", mostra TODAS as atividades,
            # mas somente dos itens que tenham a mesma matrícula informada
            if atividade_selecionada == "TODAS":
                texto_exibicao = ""
                for atividade, itens in self.dados_atividades.items():
                    # Filtra apenas os itens da matrícula digitada
                    itens_filtrados = [item for item in itens if item['matricula'] == matricula_filtro]
                    if not itens_filtrados:
                        continue

                    texto_exibicao += f"{atividade}:\n"
                    # SOMA DAS HORAS CONVERTIDAS:
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

            # 3) Caso seja uma atividade específica
            elif atividade_selecionada in self.dados_atividades:
                # Pega apenas os itens daquela atividade
                itens_atividade = self.dados_atividades[atividade_selecionada]
                # Filtra pela matrícula
                itens_filtrados = [item for item in itens_atividade if item['matricula'] == matricula_filtro]

                # Se não houver itens para essa matrícula, limpa o texto
                if not itens_filtrados:
                    self.resultado_texto.clear()
                    return

                texto_exibicao = ""
                # SOMA DAS HORAS CONVERTIDAS:
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
                # Se não encontrar a atividade no dicionário, limpa
                self.resultado_texto.clear()

        except Exception as e:
            print(f"Erro em carregar_dados_filtrados: {e}")

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
        - Aplica o percentual específico da atividade para converter as horas.
        - Salva no JSON.
        """
        try:
            matricula = self.campo_matricula.text()
            atividade = self.caixa_atividade.currentText()
            horas_digitadas = self.campo_horas_certificado.text()
            caminho_pdf = self.campo_caminho_pdf.text()

            # Se for "Todas", considera-se erro crítico
            if atividade == "Todas":
                self.mostrar_mensagem(
                    "Erro",
                    "Selecione uma atividade específica para incluir o certificado.",
                    "critical"
                )
                return

            if not matricula or not horas_digitadas or not caminho_pdf:
                self.mostrar_mensagem(
                    "Erro",
                    "Todos os campos (Matrícula, Horas do Certificado e PDF) são obrigatórios.",
                    "critical"
                )
                return

            try:
                horas_certificado = float(horas_digitadas)
            except ValueError:
                self.mostrar_mensagem(
                    "Erro",
                    "Insira um número válido em 'Horas do Certificado'.",
                    "critical"
                )
                return

            # Obtém o percentual da atividade, caso não exista na tabela, usa 0.3 como padrão
            percentual = self.percentuais_atividade.get(atividade, 0.3)
            horas_convertidas = horas_certificado * percentual

            # Monta o registro a ser salvo
            item = {
                "matricula": matricula,
                "horas": horas_certificado,  # Horas do Certificado
                "pdf": caminho_pdf,
                "horas_convertidas": horas_convertidas,  # Horas Convertidas
                "data_inclusao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
        """
        Exibe uma mensagem ao usuário, podendo ser:
         - 'information' (padrão)
         - 'warning'
         - 'critical'
        """
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
        """Limpa o conteúdo do QTextEdit ao digitar nova matrícula e reseta o combo."""
        self.resultado_texto.clear()
        self.caixa_atividade.setCurrentIndex(0)

    # =================== Relatórios ===================
    def gerar_relatorio(self):
        """
        Gera um relatório em PDF com os dados do JSON,
        convertendo as páginas dos PDFs associados em imagens,
        mas apenas para a matrícula informada no campo.
        O PDF final terá nome no formato: AAAAMMDD_HHMMSS_MATRICULA_XXXXXX.pdf
        """
        try:
            # Matrícula para filtrar
            matricula_filtro = self.campo_matricula.text().strip()
            if not matricula_filtro:
                self.mostrar_mensagem(
                    "Erro",
                    "Informe a matrícula para gerar o relatório.",
                    "critical"
                )
                return

            # Caminho do Poppler para a conversão
            poppler_path = r"C:\Poppler\Library\bin"  # Ajuste de acordo com seu ambiente

            # Gera nome do PDF no formato solicitado
            data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo_relatorio = f"{data_hora}_MATRICULA_{matricula_filtro}.pdf"

            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Abre os dados do JSON
            with open("dados_atividades.json", "r") as file:
                dados_temp = json.load(file)

            # Para cada atividade, filtra somente os itens daquela matrícula
            for atividade, itens in dados_temp.items():
                itens_filtrados = [it for it in itens if it['matricula'] == matricula_filtro]
                if not itens_filtrados:
                    continue

                # Imprime a atividade
                pdf.set_font("Arial", style="B", size=14)
                pdf.cell(0, 10, f"Atividade: {atividade}", ln=True, align="L")

                # SOMA DAS HORAS CONVERTIDAS TAMBÉM AQUI:
                converted_hours_sum = sum(item['horas_convertidas'] for item in itens_filtrados)
                pdf.cell(0, 10, f"(HORAS CONVERTIDAS: {converted_hours_sum:.2f})", ln=True, align="L")

                for item in itens_filtrados:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    pdf.cell(0, 10, f"Horas do Certificado: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)

                    # Caminho do PDF associado
                    pdf_path = os.path.normpath(item['pdf'])
                    pdf.ln(5)  # Espaçamento para as imagens

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
                        except Exception as e2:
                            pdf.cell(0, 10, f"Erro ao processar PDF ({pdf_path}): {e2}", ln=True)
                    else:
                        pdf.cell(0, 10, f"Arquivo PDF não encontrado: {pdf_path}", ln=True)

                    pdf.ln(10)  # Espaço entre registros

            # Salva o relatório
            pdf.output(nome_arquivo_relatorio)
            self.mostrar_mensagem(
                "Sucesso",
                f"Relatório gerado com sucesso:\n{nome_arquivo_relatorio}",
                "information"
            )

        except Exception as e:
            print(f"Erro em gerar_relatorio: {e}")
            self.mostrar_mensagem(
                "Erro",
                f"Erro ao gerar relatório: {e}",
                "critical"
            )

    def gerar_relatorio_(self):
        """Gera um relatório em PDF e combina com PDFs associados (versão alternativa)."""
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
                    pdf.cell(0, 10, f"Horas do Certificado: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f}", ln=True)
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

            # Aqui continua chamando QMessageBox diretamente
            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            print(f"Erro em gerar_relatorio_: {e}")
            QMessageBox.critical(self, "Erro", f"Erro ao gerar relatório: {e}")

    def obter_data_hora_atual(self):
        """Retorna data/hora no formato dd/mm/yyyy HH:MM:SS."""
        try:
            return datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        except Exception as e:
            print(f"Erro em obter_data_hora_atual: {e}")
            return ""


class Sobre(QDialog):
    """Classe para a janela 'Sobre'."""
    def __init__(self, parent=None):
        try:
            super().__init__(parent)
            self.setWindowTitle("Sobre")
            self.setWindowFlag(Qt.WindowStaysOnTopHint, True)  # Define janela sempre no topo
            self.setModal(True)  # Janela modal, bloqueia app principal

            layout = QVBoxLayout()
            self.versao = "1.0.0"
            info_label = QLabel(
                f"Atividades Complementares - Extrato do Aluno\n"
                f"Versão: {self.versao}\n"
                f"Desenvolvido por: 20223006399 - Michelson Breno Pereira Silva"
            )
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
