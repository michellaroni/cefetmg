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
            # Agora baseado no “Máximo aproveitamento relativo à carga horária de AC” (coluna da direita no Anexo III).
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
            # Em cada item do dicionário self.detalhes_atividades, adicionamos comentários baseados no ANEXO III.
            # A estrutura: (descrição, k_valor, calcHoras_valor, label_unidade).
            # Se "label_unidade" for None, significa "K x calcHoras_valor" diretamente.
            # Se "label_unidade" for uma string, faremos "K x calcHoras_valor x valor_digitado".
            # ==========================================================================================
            self.detalhes_atividades = {
                # ======================================================================================
                # ANEXO III - Item 1
                # Atividade: Produção Científica e Tecnológica
                # Definição/Caracterização: Publicação de trabalho de caráter técnico e/ou científico na
                #                           área de conhecimento do curso.
                # Forma de comprovação: Periódico (doc. expedido pela org. do evento ou cópia do trabalho),
                #                      Evento (doc. expedido pela org. do evento ou cópia do trabalho nos anais).
                # Fatores (K): Periódicos K=2; Congresso Internacional K=1,5; Congresso Nacional K=1; Outros K=0,5; Resumo K=0,1
                # Cálculo (horas): K x 30
                # Cálculo (horas/aula): K x 36
                # Máximo aproveitamento relativo à AC: 70 %
                # ======================================================================================
                "01. Produção Científica e Tecnológica": [
                    ("Periódicos (K=2, calcHoras=30)",                 2.0,  30,   None),
                    ("Congresso Internacional (K=1.5, calcHoras=30)",  1.5,  30,   None),
                    ("Congresso Nacional (K=1, calcHoras=30)",         1.0,  30,   None),
                    ("Outros eventos (K=0.5, calcHoras=30)",           0.5,  30,   None),
                    ("Resumo (K=0.1, calcHoras=30)",                   0.1,  30,   None),
                ],

                # ======================================================================================
                # ANEXO III - Item 2
                # Atividade: Apresentação de Trabalhos em Eventos
                # Definição/Caracterização: Apresentação de trabalhos em eventos na área do curso.
                # Forma de comprovação: Documento expedido pela organização do evento.
                # Fatores (K): Evento Internacional K=1,5; Evento Nacional K=1; Outros K=0,5
                # Cálculo (horas): K x 10
                # Cálculo (horas/aula): K x 12
                # Máximo aproveitamento: 70 %
                # ======================================================================================
                "02. Apresentação de Trabalhos em Eventos": [
                    ("Evento Internacional (K=1.5, calcHoras=10)", 1.5, 10, None),
                    ("Evento Nacional (K=1.0, calcHoras=10)",      1.0, 10, None),
                    ("Outros eventos (K=0.5, calcHoras=10)",       0.5, 10, None),
                ],

                # ======================================================================================
                # ANEXO III - Item 3
                # Atividade: Participação em congresso e encontro científico
                # Definição/Caracterização: Participação em congresso e encontro na área do curso.
                # Forma de comprovação: Documento expedido pela organização do evento.
                # Fatores (K): Congresso Internacional K=1,5; Congresso Nacional K=1; Outros K=0,5
                # Cálculo (horas): K x 8 x número de dias
                # Cálculo (horas/aula): K x 9,6 x número de dias
                # Máximo aproveitamento: 50 %
                # ======================================================================================
                "03. Participação em congresso e encontro científico": [
                    ("Congresso Internacional (K=1.5, calcHoras=8, DIAS)",  1.5, 8, "DIAS"),
                    ("Congresso Nacional (K=1.0, calcHoras=8, DIAS)",       1.0, 8, "DIAS"),
                    ("Outros eventos (K=0.5, calcHoras=8, DIAS)",           0.5, 8, "DIAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 4
                # Atividade: Patente/software
                # Definição/Caracterização: Patente ou desenvolvimento de software registrado com inovação
                #                           comprovada na área do curso.
                # Forma de comprovação: Documento do órgão certificador de pedido de depósito realizado
                #                      pelo CEFET-MG ou documento final.
                # Fatores (K): Autor único K=1; Co-autor K=0,5
                # Cálculo (horas): K x 100
                # Cálculo (horas/aula): K x 120
                # Máximo aproveitamento: 70 %
                # ======================================================================================
                "04. Patente/software": [
                    ("Autor único (K=1, calcHoras=100)", 1.0, 100, None),
                    ("Co-autor (K=0.5, calcHoras=100)",  0.5, 100, None),
                ],

                # ======================================================================================
                # ANEXO III - Item 5
                # Atividade: Livro/Capítulo de livro
                # Definição/Caracterização: Publicação de livro/capítulo na área do curso.
                # Forma de comprovação: Documento expedido pela editora.
                # Fatores (K): Livro K=1; Capítulo K=0,5
                # Cálculo (horas): K x 100
                # Cálculo (horas/aula): K x 120
                # Máximo aproveitamento: 70 %
                # ======================================================================================
                "05. Livro/Capítulo de livro": [
                    ("Livro (K=1.0, calcHoras=100)",    1.0, 100, None),
                    ("Capítulo (K=0.5, calcHoras=100)" ,0.5, 100, None),
                ],

                # ======================================================================================
                # ANEXO III - Item 6
                # Atividade: Participação na Organização de Eventos
                # Definição/Caracterização: Participação do(a) discente no apoio à organização do evento.
                # Forma de comprovação: Documento expedido pela Entidade responsável pelo evento.
                # Fator (K): 1
                # Cálculo (horas): K x 15
                # Cálculo (horas/aula): K x 18
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "06. Participação na Organização de Eventos": [
                    ("K=1, calcHoras=15", 1.0, 15, None),
                ],

                # ======================================================================================
                # ANEXO III - Item 7
                # Atividade: Participação em Programas de Intercâmbio cultural/estudantil
                # Definição/Caracterização: Programas do CEFET-MG com outras instituições nacionais ou internacionais.
                # Forma de comprovação: Documento expedido pela Secretaria de Relações Internacionais ou entidade.
                # Fator (K): 1
                # Cálculo (horas): K x 10 x número de meses
                # Cálculo (horas/aula): K x 12
                # Máximo aproveitamento: 60 %
                # ======================================================================================
                "07. Participação em Programas de Intercâmbio cultural/estudantil": [
                    ("K=1, calcHoras=10, MESES", 1.0, 10, "MESES"),
                ],

                # ======================================================================================
                # ANEXO III - Item 8
                # Atividade: Premiação em concurso técnico, científico e artístico
                # Definição/Caracterização: Premiação do(a) discente em concurso com trabalhos de caráter
                #                           técnico, científico e artístico na área do curso.
                # Forma de comprovação: Documento expedido pela entidade promotora do concurso.
                # Fatores (K): Três primeiras colocações = K=1, Demais colocações = K=0,5
                # Cálculo (horas): K x 30
                # Cálculo (horas/aula): K x 36
                # Máximo aproveitamento: 40 %
                # ======================================================================================
                "08. Premiação em concurso técnico, científico e artístico": [
                    ("Três primeiras colocações (K=1.0, calcHoras=30)", 1.0, 30, None),
                    ("Demais colocações (K=0.5, calcHoras=30)",        0.5, 30, None),
                ],

                # ======================================================================================
                # ANEXO III - Item 9
                # Atividade: Visita Técnica
                # Definição/Caracterização: Visita realizada em empresas e institutos de P&D na área do curso.
                # Forma de comprovação: Documento expedido pela empresa/instituição ou professor responsável.
                # Fator (K): 1,5
                # Cálculo (horas): K x (horas das visitas)
                # Cálculo (horas/aula): K x (horas das visitas) x 1,2
                # Máximo aproveitamento: 40 %
                # ======================================================================================
                "09. Visita Técnica": [
                    ("K=1.5, HORAS", 1.5, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 10
                # Atividade: Representação estudantil em Órgão Colegiado da Instituição
                # Definição/Caracterização: Participação do(a) discente em Órgão Colegiado (titular ou suplente).
                # Forma de comprovação: Documento expedido pelo presidente do Órgão Colegiado.
                # Fatores (K): Titular K=1; Suplente K=0,5
                # Cálculo (horas): K x 15 x número de semestres
                # Cálculo (horas/aula): K x 18
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "10. Representação estudantil em Órgão Colegiado da Instituição": [
                    ("Titular (K=1, calcHoras=15, SEMESTRES)",     1.0, 15, "SEMESTRES"),
                    ("Suplente (K=0.5, calcHoras=15, SEMESTRES)",  0.5, 15, "SEMESTRES"),
                ],

                # ======================================================================================
                # ANEXO III - Item 11
                # Atividade: Gestão de Órgãos de Representação Estudantil
                # Definição/Caracterização: Participação na gestão de DA ou DCE por período de 1 ano.
                # Forma de comprovação: Diretoria de Graduação (Presidente K=1, Titulares K=0,5, Suplentes K=0,25).
                # Cálculo (horas): K x 30 x número de anos
                # Cálculo (horas/aula): K x 36
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "11. Gestão de Órgãos de Representação Estudantil": [
                    ("Presidente (K=1, calcHoras=30, ANOS)",     1.0, 30, "ANOS"),
                    ("Titulares (K=0.5, calcHoras=30, ANOS)",    0.5, 30, "ANOS"),
                    ("Suplentes (K=0.25, calcHoras=30, ANOS)",  0.25, 30, "ANOS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 12
                # Atividade: Curso de Línguas Estrangeiras
                # Definição/Caracterização: Estudo de língua estrangeira, com aprovação, oferecido por
                #                           instituição regulamentada ou pelo CEFET-MG.
                # Forma de comprovação: Certificado emitido pela instituição de ensino regulamentada
                #                      ou pela Fundação CEFET-MINAS.
                # Fator (K): 0,5
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "12. Curso de Línguas Estrangeiras": [
                    ("K=0.5, HORAS", 0.5, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 13
                # Atividade: Curso extracurricular na área de concentração do curso
                # Definição/Caracterização: Curso oferecido pelo CEFET-MG ou outra instituição/empresa,
                #                           envolvendo conteúdos não contemplados na matriz.
                # Forma de comprovação: Certificado emitido pelo órgão responsável.
                # Fator (K): 1
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 40 %
                # ======================================================================================
                "13. Curso extracurricular na área de concentração do curso": [
                    ("K=1.0, HORAS", 1.0, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 14
                # Atividade: Curso extracurricular em área diferenciada da área de concentração do curso
                # Definição/Caracterização: Curso oferecido pelo CEFET-MG ou outra instituição/empresa.
                # Forma de comprovação: Certificado emitido pelo órgão responsável.
                # Fator (K): 0,5
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 20 %
                # ======================================================================================
                "14. Curso extracurricular em área diferenciada da área de concentração do curso": [
                    ("K=0.5, HORAS", 0.5, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 15
                # Atividade: Palestra na área de concentração do curso
                # Definição/Caracterização: Assistir palestra com discussão de temas de interesse do curso.
                # Forma de comprovação: Documento expedido pela organização do evento.
                # Fator (K): 1
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "15. Palestra na área de concentração do curso": [
                    ("K=1.0, HORAS", 1.0, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 16
                # Atividade: Participação em programas de intercâmbio de línguas estrangeiras
                # Definição/Caracterização: Programas internacionais de língua estrangeira.
                # Forma de comprovação: Documento expedido pela entidade responsável.
                # Fator (K): 1
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 40 %
                # ======================================================================================
                "16. Participação em programas de intercâmbio de línguas estrangeiras": [
                    ("K=1.0, HORAS", 1.0, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 17
                # Atividade: Programa de Educação Tutorial
                # Definição/Caracterização: Participação em PET.
                # Forma de comprovação: Declaração do tutor por semestre ou certificado final (2 anos).
                # Fator (K): 1
                # Cálculo (horas): K x 120 x número de semestres
                # Cálculo (horas/aula): K x 144
                # Máximo aproveitamento: 80% (certificado final) e 50% (demais)
                # ======================================================================================
                "17. Programa de Educação Tutorial": [
                    ("K=1.0, calcHoras=120, SEMESTRES", 1.0, 120, "SEMESTRES"),
                ],

                # ======================================================================================
                # ANEXO III - Item 18
                # Atividade: Liga Universitária
                # Definição/Caracterização: Participação em Liga Universitária.
                # Forma de comprovação: Declaração da Instituição ou setor responsável.
                # Fator (K): 1
                # Cálculo (horas): K x 120 x número de semestres
                # Cálculo (horas/aula): K x 144
                # Máximo aproveitamento: 80 %
                # ======================================================================================
                "18. Liga Universitária": [
                    ("K=1.0, calcHoras=120, SEMESTRES", 1.0, 120, "SEMESTRES"),
                ],

                # ======================================================================================
                # ANEXO III - Item 19
                # Atividade: Projetos de Ensino
                # Definição/Caracterização: Participação em Projetos de Ensino.
                # Forma de comprovação: Declaração da Instituição ou setor responsável.
                # Fator (K): 1
                # Cálculo (horas): K x número de horas
                # Cálculo (horas/aula): K x 1,2 x número de horas
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "19. Projetos de Ensino": [
                    ("K=1.0, HORAS", 1.0, None, "HORAS"),
                ],

                # ======================================================================================
                # ANEXO III - Item 20
                # Atividade: Participação em empresa júnior
                # Definição/Caracterização: Participação em empresa júnior.
                # Forma de comprovação: Declaração do Professor Supervisor.
                # Fator (K): 1
                # Cálculo (horas): K x 80 x número de semestres
                # Cálculo (horas/aula): K x 96 x número de semestres
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "20. Participação em empresa júnior": [
                    ("K=1.0, calcHoras=80, SEMESTRES", 1.0, 80, "SEMESTRES"),
                ],

                # ======================================================================================
                # ANEXO III - Item 21
                # Atividade: Outras Atividades
                # Definição/Caracterização: A definir
                # Forma de comprovação: A definir
                # Fator (K): A definir
                # Cálculo (horas): A definir
                # Cálculo (horas/aula): A definir
                # Máximo aproveitamento: 30 %
                # ======================================================================================
                "21. Outras Atividades": [
                    ("A definir", 0, 0, None),
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
            self.caixa_atividade.currentIndexChanged.connect(self.carregar_dados_filtrados)
            self.caixa_atividade.currentIndexChanged.connect(self.atualizar_detalhes)
            self.layout_formulario.addWidget(self.rotulo_atividade)
            self.layout_formulario.addWidget(self.caixa_atividade)

            # ComboBox Detalhe da Atividade
            self.rotulo_detalhe = QLabel("Detalhe da Atividade:")
            self.caixa_detalhe = QComboBox()
            self.layout_formulario.addWidget(self.rotulo_detalhe)
            self.layout_formulario.addWidget(self.caixa_detalhe)

            # Substituir rótulo "Horas do Certificado" por "Unidade de Cálculo"
            self.rotulo_horas_certificado = QLabel("Unidade de Cálculo:")
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

            self.centralizar_janela()
            
        except Exception as e:
            print(f"Erro em Formulario.__init__: {e}")

    def centralizar_janela(self):
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
            atividade_selecionada = self.caixa_atividade.currentText().strip()
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
                    
                    # === Adicionando o cálculo do máximo, considerando self.percentuais_atividade e 240 horas ===
                    if atividade in self.percentuais_atividade:
                        maximo_horas_atividade = self.percentuais_atividade[atividade] * 240
                        texto_exibicao += f"(MÁXIMO DE HORAS: {maximo_horas_atividade:.2f})\n"

                    for item in itens_filtrados:
                        texto_exibicao += (
                            f"  Matrícula: {item['matricula']}\n"
                            #f"  Unidade de Cálculo: {item['unidade']}\n"
                            f"  PDF: {item['pdf']}\n"
                            f"  Horas Convertidas: {item['horas_convertidas']:.2f} (Cálculo: {item.get('detalhes_calculo', '')})\n"
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
                texto_exibicao += f"(HORAS CONVERTIDAS: {converted_hours_sum:.2f})\n"
                
                # === Adicionando o cálculo do máximo também para a atividade selecionada ===
                if atividade_selecionada in self.percentuais_atividade:
                    maximo_horas_atividade = self.percentuais_atividade[atividade_selecionada] * 240
                    texto_exibicao += f"(MÁXIMO DE HORAS: {maximo_horas_atividade:.2f})\n\n"

                for item in itens_filtrados:
                    texto_exibicao += (
                        f"Matrícula: {item['matricula']}\n"
                        #f"Unidade de Cálculo: {item['unidade']}\n"
                        f"PDF: {item['pdf']}\n"
                        f"Horas Convertidas: {item['horas_convertidas']:.2f} (Cálculo: {item.get('detalhes_calculo', '')})\n"
                        f"Data Inclusão: {item['data_inclusao']}\n\n"
                    )
                self.resultado_texto.setPlainText(texto_exibicao)

            else:
                self.resultado_texto.clear()

        except Exception as e:
            print(f"Erro em carregar_dados_filtrados: {e}")

    def atualizar_detalhes(self):
        """
        Atualiza o comboBox de 'Detalhe' baseado na atividade selecionada.
        Se label_unidade não for None, mudamos self.rotulo_horas_certificado para exibir esse nome.
        Senão, deixamos "Unidade de Cálculo:".
        """
        atividade_selecionada = self.caixa_atividade.currentText().strip()
        self.caixa_detalhe.clear()

        # Reseta label para o padrão
        self.rotulo_horas_certificado.setText("Unidade de Cálculo:")

        if atividade_selecionada in ("(Selecione)", "TODAS"):
            return

        if atividade_selecionada in self.detalhes_atividades:
            opcoes = self.detalhes_atividades[atividade_selecionada]
            for (descricao, k_valor, calcHoras_valor, label_unidade) in opcoes:
                self.caixa_detalhe.addItem(descricao, (k_valor, calcHoras_valor, label_unidade))

    def selecionar_pdf(self):
        """Abre um diálogo para selecionar o arquivo PDF."""
        try:
            caminho_pdf_selecionado, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
            if caminho_pdf_selecionado:
                self.campo_caminho_pdf.setText(caminho_pdf_selecionado)
        except Exception as e:
            print(f"Erro em selecionar_pdf: {e}")

    def incluir_certificado(self):
        """
        Inclui um novo certificado:
        - Aplica K, calcHoras e/ou multiplica pelo valor que o usuário digitou, dependendo do label_unidade.
        - Registra detalhadamente como foi o cálculo em 'detalhes_calculo'.
        """
        try:
            matricula = self.campo_matricula.text()
            atividade = self.caixa_atividade.currentText()
            detalhe_escolhido = self.caixa_detalhe.currentText()
            unidade_calculo_str = self.campo_horas_certificado.text()  # Renome simbólico; interna.
            caminho_pdf = self.campo_caminho_pdf.text()

            if atividade == "TODAS":
                self.mostrar_mensagem("Erro", "Selecione uma atividade específica para incluir o certificado.", "critical")
                return
            
            indice_detalhe = self.caixa_detalhe.currentIndex()
            (k_valor, calcHoras_valor, label_unidade) = self.caixa_detalhe.itemData(indice_detalhe)

            if not matricula:
                self.mostrar_mensagem(
                    "Erro",
                    f"O campo Matrícula é obrigatório.",
                    "critical"
                )
                return
            
            if not caminho_pdf:
                self.mostrar_mensagem(
                    "Erro",
                    f"O campo PDF é obrigatório.",
                    "critical"
                )
                return
            
            if label_unidade is not None:
                if not unidade_calculo_str:
                    self.mostrar_mensagem(
                            "Erro",
                            f"O campo Unidade de Cálculo ({label_unidade}) é obrigatório.",
                            "critical"
                        )
                    return
                else:
                    try:
                        valor_informado_usuario = float(unidade_calculo_str)
                    except ValueError:
                        self.mostrar_mensagem("Erro", "Insira um número válido em 'Unidade de Cálculo'.", "critical")
                        return
                
            #valor_informado_usuario = 0
            # if label_unidade is not None:
            #     try:
            #         valor_informado_usuario = float(unidade_calculo_str)
            #     except ValueError:
            #         self.mostrar_mensagem("Erro", "Insira um número válido em 'Unidade de Cálculo'.", "critical")
            #         return

            indice_detalhe = self.caixa_detalhe.currentIndex()
            if indice_detalhe < 0:
                self.mostrar_mensagem("Erro", "Selecione o detalhe da atividade!", "critical")
                return

            (k_valor, calcHoras_valor, label_unidade) = self.caixa_detalhe.itemData(indice_detalhe)

            # Determina as horas-base (horas_base) e os detalhes de cálculo
            if calcHoras_valor is None:
                # Exemplo: "09. Visita Técnica" => K * valor_informado_usuario
                unidade_calculo = valor_informado_usuario
                texto_detalhes_calculo = f"K={k_valor} x Unidade ({label_unidade})={valor_informado_usuario}"
            else:
                # Se label_unidade for None => K x calcHoras_valor
                # Se label_unidade for algo => K x calcHoras_valor x valor_informado_usuario
                if label_unidade is not None:
                    unidade_calculo = calcHoras_valor * valor_informado_usuario
                    texto_detalhes_calculo = f"K={k_valor} x {calcHoras_valor} x Unidade ({label_unidade})={valor_informado_usuario}"
                else:
                    unidade_calculo = calcHoras_valor
                    texto_detalhes_calculo = f"K={k_valor} x {calcHoras_valor}"

            # Multiplica por K (SEM usar maxAproveitamento)
            horas_convertidas = k_valor * unidade_calculo

            # Aqui armazenamos no dicionário
            item = {
                "matricula": matricula,
                #"unidade": valor_informado_usuario,
                "pdf": caminho_pdf,
                "horas_convertidas": horas_convertidas,
                "data_inclusao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "detalhe_escolhido": detalhe_escolhido,
                "detalhes_calculo": texto_detalhes_calculo,  # Adicionando em português
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
        anexando PDFs associados, mas apenas para a matrícula informada.
        """
        try:
            matricula_filtro = self.campo_matricula.text().strip()
            if not matricula_filtro:
                self.mostrar_mensagem("Erro", "Informe a matrícula para gerar o relatório.", "critical")
                return

            # Renomeando poppler_path para caminho_poppler
            caminho_poppler = r"C:\Poppler\Library\bin"
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
                
                # Inserindo também o máximo para cada atividade no relatório
                if atividade in self.percentuais_atividade:
                    maximo_horas_atividade = self.percentuais_atividade[atividade] * 240
                    pdf.cell(0, 10, f"(MÁXIMO DE HORAS: {maximo_horas_atividade:.2f})", ln=True, align="L")

                for item in itens_filtrados:
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"Matrícula: {item['matricula']}", ln=True)
                    #pdf.cell(0, 10, f"Unidade de Cálculo: {item['unidade']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f} (Cálculo: {item.get('detalhes_calculo', '')})", ln=True)
                    if 'detalhe_escolhido' in item:
                        pdf.cell(0, 10, f"Detalhe: {item['detalhe_escolhido']}", ln=True)

                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)

                    # Renomeando pdf_path para caminho_pdf_item
                    caminho_pdf_item = os.path.normpath(item['pdf'])
                    pdf.ln(5)
                    if os.path.exists(caminho_pdf_item):
                        try:
                            imagens_convertidas = convert_from_path(caminho_pdf_item, dpi=100, poppler_path=caminho_poppler)
                            for imagem in imagens_convertidas:
                                # Renomeando unique_name -> nome_unico, image_path -> caminho_imagem
                                nome_unico = str(uuid.uuid4())
                                caminho_imagem = f"temp_image_{nome_unico}.jpg"
                                imagem.save(caminho_imagem, "JPEG")
                                pdf.image(caminho_imagem, x=10, w=150)
                                os.remove(caminho_imagem)
                        except Exception as e2:
                            pdf.cell(0, 10, f"Erro ao processar PDF ({caminho_pdf_item}): {e2}", ln=True)
                    else:
                        pdf.cell(0, 10, f"Arquivo PDF não encontrado: {caminho_pdf_item}", ln=True)

                    pdf.ln(10)

            pdf.output(nome_arquivo_relatorio)
            self.mostrar_mensagem("Sucesso", f"Relatório gerado com sucesso:\n{nome_arquivo_relatorio}", "information")
        except Exception as e:
            print(f"Erro em gerar_relatorio: {e}")
            self.mostrar_mensagem("Erro", f"Erro ao gerar relatório: {e}", "critical")

    def gerar_relatorio_exemplo(self):
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
                    #pdf.cell(0, 10, f"Unidade de Cálculo: {item['horas']}", ln=True)
                    pdf.cell(0, 10, f"Horas Convertidas: {item['horas_convertidas']:.2f}", ln=True)
                    pdf.cell(0, 10, f"Data Inclusão: {item['data_inclusao']}", ln=True)
                    pdf.ln(10)

            pdf.output(relatorio_path)

            merger = PdfMerger()
            merger.append(relatorio_path)

            for atividade, itens in dados_temp.items():
                for item in itens:
                    caminho_pdf_item = os.path.normpath(item["pdf"])
                    if os.path.exists(caminho_pdf_item):
                        merger.append(caminho_pdf_item)

            merger.write("relatorio_completo.pdf")
            merger.close()

            QMessageBox.information(self, "Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            print(f"Erro em gerar_relatorio_exemplo: {e}")
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
