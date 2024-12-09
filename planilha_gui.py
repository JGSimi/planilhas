import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QTextEdit, QPushButton, QLabel, 
    QLineEdit, QFileDialog, QMessageBox, QProgressBar,
    QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor
from planilha_ia import GeradorPlanilhas
from datetime import datetime
import requests

class StyledButton(QPushButton):
    def __init__(self, text, icon_name=None, primary=True):
        super().__init__(text)
        if icon_name:
            self.setIcon(QIcon(icon_name))
        self.primary = primary
        self.setup_style()
        
    def setup_style(self):
        if self.primary:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #007AFF;
                    color: white;
                    border: none;
                    border-radius: 8px;
                    padding: 12px 24px;
                    font-weight: bold;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                }
                QPushButton:pressed {
                    background-color: #004494;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                }
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #f8f9fa;
                    color: #333;
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    padding: 12px 24px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #e9ecef;
                    border-color: #bbb;
                }
                QPushButton:pressed {
                    background-color: #dde0e3;
                }
            """)

class AutoNameThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, prompt: str):
        super().__init__()
        self.prompt = prompt
        
    def run(self):
        try:
            headers = {"Content-Type": "application/json"}
            data = {
                "model": "mistral",
                "prompt": f"""Baseado na seguinte descrição de planilha, sugira um nome curto e descritivo para o arquivo (máximo 3 palavras, sem espaços):
                {self.prompt}
                
                Responda APENAS com o nome sugerido, sem extensão e sem explicações.
                Exemplo: controle_financeiro
                """,
                "stream": False
            }
            
            # Faz a requisição diretamente para o Ollama
            response = requests.post(
                "http://localhost:11434/api/generate",
                headers=headers,
                json=data
            )
            response.raise_for_status()
            
            # Extrai o texto da resposta
            nome = response.json()["response"].strip()
            
            # Limpa e formata o nome
            nome = nome.strip().replace(" ", "_").lower()
            nome = ''.join(e for e in nome if e.isalnum() or e == '_')
            
            if nome:
                self.finished.emit(nome)
            else:
                raise ValueError("Nome gerado está vazio")
        except Exception as e:
            self.error.emit(str(e))

class FilePathWidget(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        
        # Label
        self.label = QLabel("Local de Salvamento:")
        self.label.setStyleSheet("color: #333; font-weight: bold;")
        layout.addWidget(self.label)
        
        # Layout para nome do arquivo e caminho
        file_layout = QHBoxLayout()
        
        # Campo de nome do arquivo
        self.name_layout = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Nome do arquivo")
        self.name_input.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: white;
                color: #333;
                margin-right: 5px;
                min-width: 200px;
            }
            QLineEdit:focus {
                border-color: #007AFF;
            }
        """)
        self.name_layout.addWidget(self.name_input)
        
        # Botão de auto-nome com ícone de varinha mágica
        self.auto_name_btn = QPushButton()
        self.auto_name_btn.setIcon(QIcon.fromTheme("edit-redo"))
        self.auto_name_btn.setToolTip("Gerar nome automaticamente baseado na descrição")
        self.auto_name_btn.setFixedSize(32, 32)
        self.auto_name_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-right: 10px;
                qproperty-iconSize: QSize(16, 16);
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        self.auto_name_btn.clicked.connect(self.generate_auto_name)
        self.name_layout.addWidget(self.auto_name_btn)
        
        file_layout.addLayout(self.name_layout)
        
        # Campo de caminho
        path_layout = QHBoxLayout()
        self.path_display = QLineEdit()
        self.path_display.setPlaceholderText("Local para salvar...")
        self.path_display.setReadOnly(True)
        self.path_display.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: #f8f9fa;
                color: #333;
            }
        """)
        path_layout.addWidget(self.path_display)
        
        self.browse_btn = StyledButton("Escolher Local", primary=False)
        path_layout.addWidget(self.browse_btn)
        
        file_layout.addLayout(path_layout)
        
        layout.addLayout(file_layout)
        
        # Estilo do frame
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        
    def generate_auto_name(self):
        if not hasattr(self.parent, 'prompt_text'):
            return
            
        prompt = self.parent.prompt_text.toPlainText().strip()
        if not prompt:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, descreva a planilha primeiro para gerar um nome automático."
            )
            return
            
        # Desabilita o botão durante a geração
        self.auto_name_btn.setEnabled(False)
        self.auto_name_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-right: 10px;
                opacity: 0.5;
            }
        """)
        
        # Inicia thread de geração do nome
        self.thread = AutoNameThread(prompt)
        self.thread.finished.connect(self.auto_name_finished)
        self.thread.error.connect(self.auto_name_error)
        self.thread.start()
        
    def auto_name_finished(self, nome):
        self.name_input.setText(f"{nome}")
        self.auto_name_btn.setEnabled(True)
        self.auto_name_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-right: 10px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        
    def auto_name_error(self, error):
        self.auto_name_btn.setEnabled(True)
        self.auto_name_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-right: 10px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        QMessageBox.warning(
            self,
            "Erro",
            f"Erro ao gerar nome automático: {error}"
        )
        
    def get_full_path(self) -> str:
        """Retorna o caminho completo do arquivo"""
        path = self.path_display.text()
        name = self.name_input.text().strip()
        
        if not name:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            name = f"planilha_{timestamp}.xlsx"
        elif not name.endswith('.xlsx'):
            name += '.xlsx'
            
        if not path:
            return name
            
        return os.path.join(path, name)
        
    def choose_location(self):
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Escolher Local para Salvar",
            os.path.expanduser("~/Documents"),
            QFileDialog.Option.ShowDirsOnly
        )
        if dir_path:
            self.path_display.setText(dir_path)

class GeradorThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, prompt, nome_arquivo=None):
        super().__init__()
        self.prompt = prompt
        self.nome_arquivo = nome_arquivo
        
    def run(self):
        try:
            gerador = GeradorPlanilhas()
            arquivo = gerador.gerar_planilha(self.prompt, self.nome_arquivo)
            self.finished.emit(arquivo)
        except Exception as e:
            self.error.emit(str(e))

class PromptEnhancerThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, current_prompt: str = ""):
        super().__init__()
        self.current_prompt = current_prompt
        
    def run(self):
        try:
            headers = {"Content-Type": "application/json"}
            
            if self.current_prompt.strip():
                prompt_text = f"""Melhore o seguinte prompt para geração de planilha, adicionando mais detalhes e especificações úteis:
                "{self.current_prompt}"
                
                Responda apenas com o prompt melhorado, sem explicações adicionais.
                Exemplo de melhoria:
                Original: "lista de tarefas"
                Melhorado: "lista de tarefas com título, descrição, prioridade (alta/média/baixa), status (pendente/em andamento/concluído) e data de entrega"
                """
            else:
                prompt_text = """Sugira um prompt completo e detalhado para gerar uma planilha útil.
                O prompt deve incluir todos os campos e especificações necessários.
                
                Responda apenas com o prompt sugerido, sem explicações adicionais.
                Exemplo: "planilha de controle financeiro pessoal com data, categoria (receita/despesa), descrição, valor, método de pagamento e status (pago/pendente)"
                """
            
            data = {
                "model": "mistral",
                "prompt": prompt_text,
                "stream": False
            }
            
            response = requests.post(
                "http://localhost:11434/api/generate",
                headers=headers,
                json=data
            )
            response.raise_for_status()
            
            prompt_melhorado = response.json()["response"].strip()
            if prompt_melhorado:
                self.finished.emit(prompt_melhorado)
            else:
                raise ValueError("Prompt gerado está vazio")
        except Exception as e:
            self.error.emit(str(e))

class PromptWidget(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        
        # Label
        prompt_label = QLabel("O que você deseja criar?")
        prompt_label.setStyleSheet("color: #333; font-weight: bold; font-size: 16px;")
        layout.addWidget(prompt_label)
        
        # Layout para o campo de texto e botão
        input_layout = QHBoxLayout()
        
        # Campo de texto
        self.prompt_text = QTextEdit()
        self.prompt_text.setPlaceholderText("Ex: Crie uma planilha de controle financeiro com colunas para data, descrição, valor e categoria...")
        self.prompt_text.setMinimumHeight(120)
        self.prompt_text.setStyleSheet("""
            QTextEdit {
                padding: 12px;
                border: 2px solid #ddd;
                border-radius: 10px;
                background-color: white;
                font-size: 14px;
            }
            QTextEdit:focus {
                border-color: #007AFF;
            }
        """)
        input_layout.addWidget(self.prompt_text)
        
        # Layout vertical para os botões
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(5)
        
        # Botão de melhoria de prompt
        self.enhance_btn = QPushButton()
        self.enhance_btn.setIcon(QIcon.fromTheme("edit-find-replace"))
        self.enhance_btn.setToolTip("Melhorar/Gerar prompt automaticamente")
        self.enhance_btn.setFixedSize(32, 32)
        self.enhance_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-left: 10px;
                qproperty-iconSize: QSize(16, 16);
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        self.enhance_btn.clicked.connect(self.enhance_prompt)
        buttons_layout.addWidget(self.enhance_btn)
        
        # Adiciona espaçador para alinhar o botão ao topo
        buttons_layout.addStretch()
        
        input_layout.addLayout(buttons_layout)
        layout.addLayout(input_layout)
        
        # Estilo do frame
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        
    def enhance_prompt(self):
        # Desabilita o botão durante a geração
        self.enhance_btn.setEnabled(False)
        self.enhance_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-left: 10px;
                opacity: 0.5;
            }
        """)
        
        # Inicia thread de melhoria do prompt
        self.thread = PromptEnhancerThread(self.prompt_text.toPlainText())
        self.thread.finished.connect(self.prompt_enhanced)
        self.thread.error.connect(self.prompt_error)
        self.thread.start()
        
    def prompt_enhanced(self, novo_prompt):
        self.prompt_text.setText(novo_prompt)
        self.enhance_btn.setEnabled(True)
        self.enhance_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-left: 10px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        
    def prompt_error(self, error):
        self.enhance_btn.setEnabled(True)
        self.enhance_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 16px;
                margin-left: 10px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #007AFF;
            }
            QPushButton:pressed {
                background-color: #dde0e3;
            }
        """)
        QMessageBox.warning(
            self,
            "Erro",
            f"Erro ao melhorar prompt: {error}"
        )

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Planilhas com IA 📊")
        self.setMinimumSize(800, 600)
        self.setup_ui()
        
    def setup_ui(self):
        # Widget central com margem
        container = QWidget()
        self.setCentralWidget(container)
        main_layout = QVBoxLayout(container)
        main_layout.setSpacing(25)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # Cabeçalho
        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setSpacing(10)
        
        # Título
        title = QLabel("Gerador de Planilhas com IA")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(24)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet("color: #2C3E50;")
        header_layout.addWidget(title)
        
        # Subtítulo
        subtitle = QLabel("Crie planilhas automaticamente usando Inteligência Artificial")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #7F8C8D; font-size: 16px;")
        header_layout.addWidget(subtitle)
        
        main_layout.addWidget(header)
        
        # Área principal
        content = QFrame()
        content.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 15px;
                padding: 20px;
                border: 1px solid #ddd;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                color: #000;
            }
        """)
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(20)
        
        # Widget de prompt
        self.prompt_widget = PromptWidget(self)
        content_layout.addWidget(self.prompt_widget)
        
        # Widget de seleção de arquivo
        self.file_widget = FilePathWidget(self)
        self.file_widget.browse_btn.clicked.connect(self.file_widget.choose_location)
        content_layout.addWidget(self.file_widget)
        
        # Barra de progresso
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 10px;
                text-align: center;
                background-color: #f0f0f0;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #007AFF;
                border-radius: 10px;
            }
        """)
        content_layout.addWidget(self.progress)
        
        # Botão gerar
        self.generate_btn = StyledButton("Gerar Planilha")
        self.generate_btn.setMinimumHeight(50)
        self.generate_btn.clicked.connect(self.generate_spreadsheet)
        content_layout.addWidget(self.generate_btn)
        
        main_layout.addWidget(content)
        
        # Exemplos
        examples_frame = QFrame()
        examples_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        examples_layout = QVBoxLayout(examples_frame)
        
        examples_title = QLabel("Exemplos de prompts:")
        examples_title.setStyleSheet("color: #333; font-weight: bold;")
        examples_layout.addWidget(examples_title)
        
        examples = [
            "• Planilha de controle financeiro com data, descrição, valor e categoria",
            "• Lista de tarefas com título, prioridade, status e prazo",
            "• Planilha de contatos com nome, email, telefone e empresa",
            "• Inventário de produtos com código, nome, quantidade e preço"
        ]
        
        for example in examples:
            example_label = QLabel(example)
            example_label.setStyleSheet("color: #666; padding: 3px 0;")
            examples_layout.addWidget(example_label)
            
        main_layout.addWidget(examples_frame)
        
        # Estilização da janela principal
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
        """)
        
    def generate_spreadsheet(self):
        prompt = self.prompt_widget.prompt_text.toPlainText().strip()
        if not prompt:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, descreva a planilha que você deseja gerar."
            )
            return
            
        # Verifica se tem um nome de arquivo
        file_path = self.file_widget.get_full_path()
        if not file_path:
            QMessageBox.warning(
                self,
                "Atenção",
                "Por favor, forneça um nome para o arquivo ou escolha um local para salvar."
            )
            return
            
        # Confirma sobrescrita se arquivo existir
        if os.path.exists(file_path):
            reply = QMessageBox.question(
                self,
                "Confirmar Sobrescrita",
                f"O arquivo {os.path.basename(file_path)} já existe.\nDeseja sobrescrevê-lo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
            
        self.generate_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)  # Modo indeterminado
        
        # Iniciar thread de geração
        self.thread = GeradorThread(prompt, file_path)
        self.thread.finished.connect(self.generation_finished)
        self.thread.error.connect(self.generation_error)
        self.thread.start()
        
    def generation_finished(self, arquivo):
        self.progress.setVisible(False)
        self.generate_btn.setEnabled(True)
        
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Sucesso")
        msg.setText("Planilha gerada com sucesso!")
        msg.setInformativeText(f"Arquivo salvo em:\n{arquivo}")
        
        # Botão para abrir a pasta
        open_folder = msg.addButton("Abrir Pasta", QMessageBox.ButtonRole.ActionRole)
        msg.addButton(QMessageBox.StandardButton.Ok)
        
        msg.exec()
        
        if msg.clickedButton() == open_folder:
            os.system(f'open -R "{arquivo}"')
        
    def generation_error(self, error):
        self.progress.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(
            self,
            "Erro",
            f"Erro ao gerar planilha:\n{error}"
        )

def main():
    app = QApplication(sys.argv)
    
    # Configurar fonte padrão
    font = app.font()
    font.setFamily("SF Pro Display")  # Fonte padrão do macOS
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 