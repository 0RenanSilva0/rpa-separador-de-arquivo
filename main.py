import sys
import os
import pandas as pd
from PyQt5 import QtWidgets, QtGui, QtCore

class ExcelSplitter(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Divisor de Arquivos XLSX')
        self.setGeometry(100, 100, 600, 400)
        self.setFixedSize(600, 400)

        # Fonte personalizada
        font = QtGui.QFont('Arial', 10)

        # Layout principal
        main_layout = QtWidgets.QVBoxLayout()

        # Layout para seleção de arquivo
        file_layout = QtWidgets.QHBoxLayout()
        self.file_label = QtWidgets.QLabel('Arquivo XLSX:')
        self.file_label.setFont(font)
        self.file_input = QtWidgets.QLineEdit()
        self.file_input.setFont(font)
        self.file_button = QtWidgets.QPushButton('Procurar')
        self.file_button.setFont(font)
        self.file_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.file_input)
        file_layout.addWidget(self.file_button)

        # Layout para seleção de diretório de saída
        output_layout = QtWidgets.QHBoxLayout()
        self.output_label = QtWidgets.QLabel('Diretório de saída:')
        self.output_label.setFont(font)
        self.output_input = QtWidgets.QLineEdit()
        self.output_input.setFont(font)
        self.output_button = QtWidgets.QPushButton('Selecionar')
        self.output_button.setFont(font)
        self.output_button.clicked.connect(self.select_output_dir)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(self.output_button)

        # Layout para número de linhas por arquivo
        lines_layout = QtWidgets.QHBoxLayout()
        self.lines_label = QtWidgets.QLabel('Linhas por arquivo:')
        self.lines_label.setFont(font)
        self.lines_input = QtWidgets.QLineEdit()
        self.lines_input.setFont(font)
        lines_layout.addWidget(self.lines_label)
        lines_layout.addWidget(self.lines_input)

        # Botão de visualização
        self.preview_button = QtWidgets.QPushButton('Visualizar Dados')
        self.preview_button.setFont(font)
        self.preview_button.clicked.connect(self.preview_data)

        # Barra de progresso
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setFont(font)

        # Botões inferiores
        buttons_layout = QtWidgets.QHBoxLayout()
        self.start_button = QtWidgets.QPushButton('Iniciar')
        self.start_button.setFont(font)
        self.start_button.clicked.connect(self.start_splitting)
        self.about_button = QtWidgets.QPushButton('Sobre')
        self.about_button.setFont(font)
        self.about_button.clicked.connect(self.show_about)
        self.exit_button = QtWidgets.QPushButton('Sair')
        self.exit_button.setFont(font)
        self.exit_button.clicked.connect(self.close_application)
        buttons_layout.addWidget(self.start_button)
        buttons_layout.addWidget(self.about_button)
        buttons_layout.addWidget(self.exit_button)

        # Adicionando layouts ao layout principal
        main_layout.addLayout(file_layout)
        main_layout.addLayout(output_layout)
        main_layout.addLayout(lines_layout)
        main_layout.addWidget(self.preview_button)
        main_layout.addWidget(self.progress_bar)
        main_layout.addLayout(buttons_layout)

        self.setLayout(main_layout)

    def select_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Selecionar arquivo Excel', '', 'Arquivos Excel (*.xlsx *.xls)')
        if file_path:
            self.file_input.setText(file_path)

    def select_output_dir(self):
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, 'Selecionar diretório de saída')
        if dir_path:
            self.output_input.setText(dir_path)

    def preview_data(self):
        file_path = self.file_input.text()
        if not file_path:
            QtWidgets.QMessageBox.warning(self, 'Aviso', 'Por favor, selecione um arquivo Excel.')
            return
        try:
            df = pd.read_excel(file_path, nrows=10)
            preview = df.to_string()
            QtWidgets.QMessageBox.information(self, 'Prévia dos Dados', preview)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Erro', f'Erro ao ler o arquivo: {e}')

    def start_splitting(self):
        file_path = self.file_input.text()
        output_dir = self.output_input.text()
        lines_per_file = self.lines_input.text()

        if not file_path or not output_dir or not lines_per_file:
            QtWidgets.QMessageBox.warning(self, 'Aviso', 'Por favor, preencha todos os campos.')
            return

        try:
            lines_per_file = int(lines_per_file)
            if lines_per_file <= 0:
                raise ValueError
        except ValueError:
            QtWidgets.QMessageBox.critical(self, 'Erro', 'O número de linhas deve ser um inteiro positivo.')
            return

        try:
            df = pd.read_excel(file_path)
            total_rows = df.shape[0]
            num_files = total_rows // lines_per_file + (total_rows % lines_per_file > 0)

            base_filename = os.path.splitext(os.path.basename(file_path))[0]

            self.progress_bar.setMaximum(num_files)
            self.progress_bar.setValue(0)

            for i in range(num_files):
                start_row = i * lines_per_file
                end_row = start_row + lines_per_file
                chunk = df.iloc[start_row:end_row]
                output_filename = f"{base_filename}_parte_{i+1}.xlsx"
                output_path = os.path.join(output_dir, output_filename)
                chunk.to_excel(output_path, index=False)
                self.progress_bar.setValue(i + 1)
                QtCore.QCoreApplication.processEvents()

            QtWidgets.QMessageBox.information(self, 'Sucesso', f'Arquivo dividido em {num_files} partes.')
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Erro', f'Ocorreu um erro: {e}')

    def show_about(self):
        about_text = """
        <b>Divisor de Arquivos XLSX</b><br><br>
        Desenvolvido por Renan Silva.<br>
        Este aplicativo permite dividir arquivos Excel em partes menores.<br><br>
        <i>Versão 1.0</i>
        """
        QtWidgets.QMessageBox.about(self, 'Sobre', about_text)

    def close_application(self):
        choice = QtWidgets.QMessageBox.question(
            self, 'Confirmação', "Deseja realmente sair?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            sys.exit()

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ExcelSplitter()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
