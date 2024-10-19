import os
import pdfkit
import subprocess
import shutil
import zipfile
import re
import logging
from PyPDF2 import PdfMerger ,PdfReader, PdfWriter
from bs4 import BeautifulSoup
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import qt_material
# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
progress = 0
total_pages = 0

# Configuration for wkhtmltopdf
path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
options = {
    'page-size': 'A4',
    'encoding': 'UTF-8',
    'enable-local-file-access': '',
    'margin-top': '0.75in',
    'margin-right': '0.75in',
    'margin-bottom': '0.75in',
    'margin-left': '0.75in',
}

config = pdfkit.configuration(wkhtmltopdf=path)
main_path = os.path.abspath(__file__).replace(os.path.abspath(__file__).split('\\')[-1], "")

# Backend functions
def run_cmd_command(command):
    try:
        result = subprocess.run(command, shell=True, check=True, text=True,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                encoding='utf-8', errors='replace')
        if result.stderr:
            logging.error("Error: %s", result.stderr)
        return result.stdout
    except subprocess.CalledProcessError as e:
        logging.error("Command failed with error: %s", e)
        return None

def find_file(file_name, search_path):
    for root, dirs, files in os.walk(search_path):
        if file_name in files:
            return os.path.join(root, file_name)
    logging.warning("File '%s' not found in '%s'.", file_name, search_path)
    return None

def unzip_file(zip_file_path, extract_to_folder):
    if os.path.exists(zip_file_path):
        os.makedirs(extract_to_folder, exist_ok=True)
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to_folder)
        logging.info("Successfully unzipped %s to %s", zip_file_path, extract_to_folder)
    else:
        logging.error("The zip file does not exist: %s", zip_file_path)

def solve_xhtml_tags(xhtml):
    soup = BeautifulSoup(xhtml, "html.parser")
    return str(soup.prettify())

def remove_xhtml_comments(xhtml):
    comment_pattern = re.compile(r'<!--.*?-->', re.DOTALL)
    return re.sub(comment_pattern, '', xhtml)

def process_xhtml(item, merger, extract_to_folder):
    try:
        with open(item, "r", encoding="utf-8") as xhtml_file:
            xhtml = xhtml_file.read()

        cssjs = []
        xhtml_list = xhtml.strip().split("\"")

        for item_2 in xhtml_list:
            if item_2.endswith('.css') or item_2.endswith('.js'):
                cssjs_filename = item_2.split("/")[-1]
                if cssjs_filename:
                    filepath = find_file(cssjs_filename, extract_to_folder)
                    if filepath:
                        cssjs.append([item_2, 'file:\\\\' + filepath])
                    else:
                        logging.warning("File '%s' not found for reference in '%s'.", cssjs_filename, item)

        for change in cssjs:
            xhtml = xhtml.replace(change[0], change[1])

        xhtml = re.sub(r'<a\s+[^>]*href="([^"]+)"[^>]*>(.*?)</a>', r'\1', xhtml)
        xhtml = "\n".join(line for line in xhtml.split("\n") if "kInteractive video" not in line)

        with open(item, "w", encoding="utf-8") as new_xhtml:
            new_xhtml.write(remove_xhtml_comments(solve_xhtml_tags(xhtml)))

        pdf_output = item[:-5] + "pdf"
        pdfkit.from_file(item, pdf_output, options=options, configuration=config)
        merger.append(pdf_output)

    except Exception as e:
        logging.error("Error processing %s: %s", item, e)

def exe_to_pdf(exe_filepaths, progress_callback):
    for exe_filepath in exe_filepaths:
        exe_filedir = os.path.dirname(exe_filepath)
        shutil.copy(main_path+r"arc_unpacker.exe",exe_filedir+r"\arc_unpacker.exe")
        command = "\""+exe_filedir+"\\arc_unpacker.exe\" --dec=microsoft/exe \"" + exe_filepath+"\""
        run_cmd_command("cd /d \""+exe_filedir+"\" && "+command)

        extradatafile = os.path.join(exe_filepath.replace(".exe", "") + "~.exe", "extra_data")
        extract_to_folder = extradatafile + "_unzipped"
        unzip_file(extradatafile, extract_to_folder)

        opf_file_path = os.path.join(extract_to_folder, 'epub', 'EPUB', 'package.opf')
        if os.path.exists(opf_file_path):
            with open(opf_file_path, 'r', encoding="utf-8") as opf_file:
                opf = opf_file.read()
                opf_items = opf[opf.find("item")-1:opf.find('<item id="ncx"')].strip().split('\n')

            list_of_xhtmls = [
                os.path.join(extract_to_folder, 'epub', 'EPUB', item[item.find("href=") + 6:item.find(".xhtml") + 6].replace("/","\\"))
                for item in opf_items
            ]
            list_of_xhtmls = list_of_xhtmls[:-1]
            merger = PdfMerger()
            total_count = len(list_of_xhtmls)

            for progress_counter, item in enumerate(list_of_xhtmls):
                process_xhtml(item, merger, extract_to_folder)
                progress = int((progress_counter + 1) / total_count * 95)
                progress_callback(progress)

            merger_output = exe_filepath[:-3] + "pd"
            merger.write(merger_output)
            merger.close()

            reader = PdfReader(merger_output)
            writer = PdfWriter()
            global total_pages
            for page in reader.pages:
                if page.extract_text().strip():
                    writer.add_page(page)
                    total_pages +=1
            with open(merger_output+'f',"wb") as output_pdf:
                writer.write(output_pdf)

            logging.info("Merged PDF written to %s", merger_output)
        else:
            logging.error("The OPF file does not exist: %s", opf_file_path)

        shutil.rmtree(exe_filepath.replace(".exe", "") + "~.exe")
        os.remove(exe_filedir + r"\arc_unpacker.exe")
        os.remove(merger_output)
        logging.info("Cleaned up temporary files.")
    progress = 100
    progress_callback(progress)

# Worker Thread
class WorkerThread(QtCore.QThread):
    progress_changed = QtCore.pyqtSignal(int)

    def __init__(self, exe_filepath):
        super().__init__()
        self.exe_filepath = exe_filepath

    def run(self):
        exe_to_pdf(self.exe_filepath, self.progress_changed.emit)
        

# Frontend Integration
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(640, 350)
        MainWindow.setMinimumSize(QtCore.QSize(640, 350))
        MainWindow.setMaximumSize(QtCore.QSize(640, 350))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.locLbl = QtWidgets.QLabel(self.centralwidget)
        self.locLbl.setGeometry(QtCore.QRect(205, 17, 411, 81))
        self.locLbl.setObjectName("locLbl")
        self.strtBtn = QtWidgets.QPushButton(self.centralwidget)
        self.strtBtn.setGeometry(QtCore.QRect(190, 245, 251, 81))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        font.setWeight(75)
        self.strtBtn.setFont(font)
        self.strtBtn.setObjectName("strtBtn")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 120, 626, 86))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.browseBtn = QtWidgets.QPushButton(self.centralwidget)
        self.browseBtn.setGeometry(QtCore.QRect(10, 15, 176, 86))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        font.setWeight(75)
        self.browseBtn.setFont(font)
        self.browseBtn.setObjectName("browseBtn")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Connect buttons
        self.browseBtn.clicked.connect(self.browse_files)
        self.strtBtn.clicked.connect(self.start_conversion)

    def browse_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("EXE Files (*.exe)")
        if file_dialog.exec_():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.exe_filepath = selected_files
                self.locLbl.setText("Files Selected")
                logging.info("Selected file: %s", self.exe_filepath)

    def start_conversion(self):
        if hasattr(self, 'exe_filepath'):
            self.progressBar.setValue(0)
            self.worker = WorkerThread(self.exe_filepath)
            self.worker.progress_changed.connect(self.update_progress)
            self.worker.start()
            self.strtBtn.setDisabled(True)
            self.browseBtn.setDisabled(True)
        else:
            logging.warning("No file selected!")
        

    def update_progress(self, value):
        self.progressBar.setValue(value)
        if value == 100:
            global total_pages
            total_price = 0
            import math
            for i in range(1,total_pages+1):
                total_price += 1.99 + (0.51)**i
            total_price = math.floor(total_price/50)*50
            self.locLbl.setText("Total pages is: " +str(total_pages)+" Total Price is "+str(total_price))

            self.browseBtn.setEnabled(True)
            self.strtBtn.setEnabled(True)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Executable to PDF Converter"))
        self.locLbl.setText(_translate("MainWindow", "Select your EXE file to convert"))
        self.strtBtn.setText(_translate("MainWindow", "Convert"))
        self.browseBtn.setText(_translate("MainWindow", "Browse"))

# Main execution
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
