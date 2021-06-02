# from pptx import Presentation
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt5 import uic

from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
from docx2pdf import convert

import sys


form_class = uic.loadUiType("design.ui")[0]  # Load the UI
DOCX_filename = None
TXT_filename = None
NAMES = None
author_info = "This program was created by Anna Ilina for IT school in Vladikavkaz."


def generate_certificates(name_docx_template, names, label):
	for name in names:
		name = name.replace('\n', '')
		doc = DocxTemplate(name_docx_template)
		context = {'name': name}
		doc.render(context)
		result_doc = f'certificates/certificate_{name}.docx'
		doc.save(result_doc)
		try:
			convert(result_doc, result_doc[:-4]+"pdf")
		except NotImplementedError:
			label.setText("[ERROR] Please install Microsoft Word to generate PDF")

	label.setText("[SUCCESS] Certificates are successfully generated at certificates/")


def main():
	print("This program was created by Anna Ilina for IT school in Vladikavkaz.")
	# args = sys.argv  # [имя файла с именами ("names.txt"), имя файла шаблона docx ("template.docx")]
	# if len(sys.argv) < 3 or args[1] == "--help":
	# 	print("Please input command like:\npython3 main.py names.txt docx-templates/template.docx")
	# else:
	# 	names = None
	# 	try:
	# 		with open(args[1], "r") as f:
	# 			names = f.readlines()
	# 			if len(names) == 0:
	# 				print(f"[ERROR] {args[1]} is empty")
	# 				return
	# 	except:
	# 		print(f"[ERROR] The file {args[1]} doesn't exist. Please check the path.")
	# 		return
	# 	try:
	# 		generate_certificates(args[2], names)
	# 	except PackageNotFoundError:
	# 		print(f"[ERROR] {args[2]} doesn't exist. Please check the path.")


class MyWindowClass(QMainWindow, form_class):
    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.label.setText(self.label.text() + "\n" + author_info)

    def actionAbout_clicked(self):
    	self.label.setText(author_info)

    def bt_docx_clicked(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","*.docx *.doc", options=options)
        if fileName:
            DOCX_filename = fileName
            self.tb_docx.setText(DOCX_filename)

    def bt_txt_clicked(self):
        global NAMES
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","*.txt", options=options)
        if fileName:
            TXT_filename = fileName
            self.tb_txt.setText(TXT_filename)
            try:
                with open(TXT_filename, "r") as f:
                    NAMES = f.readlines()
                    if len(NAMES) == 0:
                        self.label.setText(f"[ERROR] {TXT_filename} is empty")
                        return
            except:
                self.label.setText(f"[ERROR] The file {TXT_filename} doesn't exist. Please check the path.")
            self.bt_gen.setEnabled(True)

    def bt_gen_clicked(self):
        try:
            generate_certificates(DOCX_filename, NAMES, self.label)
        except PackageNotFoundError:
            self.label.setText(f"[ERROR] {DOCX_filename} doesn't exist. Please check the path.")


if __name__ == "__main__":
	app = QApplication(sys.argv)
	myWindow = MyWindowClass(None)
	myWindow.show()
	app.exec_()
