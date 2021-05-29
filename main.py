# from pptx import Presentation
from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
from docx2pdf import convert
import sys


def generate_certificates(name_docx_template, names):
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
			print("[ERROR] Please install Microsoft Word to generate PDF")

	print("[SUCCESS] Certificates are successfully generated at certificates/")


def main():
	print("This program was created by Anna Ilina for IT school in Vladikavkaz.")
	args = sys.argv  # [имя файла с именами ("names.txt"), имя файла шаблона docx ("template.docx")]
	if len(sys.argv) < 3 or args[1] == "--help":
		print("Please input command like:\npython3 main.py names.txt docx-templates/template.docx")
	else:
		names = None
		try:
			with open(args[1], "r") as f:
				names = f.readlines()
				if len(names) == 0:
					print(f"[ERROR] {args[1]} is empty")
					return
		except:
			print(f"[ERROR] The file {args[1]} doesn't exist. Please check the path.")
			return
		try:
			generate_certificates(args[2], names)
		except PackageNotFoundError:
			print(f"[ERROR] {args[2]} doesn't exist. Please check the path.")


if __name__ == "__main__":
	main()
