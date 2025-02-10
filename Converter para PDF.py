from docx2pdf import convert
import os
from tkinter import filedialog as fd
from pptxtopdf import convert as pptx
from fpdf import FPDF
from aspose.cells import Workbook, PdfSaveOptions  # Removido PdfCompliance

# Selecionar Arquivos
def open_file_selection():
    filenames = fd.askopenfilenames()
    return filenames

# Converter TXT
def arquivoTXT():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=15)

    filenames = open_file_selection()

    for filename in filenames:
        with open(filename, 'r', encoding='utf-8') as f:
            for line in f:
                pdf.cell(200, 10, txt=line.strip(), ln=1, align='C')  # Removido espaço em branco

    pdf.output("arquivoTXT.pdf")

# Converter Word para PDF
def arquivoWord():
    filenames = open_file_selection()

    for filename in filenames:
        abs_filename = os.path.abspath(filename)
        output_filename = os.path.splitext(abs_filename)[0] + ".pdf"
        convert(abs_filename, output_filename)

# Converter PowerPoint para PDF
def arquivoPowerPoint():
    input_dir = open_file_selection()

    for filename in input_dir:
        output_dir = r"./"
        pptx(filename, output_dir)

# Converter Excel para PDF
def arquivoExcel():
    filenames = open_file_selection()

    for filename in filenames:
        workbook = Workbook(filename)
        pdfOptions = PdfSaveOptions()
        # Removido PdfCompliance se não estiver disponível
        workbook.save("arquivoExcel.pdf", pdfOptions)

# Menu de Opções e métodos de escolha
def display_menu():
    os.system("cls" if os.name == "nt" else "clear")  # Compatível com Windows e Unix
    print("\n## Trabalho do dia a dia: ##\n")
    print("1. arquivo TXT")
    print("2. arquivo Word")
    print("3. arquivo PowerPoint")
    print("4. arquivo Excel")
    print("0. Sair")

if __name__ == "__main__":
    while True:
        display_menu()
        choice = input("\nEscolha uma opção: ")

        if choice == "1":
            os.system('cls' if os.name == "nt" else "clear")
            arquivoTXT()
        elif choice == "2":
            os.system('cls' if os.name == "nt" else "clear")
            arquivoWord()
        elif choice == "3":
            arquivoPowerPoint()
            os.system('cls' if os.name == "nt" else "clear")
        elif choice == "4":
            arquivoExcel()
            os.system('cls' if os.name == "nt" else "clear")
        elif choice == "0":
            os.system('cls' if os.name == "nt" else "clear")
            print("Saindo...")
            break
        else:
            os.system('cls' if os.name == "nt" else "clear")
            print("Opção inválida. Tente novamente.")
