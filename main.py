import shutil
from tkinter import *
from tkinter import filedialog
import requests
from configparser import ConfigParser
from tkinter import messagebox
import pytesseract
import os
import zipfile
import pdf2image
from PyPDF2 import PdfFileMerger

config = ConfigParser()

try:
    config.read("config.ini")
    update = config["Settings"]["update"]
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    popp_path = r'popp/' + os.listdir("popp")[0] + r'/Library/bin'

except Exception as e:
    messagebox.showerror("Config file Error!", "Config File Is Corrupted or does not Exist!" + str(e))

if update == "True":
    if os.path.isdir("popp"):
        shutil.rmtree("popp")

    tess_url = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.1.20220118.exe"
    r_tess = requests.get(tess_url)
    open("tesseract.exe", "wb").write(r_tess.content)

    popp_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v22.01.0-0/Release-22.01.0-0.zip"
    r_popp = requests.get(popp_url)
    open("popp.zip", "wb").write(r_popp.content)

    with zipfile.ZipFile("popp.zip", "r") as zf:
        zf.extractall("popp")

    os.system("tesseract.exe")

    os.remove("tesseract.exe")
    os.remove("popp.zip")

    config.set("Settings", "update", "False")

    with open("config.ini", "w") as configfile:
        config.write(configfile)

def btn_clicked():
    fname = entry0.get()
    extention = fname.split(".")

    if extention[-1] != "pdf":
        if entry2.get() != "":
            pdf = pytesseract.image_to_pdf_or_hocr(fname, entry2.get(), extension='pdf')

        else:
            pdf = pytesseract.image_to_pdf_or_hocr(fname, extension='pdf')
        
        with open('Exports/output.pdf', 'w+b') as f:
            f.write(pdf)

    else:
        images = pdf2image.convert_from_path(fname, poppler_path=popp_path)
        c = 0
        pdfs = []
        for i in images:
            if entry2.get() != "":
                pdf = pytesseract.image_to_pdf_or_hocr(i, entry2.get(), extension='pdf')

            else:
                pdf = pytesseract.image_to_pdf_or_hocr(i, extension='pdf')

            with open('Exports/output' + str(c) + '.pdf', 'w+b') as f:
                f.write(pdf)
                pdfs.append('Exports/output' + str(c) + '.pdf')
        
        merger = PdfFileMerger()

        for pdf in pdfs:
            merger.append(pdf)

        merger.write("Exports/output.pdf")
        merger.close()

        for pdf in pdfs:
            os.remove(pdf)

    messagebox.showinfo("Succesfully Exported Document!", "Succesfully finished exporting your document! Its located in Output.pdf")

def select_input(event):
    fname = filedialog.askopenfilename(title="Select Input Image", filetypes=[("Images and Documents", ".png .jpg .jpeg .pdf")])

    if fname != "":

        extention = fname.split(".")
        
        if extention[-1] != "pdf":
            if entry2.get() != "":
                text = pytesseract.image_to_string(fname, entry2.get(), config='--psm 6')

            else:
                text = pytesseract.image_to_string(fname, config='--psm 6')

            entry0.delete(0,"end")
            entry0.insert(0, fname)
            entry1.delete(1.0,"end")
            entry1.insert(1.0, text)

        else:
            images = pdf2image.convert_from_path(fname, poppler_path=popp_path)
            exported_images = []

            for i in images:
                if entry2.get() != "":
                    exported_images.append(pytesseract.image_to_string(i, entry2.get(), config='--psm 6'))

                else:
                    exported_images.append(pytesseract.image_to_string(i, config='--psm 6'))

            exported_images = "\nPAGE END\n".join(exported_images)

            entry0.delete(0,"end")
            entry0.insert(0, fname)
            entry1.delete(1.0,"end")
            entry1.insert(1.0, exported_images)
            

window = Tk()

window.geometry("700x500")
window.configure(bg = "#ffffff")
window.title("PyOcr")
window.iconbitmap("Icon.ico")

canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 500,
    width = 700,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

entry0_img = PhotoImage(file = f"assets/img_textBox0.png")
entry0_bg = canvas.create_image(
    350.0, 51.0,
    image = entry0_img)

entry0 = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

entry0.place(
    x = 50, y = 34,
    width = 600,
    height = 32)

entry0.bind("<1>", select_input)

entry1_img = PhotoImage(file = f"assets/img_textBox1.png")
entry1_bg = canvas.create_image(
    350.0, 261.0,
    image = entry1_img)

entry1 = Text(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

entry1.place(
    x = 50, y = 76,
    width = 600,
    height = 368)

entry2_img = PhotoImage(file = f"assets/img_textBox2.png")
entry2_bg = canvas.create_image(
    211.5, 471.5,
    image = entry2_img)

entry2 = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

entry2.place(
    x = 154, y = 454,
    width = 115,
    height = 33)

img0 = PhotoImage(file = f"assets/img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b0.place(
    x = 292, y = 454,
    width = 115,
    height = 35)

background_img = PhotoImage(file = f"assets/background.png")
background = canvas.create_image(
    214.5, 244.5,
    image=background_img)

window.resizable(False, False)
window.mainloop()