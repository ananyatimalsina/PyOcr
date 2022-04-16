from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Progressbar
from tkinter import filedialog
import requests
from configparser import ConfigParser
import pytesseract
import os
from requests_html import HTML
import threading
import ocrmypdf

config = ConfigParser()

try:
    config.read("config.ini")
    update = config["Settings"]["update"]

except Exception as e:
    print("Config File Is Corrupted or does not Exist! Error: " + str(e))

if update == "True":
    print("Updating packages.. Please be patient")
    print("Requesting tesseract urls..")
    tess_r = requests.get("https://github.com/UB-Mannheim/tesseract/wiki")
    tess_r = HTML(html=str(tess_r.content))
    tess_url = tess_r.links
    
    print("Requesting gostscript urls..")
    ghost_r = requests.get("https://github.com/ArtifexSoftware/ghostpdl-downloads/releases")
    ghost_r = HTML(html=str(ghost_r.content))
    ghost_url = ghost_r.links

    versions = []
    links = []

    print("Getting latest tesseract version")

    for i in tess_url:
        if "tesseract" in i and "w64" in i:
            links.append(i)
            versions.append(int(i.split("v")[1][:5].replace(".", "")))

    tess_url = links[versions.index(max(versions))]
    
    r_tess = requests.get(tess_url)

    versions = []
    links = []

    print("Getting latest gostscript version")

    for i in ghost_url:
        if "gs" in i and "w64" in i:
            try:
                versions.append(int(i.split("gs")[2].split("w")[0]))
            except:
                pass
            links.append("https://github.com" + i)
    
    ghost_url = links[versions.index(max(versions))]
    
    r_ghost = requests.get(ghost_url)

    print("Saving Files")
    
    open("tesseract.exe", "wb").write(r_tess.content)
    open("ghostscript.exe", "wb").write(r_ghost.content)

    print("Almost done! Please go through the Install steps now.")

    os.system("tesseract.exe")
    os.system("ghostscript.exe")

    print("Removing Files")

    os.remove("tesseract.exe")
    os.remove("ghostscript.exe")

    config.set("Settings", "update", "False")

    with open("config.ini", "w") as configfile:
        config.write(configfile)

    print("Sucessfully downloaded and installed the Following Packages: Tesseract, Ghostscript. Please make sure to add them to Path. Also restart your PC after Path Change")

def temp_bind(event):
    pass

def disable_widgets():
    select_path_main["state"] = "disabled"
    select_path_main.bind("<1>", temp_bind)
    preview["state"] = "disabled"
    language["state"] = "disabled"
    output["state"] = "disabled"
    output.bind("<1>", temp_bind)
    export_main["state"] = "disabled"

def enable_wdgets():
    select_path_main["state"] = "normal"
    select_path_main.bind("<1>", select_input)
    preview["state"] = "normal"
    language["state"] = "normal"
    output["state"] = "normal"
    output.bind("<1>", select_output)
    export_main["state"] = "normal"

def export():
    ofname = output.get()
    fname = select_path_main.get()

    if fname == "":
        messagebox.showerror("Empty Input!", "No input File provided!")
        return

    extention = fname.split(".")[-1]
    try:
        extentionof = ofname.split(".")[-1]
    except:
        pass
    def convert_pdf():
        progress.place(x=305, y=250)
        progress.start()

        if language.get() == "":
            ocrmypdf.ocr(fname, ofname, deskew=True)

        else:
            ocrmypdf.ocr(fname, ofname, deskew=True, language=language.get())

        progress.place_forget()
        progress.stop()
        enable_wdgets()
        messagebox.showinfo("Exported PDF!", "PDF has been sucesfully exported and was saved to: " + ofname)

    def convert_pdf_non():
        progress.place(x=305, y=250)
        progress.start()

        if language.get() == "":
            pdf = pytesseract.image_to_pdf_or_hocr(fname, extension='pdf', config='--psm 6')

        else:
            pdf = pytesseract.image_to_pdf_or_hocr(fname, extension='pdf', lang=language.get(), config='--psm 6')

        with open(ofname, 'w+b') as f:
            f.write(pdf)

        progress.place_forget()
        progress.stop()
        enable_wdgets()
        messagebox.showinfo("Exported PDF!", "PDF has been sucesfully exported and was saved to: " + ofname)

    def convert_txt():
        progress.place(x=305, y=250)
        progress.start()
        
        with open(ofname, "w") as f:
            f.write(preview.get())

        progress.place_forget()
        progress.stop()
        enable_wdgets()
        messagebox.showinfo("Exported Text File!", "Text File has been sucesfully exported and was saved to: " + ofname)

    if output.get() == "" and extention == "pdf":
        ofname = select_output("NORMAL", "PDF")
        extentionof = ofname.split(".")[-1]

    elif extention == "pdf" and extentionof != "pdf":
        ofname = select_output("NORMAL", "PDF")
        extentionof = ofname.split(".")[-1]

    else:
        ofname = select_output("NORMAL")
        extentionof = ofname.split(".")[-1]

    if extentionof == "pdf" and extention == "pdf":
        disable_widgets()
        threading.Thread(target=convert_pdf, daemon=True).start()

    elif extentionof == "pdf" and extention != "pdf":
        disable_widgets()
        threading.Thread(target=convert_pdf_non, daemon=True).start()

    else:
        disable_widgets()
        threading.Thread(target=convert_txt, daemon=True).start()

def select_output(event, mode = "NORMAL"):
    if mode == "NORMAL":
        ofname = filedialog.askopenfilename(title="Select Output File", filetypes=[("TextFiles and Documents", ".txt .pdf")])

    elif mode == "PDF":
        ofname = filedialog.askopenfilename(title="Select Output File", filetypes=[("Documents", ".pdf")])

    output.delete(0,"end")
    output.insert(0, ofname)

    return ofname

def select_input(event):
    fname = filedialog.askopenfilename(title="Select Input File", filetypes=[("Images and Documents", ".png .jpg .jpeg .pdf")])

    def convert():
        progress.place(x=305, y=250)
        progress.start()

        extention = fname.split(".")
        extention = extention[-1]
            
        if extention != "pdf":
            if language.get() != "":
                text = pytesseract.image_to_string(fname, language.get(), config='--psm 6')

            else:
                text = pytesseract.image_to_string(fname, config='--psm 6')

        else:
            export()

        progress.place_forget()
        progress.stop()

        enable_wdgets()

        preview.delete(1.0,"end")
        preview.insert(1.0, text)

    if fname !="":
        select_path_main.delete(0,"end")
        select_path_main.insert(0, fname)
        disable_widgets()
        threading.Thread(target=convert, daemon=True).start()


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

select_path_main_img = PhotoImage(file = f"assets/select_path_main.png")
select_path_main_bg = canvas.create_image(
    350.0, 51.0,
    image = select_path_main_img)

select_path_main = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

select_path_main.place(
    x = 50, y = 34,
    width = 600,
    height = 32)

select_path_main.bind("<1>", select_input)

preview_img = PhotoImage(file = f"assets/preview.png")
preview_bg = canvas.create_image(
    350.0, 261.0,
    image = preview_img)

preview = Text(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

preview.place(
    x = 50, y = 76,
    width = 600,
    height = 368)

language_img = PhotoImage(file = f"assets/language.png")
language_bg = canvas.create_image(
    211.5, 471.5,
    image = language_img)

language = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

language.place(
    x = 154, y = 454,
    width = 115,
    height = 33)

output_img = PhotoImage(file = f"assets/output.png")
output_bg = canvas.create_image(
    580.0, 471.5,
    image = output_img)

output = Entry(
    bd = 0,
    bg = "#ffffff",
    highlightthickness = 0)

output.place(
    x = 510, y = 454,
    width = 140,
    height = 33)

output.bind("<1>", select_output)

export_main_img = PhotoImage(file = f"assets/export_main.png")
export_main = Button(
    image = export_main_img,
    borderwidth = 0,
    highlightthickness = 0,
    command = export,
    relief = "flat")

export_main.place(
    x = 292, y = 454,
    width = 115,
    height = 35)

background_img = PhotoImage(file = f"assets/background_main.png")
background = canvas.create_image(
    277.5, 244.5,
    image=background_img)

progress = Progressbar(orient=HORIZONTAL, length=100, mode='indeterminate')

window.resizable(False, False)
window.mainloop()
