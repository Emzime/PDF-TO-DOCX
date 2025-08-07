import os
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import tkinterdnd2 as tkdnd2
from PyPDF2 import PdfReader
from docx import Document  # pour fusionner les docx
from pdf2docx import Converter

chemin_pdf = ""
dossier_sortie = ""


def update_progress(percent):
    label_progression.config(text=f"{int(percent * 100)} %")


def update_status(message):
    label_statut.config(text=message)


def get_total_pages(pdf_path: str) -> int:
    pdf_path = str(pdf_path)
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        return len(reader.pages)


def get_unique_filepath(folder: str, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    counter = 1
    candidate = filename
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base} ({counter}){ext}"
        counter += 1
    return os.path.join(folder, candidate)


def merge_docx(files, output_path):
    merged_document = Document()

    # Just append the content of each document, without trying to remove paragraph
    for file in files:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save(output_path)


def convertir_pdf_en_docx(pdf_path: str, output_folder: str):
    pdf_path = str(pdf_path)
    output_folder = str(output_folder)
    try:
        nom_fichier = os.path.splitext(os.path.basename(pdf_path))[0]
        filename = f"{nom_fichier}.docx"
        output_path = get_unique_filepath(output_folder, filename)

        update_status("Conversion en cours...")

        # Disable buttons instead of hiding
        btn_convertir.config(state='disabled')
        btn_choisir_dossier.config(state='disabled')

        label_progression.config(text="0 %")
        label_progression.pack(pady=10)

        total_pages = get_total_pages(pdf_path)
        convertisseur = Converter(pdf_path)

        temp_dir = tempfile.mkdtemp()
        temp_files = []

        for page_number in range(total_pages):
            temp_file = os.path.join(temp_dir, f"page_{page_number + 1}.docx")
            convertisseur.convert(temp_file, start=page_number, end=page_number + 1)
            temp_files.append(temp_file)
            percent = (page_number + 1) / total_pages
            update_progress(percent)

        convertisseur.close()

        merge_docx(temp_files, output_path)
        shutil.rmtree(temp_dir)

        update_status("✅ Conversion terminée")
        messagebox.showinfo("Succès", f"Conversion terminée :\n{output_path}")

    except Exception as e:
        update_status("❌ Erreur pendant la conversion")
        messagebox.showerror("Erreur", f"Erreur pendant la conversion :\n{str(e)}")

    finally:
        # Always re-enable buttons and hide progress label
        btn_convertir.config(state='normal')
        btn_choisir_dossier.config(state='normal')
        label_progression.pack_forget()


def lancer_conversion():
    global chemin_pdf, dossier_sortie
    chemin_pdf = str(chemin_pdf)
    dossier_sortie = str(dossier_sortie)
    if not chemin_pdf or not dossier_sortie:
        messagebox.showwarning("Attention", "Veuillez sélectionner un fichier PDF et un dossier de sortie.")
        return

    threading.Thread(target=convertir_pdf_en_docx, args=(chemin_pdf, dossier_sortie), daemon=True).start()


def on_drop(event):
    global chemin_pdf
    chemin_pdf = str(event.data).strip('{}')
    label_fichier.config(text=os.path.basename(chemin_pdf))


def choisir_dossier():
    global dossier_sortie
    dossier = filedialog.askdirectory()
    if dossier:
        dossier_sortie = str(dossier)
        label_dossier.config(text=dossier)


fenetre = tkdnd2.TkinterDnD.Tk()
fenetre.title("Convertisseur PDF → DOCX")
fenetre.geometry("500x350")

label_fichier = tk.Label(fenetre, text="Glissez un fichier PDF ici", width=60, height=5, bg="#f0f0f0", relief="groove")
label_fichier.pack(pady=10)
label_fichier.drop_target_register(tkdnd2.DND_FILES)
label_fichier.dnd_bind('<<Drop>>', on_drop)

btn_choisir_dossier = tk.Button(fenetre, text="Choisir le dossier de sortie", command=choisir_dossier)
btn_choisir_dossier.pack(pady=10)

label_dossier = tk.Label(fenetre, text="Aucun dossier sélectionné", wraplength=400)
label_dossier.pack()

btn_convertir = tk.Button(fenetre, text="Convertir", command=lancer_conversion, bg="#4CAF50", fg="white", padx=10, pady=5)
btn_convertir.pack(pady=10)

label_progression = tk.Label(fenetre, text="", fg="green", font=("Helvetica", 12))

label_statut = tk.Label(fenetre, text="", fg="blue")
label_statut.pack(pady=10)

fenetre.mainloop()
