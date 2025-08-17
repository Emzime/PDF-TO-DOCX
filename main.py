import os
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import tkinterdnd2 as tkdnd2
from PyPDF2 import PdfReader
from docx import Document
from pdf2docx import Converter

pdf_path = ""
output_folder = ""

# --- Thème personnalisé ---
BG_COLOR = "#f7f9fb"
ACCENT_COLOR = "#4CAF50"
BUTTON_COLOR = "#2196F3"
BUTTON_HOVER_COLOR = "#1976D2"
FONT = ("Segoe UI", 11)


def update_progress(percent):
    progress_var.set(percent * 100)
    window.update_idletasks()


def update_status(message):
    label_status.config(text=message)


def get_total_pages(path: str) -> int:
    with open(path, "rb") as f:
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
    for file in files:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save(output_path)


def convert_pdf_to_docx(path: str, destination: str):
    try:
        base_name = os.path.splitext(os.path.basename(path))[0]
        filename = f"{base_name}.docx"
        final_output = get_unique_filepath(destination, filename)

        update_status("🔄 Conversion en cours...")

        btn_convert.config(state='disabled')
        btn_select_folder.config(state='disabled')
        progress_var.set(0)

        total_pages = get_total_pages(path)
        converter = Converter(path)

        temp_dir = tempfile.mkdtemp()
        temp_files = []

        for page_number in range(total_pages):
            temp_file = os.path.join(temp_dir, f"page_{page_number + 1}.docx")
            converter.convert(temp_file, start=page_number, end=page_number + 1)
            temp_files.append(temp_file)
            percent = (page_number + 1) / total_pages
            update_progress(percent)

        converter.close()
        merge_docx(temp_files, final_output)
        shutil.rmtree(temp_dir)

        update_status("✅ Conversion terminée")
        messagebox.showinfo("Succès", f"Conversion terminée :\n{final_output}")

    except Exception as e:
        update_status("❌ Erreur lors de la conversion")
        messagebox.showerror("Erreur", f"Une erreur est survenue pendant la conversion :\n{str(e)}")

    finally:
        btn_convert.config(state='normal')
        btn_select_folder.config(state='normal')
        update_convert_button_state()
        progress_bar.pack_forget()  # Masquer la barre après conversion


def update_convert_button_state():
    if pdf_path and output_folder:
        btn_convert.config(state="normal")
    else:
        btn_convert.config(state="disabled")


def start_conversion():
    if not pdf_path or not output_folder:
        messagebox.showwarning("Attention", "Veuillez sélectionner un fichier PDF et un dossier de destination.")
        return
    progress_bar.pack(pady=10)  # Afficher la barre ici
    threading.Thread(target=convert_pdf_to_docx, args=(pdf_path, output_folder), daemon=True).start()


def on_drop(event):
    global pdf_path
    pdf_path = str(event.data).strip('{}')
    label_file.config(text=f"📄 {os.path.basename(pdf_path)}")
    update_convert_button_state()


def select_output_folder():
    global output_folder
    folder = filedialog.askdirectory()
    if folder:
        output_folder = str(folder)
        label_folder.config(text=f"📁 {folder}")
        update_convert_button_state()


def style_button(button, color):
    button.config(bg=color, fg="white", activebackground=BUTTON_HOVER_COLOR, relief="flat", font=FONT, padx=10, pady=5)


# --- Interface principale ---
window = tkdnd2.TkinterDnD.Tk()
window.title("Convertisseur PDF → DOCX")
window.geometry("600x500")
window.configure(bg=BG_COLOR)

# Titre
label_title = tk.Label(window, text="Convertisseur PDF vers DOCX", font=("Segoe UI", 16, "bold"), bg=BG_COLOR, fg="#333")
label_title.pack(pady=(20, 10))

# Zone de dépôt
label_file = tk.Label(window, text="📄 Déposez votre fichier PDF ici", width=60, height=5,
                      bg="#dee5ec", fg="#333", relief="groove", font=FONT)
label_file.pack(pady=10)
label_file.drop_target_register(tkdnd2.DND_FILES)
label_file.dnd_bind('<<Drop>>', on_drop)

# Dossier de sortie
btn_select_folder = tk.Button(window, text="📂 Sélectionner le dossier de sortie", command=select_output_folder)
btn_select_folder.pack(pady=(10, 5))
style_button(btn_select_folder, BUTTON_COLOR)

label_folder = tk.Label(window, text="Aucun dossier sélectionné", wraplength=500,
                        bg=BG_COLOR, font=("Segoe UI", 10), fg="#444")
label_folder.pack(pady=(0, 10))

# Bouton Convertir
btn_convert = tk.Button(window, text="🚀 Convertir", command=start_conversion, state="disabled")
btn_convert.pack(pady=20)
style_button(btn_convert, ACCENT_COLOR)

# Barre de progression (initialement masquée)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(window, variable=progress_var, maximum=100, length=400)
progress_bar.pack_forget()

# Statut
label_status = tk.Label(window, text="", fg="#555", font=("Segoe UI", 10), bg=BG_COLOR)
label_status.pack(pady=5)

# Lancement
window.mainloop()
