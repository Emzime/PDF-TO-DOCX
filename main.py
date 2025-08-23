import os
import threading
import tkinter as tk
import warnings
from tkinter import filedialog, messagebox, ttk

# noinspection SpellCheckingInspection
import tkinterdnd2 as tkdnd2
from PyPDF2 import PdfReader
from pdf2docx import Converter
from docx import Document
from docx.shared import Mm

# --- (Optionnel) filtrages d'avertissements ciblés ---
# (conservé au cas où d'autres libs émettent des warnings bruyants)
warnings.filterwarnings("ignore", category=UserWarning, module=r"docxcompose\.properties")

# --- Thème personnalisé (centralisé) ---
BG_COLOR = "#f7f9fb"  # fond général
PANEL_BG = "#dee5ec"  # zones encadrées
TEXT_COLOR = "#333"  # texte principal
TEXT_COLOR_SECONDARY = "#444"
TEXT_COLOR_MUTED = "#555"

ACCENT_COLOR = "#4CAF50"  # bouton "Convertir"
BUTTON_COLOR = "#2196F3"  # bouton "Sélection dossier"
BUTTON_HOVER_COLOR = "#1976D2"

FONT = ("Segoe UI", 11)  # police standard
TITLE_FONT = ("Segoe UI", 16, "bold")
SMALL_FONT = ("Segoe UI", 10)


# --- Helper thread-safe pour Tkinter ---
def _call_with_kwargs(f, a, kw):
    f(*a, **kw)


def ui_after(func, *args, **kwargs):
    if kwargs:
        window.after(0, _call_with_kwargs, func, args, kwargs)
    else:
        window.after(0, func, *args)


# --- Helpers UI ---
def update_status(message: str):
    ui_after(label_status.config, text=message)


def set_progress_indeterminate(on: bool):
    if on:
        progress_bar.config(mode='indeterminate')
        ui_after(progress_bar.start, 10)  # 10 = vitesse
    else:
        try:
            progress_bar.stop()
        except FileNotFoundError as e:
            update_status("❌ Fichier introuvable")
            ui_after(messagebox.showerror, "Erreur", f"Impossible d'ouvrir le fichier :\n{e}")

        except PermissionError as e:
            update_status("❌ Permission refusée")
            ui_after(messagebox.showerror, "Erreur", f"Permission refusée (fichier ou dossier protégé) :\n{e}")

        except OSError as e:
            update_status("❌ Erreur système")
            ui_after(messagebox.showerror, "Erreur", f"Erreur système (I/O) :\n{e}")

        except RuntimeError as e:
            update_status("❌ Erreur interne")
            ui_after(messagebox.showerror, "Erreur", f"Erreur interne pdf2docx :\n{e}")

        except ValueError as e:
            update_status("❌ Erreur de données")
            ui_after(messagebox.showerror, "Erreur", f"Erreur de données/format :\n{e}")

        progress_bar.config(mode='determinate')


def style_button(button, color):
    button.config(bg=color, fg="white", activebackground=BUTTON_HOVER_COLOR,
                  relief="flat", font=FONT, padx=10, pady=5)


# --- Utilitaires ---
def get_unique_filepath(folder: str, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    counter = 1
    candidate = filename
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base} ({counter}){ext}"
        counter += 1
    return os.path.join(folder, candidate)


# --- Conversion (UNE PASSE) ---
def convert_pdf_to_docx(path: str, destination: str):
    converter = None
    try:
        base_name = os.path.splitext(os.path.basename(path))[0]
        filename = f"{base_name}.docx"
        final_output = get_unique_filepath(destination, filename)

        update_status("🔄 Conversion en cours…")
        ui_after(btn_convert.config, state='disabled')
        ui_after(btn_select_folder.config, state='disabled')

        # Barre de progression indéterminée
        progress_bar.pack(pady=10)
        set_progress_indeterminate(True)

        # Lancement conversion
        converter = Converter(path)
        converter.convert(final_output)   # <-- UNE PASSE
        converter.close()
        converter = None

        # 🔧 Ajuste la taille de page Word au format PDF (évite les bulles/légendes hors page)
        adjust_docx_section_to_pdf(src_pdf_path=path, docx_path=final_output)

        update_status("✅ Conversion terminée")
        ui_after(messagebox.showinfo, "Succès", f"Conversion terminée :\n{final_output}")

    except FileNotFoundError as e:
        update_status("❌ Fichier introuvable")
        ui_after(messagebox.showerror, "Erreur", f"Impossible d'ouvrir le fichier :\n{e}")
    except PermissionError as e:
        update_status("❌ Permission refusée")
        ui_after(messagebox.showerror, "Erreur", f"Permission refusée (fichier ou dossier protégé) :\n{e}")
    except OSError as e:
        update_status("❌ Erreur système")
        ui_after(messagebox.showerror, "Erreur", f"Erreur système (I/O) :\n{e}")
    except RuntimeError as e:
        update_status("❌ Erreur interne")
        ui_after(messagebox.showerror, "Erreur", f"Erreur interne pdf2docx :\n{e}")
    except ValueError as e:
        update_status("❌ Erreur de données")
        ui_after(messagebox.showerror, "Erreur", f"Erreur de données/format :\n{e}")
    finally:
        if converter is not None:
            try:
                converter.close()
            except (OSError, RuntimeError):
                pass
        # Rétablir UI
        set_progress_indeterminate(False)
        ui_after(progress_bar.pack_forget)
        ui_after(btn_convert.config, state='normal')
        ui_after(btn_select_folder.config, state='normal')
        ui_after(update_convert_button_state)


# --- UI logic ---
pdf_path = ""
output_folder = ""


def update_convert_button_state():
    if pdf_path and output_folder:
        btn_convert.config(state="normal")
    else:
        btn_convert.config(state="disabled")


def start_conversion():
    if not pdf_path or not output_folder:
        messagebox.showwarning("Attention", "Veuillez sélectionner un fichier PDF et un dossier de destination.")
        return
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


def adjust_docx_section_to_pdf(src_pdf_path: str, docx_path: str):
    with open(src_pdf_path, "rb") as f:
        r = PdfReader(f)
        m = r.pages[0].mediabox
        width_pt = float(m.width)
        height_pt = float(m.height)

    width_mm = width_pt * 25.4 / 72.0
    height_mm = height_pt * 25.4 / 72.0

    doc = Document(docx_path)
    for section in doc.sections:
        section.page_width = Mm(width_mm)
        section.page_height = Mm(height_mm)
        # marges minimales pour laisser de l’espace aux flottants/callouts
        section.left_margin = Mm(5)
        section.right_margin = Mm(5)
        section.top_margin = Mm(5)
        section.bottom_margin = Mm(5)
    doc.save(docx_path)


# --- Interface principale ---
window = tkdnd2.TkinterDnD.Tk()
window.title("Convertisseur PDF → DOCX")
window.geometry("600x500")
window.configure(bg=BG_COLOR)

# Titre
label_title = tk.Label(
    window,
    text="Convertisseur PDF vers DOCX",
    font=TITLE_FONT,
    bg=BG_COLOR,
    fg=TEXT_COLOR
)
label_title.pack(pady=(20, 10))

# Zone de dépôt
label_file = tk.Label(
    window,
    text="📄 Déposez votre fichier PDF ici",
    width=60,
    height=5,
    bg=PANEL_BG,
    fg=TEXT_COLOR,
    relief="groove",
    font=FONT
)
label_file.pack(pady=10)
label_file.drop_target_register(tkdnd2.DND_FILES)
label_file.dnd_bind('<<Drop>>', on_drop)

# Dossier de sortie
btn_select_folder = tk.Button(window, text="📂 Sélectionner le dossier de sortie", command=select_output_folder)
btn_select_folder.pack(pady=(10, 5))
style_button(btn_select_folder, BUTTON_COLOR)

label_folder = tk.Label(
    window,
    text="Aucun dossier sélectionné",
    wraplength=500,
    bg=BG_COLOR,
    font=SMALL_FONT,
    fg=TEXT_COLOR_SECONDARY
)
label_folder.pack(pady=(0, 10))

# Bouton Convertir
btn_convert = tk.Button(window, text="🚀 Convertir", command=start_conversion, state="disabled")
btn_convert.pack(pady=20)
style_button(btn_convert, ACCENT_COLOR)

# Barre de progression (indéterminée pendant la conv.)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(window, variable=progress_var, maximum=100, length=300)
progress_bar.pack_forget()

# Statut
label_status = tk.Label(
    window,
    text="",
    fg=TEXT_COLOR_MUTED,
    font=SMALL_FONT,
    bg=BG_COLOR
)
label_status.pack(pady=5)

# Lancement
window.mainloop()
