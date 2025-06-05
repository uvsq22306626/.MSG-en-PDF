import os
import tkinter as tk
from tkinter import filedialog, messagebox
import extract_msg
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import re
import html

def clean_html(raw_html):
    clean = re.sub(r'<[^>]+>', '', raw_html)  # supprimer les balises
    clean = html.unescape(clean)              # convertir les entités HTML
    return clean.strip()

def get_msg_body(msg):
    # Tente de récupérer du texte brut ou HTML
    raw_body = msg.body or msg.htmlBody or b"(Message vide)"
    
    # Si c'est des bytes, on essaie de décoder
    if isinstance(raw_body, bytes):
        try:
            raw_body = raw_body.decode("utf-8")
        except UnicodeDecodeError:
            try:
                raw_body = raw_body.decode("iso-8859-1")
            except Exception:
                raw_body = "(Contenu non décodable)"
    
    return clean_html(raw_body)

def wrap_text(text, max_width, canvas_obj, font_name="Helvetica", font_size=11):
    """Découpe une ligne trop longue en plusieurs lignes selon la largeur maximale autorisée"""
    canvas_obj.setFont(font_name, font_size)
    words = text.split()
    lines = []
    line = ""

    for word in words:
        test_line = f"{line} {word}".strip()
        if canvas_obj.stringWidth(test_line) <= max_width:
            line = test_line
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines
def convert_msg_to_pdf(msg_path, output_path):
    try:
        msg = extract_msg.Message(msg_path)
        msg_subject = msg.subject or "Sans sujet"
        msg_date = msg.date or "Date inconnue"
        msg_sender = msg.sender or "Expéditeur inconnu"
        msg_body = get_msg_body(msg)

        filename = os.path.splitext(os.path.basename(msg_path))[0]
        pdf_file = os.path.join(output_path, filename + ".pdf")

        c = canvas.Canvas(pdf_file, pagesize=A4)
        width, height = A4
        left_margin = 40
        top_margin = 50
        max_width = width - 2 * left_margin

        y = height - top_margin
        line_height = 14
        c.setFont("Helvetica", 11)

        def draw_line(text):
            nonlocal y
            if y < 40:  # nouvelle page si plus de place
                c.showPage()
                y = height - top_margin
                c.setFont("Helvetica", 11)
            c.drawString(left_margin, y, text)
            y -= line_height

        draw_line(f"Sujet : {msg_subject}")
        draw_line(f"Date : {msg_date}")
        draw_line(f"De : {msg_sender}")
        draw_line("-" * 80)

        for line in msg_body.splitlines():
            for wrapped_line in wrap_text(line, max_width, c):
                draw_line(wrapped_line)

        c.save()
        return True

    except Exception as e:
        print(f"Erreur lors de la conversion de {msg_path} : {e}")
        return False

def select_and_convert():
    msg_files = filedialog.askopenfilenames(
        title="Sélectionner des fichiers MSG",
        filetypes=[("Fichiers MSG", "*.msg")]
    )
    if not msg_files:
        return

    output_dir = filedialog.askdirectory(title="Sélectionner un dossier de sortie")
    if not output_dir:
        return

    success = 0
    for msg_file in msg_files:
        if convert_msg_to_pdf(msg_file, output_dir):
            success += 1

    messagebox.showinfo("Conversion terminée", f"{success}/{len(msg_files)} fichiers convertis avec succès.")

def main():
    root = tk.Tk()
    root.title("Convertisseur MSG → PDF")
    root.geometry("420x160")

    label = tk.Label(root, text="Convertir des fichiers MSG en PDF", font=("Arial", 14))
    label.pack(pady=10)

    btn = tk.Button(root, text="Sélectionner les fichiers MSG", command=select_and_convert)
    btn.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
