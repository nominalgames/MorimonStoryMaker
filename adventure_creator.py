import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape

import openpyxl
from docx import Document
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(PROJECT_DIR, "data.xlsx")
TEMPLATE_FILE = os.path.join(PROJECT_DIR, "Adventure_Template.docx")
OPENING_PLACEHOLDER = "[INSERT_OPENING]"
PLAYER_NAME_PLACEHOLDER = "[Player_Name]"

CHARACTER_TYPES = ["Catcher"]


def get_starting_location():
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    ws = wb["Locations"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        location_id, location_name = row[0], row[1]
        if location_id == 1:
            return location_name
    raise ValueError("No location found with Location ID 1")


def get_random_opening(player_type):
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    ws = wb["Openings"]
    matches = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        opening_text, opening_player_type = row[1], row[3]
        if opening_player_type == player_type:
            matches.append(opening_text)
    if not matches:
        raise ValueError(f"No openings found for Player_Type '{player_type}'")
    return random.choice(matches)


def create_adventure_pdf(opening_text, player_name, pdf_path):
    personalized_opening = opening_text.replace(PLAYER_NAME_PLACEHOLDER, player_name)

    template = Document(TEMPLATE_FILE)
    section = template.sections[0]

    styles = getSampleStyleSheet()
    story = []
    for paragraph in template.paragraphs:
        text = paragraph.text.replace(OPENING_PLACEHOLDER, personalized_opening)
        if not text.strip():
            story.append(Spacer(1, 12))
            continue
        story.append(Paragraph(escape(text), styles["Normal"]))
        story.append(Spacer(1, 12))

    pdf_doc = SimpleDocTemplate(
        pdf_path,
        pagesize=(section.page_width.pt, section.page_height.pt),
        leftMargin=section.left_margin.pt,
        rightMargin=section.right_margin.pt,
        topMargin=section.top_margin.pt,
        bottomMargin=section.bottom_margin.pt,
    )
    pdf_doc.build(story)


class AdventureCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Morimon Story Maker")
        self.root.geometry("400x200")

        self.menu_frame = tk.Frame(self.root)
        self.form_frame = tk.Frame(self.root)

        self.show_menu()

    def show_menu(self):
        self.form_frame.pack_forget()
        for widget in self.menu_frame.winfo_children():
            widget.destroy()

        create_button = tk.Button(
            self.menu_frame,
            text="Create New Adventure",
            command=self.show_new_adventure_form,
        )
        create_button.pack(pady=20)

        self.menu_frame.pack(fill="both", expand=True)

    def show_new_adventure_form(self):
        self.menu_frame.pack_forget()
        for widget in self.form_frame.winfo_children():
            widget.destroy()

        starting_location = get_starting_location()

        tk.Label(self.form_frame, text="Character Type").grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )
        self.character_type_var = tk.StringVar(value=CHARACTER_TYPES[0])
        character_type_dropdown = ttk.Combobox(
            self.form_frame,
            textvariable=self.character_type_var,
            values=CHARACTER_TYPES,
            state="readonly",
        )
        character_type_dropdown.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.form_frame, text="Starting Location").grid(
            row=1, column=0, sticky="w", padx=10, pady=10
        )
        self.starting_location_var = tk.StringVar(value=starting_location)
        starting_location_dropdown = ttk.Combobox(
            self.form_frame,
            textvariable=self.starting_location_var,
            values=[starting_location],
            state="readonly",
        )
        starting_location_dropdown.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.form_frame, text="Player Name").grid(
            row=2, column=0, sticky="w", padx=10, pady=10
        )
        self.player_name_var = tk.StringVar()
        player_name_entry = tk.Entry(self.form_frame, textvariable=self.player_name_var)
        player_name_entry.grid(row=2, column=1, padx=10, pady=10)

        button_frame = tk.Frame(self.form_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)

        tk.Button(button_frame, text="Create Adventure", command=self.submit_form).pack(
            side="left", padx=5
        )
        tk.Button(button_frame, text="Back", command=self.show_menu).pack(
            side="left", padx=5
        )

        self.form_frame.pack(fill="both", expand=True)

    def submit_form(self):
        Player_Name = self.player_name_var.get().strip()

        if not Player_Name:
            messagebox.showerror("Missing Player Name", "Please enter a player name.")
            return

        Player_Type = self.character_type_var.get()
        starting_location = self.starting_location_var.get()

        try:
            opening_text = get_random_opening(Player_Type)
        except ValueError as exc:
            messagebox.showerror("No Opening Found", str(exc))
            return

        pdf_path = filedialog.asksaveasfilename(
            title="Save Adventure As",
            defaultextension=".pdf",
            initialfile=f"{Player_Name}_Adventure.pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            return

        try:
            create_adventure_pdf(opening_text, Player_Name, pdf_path)
        except Exception as exc:
            messagebox.showerror("Failed to Create Adventure", str(exc))
            return

        messagebox.showinfo(
            "Adventure Created",
            f"Player Name: {Player_Name}\n"
            f"Character Type: {Player_Type}\n"
            f"Starting Location: {starting_location}\n\n"
            f"Saved to: {pdf_path}",
        )


def main():
    root = tk.Tk()
    AdventureCreatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
