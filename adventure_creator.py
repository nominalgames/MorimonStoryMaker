import os
import random
import subprocess
import sys
from tkinter import filedialog, messagebox

import customtkinter as ctk
import openpyxl

from location_reveal import get_location_reveal
from main_quest import choose_main_quest
from pdf_builder import build_adventure_pdf

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(PROJECT_DIR, "data.xlsx")
IMAGES_DIR = os.path.join(PROJECT_DIR, "images")

CHARACTER_TYPES = ["Catcher"]

# ---- UI theme: matches the generated PDF's own palette/type (pdf_builder.py) ----
COLOR_BG = "#FAF9F7"
COLOR_CARD = "#F2F1EF"
COLOR_BORDER = "#DAD6CF"
COLOR_FIELD_BG = "#FFFFFF"
COLOR_DARK = "#2A221C"
COLOR_BODY = "#4A4642"
COLOR_BLUE = "#1B5E8C"
COLOR_BLUE_HOVER = "#154A6E"
COLOR_WHITE = "#FFFFFF"

ctk.set_appearance_mode("light")


def _font(family, size, weight="normal", slant="roman"):
    return ctk.CTkFont(family=family, size=size, weight=weight, slant=slant)


def open_file(path):
    """Open `path` in the OS's default viewer (e.g. the PDF opens in whatever
    reader is set up for it). Best-effort: if nothing is registered to open it,
    the file is still saved, so this stays quiet rather than raising an error."""
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa: S606 (Windows only; not exec of user input)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except OSError:
        pass


def get_starting_location():
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    ws = wb["Locations"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        location_id, location_name = row[0], row[1]
        if location_id == 1:
            return location_name
    raise ValueError("No location found with Location ID 1")


def get_location_id(location_name):
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    ws = wb["Locations"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] == location_name:
            return int(row[0])
    raise ValueError(f"No location named '{location_name}' on the Locations tab")


def get_location_image_path(location_id):
    """images/location_<ID>.png for the starting location's Location ID."""
    path = os.path.join(IMAGES_DIR, f"location_{location_id}.png")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Missing location image: {path}")
    return path


def get_random_opening(player_type):
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    ws = wb["Openings"]
    matches = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 4:  # short/blank rows at the end of the sheet
            continue
        opening_text, opening_player_type = row[1], row[3]
        if opening_text and opening_player_type == player_type:
            matches.append(opening_text)
    if not matches:
        raise ValueError(f"No openings found for Player_Type '{player_type}'")
    return random.choice(matches)


def create_adventure(opening_text, player_name, location_id, image_path, output_path, main_quest=None):
    """Write the finished adventure PDF (built directly in Python, see pdf_builder.py)."""
    build_adventure_pdf(opening_text, player_name, location_id, image_path, output_path, main_quest)


class AdventureCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Morimon Story Maker")
        self.root.geometry("460x560")
        self.root.resizable(False, False)
        self.root.configure(fg_color=COLOR_BG)

        self.container = ctk.CTkFrame(self.root, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=36, pady=36)

        self.show_menu()

    def _clear(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def _header(self, eyebrow, title, subtitle=None):
        ctk.CTkLabel(
            self.container,
            text=" ".join(eyebrow.upper()),  # tracked caps, like the PDF's own eyebrow labels
            text_color=COLOR_BLUE,
            font=_font("Consolas", 12, weight="bold"),
            anchor="w",
        ).pack(fill="x")
        ctk.CTkLabel(
            self.container,
            text=title,
            text_color=COLOR_DARK,
            font=_font("Georgia", 28, weight="bold"),
            anchor="w",
        ).pack(fill="x", pady=(4, 0))
        if subtitle:
            ctk.CTkLabel(
                self.container,
                text=subtitle,
                text_color=COLOR_BODY,
                font=_font("Cambria", 13, slant="italic"),
                anchor="w",
                justify="left",
                wraplength=380,
            ).pack(fill="x", pady=(8, 0))

    def _field_label(self, parent, text):
        ctk.CTkLabel(
            parent,
            text=" ".join(text.upper()),
            text_color=COLOR_BLUE,
            font=_font("Consolas", 11, weight="bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 6))

    def show_menu(self):
        self._clear()
        self._header(
            "A tabletop adventure maker",
            "Morimon Story Maker",
            "Create a personalized adventure PDF to kick off your next Morimon journey.",
        )
        ctk.CTkButton(
            self.container,
            text="Create New Adventure",
            command=self.show_new_adventure_form,
            font=_font("Georgia", 16, weight="bold"),
            fg_color=COLOR_BLUE,
            hover_color=COLOR_BLUE_HOVER,
            text_color=COLOR_WHITE,
            corner_radius=10,
            height=50,
        ).pack(fill="x", pady=(32, 0))

    def show_new_adventure_form(self):
        self._clear()
        starting_location = get_starting_location()

        self._header("New adventure", "Create Your Adventure")

        card = ctk.CTkFrame(
            self.container,
            fg_color=COLOR_CARD,
            corner_radius=14,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        card.pack(fill="x", pady=(28, 0))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=24)

        field_kwargs = dict(
            font=_font("Cambria", 13),
            fg_color=COLOR_FIELD_BG,
            text_color=COLOR_DARK,
            corner_radius=8,
            height=38,
        )

        self._field_label(inner, "Character Type")
        self.character_type_var = ctk.StringVar(value=CHARACTER_TYPES[0])
        ctk.CTkOptionMenu(
            inner,
            values=CHARACTER_TYPES,
            variable=self.character_type_var,
            button_color=COLOR_BLUE,
            button_hover_color=COLOR_BLUE_HOVER,
            dropdown_fg_color=COLOR_FIELD_BG,
            dropdown_text_color=COLOR_DARK,
            **field_kwargs,
        ).pack(fill="x", pady=(0, 18))

        self._field_label(inner, "Starting Location")
        self.starting_location_var = ctk.StringVar(value=starting_location)
        ctk.CTkOptionMenu(
            inner,
            values=[starting_location],
            variable=self.starting_location_var,
            button_color=COLOR_BLUE,
            button_hover_color=COLOR_BLUE_HOVER,
            dropdown_fg_color=COLOR_FIELD_BG,
            dropdown_text_color=COLOR_DARK,
            **field_kwargs,
        ).pack(fill="x", pady=(0, 18))

        self._field_label(inner, "Player Name")
        self.player_name_var = ctk.StringVar()
        ctk.CTkEntry(
            inner,
            textvariable=self.player_name_var,
            placeholder_text="Enter your name",
            placeholder_text_color=COLOR_BODY,
            border_width=1,
            border_color=COLOR_BORDER,
            **field_kwargs,
        ).pack(fill="x")

        button_row = ctk.CTkFrame(self.container, fg_color="transparent")
        button_row.pack(fill="x", pady=(28, 0))
        ctk.CTkButton(
            button_row,
            text="Back",
            command=self.show_menu,
            font=_font("Georgia", 14),
            fg_color="transparent",
            hover_color=COLOR_CARD,
            text_color=COLOR_BODY,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=10,
            height=44,
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkButton(
            button_row,
            text="Create Adventure",
            command=self.submit_form,
            font=_font("Georgia", 14, weight="bold"),
            fg_color=COLOR_BLUE,
            hover_color=COLOR_BLUE_HOVER,
            text_color=COLOR_WHITE,
            corner_radius=10,
            height=44,
        ).pack(side="left", expand=True, fill="x", padx=(8, 0))

        self.form_frame.pack(fill="both", expand=True)

    def submit_form(self):
        Player_Name = self.player_name_var.get().strip()

        if not Player_Name:
            messagebox.showerror("Missing Player Name", "Please enter a player name.")
            return

        Player_Type = self.character_type_var.get()
        starting_location = self.starting_location_var.get()

        # runs right away, before the opening/location are even looked up: picks
        # one "Main" questline for this playthrough (or None if none is available)
        main_quest = choose_main_quest(Player_Type)

        try:
            opening_text = get_random_opening(Player_Type)
        except ValueError as exc:
            messagebox.showerror("No Opening Found", str(exc))
            return

        try:
            location_id = get_location_id(starting_location)
            image_path = get_location_image_path(location_id)
            get_location_reveal(location_id)  # fail early if the Location_<ID> tab is missing
        except (ValueError, FileNotFoundError) as exc:
            messagebox.showerror("Location Data Not Found", str(exc))
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Adventure As",
            defaultextension=".pdf",
            initialfile=f"{Player_Name}_Adventure.pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_path:
            return

        try:
            create_adventure(opening_text, Player_Name, location_id, image_path, output_path, main_quest)
        except Exception as exc:
            messagebox.showerror("Failed to Create Adventure", str(exc))
            return

        messagebox.showinfo(
            "Adventure Created",
            f"Player Name: {Player_Name}\n"
            f"Character Type: {Player_Type}\n"
            f"Starting Location: {starting_location}\n\n"
            f"Saved to: {output_path}",
        )
        open_file(output_path)


def main():
    root = ctk.CTk()
    AdventureCreatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
