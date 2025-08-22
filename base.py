import os
import csv
import subprocess
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from dataclasses import dataclass
import math

CONTACTS_FILE = "contacts.csv"
DRAG_THRESHOLD = 5  # Pixels after which drag is considered


def _norm_path(p: str) -> str:
    """Normalizes the folder path by removing extra spaces."""
    return p.strip()


@dataclass
class Contact:
    folder_path: str      # Path to folder in the form "Folder/Subfolder"
    name: str
    ammyadmin_id: str
    anydesk_id: str
    rustdesk_id: str
    notes: str


# ---------------------------------------------------------------------------
# Dialog for adding/editing a contact
# ---------------------------------------------------------------------------
class ContactDialog(simpledialog.Dialog):
    """Dialog window for entering/editing a contact."""

    def __init__(self, parent, title=None, folder="", contact=None, on_save=None):
        self.folder = folder
        self.contact = contact
        self.on_save = on_save
        super().__init__(parent, title)

    def body(self, master):
        # Name
        ttk.Label(master, text="Name:").grid(row=0, column=0, sticky="e")
        self.name_var = tk.StringVar(value=self.contact.name if self.contact else "")
        entry_name = ttk.Entry(master, textvariable=self.name_var, width=40)
        entry_name.grid(row=0, column=1, padx=5, pady=2)

        # AmmyAdmin
        ttk.Label(master, text="AmmyAdmin ID:").grid(row=1, column=0, sticky="e")
        self.ammy_var = tk.StringVar(value=self.contact.ammyadmin_id if self.contact else "")
        entry_ammy = ttk.Entry(master, textvariable=self.ammy_var, width=40)
        entry_ammy.grid(row=1, column=1, padx=5, pady=2)

        # AnyDesk
        ttk.Label(master, text="AnyDesk ID:").grid(row=2, column=0, sticky="e")
        self.anydesk_var = tk.StringVar(value=self.contact.anydesk_id if self.contact else "")
        entry_anydesk = ttk.Entry(master, textvariable=self.anydesk_var, width=40)
        entry_anydesk.grid(row=2, column=1, padx=5, pady=2)

        # RustDesk
        ttk.Label(master, text="RustDesk ID:").grid(row=3, column=0, sticky="e")
        self.rustdesk_var = tk.StringVar(value=self.contact.rustdesk_id if self.contact else "")
        entry_rustdesk = ttk.Entry(master, textvariable=self.rustdesk_var, width=40)
        entry_rustdesk.grid(row=3, column=1, padx=5, pady=2)

        # Notes
        ttk.Label(master, text="Notes:").grid(row=4, column=0, sticky="e")
        self.notes_var = tk.StringVar(value=self.contact.notes if self.contact else "")
        entry_notes = ttk.Entry(master, textvariable=self.notes_var, width=40)
        entry_notes.grid(row=4, column=1, padx=5, pady=2)

        # Folder (for information only)
        ttk.Label(master, text="Folder:").grid(row=5, column=0, sticky="e")
        folder_disp = self.folder if self.folder else "(root)"
        ttk.Label(master, text=folder_disp).grid(row=5, column=1, sticky="w")

        # Set focus to the Name field
        return entry_name

    def apply(self):
        """Saves the entered data and calls the callback."""
        name = self.name_var.get().strip()
        ammy = self.ammy_var.get().strip()
        anydesk = self.anydesk_var.get().strip()
        rustdesk = self.rustdesk_var.get().strip()
        notes = self.notes_var.get().strip()
        if not name:
            messagebox.showwarning(
                "Data Input", "Contact name is required.", parent=self
            )
            return
        if self.on_save:
            self.on_save(name, ammy, anydesk, rustdesk, notes)


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class ContactManagerApp(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("Contacts for Remote Access")
        self.geometry("960x600")

        self.contacts: list[Contact] = []
        self.folder_path_to_id: dict[str, str] = {}
        self.drag_data = None  # {'contact_idx': int, 'start_x': int, 'start_y': int, 'dragged': bool}
        self.current_folder = ""

        self.load_contacts()
        self.create_widgets()
        self.build_folder_tree()
        self.select_root()

    # ---------------------------------------------------------------------------
    # CSV loading/saving
    # ---------------------------------------------------------------------------
    def load_contacts(self):
        self.contacts.clear()
        if not os.path.exists(CONTACTS_FILE):
            with open(CONTACTS_FILE, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(
                    ["folder", "name", "ammyadmin_id", "anydesk_id", "rustdesk_id", "notes"]
                )
            return
        with open(CONTACTS_FILE, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                self.contacts.append(
                    Contact(
                        folder_path=_norm_path(row["folder"]),
                        name=row["name"],
                        ammyadmin_id=row.get("ammyadmin_id", ""),
                        anydesk_id=row.get("anydesk_id", ""),
                        rustdesk_id=row.get("rustdesk_id", ""),
                        notes=row.get("notes", ""),
                    )
                )

    def save_contacts(self):
        with open(CONTACTS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(
                ["folder", "name", "ammyadmin_id", "anydesk_id", "rustdesk_id", "notes"]
            )
            for c in self.contacts:
                writer.writerow(
                    [
                        c.folder_path,
                        c.name,
                        c.ammyadmin_id,
                        c.anydesk_id,
                        c.rustdesk_id,
                        c.notes,
                    ]
                )

    # ---------------------------------------------------------------------------
    # Building the GUI
    # ---------------------------------------------------------------------------
    def create_widgets(self):
        # Button panel
        btn_bar = ttk.Frame(self)
        btn_bar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Button(btn_bar, text="Add Folder", command=self.add_folder).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_bar, text="Add Contact", command=self.add_contact).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_bar, text="Edit", command=self.edit_selected).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(btn_bar, text="Delete", command=self.delete_selected).pack(
            side=tk.LEFT, padx=2
        )

        # Tree/table panel
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Folder tree (left side)
        folder_frame = ttk.Frame(paned, width=200)
        paned.add(folder_frame, weight=1)

        self.folder_tree = ttk.Treeview(folder_frame, show="tree")
        self.folder_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        sb_folder = ttk.Scrollbar(
            folder_frame, orient=tk.VERTICAL, command=self.folder_tree.yview
        )
        sb_folder.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.configure(yscrollcommand=sb_folder.set)

        # Contacts table (right side)
        contacts_frame = ttk.Frame(paned)
        paned.add(contacts_frame, weight=3)

        columns = ("Name", "AmmyAdmin ID", "AnyDesk ID", "RustDesk ID", "Notes")
        self.contacts_table = ttk.Treeview(
            contacts_frame, columns=columns, show="headings", selectmode="extended"
        )
        for col in columns:
            self.contacts_table.heading(col, text=col)
            self.contacts_table.column(col, width=150, anchor="w")
        self.contacts_table.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        sb_vert = ttk.Scrollbar(
            contacts_frame, orient=tk.VERTICAL, command=self.contacts_table.yview
        )
        sb_vert.pack(side=tk.RIGHT, fill=tk.Y)
        self.contacts_table.configure(yscrollcommand=sb_vert.set)

        sb_horiz = ttk.Scrollbar(
            contacts_frame, orient=tk.HORIZONTAL, command=self.contacts_table.xview
        )
        sb_horiz.pack(side=tk.BOTTOM, fill=tk.X)
        self.contacts_table.configure(xscrollcommand=sb_horiz.set)

        # Events
        self.contacts_table.bind("<ButtonPress-1>", self.on_start_drag)
        self.contacts_table.bind("<B1-Motion>", self.on_drag_motion)
        # Global handling of button release (for both tree and table)
        self.bind("<ButtonRelease-1>", self.on_end_drag)

        self.contacts_table.bind("<Double-1>", self.on_contact_double_click)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_select)

        # Highlighting and auto-expanding folders during drag-and-drop
        self.folder_tree.bind("<Motion>", self.on_folder_highlight)

    # ---------------------------------------------------------------------------
    # Folder tree
    # ---------------------------------------------------------------------------
    def build_folder_tree(self):
        self.folder_tree.delete(*self.folder_tree.get_children())
        # Root node
        self.folder_tree.insert("", "end", "__root__", text="Root", open=True)
        self.folder_path_to_id = {"": "__root__"}

        # Collect all unique paths from contacts
        paths = {c.folder_path for c in self.contacts}
        paths.add("")  # Ensure root is present

        for path in sorted(paths):
            if path == "":
                continue
            parts = path.strip("/").split("/")
            cur = ""
            parent_id = "__root__"
            for part in parts:
                cur = f"{cur}/{part}" if cur else part
                if cur not in self.folder_path_to_id:
                    node_id = cur
                    self.folder_tree.insert(parent_id, "end", node_id, text=part, open=False)
                    self.folder_path_to_id[cur] = node_id
                    parent_id = node_id
                else:
                    parent_id = self.folder_path_to_id[cur]

    def select_root(self):
        self.folder_tree.selection_set("__root__")
        self.on_folder_select(event=None)

    def on_folder_select(self, event):
        sel = self.folder_tree.selection()
        if not sel:
            self.current_folder = ""
        else:
            node_id = sel[0]
            self.current_folder = "" if node_id == "__root__" else node_id
        self.refresh_contacts_table()

    # ---------------------------------------------------------------------------
    # Contacts table
    # ---------------------------------------------------------------------------
    def refresh_contacts_table(self):
        """Displays only contacts in the current folder (including root)."""
        self.contacts_table.delete(*self.contacts_table.get_children())
        folder = self.current_folder
        for idx, contact in enumerate(self.contacts):
            if _norm_path(contact.folder_path) == _norm_path(folder):
                self.contacts_table.insert(
                    "",
                    "end",
                    iid=str(idx),
                    values=(
                        contact.name,
                        contact.ammyadmin_id,
                        contact.anydesk_id,
                        contact.rustdesk_id,
                        contact.notes,
                    ),
                )

    # ---------------------------------------------------------------------------
    # Operations on data (add/edit/delete)
    # ---------------------------------------------------------------------------
    def add_folder(self):
        sel = self.folder_tree.selection()
        parent_path = "" if not sel else ("" if sel[0] == "__root__" else sel[0])

        name = simpledialog.askstring(
            "Add Folder", "Folder name (no '/' or empty characters):", parent=self
        )
        if not name:
            return
        name = name.strip()
        if "/" in name:
            messagebox.showwarning("Folder Name", "Name should not contain '/'", parent=self)
            return

        new_path = f"{parent_path}/{name}" if parent_path else name
        if new_path in self.folder_path_to_id:
            messagebox.showwarning(
                "Folder Exists",
                f"Folder '{new_path}' already exists.",
                parent=self,
            )
            return

        parent_id = self.folder_path_to_id.get(parent_path, "__root__")
        self.folder_tree.insert(parent_id, "end", new_path, text=name, open=False)
        self.folder_path_to_id[new_path] = new_path

    def add_contact(self):
        ContactDialog(
            self,
            title="Add Contact",
            folder=self.current_folder,
            on_save=self._add_contact_callback,
        )

    def _add_contact_callback(self, name, ammy, anydesk, rustdesk, notes):
        self.contacts.append(
            Contact(
                folder_path=self.current_folder,
                name=name,
                ammyadmin_id=ammy,
                anydesk_id=anydesk,
                rustdesk_id=rustdesk,
                notes=notes,
            )
        )
        self.save_contacts()
        self.refresh_contacts_table()

    def edit_selected(self):
        sel = self.contacts_table.selection()
        if not sel:
            messagebox.showinfo("Edit", "Select a contact.", parent=self)
            return
        iid = sel[0]
        idx = int(iid)
        contact = self.contacts[idx]

        def on_save(name, ammy, anydesk, rustdesk, notes):
            contact.name = name
            contact.ammyadmin_id = ammy
            contact.anydesk_id = anydesk
            contact.rustdesk_id = rustdesk
            contact.notes = notes
            self.save_contacts()
            self.refresh_contacts_table()

        ContactDialog(
            self,
            title="Edit Contact",
            folder=contact.folder_path,
            contact=contact,
            on_save=on_save,
        )

    def delete_selected(self):
        # Delete contacts
        sel_contacts = self.contacts_table.selection()
        if sel_contacts:
            if not messagebox.askyesno(
                "Delete", "Delete selected contacts?", parent=self
            ):
                return
            indices = sorted([int(iid) for iid in sel_contacts], reverse=True)
            for i in indices:
                del self.contacts[i]
            self.save_contacts()
            self.refresh_contacts_table()
            return

        # Delete folder
        sel_folders = self.folder_tree.selection()
        if not sel_folders:
            messagebox.showinfo("Delete", "Select a folder or contact.", parent=self)
            return
        node_id = sel_folders[0]
        if node_id == "__root__":
            messagebox.showwarning("Delete", "Cannot delete root folder.", parent=self)
            return
        if not messagebox.askyesno(
            "Delete Folder",
            f"Delete folder '{self.folder_tree.item(node_id, 'text')}' and all its contents?",
            parent=self,
        ):
            return
        folder_path = node_id
        # Remove contacts from this folder and all subfolders
        self.contacts = [
            c for c in self.contacts if not c.folder_path.startswith(folder_path)
        ]
        # Remove tree node and its children
        self.folder_tree.delete(node_id)
        to_remove = [
            fp
            for fp in self.folder_path_to_id
            if fp == folder_path or fp.startswith(folder_path + "/")
        ]
        for fp in to_remove:
            del self.folder_path_to_id[fp]
        self.save_contacts()
        self.refresh_contacts_table()
        self.select_root()

    # ---------------------------------------------------------------------------
    # Helper method - check descendant
    # ---------------------------------------------------------------------------
    def _is_child_of(self, widget, container):
        """Returns True if widget is inside container (including container itself)."""
        while widget is not None:
            if widget == container:
                return True
            widget = widget.master
        return False

    # ---------------------------------------------------------------------------
    # Drag-and-Drop
    # ---------------------------------------------------------------------------
    def on_start_drag(self, event):
        """Records the initial position of the row the user is about to drag."""
        row_id = self.contacts_table.identify_row(event.y)
        if not row_id:
            self.drag_data = None
            return
        self.drag_data = {
            "contact_idx": int(row_id),
            "start_x": event.x_root,
            "start_y": event.y_root,
            "dragged": False,
        }
        self.contacts_table.configure(cursor="hand2")

    def on_drag_motion(self, event):
        """Tracks if the user has exceeded the threshold movement."""
        if not self.drag_data:
            return
        dx = event.x_root - self.drag_data["start_x"]
        dy = event.y_root - self.drag_data["start_y"]
        if not self.drag_data["dragged"] and math.hypot(dx, dy) >= DRAG_THRESHOLD:
            self.drag_data["dragged"] = True

        # Highlight folder under cursor
        if self.drag_data["dragged"]:
            x_root, y_root = self.winfo_pointerx(), self.winfo_pointery()
            widget_under = self.winfo_containing(x_root, y_root)
            if widget_under is self.folder_tree:
                rel_y = y_root - self.folder_tree.winfo_rooty()
                node = self.folder_tree.identify_row(rel_y)
                if node:
                    self.folder_tree.selection_set(node)

    def on_end_drag(self, event):
        """Handles mouse button release - moves folder or reorders."""
        if not self.drag_data:
            return

        # User just clicked - do nothing
        if not self.drag_data["dragged"]:
            self.contacts_table.configure(cursor="")
            self.drag_data = None
            return

        try:
            source_idx = self.drag_data["contact_idx"]
            source_contact = self.contacts[source_idx]

            # Cursor position in global coordinates
            x_root = self.winfo_pointerx()
            y_root = self.winfo_pointery()
            widget_under = self.winfo_containing(x_root, y_root)

            # Drop onto folder
            if self._is_child_of(widget_under, self.folder_tree):
                # Coordinates relative to tree
                rel_y = y_root - self.folder_tree.winfo_rooty()
                node_id = self.folder_tree.identify_row(rel_y)
                if not node_id:
                    node_id = "__root__"
                # Expand the node (useful for nested folders)
                self.folder_tree.item(node_id, open=True)

                new_folder = "" if node_id == "__root__" else node_id
                if source_contact.folder_path != new_folder:
                    source_contact.folder_path = new_folder
                    self.save_contacts()
                    # Switch to the new folder immediately to show the result
                    self.folder_tree.selection_set(node_id)
                    self.on_folder_select(event=None)
                # No further actions needed
                return

            # Drop onto table
            if self._is_child_of(widget_under, self.contacts_table):
                # Coordinates relative to table
                rel_y = y_root - self.contacts_table.winfo_rooty()
                target_row = self.contacts_table.identify_row(rel_y)

                if not target_row:
                    # Dropped on empty area - move to end of current folder
                    self._move_contact_to_end_of_folder(source_idx)
                else:
                    target_idx = int(target_row)
                    if target_idx != source_idx:
                        target_contact = self.contacts[target_idx]

                        if target_contact.folder_path == source_contact.folder_path:
                            # Reorder within the same folder
                            self._reorder_within_folder(source_idx, target_idx)
                        else:
                            # Move to another folder (after target contact)
                            self._move_contact_to_folder_at_position(
                                source_idx, target_contact.folder_path, target_idx
                            )
                self.save_contacts()
                self.refresh_contacts_table()

                # Select the moved contact if visible in the current folder
                new_idx = self.contacts.index(source_contact)
                if self.current_folder == source_contact.folder_path:
                    if self.contacts_table.exists(str(new_idx)):
                        self.contacts_table.selection_set(str(new_idx))

        except Exception as e:
            messagebox.showerror("Drag Error", str(e), parent=self)
        finally:
            self.contacts_table.configure(cursor="")
            self.drag_data = None

    # ---------------------------------------------------------------------------
    # Helper methods for drag-and-drop
    # ---------------------------------------------------------------------------
    def _move_contact_to_end_of_folder(self, source_idx: int):
        """Moves the contact to the end of its current folder."""
        contact = self.contacts.pop(source_idx)
        folder = contact.folder_path
        # Find the last element in the same folder
        last_idx = None
        for i, c in enumerate(self.contacts):
            if c.folder_path == folder:
                last_idx = i
        if last_idx is None:
            self.contacts.append(contact)
        else:
            self.contacts.insert(last_idx + 1, contact)

    def _reorder_within_folder(self, source_idx: int, target_idx: int):
        """Moves the contact within the same folder (preserving order)."""
        folder = self.contacts[source_idx].folder_path
        folder_indices = [i for i, c in enumerate(self.contacts) if c.folder_path == folder]
        src_pos = folder_indices.index(source_idx)
        tgt_pos = folder_indices.index(target_idx)

        folder_contacts = [self.contacts[i] for i in folder_indices]
        contact_obj = folder_contacts.pop(src_pos)
        folder_contacts.insert(tgt_pos, contact_obj)

        for local_i, global_i in enumerate(folder_indices):
            self.contacts[global_i] = folder_contacts[local_i]

    def _move_contact_to_folder_at_position(self, source_idx: int, target_folder: str, target_idx: int):
        """
        Moves the contact to another folder and places it right after the target contact.
        """
        contact = self.contacts.pop(source_idx)
        contact.folder_path = target_folder

        insert_pos = target_idx
        if source_idx < target_idx:
            insert_pos -= 1  # Adjustment due to the already removed element
        self.contacts.insert(insert_pos + 1, contact)

    # ---------------------------------------------------------------------------
    # Folder highlighting during drag (and auto-preview)
    # ---------------------------------------------------------------------------
    def on_folder_highlight(self, event):
        """During dragging, highlights (and opens) the folder under the cursor."""
        if not self.drag_data or not self.drag_data["dragged"]:
            return
        node = self.folder_tree.identify_row(event.y)
        if node:
            self.folder_tree.selection_set(node)
            self.folder_tree.item(node, open=True)

    # ---------------------------------------------------------------------------
    # Launching remote access program
    # ---------------------------------------------------------------------------
    def on_contact_double_click(self, event):
        row_id = self.contacts_table.identify_row(event.y)
        col_id = self.contacts_table.identify_column(event.x)
        if not row_id or not col_id:
            return
        idx = int(row_id)
        contact = self.contacts[idx]
        col_num = int(col_id.replace("#", ""))
        prog_map = {2: "ammyadmin", 3: "anydesk", 4: "rustdesk"}
        if col_num not in prog_map:
            return
        prog = prog_map[col_num]
        remote_id = getattr(contact, f"{prog}_id")
        if not remote_id:
            messagebox.showinfo(
                "No ID",
                f"Contact lacks {prog.upper()}-ID.",
                parent=self,
            )
            return
        self.launch_program(prog, remote_id)

    def find_executable(self, prog_name: str) -> str | None:
        """
        Searches for the executable file of the program.
        Search order:
        1) Directory where the script is located.
        2) Current working directory (where contacts.csv is).
        3) Standard Windows installation paths.
        4) System PATH (via shutil.which).
        """
        exe_name_map = {
            "ammyadmin": "AmmyyAdmin.exe",
            "anydesk": "AnyDesk.exe",
            "rustdesk": "rustdesk.exe",
        }
        exe_name = exe_name_map.get(prog_name)
        if not exe_name:
            return None

        # 1) Script directory
        script_dir = os.path.abspath(os.path.dirname(__file__))
        script_path = os.path.join(script_dir, exe_name)
        if os.path.isfile(script_path):
            return script_path

        # 2) Current working directory
        cwd_path = os.path.join(os.getcwd(), exe_name)
        if os.path.isfile(cwd_path):
            return cwd_path

        # 3) Standard installation locations
        candidates = {
            "ammyadmin": [
                r"C:\Program Files (x86)\Ammyy Admin\AmmyyAdmin.exe",
                r"C:\Program Files\Ammyy Admin\AmmyyAdmin.exe",
            ],
            "anydesk": [
                r"C:\Program Files (x86)\AnyDesk\AnyDesk.exe",
                r"C:\Program Files\AnyDesk\AnyDesk.exe",
            ],
            "rustdesk": [
                r"C:\Program Files\RustDesk\rustdesk.exe",
                r"C:\Program Files (x86)\RustDesk\rustdesk.exe",
            ],
        }.get(prog_name, [])

        for p in candidates:
            if os.path.isfile(p):
                return p

        # 4) Via PATH
        return shutil.which(exe_name)

    def launch_program(self, prog_name: str, remote_id: str):
        """
        Launches the selected remote access program with the appropriate parameters.
        """
        exe = self.find_executable(prog_name)
        if not exe:
            messagebox.showerror(
                "Error",
                f"Application {prog_name.title()} not found. Check installation.",
                parent=self,
            )
            return

        # Form the argument list depending on the program
        args = [exe]
        if prog_name == "ammyadmin":
            # AmmyyAdmin requires -connect flag
            args += ["-connect", remote_id]
        elif prog_name == "anydesk":
            # AnyDesk accepts just the ID
            args += [remote_id]
        elif prog_name == "rustdesk":
            # RustDesk requires --connect flag
            args += ["--connect", remote_id]
        else:
            # For potential new programs
            args += [remote_id]

        try:
            subprocess.Popen(args, shell=False)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to launch {prog_name.title()}: {e}",
                parent=self,
            )


if __name__ == "__main__":
    app = ContactManagerApp()
    app.mainloop()