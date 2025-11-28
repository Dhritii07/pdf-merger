import os
from tkinter import *
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger
from PIL import Image

def combine_pdfs_and_images(folder_path, file_order, output_pdf="combined.pdf"):
    merger = PdfMerger()
    temp_files = []

    for file in file_order:
        full_path = os.path.join(folder_path, file)

        if not os.path.exists(full_path):
            print(f"⚠ File not found: {file}")
            continue

        
        if file.lower().endswith(".pdf"):
            merger.append(full_path)

       
        elif file.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tiff")):
            try:
                image = Image.open(full_path).convert("RGB")
            except Exception as e:
                print(f"⚠ Error opening image {file}: {e}")
                continue

            temp_pdf = os.path.join(folder_path, f"_temp_{file}.pdf")
            image.save(temp_pdf)
            temp_files.append(temp_pdf)
            merger.append(temp_pdf)

        else:
            print(f"⚠ Skipping unsupported file: {file}")

    merger.write(output_pdf)
    merger.close()

    for temp in temp_files:
        try:
            os.remove(temp)
        except:
            pass

    messagebox.showinfo("Success", f"Combined PDF created:\n{output_pdf}")

class PDFImageMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF & Image Merger with Drag-and-Drop Order")

        self.folder_path = ""
        self.file_list = []

        Button(root, text="Select Folder", command=self.select_folder).pack(pady=5)

        self.listbox = Listbox(root, selectmode=SINGLE, width=60, height=18)
        self.listbox.pack(padx=10, pady=10)
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        Button(root, text="Move Up", command=self.move_up).pack(pady=2)
        Button(root, text="Move Down", command=self.move_down).pack(pady=2)

        Button(root, text="Create Combined PDF", command=self.create_pdf).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.folder_path = folder

        self.file_list = sorted(
            f for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f))
        )

        self.refresh_listbox()

    def on_drop(self, event):
        dropped_files = self.root.splitlist(event.data)

        for f in dropped_files:
            filename = os.path.basename(f)
            if filename not in self.file_list:
                self.file_list.append(filename)

        self.refresh_listbox()

    def move_up(self):
        idx = self.get_selected_index()
        if idx is None or idx == 0:
            return

        self.file_list[idx], self.file_list[idx - 1] = \
            self.file_list[idx - 1], self.file_list[idx]

        self.refresh_listbox()
        self.listbox.select_set(idx - 1)

    def move_down(self):
        idx = self.get_selected_index()
        if idx is None or idx == len(self.file_list) - 1:
            return

        self.file_list[idx], self.file_list[idx + 1] = \
            self.file_list[idx + 1], self.file_list[idx]

        self.refresh_listbox()
        self.listbox.select_set(idx + 1)

    def refresh_listbox(self):
        self.listbox.delete(0, END)
        for f in self.file_list:
            self.listbox.insert(END, f)

    def get_selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        return sel[0]

    def create_pdf(self):
        if not self.folder_path:
            messagebox.showerror("Error", "No folder selected.")
            return

        if not self.file_list:
            messagebox.showerror("Error", "No files selected.")
            return

        output_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Combined PDF As"
        )

        if not output_pdf:
            return

        combine_pdfs_and_images(self.folder_path, self.file_list, output_pdf)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    PDFImageMergerGUI(root)
    root.mainloop()
