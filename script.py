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
            print(f"⚠️ File not found: {file}")
            continue

        if file.lower().endswith(".pdf"):
            merger.append(full_path)
        else:
            # Convert image to temporary PDF
            image = Image.open(full_path).convert("RGB")
            temp_pdf = os.path.join(folder_path, f"_temp_{file}.pdf")
            image.save(temp_pdf)
            temp_files.append(temp_pdf)
            merger.append(temp_pdf)

    merger.write(output_pdf)
    merger.close()

    # Cleanup temporary PDFs
    for temp in temp_files:
        os.remove(temp)

    messagebox.showinfo("Success", f"Combined PDF created:\n{output_pdf}")

# ---------------- GUI ---------------- #

class PDFImageMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF & Image Merger with Drag-and-Drop Order")

        self.folder_path = ""
        self.file_list = []

        # Folder selection button
        Button(root, text="Select Folder", command=self.select_folder).pack(pady=5)

        # Drag-and-drop file listbox
        self.listbox = Listbox(root, selectmode=SINGLE, width=60, height=18)
        self.listbox.pack(padx=10, pady=10)
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self.on_drop)

        # Buttons
        Button(root, text="Move Up", command=self.move_up).pack(pady=2)
        Button(root, text="Move Down", command=self.move_down).pack(pady=2)

        Button(root, text="Create Combined PDF", command=self.create_pdf).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.folder_path = folder
        self.file_list = sorted(os.listdir(folder))

        self.listbox.delete(0, END)
        for f in self.file_list:
            self.listbox.insert(END, f)

    # Drag files into listbox
    def on_drop(self, event):
        files = self.root.splitlist(event.data)
        for f in files:
            filename = os.path.basename(f)
            if filename not in self.file_list:
                self.file_list.append(filename)
                self.listbox.insert(END, filename)

    # Move item up
    def move_up(self):
        index = self.listbox.curselection()
        if not index or index[0] == 0:
            return
        idx = index[0]

        self.file_list[idx], self.file_list[idx - 1] = self.file_list[idx - 1], self.file_list[idx]

        # Refresh listbox
        self.update_listbox()
        self.listbox.select_set(idx - 1)

    # Move item down
    def move_down(self):
        index = self.listbox.curselection()
        if not index or index[0] == len(self.file_list) - 1:
            return
        idx = index[0]

        self.file_list[idx], self.file_list[idx + 1] = self.file_list[idx + 1], self.file_list[idx]

        self.update_listbox()
        self.listbox.select_set(idx + 1)

    def update_listbox(self):
        self.listbox.delete(0, END)
        for f in self.file_list:
            self.listbox.insert(END, f)

    def create_pdf(self):
        if not self.folder_path:
            messagebox.showerror("Error", "No folder selected.")
            return

        if not self.file_list:
            messagebox.showerror("Error", "File list is empty.")
            return

        output_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Combined PDF As"
        )

        if not output_pdf:
            return

        combine_pdfs_and_images(self.folder_path, self.file_list, output_pdf)


# Run GUI
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    PDFImageMergerGUI(root)
    root.mainloop()
