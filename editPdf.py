import os
from tkinter import *
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def parse_page_ranges(range_str):
    pages = set()
    parts = [x.strip() for x in range_str.split(",") if x.strip()]

    for part in parts:
        if "-" in part:
            try:
                start, end = part.split("-")
                start, end = int(start), int(end)
                pages.update(range(start - 1, end))
            except:
                raise ValueError(f"Invalid range: {part}")
        else:
            try:
                pages.add(int(part) - 1)
            except:
                raise ValueError(f"Invalid number: {part}")

    return pages


def remove_pages(input_pdf, output_pdf, pages_to_remove_str):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    pages_to_remove = parse_page_ranges(pages_to_remove_str)
    total_pages = len(reader.pages)

    for i in range(total_pages):
        if i not in pages_to_remove:
            writer.add_page(reader.pages[i])

    with open(output_pdf, "wb") as f:
        writer.write(f)


class PDFPageRemoverGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Remover")

        self.pdf_path = ""
        self.total_pages = 0


        Button(root, text="Select PDF", command=self.select_pdf).pack(pady=8)

        self.label_pdf = Label(root, text="No PDF selected", fg="gray")
        self.label_pdf.pack()

        self.label_pages = Label(root, text="")
        self.label_pages.pack(pady=4)

        Label(root, text="Pages to remove (e.g. 1,3,5-7):").pack(pady=(15, 3))
        self.entry_pages = Entry(root, width=40)
        self.entry_pages.pack()

        Button(
            root,
            text="Remove Pages",
            command=self.remove_pages_clicked,
            bg="#d9534f",
            fg="white",
            width=20
        ).pack(pady=15)


    def select_pdf(self):
        pdf = filedialog.askopenfilename(
            filetypes=[("PDF Files", "*.pdf")],
            title="Select PDF"
        )

        if not pdf:
            return

        self.pdf_path = pdf
        self.label_pdf.config(text=os.path.basename(pdf), fg="black")

        try:
            reader = PdfReader(pdf)
            self.total_pages = len(reader.pages)
            self.label_pages.config(text=f"Total pages: {self.total_pages}")
        except Exception as e:
            messagebox.showerror("Error", f"Unable to open PDF:\n{e}")

    def remove_pages_clicked(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "No PDF selected.")
            return

        pages_str = self.entry_pages.get().strip()
        if not pages_str:
            messagebox.showerror("Error", "Enter page numbers to remove.")
            return

        try:
            pages_to_remove = parse_page_ranges(pages_str)
        except Exception as e:
            messagebox.showerror("Invalid input", str(e))
            return

        for p in pages_to_remove:
            if p < 0 or p >= self.total_pages:
                messagebox.showerror(
                    "Error",
                    f"Page {p+1} is out of range (PDF has {self.total_pages} pages)"
                )
                return

        output_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Save New PDF As"
        )

        if not output_pdf:
            return

        try:
            remove_pages(self.pdf_path, output_pdf, pages_str)
            messagebox.showinfo("Success", f"Saved new PDF:\n{output_pdf}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove pages:\n{e}")


if __name__ == "__main__":
    root = Tk()
    PDFPageRemoverGUI(root)
    root.mainloop()
