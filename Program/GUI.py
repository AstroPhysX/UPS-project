import tkinter as tk
from tkinter import filedialog

class PDFUploader:
    def __init__(self, master):
        self.master = master
        master.title("PDF Uploader")

        self.file1_button = tk.Button(master, text="File 1", command=self.upload_file1)
        self.file1_button.pack()

        self.file2_button = tk.Button(master, text="File 2", command=self.upload_file2)
        self.file2_button.pack()

    def upload_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        print("File 1:", file_path)

    def upload_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        print("File 2:", file_path)

root = tk.Tk()
pdf_uploader = PDFUploader(root)
root.mainloop()