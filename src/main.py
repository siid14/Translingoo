import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from excel_processor import ExcelProcessor

class ExcelTranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Translator")
        self.processor = ExcelProcessor()
        
        # Configure the main window
        self.root.geometry("600x400")
        self.setup_ui()

    def setup_ui(self):
        # Create and pack widgets
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(expand=True, fill='both')

        # Input file selection
        tk.Label(frame, text="Input Excel File:").pack(anchor='w')
        self.input_path_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.input_path_var, width=50).pack(fill='x', pady=(0, 10))
        tk.Button(frame, text="Browse", command=self.browse_input).pack(anchor='w')

        # Output file selection
        tk.Label(frame, text="Output Excel File:").pack(anchor='w', pady=(20, 0))
        self.output_path_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.output_path_var, width=50).pack(fill='x', pady=(0, 10))
        tk.Button(frame, text="Browse", command=self.browse_output).pack(anchor='w')

        # Process button
        tk.Button(frame, text="Process File", command=self.process_file).pack(pady=20)

        # Status label
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(frame, textvariable=self.status_var).pack()

    def browse_input(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.input_path_var.set(filename)
            # Auto-generate output path
            output_path = str(Path(filename).parent / f"{Path(filename).stem}_translated{Path(filename).suffix}")
            self.output_path_var.set(output_path)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.output_path_var.set(filename)

    def process_file(self):
        input_path = self.input_path_var.get()
        output_path = self.output_path_var.get()

        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output files")
            return

        self.status_var.set("Processing...")
        self.root.update()

        # Process the file
        if not self.processor.load_excel(input_path):
            self.status_var.set("Error loading file")
            return

        if not self.processor.process_file():
            self.status_var.set("Error processing file")
            return

        if not self.processor.save_excel(output_path):
            self.status_var.set("Error saving file")
            return

        self.status_var.set("Processing complete!")
        messagebox.showinfo("Success", "File processed successfully!")

def main():
    root = tk.Tk()
    app = ExcelTranslatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 