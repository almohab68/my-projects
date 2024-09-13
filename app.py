import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter

# Function to convert PDF to Word
def convert_pdf_to_word():
    pdf_file = filedialog.askopenfilename(
        title="Select PDF File", filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_file:
        return

    # Ask for save location for the output Word file
    word_file = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word documents", "*.docx")],
        title="Save as",
    )
    if not word_file:
          return

    try:
        # Convert PDF to Word
        cv = Converter(pdf_file)
        cv.convert(word_file, start=0, end=None)
        cv.close()

        messagebox.showinfo("Success", "PDF converted to Word successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("PDF to Word Converter")

# Set window size
root.geometry("300x150")

# Create a label
label = tk.Label(root, text="Convert PDF to Word", font=(14))
label.pack()

# Create a button for file selection
button = tk.Button(root, text="Select PDF and Convert", command=convert_pdf_to_word)
button.pack()

# Start the GUI loop
root.mainloop()
