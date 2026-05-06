import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    from src.validate_invoice import validate_invoice
except ImportError:
    import sys, os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from src.validate_invoice import validate_invoice

def run_validation():
    path = filedialog.askopenfilename(
        title="Factur-X Datei auswählen",
        filetypes=[("XML Dateien", "*.xml")]
    )
    if not path:
        return

    result = validate_invoice(path)
    report = result.to_html(path)

    output.delete("1.0", tk.END)
    output.insert(tk.END, str(result))
    messagebox.showinfo("Fertig", f"Prüfung abgeschlossen.\nHTML-Report: {report}")

root = tk.Tk()
root.title("Factur-X / EN16931 Prüftool")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

btn = tk.Button(frame, text="XML prüfen", command=run_validation, width=20)
btn.pack(pady=10)

output = scrolledtext.ScrolledText(frame, width=80, height=20)
output.pack()

root.mainloop()
