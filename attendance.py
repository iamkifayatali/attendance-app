import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def upload_file_1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file_1.delete(0, tk.END)
        entry_file_1.insert(0, file_path)

def upload_file_2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file_2.delete(0, tk.END)
        entry_file_2.insert(0, file_path)

def compare_files():
    file_1 = entry_file_1.get()
    file_2 = entry_file_2.get()
    
    if not file_1 or not file_2:
        messagebox.showerror("Error", "Please upload both files.")
        return
    
    try:
        df1 = pd.read_excel(file_1)
        df2 = pd.read_excel(file_2)
        
        # Find similarities and differences
        similarities = df1[df1.isin(df2)].dropna()
        differences = pd.concat([df1, df2]).drop_duplicates(keep=False)
        
        # Ask user where to save the results
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            with pd.ExcelWriter(save_path) as writer:
                similarities.to_excel(writer, sheet_name="Similarities", index=False)
                differences.to_excel(writer, sheet_name="Differences", index=False)
            
            messagebox.showinfo("Success", f"Comparison complete. Results saved to '{save_path}'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Create the main window
root = tk.Tk()
root.title("Excel File Comparator")

# Create and place widgets
tk.Label(root, text="human:").grid(row=0, column=0, padx=10, pady=10)
entry_file_1 = tk.Entry(root, width=50)
entry_file_1.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Upload", command=upload_file_1).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="machine:").grid(row=1, column=0, padx=10, pady=10)
entry_file_2 = tk.Entry(root, width=50)
entry_file_2.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Upload", command=upload_file_2).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Compare", command=compare_files).grid(row=2, column=1, pady=20)

# Run the application
root.mainloop()
#_______________________________________________________________________________________________