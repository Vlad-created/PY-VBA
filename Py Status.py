import tkinter as tk

def show_status(All Regions):
    root = tk.Tk()
    root.title("Query Status")
    label = tk.Label(root, text=f"Running: {All Regions}")
    label.pack()
    root.after(2000, root.destroy)  # Auto-close after 2 seconds
    root.mainloop()

