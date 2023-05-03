import tkinter as tk
import pandas as pd

# Read the Excel file
df = pd.read_excel("ip_addresses.xlsx")

# Create the Tkinter window
window = tk.Tk()
window.title("IP Addresses")

# Create a table to display the data
for i in range(len(df.index)):
    for j in range(len(df.columns)):
        cell_value = df.iloc[i, j]
        cell = tk.Label(window, text=cell_value, font=("Arial", 10), relief="solid")
        cell.grid(row=i, column=j, sticky="nsew")

        # Add a click event to the cell
        def on_click(event, cell=cell):
            print(cell.cget("text"))

        cell.bind("<Button-1>", on_click)

# Make the cells resizable
for i in range(len(df.index)):
    window.rowconfigure(i, weight=1)
for j in range(len(df.columns)):
    window.columnconfigure(j, weight=1)

# Start the Tkinter event loop
window.mainloop()
