import ctypes
import tkinter as tk
from tkinter import messagebox, ttk
import os

class OT:
    def __init__(self, ot_values, shims):
        self.OT_list = [round(float(value), 2) for value in ot_values]
        self.CT_list = []
        self.CS_list = {}
        self.goodCT = []
        self.spacers = []
        self.shim_combinations = []
        self.shims = shims

    def CT_calc(self):
        CT3 = round(sum(self.OT_list) * 0.25, 2)
        CT1 = round(CT3 - 0.02, 2)
        CT2 = round(CT3 - 0.01, 2)
        CT4 = round(CT3 + 0.01, 2)
        CT5 = round(CT3 + 0.02, 2)
        self.CT_list = [CT1, CT2, CT3, CT4, CT5]

    def CS_calc(self):
        for C in self.CT_list:
            self.CS_list[C] = []  # Initialize the list for each CT value
            for O in self.OT_list:
                if O == C:
                    return False  # Stop further calculations if there's an error
                CS = round(((O) / (O - C)) * 100, 2)
                self.CS_list[C].append(CS)
        return True

    def validate_CT(self):
        self.goodCT = []
        for C in self.CT_list:
            if all(391 < cs_value < 409 for cs_value in self.CS_list[C]):
                self.goodCT.append(C)

    def find_space(self):
        self.spacers = []
        for c in self.goodCT:
            spacer = round((c - 9.5), 2)
            self.spacers.append(spacer)

    def find_shims(self, target, shims, memo):
        if target in memo:
            return memo[target]
        if target == 0:
            return []
        if target < 0:
            return None

        best_combination = None
        for shim in shims:
            remainder = round(target - shim, 2)
            result = self.find_shims(remainder, shims, memo)
            if result is not None:
                combination = result + [shim]
                if best_combination is None or len(combination) < len(best_combination):
                    best_combination = combination

        memo[target] = best_combination
        return best_combination

    def shim_combo(self):
        self.shim_combinations = []
        memo = {}
        for idx, c in enumerate(self.goodCT):
            s = self.spacers[idx]
            best_shims = self.find_shims(s, self.shims, memo)
            if best_shims is not None:
                self.shim_combinations.append([c, s, best_shims])
        self.shim_combinations.sort(key=lambda x: (len(x[2]), abs(x[0] - sum(self.CT_list)/len(self.CT_list))))

def calculate_and_display():
    try:
        ot_values = [entry_ot1.get(), entry_ot2.get(), entry_ot3.get()]
        if not all(ot_values):
            messagebox.showerror("Error", "Please enter all values!")
            return

        ot_instance = OT(ot_values, shims_list)
        ot_instance.CT_calc()
        if not ot_instance.CS_calc():
            messagebox.showerror("Error", "Check OT values.")
            return
        ot_instance.validate_CT()
        if not ot_instance.goodCT:
            messagebox.showerror("Error", "No valid CT values found. Standardize samples.")
            return
        ot_instance.find_space()
        ot_instance.shim_combo()
        if not ot_instance.shim_combinations:
            messagebox.showerror("Error", "No valid shim combinations found.")
            return

        display_shim_combinations(ot_instance.shim_combinations)
    except ValueError:
        messagebox.showerror("Error", "Please enter valid OT values.")

def display_shim_combinations(combinations):
    def show_combination(index):
        if 0 <= index < len(combinations):
            ct_value.set(f"CT: {combinations[index][0]}")
            spacer_value.set(f"Spacer: {combinations[index][1]}")
            shims_value.set(f"Shims: {combinations[index][2]}")

    def next_combination():
        nonlocal current_index
        current_index = (current_index + 1) % len(combinations)
        show_combination(current_index)

    def prev_combination():
        nonlocal current_index
        current_index = (current_index - 1) % len(combinations)
        show_combination(current_index)

    current_index = 0
    show_combination(current_index)

    next_button.config(command=next_combination)
    prev_button.config(command=prev_combination)

    next_button.grid(row=9, column=1, pady=5)
    prev_button.grid(row=9, column=0, pady=5)

def handle_enter(event, next_entry):
    if event.widget.get() and next_entry:
        next_entry.focus_set()
    elif event.widget.get() and not next_entry:
        calculate_and_display()

def show_info():
    messagebox.showinfo("About", "This app was made by Michael Guerrero - Chemical Engineer Co-Op in July, 2024 for the Rubber Development Lab of Freudenberg NOK in Cleveland, Georgia.")

def add_shim():
    try:
        new_shim = round(float(shim_entry.get()), 2)
        if new_shim not in shims_list:
            if messagebox.askyesno("Confirm Add", f"Are you sure you want to add the shim: {new_shim}?"):
                shims_list.append(new_shim)
                shims_list.sort()
                update_shim_listbox()
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid shim value.")

def remove_shim():
    try:
        if shim_entry.get():
            remove_shim_value = round(float(shim_entry.get()), 2)
            if remove_shim_value in shims_list:
                if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove the shim: {remove_shim_value}?"):
                    shims_list.remove(remove_shim_value)
                    update_shim_listbox()
        elif shim_listbox.curselection():
            remove_shim_index = shim_listbox.curselection()[0]
            remove_shim_value = shims_list[remove_shim_index]
            if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove the shim: {remove_shim_value}?"):
                shims_list.pop(remove_shim_index)
                update_shim_listbox()
        else:
            messagebox.showerror("Error", "Please enter a shim value or select one from the list to remove.")
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid shim value.")

def update_shim_listbox():
    shim_listbox.delete(0, tk.END)
    for shim in shims_list:
        shim_listbox.insert(tk.END, shim)

# Set the application ID for Windows
myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

root = tk.Tk()
root.title("Compression Set Calculator")

# Set the window icon using relative path
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir, 'img.ico')
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)
else:
    print(f"Icon not found at: {icon_path}")

# Create the notebook (tabbed interface)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both')

# Create the main tab
main_frame = ttk.Frame(notebook)
notebook.add(main_frame, text='Compression Set')

main_frame.columnconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

tk.Label(main_frame, text="OT I:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
entry_ot1 = tk.Entry(main_frame)
entry_ot1.grid(row=0, column=1, padx=10, pady=5, sticky='w')

tk.Label(main_frame, text="OT II:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
entry_ot2 = tk.Entry(main_frame)
entry_ot2.grid(row=1, column=1, padx=10, pady=5, sticky='w')

tk.Label(main_frame, text="OT III:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
entry_ot3 = tk.Entry(main_frame)
entry_ot3.grid(row=2, column=1, padx=10, pady=5, sticky='w')

calculate_button = tk.Button(main_frame, text="Calculate", command=calculate_and_display)
calculate_button.grid(row=3, column=0, columnspan=2, pady=10)

ct_value = tk.StringVar()
spacer_value = tk.StringVar()
shims_value = tk.StringVar()

tk.Label(main_frame, textvariable=ct_value).grid(row=4, column=0, columnspan=2, pady=5)
tk.Label(main_frame, textvariable=spacer_value).grid(row=5, column=0, columnspan=2, pady=5)
tk.Label(main_frame, textvariable=shims_value).grid(row=6, column=0, columnspan=2, pady=5)

tk.Label(main_frame, text="This application conforms to ASTM D-395 standards.").grid(row=7, column=0, columnspan=2, pady=10)

next_button = tk.Button(main_frame, text="Next")
prev_button = tk.Button(main_frame, text="Previous")

# Bind Enter key to each entry widget
entry_ot1.bind("<Return>", lambda event: handle_enter(event, entry_ot2))
entry_ot2.bind("<Return>", lambda event: handle_enter(event, entry_ot3))
entry_ot3.bind("<Return>", lambda event: handle_enter(event, None))

# Create the shims management tab
shims_frame = ttk.Frame(notebook)
notebook.add(shims_frame, text='Manage Shims')

shims_frame.columnconfigure(0, weight=1)
shims_frame.columnconfigure(1, weight=1)

shim_list_var = tk.StringVar()

tk.Label(shims_frame, text="Available Shims:").grid(row=0, column=0, columnspan=2, pady=5)
shim_listbox = tk.Listbox(shims_frame, listvariable=shim_list_var, height=10)
shim_listbox.grid(row=1, column=0, columnspan=2, pady=5)

tk.Label(shims_frame, text="New Shim:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
shim_entry = tk.Entry(shims_frame)
shim_entry.grid(row=2, column=1, padx=10, pady=5, sticky='w')

add_shim_button = tk.Button(shims_frame, text="Add Shim", command=add_shim)
add_shim_button.grid(row=3, column=0, pady=5)

remove_shim_button = tk.Button(shims_frame, text="Remove Shim", command=remove_shim)
remove_shim_button.grid(row=3, column=1, pady=5)

# Initial list of shims
shims_list = [0.2, 0.25, 0.3, 0.31, 0.52, 0.56, 0.61, 0.66, 0.68, 0.75, 0.76, 0.87, 0.96, 1.07, 1.13, 1.2, 1.25, 1.28, 1.34, 1.42, 1.43, 1.44, 1.5, 1.55, 1.65, 1.68, 1.7, 1.73, 1.82, 1.86, 1.92, 1.97, 2.04, 2.13, 2.25, 2.37, 2.48, 2.65]

# Update the shim list display
update_shim_listbox()

# Information label below the shim list
info_label = tk.Label(root, text="This application conforms to ASTM D-395 standards.\n© 2024 Michael Guerrero - Chemical Engineer Co-Op", font=("Arial", 8))
info_label.pack(side='bottom', fill='x', pady=5)

# Place the "About" button at the right end
info_button = tk.Button(root, text="?", command=show_info)
info_button.place(relx=1.0, rely=0.0, anchor='ne')

root.mainloop()
