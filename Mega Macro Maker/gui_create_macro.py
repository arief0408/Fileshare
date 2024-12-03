import tkinter as tk
from tkinter import messagebox
import os

# Drag-and-drop helper variables
drag_data = {"index": None, "text": None}
edit_index = None

def on_drag_start(event):
    global drag_data
    drag_data["index"] = selected_list.index("@%s,%s" % (event.x, event.y))
    drag_data["text"] = selected_list.get(drag_data["index"])

def on_drag_motion(event):
    pass

def on_drag_release(event):
    global drag_data
    if drag_data["index"] is not None:
        new_index = selected_list.index("@%s,%s" % (event.x, event.y))
        if new_index != drag_data["index"]:
            selected_text = drag_data["text"]
            selected_list.delete(drag_data["index"])
            selected_list.insert(new_index, selected_text)
    drag_data = {"index": None, "text": None}

def generate_xml(ordered_values):
    if not ordered_values:
        messagebox.showwarning("Warning", "No items to generate XML.")
        return

    xml_output = "<HAScript name=\"create_proposal\" description=\"\" timeout=\"60000\" pausetime=\"300\" promptall=\"true\" blockinput=\"true\" author=\"\" creationdate=\"\" supressclearevents=\"false\" usevars=\"true\" ignorepauseforenhancedtn=\"true\" delayifnotenhancedtn=\"0\" ignorepausetimeforenhancedtn=\"true\" continueontimeout=\"false\">\n"
    
    hascript_templates = {
        "Create Policy": "Create_Policy_hascript.txt",
        "Create Investment": "Create_Investment_hascript.txt",
        "Create Fund": "option3_hascript.txt",
        "Create Loan": "option4_hascript.txt",
        "Create Savings": "option5_hascript.txt"
    }

    for index, item in enumerate(ordered_values):
        selected_option = item.split(":")[0].strip()
        template_file = hascript_templates.get(selected_option)

        if not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file not found: {template_file}")
            return

        with open(template_file, 'r') as file:
            hascript = file.read()

        values = item.split(":")[1].strip().split(",")
        values = [value.strip() for value in values]
        str_index = str(index+1)
        hascript = hascript.replace("{index}", str_index)

        placeholder_dict = {}
        for i, value in enumerate(values):
            placeholder_name = f"{selected_option.replace(' ', '_')}_{i+1}"
            placeholder_dict[f"{{{placeholder_name}}}"] = value

        item_value = f"{selected_option.replace(' ', '_')}_{index + 1}"
        placeholder_dict['{item_value}'] = item_value

        if index + 1 < len(ordered_values):
            next_item = ordered_values[index + 1].strip()
            next_option_value = next_item.split(":")[0].strip()
            next_option_item_value = f"{next_option_value.replace(' ', '_')}_{index + 2}"
            placeholder_dict['{next_option_item_value}'] = next_option_item_value
        else:
            placeholder_dict['{next_option_item_value}'] = "Table_Code_Maintenance_submenu_next_value"

        for placeholder, value in placeholder_dict.items():
            hascript = hascript.replace(placeholder, value)

        xml_output += hascript + "\n"

    xml_output += "</HAScript>"

    with open("output.xml", "w") as xml_file:
        xml_file.write(xml_output)

    messagebox.showinfo("Success", f"XML Hascripts generated for {len(ordered_values)} ordered values.")

def add_item_group(entries, option_name):
    global edit_index
    values = [entry.get().strip() for entry in entries]
    if not all(values):
        messagebox.showwarning("Warning", f"All fields for {option_name} must be filled in.")
        return

    item_value = f"{option_name}: " + ", ".join(values)
    
    if edit_index is None:
        selected_list.insert(tk.END, item_value)
    else:
        selected_list.delete(edit_index)
        selected_list.insert(edit_index, item_value)
        edit_index = None

def toggle_dropdown(option_name, edit_index=None):
    for frame, _ in frames.values():
        frame.pack_forget()

    if option_name:
        frame, entries = frames[option_name]
        frame.pack(after=buttons[option_name], fill='x', padx=5, pady=5)

    if edit_index is not None:
        existing_values = selected_list.get(edit_index).split(":")[1].split(",")
        for entry, value in zip(entries, existing_values):
            entry.delete(0, tk.END)
            entry.insert(0, value.strip())

def on_item_double_click(event):
    global edit_index
    selected_index = selected_list.curselection()
    if not selected_index:
        return

    edit_index = selected_index[0]
    selected_item = selected_list.get(edit_index)

    option_name, values = selected_item.split(":")
    values = [v.strip() for v in values.split(",")]

    toggle_dropdown(option_name.strip(), edit_index)

    frame, entries = frames[option_name.strip()]
    for entry, value in zip(entries, values):
        entry.delete(0, tk.END)
        entry.insert(0, value)

def delete_selected_item():
    selected_index = selected_list.curselection()
    if selected_index:
        selected_list.delete(selected_index)

def create_option_frame(root, option_name, input_labels):
    frame = tk.Frame(root)
    entries = []
    for label_text in input_labels:
        label = tk.Label(frame, text=label_text)
        label.pack(anchor='w')
        entry = tk.Entry(frame, width=30)
        entry.pack(anchor='w')
        entries.append(entry)

    add_button = tk.Button(frame, text=f"Add {option_name} Group",
                           command=lambda: add_item_group(entries, option_name))
    add_button.pack(pady=5)
    return frame, entries

# Main application window
root = tk.Tk()
root.title("Drag-and-Drop XML Generator with Custom Group Input")
root.geometry("800x600")

left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

selected_list = tk.Listbox(left_frame, selectmode=tk.SINGLE, width=50)
selected_list.pack(fill=tk.BOTH, expand=True)

selected_list.bind("<ButtonPress-1>", on_drag_start)
selected_list.bind("<B1-Motion>", on_drag_motion)
selected_list.bind("<ButtonRelease-1>", on_drag_release)
selected_list.bind("<Double-Button-1>", on_item_double_click)

delete_button = tk.Button(left_frame, text="Delete Selected", command=delete_selected_item)
delete_button.pack(pady=5)

right_frame_container = tk.Frame(root)
right_frame_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(right_frame_container)
scrollbar = tk.Scrollbar(right_frame_container, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

options = {
    "Create Policy": ["Sum Assured", "Product Name", "Billing Frequency"],
    "Create Investment": ["Coverage Amount", "Policy Name", "Payment Mode"],
    "Create Fund": ["Investment Amount", "Fund Name", "Investment Frequency"],
    "Create Loan": ["Loan Amount", "Loan Type", "Repayment Schedule"],
    "Create Savings": ["Savings Amount", "Plan Name", "Deposit Frequency"]
}

frames = {}
buttons = {}

for option_name, labels in options.items():
    dropdown_button = tk.Button(scrollable_frame, text=option_name,
                                command=lambda n=option_name: toggle_dropdown(n))
    dropdown_button.pack(pady=5, fill=tk.X)
    buttons[option_name] = dropdown_button

    frame, entries = create_option_frame(scrollable_frame, option_name, labels)
    frame.pack_forget()
    frames[option_name] = (frame, entries)

generate_button = tk.Button(root, text="Generate XML",
                            command=lambda: generate_xml([selected_list.get(i) for i in range(selected_list.size())]))
generate_button.pack(side=tk.BOTTOM, pady=10)

root.mainloop()
