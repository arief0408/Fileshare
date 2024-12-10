import tkinter as tk
from tkinter import messagebox
import os
import tkinter.filedialog as filedialog

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

def select_save_path():
    folder = filedialog.askdirectory(initialdir=save_path.get(), title="Select Save Folder")
    if folder:
        save_path.set(folder)

def generate_xml(ordered_values):
    """Generate XML based on the ordered values and save it to the selected folder."""
    if not ordered_values:
        messagebox.showwarning("Warning", "No items to generate XML.")
        return

    xml_output = "<HAScript name=\"create_proposal\" description=\"\" timeout=\"60000\" pausetime=\"300\" promptall=\"true\" blockinput=\"true\" author=\"\" creationdate=\"\" supressclearevents=\"false\" usevars=\"true\" ignorepauseforenhancedtn=\"true\" delayifnotenhancedtn=\"0\" ignorepausetimeforenhancedtn=\"true\" continueontimeout=\"false\">\n"
    
    # Define the templates for each option
    hascript_templates = {
        "Create Policy": "Create_Policy_hascript.txt",
        "Create Client": "Create_Client_hascript.txt",
        "Renewal 1Y": "Renewal_hascript.txt",
        "Cancellation": "Cancel_hascript.txt",
        "Loading Changes": "Comp_Changes_Loading_hascript.txt"
    }

    for index, item in enumerate(ordered_values):
        selected_option = item.split(":")[0].strip()  # Extract the option name
        template_file = hascript_templates.get(selected_option)

        if not template_file or not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file not found for {selected_option}: {template_file}")
            return

        # Read the template file content
        with open(template_file, 'r') as file:
            hascript = file.read()

        # Extract the values from the item
        values = item.split(":")[1].strip().split(",")
        values = [value.strip() for value in values]
        str_index = str(index + 1)

        # Replace the {index} placeholder in the template
        hascript = hascript.replace("{index}", str_index)

        # Create a dictionary for the placeholders and their corresponding values
        placeholder_dict = {}
        for i, value in enumerate(values):
            placeholder_name = f"{selected_option.replace(' ', '_')}_{i+1}"

            # Check if the value contains '$', if not, replace with &apos;Value&apos;
            if '$' not in value:
                value = f"&apos;{value}&apos;"

            placeholder_dict[f"{{{placeholder_name}}}"] = value

        # Add the item value placeholder
        item_value = f"{selected_option.replace(' ', '_')}_{index + 1}"
        placeholder_dict['{item_value}'] = item_value

        # Handle the next item placeholder
        if index + 1 < len(ordered_values):
            next_item = ordered_values[index + 1].strip()
            next_option_value = next_item.split(":")[0].strip()
            next_option_item_value = f"{next_option_value.replace(' ', '_')}_{index + 2}"
            placeholder_dict['{next_option_item_value}'] = next_option_item_value
        else:
            placeholder_dict['{next_option_item_value}'] = ""

        # Replace all placeholders in the template with the values from the dictionary
        for placeholder, value in placeholder_dict.items():
            hascript = hascript.replace(placeholder, value)

        # Add the processed hascript to the XML output
        xml_output += hascript + "\n"

    # Close the XML structure
    xml_output += "</HAScript>"

    # Save the generated XML to the selected folder
    file_path = os.path.join(save_path.get(), "output.xml")
    with open(file_path, "w") as xml_file:
        xml_file.write(xml_output)

    messagebox.showinfo("Success", f"XML Hascripts generated and saved to:\n{file_path}")

def add_item_group(entries, option_name, mandatory_fields):
    global edit_index
    values = []
    for entry, is_mandatory in zip(entries, mandatory_fields):
        value = entry.get().strip()
        if not value and is_mandatory:
            messagebox.showwarning("Warning", f"Mandatory field cannot be empty.")
            return
        values.append(value)

    item_value = f"{option_name}: " + ", ".join(values)

    if edit_index is None:
        selected_list.insert(tk.END, item_value)
    else:
        selected_list.delete(edit_index)
        selected_list.insert(edit_index, item_value)
        edit_index = None

def toggle_dropdown(option_name, edit_index=None):
    for frame_data in frames.values():
        frame_data[0].pack_forget()  # Only unpack the frame to hide it.

    if option_name:
        frame, entries, mandatory_fields = frames[option_name]
        frame.pack(after=buttons[option_name], fill='x', padx=5, pady=5)

    if edit_index is not None:
        existing_values = selected_list.get(edit_index).split(":")[1].split(",")
        for entry, value in zip(entries, existing_values):
            entry.delete(0, tk.END)
            entry.insert(0, value)

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

    frame, entries, _ = frames[option_name.strip()]
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
    mandatory_fields = []

    variable_label = tk.Label(frame, text=f"Variable returned = {option_name.replace(' ', '_')}_<index>")
    variable_label.pack(anchor='w', pady=5)
    
    for field in input_labels:
        label_text = field["label"]
        is_mandatory = field["mandatory"]

        label = tk.Label(frame, text=label_text + (" !" if is_mandatory else ""))
        label.pack(anchor='w')
        entry = tk.Entry(frame, width=30)
        entry.pack(anchor='w')
        entries.append(entry)
        mandatory_fields.append(is_mandatory)

    add_button = tk.Button(frame, text=f"Add {option_name} Group",
                           command=lambda: add_item_group(entries, option_name, mandatory_fields))
    add_button.pack(pady=5)
    return frame, entries, mandatory_fields

# Main application window
root = tk.Tk()
root.title("Drag-and-Drop XML Generator with Custom Group Input")
root.geometry("800x600")

save_path = tk.StringVar(value=os.getcwd())

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
    "Create Policy": [
        {"label": "Product", "mandatory": True}, #1
        {"label": "Client", "mandatory": True}, #2
        {"label": "Agent", "mandatory": True}, #3
        {"label": "Payment Method", "mandatory": True}, #4
        {"label": "Billing Frequency", "mandatory": True}, #5
        {"label": "Sum Assured", "mandatory": True}, #6
        {"label": "Refferal", "mandatory": False}, #7
        {"label": "Risk Term", "mandatory": False}, #8
        {"label": "Prem Term", "mandatory": False}, #9
        {"label": "Loading Death", "mandatory": False}, #10
        {"label": "Loading TI/CI", "mandatory": False}, #11
        {"label": "Loading PerMill", "mandatory": False}, #12
        {"label": "Smoking", "mandatory": False}, #13
    ],
    "Create Client": [
        {"label": "Client Name", "mandatory": True},
        {"label": "Gender (F,M)", "mandatory": True},
        {"label": "DOB (DD/MM/YYYY)", "mandatory": True},
    ],
    "Renewal 1Y": [
        {"label": "Premium Amount", "mandatory": True},
        {"label": "Year", "mandatory": True},
    ],
    "Cancellation": [
        {"label": "Loan Amount", "mandatory": True},
        {"label": "Loan Type", "mandatory": False},
        {"label": "Repayment Schedule", "mandatory": False},
    ],
    "ReIssuance AFI": [
        {"label": "Policy", "mandatory": True},
        {"label": "Year", "mandatory": True},
    ],
    "Loading Changes": [
        {"label": "Savings Amount", "mandatory": True},
        {"label": "Plan Name", "mandatory": False},
        {"label": "Deposit Frequency", "mandatory": False},
    ],
}

frames = {}
buttons = {}

for option_name, labels in options.items():
    dropdown_button = tk.Button(scrollable_frame, text=option_name,
                                command=lambda n=option_name: toggle_dropdown(n))
    dropdown_button.pack(pady=5, fill=tk.X)
    buttons[option_name] = dropdown_button

    frame, entries, mandatory_fields = create_option_frame(scrollable_frame, option_name, labels)
    frame.pack_forget()
    frames[option_name] = (frame, entries, mandatory_fields)

generate_button_frame = tk.Frame(root)
generate_button_frame.pack(fill='x', padx=10, pady=5, side=tk.BOTTOM)

generate_button = tk.Button(generate_button_frame, text="Generate XML",
                            command=lambda: generate_xml([selected_list.get(i) for i in range(selected_list.size())]))
generate_button.pack(side=tk.LEFT, padx=5)

save_path_label = tk.Label(generate_button_frame, text="Save Path:")
save_path_label.pack(side=tk.LEFT, padx=5)

save_path_entry = tk.Entry(generate_button_frame, textvariable=save_path, width=30)
save_path_entry.pack(side=tk.LEFT, padx=5)

browse_button = tk.Button(generate_button_frame, text="Browse", command=select_save_path)
browse_button.pack(side=tk.LEFT, padx=5)

root.mainloop()
