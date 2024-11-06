import tkinter as tk
from tkinter import messagebox
import os

# Drag-and-drop helper variables
drag_data = {"index": None, "text": None}

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
        selected_option = item.split(":")[0].strip()  # Get the option name from the list
        template_file = hascript_templates.get(selected_option)

        if not os.path.exists(template_file):
            messagebox.showerror("Error", f"Template file not found: {template_file}")
            return

        with open(template_file, 'r') as file:
            hascript = file.read()

        # Replace placeholders with actual values
        values = item.split(":")[1].strip().split(",")  # Get the input values
        values = [value.strip() for value in values]  # Strip whitespace from each value
        str_index = str(index+1)
        hascript = hascript.replace("{index}", str_index)

        # Create a dictionary of placeholders and values
        placeholder_dict = {}
        for i, value in enumerate(values):
            # Creating placeholders with specific naming conventions
            placeholder_name = f"{selected_option.replace(' ', '_')}_{i+1}"
            placeholder_dict[f"{{{placeholder_name}}}"] = value  # e.g., {Create_Policy_1}: value

        # Add placeholders for the item_value and next_option_item_value
        item_value = f"{selected_option.replace(' ', '_')}_{index + 1}"
        placeholder_dict['{item_value}'] = item_value

        # Determine the next item value for the nextscreen tag
        if index + 1 < len(ordered_values):
            next_item = ordered_values[index + 1].strip()
            next_option_value = next_item.split(":")[0].strip()  # Get next option name
            next_option_item_value = f"{next_option_value.replace(' ', '_')}_{index + 2}"
            placeholder_dict['{next_option_item_value}'] = next_option_item_value
        else:
            placeholder_dict['{next_option_item_value}'] = "Table_Code_Maintenance_submenu_next_value"  # Fallback

        # Replace all placeholders in the hascript
        for placeholder, value in placeholder_dict.items():
            hascript = hascript.replace(placeholder, value)

        # Add hascript to XML output
        xml_output += hascript + "\n"

    # Add the closing tag
    xml_output += "</HAScript>"

    # Write to XML file
    with open("output.xml", "w") as xml_file:
        xml_file.write(xml_output)

    messagebox.showinfo("Success", f"XML Hascripts generated for {len(ordered_values)} ordered values.")




def add_item_group(entries, option_name):
    values = [entry.get().strip() for entry in entries]
    if not all(values):
        messagebox.showwarning("Warning", f"All fields for {option_name} must be filled in.")
        return

    item_value = f"{option_name}: " + ", ".join(values)
    selected_list.insert(tk.END, item_value)

def toggle_dropdown(option_name):
    # Hide all frames first
    for frame in frames.values():
        frame.pack_forget()
    # Show the selected frame
    frames[option_name].pack(after=buttons[option_name], fill='x', padx=5, pady=5)

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

# Create main application window
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

right_frame_container = tk.Frame(root)
right_frame_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(right_frame_container)
scrollbar = tk.Scrollbar(right_frame_container, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
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
buttons = {}  # Changed to a dictionary
option1_entries = []

for option_name, labels in options.items():
    dropdown_button = tk.Button(scrollable_frame, text=option_name,
                                command=lambda n=option_name: toggle_dropdown(n))
    dropdown_button.pack(pady=5, fill=tk.X)
    buttons[option_name] = dropdown_button  # Store the button in a dictionary

    frame, entries = create_option_frame(scrollable_frame, option_name, labels)
    frame.pack_forget()  # Start with frames hidden
    frames[option_name] = frame  # Store the frame in a dictionary

    # Store entries for the first option to generate XML
    if option_name == "Create Policy":
        option1_entries = entries

generate_button = tk.Button(root, text="Generate XML",
                            command=lambda: generate_xml([selected_list.get(i) for i in range(selected_list.size())]))
generate_button.pack(side=tk.BOTTOM, pady=10)

root.mainloop()
