import pandas as pd
from tabulate import tabulate
from tkinter import Tk, Label, Entry, Button, Text, Scrollbar, filedialog, messagebox, END
from tkinter.ttk import Style
from ttkthemes import ThemedTk

# Function to load and process the Excel file
def load_excel(file_path):
    data = pd.read_excel(file_path, header=1)  # Use the second row as the header
    data.columns = data.columns.str.strip()  # Clean column names
    return data[['MOUZA', 'PLOT NO', 'KH. NO.', 'OWNERS NAME']]  # Return relevant columns

# Function to filter data by owner's name
def filter_data(owner_name):
    global extracted_data
    return extracted_data[extracted_data['OWNERS NAME'].str.contains(owner_name, case=False, na=False)]

# Function to proceed to the main app after selecting the file
def proceed_to_main_app(file_path):
    global extracted_data, root

    try:
        # Load and prepare the data
        extracted_data = load_excel(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load the file: {e}")
        return

    # Destroy the welcome window
    root.destroy()

    # Launch the main app
    main_app()

# Function to open the welcome page
def welcome_page():
    global root
    root = ThemedTk(theme="adapta")  # Vibrant theme for the UI
    root.title("Welcome")

    Label(
        root, 
        text="Welcome to the Customer Data Search App!", 
        font=("Helvetica", 18, "bold"), 
        fg="darkblue"
    ).pack(pady=20)
    Label(root, text="Please select an Excel file to continue.", font=("Helvetica", 12), fg="green").pack(pady=10)

    Button(
        root, text="Select Excel File", font=("Helvetica", 14, "bold"), bg="lightblue", fg="black",
        command=select_file
    ).pack(pady=20)

    root.mainloop()

# Function to select a file and proceed
def select_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        proceed_to_main_app(file_path)
    else:
        messagebox.showwarning("File Selection", "No file selected. Please select a file to proceed.")

# Main application for searching data
def main_app():
    main_root = ThemedTk(theme="breeze")  # A light and modern theme for the main UI
    main_root.title("Customer Data Search")

    Label(
        main_root, 
        text="Customer Data Search", 
        font=("Helvetica", 20, "bold"), 
        fg="darkgreen"
    ).grid(row=0, column=0, columnspan=3, pady=10)

    # Owner's name input
    Label(main_root, text="Enter Owner's Name:", font=("Helvetica", 14), fg="black").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    owner_name_entry = Entry(main_root, width=30, font=("Helvetica", 14))
    owner_name_entry.grid(row=1, column=1, padx=10, pady=10)

    # Search button
    search_button = Button(
        main_root, text="Search", font=("Helvetica", 14, "bold"), bg="orange", fg="black",
        command=lambda: search_owner(owner_name_entry, result_text)
    )
    search_button.grid(row=1, column=2, padx=10, pady=10)

    # Bind Enter key to trigger search
    main_root.bind("<Return>", lambda event: search_owner(owner_name_entry, result_text))

    # Results display
    Label(main_root, text="Search Results:", font=("Helvetica", 14), fg="blue").grid(row=2, column=0, padx=10, pady=10, sticky="nw")
    result_text = Text(main_root, width=80, height=20, font=("Courier", 12), bg="#f7f7f7", fg="black")
    result_text.grid(row=2, column=1, padx=10, pady=10, columnspan=2)

    # Scrollbar for results
    scrollbar = Scrollbar(main_root, command=result_text.yview)
    result_text.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=2, column=3, sticky="ns")

    # Add an animated status bar
    status_bar = Label(main_root, text="Welcome! Ready to search.", font=("Helvetica", 12), fg="darkgreen", anchor="w")
    status_bar.grid(row=3, column=0, columnspan=3, sticky="we", pady=10)

    def animate_status(message, repeat=3):
        for i in range(repeat):
            main_root.after(i * 300, lambda: status_bar.config(text=message))
            main_root.after((i * 300) + 150, lambda: status_bar.config(text=""))

    owner_name_entry.bind("<KeyRelease>", lambda e: animate_status("Searching...", repeat=1))

    main_root.mainloop()

# Function to search and display results
def search_owner(owner_name_entry, result_text):
    owner_name = owner_name_entry.get().strip()
    if not owner_name:
        messagebox.showwarning("Input Error", "Please enter an owner's name.")
        return

    # Filter data
    filtered_data = filter_data(owner_name)

    # Clear previous results
    result_text.delete(1.0, END)

    if not filtered_data.empty:
        # Display results in a tabular format
        result_table = tabulate(filtered_data, headers='keys', tablefmt='grid', showindex=False)
        result_text.insert(END, result_table)
    else:
        result_text.insert(END, f"No data found for owner '{owner_name}'.")

# Start the welcome page
welcome_page()
