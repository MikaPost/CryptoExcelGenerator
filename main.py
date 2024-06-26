from tkinter import *
from tkinter import filedialog, messagebox
import requests
from os import getcwd, path
import xlsxwriter


def fetch_crypto_data():
    """
    Function: fetch_crypto_data
    Brief: Fetches cryptocurrency data from an API and saves it to a text file.
    Params: None
    Return: None
    """
    try:
        url1 = "https://api.coincap.io/v2/assets"
        params = {"limit": 50}
        r = requests.get(url1, params=params)
        r.raise_for_status()
        data = r.json()["data"]
        save_crypto_data_to_txt(data)
    except:
        messagebox.showerror("Error Check your internet connection")


def save_crypto_data_to_txt(data):
    """
    Function: save_crypto_data_to_txt
    Brief: Saves cryptocurrency data to a text file.
    Params:
        - data (list): List of cryptocurrency data.
    Return: None
    """
    try:
        # Open the file in write mode
        with open("a.txt", "w") as f:
            # Write each item's details in a formatted manner
            for item in data:
                f.write(f"{item['name']} {item['symbol']} {item['priceUsd']}$ {item['volumeUsd24Hr']}$ "
                        f"{item['changePercent24Hr']}%\n")
    except IOError as e:
        messagebox.showerror("Error", f"Failed to save data: {e}")


def read_crypto_data_from_txt(file_path):
    """
    Function: read_crypto_data_from_txt
    Brief: Reads cryptocurrency data from a text file.
    Params:
        - file_path (str): Path to the text file.
    Return: List of lists, where each inner list represents a line of data split into components.
    """
    with open(file_path, "r") as f:
        lines = f.readlines()

    return [line.split() for line in lines]


def browse_file(entry1):
    """
    Function: browse_file
    Brief: Opens a file dialog to browse and select a text file.
    Params:
        - entry1 (Entry): Entry widget to display the selected file path.
    Return: None
    """
    # Open a file dialog to select a text file
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

    # If a file is selected, update the entry widget with the file path
    if file_path:
        entry1.delete(0, END)
        entry1.insert(0, file_path)


def browse_directory(entry):
    """
    Function: browse_directory
    Brief: Opens a directory dialog to browse and select a directory path.
    Params:
        - entry (Entry): Entry widget to display the selected directory path.
    Return: None
    """
    # Open a directory dialog to select a directory
    directory_path = filedialog.askdirectory()

    # If a directory is selected, update the entry widget with the directory path
    if directory_path:
        entry.delete(0, END)
        entry.insert(0, directory_path)


def create_file_selector(frame):
    """
    Function: create_file_selector
    Brief: Creates file selection UI components.
    Params:
        - frame (Frame): Parent frame to place UI components.
    Return: Entry widget representing the file path selection.
    """
    # Create and pack a label
    label1 = Label(frame, text="Select File:", font=15, bg="#CE7816")
    label1.pack(side=LEFT, padx=10, pady=5)

    # Create and pack an entry widget for file path input
    entry1 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry1.pack(side=LEFT, padx=10, pady=5)
    entry1.insert(0, getcwd() + "\\a.txt")

    # Create and pack a browse button
    button1 = Button(frame, text="Browse", bg="#CE7816", command=lambda: browse_file(entry1))
    button1.pack(side=LEFT, padx=10, pady=5)

    return entry1


def create_filename_entry(frame):
    """
    Function: create_filename_entry
    Brief: Creates filename entry UI components.
    Params:
        - frame (Frame): Parent frame to place UI components.
    Return: Entry widget representing the filename input.
    """
    # Create and pack a label
    label2 = Label(frame, text="File Name:", font=15, bg="#CE7816")
    label2.pack(side=LEFT, padx=10, pady=5)

    # Create and pack an entry widget for filename input
    entry2 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry2.pack(side=LEFT, padx=10, pady=5)

    return entry2


def create_directory_selector(frame):
    """
    Function: create_directory_selector
    Brief: Creates directory selection UI components.
    Params:
        - frame (Frame): Parent frame to place UI components.
    Return: Entry widget representing the directory path selection.
    """
    # Create and pack a label
    label3 = Label(frame, text="Select Directory:", font=15, bg="#CE7816")
    label3.pack(side=LEFT, padx=10, pady=5)

    # Create and pack an entry widget for directory path input
    entry3 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry3.pack(side=LEFT, padx=10, pady=5)
    entry3.insert(0, "C:\\Users\\Ervand\\Downloads")

    # Create and pack a browse button
    button3 = Button(frame, text="Browse", bg="#CE7816", command=lambda: browse_directory(entry3))
    button3.pack(side=LEFT, padx=10, pady=5)

    return entry3


def generate_excel_file(data, entry2, entry3):
    """
    Function: generate_excel_file
    Brief: Generates an Excel file with cryptocurrency data.
    Params:
        - data (list): List of cryptocurrency data to write to Excel.
        - entry2 (Entry): Entry widget containing the filename.
        - entry3 (Entry): Entry widget containing the directory path.
    Return: None
    """
    # Get the filename and directory path from the entry widgets
    filename = entry2.get()
    directory = entry3.get()
    file_path = path.join(directory, f"{filename}.xlsx")

    # Create an Excel workbook and worksheet
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    # Define the headers for the Excel file
    headers = ["Name", "Symbol", "Price (USD)", "Volume (24Hr)", "Change (24Hr %)"]

    # Write the headers to the first row of the worksheet
    for i, header in enumerate(headers):
        worksheet.write(0, i, header)

    # Write the data to the worksheet, starting from the second row
    for j, line in enumerate(data):
        for i, value in enumerate(line):
            worksheet.write(j + 1, i, value)

    # Close the workbook to save the file
    workbook.close()

    # Show a success message
    messagebox.showinfo("Success", f"Excel file created at {file_path}")


def main():
    """
    Function: main
    Brief: Entry point of the program.
    """
    # Fetch the cryptocurrency data
    fetch_crypto_data()

    # Initialize the Tkinter root window
    root = Tk()
    root.title('Crypto')

    # Set the window icon and background color
    icon = PhotoImage(file="bitcoin.png")
    root.iconphoto(True, icon)
    root.config(bg="#F7931A")
    root.geometry('600x300+200+200')
    root.resizable(width=False, height=False)

    # Create and pack the first frame for file selection
    frame1 = Frame(root, bg="#F7931A")
    frame1.pack(pady=10, fill=X)
    entry1 = create_file_selector(frame1)

    # Create and pack the second frame for filename input
    frame2 = Frame(root, bg="#F7931A")
    frame2.pack(pady=10, fill=X)
    entry2 = create_filename_entry(frame2)

    # Create and pack the third frame for directory selection
    frame3 = Frame(root, bg="#F7931A")
    frame3.pack(pady=10, fill=X)
    entry3 = create_directory_selector(frame3)

    # Read the data from the selected text file
    data = read_crypto_data_from_txt(entry1.get())

    # Create and pack the button to generate the Excel file
    exel_button = Button(root, text="Get Excel", bg="#CE7816", width=15, height=2,
                         command=lambda: generate_excel_file(data, entry2, entry3))
    exel_button.pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    main()
