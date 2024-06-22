from tkinter import *
from tkinter import filedialog, messagebox
import requests
from os import getcwd, path
import xlsxwriter

def fetch_crypto_data():
    url1 = "https://api.coincap.io/v2/assets"
    params = {"limit": 50}
    r = requests.get(url1, params=params)
    r.raise_for_status()
    data = r.json()["data"]
    save_crypto_data_to_txt(data)

def save_crypto_data_to_txt(data):
    with open("a.txt", "w") as f:
        for item in data:
            f.write(f"{item['name']} {item['symbol']} {item['priceUsd']}$ {item['volumeUsd24Hr']}$ "
                    f"{item['changePercent24Hr']}%\n")

def read_crypto_data_from_txt(file_path):
    with open(file_path, "r") as f:
        lines = f.readlines()
    return [line.split() for line in lines]

def browse_file(entry1):
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path:
        entry1.delete(0, END)
        entry1.insert(0, file_path)

def browse_directory(entry):
    directory_path = filedialog.askdirectory()
    if directory_path:
        entry.delete(0, END)
        entry.insert(0, directory_path)

def create_file_selector(frame):
    label1 = Label(frame, text="Select File:", font=15, bg="#CE7816")
    label1.pack(side=LEFT, padx=10, pady=5)

    entry1 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry1.pack(side=LEFT, padx=10, pady=5)
    entry1.insert(0, getcwd() + "\\a.txt")

    button1 = Button(frame, text="Browse", bg="#CE7816", command=lambda: browse_file(entry1))
    button1.pack(side=LEFT, padx=10, pady=5)
    return entry1

def create_filename_entry(frame):
    label2 = Label(frame, text="File Name:", font=15, bg="#CE7816")
    label2.pack(side=LEFT, padx=10, pady=5)

    entry2 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry2.pack(side=LEFT, padx=10, pady=5)
    return entry2

def create_directory_selector(frame):
    label3 = Label(frame, text="Select Directory:", font=15, bg="#CE7816")
    label3.pack(side=LEFT, padx=10, pady=5)

    entry3 = Entry(frame, width=40, font=40, bg="#CE7816")
    entry3.pack(side=LEFT, padx=10, pady=5)
    entry3.insert(0, "C:\\Users\\Ervand\\Downloads")

    button3 = Button(frame, text="Browse", bg="#CE7816", command=lambda: browse_directory(entry3))
    button3.pack(side=LEFT, padx=10, pady=5)
    return entry3


def generate_excel_file(data, entry2, entry3):
    filename = entry2.get()
    directory = entry3.get()
    file_path = path.join(directory, f"{filename}.xlsx")
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    # Headers
    headers = ["Name", "Symbol", "Price (USD)", "Volume (24Hr)", "Change (24Hr %)"]
    for i, header in enumerate(headers):
        worksheet.write(0, i, header)

    # Write data
    for j, line in enumerate(data):
        for i, value in enumerate(line):
            worksheet.write(j + 1, i, value)

    workbook.close()
    messagebox.showinfo("Success", f"Excel file created at {file_path}")


def main():
    fetch_crypto_data()

    root = Tk()
    root.title('Crypto')
    icon = PhotoImage(file="bitcoin.png")
    root.iconphoto(True, icon)
    root.config(bg="#F7931A")
    root.geometry('600x300+200+200')
    root.resizable(width=False, height=False)

    # Create and pack frames
    frame1 = Frame(root, bg="#F7931A")
    frame1.pack(pady=10, fill=X)
    entry1 = create_file_selector(frame1)

    frame2 = Frame(root, bg="#F7931A")
    frame2.pack(pady=10, fill=X)
    entry2 = create_filename_entry(frame2)

    frame3 = Frame(root, bg="#F7931A")
    frame3.pack(pady=10, fill=X)
    entry3 = create_directory_selector(frame3)

    data = read_crypto_data_from_txt(entry1.get())

    exel_button = Button(root, text="Get Excel", bg="#CE7816", width=15, height=2,
                         command=lambda: generate_excel_file(data, entry2, entry3))
    exel_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
