from tkinter import *
from tkinter import filedialog, messagebox
import requests
from os import getcwd, path
import xlsxwriter


def fetch_crypto_data():
    url1 = "https://api.coincap.io/v2/assets"
    params = {"limit": 20}
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


def create_file_selector(root):
    label1 = Label(root, text="Select File:", font=15, bg="#CE7816")
    label1.place(x=30, y=60)

    entry1 = Entry(root, width=40, font=40, bg="#CE7816")
    entry1.place(x=140, y=60)
    entry1.insert(0, getcwd() + "\\a.txt")

    button1 = Button(root, text="Browse", bg="#CE7816", command=lambda: browse_file(entry1))
    button1.place(x=580, y=60)
    return entry1


def create_filename_entry(root):
    label2 = Label(root, text="File Name:", font=15, bg="#CE7816")
    label2.place(x=30, y=100)

    entry2 = Entry(root, width=40, font=40, bg="#CE7816")
    entry2.place(x=140, y=100)
    return entry2


def create_directory_selector(root):
    label3 = Label(root, text="Select Directory:", font=15, bg="#CE7816")
    label3.place(x=30, y=140)

    entry3 = Entry(root, width=40, font=40, bg="#CE7816")
    entry3.place(x=190, y=140)
    entry3.insert(0, "C:\\Users\\Ervand\\Downloads")

    button3 = Button(root, text="Browse", bg="#CE7816", command=lambda: browse_directory(entry3))
    button3.place(x=620, y=140)
    return entry3


def generate_excel_file(data, entry2, entry3):
    filename = entry2.get()
    directory = entry3.get()
    found = False
    for line in data:
        if filename.lower() == line[1].lower():
            create_excel_file(line, directory, filename)
            found = True
            break
    if not found:
        messagebox.showwarning("Warning", "No matching cryptocurrency found")


def create_excel_file(line, directory, filename):
    file_path = path.join(directory, f"{filename}.xlsx")
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Name")
    worksheet.write(0, 1, "Symbol")
    worksheet.write(0, 2, "Price (USD)")
    worksheet.write(0, 3, "Volume (24Hr)")
    worksheet.write(0, 4, "Change (24Hr %)")

    for i in range(len(line)):
        worksheet.write(1, i, line[i])

    workbook.close()
    messagebox.showinfo("Success", f"Excel file created at {file_path}")
    exit()


def main():
    fetch_crypto_data()

    root = Tk()
    root.title('Crypto')
    icon = PhotoImage(file="bitcoin.png")
    root.iconphoto(True, icon)
    root.config(bg="#F7931A")
    root.geometry('700x300+200+200')
    root.resizable(width=False, height=False)

    entry1 = create_file_selector(root)
    entry2 = create_filename_entry(root)
    entry3 = create_directory_selector(root)

    data = read_crypto_data_from_txt(entry1.get())

    exel_button = Button(root, text="Get Excel", bg="#CE7816", width=15, height=2,
                         command=lambda: generate_excel_file(data, entry2, entry3))
    exel_button.place(x=290, y=180)

    root.mainloop()


if __name__ == "__main__":
    main()
