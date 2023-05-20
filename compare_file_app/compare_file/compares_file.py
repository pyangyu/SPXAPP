import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import re
import xlwt
import os
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import messagebox

pattern1 = r"\d{3}-\d{8}"
pattern2 = r"\d{11}"


def extract_file_name(file_name):
    file_name = file_name.replace(" ", ",")
    matches = re.findall(pattern1, file_name)
    if not matches:
        matches = re.findall(pattern2, file_name)
        if matches:
            matches = matches[0]
            return matches[:3] + "-" + matches[3:]
        else:
            return ""
    return matches[0]


# have the T86 folder named "T86" and the scanned folder named "scanned" under the same directory
# this app will generate the new folder under the same directory called "compare result"

def compare_files():
    # get the selected folder path
    # open T86 folder
    folder_path_T86 = folder_path_var1.get()
    folder_path_scanned = folder_path_var2.get()

    # create a new subdirectory called "complete_audit_file" under the specified folder path
    parent_folder = os.path.dirname(folder_path_T86)
    complete_files = os.path.join(parent_folder, "compare result")
    complete_files_T86 = os.path.join(folder_path_T86, "rename T86 files")
    complete_files_scanned = os.path.join(folder_path_scanned, "rename scanned files")
    os.makedirs(complete_files, exist_ok=True)
    os.makedirs(complete_files_T86, exist_ok=True)
    os.makedirs(complete_files_scanned, exist_ok=True)
    excel_files_T86 = [f for f in os.listdir(folder_path_T86) if f.endswith(".xls") or f.endswith(".xlsx")]
    excel_files_scanned = [f for f in os.listdir(folder_path_scanned) if f.endswith(".xls") or f.endswith(".xlsx")]
    excel_files_T86_rename = []
    excel_files_scanned_rename = []

    for excel_file in excel_files_T86:
        # read the Excel file into a Pandas dataframe
        df = pd.read_excel(os.path.join(folder_path_T86, excel_file))
        # if "consignor_item_id" in df.columns:
        #     df = df[df['consignor_item_id'].str.contains('SPX')]
        # elif "服务商单号" in df.columns:
        #     df = df[df['服务商单号'].str.contains('SPX')]
        # else:
        #     print("error")
        new_file_name = extract_file_name(excel_file)
        excel_files_T86_rename.append(new_file_name)
        if new_file_name == "":
            continue

        # write the new dataframe to a new Excel file with the trimmed substring as the name,
        # in the "complete_audit_file" subdirectory
        new_file_name = new_file_name + "_T86.xlsx"
        new_file_path = os.path.join(complete_files_T86, new_file_name)

        # create a new workbook and worksheet using openpyxl
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

        # write the dataframe to the worksheet
        for row in dataframe_to_rows(df, index=False, header=True):
            worksheet.append(row)

        # save the workbook to the new Excel file
        workbook.save(new_file_path)
    final_string_presenting = ""

    for excel_file in excel_files_scanned:
        # read the Excel file into a Pandas dataframe
        df = pd.read_excel(os.path.join(folder_path_scanned, excel_file), header=None)
        df = df.iloc[:, 0]
        df.columns = ['scanned result']
        new_file_name = extract_file_name(excel_file)
        excel_files_scanned_rename.append(new_file_name)
        if new_file_name == "":
            continue
        scanned_df = df.iloc[:, 0]
        scanned_df = scanned_df.to_frame()
        # write the new dataframe to a new Excel file with the trimmed substring as the name,
        # in the "complete_audit_file" subdirectory
        new_file_name = new_file_name + "_scanned.xlsx"
        new_file_path = os.path.join(complete_files_scanned, new_file_name)

        # create a new workbook and worksheet using openpyxl
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

        # write the dataframe to the worksheet
        for row in dataframe_to_rows(scanned_df, index=False, header=True):
            worksheet.append(row)

        # save the workbook to the new Excel file
        workbook.save(new_file_path)

    final_string_presenting += "The Revd A/R values would be: \n"
    final_string_presenting += "-" * 50 + '\n'
    for excel_file_scanned in excel_files_scanned_rename:
        if excel_file_scanned in excel_files_T86_rename:
            df_T86 = pd.read_excel(os.path.join(complete_files_T86, excel_file_scanned + "_T86.xlsx"))
            df_scanned = pd.read_excel(os.path.join(complete_files_scanned, excel_file_scanned + "_scanned.xlsx"))
            if "箱号" in df_T86.columns.tolist():
                df_T86 = df_T86.rename(columns={"箱号": "receptacle_id"})
                print(df_T86.columns)
            consignor_item_id_scan = df_scanned.iloc[:, 0]
            consignor_item_id_scan = consignor_item_id_scan.drop_duplicates()
            consignor_item_id_scan = consignor_item_id_scan.tolist()
            consignor_item_id_scan = list(set(consignor_item_id_scan))
            consignor_item_id_T86 = df_T86["receptacle_id"]
            consignor_item_id_T86 = consignor_item_id_T86.drop_duplicates()
            consignor_item_id_T86 = consignor_item_id_T86.tolist()
            consignor_item_id_T86 = list(set(consignor_item_id_T86))
            # the receptacle id label is broken, use the package inside instead
            for consignor_item_id_scan_item in consignor_item_id_scan:
                if "SPX" in consignor_item_id_scan_item:
                    if df_T86['consignor_item_id'].str.contains(consignor_item_id_scan_item).any():
                        continue
            scanned_finished = 0
            for element in consignor_item_id_scan:
                if element in consignor_item_id_T86:
                    scanned_finished += 1
            final_string_presenting += excel_file_scanned + '\n'
            final_string_presenting += f"{scanned_finished} / {len(consignor_item_id_T86)} \n"
            final_string_presenting += "-" * 50 + '\n'

    messagebox.showinfo("Comparison", final_string_presenting)


def browse_folder_T86():
    folder_path = filedialog.askdirectory()
    folder_path_var1.set(folder_path)


def browse_folder_Scan():
    folder_path = filedialog.askdirectory()
    folder_path_var2.set(folder_path)


if __name__ == "__main__":
    # initialize the UI
    root = tk.Tk()
    root.title("Matches")

    # set the window size
    root.geometry("500x500")

    # create a canvas widget
    canvas = tk.Canvas(root, width=500, height=500)
    # create an image object from the icon.webp file
    # get the selected folder path
    image_folder_path = os.path.join(os.getcwd(), 'combine_icon.png')
    image = Image.open(image_folder_path)
    # adjust the alpha channel to 0.3
    image.putalpha(int(255 * 0.2))
    photo_image = ImageTk.PhotoImage(image)
    # set the canvas background color to white
    canvas.configure(bg='#DAE6E6')
    # create a rectangle with the same size as the canvas to serve as the background
    background = canvas.create_rectangle(0, 0, 500, 500, fill="#DAE6E6", outline="#DAE6E6")
    # create an image item on the canvas with the icon.webp image
    canvas.create_image(0, 0, image=photo_image, anchor="nw")
    # pack the canvas widget to fill the window
    canvas.pack(fill="both", expand=True)

    # add a label and entry for the folder path
    folder_path_var1 = tk.StringVar()
    folder_path_label = tk.Label(root, text="Original Folder Path (contains T86 files):")
    folder_path_label.pack(side=tk.TOP)
    folder_path_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
    folder_path_entry = tk.Entry(root, textvariable=folder_path_var1, width=80)
    folder_path_entry.pack(side=tk.TOP)
    folder_path_entry.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

    browse_button = tk.Button(root, text="Browse T86 folder", command=browse_folder_T86, font=("Helvetica", 10),
                              bg="lightblue",
                              bd=0.8,
                              relief=tk.RAISED, activebackground="#FF9999", activeforeground="white")
    browse_button.pack(pady=10, side=tk.TOP)
    browse_button.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # -----------------------------------------------------------
    # add a label and entry for the folder path
    folder_path_var2 = tk.StringVar()
    folder_path_label = tk.Label(root, text="Scan Folder Path (contains T86 files):")
    folder_path_label.pack(side=tk.TOP)
    folder_path_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
    folder_path_entry = tk.Entry(root, textvariable=folder_path_var2, width=80)
    folder_path_entry.pack(side=tk.TOP)
    folder_path_entry.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

    browse_button = tk.Button(root, text="Browse Scan folder", command=browse_folder_Scan, font=("Helvetica", 10),
                              bg="lightblue",
                              bd=0.8,
                              relief=tk.RAISED, activebackground="#FF9999", activeforeground="white")
    browse_button.pack(pady=10, side=tk.TOP)
    browse_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # add a button to perform the combination
    combine_button = tk.Button(root, text="Generate Files", command=compare_files)
    combine_button.pack(pady=5, side=tk.BOTTOM)
    combine_button.place(relx=0.5, rely=0.9, anchor=tk.CENTER)

    # start the UI loop
    root.mainloop()
