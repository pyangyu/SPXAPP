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
    file_name = file_name.replace(" ", "")
    matches = re.findall(pattern1, file_name)
    if not matches:
        matches = re.findall(pattern2, file_name)
        if matches:
            matches = matches[0]
            return matches[:3] + "-" + matches[3:]
        else:
            return []
    return matches[0]




def audit_files_new():
    # get the selected folder path
    folder_path = folder_path_var.get()

    # create a new subdirectory called "complete_audit_file" under the specified folder path
    complete_filtered_files = os.path.join(folder_path, "completed filter files")
    os.makedirs(complete_filtered_files, exist_ok=True)

    # get a list of all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xls") or f.endswith(".xlsx")]
    final_string = ""
    for excel_file in excel_files:
        new_file_name = extract_file_name(excel_file)
        final_string += f"AWB is: {new_file_name} \n"

        # read the Excel file into a Pandas dataframe
        df = pd.read_excel(os.path.join(folder_path, excel_file))

        # check if the "Tracking Number" column exists in the dataframe
        if "consignor_item_id" in df.columns:
            # filter the rows containing "SPX"
            filtered_df = df[df['consignor_item_id'].str.contains('SPX')]
            filtered_df = filtered_df.drop_duplicates(subset="consignor_item_id")
            if "receptacle_id" in df.columns:
                receptacle_id_unique = filtered_df['receptacle_id']
                receptacle_id_unique = receptacle_id_unique.drop_duplicates()
                final_string += f"There are {len(receptacle_id_unique)} boxes.\n"
            total_rows = len(filtered_df)
            num_rows_containing_ah = filtered_df["consignor_item_id"].str.contains('AH', case=False).sum()
            num_rows_containing_ge = filtered_df["consignor_item_id"].str.contains('GE', case=False).sum()
            num_rows_containing_ud = filtered_df["consignor_item_id"].str.contains('UD', case=False).sum()
            final_string += f"There are {total_rows} total parcels.\n"
            final_string += f"There are {num_rows_containing_ah} rows containing 'AH' in the DataFrame.\n"
            final_string += f"There are {num_rows_containing_ge} rows containing 'GE' in the DataFrame.\n"
            final_string += f"There are {num_rows_containing_ud} rows containing 'UD' in the DataFrame.\n"
            final_string += "-" * 50 + '\n'

            # write the new dataframe to a new Excel file with the trimmed substring as the name,
            # in the "complete_audit_file" subdirectory
            new_file_name = new_file_name + "_T86.xlsx"
            new_file_path = os.path.join(complete_filtered_files, new_file_name)

            # create a new workbook and worksheet using openpyxl
            workbook = openpyxl.Workbook()
            worksheet = workbook.active

            # write the dataframe to the worksheet
            for row in dataframe_to_rows(filtered_df, index=False, header=True):
                worksheet.append(row)

            # save the workbook to the new Excel file
            workbook.save(new_file_path)
        else:
            # skip this file if the "Tracking Number" column does not exist
            print(f"Skipping {excel_file} because the 'consignor_item_id' column does not exist.")

    # display a message box to indicate the operation is complete
    messagebox.showinfo("T86 Filter Files", final_string)

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_var.set(folder_path)


if __name__ == "__main__":
    # initialize the UI
    root = tk.Tk()
    root.title("Filter T86")

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
    # set the canvas background color to #DAE6E6
    canvas.configure(bg='#DAE6E6')
    # create a rectangle with the same size as the canvas to serve as the background
    background = canvas.create_rectangle(0, 0, 500, 500, fill="#DAE6E6", outline="#DAE6E6")
    # create an image item on the canvas with the icon.webp image
    canvas.create_image(0, 0, image=photo_image, anchor="nw")
    # pack the canvas widget to fill the window
    canvas.pack(fill="both", expand=True)

    # add a label and entry for the folder path
    folder_path_var = tk.StringVar()
    folder_path_label = tk.Label(root, text="Folder Path:", font=("Helvetica", 20, "bold"), fg="darkblue", bg="#DAE6E6")
    folder_path_label.pack(side=tk.TOP)
    folder_path_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
    folder_path_entry = tk.Entry(root, textvariable=folder_path_var, width=40, font=("Helvetica", 14))
    folder_path_entry.pack(side=tk.TOP)
    folder_path_entry.place(relx=0.5, rely=0.4, anchor=tk.CENTER)

    browse_button = tk.Button(root, text="Browse", command=browse_folder, font=("Helvetica", 12), bg="orange",
                              bd=2, relief=tk.RAISED, activebackground="#FF9999", activeforeground="white",
                              padx=10, pady=5)
    browse_button.pack(pady=10, side=tk.TOP)
    browse_button.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

    # add a button to perform the combination
    combine_button = tk.Button(root, text="Filter T86 Files", command=audit_files_new, font=("Helvetica", 12), bg="lightblue",
                              bd=2, relief=tk.RAISED, activebackground="#FF9999", activeforeground="white",
                              padx=10, pady=5)
    combine_button.pack(pady=5, side=tk.BOTTOM)
    combine_button.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

    # start the UI loop
    root.mainloop()
