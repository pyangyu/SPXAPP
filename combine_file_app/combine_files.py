import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

# define a function to perform the concatenation and combination
'''
def combine_files():
    # get the selected folder path
    folder_path = folder_path_var.get()
    
    # get a list of all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

    # initialize an empty list to hold the DataFrames for each file
    df_list = []

    # iterate through each Excel file and concatenate the specified columns
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        if file_name == "combined.xlsx":
            continue

        # check if the file contains the specified columns
        df_cols = pd.read_excel(file_path, sheet_name="Parcel scan result", nrows=0).columns
        while "Tracking Number" not in df_cols:
            # try checking the next row in case the headers are not in the first row
            df_cols = pd.read_excel(file_path, sheet_name="Parcel scan result", skiprows=1, nrows=0).columns
            print(df_cols)
            if "Tracking Number" not in df_cols:
                # if there is no "Tracking Number" column, move to the next file
                break

        if "Tracking Number" not in df_cols:
            # if there is no "Tracking Number" column in this file, move to the next file
            continue

        # extract the date and label from the file name
        date_str, label = file_name.split(" ")[0:2]
        date = pd.to_datetime(date_str, format="%m%d").strftime("%m-%d")

        skip_rows = 0
        while True:
            df = pd.read_excel(file_path, sheet_name="Parcel scan result", header=skip_rows)
            if "Tracking Number" in df.columns:
                break
            skip_rows += 1
        

        # define the list of column names
        columns = ["Tracking Number", "Scan Date", "Scan Operator"]

        # read the DataFrame selecting columns that contain any of the specified column names
        df = pd.read_excel(file_path, sheet_name="Parcel scan result", usecols=lambda x: any(col in x for col in columns), skiprows=skip_rows)

        # define the list of column names to look for and their corresponding names
        columns = {"Tracking Number": "Tracking Number", "Scan Date": "Scan Date", "Scan Operator": "Scan Operator"}

        # function to check if the column name contains any of the keywords and rename it
        def rename_columns(column_name):
            for keyword in columns:
                if keyword in column_name:
                    return columns[keyword]
            return column_name


        # rename the columns
        df.columns = [rename_columns(col.strip()) for col in df.columns]


        # add the date and label columns to the DataFrame
        df["Date"] = date
        df["Label"] = label

        def replace_empty_with_no_sca(df):
            df["Scan Date"].fillna("no scan", inplace=True)
            df["Scan Operator"].fillna("no scan", inplace=True)
            return df

        df = replace_empty_with_no_sca(df)

        # append the DataFrame to the list
        df_list.append(df[["Tracking Number", "Scan Date", "Scan Operator", "Date", "Label"]])

    if not df_list:
        # if there are no files with "Tracking Number" column, show a message box and return
        messagebox.showinfo("Combine Files", "No files with 'Tracking Number' column found.")
        return

    # concatenate all the DataFrames together
    combined_df = pd.concat(df_list, ignore_index=True)

    # write the combined DataFrame to a new Excel file
    combined_file_path = os.path.join(folder_path, "combined.xlsx")
    combined_df.to_excel(combined_file_path, index=False)
    
    # display a message box to indicate the operation is complete
    messagebox.showinfo("Combine Files", "The files have been successfully combined.")
'''


def combine_files_new():
    # get the selected folder path
    folder_path = folder_path_var.get()

    # get a list of all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and f != "combined.xlsx"]

    # initialize an empty list to hold the DataFrames for each file
    df_list = []
    excel_files = sorted(excel_files, key=lambda x: int(x[x.find("Route") + 5:].split(".")[0].rstrip(" ")))
    print(excel_files)

    # iterate through each Excel file and concatenate all columns
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        if file_name == "combined.xlsx":
            continue

        # # extract the date and label from the file name
        # date_str, label = file_name.split(" ")[0:2]
        # date = pd.to_datetime(date_str, format="%m%d").strftime("%m-%d")

        skip_rows = 0
        while True:
            df = pd.read_excel(file_path, sheet_name="Parcel scan result", header=skip_rows)
            if "Tracking Number" in df.columns:
                break
            skip_rows += 1

        # # add the date and label columns to the DataFrame
        # df["Date"] = date
        # # df["Label"] = label
        #
        # def replace_empty_with_no_sca(df):
        #     columns_to_fill = [col for col in df.columns if "Scan Date" in col or "Scan Operator" in col]
        #     df[columns_to_fill] = df[columns_to_fill].fillna("no scan")
        #     return df
        #
        # df = replace_empty_with_no_sca(df)

        # append the DataFrame to the list
        df_list.append(df)

    if not df_list:
        # if there are no files with "Tracking Number" column, show a message box and return
        messagebox.showinfo("Combine Files", "No files with 'Tracking Number' column found.")
        return

    # concatenate all the DataFrames together
    # combined_df = pd.concat(df_list, ignore_index=True)
    combined_df = pd.concat(df_list, join='outer', sort=False)
    # combined_df = combined_df.sort_values(by='Route Code')

    # write the combined DataFrame to a new Excel file
    combined_file_path = os.path.join(folder_path, "combined.xlsx")
    combined_df.to_excel(combined_file_path, index=False)

    # display a message box to indicate the operation is complete
    messagebox.showinfo("Combine Files", "The files have been successfully combined.")


# initialize the UI
root = tk.Tk()
root.title("Combine Excel Files")

# set the window size
root.geometry("500x500")

# create a canvas widget
canvas = tk.Canvas(root, width=500, height=500)
# create an image object from the icon.webp file
# get the selected folder path
image_folder_path = os.path.join(os.getcwd(), 'combine_icon.png')
image = Image.open(image_folder_path)
# adjust the alpha channel to 0.3
image.putalpha(int(255*0.2))
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
folder_path_var = tk.StringVar()
folder_path_label = tk.Label(root, text="Folder Path:")
folder_path_label.pack(side=tk.TOP)
folder_path_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
folder_path_entry = tk.Entry(root, textvariable=folder_path_var, width=50)
folder_path_entry.pack(side=tk.TOP)
folder_path_entry.place(relx=0.5, rely=0.4, anchor=tk.CENTER)


def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_var.set(folder_path)


browse_button = tk.Button(root, text="Browse", command=browse_folder, font=("Helvetica", 12), bg="lightblue", bd=0.8,
                          relief=tk.RAISED, activebackground="#FF9999", activeforeground="white")
browse_button.pack(pady=10, side=tk.TOP)
browse_button.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

# add a button to perform the combination
combine_button = tk.Button(root, text="Combine Files", command=combine_files_new)
combine_button.pack(pady=5, side=tk.BOTTOM)
combine_button.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

# start the UI loop
root.mainloop()
