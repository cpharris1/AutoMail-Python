import csv
from tkinter import Tk
from tkinter import filedialog
import os


def select_file(kind):
    while True:
        root = Tk()
        #root.lift()
        root.attributes("-topmost", True)
        root.withdraw()
        if kind == "csv":
            root.filename = filedialog.askopenfilename(filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        elif kind == "html":
            root.filename = filedialog.askopenfilename(filetypes=(("HTML files", "*.html"), ("txt files", "*.txt"), ("All files", "*.*")))
        elif kind == "dir":
            root.filename = filedialog.askdirectory()

        while True:
            print("\nSelected file: " + root.filename)
            ans = input("\nIs this the file you wish to import? Press ENTER to continue, or 'N' to select new file. ").lower()

            if ans == 'n':
                break
            else:
                return root.filename


def split_csv(source_filepath, dest_folder, split_file_prefix, records_per_file):
    """
    Split a source csv into multiple csvs of equal numbers of records,
    except the last file.

    Includes the initial header row in each split file.

    Split files follow a zero-index sequential naming convention like so:

        `{split_file_prefix}_0.csv`
    """
    if records_per_file <= 0:
        raise Exception('records_per_file must be > 0')

    with open(source_filepath, 'r') as source:
        reader = csv.reader(source)
        headers = next(reader)

        file_idx = 0
        records_exist = True

        while records_exist:

            i = 0
            target_filename = f'{split_file_prefix}_{file_idx}.csv'
            target_filepath = os.path.join(dest_folder, target_filename)

            with open(target_filepath, 'w', newline='') as target:
                writer = csv.writer(target)

                while i < records_per_file:
                    if i == 0:
                        writer.writerow(headers)

                    try:
                        writer.writerow(next(reader))
                        i += 1
                    except:
                        records_exist = False
                        break

            if i == 0:
                # we only wrote the header, so delete that file
                os.remove(target_filepath)

            file_idx += 1


#main program
splitName = "split_email_list"
numRecords = input("Please indicate the number of records per split file: ")
numRecords = int(numRecords)

print("\nSelect file you wish to import. Make sure it is a .csv file.")
fileSource = select_file("csv")
print("\nSelect folder you wish to output files to.")
destPath = select_file("dir")

split_csv(fileSource, destPath, splitName, numRecords)

print("\nFile successfully split!")
