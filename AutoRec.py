import tkinter as tk
from openpyxl import Workbook
from openpyxl import load_workbook

############# TASKS ###################
# PUTTING EVERYTHING IN ONE CONSTANT FILE
# CREATE NEW FILE, COPY REC FILE, REC
# GUI
# DUPLICATE RESULT IR
# CONVERT ALPHABET TO NUMBER
# DESIGN ERROR TEXT FIELD
# PRINT ERRORS TO TEXT FIELD

#

# #FUNCTION THAT RECONCILES SURVEY DOCUMENT#
def reconcile():
    print("")
#     workbook = load_workbook(
#     filename="C:/Users/okosu/Google Drive/Work/VODACOM MAIN COPY OF SURVEY SPREADSHEET.xlsx")

# # For loop for search and cell assignment
#     for ir in workbook["RecNew"].iter_rows(min_row=2, min_col=3):
#         # print(ir) # For debugging
#         for sheet in workbook:
#             for row in sheet.iter_rows(min_row=2, min_col=8):
#                 # print(row) # For debugging
#                 try:
#                     if ir[0].value == row[0].value:
#                         if row[6].value != "NIL":
#                             # print("Present") # For debugging
#                             # print("VBN is {} and Recon is {}".format(
#                             #     row[0].value, ir[0].value))
#                             ir[21].value = "Failed"  # 19
#                             ir[22].value = row[6].value  # 20
#                             break  # Since IR has been found loop should break to next IR
#                         else:
#                             ir[21].value = "OK"
#                 except:
#                     print("Didn't work!")
#             else:
#                 continue
#             break  # Break the outer loop

# # For loop for duplicate search
#     for ir in workbook["RecNew"].iter_rows(min_row=2, min_col=3):
#         for sheet in workbook:
#             for row in sheet.iter_rows(min_row=2, min_col=3):
#                 try:
#                     if ir[0].value == row[0].value:
#                         if row[21].value == "Approved":
#                             # print("Duplicate") # For debugging
#                             # print("VBN is {} and Recon is {}".format(
#                             #     row[0].value, ir[0].value))
#                             ir[21].value = "DUPLICATE"  # 19
#                             ir[22].value = row[22].value  # 20
#                             break  # Since IR has been found loop should break to next IR
#                 except:
#                     print("Didn't work!")
#             else:
#                 continue
#             break  # Break the outer loop

# workbook.save(
#     filename="C:/Users/okosu/Google Drive/Work/VODACOM MAIN COPY OF SURVEY SPREADSHEET.xlsx")

#

win = tk.Tk()
win.title("AutoRec")

frm_head = tk.Frame(bg="yellow")
frm_head.pack(fill=tk.X)

lbl_greeting = tk.Label(relief=tk.RIDGE, height=2, bg="#641822",text="Inq SLA Monthly Reconciler",
                            master=frm_head)
lbl_greeting.pack(fill=tk.X)

frm_body = tk.Frame()
frm_body.pack(padx=15, pady=15)

lbl_approvalColumn = tk.Label(text = "inq. Approval Column Alphabet Identifier", master = frm_body)
lbl_approvalColumn.grid(row=0, column=0)
ent_approvalColumn = tk.Entry(master = frm_body, width = 5)
approvalColumn = ent_approvalColumn.get()
ent_approvalColumn.grid(row=0, column=1)

lbl_commentColumn = tk.Label(text = "inq. Comment Column Alphabet Identifier", master = frm_body)
lbl_commentColumn.grid(row=1, column=0)
ent_commentColumn = tk.Entry(master = frm_body, width = 5)
commentColumn = ent_commentColumn.get()
ent_commentColumn.grid(row=1, column=1)

frm_directory = tk.Frame(padx = 10,pady = 10)

lbl_directory = tk.Label(master = frm_directory, width=35, bg="gray", padx=10, text="DIRECTORY PATH")
lbl_directory.pack()

file_directory = "C:/Users/okosu/Google Drive/Work/VODACOM MAIN COPY OF SURVEY SPREADSHEET.xlsx"
directory_ent = tk.Entry(master=frm_directory, width=55)
directory_ent.insert(0, file_directory)
file_directory = directory_ent.get()
directory_ent.pack()

frm_directory.pack()

btn_reconcile = tk.Button(text="Reconcile", bg="#a00008", width=15, height=3, command=reconcile)
btn_reconcile.pack(pady=5)

win.mainloop()