import tkinter as tk

win = tk.Tk()
win.title("AutoRec")

frm_head = tk.Frame(bg="yellow")
frm_head.pack(fill=tk.X)

lbl_greeting = tk.Label(relief=tk.RIDGE, height=2, bg="red",text="Inq SLA Monthly Reconciler",
                            master=frm_head)
lbl_greeting.pack(fill=tk.X)

frm_body = tk.Frame()
frm_body.pack(padx=15, pady=15)

lbl_val1 = tk.Label(text = "Value 1", master = frm_body)
lbl_val1.grid(row=0, column=0)

lbl_val2 = tk.Label(text = "Value 2", master = frm_body)
lbl_val2.grid(row=1, column=0)

ent1 = tk.Entry(master = frm_body)
ir21 = ent1.get()

ent2 = tk.Entry(master = frm_body)
ir22 = ent2.get()

ent1.grid(row=0, column=1)
ent2.grid(row=1, column=1)

lbl_directory = tk.Label(width=35, bg="gray", padx=10, text="DIRECTORY PATH")
lbl_directory.pack()

rec = tk.Button(text="Reconcile", bg="red", width=15, height=3)
rec.pack()

win.mainloop()