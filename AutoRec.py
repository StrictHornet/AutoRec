import tkinter as tk

win = tk.Tk()
win.title("AutoRec")

frm_head = tk.Frame(bg="yellow")
frm_head.pack(fill=tk.X)

lbl_greeting = tk.Label(relief=tk.RIDGE, height=2, bg="red",text="Inq SLA Monthly Reconciler",
                            master=frm_head)
lbl_greeting.pack(fill=tk.X)

ent1 = tk.Entry()
ir21 = ent1.get()

ent2 = tk.Entry()
ir22 = ent2.get()

ent1.pack()
ent2.pack()

rec = tk.Button(text="Reconcile", bg="red", width=15, height=3)
rec.pack()

win.mainloop()