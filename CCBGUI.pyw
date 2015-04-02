import tkinter as tk
import MailingList as mail
import CustomList as custom


def run_full_list():
    mail.main()
    root.destroy()


def run_custom_list():
    custom.main()
    root.destroy()


root = tk.Tk()
root.title("CCB Automation")
label = tk.Label(root, text="CCB API", fg="black")
label.pack()
btWidth = 50
btHeight = 5
button = tk.Button(root, text='Export Names', width=btWidth, height=btHeight, command=run_full_list)
button_custom = tk.Button(root, text='Custom Search Export', width=btWidth, height=btHeight, command=run_custom_list)
button_close = tk.Button(root, text='Close App', width=btWidth, height=btHeight, command=root.destroy)
button.pack()
button_custom.pack()
button_close.pack()
root.mainloop()
