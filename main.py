import extract_msg
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import datetime
from funcs import *



root = Tk()
root.title('Code finder')
root.geometry("500x500")


def browse_button():
    # Allow user to select a directory and store it in global var
    # called dir
    global dir
    filename = filedialog.askdirectory()
    dir.set(filename)



def retrieve_input():
    ext_list = ['xlsx', 'xls', 'csv', 'txt']
    global codes_list
    codes_list = []
    global req_code
    req_code = inputtxt.get("1.0", "end-1c")
    codes_list = [code.strip() for code in req_code.split(",")]
    global dir_file_names
    dir_file_names = glob.glob(dir.get() + "/*msg")
    global output
    output = dir.get() + "/output"
    if not os.path.isdir(output):
        os.makedirs(output)

    for name in dir_file_names:
        msg = extract_msg.openMsg(name)
        msg.saveAttachments(customPath=output, skipHidden=True)

    global attached_names
    attached_names = glob.glob(output + "/*")
    if button3['state'] == DISABLED:
        button3['state'] = NORMAL

    return attached_names




style = ttk.Style(root)
style.theme_use('clam')
style.configure('TLabel', font=('Arial', 12), foreground='black')
style.configure('TButton', font=('Arial', 12), foreground='black', background='#d9d9d9', borderwidth=0)
style.map('TButton', background=[('active', '#ececec')])


dir = StringVar()
lbl1 = Label(master=root, textvariable=dir)
lbl1.grid(row=0, column=1)
button2 = Button(text="Browse", command=browse_button)
button2.grid(row=0, column=3)
inputtxt = Text(root, width=20, height=5)
inputtxt.grid(row=1, column=1)
button = Button(text="Confirm codes", command=retrieve_input)
button.grid(row=1, column=3)



def copy_paste_data():
    start_time = datetime.datetime.now()
    counter = 0
    with open(dir.get() + "/out.txt", "w") as f1:
        for name in attached_names:
            ext = os.path.splitext(name)[-1].lower()
            if ext == ".xls":
                for code in codes_list:
                    copy_rows_with_text(name, code)
                    for row in rows_list_xls:
                        f1.write(row + "\n")
                    rows_list_xls.clear()
            elif ext == ".csv":
                for code in codes_list:
                    copy_csv_row(name, code)
                    for row in rows_csv:
                        f1.write(row + "\n")
                    rows_csv.clear()
            elif ext == ".xlsx":
                for code in codes_list:
                    copy_row_xlsx(name, code)
                    for row in rows_xlsx:
                        f1.write(row + "\n")
                    rows_xlsx.clear()
            elif ext == ".txt":
                for code in codes_list:
                    copy_txt_row(name, code)
                    for row in rows_txt:
                        f1.write(row + "\n")
                    rows_txt.clear()
            counter += 1
            if counter == len(attached_names):
                msg_finished()
            else:
                pass
        end_time = datetime.datetime.now()
        print(end_time - start_time)
    button3['state'] == DISABLED

button3 = Button(text="Start",command=copy_paste_data, state=DISABLED)
button3.grid(row=2, column=3)


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)
mainloop()