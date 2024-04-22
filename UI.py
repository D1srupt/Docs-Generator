import customtkinter as CTk
import Second


def disactive(event):
    if fio.get() != "" and phn.get() != "":
        ok_but.configure(state="normal", command=lambda: Second.second_tile(root))
    else:
        ok_but.configure(state="disabled")


CTk.set_appearance_mode("Dark")
CTk.set_default_color_theme("blue")

root = CTk.CTk()
root.geometry("250x250+{}+{}".format(int(root.winfo_screenwidth() / 2 - 150),
                                     int(root.winfo_screenheight() / 2 - 100)))
root.resizable(width=False, height=False)
root.title("Генератор документов")

ispol = CTk.CTkLabel(root, text='Введите данные исполнителя:')
ispol.pack()

fiol = CTk.CTkLabel(root, text='ФИО')
fiol.pack()

fio = CTk.CTkEntry(root, width=220)
fio.pack()

ph = CTk.CTkLabel(root, text='Номер рабочего телефона')
ph.pack()

phn = CTk.CTkEntry(root, width=220)
phn.pack()

ok_but = CTk.CTkButton(root, text='ОК')
ok_but.configure(state='disabled')
ok_but.pack()
ok_but.place(x=60, y=180)

fio.bind('<KeyRelease>', disactive)
phn.bind('<KeyRelease>', disactive)

root.mainloop()
