import customtkinter as CTk
import Second


def create():


    CTk.set_appearance_mode("Dark")
    CTk.set_default_color_theme("blue")

    root = CTk.CTk()
    root.geometry("250x250+{}+{}".format(int(root.winfo_screenwidth() / 2 - 150),
                                         int(root.winfo_screenheight() / 2 - 100)))
    root.resizable(width=False, height=False)
    root.title("Генератор документов")

    def disactive(event):
        if fio.get() != "" and phn.get() != "":
            ok_but.configure(state="normal", command=lambda: Second.second_tile(root, fio, phn))
        else:
            ok_but.configure(state="disabled")

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

    global number, value
    number = phn.get()
    value = fio.get()

    ok_but = CTk.CTkButton(root, text='Далее')
    ok_but.configure(state='disabled')
    ok_but.pack()
    ok_but.place(x=60, y=180)

    fio.bind('<KeyRelease>', disactive)
    phn.bind('<KeyRelease>', disactive)

    root.mainloop()


create()
