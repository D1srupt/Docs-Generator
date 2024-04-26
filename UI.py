import customtkinter as CTk
import Second
import os.path
from PIL import Image

file_path = os.path.dirname(os.path.realpath(__file__))
image_1 = CTk.CTkImage(Image.open(file_path + "/dark.png"), size=(35, 35))
image_2 = CTk.CTkImage(Image.open(file_path + "/light.png"), size=(35, 35))


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

    ok_but = CTk.CTkButton(root, text='Далее')
    ok_but.configure(state='disabled')
    ok_but.pack()
    ok_but.place(x=60, y=160)

    def set_light_theme():
        CTk.set_appearance_mode("Dark")
        dark_on = CTk.CTkButton(root, width=35, height=35, text="", command=set_dark_theme, image=image_2)
        dark_on.pack()
        dark_on.place(x=190, y=200)

    def set_dark_theme():
        CTk.set_appearance_mode("light")
        dark_off = CTk.CTkButton(root, width=35, height=35, text="", command=set_light_theme, image=image_1)
        dark_off.pack()
        dark_off.place(x=190, y=200)

    dark_on = CTk.CTkButton(root, width=35, height=35, text="", command=set_dark_theme, image=image_2)
    dark_on.pack()
    dark_on.place(x=190, y=200)

    fio.bind('<KeyRelease>', disactive)
    phn.bind('<KeyRelease>', disactive)

    root.mainloop()


create()
