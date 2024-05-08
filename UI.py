import sys
import customtkinter as CTk
import Second
import os.path
from PIL import Image

file_path = os.path.dirname(os.path.realpath(__file__))
image_1 = CTk.CTkImage(Image.open(file_path + "/dark.png"), size=(25, 25))
image_2 = CTk.CTkImage(Image.open(file_path + "/light.png"), size=(25, 25))


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

    fiol = CTk.CTkLabel(root, text='ФИО:')
    fiol.pack()

    fio = CTk.CTkEntry(root, width=220, placeholder_text='Иванов И.И.', corner_radius=4, border_width=1,
                       border_color="#1f6aa5")
    fio.pack()

    ph = CTk.CTkLabel(root, text='Номер рабочего телефона:')
    ph.pack()

    phn = CTk.CTkEntry(root, width=220, placeholder_text='1111', corner_radius=4, border_width=1,
                       border_color="#1f6aa5")
    phn.pack()

    ok_but = CTk.CTkButton(root, text='Далее')
    ok_but.configure(state='disabled')
    ok_but.pack()
    ok_but.place(x=60, y=160)

    def set_light_theme():
        CTk.set_appearance_mode("Dark")
        dark_on = CTk.CTkButton(root, width=25, height=25, text="", command=set_dark_theme, image=image_2)
        dark_on.pack()
        dark_on.place(x=200, y=210)

    def set_dark_theme():
        CTk.set_appearance_mode("light")
        dark_off = CTk.CTkButton(root, width=25, height=25, text="", command=set_light_theme, image=image_1)
        dark_off.pack()
        dark_off.place(x=200, y=210)

    dark_on = CTk.CTkButton(root, width=25, height=25, text="", command=set_dark_theme, image=image_2)
    dark_on.pack()
    dark_on.place(x=200, y=210)

    disrupt = CTk.CTkLabel(root, text='Developed by E.Skorobogatko', width=20, height=20)
    disrupt.pack()
    disrupt.place(x=10, y=220)

    fio.bind('<KeyRelease>', disactive)
    phn.bind('<KeyRelease>', disactive)

    def on_closing():
        sys.exit()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


create()
