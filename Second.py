import sys

import customtkinter as CTk
import SZ_reform
import Prikaz_reform
import Pismo_reform


def second_tile(root, fio, phn):
    root.withdraw()
    CTk.set_default_color_theme("blue")
    second = CTk.CTk()
    second.geometry(
        "250x250+{}+{}".format(int(second.winfo_screenwidth() / 2 - 150),
                               int(second.winfo_screenheight() / 2 - 100)))
    second.resizable(width=False, height=False)
    second.title("Генератор документов")
    second.iconbitmap('капля.ico')

    typedoc = CTk.CTkLabel(second, text='Выберите тип документа:')
    typedoc.pack()

    sz = CTk.CTkButton(second, text='Служебная записка', corner_radius=10, width=200,
                       command=lambda: SZ_reform.on_document_sz_select(root, second, fio, phn))
    sz.pack()
    sz.place(x=30, y=55)

    prikaz = CTk.CTkButton(second, text='Приказ', corner_radius=10, width=200,
                           command=lambda: Prikaz_reform.on_document_prikaz_select(root, second, fio, phn))
    prikaz.pack()
    prikaz.place(x=30, y=100)

    pismo = CTk.CTkButton(second, text='Письмо', corner_radius=10, width=200,
                          command=lambda: Pismo_reform.on_document_pismo_select(root, second, fio, phn))
    pismo.pack()
    pismo.place(x=30, y=145)

    def back():
        second.destroy()
        root.deiconify()

    arrow = CTk.CTkButton(second, width=25, height=25, text="←", command=back)
    arrow.pack()
    arrow.place(x=15, y=215)

    def on_closing():
        sys.exit()

    second.protocol("WM_DELETE_WINDOW", on_closing)
    second.mainloop()
