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

    typedoc = CTk.CTkLabel(second, text='Выберите тип документа')
    typedoc.pack()

    sz = CTk.CTkButton(second, text='Служебная записка',
                       command=lambda: SZ_reform.on_document_sz_select(root, second, fio, phn))
    sz.pack()
    sz.place(x=55, y=40)

    prikaz = CTk.CTkButton(second, text='Приказ',
                           command=lambda: Prikaz_reform.on_document_prikaz_select(root, second, fio, phn))
    prikaz.pack()
    prikaz.place(x=55, y=110)

    pismo = CTk.CTkButton(second, text='Письмо',
                          command=lambda: Pismo_reform.on_document_pismo_select(root, second, fio, phn))
    pismo.pack()
    pismo.place(x=55, y=180)

    second.mainloop()
