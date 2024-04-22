import customtkinter as CTk


def second_tile(root):
    root.destroy()
    CTk.set_appearance_mode("Dark")
    CTk.set_default_color_theme("blue")
    second = CTk.CTk()
    second.geometry(
        "250x250+{}+{}".format(int(second.winfo_screenwidth() / 2 - 150), int(second.winfo_screenheight() / 2 - 100)))
    second.resizable(width=False, height=False)
    second.title("Генератор документов")

    typedoc = CTk.CTkLabel(second, text='Выберите тип документа')
    typedoc.pack()

    second.mainloop()
