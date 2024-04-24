import customtkinter as CTk
from PIL import Image, ImageTk
import os.path


def on_document_prikaz_select(second):
    second.withdrow()
    file_path = os.path.dirname(os.path.realpath((__file__)))
    image_1 = CTk.CTkImage(Image.open(file_path + "/dark.png"), size=(35, 35))
    image_2 = CTk.CTkImage(Image.open(file_path + "/light.png"), size=(35, 35))
    print("Выбран приказ")
    # addr_entry.configure(state="disabled")
    # answer_entry.configure(state="disabled")
    # male_radio.configure(state="disabled")
    # female_radio.configure(state="disabled")
    CTk.set_appearance_mode("Dark")
    CTk.set_default_color_theme("blue")

    third = CTk.CTk()
    third.geometry(
        "400x750+{}+{}".format(int(third.winfo_screenwidth() / 2 - 150), int(third.winfo_screenheight() / 2 - 300)))
    third.resizable(width=False, height=False)
    third.title("Приказ")

    theme_label = CTk.CTkLabel(third, text="Введите тему:")
    theme_label.pack()

    theme_entry = CTk.CTkEntry(third, width=300)
    theme_entry.pack()

    content_label = CTk.CTkLabel(third, text="Введите содержание:")
    content_label.pack()

    def on_paste():
        content_entry.paste()

    content_entry = CTk.CTkTextbox(third, height=200, width=300)
    content_entry.bind('<Control-v>', on_paste)
    content_entry.pack()

    save_path_label = CTk.CTkLabel(third, text="Выберите место сохранения:")
    save_path_label.pack()

    save_path_entry = CTk.CTkEntry(third, width=300)
    save_path_entry.pack()

    def select_save_path():
        save_path = CTk.filedialog.asksaveasfilename(defaultextension=".docx",
                                                     filetypes=[("Word files", "*.docx")])
        save_path_entry.delete(0, "end")
        save_path_entry.insert(0, save_path)

    select_button = CTk.CTkButton(third, text="Выбрать место сохранения", command=select_save_path)
    select_button.pack()
    select_button.place(x=110, y=615)

    generate_button = CTk.CTkButton(third, text="Сформировать документ")
    generate_button.pack()
    generate_button.place(x=117, y=650)

    def set_light_theme():
        CTk.set_appearance_mode("Dark")
        dark_on = CTk.CTkButton(third, width=35, height=35, text="", command=set_dark_theme, image=image_2)
        dark_on.pack()
        dark_on.place(x=340, y=700)

    def set_dark_theme():
        CTk.set_appearance_mode("light")
        dark_off = CTk.CTkButton(third, width=35, height=35, text="", command=set_light_theme, image=image_1)
        dark_off.pack()
        dark_off.place(x=340, y=700)

    dark_on = CTk.CTkButton(third, width=35, height=35, text="", command=set_dark_theme, image=image_2)
    dark_on.pack()
    dark_on.place(x=340, y=700)

    third.mainloop()