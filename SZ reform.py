import customtkinter as CTk
from PIL import Image, ImageTk
import os.path

file_path = os.path.dirname(os.path.realpath((__file__)))
image_1 = CTk.CTkImage(Image.open(file_path + "/dark.png"), size=(35, 35))
image_2 = CTk.CTkImage(Image.open(file_path + "/light.png"), size=(35, 35))

def on_document_sz_select():
    print("Выбрана служебка")
    # addr_entry.configure(state="disabled")
    # answer_entry.configure(state="disabled")
    # male_radio.configure(state="disabled")
    # female_radio.configure(state="disabled")
    CTk.set_appearance_mode("Dark")
    CTk.set_default_color_theme("blue")

    root = CTk.CTk()
    root.geometry("400x750")
    root.resizable(width=False, height=False)
    root.title("Генератор документов")

    type_label = CTk.CTkLabel(root, text="Выберите вид документа:")
    type_label.pack()

    address_label = CTk.CTkLabel(root, text="Выберите адресата:")
    address_label.pack()

    gender_var = CTk.StringVar()
    gender_var.set("Мужской")
    male_radio = CTk.CTkRadioButton(root, text="Мужской", variable=gender_var, radiobutton_width=15,
                                              radiobutton_height=15, value="Мужской")
    male_radio.pack()
    female_radio = CTk.CTkRadioButton(root, text="Женский", variable=gender_var, radiobutton_width=15,
                                                radiobutton_height=15, value="Женский")
    female_radio.pack()

    address_var = CTk.StringVar()
    address_combobox = CTk.CTkComboBox(root, width=300, variable=address_var,
                                                 values=["Генеральный директор", "Главный инженер",
                                                         "Директор по безопасности"])
    address_combobox.pack()

    theme_label = CTk.CTkLabel(root, text="Введите тему:")
    theme_label.pack()

    theme_entry = CTk.CTkEntry(root, width=300)
    theme_entry.pack()

    answer_label = CTk.CTkLabel(root, text="Введите номер и дату входящего письма:")
    answer_label.pack()

    answer_entry = CTk.CTkEntry(root, width=300)
    answer_entry.pack()

    addr_label = CTk.CTkLabel(root, text="Введите адрес получателя:")
    addr_label.pack()

    addr_entry = CTk.CTkEntry(root, width=300)
    addr_entry.pack()

    content_label = CTk.CTkLabel(root, text="Введите содержание:")
    content_label.pack()

    def on_paste():
        content_entry.paste()

    content_entry = CTk.CTkTextbox(root, height=200, width=300)
    content_entry.bind('<Control-v>', on_paste)
    content_entry.pack()

    save_path_label = CTk.CTkLabel(root, text="Выберите место сохранения:")
    save_path_label.pack()

    save_path_entry = CTk.CTkEntry(root, width=300)
    save_path_entry.pack()

    def select_save_path():
        save_path = CTk.filedialog.asksaveasfilename(defaultextension=".docx",
                                                               filetypes=[("Word files", "*.docx")])
        save_path_entry.delete(0, "end")
        save_path_entry.insert(0, save_path)

    select_button = CTk.CTkButton(root, text="Выбрать место сохранения", command=select_save_path)
    select_button.pack()
    select_button.place(x=110, y=615)

    generate_button = CTk.CTkButton(root, text="Сформировать документ" """command=generate_document""")
    generate_button.pack()
    generate_button.place(x=117, y=650)

    def set_light_theme():
        CTk.set_appearance_mode("Dark")
        dark_on = CTk.CTkButton(root, width=35, height=35, text="", command=set_dark_theme, image=image_2)
        dark_on.pack()
        dark_on.place(x=340, y=700)

    def set_dark_theme():
        CTk.set_appearance_mode("light")
        dark_off = CTk.CTkButton(root, width=35, height=35, text="", command=set_light_theme, image=image_1)
        dark_off.pack()
        dark_off.place(x=340, y=700)

    dark_on = CTk.CTkButton(root, width=35, height=35, text="", command=set_dark_theme, image=image_2)
    dark_on.pack()
    dark_on.place(x=340, y=700)

    root.mainloop()

on_document_sz_select()