import customtkinter as CTk
from docx import Document
from docx.shared import Pt
import CTkMessagebox
import os.path
import Second


def on_document_sz_select(root, second, fio, phn):
    second.withdraw()
    CTk.set_default_color_theme("blue")

    third = CTk.CTk()
    third.geometry(
        "400x750+{}+{}".format(int(third.winfo_screenwidth() / 2 - 150), int(third.winfo_screenheight() / 2 - 300)))
    third.resizable(width=False, height=False)
    third.title("Служебная записка")

    address_label = CTk.CTkLabel(third, text="Выберите адресата:")
    address_label.pack()

    address_var = CTk.StringVar()
    address_combobox = CTk.CTkComboBox(third, width=300, variable=address_var,
                                       values=["Генеральный директор", "Главный инженер",
                                               "Директор по безопасности"])
    address_combobox.pack()

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

    def sz_reform_def():
        theme = theme_entry.get()
        choice = address_combobox.get()
        content = content_entry.get("1.0", "end-1c")
        save_path = save_path_entry.get()
        doc = Document('СЗ_шаблон.docx')
        section = doc.sections[-1]
        footer = section.footer
        footer_para = footer.paragraphs[-1]
        footer_para.text = "Исп. " + fio.get() + ' ' + 'тел. ' + phn.get()
        run = footer_para.runs[0]
        run.font.name = "Arial"
        run.font.size = Pt(10)
        doc.save('Temp.docx')

        docu = Document('Temp.docx')

        for table in docu.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if "(Адресат)" in paragraph.text:
                            if choice == "Генеральный директор":
                                paragraph.text = paragraph.text.replace("(Адресат)", "Генеральному директору")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)
                            elif choice == "Главный инженер":
                                paragraph.text = paragraph.text.replace("(Адресат)", "Главному инженеру")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)
                            elif choice == "Директор по безопасности":
                                paragraph.text = paragraph.text.replace("(Адресат)", "Директору по безопасности")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)

                        if "(Имя и Отчество)" in paragraph.text:
                            if choice == "Генеральный директор":
                                paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Павел Гаврилович")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)
                            elif choice == "Главный инженер":
                                paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Денис Александрович")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)
                            elif choice == "Директор по безопасности":
                                paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Андрей Владимирович")
                                run = paragraph.runs[0]
                                run.font.name = "Arial"
                                run.font.size = Pt(14)

                        if "(Тема)" in paragraph.text:
                            print('найдена тема')
                            paragraph.text = paragraph.text.replace("(Тема)", theme)
                            run = paragraph.runs[0]
                            run.font.name = "Arial"
                            run.font.size = Pt(14)

        # Добавим замену для текста, не находящегося в таблице
        for paragraph in docu.paragraphs:
            if choice == "Генеральный директор":
                paragraph.text = paragraph.text.replace("(Адресат)", "Генеральному директору")
                run = paragraph.runs[0]
                run.font.name = "Arial"
                run.font.size = Pt(14)
            elif choice == "Главный инженер":
                paragraph.text = paragraph.text.replace("(Адресат)", "Главный инженер")
                run = paragraph.runs[0]
                run.font.name = "Arial"
                run.font.size = Pt(14)
            elif choice == "Директор по безопасности":
                paragraph.text = paragraph.text.replace("(Адресат)", "Директор по безопасности")
                run = paragraph.runs[0]
                run.font.name = "Arial"
                run.font.size = Pt(14)
            if "(Имя и Отчество)" in paragraph.text:
                if choice == "Генеральный директор":
                    paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Павел Гаврилович")
                    run = paragraph.runs[0]
                    run.font.name = "Arial"
                    run.font.size = Pt(14)
                elif choice == "Главный инженер":
                    paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Денис Александрович")
                    run = paragraph.runs[0]
                    run.font.name = "Arial"
                    run.font.size = Pt(14)
                elif choice == "Директор по безопасности":
                    paragraph.text = paragraph.text.replace("(Имя и Отчество)", "Андрей Владимирович")
                    run = paragraph.runs[0]
                    run.font.name = "Arial"
                    run.font.size = Pt(14)
                if "(Тема)" in paragraph.text:
                    paragraph.text = paragraph.text.replace("(Тема)", theme)
                    run = paragraph.runs[0]
                    run.font.name = "Arial"
                    run.font.size = Pt(14)

            if "(Содержание)" in paragraph.text:
                paragraph.text = paragraph.text.replace("(Содержание)", content)
                run = paragraph.runs[0]
                run.font.name = "Arial"
                run.font.size = Pt(14)

        docu.save(save_path)
        os.remove("Temp.docx")
        msg = CTkMessagebox.CTkMessagebox(
            title="Docs Generator",
            message="Документ успешно сформирован! Хотите сформировать еще?",
            icon='check',
            option_1='Да',
            option_2='Нет')

        if msg.get() == 'Да':
            third.destroy()
            Second.second_tile(root, fio, phn)
        if msg.get() == 'Нет':
            root.destroy()
            second.destroy()
            third.destroy()

    generate_button = CTk.CTkButton(third, text="Сформировать документ", command=sz_reform_def)
    generate_button.pack()
    generate_button.place(x=117, y=650)

    select_button = CTk.CTkButton(third, text="Выбрать место сохранения", command=select_save_path)
    select_button.pack()
    select_button.place(x=110, y=615)

    third.mainloop()
