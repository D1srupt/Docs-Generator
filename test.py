import customtkinter

customtkinter.set_default_color_theme("RVK.json")
root = customtkinter.CTk()
root.iconbitmap('C:/Users/e.skorobogatko/PycharmProjects/Docs Generator/капля.ico')


root.geometry("250x250+{}+{}".format(int(root.winfo_screenwidth() / 2 - 150),
                                     int(root.winfo_screenheight() / 2 - 100)))
root.resizable(width=False, height=False)

btn = customtkinter.CTkButton(root, text='Test')
btn.pack()
btn.place(x=50, y=100)

ent = customtkinter.CTkEntry(root)
ent.pack()


root.mainloop()
