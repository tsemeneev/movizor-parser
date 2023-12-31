import customtkinter
import subprocess
from main import parser
from auth import authenticate


customtkinter.set_appearance_mode("light")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green




def button_function(proxy, pages_count):
   
    proxy = proxy_entry.get()
    pages_count = pages_entry.get()
    
    label.destroy()
    button.destroy()
    pages_entry.destroy()
    pages.destroy()
    proxy_entry.destroy()
    proxy_text.destroy()
    auth_text.destroy()
    auth_button.destroy()
    
    text = customtkinter.CTkLabel(master=app, text="Парсинг начался. Не закрывайте окно.\nПосле выполения парсинга, вся информация будет отображена в этом окне",
                                   wraplength=300, 
                                   font=("Arial", 20))
    text.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
    app.update()
    
    
    result = parser(proxy, pages_count)
    if result == 'Success':
        run_bat_script()
        text.configure(text="Поздравляю! Мы успешно собрали данные.\nСохранили их в Excel файле и положили в папку Documents на С: диске.")
        app.update() 
    else:
        text.configure(text="Что-то пошло не так. Обратитесь к создателю парсера.")
        app.update()
    
    
def reauth(login_entry, password_entry):
    login = login_entry.get()
    password = password_entry.get()
    result = customtkinter.CTkLabel(master=app, text="Происходит авторизация. Не закрывайте окно", wraplength=200)
    result.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
    app.update()
    try:
        authenticate(login, password)
        result.configure(text="Авторизация прошла успешно")
        app.update()
    except Exception as e:
        result.configure(text="Что-то пошло не так. Авторизация не удалась. Ошибка: " + str(e))
        app.update()
    
    
    
    
def update_cookies():
    label.destroy()
    button.destroy()
    pages_entry.destroy()
    pages.destroy()
    proxy_entry.destroy()
    proxy_text.destroy()
    auth_text.destroy()
    auth_button.destroy()
    
    
    login_entry = customtkinter.CTkEntry(master=app)
    login_entry.place(relx=0.5, rely=0.20, anchor=customtkinter.N)
    login_text = customtkinter.CTkLabel(master=app, text="Введите логин (email)")
    login_text.place(relx=0.5, rely=0.16, anchor=customtkinter.CENTER)
    password_entry = customtkinter.CTkEntry(master=app)
    password_entry.place(relx=0.5, rely=0.35, anchor=customtkinter.N)
    password_text = customtkinter.CTkLabel(master=app, text="Введите пароль")
    password_text.place(relx=0.5, rely=0.31, anchor=customtkinter.CENTER)
    main_label = customtkinter.CTkLabel(master=app, text="Введите email и пароль авторизации сайта movizor-info.ru", font=("Arial", 15, "bold"), wraplength=300)
    main_label.place(relx=0.5, rely=0.05, anchor=customtkinter.CENTER)
    auth_button2 = customtkinter.CTkButton(master=app, text="Обновить", command=lambda: reauth(login_entry, password_entry))
    auth_button2.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
    
    
def run_bat_script():
    process = subprocess.Popen(rf'.\copy_files.bat', shell=True)
    # Вы можете выполнять другие действия во время выполнения скрипта
    process.wait() # Ожидайте завершения процесса




app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.geometry("400x600")
app.title("Парсер сайта movizor-info.ru")

# Use CTkButton instead of tkinter Button
label = customtkinter.CTkLabel(master=app, 
                               text="Перед началом парсинга укажите прокси и колличество страниц, которое нужно парсить.\nНа одной странице 10 карточек", 
                               font=("Arial", 15, "bold"),
                               wraplength=380)
label.place(relx=0.5, rely=0.12, anchor=customtkinter.CENTER)

pages_entry = customtkinter.CTkEntry(master=app)
pages_entry.place(relx=0.5, rely=0.55, anchor=customtkinter.N)


pages = customtkinter.CTkLabel(master=app, text="Сколько страниц нужно парсить?\nмин. значение - 3, макс. значение - 100")
pages.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)

proxy_entry = customtkinter.CTkEntry(master=app)
proxy_entry.place(relx=0.5, rely=0.37, anchor=customtkinter.CENTER)

proxy_text = customtkinter.CTkLabel(master=app, text="Введите прокси в поле ниже (обязательно)\nФормат: 127.0.0.1:8080")
proxy_text.place(relx=0.5, rely=0.3, anchor=customtkinter.CENTER)

auth_text = customtkinter.CTkLabel(master=app, 
                                   text="Обновить данные куки и повторно пройти авторизацию (по требованию)",
                                   wraplength=250)
auth_text.place(relx=0.5, rely=0.85, anchor=customtkinter.CENTER)
auth_button = customtkinter.CTkButton(master=app, text="Обновить", command=update_cookies)
auth_button.place(relx=0.5, rely=0.92, anchor=customtkinter.CENTER)

separator_frame = customtkinter.CTkFrame(master=app, bg_color="white", fg_color="gray", height=4, width=400)
separator_frame.place(relx=0.5, rely=0.80, anchor=customtkinter.CENTER,)


button = customtkinter.CTkButton(master=app, text="Запустить", command=lambda: button_function(proxy_entry, pages_entry))
button.place(relx=0.5, rely=0.7, anchor=customtkinter.CENTER)
app.mainloop()