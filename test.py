# test_ctk.py
import customtkinter as ctk
ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.geometry("400x200")
ctk.CTkLabel(app, text="CTk OK").pack(pady=40)
app.mainloop()
