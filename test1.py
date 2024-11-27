import threading as th
import time as t
import customtkinter as ctk

def root_application():
    sc.destroy()        
    app=ctk.CTk()
    app.mainloop()


sc=ctk.CTk()
button=ctk.CTkButton(master=sc, text="Test")
button.place(x=20,y=20)


print("Before delay")
delay1=th.Timer(3.0,root_application)
delay1.start()
print("After Delay")

sc.mainloop()
