import customtkinter as ctk
import ctypes
from PIL import Image, ImageTk
import threading as th
import os

user32 = ctypes.windll.user32
screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


stop_splash=False
def root_application():
    splash_screen.withdraw()
    os.system('start root.bat')
    exit()

#Splash Screen
#--------------------------------------------------------------------------------------------------------------------------

splash_screen=ctk.CTk()
splash_screen.geometry("400x200+{a}+{b}".format(a=(int(screensize[0]/2)-200),b=(int(screensize[1]/2)-100)))
splash_screen.resizable(0,0)
splash_screen.overrideredirect(True)
splash_screen.attributes('-topmost', True) 

splash_heading_font=ctk.CTkFont(family="Alphacorsa",size=35)
splash_loading_font=ctk.CTkFont(family="Arial", size=11)

software_name_label=ctk.CTkLabel(master=splash_screen, text="Word Format Wizard", font=splash_heading_font)
software_name_label.place(x=40,y=80)

loading_label=ctk.CTkLabel(master=splash_screen, text="Loading...", font=splash_loading_font)
loading_label.place(x=20,y=165)

c2=ctk.CTkImage(light_image=Image.open(r'c2.png'),dark_image=Image.open(r'c2.png'),size=(20,12))
c1=ctk.CTkImage(light_image=Image.open(r'c1.png'),dark_image=Image.open(r'c1.png'),size=(20,12))


def animation1(i):
    #print(i)
    l1=ctk.CTkLabel(splash_screen, text="", image=c1).place(x=150,y=115)
    l2=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=170,y=115)
    l3=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=190,y=115)
    l4=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=210,y=115)
    splash_screen.update_idletasks()
    splash_screen.after(500, animation2,i)

def animation2(i):
    l1=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=150,y=115)
    l2=ctk.CTkLabel(splash_screen, text="", image=c1).place(x=170,y=115)
    l3=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=190,y=115)
    l4=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=210,y=115)
    splash_screen.update_idletasks()
    splash_screen.after(500, animation3,i)

def animation3(i):
    l1=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=150,y=115)
    l2=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=170,y=115)
    l3=ctk.CTkLabel(splash_screen, text="", image=c1).place(x=190,y=115)
    l4=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=210,y=115)
    splash_screen.update_idletasks()
    splash_screen.after(500, animation4,i)

def animation4(i):
    l1=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=150,y=115)
    l2=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=170,y=115)
    l3=ctk.CTkLabel(splash_screen, text="", image=c2).place(x=190,y=115)
    l4=ctk.CTkLabel(splash_screen, text="", image=c1).place(x=210,y=115)
    splash_screen.update_idletasks()
    splash_screen.after(500, animation1,i+1)

splash_screen.update_idletasks()
root_start=th.Timer(3.0,root_application)
root_start.start()
animation1(0)

splash_screen.mainloop()
#--------------------------------------------------------------------------------------------------------------------------
