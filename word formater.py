import customtkinter as ctk
from PIL import Image, ImageTk
import grid
from spire.doc import *
from spire.doc.common import *
import ctypes
import pywinstyles
import time as t

user32 = ctypes.windll.user32
screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)




def word_function(modules):
    
    doc=Document()

    section = doc.AddSection()

    section.PageSetup.PageSize = PageSize.A4()
    section.PageSetup.Margins.All = 36.1

    tstyle=ParagraphStyle(doc)
    tstyle.Name = "titlestyle"
    tstyle.CharacterFormat.Bold=True
    tstyle.CharacterFormat.FontName = "Times New Roman"
    tstyle.CharacterFormat.FontSize = 18
    doc.Styles.Add(tstyle)

    hstyle=ParagraphStyle(doc)
    hstyle.Name = "headingstyle"
    hstyle.CharacterFormat.Bold=True
    hstyle.CharacterFormat.FontName = "Times New Roman"
    hstyle.CharacterFormat.FontSize = 12
    doc.Styles.Add(hstyle)

    astyle=ParagraphStyle(doc)
    astyle.Name = "authordetailstyle"
    astyle.CharacterFormat.FontName = "Times New Roman"
    astyle.CharacterFormat.FontSize = 12
    doc.Styles.Add(astyle)

    pstyle=ParagraphStyle(doc)
    pstyle.Name = "parastyle"
    pstyle.CharacterFormat.FontName = "Times New Roman"
    pstyle.CharacterFormat.FontSize = 12
    doc.Styles.Add(pstyle)
    
    titlepara = section.AddParagraph()
    titlepara.AppendText(modules[0])
    titlepara.ApplyStyle("titlestyle")
    titlepara.Format.HorizontalAlignment = HorizontalAlignment.Center
    titlepara.Format.AfterSpacing = 10

    namepara=section.AddParagraph()
    namepara.AppendText(modules[1])
    namepara.ApplyStyle("authordetailstyle")
    namepara.Format.HorizontalAlignment = HorizontalAlignment.Center
    namepara.Format.AfterSpacing = 10

    addrpara=section.AddParagraph()
    addrpara.AppendText(modules[2])
    addrpara.ApplyStyle("authordetailStyle")
    addrpara.Format.HorizontalAlignment = HorizontalAlignment.Center
    addrpara.Format.AfterSpacing = 10

    mailpara=section.AddParagraph()
    mailpara.AppendText(modules[3])
    mailpara.ApplyStyle("authordetailStyle")
    mailpara.Format.HorizontalAlignment = HorizontalAlignment.Center
    mailpara.Format.AfterSpacing = 10

    abstracthead=section.AddParagraph()
    abstracthead.AppendText("Abstract")
    abstracthead.ApplyStyle("headingstyle")
    abstracthead.Format.HorizontalAlignment = HorizontalAlignment.Justify
    abstracthead.Format.AfterSpacing = 8

    abstractpara=section.AddParagraph()
    abstractpara.AppendText(modules[4])
    abstractpara.ApplyStyle("parastyle")
    abstractpara.Format.HorizontalAlignment = HorizontalAlignment.Justify
    abstractpara.Format.AfterSpacing = 8

    doc.SaveToFile(r"C:\Users\athar\Desktop\Projects\Python Projects\Word formater\test.docx",FileFormat.Docx2019)
    doc.Close()


#Done Module
#----------------------------------------------------------------------------------
def done_module():
    done_dialog_box=ctk.CTkToplevel(app)
    done_dialog_box.geometry("300x150+{a}+{b}".format(a=int((screensize[0]/2)-150),b=int((screensize[1]/2)-75)))
    done_dialog_box.title("Task Complete")
    done_label=ctk.CTkLabel(master=done_dialog_box,text="Done",font=heading_font)
    done_label.place(x=130,y=50)

    block_frame=ctk.CTkFrame(master=app, width=screensize[0], height=screensize[1], border_width=1,border_color="black")
    block_frame.place(x=0,y=0)
    pywinstyles.set_opacity(block_frame,value=0.5)
    done_dialog_box.transient(app)
    done_dialog_box.grab_set()
    app.wait_window(done_dialog_box)
    block_frame.destroy()
#----------------------------------------------------------------------------------

#Submit_button
#----------------------------------------------------------------------------------
def submit_func():
    modules=[]
    modules.append(title_input.get())
    modules.append(author_name_input.get())
    modules.append(author_addr_input.get())
    modules.append(author_mail_input.get())
    modules.append(abstract_input.get())
    word_function(modules)
    done_module()
#----------------------------------------------------------------------------------
#Add function for new titles
#----------------------------------------------------------------------------------
def add_heading_dialog_box():
    #main window
    #-----------------------------------------------------------
    add_dialog_box=ctk.CTkToplevel(app)
    add_dialog_box.geometry("620x240+{a}+{b}".format(a=int((screensize[0]/2)-310),b=int((screensize[1]/2)-120)))
    add_dialog_box.title("Add Heading/Paragraph")
    add_dialog_box.transient(app)
    add_dialog_box.grab_set()
    #add_dialog_box.overrideredirect(True)
    #-----------------------------------------------------------
    #block frame on root_application when add window starts
    #-----------------------------------------------------------
    block_frame=ctk.CTkFrame(master=app, width=screensize[0], height=screensize[1], border_width=1,border_color="black")
    block_frame.place(x=0,y=0)
    pywinstyles.set_opacity(block_frame,value=0.5)
    #----------------------------------------------------------
    #functions for checkbox widets
    #----------------------------------------------------------
    heading_checkvar=ctk.StringVar(value="off")
    def heading_func():
        print(heading_checkvar.get())
    #----------------------------------------------------------
    # buttons and event boxs
    #----------------------------------------------------------
    heading_checkbox=ctk.CTkCheckBox(master=add_dialog_box, text="Heading", command=heading_func, variable=heading_checkvar,onvalue="on",offvalue="off")
    heading_checkbox.place(x=20,y=20)
    #----------------------------------------------------------
    app.wait_window(add_dialog_box)
    block_frame.destroy()
    
#----------------------------------------------------------------------------------


#GUI Section
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    
ctk.set_appearance_mode('dark')
ctk.set_default_color_theme("dark-blue")

app=ctk.CTk()
app.geometry("{a}x{b}+0+0".format(a=screensize[0],b=screensize[1]))
app.title("Word Format Wizard")


heading_font=ctk.CTkFont(family="Arial Black",size=21)
subheading_font=ctk.CTkFont(family="Arial Black",size=18)

background_image = ctk.CTkImage(light_image=Image.open(r"C:\Users\athar\Desktop\Projects\Python Projects\Word formater\background.png"),
                                dark_image=Image.open(r"C:\Users\athar\Desktop\Projects\Python Projects\Word formater\background.png"),
                                size=(1920,1080))

Background_frame = ctk.CTkLabel(app, text="",image=background_image)
Background_frame.pack(pady=10)



frame1 = ctk.CTkFrame(master=app, width = 1280, height = 820)
frame1.place(x=320,y=130)


x=20
y=20

#Artite Title
title_label=ctk.CTkLabel(master=frame1,text="Article Title",font=heading_font)
title_label.place(x=x+550,y=y)
y+=30
title_input=ctk.CTkEntry(master=frame1,placeholder_text="Enter Title of Article",width=1220)
title_input.place(x=x,y=y)
y+=60


#Author Details
author_detail_label=ctk.CTkLabel(master=frame1, text="Author Details",font=heading_font)
author_detail_label.place(x=x+545,y=y)
y+=40

author_name_label=ctk.CTkLabel(master=frame1, text="Author Name(s)",font=subheading_font)
author_name_label.place(x=x,y=y)
y+=30
author_name_input=ctk.CTkEntry(master=frame1, placeholder_text="Enter Author Name(s), Note: If  multiple seprate by commas (,)",width=1220)
author_name_input.place(x=x,y=y)
y+=50

author_addr_label=ctk.CTkLabel(master=frame1, text="Author Address(s)",font=subheading_font)
author_addr_label.place(x=x,y=y)
y+=30
author_addr_input=ctk.CTkEntry(master=frame1, placeholder_text="Enter Author Address(s), Note: Follow the sequence of Author Names",width=1220, height=70)
author_addr_input.place(x=x,y=y)
y+=90

author_mail_label=ctk.CTkLabel(master=frame1,text="Author Mail(s)",font=subheading_font)
author_mail_label.place(x=x,y=y)
y+=30
author_mail_input=ctk.CTkEntry(master=frame1,placeholder_text="Enter Author Mail(s), Note: Follow the sequence of Author Names",width=1220)
author_mail_input.place(x=x,y=y)
y+=50


#Article body
y+=40
article_body_label=ctk.CTkLabel(master=frame1, text="Article Body", font=heading_font)
article_body_label.place(x=x+545,y=y)
y+=40

abstract_label=ctk.CTkLabel(master=frame1, text="Abstract", font=subheading_font)
abstract_label.place(x=x,y=y)
y+=30
abstract_input=ctk.CTkEntry(master=frame1, placeholder_text="Enter Article Abstract",width=1220, height=70)
abstract_input.place(x=x,y=y)
y+=100



#Buttons and scrolls
#---------------------------------------------------------------------------------------------------------------------------------------
add_button=ctk.CTkButton(master=frame1,text="Add", fg_color=("#cacaca","#262626"),hover_color="#6befff", corner_radius=10, command=add_heading_dialog_box)
add_button.place(x=x,y=770)

submit_button=ctk.CTkButton(master=frame1,text="Submit", fg_color=("#539dfa","#539dfa"),hover_color="#f9990f", corner_radius=10, command=submit_func)
submit_button.place(x=x+170,y=770)
#---------------------------------------------------------------------------------------------------------------------------------------

#print(y)
#grid.grid(app)

app.mainloop()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
