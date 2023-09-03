from tkinter import *
from tkinter.filedialog import *
from PIL import ImageTk
import PIL
from pyautogui import alert
import pymsgbox
import tabula
from regular import *
from Statistics import *
from User_guide import *
import os
import sys
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

master=Tk()
master.geometry("1250x900+300+0")
master.configure(bg="#1E90FF")
master.title("CGPA/SGPA Calculator")
img=PhotoImage(file=resource_path("Background.png"))
logo=PIL.Image.open(resource_path("JNTUK logo.png"))
new_logo=logo.resize((100,100))
new_logo=ImageTk.PhotoImage(new_logo)
my_canvas=Canvas(master,width=1000,height=1000)
my_canvas.pack(fill='both',expand=True)
my_canvas.create_image(0,0,image=img,anchor="nw")
my_canvas.create_image(20,20,image=new_logo,anchor="nw")
my_canvas.create_text(110,20,text="University College of Engineering Narasaraopet",anchor="nw",font=('Algerian',25),fill="Black")
my_canvas.create_text(150,60,text="Jawaharlal Nehru Technological University Kakinada",anchor="nw",font=('Algerian',20),fill="Black")
my_canvas.create_text(1100,850,text="K. Devika Yasaswi",font=("Vladimir Script",18))
root=Frame(my_canvas,padx=20,pady=20,bg="#FFE9E3")
root.pack(pady=110)
input_file=''
input_file_excel=''
Label_font=('Times new Roman',18)   #Font style for labels
Entry_font=('Times new Roman',15)   #Font style for entry boxes
pdf_type = [("pdf Files",'.*pdf')]
excel_type=[("xlsx Files",".*xlsx")]
status1=0
#File reading function for SGPA-Regular
def input_marks_file():
    global input_file
    input_file=askopenfilename(filetypes=pdf_type)

#File reading function for SGPA-Regular
def input_marks_excel():
    global input_file_excel
    input_file_excel=askopenfilename(filetypes=excel_type)

upload=Label(root,text="Upload the result pdf  ",bg="#FFE9E3",font=Entry_font)
upload.grid(row=1,column=0,sticky='w',pady=6)
upload_button=Button(root, text='Upload File', width=20,command = input_marks_file)
upload_button.grid(row=1,column=2,sticky='w')


file_error=Label(root,text='The uploaded pdf format is not suitable.',font=Entry_font,fg='red',bg="#FFE9E3")
file_error_continue=Label(root,text='So please try uploading excel.',font=Entry_font,fg='red',bg="#FFE9E3")
new_upload=Label(root,text="Upload the result excel",bg="#FFE9E3",font=Entry_font)
new_upload_button=Button(root, text='Upload File', width=20,command = input_marks_excel)
#wrong_upload=Label(root,text='The uploaded excel format is not suitable.',font=Entry_font,fg='red',bg="#FFE9E3")
wrong_file=Label(root,text='The uploaded file is wrong! Try again.',font=Entry_font,fg='red',bg="#FFE9E3")
upload_result_label=Label(root,text='Please upload the result file',font=Entry_font,fg='red',bg="#FFE9E3")
upload_regular_gpa=Label(root,text='Please upload the result excel',font=Entry_font,fg='red',bg="#FFE9E3")
#saving functionality
def save():
    global data,status1,civil_credits,mech_credits,eee_credits,ece_credits,cse_credits,input_file,GPA_file
    try:
        upload_regular_gpa.grid_forget()
    except:
        pass
    try:
        upload_result_label.grid_forget()
    except:
        pass
    try:
        file_error.grid_forget()
        file_error_continue.grid_forget()
    except:
        pass
    try:
        wrong_file.grid_forget()
    except:
        pass
    if input_file!="":
        def calculation(data):                            
            rno_list=[]
            for i in range(len(data)):
                x=int(data['Htno'][i][7:10])
                rno_list.append(data['Htno'][i][0:6])
            new_rno_list=list(set(rno_list))
            new=[]
            for i in new_rno_list:
                new.append(rno_list.count(i))
            series=new_rno_list[new.index(max(new))]
            new_df=pd.DataFrame(columns=data.columns)
            series1=str(int(series[0:2])+1)+"035A"
            for i in range(len(data)):
                if data.iloc[i,0][0:6]== series or data.iloc[i,0][0:6]==series1:
                    new_df.loc[len(new_df.index)]=list(data.iloc[i,:])
            try:
                Sgpa(new_df,input_file) 
            except ZeroDivisionError:
                wrong_file.grid(row=2,column=0,sticky='w',pady=6) 
                return                      
            pymsgbox.rootWindowPosition="+700+350"
            result=alert(text="Result Analysis file generation is completed",title="Status",button="Ok")                            
            if result=="Ok":
                master.destroy()   
        if status1==0:               
            df=tabula.read_pdf(input_file,pages="all")
            data=pd.DataFrame()
            for i in range(len(df)):
                data=pd.concat([data,df[i]],ignore_index=True)
            try:
                if data[list(data.columns)[-1]].isnull().values.any():
                    status1=1                                      
                    file_error.grid(row=2,column=0,sticky='w')    
                    file_error_continue.grid(row=2,column=2,sticky='w') 
                    upload.grid_forget()
                    upload_button.grid_forget()
                    new_upload.grid(row=1,column=0,sticky='w',pady=6)
                    new_upload_button.grid(row=1,column=2,sticky='w')                           
            except: 
                status1=1
                file_error.grid(row=2,column=0,sticky='w')    
                file_error_continue.grid(row=2,column=2,sticky='w') 
                upload.grid_forget()
                upload_button.grid_forget()
                new_upload.grid(row=1,column=0,sticky='w',pady=6)
                new_upload_button.grid(row=1,column=2,sticky='w')   
                                            
            if status1==0 and data.empty==False:
                calculation(data)
        elif status1==1:
            if input_file_excel =="":
                upload_regular_gpa.grid(row=2,column=0,sticky='w')
            else:
                data=read_excel(input_file_excel)
                if "Htno" not in data.columns:
                    wrong_file.grid(row=2,column=0,sticky='w')
    else:
        upload_result_label.grid(row=2,column=0,sticky='w')

Label(root,text="   ",font=('Times new Roman',16),bg="#FFE9E3").grid(row=3,column=2)       
#Reset functionality
def userGuide():
    user_guide()
#Get result button
Button(root,text="Result Analysis",command=save,font=('Times new Roman',16)).grid(row=4,column=2,sticky='w')
#Reset button
Button(root,text="User Guide",command=userGuide,font=('Times new Roman',16)).grid(row=4,column=0)
master.mainloop()