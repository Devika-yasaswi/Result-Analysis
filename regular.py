import pandas as pd
from tkinter.filedialog import *
from Statistics import get_statistics
import sys
import os
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
def Sgpa(data,input):
    #Initializations
    global GPA,a,roll_no,student_data,start,start_x,df,civil_credits,eee_credits,mech_credits,ece_credits,cse_credits,GBM,tc
    roll_no=0    #Variable for collectig last three digis of the RollNo
    a=0    #
    GPA=0.0
    df=pd.DataFrame({"Roll_No":[]})
    student_data=[data['Htno'][1]]
    sub=[]
    start=int(data['Htno'][1][0:4])  
    start_x=1
    cse=0
    total=0
    civil_credits=0
    eee_credits=0
    mech_credits=0
    ece_credits=0
    cse_credits=0
    GBM=0
    tc=0
    df1=pd.DataFrame({"Files updated ":[input]})

    #Deleting data frame for the creation of new branch dataframe with same name
    def delete():
            global df
            d=df
            del df
            del d
            df=pd.DataFrame({"Roll_No":[]})
            #print(len(df))
        
    #Calculation and entering marks of the student into data frame
    files=[('xlsx files','*.xlsx')]
    file=asksaveasfile(mode='wb',filetypes = files,defaultextension=files)

    #Calculating credits
    for i in range(len(data)):
        x=int(data.iloc[i,0][7:10])
        if data.iloc[i,1] not in sub:
            sub.append(data.iloc[i,1])
            total+=float(data.iloc[i,-1])
            student_data.append(data.iloc[i,-2])
        if data.iloc[i,0] not in student_data:
            if x//100==1:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data:
                     civil_credits=total
            elif x//100==2:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data:
                     eee_credits=total
            elif x//100==3:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data:
                     mech_credits=total
            elif x//100==4:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data:
                     ece_credits=total
            elif x//100==5:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data:
                     cse_credits=total
            #print(student_data)
            #print(x," ",civil_credits," ",eee_credits," ",mech_credits," ",ece_credits," ",cse_credits)
            sub=[]
            total=0
            student_data=[data.iloc[i,0]]        
    #print(civil_credits," ",eee_credits," ",mech_credits," ",ece_credits," ",cse_credits)
    student_data=[]
    #calculating and writing GPA to output file
    with pd.ExcelWriter(resource_path(file.name),engine='openpyxl',mode='w') as output:    
        for i in range(len(data)):        
            #print(i,data.iloc[i,0])
            d=str(data.iloc[i,0])
            x=int(d[7:10])
                    
            #Entering the list of students values stored in the dataframe into the marks excel sheet
            if data.iloc[i,0] not in student_data:
                def enter():  
                    global a,roll_no,GPA,GBM,tc
                    if 'SGPA' not in df.columns:
                        df['GBM']=[]
                        df['Total Credits']=[] 
                        df['Status']=[]
                        df['Backlogs']=[]
                        df['TC']=[]
                        df['Pass Percentage']=[]
                        df['Points']=[]
                        df['SGPA']=[]
                        
                    student_data.append(GBM)    
                    
                    student_data.append(total_credits)
                    if "F" not in student_data and "AB" not in student_data and "MP" not in student_data:
                        student_data.append("Pass")
                    else:
                        student_data.append("Fail")
                    student_data.append(student_data.count("F")+student_data.count("AB")+student_data.count("MP")+student_data.count("ABSENT"))
                    student_data.append(tc)
                    student_data.append(GBM/(len(sub)-(student_data.count("COMPLE")+student_data.count("COMPLETED"))))
                    student_data.append(GPA)                       
                    GPA=GPA/total_credits                 
                    student_data.append(GPA)                    
                    #print(student_data)
                    #print(a,'=',GPA)
                    try:
                        df.loc[len(df.index)]=student_data 
                    except:
                        pass
                    student_data.clear()
                    a=a+1
                    GPA=0
                    roll_no+=1
                    GBM=0
                    tc=0
                if i>0:
                    enter()
                student_data.append(data.iloc[i,0])
            
            #Entering the excel sheets based on the branch (sheet1=Civil,sheet2=Mechanical,sheet3=EEE,sheet4=ECE,sheet5=CSE)
            if int(d[7])>(a/100) or int(d[0:4])>start:
                if int(d[0:4])>start:
                    start=int(d[0:4])
                    cse=len(df.index)+1
                    df.to_excel(output,sheet_name="CSE",index=False)
                    start_x=2
                    delete()
                a=int(d[7])*100+1
                #print("a=",a)
                if int(d[7])==1:
                    total_credits=civil_credits                
                if int(d[7])==2:                                       
                    if start_x==2:
                        df.to_excel(output,sheet_name="CE",index=False,startrow=civil,header=None)
                    else:
                        civil=len(df.index)+1
                        df.to_excel(output,sheet_name="CE",index=False)
                        #df.to_excel(output,sheet_name="CE stats",index=False)
                    total_credits=eee_credits
                    delete()
                    df["Roll_No"]=[]
                if int(d[7])==3:
                    if start_x==2:
                        df.to_excel(output,sheet_name="EEE",index=False,startrow=eee,header=None)
                    else:
                        eee=len(df.index)+1
                        df.to_excel(output,sheet_name="EEE",index=False)
                        #df.to_excel(output,sheet_name="EEE stats",index=False)
                    total_credits=mech_credits
                    delete()
                    df["Roll_No"]=[]
                if int(d[7])==4:
                    if start_x==2:
                        df.to_excel(output,sheet_name="ME",index=False,startrow=mech,header=None)
                    else:
                        mech=len(df.index)+1
                        df.to_excel(output,sheet_name='ME',index=False)
                        #df.to_excel(output,sheet_name='ME stats',index=False)
                    total_credits=ece_credits
                    delete()
                    df["Roll_No"]=[]
                if int(d[7])==5:
                    if start_x==2:
                        df.to_excel(output,sheet_name="ECE",index=False,startrow=ece,header=None)
                    else:
                        ece=len(df.index)+1
                        df.to_excel(output,sheet_name="ECE",index=False)
                        #df.to_excel(output,sheet_name="ECE stats",index=False)
                    total_credits=cse_credits
                    delete()
                    df["Roll_No"]=[]
                sub=[]
        
        #Adding subject name columns in the dataframe based on the subject code
            if data.iloc[i,1] not in sub:            
                try:                    
                    df[data['Subname'][i]+' ('+data.iloc[i,1]+')']=[]
                    sub.append(data.iloc[i,1])
                except:
                    continue

            
            #Grades acquired based on the marks of the students 
            #90-100  A+ grade
            #80-89  A grade
            #70-79  B grade
            #60-69  C grade
            #50-59  D grade
            #40-59  E grade
            #<40 Fail
            #AB Absent
            tc+=data.iloc[i,-1]
            if data.iloc[i,-2]=='A+':
                    grade=10
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='A':
                    grade=9
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='B':
                    grade=8
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='C':
                    grade=7
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='D':
                    grade=6
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='E':
                    grade=5
                    GBM+=grade*10
                    GPA+=grade*data.iloc[i,-1]
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='F':             
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='AB' or data.iloc[i,-2]=='ABSENT' or data.iloc[i,-2]=="MP":
                    student_data.append(data.iloc[i,-2])
            elif data.iloc[i,-2]=='COMPLETED' or data.iloc[i,-2]=='COMPLE':
                    student_data.append(data.iloc[i,-2])
                
        #Adding final sheet CSE to the Excel
        enter()
        if start_x==2:
            df.to_excel(output,sheet_name="CSE",index=False,startrow=cse,header=None)
        else:
            df.to_excel(output,sheet_name="CSE",index=False,startrow=cse)
        #df.to_excel(output,sheet_name="CSE stats",index=False,startrow=cse,header=None)
        #df1.to_excel(output,sheet_name="Updated files",index=False)
    get_statistics(file.name)  