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
    global GPA,a,roll_no,student_data,start,start_x,df,civil_credits,eee_credits,mech_credits,ece_credits,cse_credits,GBM,tc,civil_subs,eee_subs,mech_subs,ece_subs,cse_subs,total_subs
    roll_no=0    #Variable for collectig last three digis of the RollNo
    a=0    #
    GPA=0.0
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
    civil_list=[]
    eee_list=[]
    mech_list=[]
    ece_list=[]
    cse_list=[]
    GBM=0
    tc=0

    #Deleting data frame for the creation of new branch dataframe with same name
    def delete(cols):
        global df
        d=df
        del df
        del d
        cols.insert(0,"Roll No")
        df=pd.DataFrame(columns=cols)
        cols.pop(0)
        return df
        
    #Calculation and entering marks of the student into data frame
    files=[('xlsx files','*.xlsx')]
    file=asksaveasfile(mode='wb',filetypes = files,defaultextension=files)

    #Finding subjects
    for i in range(len(data)):
        x=int(data.iloc[i,0][7:10])
        if data.iloc[i,0] not in student_data:
            if x//100==1 and int(data.iloc[i-1,0][7])!=5:
                 civil_list.append(sub)
            elif x//100==2 and int(data.iloc[i-1,0][7])!=1:
                 eee_list.append(sub)
            elif x//100==3 and int(data.iloc[i-1,0][7])!=2:
                 mech_list.append(sub)
            elif x//100==4 and int(data.iloc[i-1,0][7])!=3:
                 ece_list.append(sub)
            elif x//100==5 and int(data.iloc[i-1,0][7])!=4:
                 cse_list.append(sub)
            sub=[]
            student_data=[data.iloc[i,0]]
        if data.iloc[i,1] not in sub:
            sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
            total+=float(data.iloc[i,-1])
            student_data.append(data.iloc[i,-2]) 
    def sub_selection(subs):
        new_sub=[]
        count=[]
        final_sub=[]
        for i in subs:
             if i not in new_sub:
                  new_sub.append(i)
        #print(new_sub)
        for i in range(len(new_sub)):
            count.append(len(new_sub[i]))
        for i in range(len(count)):
            if count[i] == min(count):
                final_sub.append(new_sub[i])
        new_final_sub=final_sub[0]
        for i in range(1,len(final_sub)):
            for j in range(len(final_sub[i])):
                 if final_sub[i][j] not in new_final_sub:
                    new_final_sub.append(final_sub[i][j])
        return new_final_sub
    civil_subs=sub_selection(civil_list)
    eee_subs=sub_selection(eee_list)
    mech_subs=sub_selection(mech_list)
    ece_subs=sub_selection(ece_list)
    cse_subs=sub_selection(cse_list)
    #print(civil_subs,eee_subs,mech_subs,ece_subs,cse_subs)
    #Calculating credits
    for i in range(len(data)):
        x=int(data.iloc[i,0][7:10])
        if data.iloc[i,0] not in student_data:
            #print(set(sub)-set(civil_subs))
            #print(set(sub)-set(eee_subs))
            #print(set(sub)-set(mech_subs))
            if x//100==1:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data and set(sub)-set(civil_subs)==set():
                    civil_credits=total
                     
            elif x//100==2:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data and set(sub)-set(eee_subs)==set():
                     eee_credits=total
                
            elif x//100==3:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data and set(sub)-set(mech_subs)==set():
                    mech_credits=total
            elif x//100==4:                
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data and set(sub)-set(ece_subs)==set():
                     ece_credits=total
            elif x//100==5:
                if "MP" not in student_data and 'F' not in student_data and "AB" not in student_data and set(sub)-set(cse_subs)==set():
                     cse_credits=total
            #print(student_data)
            #print(x," ",civil_credits," ",eee_credits," ",mech_credits," ",ece_credits," ",cse_credits)
            sub=[]
            total=0
            student_data=[data.iloc[i,0]]  
        if data.iloc[i,1] not in sub:
            sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
            total+=float(data.iloc[i,-1])
            student_data.append(data.iloc[i,-2])      
    #print(civil_credits," ",eee_credits," ",mech_credits," ",ece_credits," ",cse_credits)
    student_data=[]
    civil_subs.insert(0,"Roll No")
    df=pd.DataFrame(columns=civil_subs)
    civil_subs.pop(0)
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
                    student_data.append(student_data.count("F")+student_data.count("AB")+student_data.count("MP")+student_data.count("ABSENT")+student_data("WH"))
                    student_data.append(tc)
                    student_data.append(GBM/(len(total_subs)-(student_data.count("COMPLE")+student_data.count("COMPLETED"))))
                    student_data.append(GPA)                       
                    GPA=GPA/total_credits                 
                    student_data.append(GPA)                    
                    #print(student_data)
                    #print(a,'=',GPA)
                    try:
                        df.loc[len(df.index)]=student_data 
                    except ValueError:
                        #print(total_subs)
                        #print(sub)
                        for b in range(len(total_subs)):
                             if total_subs[b] not in sub:
                                #print(total_subs)
                                student_data.insert(b+1,"-")
                        #print(student_data)
                        df.loc[len(df.index)]=student_data                                 
                    student_data.clear()
                    sub.clear()
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
                    df=delete(civil_subs)
                a=int(d[7])*100+1
                #print("a=",a)
                if int(d[7])==1:
                    total_credits=civil_credits   
                    total_subs=civil_subs             
                if int(d[7])==2:                                       
                    if start_x==2:
                        df.to_excel(output,sheet_name="CE",index=False,startrow=civil,header=None)
                    else:
                        civil=len(df.index)+1
                        df.to_excel(output,sheet_name="CE",index=False)
                        #df.to_excel(output,sheet_name="CE stats",index=False)
                    total_credits=eee_credits
                    total_subs=eee_subs
                    df=delete(eee_subs)
                    #print(df)
                if int(d[7])==3:
                    if start_x==2:
                        df.to_excel(output,sheet_name="EEE",index=False,startrow=eee,header=None)
                    else:
                        eee=len(df.index)+1
                        df.to_excel(output,sheet_name="EEE",index=False)
                        #df.to_excel(output,sheet_name="EEE stats",index=False)
                    total_credits=mech_credits
                    total_subs=mech_subs
                    df=delete(mech_subs)
                if int(d[7])==4:
                    if start_x==2:
                        df.to_excel(output,sheet_name="ME",index=False,startrow=mech,header=None)
                    else:
                        mech=len(df.index)+1
                        df.to_excel(output,sheet_name='ME',index=False)
                        #df.to_excel(output,sheet_name='ME stats',index=False)
                    total_credits=ece_credits
                    total_subs=ece_subs
                    df=delete(ece_subs)
                if int(d[7])==5:
                    if start_x==2:
                        df.to_excel(output,sheet_name="ECE",index=False,startrow=ece,header=None)
                    else:
                        ece=len(df.index)+1
                        df.to_excel(output,sheet_name="ECE",index=False)
                        #df.to_excel(output,sheet_name="ECE stats",index=False)
                    total_credits=cse_credits
                    total_subs=cse_subs
                    df=delete(cse_subs)

            
            #Grades acquired based on the marks of the students 
            #90-100  A+ grade
            #80-89  A grade
            #70-79  B grade
            #60-69  C grade
            #50-59  D grade
            #40-59  E grade
            #<40 Fail
            #AB Absent
            if data.iloc[i,2]+" "+data.iloc[i,1] in total_subs:
                tc+=data.iloc[i,-1]
                if data.iloc[i,-2]=='A+':
                        grade=10
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='A':
                        grade=9
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='B':
                        grade=8
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='C':
                        grade=7
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='D':
                        grade=6
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='E':
                        grade=5
                        GBM+=grade*10
                        GPA+=grade*data.iloc[i,-1]
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='F':             
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='AB' or data.iloc[i,-2]=='ABSENT' or data.iloc[i,-2]=="MP":
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='COMPLETED' or data.iloc[i,-2]=='COMPLE':
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                elif data.iloc[i,-2]=='WH':
                        student_data.append(data.iloc[i,-2])
                        sub.append(data.iloc[i,2]+" "+data.iloc[i,1])
                
        #Adding final sheet CSE to the Excel
        enter()
        if start_x==2:
            df.to_excel(output,sheet_name="CSE",index=False,startrow=cse,header=None)
        else:
            df.to_excel(output,sheet_name="CSE",index=False,startrow=cse)
        #df.to_excel(output,sheet_name="CSE stats",index=False,startrow=cse,header=None)
        #df1.to_excel(output,sheet_name="Updated files",index=False)
    get_statistics(file.name)  