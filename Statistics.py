from pandas import *
from tkinter.filedialog import *
def branch_calculation(data,overall_stat,branch):
    df=DataFrame(columns=["subject","Registered","Appeared","Absent","Failed","Passed","Pass Percentage"])
    for i in list(data.columns)[1:-8]:
        temp_list=[i]
        grades=data[i].tolist()
        temp_list.append(len(grades)-grades.count("-"))
        total=len(grades)-(grades.count("AB")+grades.count("ABSENT")+grades.count("-"))
        temp_list.append(total)
        temp_list.append(grades.count("AB")+grades.count("ABSENT"))
        passed=grades.count("A+")+grades.count("A")+grades.count("B")+grades.count("C")+grades.count("D")+grades.count("E")+grades.count("COMPLE")+grades.count("COMPLETED")
        temp_list.append(len(grades)-passed-grades.count("-")-grades.count("WH"))
        temp_list.append(passed)
        temp_list.append(passed/total*100)
        df.loc[len(df.index)]=temp_list
    count=0
    count1=0
    for i in range(len(data)):
        new=data.iloc[i,1:-7]    
        new=list(set(new))
        if len(new)==1 and (new[0]=="AB" or new[0]=="ABSENT"):
            count+=1
        if len(new)!=1 and ("F" in new or "MP" in new or  "AB" in new or "ABSENT" in new or "MP" in new or "WH" in new):
            count1+=1
    new=[branch,len(data),len(data)-count,len(data)-count1,count1,(len(data)-count1)/(len(data)-count)*100]
    overall_stat.loc[len(overall_stat.index)]=new
    return df,overall_stat
    
def failure_count(data):
        count=len(data)
        df=DataFrame(columns=["No.of Failed Subjects","No.of Failures"])
        backlog_list=data["Backlogs"].tolist()
        new=["All Pass"]
        new.append(backlog_list.count(0))
        df.loc[len(df.index)]=new
        for i in range(1,max(backlog_list)+1):
            new=[str(i)+" Subject Failed"]
            new.append(backlog_list.count(i))
            df.loc[len(df.index)]=new
        return df,count

def get_statistics(file):
    civil_data=read_excel(file,sheet_name=["CE"])
    civil_data=civil_data["CE"]
    eee_data=read_excel(file,sheet_name=["EEE"])
    eee_data=eee_data["EEE"]
    mech_data=read_excel(file,sheet_name=["ME"])
    mech_data=mech_data["ME"]
    ece_data=read_excel(file,sheet_name=["ECE"])
    ece_data=ece_data["ECE"]
    cse_data=read_excel(file,sheet_name=["CSE"])
    cse_data=cse_data["CSE"]
    overall_data=DataFrame(columns=["Branch","Total","Appeared","Pass","Fail","Percentage"])
    civil_df,civil=failure_count(civil_data)
    eee_df,eee=failure_count(eee_data)
    mech_df,mech=failure_count(mech_data)
    ece_df,ece=failure_count(ece_data)
    cse_df,cse=failure_count(cse_data)
    civil_data,overall_data=branch_calculation(civil_data,overall_data,"CE")
    eee_data,overall_data=branch_calculation(eee_data,overall_data,"EEE")
    mech_data,overall_data=branch_calculation(mech_data,overall_data,"ME")
    ece_data,overall_data=branch_calculation(ece_data,overall_data,"ECE")
    cse_data,overall_data=branch_calculation(cse_data,overall_data,"CSE")
    new=["Total"]
    new.append(sum(list(overall_data["Total"])))
    new.append(sum(list(overall_data["Appeared"])))
    new.append(sum(list(overall_data["Pass"])))
    new.append(sum(list(overall_data["Fail"])))
    new.append(new[3]/new[2]*100)
    overall_data.loc[len(overall_data.index)]=new
    with ExcelWriter(file,engine='openpyxl',mode='a',if_sheet_exists="replace") as output:
        civil_data.to_excel(output,sheet_name="CE Analysis",index=False)
        eee_data.to_excel(output,sheet_name="EEE Analysis",index=False)
        mech_data.to_excel(output,sheet_name="ME Analysis",index=False)
        ece_data.to_excel(output,sheet_name="ECE Analysis",index=False)
        cse_data.to_excel(output,sheet_name="CSE Analysis",index=False)
        overall_data.to_excel(output,sheet_name="Overall Analysis",index=False)

    with ExcelWriter(file,engine='openpyxl',mode='a',if_sheet_exists="overlay") as output:
        civil_df.to_excel(output,sheet_name="CE",startrow=civil+3,startcol=4,index=False,header=False)
        eee_df.to_excel(output,sheet_name="EEE",startrow=eee+3,startcol=4,index=False,header=False)
        mech_df.to_excel(output,sheet_name="ME",startrow=mech+3,startcol=4,index=False,header=False)
        ece_df.to_excel(output,sheet_name="ECE",startrow=ece+3,startcol=4,index=False,header=False)
        cse_df.to_excel(output,sheet_name="CSE",startrow=cse+3,startcol=4,index=False,header=False)