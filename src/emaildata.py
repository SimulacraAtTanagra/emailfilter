# -*- coding: utf-8 -*-
"""
Created on Thu Feb 18 10:36:13 2021

@author: sayers
"""

"""
This is the script to take latest employee data and sort into categories
specifically for the purposes of feeding email bot
This process should be run between twice a day and once every two days.
Input is Current Job Report from CUNYfirst as excel file.
Output is Json file with dict, key is group name, value is list of names
"""

from admin import newest, colclean, rehead ,read_json,to_records,write_json,trydict,fileverify
import pandas as pd

def subsetter(df,crit_field,criteria):
    try:
        newdf=df[df[crit_field].str.contains(criteria)]
        df=df[~df['empl_id'].isin(list(newdf.empl_id.unique()))]
        return(df,newdf)
    except:
        newdf=df[df[crit_field]==criteria]
        df=df[~df['empl_id'].isin(list(newdf.empl_id.unique()))]
        return(df,newdf)
       



def refresh_lists():
    infolder='s://downloads'
    outfile='Y://Program Data//emaildata.json'
    #first read in df by opening most recent  CJR
    df=colclean(rehead(pd.read_excel(newest(infolder,'FULL_FILE')),2))
    df.person_nm=df.person_nm.replace("Audre Jackson","Annie Jackson")
    #pull in current department data for depthead fields
    deptfile='Y://Current Data//Lookup Tables//departments_file.xlsx'
    df2=colclean(pd.read_excel(deptfile))
    chairs=list(df2.chairperson.unique())
    support=list(df2[df2.support_staff.isnull()==False].support_staff.unique())
    support=[x for x in support if x!=0]
    newlist=[]
    for x in support:
        y=x.split(',')
        if type(y)==list:
            newlist.extend(y)
        else:
            newlist.append(y[0])
    [x.strip() for x in newlist]
    #subset df to include only required fields 
    #minimally, this should be paygroup, reports to, name, department, title
    df=df[['empl_id','hr_status','person_nm','labor_job_ld','dept_descr_job','union_job_cd','empl_cls_ld','reports_to_emplid']]
    
    #add a flag for department heads, department secretaries, sueprvisors, etc
    df.loc[(df['person_nm'].isin(chairs)) & (df['hr_status'] == 'Active'),'dept_head'] = 'chair'
    df.loc[(df['person_nm'].isin(newlist)) & (df['hr_status'] == 'Active'),'support'] = 'support'
    
    #identifying people who ae supervisors
    reports=list(df[df.reports_to_emplid.isnull()==False].reports_to_emplid.unique())
    df.loc[(df['empl_id'].isin(reports)) & (df['hr_status'] == 'Active'),'supervisor'] = 'supervisor'
    
    
    critloc=[('person_nm',"Sabrina Johnson Chandler",'From Sabrina Johnson Chandler'),
    ('reports_to_emplid','23134914','From HRIS Team'),
    ('dept_descr_job','Human Resources','From other HR'),
    ('dept_descr_job','Computer Svcs','From IT'),
    ('dept_descr_job','Budget','From Budget'),
    ('dept_descr_job','Office of Planning and Budget','From Budget'),
    ('union_job_cd','ECP','From Vice President'),
    ('dept_head','chair','From Department Head'),
    ('supervisor','supervisor','From Supervisor'),
    ('support','support','From Department Secretary'),
    ('labor_job_ld','College Assistant',"From College Assistants"),
    ('empl_cls_ld',"Tenured","From Faculty")]
    write_json(critloc,'y://program data//emaildata_criteria_list')
    #open current json document or prepare to create
    if fileverify(outfile):
        maindict=read_json(outfile)
    else:
        maindict={}
    writedict={}
    #have a standing list of criteria fields and criteria, then iterate 
    #have 3 part tuple with crit_field, criteria, and label
    critloc='Y://Program Data//emaildata_criteria_list.json'
    critlist=read_json(critloc)
    for crit_field,criteria,label in critlist:
        df,newdf=subsetter(df,crit_field,criteria)
        #take names from newdf and add to list
        newlist=list(newdf.person_nm.unique())
        #take list and add to maindict
        writedict[label]=newlist

    #compare to discover any updates that were made. This is a late stage
    #error check
    try:
        diff_dict= [v for k,v in writedict.items() if maindict[k]!=v]
    except:
        diff_dict=[]

    if len(diff_dict)!=0:
        print("These are differences from the last run of refresh_lists")
        print(diff_dict)
        write_json(writedict,outfile[:-5])
    else:
        write_json(writedict,outfile[:-5])
    
if __name__=="__main__":
    refresh_lists()