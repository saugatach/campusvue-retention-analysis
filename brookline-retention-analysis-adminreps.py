#!/usr/bin/python

import sys,os
import pandas as pd
import numpy as np
import re
from IPython.display import display
from matplotlib import pyplot as plt
#display options
#pd.set_option('display.expand_frame_repr', False)

#import XLS file
xls=pd.ExcelFile("adminreps.xlsx")
df=pd.DataFrame()
for sheet in xls.sheet_names:
    print("Appending sheet ...", sheet)
    tempdf=pd.read_excel(xls,sheet)
    tempdf.fillna(0,inplace=True)
    df=df.append(tempdf,ignore_index=True)

# aggregate data for lead sources
df_lead=df[['SchoolStatus','LeadSrcDesc']]
selectrows=((df['SchoolStatus']=="Active") | (df['SchoolStatus']=="NDS - Active"))
leads=df_lead[selectrows].groupby(['LeadSrcDesc']).size().reset_index(name='active')
leads1=df_lead.groupby(['LeadSrcDesc']).size().reset_index(name='total')
leads=pd.merge(leads,leads1)
leads['retentionrate']=round(leads['active']/leads['total']*100,2)
leads=leads.sort_values(by='retentionrate',ascending=False)

# aggregate data for reps
reps_student_ditribution=df.groupby(['AdmRep','CampusDescrip']).size().reset_index(name='Freq')[1:]
reps_retention=df.groupby(['AdmRep','SchoolStatus']).size().reset_index(name='freq')[1:]
reps_retention_allstatus=df.groupby(['AdmRep']).size().reset_index(name='total')[1:]
reps=pd.merge(reps_retention,reps_retention_allstatus)
reps['retentionrate']=round(reps['freq']/reps['total']*100,2)
reps_pivot_table = pd.pivot_table(reps, index=['AdmRep'], columns=['SchoolStatus'],  values=['retentionrate'], aggfunc=np.sum, dropna=True).fillna(0)
#flatten pivottable columns
reps=reps_pivot_table.copy()
reps.columns=reps.columns.get_level_values(1)
reps.reset_index(inplace=True)

reps['Total']=reps.sum(axis=1)

# Generate required columns for Retention rate and Cancel rate

reps['Activerate']=reps[['Active','NDS - Active']].sum(axis=1)
# find all columns with the word "Cancel" in them using REGEX
regex=re.compile(r'ancel')
cancel_cols=list(filter(regex.search,reps.columns))
reps['Cancelrate']=reps[cancel_cols].sum(axis=1)

# extract relevant columns and sort by retention rate
reps_rates=reps.sort_values(by=['Activerate','AdmRep'], ascending=False)[['AdmRep','Activerate','Cancelrate']]
reps_rates.reset_index(inplace=True)
reps_rates.drop('index',axis=1,inplace=True)
reps_rates.set_index('AdmRep',inplace=True)




# print rep rates
reps_rates




plt.rcParams['figure.figsize'] = [80, 100]
plt.rcParams.update({'font.size': 48})

pdplot=reps_rates.plot.barh()
fig = pdplot.get_figure()
fig.savefig('admin-retention-rates-plot.png')




pdplot=reps_rates.plot.barh(stacked=True)
fig = pdplot.get_figure()
fig.savefig('admin-retention-rates-stacked-plot.png')




leads




pdplot=leads.plot.barh(x='LeadSrcDesc',y='retentionrate')
fig = pdplot.get_figure()
fig.savefig('leads-retention-rates-stacked-plot.png')

