#!/usr/bin/python

import sys,os
import pandas as pd
import numpy as np
import re
from matplotlib import pyplot as plt


def mergeXLsheets(xls):
#import XLS file
    df=pd.DataFrame()
    for sheet in xls.sheet_names:
        print("\n Appending sheet ...", sheet)
        tempdf=pd.read_excel(xls,sheet)
        tempdf.fillna(0,inplace=True)
        df=df.append(tempdf,ignore_index=True)
        
    return df

def calc_drop_rates(df,groupbycol,aggcols):
    # aggregates(sums over) the list of columns in "aggcols"
    # after filtering the column "groupbycol"
    
    # create a dict of columns to aggregate
    # creates a dict like {'FinalRegDropStudents':sum,'FinalRegStudents':sum}
    aggdict={x:sum for x in aggcols}
    
    # sum "aggcols" by grouping over "groupbycol"
    df_drop_rate=df.groupby(groupbycol).agg(aggdict).reset_index()

    df_drop_rate['DropRate']=round(df_drop_rate['FinalRegDropStudents']/df_drop_rate['FinalRegStudents']*100,2)

    df_drop_rate.sort_values(by=['DropRate'],ascending=False,inplace=True)
    print("\n Saving table to file : " + groupbycol +'_drop_rate.csv')
    df_drop_rate.to_csv(groupbycol+'_drop_rate.csv',index=False)
    
    # dynamically update the figure and the font size to accomodate more rows of data     
    data_count=df_drop_rate.shape[0]
    plt.rcParams['figure.figsize'] = [30, 5*round(np.sqrt(data_count))]
    plt.rcParams.update({'font.size': 8*round(np.log(data_count))})

    # Do not plot zeros
    df_drop_rate=df_drop_rate[(df_drop_rate['DropRate']!=0)].reset_index()

    # plot horizontal bar chart to allow the names to accommodate long labels
    print("\n Saving image to file : " + groupbycol + '_drop_rate.png')
    pdplot=df_drop_rate.plot.barh(x=groupbycol, y='DropRate')

    fig = pdplot.get_figure()

    fig.savefig(groupbycol+'_drop_rate.png')
    return df_drop_rate

# start of __main__
xlfile=(input("Enter XL filename [default: 'master.xlsx']: \n>") or "master.xlsx")
xls=pd.ExcelFile(xlfile)
df=mergeXLsheets(xls)
print("\n All the sheets from the XL file"+xlfile+" has been merged into a single file: merged.csv \n")
df.to_csv('merged.csv', index=False)

# Select relevant columns and drop others
tempdf=df[['Instructor', 'Campus', 'CourseCode', 'CourseSection', 'FinalRegDropStudents', 'FinalRegStudents']]
# remove all whitespaces
tempdf = tempdf.replace(to_replace=r"\s", value="", regex=True)

# the data is dirty. The Campus code should be "Online" for CourseSection=[ON-1 , ON-2, ON-3]
# select all rows with CourseSection=[ON-1 , ON-2 ] and then replace the Campus code to read "Online"
tempdf['Campus'].replace(to_replace=tempdf[((tempdf['CourseSection'] == 'OL-1') |
                                            (tempdf['CourseSection'] == 'OL-2') |
                                            (tempdf['CourseSection'] == 'OL-3'))]
                         ['Campus'], value='Online', inplace=True)

# select unique rows only
df_retention = tempdf.groupby(['Instructor', 'Campus', 'CourseCode', 'CourseSection', 'FinalRegDropStudents',
                               'FinalRegStudents']).size().reset_index(name='freq')
# Don't need the frequency column
df_retention.drop(['freq'], axis=1, inplace=True)

# drop rate aggregate by sections
sections_drop_rate = df_retention.copy()
sections_drop_rate['DropRate'] = round(sections_drop_rate['FinalRegDropStudents']/sections_drop_rate['FinalRegStudents']*100, 2)
# this is the cleaned table that contains all relevant enrolled uders' data. save it.
sections_drop_rate.to_csv('sections_drop_rate.csv', index=False)

# now we can extract retention rates for each parameter among ('Instructor', 'Campus', 'CourseCode')
colnames = str(list(sections_drop_rate.columns))
groupbycol = (input("Groupby which columns (select one) [default:'Instructor']: \n" + colnames + "\n >") or 'Instructor')

aggcols=['FinalRegDropStudents','FinalRegStudents']
aggdict={x:sum for x in aggcols}

# aggregation, exporting the aggregated table, and plotting all will be done by the function calc_drop_rates()
df_drop_rate=calc_drop_rates(sections_drop_rate, groupbycol, aggcols)

