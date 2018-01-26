# coding: utf-8
#!/usr/bin/python
# -*- coding: iso-8859-15 -*-


# Collect short positions and display daily and three months rolling. 
#Libraries in use
import urllib
import pandas as pd
import sys
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import time
import datetime
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta


# Time the execution of the script
#starttime = datetime.now()


# Daily adjustment of web address 
yesterday = date.today() - timedelta(1)
now_var2 = str(yesterday)
###now_var3 = list(now_var2)
now_var3 = list("2017-12-08")

three_months_ago = datetime.now() - relativedelta(months=+3) # for fliter of df below
three_months_ago_var2 = str(three_months_ago)
three_months_ago_var3 = three_months_ago_var2[0:10]
#print three_months_ago_var3
#change the data in the web adress and create a new adress
address = list("http://www.fi.se/contentassets/71a61417bb4c49c0a4a3a2582ea8af6c/korta_positioner_2017-03-15.xlsx")
address[81:91] = now_var3
address_var1 = "".join(address)

#Scrapes the web and downloads excel file
blanka = urllib.URLopener()
try:
    blanka.retrieve(address_var1, "short_positions.xlsx")
except:
    print "File not available today"
    exit()
    

# dataframe for analysis
###time.sleep(10)
df = pd.read_excel("short_positions.xlsx", sheetname=0, header=0, skiprows=7, skip_footer=0, index_col=None, parse_cols = 7,  encoding = 'utf8', names=["Date", "Holder", "Issuer/stock", "ISIN", "Percent", "Position date", "Comment", ""])
#df.apply(lambda x: x.astype(str).str.upper()) ## tried to decode and encode with unicode. need more work
#df.apply(lambda x: pd.lib.infer_dtype(x.values))

df_2 = df.iloc[:, [2,4,5]]

df_2["Position date"] = pd.to_datetime(df_2["Position date"], format='%m%d%y')
print df_2.head()


#datetime.datetime.strptime("2013-1-25", '%Y-%m-%d').strftime('%m/%d/%y')



#df_2['Issuer/stock'] = map(lambda x: unicode(x).upper(), df_2['Issuer/stock'])
df_2 = df_2.set_index("Position date")
#df_var1 = df_var1.iloc[:, [0,1,3]]
#print df_var1.describe
print df_2

#Sort the dataframe to fit analysis
bystock = df_2.groupby(['Issuer/stock', 'Percent'])
#print bystock.head()#['Percent'].describe

#Bring in the established dates above and adjust the dataframe
last_day = df_2.ix[now_var2 : now_var2]
  
#last_day = df_2.ix["2017-12-06" : "2017-12-06"]
print last_day 
three_months = df_2.ix[three_months_ago_var3 : now_var2]
#print last_day
shorts_last_day = last_day["Issuer/stock"].value_counts()
shorts_last_day_var2 = shorts_last_day
#print shorts_last_day_var2
shorts_three_months = three_months["Issuer/stock"].value_counts()
shorts_three_months_var2 = shorts_three_months.head(20)
#print shorts_three_months_var2


# Visualize Daily
import seaborn as sns
sns.set_palette("Set1", 2, 0.75)
sns.set_style("darkgrid")
plt.figure();
fig = shorts_last_day_var2.plot.bar(stacked=False); plt.ylabel('No of Shorts'); plt.xlabel('Name'); plt.title(now_var2, fontsize=11); plt.ylim(0,10) 
fig = plt.gcf()
fig.savefig('shorts.png', bbox_inches='tight')

# Visualize 3 months rolling
sns.set_palette("Set1", 2, 0.75)
sns.set_style("darkgrid")
plt.figure();
fig = shorts_three_months_var2.plot.bar(stacked=False); plt.ylabel('No of Shorts');plt.xlabel('Name'); plt.title("3 months (rolling)", fontsize=11)
fig = plt.gcf()
fig.savefig('shorts_3_months.png', bbox_inches='tight')

# Print execution time of the script
#print datetime.now() - starttime

