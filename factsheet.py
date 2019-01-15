# -*- coding: utf-8 -*-
"""
Created on Wed Nov 28 14:14:08 2018

@author: Giulio Rossetti
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mtick
import datetime as dt


nround = 2
writer = pd.ExcelWriter('pandas_multiple.xlsx', engine='xlsxwriter')
workbook=writer.book
worksheet=workbook.add_worksheet('factsheet_output')
writer.sheets['factsheet_output'] = worksheet

date_to_nav_and_share = 20181130

file = r"Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\performance_pir_data.csv"
file2=r"Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\allocation.csv"
file3 = r"Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\index.csv"

ext = '.jpg'
dpi = 100

cum_ret = False
asset_all_barh = False
industry_plot = False
top_holding = False
pie_all = True

rf = 0.00959

raw_data = pd.read_csv(file,sep = ',',decimal = ',',encoding='utf-8')
ret = pd.Series(raw_data['DAILY'],dtype='float')
nav_share = pd.Series(raw_data['NAV/share'],dtype='float')

raw_allocation = pd.read_csv(file2,sep = ',',decimal = ',',encoding='utf-8')
raw_allocation['MKT VALUE'] = pd.to_numeric(raw_allocation['MKT VALUE'])
index_ret =pd.read_csv(file3,sep = ',',decimal = ',',encoding='utf-8')
index_ret.set_index('Date',inplace=True)
index_ret = index_ret.replace("n.d.",np.nan).dropna()

pir_ret_beta = pd.DataFrame(np.array(raw_data['DAILY']),index=(raw_data['Date']))
date_index_beta = index_ret.index.tolist()
date_index_beta  = pir_ret_beta.index.intersection(date_index_beta)

index_ret['PIR'] = pir_ret_beta.loc[date_index_beta]
index_ret = index_ret.apply(pd.to_numeric)

ex_nav = raw_data.copy()
ex_nav.set_index('Date',inplace = True)

nav = np.round(np.float(ex_nav.loc[date_to_nav_and_share,'NAV totale'])/1000000,nround)
share = ex_nav.loc[date_to_nav_and_share,'NAV/share']
 #=============================================================
#Statistics
#=============================================================
share_inception = 500
start_date = 20181031
end_date = 20181130
stat= pd.DataFrame([ret,nav_share]).T
stat.set_index(raw_data['Date'],inplace=True)
#stat = stat.loc[start_date:end_date,:]

df_monthly_nav =[] 

for i in range(len(stat.index)):

    if i ==0:
        df_monthly_nav.append([stat.index[i],stat.iloc[i,1]])
    else:
                 
        if not str(stat.index[i])[4:6] == str(stat.index[i-1])[4:6]:
            df_monthly_nav.append([stat.index[i],stat.iloc[i-1,1]])
        

df_monthly_nav.append            

df_monthly_nav= pd.DataFrame(np.stack(df_monthly_nav),columns=['Date','Share'])
df_monthly_nav['monthly ret'] =  df_monthly_nav['Share']/df_monthly_nav['Share'].shift()-1

avg_monthly_ret = df_monthly_nav['monthly ret'].mean()
avg_monthly_std = df_monthly_nav['monthly ret'].std()

ret_inception = df_monthly_nav.iloc[-1,1]/share_inception -1

sr = (avg_monthly_ret-rf)/avg_monthly_std


avg_annualized_ret = (1+avg_monthly_ret)**12-1     
avg_annualized_std =  avg_monthly_std * 12**0.5

s_bm = str(int(df_monthly_nav.iloc[np.argmax(df_monthly_nav.iloc[:,2])-1,0]))
s_wm = str(int(df_monthly_nav.iloc[np.argmin(df_monthly_nav.iloc[:,2])-1,0]))

best_month = "{}/{}/{}".format(s_bm[6:],s_bm[4:6],s_bm[0:4])
best_month = dt.datetime.strptime(best_month,'%d/%m/%Y')

worst_month = "{}/{}/{}".format(s_wm[6:],s_wm[4:6],s_wm[0:4])
worst_month = dt.datetime.strptime(worst_month,'%d/%m/%Y')

m_last =  str(int(df_monthly_nav.iloc[-1,0]))

downside_deviation = (((df_monthly_nav[df_monthly_nav['monthly ret'] <0])['monthly ret'].sum()**2)/ \
    len(df_monthly_nav[df_monthly_nav['monthly ret'] <0]))**0.5
                      
sortino_ratio = avg_monthly_ret / downside_deviation

ytd = [str(x)[:4] for x in df_monthly_nav['Date'] ]
ytd_m = 0
for i in range(len(ytd)-1,0,-1):
    if not ytd[i]==ytd[i-1]:
        ytd_m = len(ytd)-i-1
        break

array_rolling_ret = [1,3,6,ytd_m,12,36,48,len(df_monthly_nav)]  #array woth rolling month for cumulative ret calc [1m 3m 6m ytd 1y Incp]
cum_ret_head = ['1M','3M','6M','YTD','1Y','3Y','5Y','Incpt']

ret_cumulative = []
for i in array_rolling_ret:
    if i>len(df_monthly_nav):
        cum_ret_roll = '-'
        ret_cumulative.append(cum_ret_roll)
    else:
        cum_ret_roll = np.round((((df_monthly_nav.iloc[-i:,2].add(1)).cumprod()).iloc[-1]-1)*100,nround)
        ret_cumulative.append(cum_ret_roll)

table_ret_cumulative = pd.DataFrame(np.vstack((cum_ret_head,ret_cumulative)),\
                               index=['','Fund'],columns= ['Cumulative Returns']\
                               +['' for x in range(len(cum_ret_head)-1)])

array_rolling_std = [63,126,252,756,len(raw_data)]
risk_analysis_head = ['3M','6M','1Y','3Y','Incpt']
risk_analysis_ind = ['','STD','Beta']

std_rolling = []
for i in array_rolling_std:
    if i > len(raw_data):
        std_rolling.append('-')
    else:
        std_calc = np.round((ret[len(raw_data)-i:].std() *((i)**0.5))*100,nround)
        std_rolling.append(std_calc)

beta_rolling = []
for i in array_rolling_std:
    if i > len(index_ret):
        beta_rolling.append('-')
    else:
        ret_ind = np.array(index_ret.iloc[len(index_ret)-i:,0])
        ret_pir = np.array(index_ret.iloc[len(index_ret)-i:,1])
        X = np.vstack((np.ones_like(ret_pir),ret_ind)).T
        beta = np.linalg.inv(X.T @ X )@(X.T@ret_pir)
        beta = np.round(beta[1],nround)
        beta_rolling.append(beta)
        
table_risk_analysis = pd.DataFrame(np.vstack((risk_analysis_head,std_rolling,beta_rolling)), \
                                   index=[x for x in risk_analysis_ind],\
                                   columns = ['Risk Analysis(Rolling)']+['' for x in range(len(risk_analysis_head)-1)])
        
        
month_hist= ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sept','Oct', \
        'Nov','Dec']
year_hist = [x for x in np.unique(np.stack([int(x/10000) for x in raw_data['Date']]))]

hist_perf= np.empty(len(year_hist)*12)
hist_perf[:] =np.nan
inception_index = 5


for i in df_monthly_nav['monthly ret']:

    hist_perf[inception_index] = np.round(i*100,nround)
#    print(i,inception_index,a[inception_index])
    inception_index += 1

hist_perf = np.reshape(hist_perf,(len(year_hist),12)) 

year_hist1 = ['Year']+ year_hist
table_hist_perf = pd.DataFrame(np.vstack((month_hist,hist_perf)),index = year_hist1,\
                               columns = ['Historical Performance']+['' for x in range(11)])   
    
columns = ['Return Annualized','Return Monthly','STD Annualized','STD Monthly', \
           'Sharpe Ratio','Sortino Ratio','Best Month','Worst Month','Return(Inception)']

#=============================================================
#DICT to OUTPUT
#=============================================================

dict_table_fund_info = {'Fund Information':"", \
                            'Structure':'Open-End UCITS', \
                            'Subscription':'Daily',\
                            'Min. Sub.':'500€',\
                            'Subsequent subs': '50€ or multiples',\
                            'Domicile Country': 'Italy',\
                            'Redemption': 'Daily',\
                            'Subsc. Fee': '0-3%',\
                            'Management Fee': '1.5%',\
                            'Performance Fee': '5% hwm Ass.',\
                            'Custodian Bank': 'State Street',\
                            'Max Leverage': '1.00',\
                            'Inception Date': '05/07/2017'}

dict_table_fund_fact = {'Fund Fact':"", \
                            'Morningstar Cat':'Conservative Alloc.', \
                            'NAV':str(nav)+' Mln',\
                            'Class A Share':'',\
                            'ISIN ': 'IT0005261125',\
                            'Bloomberg ': 'FIERITA IM',\
                            'Share Price': str(share),\
                            'Class PIR Share': '',\
                            'ISIN': 'IT0005273575',\
                            'Bloomberg': 'FIERPIR IM',\
                            'Share Price ': str(share),\
                            'Morningstar': 'Bil. Prud',}

dict_table_fund_stat = {'Performance Statistics':"", \
                            'Return Annualized':str(round(avg_annualized_ret*100,nround))+"%", \
                            'Return Monthly':str(round(avg_monthly_ret*100,nround))+"%",\
                            'STD Annualized':str(round(avg_annualized_std*100,nround))+"%",\
                            'STD Monthly': str(round(avg_monthly_std*100,nround))+"%",\
                            'Sharpe Ratio': str(round(sr,nround)),\
                            'Sortino Ratio': str(round(sortino_ratio,nround)),\
                            'Best Month': str(best_month.strftime('%b-%y')),\
                            'Worst Month': str(worst_month.strftime('%b-%y')),\
                            'Return(Inception)': str(round(ret_inception*100,nround))+"%"}

table_fund_info = pd.Series(dict_table_fund_info,index=dict_table_fund_info.keys())
table_fund_fact = pd.Series(dict_table_fund_fact,index=dict_table_fund_fact.keys())
table_fund_stats = pd.Series(dict_table_fund_stat,index=dict_table_fund_stat.keys())
#=============================================================
#PLOTTING
#=============================================================
if cum_ret:

    
    dt_x =pd.DataFrame([raw_data['Date']]).applymap(str).applymap(lambda s: "{}/{}/{}".format(s[6:],s[4:6],s[0:4])).T
    dt_x = dt_x['Date'].tolist()
    dt_x = list(map(dt.datetime.strptime,dt_x,len(dt_x)*['%d/%m/%Y']))
    
    dollar_invested = 100
    dollar_perf = ret.add(1).cumprod()*dollar_invested
    
    #PLOT
    formatter = mdates.DateFormatter('%B %y')
    fig = plt.figure(num=None, figsize=(16, 13), dpi=dpi, )
    
    #ax = fig.add_axes([0,0,1,1])
    ax = plt.plot(dt_x,dollar_perf,color = 'darkblue')
    
    axs = plt.gcf().axes[0]
    axs.xaxis_date()
    axs.xaxis.set_major_formatter(formatter)
    
    plt.gcf().autofmt_xdate(rotation=45)
    
    #plt.title('100€ Invested since Inception\n',fontsize=34,color = 'darkblue')
    plt.grid(True,axis='both',linestyle = '--',linewidth = 0.5)
    
    plt.ylim(bottom=80)
    plt.xlim(min(dt_x))
    
    plt.tick_params(axis='both', which='major', labelsize=18)
    
    [a1,a2,a3,a4] = plt.gca().spines.values()
    #a1.set_visible(False)
    a2.set_visible(False)
    #a3.set_visible(False)
    a4.set_visible(False)
    
    plt.tight_layout()
    plt.savefig('perf_plt'+ext)

if asset_all_barh:
    
    df_assetclass = pd.pivot_table(raw_allocation,['MKT VALUE'],['ASSET CLASS'],aggfunc=np.sum)/raw_allocation['MKT VALUE'].sum()*100
    
    #fig2 =plt.figure(num=None, figsize=(16, 13), dpi=50, )
    fig2,ax = plt.subplots(num=None, figsize=(16, 13), dpi=dpi, )
    
    ax1 = plt.barh(df_assetclass.index.tolist(),df_assetclass['MKT VALUE'],height=0.4, color='darkblue')
    
    fmt = '%.f%%' # Format you want the ticks, e.g. '40%'
    xticks = mtick.FormatStrFormatter(fmt)
    ax.xaxis.set_major_formatter(xticks)
    plt.tick_params(axis='both', which='major', labelsize=18)
    plt.tight_layout()
    plt.savefig('asset_allocation.jpg')
    
if industry_plot:    
    df_industry = (pd.pivot_table(raw_allocation[:-1],['MKT VALUE'],['INDUSTRY SECTOR'],aggfunc=np.sum)/raw_allocation[:-1]['MKT VALUE'].sum()*100).sort_values(['MKT VALUE'])
    fig3,ax3 =   plt.subplots(num=None, figsize=(16, 13), dpi=dpi, ) 
    ind = plt.barh(df_industry.index.tolist(),df_industry['MKT VALUE'],height=0.4, color='darkblue') 
    fmt = '%.f%%' # Format you want the ticks, e.g. '40%'
    xticks = mtick.FormatStrFormatter(fmt)
    ax3.xaxis.set_major_formatter(xticks)
    plt.tick_params(axis='both', which='major', labelsize=18)
    plt.tight_layout()
    plt.savefig('industry_allocation'+ext)
    
if top_holding:

    n = np.size(raw_allocation,0)
    t_10 = 11 - n
    
    df_top_hold =(raw_allocation[:-1].sort_values(['MKT VALUE'], ascending=False))[:t_10]
    df_top_hold['MKT VALUE'] = df_top_hold['MKT VALUE']/raw_allocation['MKT VALUE'].sum()*100
    df_top_hold = df_top_hold.sort_values(['MKT VALUE'], ascending=True)
    
    fig4,ax4 = plt.subplots(num=None, figsize=(16, 13), dpi=150, ) 
    top_10  = plt.barh(df_top_hold['SECURITY NAME'].tolist(),df_top_hold['MKT VALUE'],height=0.4, color='darkblue')
    fmt = '%.f%%' # Format you want the ticks, e.g. '40%'
    xticks = mtick.FormatStrFormatter(fmt)
    ax4.xaxis.set_major_formatter(xticks)
    plt.tick_params(axis='both', which='major', labelsize=18)
    plt.tight_layout()
    plt.savefig('top_10_holding'+ext)
    
   
if pie_all:     
    df_assetclass = pd.pivot_table(raw_allocation,['MKT VALUE'],['ASSET CLASS'],aggfunc=np.sum)/raw_allocation['MKT VALUE'].sum()*100
    df_assetclass['labels']=['CASH','EQUITY','FI-CORP','FI-GOVT']
    df_assetclass.set_index('labels',inplace=True)
    colors = ['darkblue','steelblue','royalblue','midnightblue']
    size = 0.4
    #fig2 =plt.figure(num=None, figsize=(16, 13), dpi=50, )
    fig5,ax5 = plt.subplots(num=None, figsize=(16, 13) )
    
    patches, texts, autotexts = ax5.pie(df_assetclass['MKT VALUE'],autopct='%1.1f ',radius=1.25, wedgeprops=dict(width=size, edgecolor='w'),colors=colors,labels =df_assetclass.index.tolist() )
    ax5.set(aspect="equal")
    
    for i in range(len(texts)):
        autotexts[i].set_fontsize(50)
        texts[i].set_fontsize(50)
        texts[i].set_rotation(-60)
    
    texts[0].set_rotation(-90)
    texts[1].set_rotation(-0)
    texts[2].set_rotation(-0)
    texts[3].set_rotation(-90)
    
    
    fmt = '%.f%%' # Format you want the ticks, e.g. '40%'
    xticks = mtick.FormatStrFormatter(fmt)
    ax5.xaxis.set_major_formatter(xticks)
    #plt.tick_params(labelsize=60)
    plt.tight_layout()
    plt.savefig('pie_asset_allocation'+ext,bbox_inches = 'tight', dpi=dpi,)
    

#table_ret_cumulative = table_ret_cumulative.apply(lambda x: x.str.replace(".",",",1))
table_ret_cumulative.loc['Fund'] = table_ret_cumulative.loc['Fund'].apply(lambda x: x + "%"if len(x)>1 else x)

#table_risk_analysis = table_risk_analysis.apply(lambda x: x.str.replace(".",",",1))
table_risk_analysis.loc['STD'] = table_risk_analysis.loc['STD'].apply(lambda x: x + "%"if len(x)>1 else x)

#table_hist_perf = table_hist_perf.apply(lambda x: x.str.replace(".",",",1))
table_hist_perf = table_hist_perf.apply(lambda x: x.str.replace("nan","",1))
for x in year_hist:
    table_hist_perf.loc[x] = table_hist_perf.loc[x].apply(lambda x: x + "%"if len(x)>1 else x)
    


#====================================================================
#writing to excel
#====================================================================


writer = pd.ExcelWriter('factsheet_test1.xlsx', engine='xlsxwriter')

workbook  = writer.book
worksheet=workbook.add_worksheet('Foglio1')


workbook.filename = 'factsheet_test1.xlsm'
workbook.add_vba_project('./vbaProject.bin')

writer.sheets['Foglio1'] = worksheet

cell_format = workbook.add_format({'num_format': '0.00%'})
cell_format.set_align('right')

worksheet.set_column(2,20,18,cell_format)

table_fund_info.to_excel(writer,sheet_name='Foglio1',startrow=4 , startcol=2,header=True)
table_fund_stats.to_excel(writer,sheet_name='Foglio1',startrow=32 , startcol=2,header=True) 
table_fund_fact.to_excel(writer,sheet_name='Foglio1',startrow=19 , startcol=2,header=True) 


table_ret_cumulative.to_excel(writer,sheet_name='Foglio1',startrow=5 , startcol=6,header=True)
table_risk_analysis.to_excel(writer,sheet_name='Foglio1',startrow=11 , startcol=6,header=True)  


table_hist_perf.to_excel(writer,sheet_name='Foglio1',startrow=17 , startcol=6,header=True)  

worksheet.write('C2', 'Investment Objective')
worksheet.write('C3', 'Finint Economia Reale Italia is an open-end fund incorporated in Italy. The objective of the Fund is capital appreciation. The Fund invests up to 60% of its assets in corporate bonds, 35% of its assets in' \
                + ' small and mid-cap Italian equities, 30% of its assets in EU government/agency bonds. The fund belongs to the "Bilanciati Obbligazionari" category (Assogestioni)')

worksheet.insert_button('G24', {'macro':   'do_2',
                               'caption': 'FACTSHEET',
                               'width':   800,
                               'font_size':40,
                               'height':  300})
writer.save()
workbook.close()



    
