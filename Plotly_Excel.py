# -*- coding: utf-8 -*-
"""
Created on Wed May 22 18:10:49 2019

@author: mtree
"""
import xlwings as xw
from Pickler import  To_pkl, From_pkl
from tkinter.filedialog import askopenfilename, askdirectory
import plotly
import plotly_express as px
import pandas as pd
from tkinter import Tk


def Clear_Vals():
    wb = xw.Book.caller()
    wb.sheets('Dash').range('B6').expand('right').clear_contents()
    wb.sheets('Dash').range('B12').expand('right').clear_contents()
    wb.sheets('Dash').range('B9').expand('right').clear_contents()
    wb.sheets('Dash').range('C11').expand('right').clear_contents()
    wb.sheets('Hide').range('A1').expand('right').clear()
    wb.sheets('Hide').range('A2').expand('right').clear()
    wb.sheets('Hide').range('A3').expand('right').clear()

    df = From_pkl()
    wb.sheets('Hide').range("A1").expand().value = list(df)


def Data_Frame():
    X = PdLoader()
    To_pkl(X)
    Clear_Vals()
    
    
def PdLoader():
    Tk().withdraw() 
    filename = askopenfilename() 
    
    if filename.endswith('.csv'):
        X = pd.read_csv(filename)
        
    elif (filename).endswith('.xlsx') or (filename).endswith('.xls') :
        X = pd.read_excel(filename)
        
    elif filename.endswith('.json'):
        X = pd.read_json(filename)

    elif filename.endswith('.html'):
        X = pd.read_html(filename)

    elif (filename).endswith('.hdf'):
        X = pd.read_hdf(filename)

    elif (filename).endswith('.feather'):
        X = pd.read_feather(filename)

    elif (filename).endswith('.parquet'):
        X = pd.read_parquet(filename)

    elif (filename).endswith('.mspack'):
        X = pd.read_msgpack(filename)

    elif (filename).endswith('.stata'):
        X = pd.read_stata(filename)

    elif (filename).endswith('.sas'):
        X = pd.read_sas(filename)

    elif (filename).endswith('.pkl'):
        X = pd.read_pickle(filename)

    elif (filename).endswith('.sql'):
        X = pd.read_sql(filename)

    elif (filename).endswith('.gbq'):
        X = pd.read_sql(filename)
    
    return(X)


def Plot():
    wb = xw.Book.caller()
    df = From_pkl()
    
    Animation = wb.sheets('Dash').range('B6:B7').value
    inputs = wb.sheets('Dash').range('B12:M12').value
    Layout = wb.sheets('Dash').range('B9:F9').value
    
    Filter = wb.sheets('Dash').range('G3').value
    
    af = Animation[0]
    ag = Animation[1]
    
    Title = Layout[0]
    hover_name = Layout[1]
    Log_x = Layout[3]
    Log_y = Layout[4]

    
    if Filter is not None:
        df = df.query(Filter)
        
    if inputs[0] =='scatter':
        Plot = px.scatter(df,title = Title, x=inputs[1], y=inputs[2], color=inputs[3],
                    size=inputs[4], facet_row=inputs[5], 
                    facet_col=inputs[6],hover_name = hover_name,animation_frame= af ,animation_group=ag,log_x=Log_x,log_y=Log_y, 
                    trendline=inputs[7],marginal_x=inputs[8], marginal_y=inputs[9])
            
    elif inputs[0] == 'line':
        Plot = px.line(df,title = Title,x=inputs[1], y=inputs[2], color=inputs[3], log_x=Log_x,log_y=Log_y,
                       facet_row=inputs[4], facet_col=inputs[5],line_group=inputs[6],line_dash=inputs[7])
            
    elif inputs[0] == 'scatter matrix':
        Plot = px.scatter_matrix(df,title = Title,color=inputs[1])
            
    elif inputs[0] == 'bar':
        Plot = px.bar(df, x=inputs[1],title = Title,y=inputs[2],color=inputs[3], log_x=Log_x, log_y=Log_y, facet_row=inputs[4],facet_col=inputs[5])
            
    elif inputs[0] == 'density':
        Plot = px.density_contour(df, title = Title,x=inputs[1],y=inputs[2],color=inputs[3],facet_row=inputs[4],
                                  facet_col=inputs[5], marginal_x=inputs[8], marginal_y=inputs[7],log_x=Log_x,log_y=Log_y,)
        
    elif  inputs[0] == 'box':
        Plot = px.box(df, title = Title,x=inputs[1],y=inputs[2],color=inputs[3],log_x=Log_x,log_y=Log_y,facet_row=inputs[4],facet_col=inputs[5])
            
    elif inputs[0] == 'histogram':
        Plot = px.histogram(df,title = Title,x=inputs[1],y=inputs[2],log_x=Log_x,log_y=Log_y, color=inputs[3],facet_row=inputs[4],
                            facet_col=inputs[5],histfunc=inputs[6],marginal=inputs[8])
        
    elif inputs[0] == 'violin':
        Plot = px.violin(df,title = Title,x=inputs[1],y=inputs[2],color=inputs[3],facet_row=inputs[4],log_x=Log_x,log_y=Log_y,
                         facet_col=inputs[5])
            
    elif inputs[0] == '3d scatter':
        Plot = px.scatter_3d(df,title = Title,x=inputs[1], y=inputs[2], z=inputs[3],log_x=Log_x,log_y=Log_y,
                          color= inputs[4], size= inputs[5],hover_name = hover_name)
       
    elif inputs[0] == '3d line':
        Plot = px.line_3d(df,title = Title,x=inputs[1], y=inputs[2], z=inputs[3],log_x=Log_x,log_y=Log_y,
                          color= inputs[4], size= inputs[5],)
       
    elif inputs[0] == 'scatter polar':
       Plot = px.scatter_polar(df,title = Title,r= inputs[1] , theta= inputs[2], color= inputs[3],log_x=Log_x,log_y=Log_y, symbol= inputs[4])
            
    elif inputs[0] == 'line polar':
        Plot = px.line_polar(df,title = Title,r= inputs[1], theta= inputs[2] , log_x=Log_x,log_y=Log_y, color= inputs[2], line_close=True)
        
    elif inputs[0] == 'bar polar':
       Plot = px.bar_polar(df,title = Title,r=inputs[1], theta= inputs[2], log_x=Log_x,log_y=Log_y,color=inputs[3])
        
            
    elif inputs[0] == 'parallel_categories':
      Plot = px.parallel_categories(df,title = Title, color=inputs[1])
                              
        
    plotly.offline.plot(Plot)