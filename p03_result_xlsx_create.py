# user: shenzhouyang
# time: 2020-10-12

import pandas as pd
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
import numpy as np
import scorecardpy as sc
import xlsxwriter as xl
import sys
sys.path.append('/Users/shenzhouyang/mycode/标准评分卡')
import scorecard_xlsx.p01_data_prepare as p1 

class  result_create():
    '''
    create the finanl xlsx file
    '''
    def __init__(self,data_dict,train_name='train',y_flag='y'):
        self.data_dict = data_dict
        self.train_name = train_name
        self.y_flag = y_flag

        dp = p1.woe_trans(data_dict=self.data_dict,train_name=self.train_name,y_flag=self.y_flag)
        dp.iv_create()
        self.is_iv_create = dp.is_iv_create
        self.tot_iv = dp.tot_iv
        
        dp.psi_create()
        self.is_psi_create = dp.is_psi_create
        self.tot_psi = dp.tot_psi

        self.is_bins_create = dp.is_bins_create
        self.train_bins = dp.train_bins
        self.oot_name = dp.oot_name
        for i in self.oot_name:
            exec("self.%s_bins = dp.%s_bins" % (i,i))



    def result_create(self,path=''):
        writer = pd.ExcelWriter(path+'数据模型结果文档.xlsx', engine='xlsxwriter')
        workbook = writer.book
        # format set
        value_format = workbook.add_format({'border':1, 'align':'center'})
        index_format = workbook.add_format({'bold': True, 'border':1, 'fg_color':'#C5D9F1','font_size':12, 'align':'center'})
        name_format = workbook.add_format({'font':'Arial', 'font_color':'#C00000', 'font_size':20, 'bold':True, 'align':'center'})
        content_format = workbook.add_format({'font':'Arial', 'font_color':'#366092', 'font_size':12, 'bold':True, 'align':'center'})
        
        # create contents
        worksheet = workbook.add_worksheet('目录')
        worksheet.set_column('A:D',22)
        content_value = []
        if self.is_iv_create:
            content_value.append('iv列表')
        if self.is_psi_create:
            content_value.append('psi列表')
        for i in content_value:
            worksheet.write_url('B%d' % (content_value.index(i)+2), 'internal:%s!A1' % i,content_format,string = i)
        worksheet.write_url('B%d' % (len(content_value)+2), 'internal:目录!C1',content_format,string = '变量明细')

        # write sheet "iv list" 
        if self.is_iv_create:
            tot_iv_df = pd.DataFrame(self.tot_iv)
            tot_iv_df.sort_values(by='train',ascending=False,inplace=True)
            worksheet_iv = workbook.add_worksheet('iv列表')
            worksheet_iv.write_url('A1', 'internal:目录!A1',content_format,string = 'Back to Content')
            worksheet_iv.set_column('%s:%s' % (cell_ch(0),cell_ch(tot_iv_df.shape[1]+1)),18)
            for i in range(tot_iv_df.shape[0]):
                worksheet_iv.write(i+1+1,0,tot_iv_df.index[i],index_format)
                for j in range(tot_iv_df.shape[1]):
                    worksheet_iv.write(i+1+1,j+1,tot_iv_df.iloc[i][j],value_format)
            for j in range(tot_iv_df.shape[1]):
                worksheet_iv.write(1,j+1,tot_iv_df.columns[j],index_format)
                worksheet_iv.conditional_format(1+1,j+1,tot_iv_df.shape[0]+1+1,j+1, {'type': 'data_bar'})

        # write sheet "psi list" 
        if self.is_psi_create:
            tot_psi_df = pd.DataFrame(self.tot_psi)
            worksheet_psi = workbook.add_worksheet('psi列表')
            worksheet_psi.write_url('A1', 'internal:目录!A1',content_format,string = 'Back to Content')
            worksheet_psi.set_column('%s:%s' % (cell_ch(0),cell_ch(tot_psi_df.shape[1]+1)),18)
            for i in range(tot_psi_df.shape[0]):
                worksheet_psi.write(i+1+1,0,tot_psi_df.index[i],index_format)
                for j in range(tot_psi_df.shape[1]):
                    worksheet_psi.write(i+1+1,j+1,tot_psi_df.iloc[i][j],value_format)
            for j in range(tot_psi_df.shape[1]):
                worksheet_psi.write(1,j+1,tot_psi_df.columns[j],index_format)
                worksheet_psi.conditional_format(1+1,j+1,tot_psi_df.shape[0]+1+1,j+1, {'type': 'data_bar'})

        # first check if the bins have been created
        if not self.is_bins_create:
            self.bins_create()
            self.is_bins_create = True

        show_var_list = list(tot_iv_df.index[:50])

        # train data describe including bad rate chart & woe value chart
        for var in show_var_list:
            exec("worksheet_%s = workbook.add_worksheet('%s')" % (var,var))
            exec("worksheet_%s.write_url('A1', 'internal:目录!A1',content_format,string = 'Back to Content')" % var)
            exec("worksheet.write_url('C%d', 'internal:%s!A1',content_format,string = '%s')" % (show_var_list.index(var)+2,var,var))
            exec("worksheet_%s.set_column('A:M',18)" % var)
            
            tot_data_name = [self.train_name] +self.oot_name
            for k in tot_data_name:
                m = tot_data_name.index(k)
                var_bin = eval("self.%s_bins[var]" % k)
                # write the bins data 
                exec("worksheet_%s.write(1+(20+var_bin.shape[0])*m,0,'%s',name_format)" % (var,k))
                for i in range(var_bin.shape[0]):
                    exec("worksheet_%s.write(%d+1+1+(20+var_bin.shape[0])*m,0,var_bin.index[%d],index_format)" % (var,i,i))
                    for j in range(var_bin.shape[1]):
                        exec("worksheet_%s.write(%d+1+1+(20+var_bin.shape[0])*m,%d+1,var_bin.iloc[%d][%d],value_format)" % (var,i,j,i,j))
                for j in range(var_bin.shape[1]):
                    exec("worksheet_%s.write(1+(20+var_bin.shape[0])*m,%d+1,var_bin.columns[%d],index_format)" % (var,j,j))
            
                # create bad rate chart
                col_par = {}
                col_par.update({'name':'=%s!E2' % var,'categories':'=%s!C3:C%s' % (var,str(var_bin.shape[0]+2)),'values':'=%s!E%s:E%s' % (var,str(3+(20+var_bin.shape[0])*m),str(var_bin.shape[0]+2+(20+var_bin.shape[0])*m))})
                line_par = {}
                line_par.update({'name':'=%s!H2' % var,'categories':'=%s!C3:C%s' % (var,str(var_bin.shape[0]+2)),'values':'=%s!H%s:H%s' % (var,str(3+(20+var_bin.shape[0])*m),str(var_bin.shape[0]+2+(20+var_bin.shape[0])*m))})
                insert_loc_col = 'A%s' % str(var_bin.shape[0]+4+(20+var_bin.shape[0])*m)
                col_line_chart(col_par,line_par,workbook,eval("worksheet_%s" % var),insert_loc_col,var)

                # create woe value chart
                bar_par = {}
                bar_par.update({'name':'=%s!I2' % var,'categories':'=%s!C3:C%s' % (var,str(var_bin.shape[0]+2)),'values':'=%s!I%s:I%s' % (var,str(3+(20+var_bin.shape[0])*m),str(var_bin.shape[0]+2+(20+var_bin.shape[0])*m))}) 
                insert_loc_bar = 'E%s' % str(var_bin.shape[0]+4+(20+var_bin.shape[0])*m)
                bar_chart(bar_par,workbook,eval("worksheet_%s" % var),insert_loc_bar,var)


        workbook.close()


def cell_ch(i):
    chart_list = [chr(i) for i in range(65,91)]
    if i < 26:
        c = chart_list[i]
    else:
        c = chart_list[int(i/26)-1]+chart_list[i%26]
    return c

def col_line_chart(col_par,line_par,insert_book,insert_sheet,insert_loc,var):
    # Create a new column chart. This will use this as the primary chart.
    column_chart = insert_book.add_chart({'type': 'column'})
    # Configure the data series for the primary chart.
    column_chart.add_series({
        'name':       col_par['name'],
        'categories': col_par['categories'],
        'values':     col_par['values'],
    })
    # Create a new column chart. This will use this as the secondary chart.
    line_chart = insert_book.add_chart({'type': 'line'})
    # Configure the data series for the secondary chart.
    line_chart.add_series({
        'name':       line_par['name'],
        'categories': line_par['categories'],
        'values':     line_par['values'],
        'marker':     {'type': 'automatic'},
        'y2_axis':    True,
    })
    # Combine the charts.
    column_chart.combine(line_chart)
    # Add a chart title and some axis labels. 
    column_chart.set_title({'name': 'Bad Rate chart -- %s' % var})
    column_chart.set_x_axis({'name': 'Bins'})
    column_chart.set_y_axis({'name': 'count_distr'})
    column_chart.set_y2_axis({'name': 'badprob'})
    # Insert the chart into the worksheet
    insert_sheet.insert_chart('%s' % insert_loc, column_chart)

def bar_chart(bar_par,insert_book,insert_sheet,insert_loc,var):
    bar_chart = insert_book.add_chart({'type': 'bar'})
    bar_chart.add_series({
        'name':       bar_par['name'],
        'categories': bar_par['categories'],
        'values':     bar_par['values'],
    })

    # Add a chart title and some axis labels. 
    bar_chart.set_title({ 'name': 'Woe Value Chart -- %s' % var})
    bar_chart.set_x_axis({'name': 'Bins'})
    bar_chart.set_y_axis({'name': 'woe value'})

    # Insert the chart into the worksheet
    insert_sheet.insert_chart('%s' % insert_loc, bar_chart)