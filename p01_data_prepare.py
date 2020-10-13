# user: shenzhouyang
# time: 2020-10-06

import pandas as pd
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
import numpy as np
import scorecardpy as sc
import xlsxwriter as xl


class woe_trans():
    """docstring for woe_trans"""
    def __init__(self,data_dict,train_name='train',y_flag='y'):
        self.data_dict = data_dict
        self.train_name = train_name
        self.y_flag = y_flag
        # 
        self.is_bins_create = False
        self.is_iv_create = False
        self.is_psi_create = False

        # create all oot name
        self.oot_name = []
        for i in self.data_dict.keys():
            exec("self.%s_data = self.data_dict['%s']" % (i,i))
            if i != self.train_name:
                self.oot_name.append(i)


    def bins_breaks_create(self):
        '''
        create the bins breaks based on train data 
        '''
        exec("self.train_bins = sc.woebin(self.%s_data,y=self.y_flag)" % self.train_name)
        self.breaks_dict = {}
        for i in self.train_bins.keys():
            bin_tmp = self.train_bins[i]
            bin_tmp = bin_tmp[~bin_tmp['breaks'].isin(["-inf","inf","missing"]) & ~bin_tmp['is_special_values']]
            if is_numeric_dtype(self.train_data[i]):
                bin_tmp['breaks'] = bin_tmp['breaks'].astype('float')
            if bin_tmp.shape[0] > 0:
                self.breaks_dict.update({i:bin_tmp['breaks'].tolist()})


    def bins_create(self):
        '''
        create all oot data bins based on the train data bins
        '''
        self.bins_breaks_create()
        for i in self.oot_name:
            exec("self.%s_bins = sc.woebin(self.%s_data,y=self.y_flag,breaks_list=self.breaks_dict)" % (i,i))
        # select all keys every data own
        self.keys_used = self.train_bins.keys()
        for i in self.oot_name:
            exec("self.keys_used = list(set(self.keys_used) & set(self.%s_bins.keys()))" % i)

    def iv_create(self):
        '''
        calculate the iv list 
        '''
        # first check if the bins have been created
        if not self.is_bins_create:
            self.bins_create()
            self.is_bins_create = True
        # calculate the train iv first
        self.tot_iv = {}
        train_iv = {}
        for i in self.train_bins.keys():
            train_iv.update({i:sum(self.train_bins[i]['bin_iv'])})
        self.tot_iv.update({'train':train_iv})
        # calculate oot data iv
        for i in self.oot_name:
            exec("%s_iv = {}" % i)
            for j in eval('self.%s_bins.keys()' % i):
                exec("%s_iv.update({'%s':sum(self.%s_bins['%s']['bin_iv'])})" % (i,j,i,j))
            exec("self.tot_iv.update({'%s':%s_iv})" % (i,i))
        self.is_iv_create = True


    def psi_create(self):
        '''
        calculate the psi list
        '''
        # first check if the bins have been created
        if not self.is_bins_create:
            self.bins_create()
            self.is_bins_create = True
        self.tot_psi = {}
        # calculate psi of each oot data based on train data
        for i in self.oot_name:
            exec("%s_psi = {}" % i)
            for j in self.keys_used:
                exec("tmp_psi = pd.merge(self.train_bins['%s'][['variable','bin','count_distr']],self.%s_bins['%s'][['bin','count_distr']],left_on='bin',right_on='bin',suffixes=('_train','_ot'))" % (j,i,j))
                exec("tmp_psi['psi'] = tmp_psi.apply(lambda x: (x['count_distr_train']-x['count_distr_ot'])*(np.log(x['count_distr_train'])-np.log(x['count_distr_ot'])),axis=1)")
                exec("%s_psi.update({'%s':sum(tmp_psi['psi'])})" % (i,j))     
            exec("self.tot_psi.update({'%s':%s_psi})" % (i,i))
        self.is_psi_create = True

