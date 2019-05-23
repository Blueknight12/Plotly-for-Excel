# -*- coding: utf-8 -*-
"""
Created on Wed Feb 20 21:14:52 2019

@author: mtree
"""

import pandas as pd
import numpy as np
import os
import xlwings as xw

loc = os.path.dirname(os.path.abspath(__file__))+r'\My_PKL.pkl'

def To_pkl(q):
    q.to_pickle(loc)

def From_pkl():
   q = pd.read_pickle(loc)
   return(q)
   

