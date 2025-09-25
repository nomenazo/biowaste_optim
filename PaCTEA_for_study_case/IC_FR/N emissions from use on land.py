#!/usr/bin/env python
# coding: utf-8

# In[2]:


import numpy as np
import pandas as pd
from pandas import Series,DataFrame
from scipy.integrate import solve_ivp
from scipy.integrate import odeint
import matplotlib.pyplot as plt
from scipy.special import logsumexp
from math import exp,expm1
from math import sqrt
from scipy import interpolate
from scipy.interpolate import interp1d
from scipy.integrate import trapz
from math import log


# In[3]:


class MFE2:
    def __init__(self, compost_composition, regional, temperature, infiltration, precipitation, soil_pH):
        
        self.temperature = temperature
        self.infiltration = infiltration
        self.precipitation = precipitation
        self.soil_pH = soil_pH
        
        with pd.ExcelFile(compost_composition) as f:
            composition = pd.read_excel(f, 'Composition', index_col=0).fillna(0.0)
            Values = [composition[c] for c in composition.columns]
            self.Values = Values
            self.c = Values[0]  
            
        with pd.ExcelFile(regional) as f:
            self.inf = pd.read_excel(f, 'Infiltration', index_col=0).fillna(0.0)
            self.prec = pd.read_excel(f, 'Precipitation', index_col=0).fillna(0.0)
            
            self.soil = pd.read_excel(f, 'Soil pH', index_col=0).fillna(0.0)
            EF = [self.soil[c] for c in self.soil.columns]
            self.EF = EF
            self.pHfact = EF[0]
            

    def NH3_comp(self):
        try:
            inf_value = self.inf.at[self.temperature, self.infiltration]
            prec_value = self.prec.at[self.temperature, self.precipitation]
            return self.c['NH4'] * inf_value * prec_value
        except KeyError as e:
            print(f"Erreur d'indice dans NH3_comp : {e}")
            return None
    
    def N_NH3_comp(self):
        return self.NH3_comp() * 14/17
    

    
    def N_NO3_comp(self):
        return 0.4 * self.c['Ntotal']
       
    
    def NO3_comp(self):
        return self.N_NO3_comp()*62/14
    
    def N_N2O_comp(self):
        return 0.0125*self.c['Ntotal']
    
    def N2O_comp(self):
        return self.N_N2O_comp()*44/28
    
    def N2_comp(self):
        return 0.09 *self.c['Ntotal']
    
    def MFE_comp(self):
        MFE_comp = (self.c['Norg'])*0.35 + self.c['N-NH4'] + self.c['N-NO3'] - self.N_NH3_comp()-        self.N_NO3_comp() - self.N_N2O_comp() - self.N2_comp()
        return MFE_comp
    
    def Nfert(self):
        Nfert = self.MFE_comp()/(1-self.pHfact[self.soil_pH] - 0.3 - 0.0125 - 0.09)
        return Nfert
    
    def fertilizer(self):#mass total of fertilizer
        return self.Nfert() * (80/28)
    
    def NH3_fert(self):
        
        return self.pHfact[self.soil_pH]*self.Nfert() * 17/14
    
    def NO3_fert(self):
        return 0.3 * self.Nfert() * 62/14
    
    def N2O_fert(self):
        return 0.0125 * self.Nfert() * 44/28
    
    def NH3net(self):
        return self.NH3_comp() - self.NH3_fert()
    
    def NO3net(self):
        return self.NO3_comp() -self.NO3_fert()
    
    def N2Onet(self):
        return self.N2O_comp() - self.N2O_fert()


# In[5]:


#IMPORT THE EXCEL FILE OF YOUR SCENARIO AND CHANGE data0 TO IT
data0 = 'N_composition_compost.xlsx'
#DO NOT CHANGE data1 
data1 = 'use_on_land_correction_fact_emissions.xlsx'


# In[ ]:




