#!/usr/bin/env python
# coding: utf-8

# In[145]:


import numpy as np
import pandas as pd
from collections import defaultdict
import pulp


# In[146]:


class territory:
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of location of technology
            site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation/non activation of technology
        #z = {} #location of centralized technology
        
        
        
        for M in self.mun:
            for origin in ["house", "building"]:
                for T in self.idtech:
                    #for S in self.S_list:
                    x[(M, origin, T)] = pulp.LpVariable(f"x_{M}_{origin}_{T}", lowBound=0)
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech}
        
        #for T in self.idtech:
         #   if self.scale[T] !=  "Decentralized": 
          #      for S in self.S_list:
           #         z[(T,S)] = pulp.LpVariable(f"z_{T}_{S}")
        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M,"house", T)] for T in self.idtech) == self.home[M], f"HouseWaste_{M}"
            model += pulp.lpSum(x[(M,"building", T)] for T in self.idtech) == self.build[M], f"BuildingWaste_{M}"

            

        ##building can not use home composting
        for M in self.mun:
            model +=pulp.lpSum(x[M,"building", 'HC']) == 0
        
        ##capacity constaints of technology
        for T in self.idtech:
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, origin, T)] for M in self.mun for origin in ["house", "building"]) >= self.lb[T] * y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, origin, T)] for M in self.mun for origin in ["house", "building"] for S in self.S_list) <= self.ub[T] * y[T], f"Max_{T}"
                
        ##constraint of site location
        #for T in self.idtech:
         #   if self.scale[T] != "Decentralized":
          #      for S in self.S_list:
           #         z[(T,S)] == self.site[T][S]
        
        #for T in self.idtech:
         #   if self.scale[T] !=  "Decentralized":
          #      for M in self.mun:
           #         for S in self.S_list:
            #            for origin in ["house", "building"]:
             #               model += pulp.lpSum(z[(T,S)]) == 1

        
            

        #objective
        model += pulp.lpSum(x[(M, origin, T)] * self.GHG[T] for M in self.mun for origin in ["house", "building"] for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for origin in ["house", "building"]:
                for T in self.idtech:
                    var = x[(M, origin, T)]
                    if var.varValue is not None and var.varValue > 1e-3:
                            print(f"  {origin} → {T} : {var.varValue:.1f} tonnes")
       # print("\nImplantation des technologies centralisées :")
        
       # for (T, S), var in z.items():
          #  if var.varValue is not None and var.varValue > 0.5:
           #     print(f"  Technologie {T} implantée à {S}")
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


            
            
    


# In[147]:


data_test = 'territory.xlsx'
data_test_tech = 'technology.xlsx'
tech2 = 'technology2.xlsx'


# In[148]:


class v1: ##most simple, no constraint
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.mass = mass
            self.home = self.mass['Home']
            self.build = self.mass['Building']
            self.total = self.mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        #variables
        x = {} #quantity of waste sent to each technology
        
        for M in self.mun:
            for T in self.idtech:
                x[(M, T)] = pulp.LpVariable(f"x_{M}_{T}", lowBound=0)
        
        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, T)] for T in self.idtech) == self.total[M]
            

        #objective
        model += pulp.lpSum(x[(M, T)] * self.GHG[T] for M in self.mun for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                var = x[(M, T)]
                if var.varValue is not None and var.varValue > 1e-3:
                    print(f"  {M} → {T} : {var.varValue:.1f} tonnes")
        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[149]:


testv1 = v1(data_test,tech2)


# In[150]:


testv1.optimize()


# In[151]:


class v2: ##add the constraint HC vs building
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipaligties
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        #variables
        x = {} #quantity of waste sent to each technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    x[(M, O, T)] = pulp.LpVariable(f"x_{M}_{O}_{T}", lowBound=0)

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T)] for T in self.idtech) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T)] for T in self.idtech) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC']) == 0
            

        #objective
        model += pulp.lpSum(x[(M, O, T)] * self.GHG[T] for M in self.mun for O in self.origin for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    var = x[(M, O, T)]
                    if var.varValue is not None and var.varValue > 1e-3:
                        print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[152]:


testv2 = v2(data_test,tech2)


# In[153]:


testv2.optimize()


# In[154]:


class v3: ##add the constraint of technology capacity, without activation/non activation
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipaligties
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        #variables
        x = {} #quantity of waste sent to each technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    x[(M, O, T)] = pulp.LpVariable(f"x_{M}_{O}_{T}", lowBound=0)

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T)] for T in self.idtech) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T)] for T in self.idtech) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC']) == 0
        
        ##constraint capacity
        for T in self.idtech:
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) >= self.lb[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) <= self.ub[T], f"Max_{T}"
            

        #objective
        model += pulp.lpSum(x[(M, O, T)] * self.GHG[T] for M in self.mun for O in self.origin for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    var = x[(M, O, T)]
                    if var.varValue is not None and var.varValue > 1e-3:
                        print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[155]:


testv3 = v3(data_test,data_test_tech)


# In[156]:


testv3.optimize()


# In[157]:


testv3_2 =v3(data_test,tech2)


# In[158]:


testv3_2.optimize()


# In[159]:


#v3 : pas de solution, car les contraintes sont trop rigides (contrainte minimal qui exige l'envoi de déchets dans chaque technologie)


# In[160]:


class v4: ##add the activation variable
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipaligties
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    x[(M, O, T)] = pulp.LpVariable(f"x_{M}_{O}_{T}", lowBound=0)
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T)] for T in self.idtech) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T)] for T in self.idtech) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC']) == 0
        
        ##constraint capacity
        for T in self.idtech:
            if not pd.isna(self.lb[T]): #HC do not have constraint capacity
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) <= self.ub[T]*y[T], f"Max_{T}"
            

        #objective
        model += pulp.lpSum(x[(M, O, T)] * self.GHG[T] for M in self.mun for O in self.origin for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    var = x[(M, O, T)]
                    if var.varValue is not None and var.varValue > 1e-3:
                        print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[161]:


testv4 = v4(data_test,tech2)


# In[162]:


testv4.optimize()


# In[163]:


class v5: ##add the impact of transport, assuming that centralize industry can be installed in one specific place
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipaligties
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            self.lieu = self.dist['w'] #distance from the only site option to the community m
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    x[(M, O, T)] = pulp.LpVariable(f"x_{M}_{O}_{T}", lowBound=0)
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T)] for T in self.idtech) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T)] for T in self.idtech) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC']) == 0
        
        ##constraint capacity
        for T in self.idtech:
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T)] for M in self.mun for O in self.origin) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
        #impact of transport
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                if self.scale[T] == 'Centralized':
                    trans_GHG[(T,M)] = self.transport * self.lieu[M]
                else:
                    trans_GHG[(T,M)] = 0
        #print (trans_GHG)
                        

        #objective
        
        
        model += pulp.lpSum(x[(M, O, T)] * (self.GHG[T] + trans_GHG[(T,M)]) for M in self.mun for O in self.origin for T in self.idtech) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    var = x[(M, O, T)]
                    if var.varValue is not None and var.varValue > 1e-3:
                        print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[164]:


testv5 = v5(data_test,tech2)


# In[165]:


testv5.optimize()


# In[166]:


class v6: ##add the variable of choice of centralized technology place
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
       
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary") 

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        #impact of transport
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0 ##assumption : people only walk to bring their waste to CC
        
                        

        #objective
        
        
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] +self.GHG[T] ) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            #if self.scale[T] == 'Centralized':
                            print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            #else:
                             #  print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[167]:


terr3 = 'territory3.xlsx'


# In[168]:


tech2 = 'technology2.xlsx'


# In[169]:


testv6 = v6(data_test,tech2)


# In[170]:


testv6.optimize()


# In[171]:


testv6_2 = v6(terr3,tech2)


# In[172]:


testv6_2.optimize()


# In[173]:


testv6_2.optimize()


# In[174]:


testv6_2.S_list


# In[175]:


##v7 boucle qui se répète et donc l'impact minimal est très grand? V7 = FAUX


# In[176]:


class v8: ##add variable : type of coproduct
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            
            
            #list of possible conventional product 
            c_p_list = pd.read_excel(technologies, 'subs_ratio')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            self.sub_ratio = sub_ratio
            
            #impact of conventional product
            c_p_imp = pd.read_excel(technologies, 'conv_prod_impact', index_col=0)
            self.GHG_cp = c_p_imp['GHG']
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary")

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        #impact of transport
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0
        #print (trans_GHG)
        
        #impact of substitution
# dictionnaire des émissions évitées par technologie
        self.sub_GHG = {}
        #self.sub_GHG_total = {}

        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_GHG[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.GHG_cp[conv_p]
        
        print (self.sub_GHG)
        
        self.sub_GHG_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_GHG.items():
            if T not in self.sub_GHG_total:
                self.sub_GHG_total[T] = 0
            self.sub_GHG_total[T] += value
            
        print (self.sub_GHG_total)
        
        for T, total in self.sub_GHG_total.items():
            print(f"{T} → {total:.2f} kgCO₂ évités")
        
        

        #objective
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] + self.GHG[T] - self.sub_GHG_total[T]) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)
       
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            else:
                                print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[177]:


testv8.optimize()


# In[178]:


testv8_2 = v8(terr3,tech2)


# In[179]:


testv8_2.optimize()


# In[180]:


class v9: ##checking of v8 version : define in excel sheet the impact of substitution for each technology
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #possibility of location of technology
            #site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            #self.site = site
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            
            
            #list of possible conventional product 
            c_p_list = pd.read_excel(technologies, 'subs_ratio')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            self.sub_ratio = sub_ratio
            
            #impact of conventional product
            c_p_imp = pd.read_excel(technologies, 'conv_prod_impact', index_col=0)
            self.GHG_cp = c_p_imp['GHG']
            
            #avoided impact from substitution
            avoid_impact = pd.read_excel(technologies ,'sub_imp', index_col=0)
            self.sub_GHG = avoid_impact['sub_imp']
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
                
                
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        
        #impact of transport
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0
        #print (trans_GHG)
        
      
           

        #objective
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] + self.GHG[T] - self.sub_GHG[T]) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)
       
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            else:
                                print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")
        
        


# In[181]:


testv9 = v8(terr3,tech2)


# In[182]:


testv9.optimize()


# In[191]:


class v10: ##add impact of coproduct transport : assumption there is only one market possible for coproduct (E)
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            
            #list of potential sites
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possibility of coproduct market 
            cop_mark = pd.read_excel(municipalities, 'market', index_col=0).fillna(0.0)
            self.cop_mark = cop_mark
           
            
            #distances between technology potentials sites and sites for market
            dist_market = pd.read_excel(municipalities, 'dist_market', index_col = 0).fillna(0.0)
            self.dist_mark = dist_market
            self.one_mark = self.dist_mark['E']
            
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            
            
            #list of possible conventional product 
            c_p_list = pd.read_excel(technologies, 'subs_ratio')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            self.sub_ratio = sub_ratio
            
            #impact of conventional product
            c_p_imp = pd.read_excel(technologies, 'conv_prod_impact', index_col=0)
            self.GHG_cp = c_p_imp['GHG']
            
            #impact of transport of coproduct
            tr_coprod_GHG = pd.read_excel(technologies, 'trans_cop_imp', index_col=0)
            self.tr_cop_GHG = tr_coprod_GHG['GHG']
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        #z = {} #quantity of coproduct sent to market segment
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
        #for T in self.idtech:
         #   for cp in self.coprod_list:
          #      for mark in self.market_list:
           #         z[(T,cp,mark)] = pulp.LpVariable(f"y_{T}_{cp}_{mark}", lowBound = 0)
        

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        
        #impact of transport of waste
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0
        #print (trans_GHG)
        
        #impact of substitution
        self.sub_GHG = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_GHG[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.GHG_cp[conv_p]
        
        #print (self.sub_GHG)
        
        self.sub_GHG_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_GHG.items():
            if T not in self.sub_GHG_total:
                self.sub_GHG_total[T] = 0
            self.sub_GHG_total[T] += value
            
        #print (self.sub_GHG_total)
        
        for T, total in self.sub_GHG_total.items():
            print(f"{T} → {total:.2f} kgCO₂ évités")
        
        #impact of coproduct transport
        self.trans_cop_GHG = {}
        for T in self.idtech:
            for cp in self.coprod_list:
                for S in self.S_list:
                    self.trans_cop_GHG[T,cp,S]= self.coprod[T][cp]*self.tr_cop_GHG[cp] * self.one_mark[S]
                    #for mark in self.market_list:
                        #self.trans_cop_GHG[T,cp,S,mark] = self.cop_mark[cp][mark]*self.coprod[T][cp] * self.tr_cop_GHG[cp] * self.dist_mark[mark][S]
        
        #print (self.trans_cop_GHG) 
        
        
        self.tr_cop_GHG_tot = {}
        
        for ((T,cp,S),val) in self.trans_cop_GHG.items():
            if not (T,S) in self.tr_cop_GHG_tot:
                self.tr_cop_GHG_tot[T,S] =0
            self.tr_cop_GHG_tot[T,S] += val
        
        print (f"total transport : {self.tr_cop_GHG_tot}")
    

        #objective
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] + self.GHG[T] - self.sub_GHG_total[T]+ self.tr_cop_GHG_tot[T,S] )                             for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)
       
        
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            else:
                                print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[192]:


##compost from HC is sent to market --> CORRECT it later


# In[193]:


testv10 = v10(terr3,tech2)


# In[194]:


testv10.optimize()


# In[134]:


##compost from HC are also sent to the same market


# In[267]:


class v11: ##add variable : market for coproduct --> influence of distance between market and treatment place
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            
            #list of potential sites
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark
            
            
            #distances between technology potentials sites and sites for market
            dist_market = pd.read_excel(municipalities, 'dist_market', index_col = 0).fillna(0.0)
            self.dist_mark = dist_market
    
            
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            
            
            #list of possible conventional product 
            c_p_list = pd.read_excel(technologies, 'subs_ratio')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            self.sub_ratio = sub_ratio
            
            #impact of conventional product
            c_p_imp = pd.read_excel(technologies, 'conv_prod_impact', index_col=0)
            self.GHG_cp = c_p_imp['GHG']
            
            #impact of transport of coproduct
            tr_coprod_GHG = pd.read_excel(technologies, 'trans_cop_imp', index_col=0)
            self.tr_cop_GHG = tr_coprod_GHG['GHG']
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        z = {} #quantity of coproduct sent to market segment
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
        
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    z[(T,cp,mark)] = pulp.LpVariable(f"y_{T}_{cp}_{mark}", lowBound = 0)
        

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        #impact of transport of waste to the facility
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0
        #print (trans_GHG)
        
        #impact of substitution
        self.sub_GHG = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_GHG[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.GHG_cp[conv_p]
        
        #print (self.sub_GHG)
        
        self.sub_GHG_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_GHG.items():
            if T not in self.sub_GHG_total:
                self.sub_GHG_total[T] = 0
            self.sub_GHG_total[T] += value
            
        #print (self.sub_GHG_total)
        
        for T, total in self.sub_GHG_total.items():
            print(f"{T} → {total:.2f} kgCO₂ évités")
        
        #impact of coproduct transport
        self.trans_cop_GHG = {}
        for T in self.idtech:
            for cp in self.coprod_list:
                for S in self.S_list:
                    for mark in self.market_list:
                        self.trans_cop_GHG[T,cp,S,mark] = self.coprod[T][cp] * self.tr_cop_GHG[cp] * self.dist_mark[mark][S]
        
        #print (self.trans_cop_GHG) 
        
        self.tr_cop_GHG_tot = {}
        
        for ((T,cp,S,mark),val) in self.trans_cop_GHG.items():
            if T not in self.tr_cop_GHG_tot:
                self.tr_cop_GHG_tot[T] = 0
            self.tr_cop_GHG_tot[T] += val
        
        print (self.tr_cop_GHG_tot)
        
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,cp,mark)] for T in self.idtech) ==0
        
        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                model += pulp.lpSum(z[(T,cp,mark)] for mark in self.market_list) ==                 pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin for S in self.S_list) 
            
    

        #objective
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_GHG[(M, T, S)] + self.GHG[T] - self.sub_GHG_total[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, cp, mark)] * self.tr_cop_GHG[cp] * self.dist_mark[mark][S] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list if self.scale[T] == 'Centralized')
                           #for cp in self.coprod_list for mark in self.market_list)
       
        #(z[(T,cp,mark)]* self.tr_cop_GHG[cp] * self.dist_mark[mark][S])
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            else:
                                print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
    
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    var = z[(T, cp, mark)]
                    if var.varValue is not None and var.varValue > 1e-3:
                        print(f"  {cp} depuis {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[268]:


testv11 = v11(terr3,tech2)


# In[269]:


testv11.optimize()


# In[239]:


ter2 =  ('territory2.xlsx')


# In[270]:


class v11_2: ##ajout de S dans la variable z
    def __init__(self, municipalities, technologies):
        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            #distances between municipalities and potential sites
            dist = pd.read_excel(municipalities, 'dist', index_col=0).fillna(0.0)
            self.dist = dist
            
            #list of potential sites
            self.S_list = np.array(self.dist.columns)
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark
            
            
            #distances between technology potentials sites and sites for market
            dist_market = pd.read_excel(municipalities, 'dist_market', index_col = 0).fillna(0.0)
            self.dist_mark = dist_market
    
            
            
            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            
            
            #list of possible conventional product 
            c_p_list = pd.read_excel(technologies, 'subs_ratio')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            self.sub_ratio = sub_ratio
            
            #impact of conventional product
            c_p_imp = pd.read_excel(technologies, 'conv_prod_impact', index_col=0)
            self.GHG_cp = c_p_imp['GHG']
            
            #impact of transport of coproduct
            tr_coprod_GHG = pd.read_excel(technologies, 'trans_cop_imp', index_col=0)
            self.tr_cop_GHG = tr_coprod_GHG['GHG']
        
        
    
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        z = {} #quantity of coproduct sent to market segment
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        for T in self.idtech:
            for S in self.S_list:
                y[(T,S)] = pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
        
        for T in self.idtech:
            for S in self.S_list:
                for cp in self.coprod_list:
                    for mark in self.market_list:
                        z[(T,S,cp,mark)] = pulp.LpVariable(f"y_{T}_{S}_{cp}_{mark}", lowBound = 0)
        

        
        #constraints
        ##all waste has to be treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]
        
        ##HC can not be used in buildings
        for M in self.mun:
            model += pulp.lpSum(x[M, 'Building', 'HC',S] for S in self.S_list) == 0
        
        ##constraint capacity
        for T in self.idtech:
            for S in self.S_list:
                if not pd.isna(self.lb[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T]*y[(T,S)]
                if not pd.isna(self.ub[T]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T]*y[(T,S)]
        
        self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += pulp.lpSum(y[(T,S)]) == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0
        
        ##decentralized technology can receive only waste from the site where it is installed
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += pulp.lpSum(y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0
        
        
        #impact of transport of waste to the facility
        trans_GHG = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_GHG[(M,T,S)] = self.transport * self.dist[S][M]
                    else:
                        trans_GHG[(M,T,S)] = 0
        #print (trans_GHG)
        
        #impact of substitution
        self.sub_GHG = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_GHG[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.GHG_cp[conv_p]
        
        #print (self.sub_GHG)
        
        self.sub_GHG_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_GHG.items():
            if T not in self.sub_GHG_total:
                self.sub_GHG_total[T] = 0
            self.sub_GHG_total[T] += value
            
        #print (self.sub_GHG_total)
        
        for T, total in self.sub_GHG_total.items():
            print(f"{T} → {total:.2f} kgCO₂ évités")
        
        #impact of coproduct transport
        #self.trans_cop_GHG = {}
        #for T in self.idtech:
         #   for cp in self.coprod_list:
          #      for S in self.S_list:
           #         for mark in self.market_list:
            #            self.trans_cop_GHG[T,cp,S,mark] = self.coprod[T][cp] * self.tr_cop_GHG[cp] * self.dist_mark[mark][S]
        
        #print (self.trans_cop_GHG) 
        
        #self.tr_cop_GHG_tot = {}
        
        #for ((T,cp,S,mark),val) in self.trans_cop_GHG.items():
         #   if T not in self.tr_cop_GHG_tot:
          #      self.tr_cop_GHG_tot[T] = 0
           # self.tr_cop_GHG_tot[T] += val
        
        #print (self.tr_cop_GHG_tot)
        
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
        
        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin) 
            
    

        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_GHG[(M, T, S)] + self.GHG[T] - self.sub_GHG_total[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * self.tr_cop_GHG[cp] * self.dist_mark[mark][S] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)
                           #for cp in self.coprod_list for mark in self.market_list)
       
        #(z[(T,cp,mark)]* self.tr_cop_GHG[cp] * self.dist_mark[mark][S])
        model.solve()
        
        for M in self.mun:
            print(f"\nCommune : {M}")
            for T in self.idtech:
                for O in self.origin:
                    for S in self.S_list:
                        var = x[(M, O, T,S)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {O} → {T} in {S} : {var.varValue:.1f} tonnes")
                            else:
                                print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f}")


# In[271]:


testv11_2 = v11_2(terr3,tech2)


# In[272]:


testv11_2.optimize()


# In[ ]:




