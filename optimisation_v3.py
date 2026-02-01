#!/usr/bin/env python
# coding: utf-8

# In[11]:


import numpy as np
import pandas as pd
from collections import defaultdict
import pulp


# In[12]:


#add the distance collection


# In[13]:


class v8:  
    def __init__(self, municipalities, technologies, method, ind):
        
        self.ind = ind
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites, transfer site, and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            #self.GHG = tech['GHG'] #impact of technnology
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
            

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            
            ###Transport of coproduct
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
        
    
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
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,T,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,T,S)] = 0
        
        
        #print(trans_I)
        
        #impact of substitution/avoided impact
        self.sub_I = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_I[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.I_av[conv_p]
        
        self.sub_I_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_I.items():
            if T not in self.sub_I_total:
                self.sub_I_total[T] = 0
            self.sub_I_total[T] += value
            
        for T, total in self.sub_I_total.items():
            print(f"{T} → {total:.2f} {self.unit} avoided")
        
        #impact net of coproduct use:
        self.use_I = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.use_I[T,cp,conv_p] = self.coprod[T][cp]*(self.I_u_cop[cp])
        
        self.use_I_total = {}
        
        for ((T, cp, conv_p), val) in self.use_I.items():
            if T not in self.use_I_total:
                self.use_I_total[T] = 0
            self.use_I_total[T] += val
        
        for T, total in self.use_I_total.items():
            print(f"{T} → {total:.2f} {self.unit} related to use of coproducts")

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M, T, S)] + self.I_T[T] - self.sub_I_total[T]+self.use_I_total[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * self.I_tr_cop[cp] * self.dist[mark][S] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            #if self.scale[T] == 'Centralized':
                            print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                             #   print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f} {self.unit}")
        
        #for v in model.variables():
         #   print(v.name, "=", v.varValue)


# In[14]:


terr = 'ARC_simple.xlsx'
impact = 'impact.xlsx'
tech = 'tech_ARC_simple.xlsx'


# In[15]:


testv8 = v8(terr, tech, impact, 'Carbon')


# In[16]:


testv8.optimize()


# In[17]:


testv8_water = v8(terr, tech, impact, 'Water')


# In[18]:


testv8_water.optimize()


# In[19]:


#tester si les solutions decentralisées (compostage) produisent des composts qui ne substituent pas des engrais


# In[20]:


terr2 = 'ARC_simple_2.xlsx'
impact2 = 'impact2.xlsx'
tech2 = 'tech_ARC_simple2.xlsx'


# In[21]:


testv8_2 = v8(terr2, tech2, impact2, 'Carbon')


# In[22]:


testv8_2.optimize()


# In[23]:


testv8_2.coprod_list


# In[24]:


#valorisation locale des coproduits issus des technologies decentralisées --> transport coproduit decentralisé = 0 (compost)


# In[25]:


class v9:  
    def __init__(self, municipalities, technologies, method, ind):
        
        self.ind = ind
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites, transfer site, and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            #self.GHG = tech['GHG'] #impact of technnology
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
            

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            
            ###Transport of coproduct
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
        
    
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
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,T,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,T,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp[(S,cp,mark)] = 0
        
        
        #print(trans_I)
        
        #impact of substitution/avoided impact
        self.sub_I = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.sub_I[T,cp,conv_p] = self.coprod[T][cp]*self.sub_ratio[cp][conv_p]*self.I_av[conv_p]
        
        self.sub_I_total = {}
        
        for ((T, cp, conv_p), value) in self.sub_I.items():
            if T not in self.sub_I_total:
                self.sub_I_total[T] = 0
            self.sub_I_total[T] += value
            
        for T, total in self.sub_I_total.items():
            print(f"{T} → {total:.2f} {self.unit} avoided")
        
        #impact net of coproduct use:
        self.use_I = {}
    
        for T in self.idtech:  # technologies (HC, IC, AD-a, AD-b)
            for cp in self.coprod_list:  # coproduits : Comp, Dig, Heat, Elect
                for conv_p in self.c_p_list:  # produits conventionnels : Fertilizer, Heat, Ele
                    self.use_I[T,cp,conv_p] = self.coprod[T][cp]*(self.I_u_cop[cp])
        
        self.use_I_total = {}
        
        for ((T, cp, conv_p), val) in self.use_I.items():
            if T not in self.use_I_total:
                self.use_I_total[T] = 0
            self.use_I_total[T] += val
        
        for T, total in self.use_I_total.items():
            print(f"{T} → {total:.2f} {self.unit} related to use of coproducts")

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
        
        ##distance of coproduct transport
        
                    
            



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M, T, S)] + self.I_T[T] - self.sub_I_total[T]+self.use_I_total[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * trans_I_cp[(S,cp,mark)] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f} {self.unit}")
        
        #for v in model.variables():
         #   print(v.name, "=", v.varValue)


# In[235]:


terr_sc1 = 'ARC_scenario1.xlsx'


# In[236]:


sc1 = v9(terr_sc1, tech2, impact2, 'Carbon')


# In[237]:


sc1.optimize()


# In[238]:


sc1.dist


# In[239]:


sc1.dist['Compiègne']['X']


# In[240]:


sc1.mun


# In[241]:


sc1W = v9(terr_sc1, tech2, impact2, 'Water')


# In[242]:


sc1W.optimize()


# In[243]:


#Impact technologie = Impact total -- sans transport


# In[402]:


class v10:  
    def __init__(self, municipalities, technologies, method, ind):
        
        self.ind = ind
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            
            tech = pd.read_excel(technologies, 'tech', index_col=0)
            #self.GHG = tech['GHG'] #impact of technnology
            self.lb = tech['l.b'] #lower bound constraint
            self.ub = tech['u.b'] #upper bound constraint
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
        
    
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
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp[(S,cp,mark)] = 0
        
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
        
        ##distance of coproduct transport
        
                    
            



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M,S)] + self.I_T[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * trans_I_cp[(S,cp,mark)] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")
        
        #print("\n=== Activation des technologies HC ===")
        
        #for S in self.S_list:
         #   print(f"y(HC,{S}) = {y[('HC', S)].varValue}")
               
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f} {self.unit}")


# In[403]:


impact3 = 'impact3.xlsx'
tech3 = 'tech_ARC_simple3.xlsx'


# In[362]:


sc1_v10 = v10(terr_sc1, tech3, impact3, 'Carbon')


# In[363]:


sc1_v10.optimize()


# In[364]:


sc1_v10_W = v10(terr_sc1, tech3, impact3, 'Water')


# In[365]:


#prise en compte de contrainte de capacité : 10% de leur capacité annuelle


# In[366]:


tech3_ub = 'tech_ARC_simple3_ub.xlsx'


# In[367]:


sc1_v10_ub = v10(terr_sc1, tech3_ub, impact3, 'Carbon')


# In[368]:


sc1_v10_ub.optimize()


# In[369]:


sc1_v10_ub_W = v10(terr_sc1, tech3_ub, impact3, 'Water')


# In[370]:


sc1_v10_ub_W.optimize()


# In[371]:


impact3_HC_ss_sub = 'impact3_HC_sans_sub.xlsx'


# In[372]:


#scénario où il n'y a pas de substitution pour le HC


# In[373]:


sc1_v10_ub_HC = v10(terr_sc1, tech3_ub, impact3_HC_ss_sub, 'Carbon')


# In[374]:


sc1_v10_ub_HC.optimize()


# In[375]:


#impacts d'utilisation des coproduits sont tous inclus


# In[661]:


impact4 = 'impact4.xlsx' #impacts d'utilisation des coproduits sont tous inclus


# In[662]:


sc1_cop_use = v10(terr_sc1, tech3_ub, impact4, 'Carbon')


# In[663]:


sc1_cop_use.optimize()


# In[664]:


sc1_cop_use_W = v10(terr_sc1, tech3_ub, impact4, 'Water')


# In[665]:


#contraintes de capacité de technologies existantes différentes de celles à implanter


# In[666]:


class v11:  
    def __init__(self, municipalities, technologies, method, ind):
        
        self.ind = ind
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            
            #mass of waste from individual home and buildings
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.home = mass['Home']
            self.build = mass['Building']
            self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
        
    
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
                if not pd.isna(self.lb[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T][S]*y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T][S]*y[(T,S)]
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp[(S,cp,mark)] = 0
        
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
        
        ##distance of coproduct transport
        
                    
            



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M,S)] + self.I_T[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * trans_I_cp[(S,cp,mark)] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")

        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        print(f"Impact environnemental total minimal : {pulp.value(model.objective):.2f} {self.unit}")


# In[667]:


ARC_sc3 = 'ARC_scenario3.xlsx'


# In[668]:


tech_ARC_sc3 = 'tech_ARC_sc3.xlsx'


# In[669]:


sc3 = v11(ARC_sc3, tech_ARC_sc3, impact4, 'Carbon')


# In[670]:


sc3.optimize()


# In[671]:


sc3_W = v11(ARC_sc3, tech_ARC_sc3, impact4, 'Water')


# In[672]:


#ajout du taux de tri


# In[673]:


class v12:  
    def __init__(self, municipalities, technologies, method, ind, sort):
        
        self.ind = ind
        self.sort = sort
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #total mass of waste generated
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            
            self.home = mass['Home'] *self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total']  * self.sort
            
            self.unsort = self.tot_mass * (1 - self.sort)
            
            
            #mass of waste sorted from individual home and buildings
            #mass = pd.read_excel(municipalities, 'waste_sorted', index_col=0).fillna(0.0)
            #self.home = mass['Home']
            #self.build = mass['Building']
            #self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            self.I_inc = self.I_T['Inc']
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
        
    
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
                if not pd.isna(self.lb[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T][S]*y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T][S]*y[(T,S)]
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp[(S,cp,mark)] = 0
        
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
                    
            



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M,S)] + self.I_T[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * trans_I_cp[(S,cp,mark)] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

        model.solve()
        
        #print("\n=== Valeurs des variables d’activation y(T,S) ===")
        
        for T in self.idtech:
            for S in self.S_list:
                val = y[(T, S)].varValue
                #print(f"y[{T},{S}] = {val}")

        
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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")

        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        
        self.tot_imp = pulp.value(model.objective) +  (self.unsort * self.I_inc)
        print(f"Impact environnemental total minimal : {(self.tot_imp):.2f} {self.unit}")


# In[680]:


##scenario où la substitution est nulle pour les énergies


# In[681]:


impact5 = 'impact5_energie_decarbonee.xlsx'


# In[682]:


sc3_sust = v12(ARC_sc3, tech_ARC_sc3, impact5, 'Carbon', 0.8)


# In[683]:


sc3_sust.optimize()


# In[684]:


sc3_sust_0_3 = v12(ARC_sc3, tech_ARC_sc3, impact5, 'Carbon', 0.3)


# In[685]:


sc3_sust_0_3.optimize()


# In[686]:


##scénarios où l'on utilise les sites existants qui acceptent les biodéchets + Villers-st-Paul


# In[687]:


ARC_reel = 'ARC_reel.xlsx'
tech_ARC_reel = 'tech_ARC_reel.xlsx'


# In[688]:


#Civic-led pathway


# In[689]:


sc_reel_civic_led = v12(ARC_reel, tech_ARC_reel, impact4, 'Carbon', 0.8)


# In[690]:


sc_reel_civic_led.optimize()


# In[691]:


sc_reel_civic_led_EQ = v12(ARC_reel, tech_ARC_reel, impact4, 'Total ecosystem quality', 0.8)


# In[692]:


sc_reel_civic_led_EQ.optimize()


# In[693]:


sc_reel_civic_led_HH = v12(ARC_reel, tech_ARC_reel, impact4, 'Total human health', 0.8)


# In[694]:


sc_reel_civic_led_HH.optimize()


# In[695]:


sc_reel_sust = v12(ARC_reel, tech_ARC_reel, impact5, 'Carbon', 0.8)


# In[696]:


sc_reel_sust.optimize()


# In[697]:


##scenario extrême : energie_decarb - population desengagée #mismatch


# In[698]:


impact6 = 'impact6_decarb_deseng.xlsx'


# In[699]:


mismatch = v12(ARC_reel, tech_ARC_reel, impact6, 'Carbon', 0.2)


# In[700]:


mismatch.optimize()


# In[701]:


##sustainable pathway, avec scénario réel


# In[702]:


sust = v12(ARC_reel, tech_ARC_reel, impact5, 'Carbon', 0.8)


# In[703]:


sust.optimize()


# In[704]:


##desengagement total


# In[705]:


impact7 ='impact7_diseng_tot.xlsx'


# In[706]:


diseng_path = v12(ARC_reel, tech_ARC_reel, impact7, 'Carbon', 0.2)


# In[707]:


diseng_path.optimize()


# In[486]:


ARC_reel_ss_inv = 'ARC_reel_sans_inv.xlsx'


# In[487]:


tech_reel_ss_inv = 'tech_ARC_reel_sans_inv.xlsx'


# In[489]:


civic_ss_inv = v12(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Carbon', 0.8)


# In[490]:


civic_ss_inv.optimize()


# In[710]:


civic_ss_inv_EQ = v12(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total ecosystem quality', 0.8)


# In[711]:


civic_ss_inv_EQ.optimize()


# In[712]:


civic_ss_inv_HH = v12(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total human health', 0.8)


# In[713]:


civic_ss_inv_HH.optimize()


# In[714]:


#v13: stockage des résultats - calcul des impacts pour les autres indicateurs


# In[715]:


class v13:  
    def __init__(self, municipalities, technologies, method, ind, sort, ind1, ind2):
        
        self.ind = ind
        self.sort = sort
        self.ind1= ind1
        self.ind2 = ind2

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #total mass of waste generated
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            
            self.home = mass['Home'] *self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total']  * self.sort
            
            self.unsort = self.tot_mass * (1 - self.sort)
            
            
            #mass of waste sorted from individual home and buildings
            #mass = pd.read_excel(municipalities, 'waste_sorted', index_col=0).fillna(0.0)
            #self.home = mass['Home']
            #self.build = mass['Building']
            #self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            self.unit1 = imp[self.ind1][0]
            self.unit2 = imp[self.ind2][0]
            
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            self.tr_w1 = imp[self.ind1][1]
            self.tr_w2 = imp[self.ind2][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            self.col_w1 = imp[self.ind1][2]
            self.col_w2 = imp[self.ind2][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            self.I_T1 = tech[self.ind1]
            self.I_T2 = tech[self.ind2]
            
            self.I_inc = self.I_T['Inc']
            self.I_inc1 = self.I_T1['Inc']
            self.I_inc2 = self.I_T2['Inc']
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            
            self.I_tr_cop = tr_cop[self.ind]
            self.I_tr_cop1 = tr_cop[self.ind1]
            self.I_tr_cop2 = tr_cop[self.ind2]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            
            self.I_av = av[self.ind]
            self.I_av1 = av[self.ind1]
            self.I_av2 = av[self.ind2]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            
            self.I_u_cop = u_cop[self.ind]
            self.I_u_cop1 = u_cop[self.ind1]
            self.I_u_cop2 = u_cop[self.ind2]
        
    
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
                if not pd.isna(self.lb[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T][S]*y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T][S]*y[(T,S)]
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I[(M,S)] =(self.col_w*self.dist_col[M])+ (self.tr_w * self.dist[M][S]) 
                    else:
                        trans_I[(M,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp[(S,cp,mark)] = 0
        
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
                    
            



        #objective
        
        model += pulp.lpSum(x[(M, O, T, S)] * (trans_I[(M,S)] + self.I_T[T])                            for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * trans_I_cp[(S,cp,mark)] for T in self.idtech for cp in self.coprod_list for mark in self.market_list 
                     for S in self.S_list)

        model.solve()
        
        #stockage des résultats
        self.x_opt = {k: v.varValue for k, v in x.items() if v.varValue is not None}
        self.y_opt = {k: v.varValue for k, v in y.items() if v.varValue is not None}
        self.z_opt = {k: v.varValue for k, v in z.items() if v.varValue is not None}

        
        #print("\n=== Valeurs des variables d’activation y(T,S) ===")
        
        for T in self.idtech:
            for S in self.S_list:
                val = y[(T, S)].varValue
                #print(f"y[{T},{S}] = {val}")

        
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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")

        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        
        self.tot_imp = pulp.value(model.objective) +  (self.unsort * self.I_inc)
        print(f"Impact environnemental total minimal : {(self.tot_imp):.2f} {self.unit}")
        
    
    def impact_tot(self):
        
        total1 = 0
        total2 = 0
        
        
    #Calcule les impacts environnementaux pour plusieurs catégories en utilisant la solution optimale déjà trouvée.

        # ---------- Impact traitement + transport des déchets ----------
        for (M, O, T, S), val in self.x_opt.items():
# transport déchets
            if self.scale[T] == 'Centralized':
                trans1 = (self.col_w1 * self.dist_col[M]) + (self.tr_w1 * self.dist[M][S])
                trans2 = (self.col_w2 * self.dist_col[M]) + (self.tr_w2 * self.dist[M][S])
            else:
                trans1 = 0.0
                trans2 = 0

            # traitement
            total1 += val * (trans1 + self.I_T1[T])
            total2 += val * (trans2 + self.I_T2[T])

        # ---------- Transport des coproduits ----------
        for (T, S, cp, mark), val in self.z_opt.items():
            if self.scale[T] == 'Centralized':
                total1 += val * self.I_tr_cop1[cp] * self.dist[mark][S]
                total2 += val * self.I_tr_cop2[cp] * self.dist[mark][S]
            else:
                total1 += 0
                total2 += 0


        # ---------- Incinération des déchets non triés ----------
        total1 += self.unsort * self.I_inc1
        total2 += self.unsort * self.I_inc2
        

        print(f"Impact total {self.ind1} : {total1:.2f} {self.unit1}")
        print(f"Impact total {self.ind2} : {total2:.2f} {self.unit2}")
        

        
        


# In[716]:


civic_SI_other_ind = v13(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Carbon', 0.8, 'Total ecosystem quality', 'Total human health')


# In[717]:


civic_SI_other_ind.optimize()


# In[718]:


civic_SI_other_ind.impact_tot()


# In[719]:


civic_SI_other_ind_EQ = v13(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total ecosystem quality', 0.8, 'Carbon', 'Total human health')


# In[720]:


civic_SI_other_ind_EQ.optimize()


# In[721]:


civic_SI_other_ind_EQ.impact_tot()


# In[722]:


civic_SI_other_ind_HH = v13(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total human health', 0.8, 'Carbon', 'Total ecosystem quality')


# In[723]:


civic_SI_other_ind_HH.optimize()


# In[724]:


civic_SI_other_ind_HH.impact_tot()


# In[725]:


civic_other_ind_Carb = v13(ARC_reel, tech_ARC_reel, impact4, 'Carbon', 0.8, 'Total ecosystem quality', 'Total human health')


# In[726]:


civic_other_ind_Carb.optimize()


# In[727]:


civic_other_ind_Carb.impact_tot()


# In[728]:


civic_other_ind_EQ = v13(ARC_reel, tech_ARC_reel, impact4, 'Total ecosystem quality', 0.8, 'Carbon', 'Total human health')


# In[729]:


civic_other_ind_EQ.optimize()


# In[730]:


civic_other_ind_EQ.impact_tot()


# In[731]:


civic_other_ind_HH = v13(ARC_reel, tech_ARC_reel, impact4, 'Total human health', 0.8, 'Carbon', 'Total ecosystem quality')


# In[732]:


civic_other_ind_HH.optimize()


# In[655]:


civic_other_ind_HH.impact_tot()


# In[656]:


#multi-objectif


# In[657]:


class v14:  
    def __init__(self, municipalities, technologies, method, ind, sort, ind1, ind2, W1, W2):
        
        self.ind = ind
        self.sort = sort
        self.ind1= ind1
        self.ind2 = ind2
        self.W1 = W1
        self.W2 = W2
        
        self.N1 = 173048 #EQ
        self.N2 = 2.42 #HH
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #total mass of waste generated
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            
            self.home = mass['Home'] *self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total']  * self.sort
            
            self.unsort = self.tot_mass * (1 - self.sort)
            
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            self.unit1 = imp[self.ind1][0]
            self.unit2 = imp[self.ind2][0]
            
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            self.tr_w1 = imp[self.ind1][1]
            self.tr_w2 = imp[self.ind2][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            self.col_w1 = imp[self.ind1][2]
            self.col_w2 = imp[self.ind2][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            self.I_T1 = tech[self.ind1]
            self.I_T2 = tech[self.ind2]
            
            self.I_inc = self.I_T['Inc']
            self.I_inc1 = self.I_T1['Inc']
            self.I_inc2 = self.I_T2['Inc']
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            
            self.I_tr_cop = tr_cop[self.ind]
            self.I_tr_cop1 = tr_cop[self.ind1]
            self.I_tr_cop2 = tr_cop[self.ind2]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            
            self.I_av = av[self.ind]
            self.I_av1 = av[self.ind1]
            self.I_av2 = av[self.ind2]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            
            self.I_u_cop = u_cop[self.ind]
            self.I_u_cop1 = u_cop[self.ind1]
            self.I_u_cop2 = u_cop[self.ind2]
        
    
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
                if not pd.isna(self.lb[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T][S]*y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T][S]*y[(T,S)]
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        for M in self.mun:
            for S in self.S_list:
                for T in self.idtech:
                    if self.scale[T] == 'Decentralized':
                        if M != S :
                            model += (y[(T,S)]) == 0
                            model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I1 = {}
        trans_I2 = {}
        
        for T in self.idtech:
            for M in self.mun:
                for S in self.S_list:
                    if self.scale[T] == 'Centralized':
                        trans_I1[(M,S)] = (self.col_w1 * self.dist_col[M]) + (self.tr_w1 * self.dist[M][S])
                        trans_I2[(M,S)] = (self.col_w2 * self.dist_col[M]) + (self.tr_w2 * self.dist[M][S]) 
                    else:
                        trans_I1[(M,S)] = 0
                        trans_I2[(M,S)] = 0
                        
        #impact of transport of coproduct
        trans_I_cp1 = {}
        trans_I_cp2 = {}
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for T in self.idtech:
                        if self.scale[T] == 'Centralized':
                            trans_I_cp1[(S,cp,mark)] = self.I_tr_cop1[cp] * self.dist[mark][S]
                            trans_I_cp2[(S,cp,mark)] = self.I_tr_cop2[cp] * self.dist[mark][S]
                        else:
                            trans_I_cp1[(S,cp,mark)] = 0
                            trans_I_cp2[(S,cp,mark)] = 0
        
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
                    
            



        #objective
        
        model +=  pulp.lpSum(x[(M, O, T, S)] * (((self.W1/self.N1)*(trans_I1[(M,S)] + self.I_T1[T]))+                                                ((self.W2/self.N2)*(trans_I2[(M,S)] + self.I_T2[T])))                             for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)         + pulp.lpSum(z[(T, S, cp, mark)] * (((self.W1/self.N1)*trans_I_cp1[(S,cp,mark)])+                                            ((self.W2/self.N2)*trans_I_cp2[(S,cp,mark)]))                     for T in self.idtech for cp in self.coprod_list for mark in self.market_list for S in self.S_list) 

        model.solve()
        
        #stockage des résultats
        self.x_opt = {k: v.varValue for k, v in x.items() if v.varValue is not None}
        self.y_opt = {k: v.varValue for k, v in y.items() if v.varValue is not None}
        self.z_opt = {k: v.varValue for k, v in z.items() if v.varValue is not None}

        
        #print("\n=== Valeurs des variables d’activation y(T,S) ===")
        
        for T in self.idtech:
            for S in self.S_list:
                val = y[(T, S)].varValue
                #print(f"y[{T},{S}] = {val}")

        
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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")

        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        
        self.tot_imp = pulp.value(model.objective) +  (self.unsort * ((self.W1/self.N1)*(self.I_inc1)+                                                                      (self.W2/self.N2)*(self.I_inc2)))
                                                                      
        print(f"Impact environnemental total minimal : {(self.tot_imp):.2f}")
        
    
    def impact_tot(self):
        
        total = 0
        total1 = 0
        total2 = 0
        
        
    #Calcule les impacts environnementaux pour plusieurs catégories en utilisant la solution optimale déjà trouvée.

        # ---------- Impact traitement + transport des déchets ----------
        for (M, O, T, S), val in self.x_opt.items():
# transport déchets
            if self.scale[T] == 'Centralized':
                trans1 = (self.col_w1 * self.dist_col[M]) + (self.tr_w1 * self.dist[M][S])
                trans2 = (self.col_w2 * self.dist_col[M]) + (self.tr_w2 * self.dist[M][S])
                trans = (self.col_w * self.dist_col[M]) + (self.tr_w * self.dist[M][S])
                
            else:
                trans1 = 0.0
                trans2 = 0
                trans = 0

            # traitement
            total1 += val * (trans1 + self.I_T1[T])
            total2 += val * (trans2 + self.I_T2[T])
            total += val * (trans + self.I_T[T])
    

        # ---------- Transport des coproduits ----------
        for (T, S, cp, mark), val in self.z_opt.items():
            if self.scale[T] == 'Centralized':
                total1 += val * self.I_tr_cop1[cp] * self.dist[mark][S]
                total2 += val * self.I_tr_cop2[cp] * self.dist[mark][S]
                total += val * self.I_tr_cop2[cp] * self.dist[mark][S]
            else:
                total1 += 0
                total2 += 0
                total += 0


        # ---------- Incinération des déchets non triés ----------
        total1 += self.unsort * self.I_inc1
        total2 += self.unsort * self.I_inc2
        total += self.unsort * self.I_inc
        

        print(f"Impact total {self.ind1} : {total1:.2f} {self.unit1}")
        print(f"Impact total {self.ind2} : {total2:.2f} {self.unit2}")
        print(f"Impact total {self.ind} : {total:.2f} {self.unit}")
        


# In[739]:


test_multi = v14(ARC_reel, tech_ARC_reel, impact4, 'Carbon', 0.8, 'Total ecosystem quality', 'Total human health',1, 0)


# In[740]:


test_multi.optimize()


# In[741]:


test_multi.impact_tot()


# In[919]:


class v13_debugg:  
    def __init__(self, municipalities, technologies, method, ind, sort, ind1, ind2):
        
        self.ind = ind
        self.sort = sort
        self.ind1= ind1
        self.ind2 = ind2


        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #total mass of waste generated
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            
            self.home = mass['Home'] *self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total']  * self.sort
            
            self.unsort = self.tot_mass * (1 - self.sort)
            
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            self.unit1 = imp[self.ind1][0]
            self.unit2 = imp[self.ind2][0]

            
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            self.tr_w1 = imp[self.ind1][1]
            self.tr_w2 = imp[self.ind2][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            self.col_w1 = imp[self.ind1][2]
            self.col_w2 = imp[self.ind2][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            self.I_T1 = tech[self.ind1]
            self.I_T2 = tech[self.ind2]
            
            self.I_inc = self.I_T['Inc']
            self.I_inc1 = self.I_T1['Inc']
            self.I_inc2 = self.I_T2['Inc']
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            
            self.I_tr_cop = tr_cop[self.ind]
            self.I_tr_cop1 = tr_cop[self.ind1]
            self.I_tr_cop2 = tr_cop[self.ind2]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            
            self.I_av = av[self.ind]
            self.I_av1 = av[self.ind1]
            self.I_av2 = av[self.ind2]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            
            self.I_u_cop = u_cop[self.ind]
            self.I_u_cop1 = u_cop[self.ind1]
            self.I_u_cop2 = u_cop[self.ind2]
        
    
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
                if not pd.isna(self.lb[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) >= self.lb[T][S]*y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) <= self.ub[T][S]*y[(T,S)]
        
        #self.transport = 2e-1 #kgCO2/t.km
        
        ##a technology can be installed only at a site which can host it
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin for M in self.mun) == 0
        
        
        ##decentralized technology can receive only waste from the site where it is installed
        #(this constraint may be unnecessary)
        #for M in self.mun:
         #   for S in self.S_list:
          #      for T in self.idtech:
           #         if self.scale[T] == 'Decentralized':
            #            if M != S :
                            #model += (y[(T,S)]) == 0
             #               model +=  pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0  
                            #--> this last line seems to mean that decentralized process should not exist ...

        

        
        
        #impact of transport of waste to the facility (collection+transpot to facility)
        trans_I = {}
        
        
        #for T in self.idtech:
        for M in self.mun:
            for S in self.S_list:
                #if self.scale[T] == 'Centralized':
                trans_I[(M,S)] = (self.col_w * self.dist_col[M]) + (self.tr_w * self.dist[M][S])
                #else:
                    #trans_I[(M,S)] = 0
                        
                        
        #impact of transport of coproduct
        trans_I_cp = {}
        
        for S in self.S_list:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    #for T in self.idtech:
                        #if self.scale[T] == 'Centralized':
                    trans_I_cp[(S,cp,mark)] = self.I_tr_cop[cp] * self.dist[mark][S]
                        #else:
                         #   trans_I_cp[(S,cp,mark)] = 0
                            
        
    
 

        ##all coproduct have to find its market.
        for cp in self.coprod_list:
            for T in self.idtech:
                for S in self.S_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                     pulp.lpSum(x[(M,O,T,S)]* self.coprod[T][cp] for M in self.mun                            for O in self.origin)
                    
        ##constraint : coproduct should be sent to a market which can valorize it.
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) ==0
                    
            



        #objective
        
        model += (pulp.lpSum(x[(M, O, T, S)] * ((trans_I[(M, S)] if self.scale[T] == 'Centralized' else 0.0)            + self.I_T[T]) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list)    + pulp.lpSum(z[(T, S, cp, mark)] * (trans_I_cp[(S, cp, mark)] if self.scale[T] == 'Centralized' else 0.0)        for T in self.idtech for S in self.S_list for cp in self.coprod_list for mark in self.market_list)    + (self.unsort * self.I_inc))

        model.writeLP("biowaste_model.lp")

        
        model.solve()
        
        #stockage des résultats
        self.x_opt = {k: v.varValue for k, v in x.items() if v.varValue is not None}
        self.y_opt = {k: v.varValue for k, v in y.items() if v.varValue is not None}
        self.z_opt = {k: v.varValue for k, v in z.items() if v.varValue is not None}

        
        #print("\n=== Valeurs des variables d’activation y(T,S) ===")
        
        for T in self.idtech:
            for S in self.S_list:
                val = y[(T, S)].varValue
                #print(f"y[{T},{S}] = {val}")

        
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
                                #print(f"  {O} → {T} : {var.varValue:.1f} tonnes")
        
        print("\nFlux de coproduits vers les marchés :")
        for T in self.idtech:
            for cp in self.coprod_list:
                for mark in self.market_list:
                    for S in self.S_list:
                        var = z[(T,S, cp, mark)]
                        if var.varValue is not None and var.varValue > 1e-3:
                            if self.scale[T] == 'Centralized':
                                print(f"  {cp} issu de {T} installé dans {S} → marché {mark} : {var.varValue:.2f}")
                            #else:
                                #print(f"  {cp} issu de {T} → marché {mark} : {var.varValue:.2f}")
        print("\nTechnologies activées et flux associés :")

        for T in self.idtech:
            for S in self.S_list:
                if y[(T, S)].varValue == 1:
                    total_TS = sum(x[(M, O, T, S)].varValue for M in self.mun for O in self.origin                                   if x[(M, O, T, S)].varValue is not None)
                    print(f"  {T} installé en {S} : {total_TS:.1f} tonnes")

        
                        
        
        print(f"\nStatut de l'optimisation : {pulp.LpStatus[model.status]}")
        
        self.tot_imp = pulp.value(model.objective)
        

                                                                    
                                                                      
        print(f"Impact environnemental total minimal : {(self.tot_imp):.2f}")
        
    
    def impact_tot(self):
        
        total = 0
        total1 = 0
        total2 = 0
        
        
    #Calcule les impacts environnementaux pour plusieurs catégories en utilisant la solution optimale déjà trouvée.

        # ---------- Impact traitement + transport des déchets ----------
        for (M, O, T, S), val in self.x_opt.items():
# transport déchets
            if self.scale[T] == 'Centralized':
                trans1 = (self.col_w1 * self.dist_col[M]) + (self.tr_w1 * self.dist[M][S])
                trans2 = (self.col_w2 * self.dist_col[M]) + (self.tr_w2 * self.dist[M][S])
                trans = (self.col_w * self.dist_col[M]) + (self.tr_w * self.dist[M][S])
                
            else:
                trans1 = 0.0
                trans2 = 0
                trans = 0

            # traitement
            total1 += val * (trans1 + self.I_T1[T])
            total2 += val * (trans2 + self.I_T2[T])
            total += val * (trans + self.I_T[T])
    

        # ---------- Transport des coproduits ----------
        for (T, S, cp, mark), val in self.z_opt.items():
            if self.scale[T] == 'Centralized':
                total1 += val * self.I_tr_cop1[cp] * self.dist[mark][S]
                total2 += val * self.I_tr_cop2[cp] * self.dist[mark][S]
                total += val * self.I_tr_cop[cp] * self.dist[mark][S]
            else:
                total1 += 0
                total2 += 0
                total += 0


        # ---------- Incinération des déchets non triés ----------
        total1 += self.unsort * self.I_inc1
        total2 += self.unsort * self.I_inc2
        total += self.unsort * self.I_inc
        

        print(f"Impact total {self.ind1} : {total1:.2f} {self.unit1}")
        print(f"Impact total {self.ind2} : {total2:.2f} {self.unit2}")
        print(f"Impact total {self.ind} : {total:.2f} {self.unit}")
        


# In[922]:


Car_SI = v13_debugg(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Carbon', 0.6, 'Total ecosystem quality', 'Total human health')


# In[924]:


Car_SI.optimize()


# In[925]:


Car_SI.impact_tot()


# In[928]:


EQ_SI = v13_debugg(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total ecosystem quality', 0.6, 'Carbon', 'Total human health')


# In[929]:


EQ_SI.optimize()


# In[931]:


EQ_SI.impact_tot()


# In[932]:


HH_SI = v13_debugg(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Total human health', 0.6, 'Carbon', 'Total ecosystem quality')


# In[933]:


HH_SI.optimize()


# In[934]:


HH_SI.impact_tot()


# In[936]:


Carb_I = v13_debugg(ARC_reel, tech_ARC_reel, impact4, 'Carbon', 0.6, 'Total ecosystem quality', 'Total human health')


# In[937]:


Carb_I.optimize()


# In[938]:


Carb_I.impact_tot()


# In[939]:


EQ_I= v13_debugg(ARC_reel, tech_ARC_reel, impact4, 'Total ecosystem quality', 0.6, 'Carbon', 'Total human health')


# In[940]:


EQ_I.optimize()


# In[941]:


EQ_I.impact_tot()


# In[942]:


HH_I= v13_debugg(ARC_reel, tech_ARC_reel, impact4, 'Total human health', 0.6, 'Carbon', 'Total ecosystem quality')


# In[943]:


HH_I.optimize()


# In[944]:


HH_I.impact_tot()


# In[911]:


ARC_test = 'ARC_test.xlsx'
test_tech = 'test_tech.xlsx'
impact_test = 'impact-test-simple.xlsx'


# In[912]:


Car_debb = v13_debugg(ARC_test, test_tech, impact_test, 'Carbon', 0.8, 'Total ecosystem quality', 'Total human health')


# In[913]:


Car_debb.optimize()


# In[880]:


eq_debugg = v13_debugg(ARC_test, test_tech, impact_test, 'Total ecosystem quality', 0.8, 'Carbon', 'Total human health')


# In[914]:


Car_debb.impact_tot()


# In[881]:


eq_debugg.optimize()


# In[ ]:





# In[858]:


eq_debugg.impact_tot()


# In[859]:


hh_debugg = v13_debugg(ARC_test, test_tech, impact_test, 'Total human health', 0.8, 'Carbon', 'Total ecosystem quality')


# In[860]:


hh_debugg.optimize()


# In[861]:


hh_debugg.impact_tot()


# In[820]:


class DEBUGG:  
    def __init__(self, municipalities, technologies, method, ind, sort):
        
        self.ind = ind
        self.sort = sort
        

        
        with pd.ExcelFile(municipalities) as f:
            
            #list of municipalities
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])
            
            #total mass of waste generated
            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            
            self.home = mass['Home'] *self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total']  * self.sort
            
            self.unsort = self.tot_mass * (1 - self.sort)
            
            
            #mass of waste sorted from individual home and buildings
            #mass = pd.read_excel(municipalities, 'waste_sorted', index_col=0).fillna(0.0)
            #self.home = mass['Home']
            #self.build = mass['Building']
            #self.total = mass['Total']
            
            
            #possibility of a site to host a technology
            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech
            
            #list of potential sites
            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])
            
            #collection distance of each municipality
            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']
            
            
            #distances between municipalities, potential sites and potential markets
            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist
            
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
            #possible market
            poss_mark = pd.read_excel(municipalities, 'market', index_col = 0)
            self.poss_mark = poss_mark

            
            
        with pd.ExcelFile(technologies) as f:
            #list of technologies
            
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])
            
            #scale of technology
            tech_scale = pd.read_excel(technologies, 'id_tech',index_col=0)
            self.scale = tech_scale['Type']
            
            #constraint capacity
            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            self.lb = lb #lower bound constraint
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.ub = ub  #upper bound constraint
            
            
            #list of possible coproduct
            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])
            
            #yield of coproduct production for each technology
            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod
            

            
            
            #list of possible conventional product 
            #c_p_list = pd.read_excel(technologies, 'subs_ratio')
            
            c_p_list = pd.read_excel(technologies, 'avoid')
            self.c_p_list = np.array(c_p_list['Conventional product'])
            
            #substitution ratio conventional product/co-product
            #sub_ratio = pd.read_excel(technologies, 'subs_ratio', index_col=0)
            #self.sub_ratio = sub_ratio
            
            #avoided product for each technology
            avoid = pd.read_excel(technologies, 'avoid',index_col=0)
            self.avoid = avoid

                
        
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')
            
            #unit
            self.unit = imp[self.ind][0]
            
            #impacts of processes
            ##impact of transport (after collection, from the municipality to the treatment site)
            self.tr_w = imp[self.ind][1]
            ##impact of collection
            self.col_w = imp[self.ind][2]
            
            ##impact of other processes
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())  # les titres ont Unit reference vide
            imp["Step"] = imp["Step"].ffill()
            
            ###treatment technology
            tech = imp[imp["Step"] == "Treatment technologies"]
            tech = tech.set_index('Process')
            self.I_T = tech[self.ind]
            self.I_inc = self.I_T['Inc']
            
            ###Impact coproduct transport
            tr_cop = imp[imp["Step"] == "Transport of coproduct"]
            tr_cop = tr_cop.set_index('Process')
            self.I_tr_cop = tr_cop[self.ind]
            
            ###Avoided impact related to conventional product production
            av = imp[imp["Step"] == "Conventional products"]
            av = av.set_index('Process')
            self.I_av = av[self.ind]
            
            ###Net impact of coproduct use
            u_cop = imp[imp["Step"] == 'Net impact of coproduct use']
            u_cop = u_cop.set_index('Process')
            self.I_u_cop = u_cop[self.ind]
            
            
    def optimize(self):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)
        self.origin = ['Home', 'Building']

    # Variables
        x = {(M,O,T,S): pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
             for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list}

        y = {(T,S): pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
             for T in self.idtech for S in self.S_list}

        z = {(T,S,cp,mark): pulp.LpVariable(f"z_{T}_{S}_{cp}_{mark}", lowBound=0)
             for T in self.idtech for S in self.S_list for cp in self.coprod_list for mark in self.market_list}

    # Contraintes : tout le trié doit être traité
        for M in self.mun:
            model += pulp.lpSum(x[(M,'Home',T,S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M,'Building',T,S)] for T in self.idtech for S in self.S_list) == self.build[M]

    # HC interdit en immeubles
        for M in self.mun:
            if 'HC' in self.idtech:
                model += pulp.lpSum(x[(M,'Building','HC',S)] for S in self.S_list) == 0

    # Site impossible -> y=0 et flux=0
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T,S)] == 0
                    model += pulp.lpSum(x[(M,O,T,S)] for M in self.mun for O in self.origin) == 0

    # Capacités : lier flux et activation
    # IMPORTANT : ceci suffit, pas besoin d’autres contraintes sur y ici
        for T in self.idtech:
            for S in self.S_list:
                tot_TS = pulp.lpSum(x[(M,O,T,S)] for M in self.mun for O in self.origin)

                if not pd.isna(self.lb[T][S]):
                    model += tot_TS >= self.lb[T][S] * y[(T,S)]
                if not pd.isna(self.ub[T][S]):
                    model += tot_TS <= self.ub[T][S] * y[(T,S)]

    # Décentralisé : ne reçoit que de son site
    # ⚠️ Correction : on NE TOUCHE PAS à y(T,S) ici !
        for T in self.idtech:
            if self.scale[T] == 'Decentralized':
                for S in self.S_list:
                    for M in self.mun:
                        if M != S:
                            model += pulp.lpSum(x[(M,O,T,S)] for O in self.origin) == 0

    # Coproduits : bilan matière
        for T in self.idtech:
            for S in self.S_list:
                for cp in self.coprod_list:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for mark in self.market_list) ==                              pulp.lpSum(x[(M,O,T,S)] * self.coprod[T][cp] for M in self.mun for O in self.origin)

    # Marchés impossibles
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T,S,cp,mark)] for T in self.idtech for S in self.S_list) == 0

    # OBJECTIF : utiliser UNE SEULE expression d'impact (même formule partout)
        model += self.impact_total(self.ind, x, z)

        model.solve()

    # Stockage solution
        self.x_opt = {k: v.varValue for k, v in x.items() if v.varValue is not None}
        self.y_opt = {k: v.varValue for k, v in y.items() if v.varValue is not None}
        self.z_opt = {k: v.varValue for k, v in z.items() if v.varValue is not None}

    # Score "solveur" (doit maintenant coller au recalcul)
        self.tot_imp = pulp.value(model.objective)
        print(f"Statut : {pulp.LpStatus[model.status]}")
        print(f"Impact optimisé (objectif) : {self.tot_imp:.6g} {self.unit}")


# In[822]:


debugg = DEBUGG(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Carbon',0.8)


# In[823]:


debugg.optimize()


# In[824]:


import numpy as np
import pandas as pd
import pulp


class v13_clean:
    def __init__(self, municipalities, technologies, method, ind, sort, ind1=None, ind2=None):
        self.ind = ind
        self.sort = sort
        self.ind1 = ind1
        self.ind2 = ind2

        # -------------------------
        # Read municipalities file
        # -------------------------
        with pd.ExcelFile(municipalities) as f:
            mun = pd.read_excel(municipalities, 'mun')
            self.mun = np.array(mun['municipalities'])

            mass = pd.read_excel(municipalities, 'waste_gen', index_col=0).fillna(0.0)
            self.tot_mass = mass['Total']['Total']
            self.home = mass['Home'] * self.sort
            self.build = mass['Building'] * self.sort
            self.total = mass['Total'] * self.sort
            self.unsort = self.tot_mass * (1 - self.sort)

            poss_tech = pd.read_excel(municipalities, 'tech_pot', index_col=0).fillna(0.0)
            self.poss_tech = poss_tech

            S_list = pd.read_excel(municipalities, 'tech_pot')
            self.S_list = np.array(S_list['Site'])

            dist_col = pd.read_excel(municipalities, 'dist_col', index_col=0).fillna(0.0)
            self.dist_col = dist_col['dist_col']

            dist = pd.read_excel(municipalities, 'dist_market', index_col=0).fillna(0.0)
            self.dist = dist

            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])

            poss_mark = pd.read_excel(municipalities, 'market', index_col=0)
            self.poss_mark = poss_mark

        # -------------------------
        # Read technologies file
        # -------------------------
        with pd.ExcelFile(technologies) as f:
            type_tech = pd.read_excel(technologies, 'id_tech')
            self.idtech = np.array(type_tech['Symbol'])

            tech_scale = pd.read_excel(technologies, 'id_tech', index_col=0)
            self.scale = tech_scale['Type']  # 'Centralized' or 'Decentralized'

            lb = pd.read_excel(technologies, 'tech_lb', index_col=0)
            ub = pd.read_excel(technologies, 'tech_ub', index_col=0)
            self.lb = lb
            self.ub = ub

            coprod_list = pd.read_excel(technologies, 'coproduct')
            self.coprod_list = np.array(coprod_list['Co-product'])

            coprod = pd.read_excel(technologies, 'coproduct', index_col=0)
            self.coprod = coprod  # yields

            avoid = pd.read_excel(technologies, 'avoid', index_col=0)
            self.avoid = avoid  # mapping tech->coprod->conventional product

        # -------------------------
        # Read LCIA method file
        # -------------------------
        with pd.ExcelFile(method) as f:
            imp = pd.read_excel(method, 'Feuil1')

            # Helper to access all indicators
            self._imp_table = imp

            # Prepare step column
            imp["Step"] = imp['Process'].where(imp["Unit reference"].isna())
            imp["Step"] = imp["Step"].ffill()

            # Units are stored in first row for each indicator (as in your file)
            self.units = {col: imp[col][0] for col in imp.columns if col not in ["Process", "Unit reference", "Step"]}

            # transport waste + collection (by indicator)
            self.tr_w = {col: imp[col][1] for col in self.units.keys()}
            self.col_w = {col: imp[col][2] for col in self.units.keys()}

            # Treatment technologies impacts (rows indexed by tech symbol e.g., 'AD', 'Inc', etc.)
            tech = imp[imp["Step"] == "Treatment technologies"].set_index('Process')
            self.I_T_all = tech[self.units.keys()]  # DataFrame: index=Process, columns=indicators

            # Coproduct transport impacts (rows indexed by cp)
            tr_cop = imp[imp["Step"] == "Transport of coproduct"].set_index('Process')
            self.I_tr_cop_all = tr_cop[self.units.keys()]  # DataFrame

            # Conventional products (if you later want avoided burdens)
            av = imp[imp["Step"] == "Conventional products"].set_index('Process')
            self.I_av_all = av[self.units.keys()]

            # Net impact of coproduct use (if you later want to include)
            u_cop = imp[imp["Step"] == "Net impact of coproduct use"].set_index('Process')
            self.I_u_cop_all = u_cop[self.units.keys()]

        # Common
        self.origin = ['Home', 'Building']

        # Placeholders after solve
        self.model = None
        self.x = None
        self.y = None
        self.z = None
        self.x_opt = None
        self.y_opt = None
        self.z_opt = None

    # -------------------------
    # Impact engine (ONE source of truth)
    # -------------------------
    def impact_expr(self, ind, x, z):
        """
        Returns a PuLP expression if x/z are PuLP variables (dicts),
        or a numeric total if x/z are dicts of floats.
        Must be identical for optimization and post-evaluation.
        """
        # precompute transport waste (M,S)
        trans_MS = {(M, S): (self.col_w[ind] * self.dist_col[M]) + (self.tr_w[ind] * self.dist[M][S])
                    for M in self.mun for S in self.S_list}

        # precompute transport coproduct (cp,S,mark)
        tr_cp = {(cp, S, mark): self.I_tr_cop_all.loc[cp, ind] * self.dist[mark][S]
                 for cp in self.coprod_list for S in self.S_list for mark in self.market_list}

        total = 0

        # Waste: treatment + transport (transport = 0 for decentralized)
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        if self.scale[T] == 'Centralized':
                            total += x[(M, O, T, S)] * (trans_MS[(M, S)] + self.I_T_all.loc[T, ind])
                        else:
                            total += x[(M, O, T, S)] * (0.0 + self.I_T_all.loc[T, ind])

        # Coproduct transport (0 for decentralized)
        for T in self.idtech:
            for S in self.S_list:
                for cp in self.coprod_list:
                    for mark in self.market_list:
                        if self.scale[T] == 'Centralized':
                            total += z[(T, S, cp, mark)] * tr_cp[(cp, S, mark)]
                        else:
                            total += 0

        # Unsorted waste incineration (constant; include for consistent reporting)
        total += self.unsort * self.I_T_all.loc['Inc', ind]

        return total

    def impact_value(self, ind, x_opt=None, z_opt=None):
        """Numeric impact using stored solution (or provided dicts)."""
        if x_opt is None:
            x_opt = self.x_opt
        if z_opt is None:
            z_opt = self.z_opt

        # Build numeric dicts with full keys if missing (safer)
        # Here we assume x_opt / z_opt have all relevant keys; we just sum existing ones.
        total = 0.0

        # precompute transport waste (M,S)
        trans_MS = {(M, S): (self.col_w[ind] * self.dist_col[M]) + (self.tr_w[ind] * self.dist[M][S])
                    for M in self.mun for S in self.S_list}

        for (M, O, T, S), val in x_opt.items():
            if val is None or abs(val) < 1e-12:
                continue
            if self.scale[T] == 'Centralized':
                total += val * (trans_MS[(M, S)] + float(self.I_T_all.loc[T, ind]))
            else:
                total += val * (0.0 + float(self.I_T_all.loc[T, ind]))

        for (T, S, cp, mark), val in z_opt.items():
            if val is None or abs(val) < 1e-12:
                continue
            if self.scale[T] == 'Centralized':
                total += val * float(self.I_tr_cop_all.loc[cp, ind]) * float(self.dist[mark][S])

        total += float(self.unsort) * float(self.I_T_all.loc['Inc', ind])
        return total

    # -------------------------
    # Optimization
    # -------------------------
    def optimize(self, solver=None):
        model = pulp.LpProblem("biowaste_management", pulp.LpMinimize)

        # Variables
        x = {(M, O, T, S): pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
             for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list}

        y = {(T, S): pulp.LpVariable(f"y_{T}_{S}", cat="Binary")
             for T in self.idtech for S in self.S_list}

        z = {(T, S, cp, mark): pulp.LpVariable(f"z_{T}_{S}_{cp}_{mark}", lowBound=0)
             for T in self.idtech for S in self.S_list for cp in self.coprod_list for mark in self.market_list}

        # Constraints: all sorted waste treated
        for M in self.mun:
            model += pulp.lpSum(x[(M, 'Home', T, S)] for T in self.idtech for S in self.S_list) == self.home[M]
            model += pulp.lpSum(x[(M, 'Building', T, S)] for T in self.idtech for S in self.S_list) == self.build[M]

        # HC cannot be used in buildings (if exists)
        if 'HC' in list(self.idtech):
            for M in self.mun:
                model += pulp.lpSum(x[(M, 'Building', 'HC', S)] for S in self.S_list) == 0

        # Technology can be installed only where possible
        for T in self.idtech:
            for S in self.S_list:
                if self.poss_tech[T][S] == 0:
                    model += y[(T, S)] == 0
                    model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin) == 0

        # Capacity constraints linked to activation
        for T in self.idtech:
            for S in self.S_list:
                tot_TS = pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin)

                if not pd.isna(self.lb[T][S]):
                    model += tot_TS >= self.lb[T][S] * y[(T, S)]
                if not pd.isna(self.ub[T][S]):
                    model += tot_TS <= self.ub[T][S] * y[(T, S)]

        # Decentralized: only waste from its own site
        # IMPORTANT: do NOT constrain y here.
        for T in self.idtech:
            if self.scale[T] == 'Decentralized':
                for S in self.S_list:
                    for M in self.mun:
                        if M != S:
                            model += pulp.lpSum(x[(M, O, T, S)] for O in self.origin) == 0

        # Coproduct balance
        for T in self.idtech:
            for S in self.S_list:
                for cp in self.coprod_list:
                    model += pulp.lpSum(z[(T, S, cp, mark)] for mark in self.market_list) ==                              pulp.lpSum(x[(M, O, T, S)] * self.coprod.loc[T, cp] for M in self.mun for O in self.origin)

        # Market feasibility
        for cp in self.coprod_list:
            for mark in self.market_list:
                if self.poss_mark[cp][mark] == 0:
                    model += pulp.lpSum(z[(T, S, cp, mark)] for T in self.idtech for S in self.S_list) == 0

        # Objective: EXACT same expression as post-evaluation
        model += self.impact_expr(self.ind, x, z)

        # Solve
        if solver is None:
            model.solve()
        else:
            model.solve(solver)

        # Save model + vars
        self.model = model
        self.x = x
        self.y = y
        self.z = z

        # Store solution
        self.x_opt = {k: v.varValue for k, v in x.items() if v.varValue is not None}
        self.y_opt = {k: v.varValue for k, v in y.items() if v.varValue is not None}
        self.z_opt = {k: v.varValue for k, v in z.items() if v.varValue is not None}

        print(f"Statut : {pulp.LpStatus[model.status]}")

        # Objective and recomputed check (should match)
        obj = pulp.value(model.objective)
        chk = self.impact_value(self.ind)
        print(f"Objectif (solveur) : {obj:.6g} {self.units.get(self.ind,'')}")
        print(f"Recalcul (post)    : {chk:.6g} {self.units.get(self.ind,'')}")

        return obj

    # -------------------------
    # Multi-impact evaluation (same engine)
    # -------------------------
    def evaluate_impacts(self, indicators):
        out = {}
        for ind in indicators:
            out[ind] = self.impact_value(ind)
        return pd.Series(out)

    # -------------------------
    # Store flows per commune
    # -------------------------
    def flows_by_commune(self, tol=1e-6):
        """
        Returns a DataFrame with flows per commune (M), origin (O), technology (T), site (S), amount.
        """
        rows = []
        for (M, O, T, S), v in self.x_opt.items():
            if v is None or v <= tol:
                continue
            rows.append({"Commune": M, "Origin": O, "Tech": T, "Site": S, "Flow": v})
        return pd.DataFrame(rows)


# In[828]:


test_v13_clean =v13_clean(ARC_reel_ss_inv, tech_reel_ss_inv, impact4, 'Carbon', 0.8, 'Total ecosystem quality', 'Total human health')


# In[829]:


test_v13_clean.optimize()


# In[ ]:




