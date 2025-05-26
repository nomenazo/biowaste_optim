#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pulp

# Création du problème d'optimisation
prob = pulp.LpProblem("Optimisation_Gestion_Biodechets", pulp.LpMinimize)

# Variables de décision : quantités de biodéchets envoyées vers chaque traitement (en tonnes)
#x_HC = pulp.LpVariable("x_HC", lowBound=0)  # Compostage domestique
#x_CC = pulp.LpVariable("x_CC", lowBound=0) # Compostage partagé
x_IC = pulp.LpVariable("x_IC", lowBound=0) #compostage industriel
x_M_inj = pulp.LpVariable("x_M_inj", lowBound=0) #methanisation avec injection
x_M_cog = pulp.LpVariable("x_M_cog", lowBound=0)  # Méthanisation avec cogeneration

# Paramètres (exemple de valeurs, à adapter selon les données réelles)
Q_total = 18e6   # Quantité totale de biodéchets (t/an)
Comp_max = 1080e6      # Capacité maximale du marché pour le compost (kWh/an)
chal_max =   3e9  # Capacité maximale du marché pour la chaleur (kWh/an)
elec_max = 6e9 # Capacité max du marché d'élec (kWh/an)
M_capacite = 7000 # Capacité maximale d'infrastructure pour la méthanisation (t/an)

# Rendements des coproduits (exemple)
r_C = 0.5   # Rendement en compost (t compost/t biodéchets)
r_biometh = 80   # Rendement en biométhane (kWh biométhane/t biodéchets)
r_chal = 410 #rendement en chaleur (kWh/t biodéchets)
r_elec = 290 #rendement en électricité (kWh/t biodéchets)


# Coefficients d'émission de GES (exemple en kg CO2e/t biodéchets)
#GES_HC = 0.109e3  # Émissions pour HC (kg CO2e/t)
GES_IC = 0.16e3 # 0.147e3  # Émissions pour la méthanisation (kg CO2e/t)
GES_M_inj = 0.187e3
GES_M_cog = 0.177e3

# Fonction objectif : minimisation des émissions de GES
prob +=  GES_IC * x_IC + GES_M_inj * x_M_inj + GES_M_cog * x_M_cog, "Minimisation_GES"

# Contraintes
prob += x_IC+ x_M_inj + x_M_cog == Q_total, "Biodéchets_disponibles"
prob += x_IC * r_C <= Comp_max, "Marché_compost"
prob += x_M_inj * r_biometh + x_M_cog * r_chal <= chal_max, "Marché_chaleur"
prob +=  x_M_cog * r_elec <= elec_max, "Marché_elec"
#prob += x_C <= C_capacite, "Capacité_compost"
#prob += x_M <= M_capacite, "Capacité_méthanisation"

# Résolution du problème
prob.solve()

# Affichage des résultats
print("Statut:", pulp.LpStatus[prob.status])
#print("Quantité envoyée vers le compostage domestique:", x_HC.varValue, "tonnes")
print("Quantité envoyée vers le compostage industriel:", x_IC.varValue, "tonnes")
print("Quantité envoyée vers la methanisation suivie d'injection:", x_M_inj.varValue, "tonnes")
print("Quantité envoyée vers la méthanisation suivie de cogénération:", x_M_cog.varValue, "tonnes")
print("Émissions minimales de GES:", pulp.value(prob.objective), "kg CO2e")


# In[2]:


from scipy.optimize import linprog
import numpy as np

# Paramètres
Q_total = 10000   # Quantité totale de biodéchets (t/an)
C_max = 4000      # Capacité maximale du marché pour le compost (t/an)
B_max = 5000      # Capacité maximale du marché pour le biométhane (t/an)
C_capacite = 6000 # Capacité maximale d'infrastructure pour le compost (t/an)
M_capacite = 7000 # Capacité maximale d'infrastructure pour la méthanisation (t/an)

# Rendements des coproduits (exemple)
r_C = 0.5   # Rendement en compost (t compost/t biodéchets)
r_B = 0.3   # Rendement en biométhane (t biométhane/t biodéchets)

# Coefficients d'émission de GES (kg CO2e/t biodéchets)
GES_C = 50  # Émissions pour le compostage
GES_M = 30  # Émissions pour la méthanisation

# Fonction objectif (minimiser GES)
c = np.array([GES_C, GES_M])

# Matrice des contraintes (Ax = b)
A_eq = np.array([[1, 1]])  # x_C + x_M = Q_total
b_eq = np.array([Q_total])

# Contraintes inégalités (Ax <= b)
A_ub = np.array([
    [r_C, 0],   # x_C * r_C <= C_max (marché compost)
    [0, r_B],   # x_M * r_B <= B_max (marché biométhane)
    [1, 0],     # x_C <= C_capacite
    [0, 1]      # x_M <= M_capacite
])
b_ub = np.array([C_max, B_max, C_capacite, M_capacite])

# Bornes des variables (x_C, x_M >= 0)
bounds = [(0, None), (0, None)]

# Résolution du problème
res = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method='highs')

# Affichage des résultats
if res.success:
    print("Solution trouvée:")
    print("Quantité envoyée vers le compostage:", res.x[0], "tonnes")
    print("Quantité envoyée vers la méthanisation:", res.x[1], "tonnes")
    print("Émissions minimales de GES:", res.fun, "kg CO2e")
else:
    print("Aucune solution optimale trouvée.")


# In[3]:


import pulp

# Modèle de minimisation
model = pulp.LpProblem("Optimisation_Biodechets_GES", pulp.LpMinimize)

# Quantité totale de biodéchets à traiter
total_biodechets = 300

# Capacité min et max (en tonnes)
capacities = {
    "MethC": (50, 200),
    "MethD": (10, 100),
    "CompostInd": (30, 150),
    "CompostIndiv": (5, 80)
}

# Facteurs d’émission GES (kg CO2e/tonne)
emissions = {
    "MethC": 50,
    "MethD": 40,
    "CompostInd": 100,
    "CompostIndiv": 20
}

# Création des variables de flux (x) et d’activation (y)
x = {}  # quantités
y = {}  # activation binaire

for tech in capacities:
    Cmin, Cmax = capacities[tech]
    x[tech] = pulp.LpVariable(f"x_{tech}", lowBound=0)
    y[tech] = pulp.LpVariable(f"y_{tech}", cat="Binary")
    # Contraintes de min et max conditionnées à l’activation
    model += x[tech] >= Cmin * y[tech], f"Min_{tech}"
    model += x[tech] <= Cmax * y[tech], f"Max_{tech}"

# Contrainte : somme des flux = total à traiter
model += pulp.lpSum(x[tech] for tech in x) == total_biodechets, "Total_biodechets"

# Objectif : minimiser les GES
model += pulp.lpSum(emissions[tech] * x[tech] for tech in x), "Min_GES"

# Résolution
model.solve()

# Affichage des résultats
print("Résultat :")
for tech in x:
    print(f"{tech} : {x[tech].varValue:.1f} tonnes (activé : {int(y[tech].varValue)})")
print(f"Total GES : {pulp.value(model.objective):.1f} kg CO₂e")


# In[4]:


import pulp

# Création du modèle
model = pulp.LpProblem("Multi_installations_MethC", pulp.LpMinimize)

# Demande totale à traiter pour cette technologie (en tonnes)
demand = 300

# Nombre maximum de sites potentiels
N_sites = 5

# Capacités minimale et maximale pour chaque installation
Cmin = 50   # en tonnes
Cmax = 200  # en tonnes

# Variables
x = {}  # Quantité traitée par site i
y = {}  # Variable binaire d'ouverture pour le site i

for i in range(1, N_sites+1):
    x[i] = pulp.LpVariable(f"x_site_{i}", lowBound=0)
    y[i] = pulp.LpVariable(f"y_site_{i}", cat="Binary")
    
    # Contrainte de capacité minimum si le site est ouvert
    model += x[i] >= Cmin * y[i], f"Min_site_{i}"
    # Contrainte de capacité maximum
    model += x[i] <= Cmax * y[i], f"Max_site_{i}"

# Contrainte : satisfaire la demande totale
model += pulp.lpSum(x[i] for i in range(1, N_sites+1)) == demand, "Total_demande"

# Exemple d'objectif : minimiser les coûts totaux (ou émissions, voir la suite)
model += pulp.lpSum(x[i] for i in range(1, N_sites+1)), "Objectif_totale"

# Résolution
model.solve()

# Affichage des résultats
for i in range(1, N_sites+1):
    print(f"Site {i} : {x[i].varValue:.1f} tonnes, ouvert ? {int(y[i].varValue)}")


# In[5]:


import pulp

# Création du modèle
model = pulp.LpProblem("Multi_unités", pulp.LpMinimize)

# Demande totale à traiter
demand = 300

# Capacités par unité
Cmin = 50   # en tonnes par unité
Cmax = 200  # en tonnes par unité

# Variables
x_tot = pulp.LpVariable("x_total", lowBound=0)    # Quantité totale assignée à la technologie
z = pulp.LpVariable("z_units", lowBound=0, cat="Integer")  # Nombre d'unités ouvertes

# Contraintes pour assurer que la quantité traitée correspond au nombre d'unités déployées
model += x_tot >= Cmin * z, "Min_global"
model += x_tot <= Cmax * z, "Max_global"

# Contrainte sur la demande totale
model += x_tot == demand, "Demande_totale"

# Exemple d'objectif : minimiser x_total (ou un autre objectif pertinent)
model += x_tot, "Objectif_totale"

# Résolution
model.solve()

print("Quantité totale traitée : {:.1f}".format(x_tot.varValue))
print("Nombre d'unités ouvertes :", int(z.varValue))


# In[238]:


import pulp

# Technologies
technos = ["CI", "CP", "CC", "MC"]
techs_contraintes = ["CP", "CC", "MC"]

# Données
D_M = 1000  # tonnes depuis maisons
D_I = 2000  # tonnes depuis immeubles

emin = {"CP": 100, "CC": 150, "MC": 200}
emax = {"CP": 500, "CC": 800, "MC": 1200}

emissions = {"CI": 5, "CP": 90, "CC": 80, "MC": 25}  # kg CO2e/tonne

# Modèle
model = pulp.LpProblem("Optimisation_Biodechets_Commune", pulp.LpMinimize)

# Variables
x = {}
for tech in technos:
    for origine in ["M", "I"]:
        if tech == "CI" and origine == "I":
            continue  # immeubles ne peuvent pas faire du compostage individuel
        x[(origine, tech)] = pulp.LpVariable(f"x_{origine}_{tech}", lowBound=0)
# Binaires pour activation
y = {tech: pulp.LpVariable(f"y_{tech}", cat="Binary") for tech in techs_contraintes}

# Contraintes : tous les déchets doivent être traités
model += pulp.lpSum(x[("M", tech)] for tech in technos) == D_M, "Traitement_maisons"
model += pulp.lpSum(x[("I", tech)] for tech in techs_contraintes) == D_I, "Traitement_immeubles"

# Contraintes de capacité pour CP, CC, MC
for tech in techs_contraintes:
    model += (x[("M", tech)] + x[("I", tech)] >= emin[tech] * y[tech], f"Min_{tech}")
    model += (x[("M", tech)] + x[("I", tech)] <= emax[tech] * y[tech], f"Max_{tech}")

# Objectif : minimisation des GES
model += pulp.lpSum(
    emissions[tech] * (x[("M", tech)] + x[("I", tech)] if tech != "CI" else x[("M", tech)])
    for tech in technos
), "Min_GES"

# Résolution
model.solve()

# Résultats
for (origine, tech), var in x.items():
    print(f"{origine} {tech} : {var.varValue:.1f} tonnes")


# In[239]:


import numpy as np
import pandas as pd


# In[875]:


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


            
            
    


# In[876]:


data_test = 'territory.xlsx'
data_test_tech = 'technology.xlsx'


# In[877]:


test = territory(data_test,data_test_tech)


# In[795]:


test.scale['HC']


# In[796]:


test.site


# In[797]:


test.site['AD-b']['A']


# In[798]:


test.S_list


# In[799]:


np.array(test.site.columns)


# In[878]:


test.optimize()


# In[264]:


test2 = territory2(data_test,data_test_tech)


# In[ ]:


type_tech = pd.read_excel(technologies, 'id_tech')


# In[1185]:


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


# In[1186]:


testv1 = v1(data_test,data_test_tech)


# In[1187]:


testv1.optimize()


# In[812]:


A= np.array(testv1.mass.columns)
A


# In[1188]:


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


# In[1189]:


testv2 = v2(data_test,data_test_tech)


# In[1190]:


testv2.optimize()


# In[1191]:


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


# In[1192]:


testv3 = v3(data_test,data_test_tech)


# In[1193]:


testv3.optimize()


# In[1194]:


#v3 : pas de solution, car les contraintes sont trop rigides (contrainte minimal qui exige l'envoi de déchets dans chaque technologie)


# In[1195]:


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


# In[1196]:


testv4 = v4(data_test,data_test_tech)


# In[1197]:


testv4.optimize()


# In[1213]:


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


# In[1214]:


testv5 = v5(data_test,data_test_tech)


# In[1215]:


testv5.lieu['C']


# In[1216]:


testv5.optimize()


# In[1217]:


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
        
        self.origin = ['Home','Building']
        
        #variables
        x = {} #quantity of waste sent to each technology
        y = {} #activation or not of a technology
        
        for M in self.mun:
            for O in self.origin:
                for T in self.idtech:
                    for S in self.S_list:
                        x[(M, O, T, S)] = pulp.LpVariable(f"x_{M}_{O}_{T}_{S}", lowBound=0)
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
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
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list ) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
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
        
        
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] +self.GHG[T] ) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list) # if not (T == "HC" and origin == "building") ) # HC interdit pour immeubles, "Total_Environmental_Impact"
        
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


# In[1218]:


testv6 = v6(data_test,data_test_tech)


# In[1219]:


testv6.optimize()


# In[1245]:


class v7: ##add the subsitution part
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
            
            #possibility of location of technology
            site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            self.site = site
            
            
            
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
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
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
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list ) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
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
        sub_GHG = {}
        for T in self.idtech:
            for cop in self.coprod_list:
                for conv_p in self.c_p_list:
                    sub_GHG[(T,cop,conv_p)] = self.coprod[T][cop] * self.sub_ratio[cop][conv_p] * self.GHG_cp[conv_p] 
        
        print (sub_GHG)               

        #objective
        
        
        model += pulp.lpSum(x[(M, O, T,S)] * ( trans_GHG[(M,T,S)] + self.GHG[T] - sub_GHG[(T,cop,conv_p)]) for M in self.mun for O in self.origin for T in self.idtech for S in self.S_list for cop in self.coprod_list for conv_p in self.c_p_list)
       
        
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


# In[1246]:


testv7 = v7(data_test,data_test_tech)


# In[1255]:


testv7.optimize()


# In[1256]:


tech2 = 'technology2.xlsx'


# In[1257]:


testv7_2 = v7(data_test,tech2)


# In[1258]:


testv7_2.optimize()


# In[1259]:


##v7 boucle qui se répète et donc l'impact minimal est très grand? V7 = FAUX


# In[1564]:


class v8: ##correction of v7 : add variable : type of coproduct
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
            
            #possibility of location of technology
            site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            self.site = site
            
            
            
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
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
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
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list ) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
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


# In[1565]:


testv8 = v8(data_test,tech2)


# In[1572]:


testv8.optimize()


# In[1573]:


testv8_2 = v8(data_test,data_test_tech)


# In[1574]:


testv8_2.optimize()


# In[1384]:


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
            
            #possibility of location of technology
            site = pd.read_excel(municipalities, 'site', index_col=0).fillna(0.0)
            self.site = site
            
            
            
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
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech} 

        
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
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list ) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
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
        
        


# In[1385]:


testv9 = v9(data_test,tech2)


# In[1386]:


testv9.optimize()


# In[1600]:


class v10: ##add variable : market for coproduct
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
            
            #list of coproduct market
            market = pd.read_excel(municipalities, 'market')
            self.market_list = np.array(market['market'])
            
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
        
        y = {T: pulp.LpVariable(f"y_{T}", cat="Binary") for T in self.idtech}
        
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
            if not pd.isna(self.lb[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list ) >= self.lb[T]*y[T], f"Min_{T}"
            if not pd.isna(self.ub[T]):
                model += pulp.lpSum(x[(M, O, T, S)] for M in self.mun for O in self.origin for S in self.S_list) <= self.ub[T]*y[T], f"Max_{T}"
        
        self.transport = 2e-1 #kgCO2/t.km
        
        
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
        
        print (self.trans_cop_GHG) 

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


# In[1601]:


testv10 = v10(data_test,tech2)


# In[1602]:


testv10.optimize()


# In[1597]:


testv10.market_list


# In[1598]:


testv10.dist_mark


# In[1599]:


testv10.tr_cop_GHG


# In[ ]:




