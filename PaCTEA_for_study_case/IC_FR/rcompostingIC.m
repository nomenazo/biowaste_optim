clear;

%initialization of variables
% -> IMPORT INITIAL VALUES FROM "waste characterization.xlsx"

y0=[0.208470299 0.064311765 0.039781924 0.01422 0.035865102 0.031701992 0.013850399 0 0 0 0 0 5e-4 ...
5e-4 1e-4 1e-4 1e-6 1e-6 0 0 0.6397 298 0 0 0 1e-4 0 0 0 0 0.00021 0 0.003 5e-2 0 0 0];

%Assessment of 90 days of composting
tspan(1)=0;
for i=2:90;
    tspan(i)=tspan(i-1)+24;
end;

options = odeset( 'RelTol',1e-12,'AbsTol',1e-14);

[t,y]=ode15s('compostingIC',[tspan],[y0], options);



C=y(:,1);
P=y(:,2);
L=y(:,3);
H=y(:,4);
CE=y(:,5);
LG=y(:,6);
Xi=y(:,7);
Sc=y(:,8);
Sp=y(:,9);
Sl=y(:,10);
Sh=y(:,11);
Slg=y(:,12);
Xmb=y(:,13);
Xtb=y(:,14);
Xma=y(:,15);
Xta=y(:,16);
Xmf=y(:,17);
Xtf=y(:,18);
Xdb=y(:,19);
CO2=y(:,20);
W=y(:,21);
T=y(:,22);
CH4gen=y(:,23);
CH4oxi=y(:,24);
CH4=y(:,25);
Xa=y(:,26);
NO3=y(:,27);
N2O=y(:,28);
N2=y(:,29);
NH3=y(:,30);
NH4=y(:,31);
Wvap = y(:,32);
O2diss = y(:,33);
O2gaz = y(:,34);
O2cons = y(:,35);
O2out = y(:,36);
NH3gaz = y(:,37);

%y = max(0,y);
CH4 = max(0,CH4);
CH4gen = max(0,CH4gen);
CH4oxi = max(0,CH4oxi);

NH4 = max(0,NH4);

CCO2= (0.012/0.044) * CO2;
OCO2 = (0.032/0.044) * CO2;

CCH4 = (0.012/0.016)*CH4;

CCH4gen = (0.012/0.016)*CH4gen;

HCH4gen = (0.004/0.016)*CH4gen;

HH20comp = (0.002/0.018)*W;
OH20comp = (0.016/0.018)*W;

HH20vap = (0.002/0.018)*Wvap;
OH20vap = (0.016/0.018)*Wvap;

NNH3 = (0.014/0.017)*NH3;
HNH3 = (0.003/0.017)*NH3;

HNH4 = (0.004/0.018)*NH4;

NNH3gaz=(0.014/0.017)*NH3gaz;
HNH3gaz = (0.003/0.017)*NH3gaz;


NN2 = N2;
NNH4 = (0.014/0.018)*NH4; %kg/kgTM
NNO3 = (0.014/0.062)*NO3; 

NN2O = (0.028/0.044)*N2O;
ON2O = (0.016/0.044)*N2O;

OO2diss = O2diss;
OO2gas = O2gaz;

%C and N balances
Cinit = (C(1)*(6*0.012/0.180)) + (P(1)*(16*0.012/0.352)) + (L(1)*(25*0.012/0.393))+...
    (H(1)*(10*0.012/0.282))+(CE(1)*(6*0.012/0.180))+(LG(1)*(20*0.012/0.366)) +(Xmb(1)*(5*0.012/0.113))+(Xtb(1)*(5*0.012/0.113))+(Xma(1)*(5*0.012/0.113))+...
    (Xta(1)*(5*0.012/0.113))+(Xmf(1)*(10*0.012/0.247))+(Xtf(1)*(10*0.012/0.247))+(Xa(1)*(5*0.012/113)); %kg/kgTM

Ninit = (P(1)*(4*0.014/0.352)) + ((Xmb(1)*(1*0.014/0.113))+(Xtb(1)*(1*0.014/0.113))+(Xma(1)*(1*0.014/0.113))+...
    (Xta(1)*(1*0.014/0.113))+(Xmf(1)*(1*0.014/0.247))+(Xtf(1)*(1*0.014/0.247))+(Xa(1)*(1*0.014/0.113))); %kg/kgTM
  


Ccompost= Cinit - CCO2 - cumsum(CCH4); %kg/kgTM %Carbon in compost

Cemit = (CCO2+cumsum(CCH4))/Cinit; %cumsum(CCH4)
Cemitcum = Cemit(end);


Ncompost = Ninit - NNH3- NN2 - NN2O-NNO3-NNH3gaz; %kg/kgTM %Nitrogen in compost

Nemit = (NNH3+NN2+NN2O+NNH3gaz)/Ninit;
Nmin = (NNO3+NNH4)/Ninit;
Nemitcum = Nemit(end);
Nmincum = Nmin(end);


CNratio = Ccompost/Ncompost;
%y(y<0)=0;


%direct emissions
%per kgTM (biowaste+wood chips)
CH4tot= cumsum(CH4);
CH4final = CH4tot(end);
CO2final= CO2(end);
NH3final = NH3(end);
N2Ofinal=N2O(end);
N2final= N2(end);
Wevapfinal = Wvap(end);



%Compost mass and composition (per kgTM) 
Mcompost = (C(end) +P(end)+ L(end)+ H(end)+LG(end)+CE(end)+Sc(end)+Sp(end)+...
    Sl(end)+Sh(end)+Slg(end)+Xmb(end)+Xtb(end)+Xma(end)+Xta(end)+Xmf(end)+Xtf(end)+...
    Xdb(end)+Xi(end)+NH4(end)+W(end)+NO3(end)+O2diss(end)+Xa(end)+NNH3(end));

Mcompost ; %kg/kgTM

Ccomfinal = Ccompost(end) 
Ncomfinal = Ncompost(end)

MS = (Mcompost - (W(end))) ; %kg/kgTM
MSfinal = MS(end);
MSfinal; 


Wvapfinal = cumsum(Wvap);


%calculation of C and N in inert material formed C1.48H1.5O0.85N0.12

Xi_init = Xi(1); %kg/kgTM
Xi_formed = Xi(end) - Xi_init; %kg/kgTM
M_Xi = 0.03454 ; %kg/mol

M_element = [0.012, 0.014]; %molar mass
n_element = [1.48, 0.12];

C_Xi =(Xi_formed*M_element(1)*n_element(1))/(M_Xi *Cinit); %kgC_Xi/kgCinit
N_Xi =(Xi_formed*M_element(2)*n_element(2))/(M_Xi *Ninit); %kgN_Xi/kgNinit


%calculation of C and N in organic matter of compost
compounds = {'C', 'P', 'L', 'H', 'CE', 'LG', 'Sc', 'Sp', 'Sl', 'Sh', 'Slg'};
Mmol = [0.180, 0.352, 0.393, 0.282, 0.180, 0.366, 0.180, 0.352, 0.393, 0.282, 0.366]; %molar mass
atomC = [6, 16, 25, 10, 6, 20, 6, 16, 25, 10, 20];
atomN = [0, 4, 0,0,0,0,0,4,0,0,0];

% Initialization
Corg = 0;
Norg = 0;

for i = 1:length(compounds)
    end_value = eval([compounds{i}, '(end)']); 
    C_value = ((end_value / Mmol(i)) * atomC(i) * 12e-3) / Cinit;
    N_value = ((end_value / Mmol(i)) * atomN(i) * 14e-3) / Ninit;
    Corg = Corg + C_value; %kgC/kgCinit
    Norg = Norg + N_value;
end;

Corg;
Norg;

%Calculation of elements in biomass in the compost
biomass = {'Xmb', 'Xtb', 'Xma', 'Xta', 'Xmf', 'Xtf', 'Xa', 'Xdb'};
Mmol_X = [0.113, 0.113, 0.113, 0.113, 0.247, 0.247, 0.113, 0.113];
atomC_X = [5, 5, 5, 5, 10, 10, 5, 5];
atomN_X = [1,1,1,1,1,1,1,1];

% Initialization
Cbio = 0;
Nbio = 0;

for i = 1:length(biomass)
    end_value_X = eval([biomass{i}, '(end)']); % Récupérer la valeur finale du composé
    C_value_X = ((end_value_X / Mmol_X(i)) * atomC_X(i) * 12e-3) / Cinit;
    N_value_X = ((end_value_X / Mmol_X(i)) * atomN_X(i) * 14e-3) / Ninit;
    Cbio = Cbio + C_value_X;
    Nbio = Nbio + N_value_X;
end
Cbio;
Nbio;


%to convert unity in 'per kg of biowaste entering the active phase'
%if biowaste/WC = 3/1 -> quantity of biowaste/TM = tot_biowaste = 3/4
%if biowaste/WC = 8/1 -> quantity of biowaste/TM = tot_biowaste = 8/9

tot_biowaste = 8/9;

CH4_biowaste = CH4final/tot_biowaste;
CH4_biowaste
CO2_biowaste = CO2final/tot_biowaste;
CO2_biowaste
NH3_biowaste = NH3final/tot_biowaste;
NH3_biowaste
N2O_biowaste = N2Ofinal/tot_biowaste;
N2O_biowaste
compost_biowaste = Mcompost/tot_biowaste;
compost_biowaste

NO3_biowaste = NO3(end)/tot_biowaste;
NO3_biowaste
NNO3_biowaste = NNO3(end)/tot_biowaste;
NNO3_biowaste
NH4_biowaste = NH4(end)/tot_biowaste;
NH4_biowaste
NNH4_biowaste = NNH4(end)/tot_biowaste;
NNH4_biowaste
Norg_biowaste = (Norg*Ninit)/tot_biowaste;
Norg_biowaste
N_Xi_biowaste = (N_Xi*Ninit)/tot_biowaste;
N_Xi_biowaste
Nbio_biowaste = (Nbio*Ninit)/tot_biowaste;
Nbio_biowaste

Ccompost_biowaste = Ccomfinal/tot_biowaste;
Ccompost_biowaste
