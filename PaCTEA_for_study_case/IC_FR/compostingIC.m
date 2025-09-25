function dydt=compostingIC(t,y);

C=y(1);	
P=y(2);
L=y(3);
H=y(4);
CE=y(5);
LG=y(6);
Xi=y(7);
Sc=y(8);	
Sp=y(9);
Sl=y(10);
Sh=y(11);
Slg=y(12);
Xmb=y(13);
Xtb=y(14);
Xma=y(15);
Xta=y(16);
Xmf=y(17);
Xtf=y(18);
Xdb=y(19);
CO2=y(20);
W=y(21);  
T=y(22);
CH4gen=y(23);
CH4oxi=y(24);
CH4=y(25);
Xa = y(26);
NO3 = y(27);
N2O =y(28);
N2 = y(29);
NH3 = y(30);
NH4 = y(31);
Wvap = y(32);
O2diss = y(33);
O2gaz = y(34);
O2cons = y(35);
O2out = y(36);
NH3gaz = y(37);

%Choice of formula
%if biowaste/WC = 3/1 -> BD = 599 kg/m3 (bulk density) and FAS = 0.4 (free air space)
%if biowaste/WC = 8/1 -> BD = 750 kg/m3 (bulk density) and FAS = 0.21 (free air space)

BD = 750; %m3
FAS =  0.21 ;


rhoair = 1.2; %kg/m3
Vtotal = (((3*1.5)/2)*6)/100000; %Un andain de 3m de larg, 1.5m de H, 6m de L %(0.80*0.80*1)/10000; %/100000; %
TM= Vtotal * BD;

%Choice of Ambient temperature
%if Ta = 25°C -> Ta = 298K
%if Ta = 5°C -> Ta = 280K
Ta = 298; %K

%Choice of aeration technology
%Passive aeration aer = 0.08 l/mn.kg
%Turning heaps aer = 0.08 l/mn.kg
%active aeration aer = 0.3 l/mn.kg

air = 0.3; %l/mn.kg

kh=[0.0293    0.1508    1.4563e-07    0.0053    0.1731    0.0182    0.0090 0.0070    0.0068    0.0090    0.0070    0.0068];

mu = [0.2 0.18 0.1 0.12 0.1 0.1 0.03];

bd = [0.03 0.02 0.01 0.015 0.01 0.01 0.0083];

K=[6.2e-5 1e-4 0.31 0.0025 0.007e-6 0.0608e-5];

KT = [440 1 2 0.09 10];

Yx_s= 0.35; %biomass yield on substrate kgX/kgS

Yx_co2 = [0.445717506 0.26543624	0.165284084	0.445717506	0.26543624	0.165284084	0.445717506	0.26543624	0.165284084	0.139609659	0.445717506	0.26543624	0.165284084	0.139609659	0.41509283	0.254264706	0.160882526	0.136456284	0.19653097	0.41509283	0.254264706	0.160882526	0.136456284	0.19653097
]; %Yield coeff of biomass on CO2 kgX/kgCO2

Yx_h = [0.716485507	1.791988467	0.380836347	0.716485507	1.791988467	0.380836347	0.716485507	1.791988467	0.380836347	...
    0.329440099	0.716485507	1.791988467	0.380836347	0.329440099	0.830451489	2.728498673	0.410802056	0.351627849	...
    0.625936213	0.830451489	2.728498673	0.410802056	0.351627849	0.625936213]; %Yield coeff of biomass on H2O kgX/kgH2O

Yx_O2mol = [0.173553719	0.098678414	0.044403196	0.173553719	0.098678414	0.044403196	0.173553719	0.098678414	0.044403196	0.054361283	0.173553719	0.098678414	0.044403196	0.054361283	0.07678245	0.044286279	0.019267472	0.024607076	0.027217967	0.07678245	0.044286279	0.019267472	0.024607076	0.027217967];
%yield coeff biomass on O2 molX/molO2

Yx_o2 = [0.61286157	0.34845815	0.156798785	0.61286157	0.34845815	0.156798785	0.61286157	0.34845815	0.156798785	0.191963281	0.61286157	0.34845815	0.156798785	0.191963281	0.592664534	0.341834717	0.148720799	0.189935871	0.210088681	0.592664534	0.341834717	0.148720799	0.189935871	0.210088681	0.040522541
]; %Yield coeff biomass on O2 kgX/kgO2

Yxa_no3 = 0.0390553; %Yield coeff of autotroph biomass on NO3 kgXa/kgNO3

Yxa_nh4 = 0.13; %Yield coeff of autotroph biomass on NO3 kgXa/kgNH4

Yx_nh4 = [6.64705882352941	2.490625447	6.647058824	6.64705882352941	2.490625447	6.647058824	0.705882353	2.490625447	6.647058824	6.647058824	0.705882353	2.490625447	6.647058824	6.647058824	14.52941176	2.069869946	14.52941176	14.52941176	14.52941176	14.52941176	2.069869946	14.52941176	14.52941176	14.52941176
]; %Yield coeff of heterotrophic biomass on NH3 kgX/kgNH3


Ks = K(1); %substrate saturation for monod kinetics (kgSc/dm3W)
Khs = K(2); %saturation coefficient for contois kinetics 
fi = K(3); %Proportion of dead biomass recycled to inert materials
kdec = K(4); %microorganisms decomposition constant
kO2=K(5); %oxygen saturation for heterotrophic activities (kgO2/l)
kO2nit = K(6); %oxygen saturation for nitrification (kgO2/l)


Vgas = FAS*Vtotal;

R = 8.134; %Pa.m3/mol.K

%Methane module parameters
Ych4_Sc = 0.267; %kgCH4/kgSC
Ych4_Sp = 0.375 ; %kgCH4/kgSp
Ych4_Sl = 0.707; %kgCH4/kgSl
Ych4_Sh = 0.284 ; %kgCH4/kgSh
Ych4_Slg = 0.535 ; %kgCH4/kgSlg
eta = 4e5; %; %L/mol %sensitivity of methangogenesis to inhibition by oxygen
Vmax = 5.35e-4; %kgCH4/kgTM.h maximum rate of methane oxidation
km = 0.72 ; %kg/l Michaelis constant for methane oxidation
Kch4_O2 = 0.033; %mol/l Michelis constant for oxygen in methane oxidation



hbio = KT(1); %heat released per mol of oxygen consumed(kJ/mol d'O2)
Ca = KT(2); %specific heat of dry air (kJ/K.kg)
Cw = KT(3); %specific heat of biowaste(kJ/K.kg)
U = KT(4); %heat transfer coefficient of wall (kJ/m2.K.h) %%valeur dans de Guardia 2012 : 7W/m2.C = 0.09 kJ/m2.h.K
A = KT (5)/100; %surface area of heat conduction (m2) (0.8*1*4)/1000 ; %

Qair  =  air*1e-3*TM*rhoair*60; %kg/h %  %passive 0.00396*TM /0.4;%

%hydrolysis constant
kh1C = kh(1); 	
kh2P = kh(2)*0.5;
kh3L = kh(3); 	
kh4C = kh(4);	
kh5P = kh(5)*0.5;	
kh6L = kh(6); 
kh7H = kh(7);	
kh8CE = kh(8);	
kh9LG = kh(9);	
kh10H = kh(10);	
kh11CE = kh(11);	
kh12LG = kh(12);

%growth rate
mmb = mu(1);
mtb = mu(2);
mma = mu(3);
mta = mu(4);
mmf = mu(5);
mtf = mu(6);
ma = mu(7)*0.1; 

%death rate
bmb = bd(1)*0.4;
btb = bd(2)*0.4;
bma = bd(3)*1;
bta = bd(4)*1;
bmf = bd(5)*1;
btf = bd(6)*1;
ba = bd(7); 

%Inverse of yield of biomass on CO2
Ymb_c_c = 1/Yx_co2(1) ; %MB on Sc
Ymb_p_c = 1/Yx_co2(2) ; %MB on Sp
Ymb_l_c = 1/Yx_co2(3) ; %MB on Sl
Ytb_c_c = 1/Yx_co2(4) ; %TB on Sc
Ytb_p_c = 1/Yx_co2(5) ; %TB on Sp
Ytb_l_c = 1/Yx_co2(6) ; %TB on Sl
Yma_c_c = 1/Yx_co2(7) ; %MA on Sc
Yma_p_c = 1/Yx_co2(8) ; %MA on Sp
Yma_l_c = 1/Yx_co2(9) ; %MA on Sl
Yma_h_c = 1/Yx_co2(10) ; %MA on Sh
Yta_c_c = 1/Yx_co2(11) ; %TA on Sc
Yta_p_c = 1/Yx_co2(12) ; %TA on Sp
Yta_l_c = 1/Yx_co2(13) ; %TA on Sl
Yta_h_c = 1/Yx_co2(14) ; %TA on Sh
Ymf_c_c = 1/Yx_co2(15) ; %MF on Sc
Ymf_p_c = 1/Yx_co2(16) ; %MF on Sp
Ymf_l_c = 1/Yx_co2(17) ; %MF on Sl
Ymf_h_c = 1/Yx_co2(18) ; %MF on Sh
Ymf_lg_c = 1/Yx_co2(19) ; %MF on Slg
Ytf_c_c = 1/Yx_co2(20) ; %TF on Sc
Ytf_p_c = 1/Yx_co2(21) ; %TF onSp
Ytf_l_c = 1/Yx_co2(22) ; %TF on Sl
Ytf_h_c = 1/Yx_co2(23) ; %TF on Sh
Ytf_lg_c = 1/Yx_co2(24) ; %TF on Slg

%Inverse of yield of biomass on H2O
Ymb_c_h = 1/Yx_h(1) ; %MB on Sc
Ymb_p_h = 1/Yx_h(2) ; %MB on Sp
Ymb_l_h = 1/Yx_h(3) ; %MB on Sl
Ytb_c_h = 1/Yx_h(4) ; %TB on Sc
Ytb_p_h = 1/Yx_h(5) ; %TB on Sp
Ytb_l_h = 1/Yx_h(6) ; %TB on Sl
Yma_c_h = 1/Yx_h(7) ; %MA on Sc
Yma_p_h = 1/Yx_h(8) ; %MA on Sp
Yma_l_h = 1/Yx_h(9) ; %MA on Sl
Yma_h_h = 1/Yx_h(10) ; %MA on Sh
Yta_c_h = 1/Yx_h(11) ; %TA on Sc
Yta_p_h = 1/Yx_h(12) ; %TA on Sp
Yta_l_h = 1/Yx_h(13) ; %TA on Sl
Yta_h_h = 1/Yx_h(14) ; %TA on Sh
Ymf_c_h = 1/Yx_h(15) ; %MF on Sc
Ymf_p_h = 1/Yx_h(16) ; %MF on Sp
Ymf_l_h = 1/Yx_h(17) ; %MF on Sl
Ymf_h_h = 1/Yx_h(18) ; %MF on Sh
Ymf_lg_h = 1/Yx_h(19) ; %MF on Slg
Ytf_c_h = 1/Yx_h(20) ; %TF on Sc
Ytf_p_h = 1/Yx_h(21) ; %TF onSp
Ytf_l_h = 1/Yx_h(22) ; %TF on Sl
Ytf_h_h = 1/Yx_h(23) ; %TF on Sh
Ytf_lg_h = 1/Yx_h(24) ; %TF on Slg

%Inverse of yield of biomass on O2
Ymb_c_o2 = 1/Yx_o2(1) ; %MB on Sc
Ymb_p_o2 = 1/Yx_o2(2) ; %MB on Sp
Ymb_l_o2 = 1/Yx_o2(3) ; %MB on Sl
Ytb_c_o2 = 1/Yx_o2(4) ; %TB on Sc
Ytb_p_o2 = 1/Yx_o2(5) ; %TB on Sp
Ytb_l_o2 = 1/Yx_o2(6) ; %TB on Sl
Yma_c_o2 = 1/Yx_o2(7) ; %MA on Sc
Yma_p_o2 = 1/Yx_o2(8) ; %MA on Sp
Yma_l_o2 = 1/Yx_o2(9) ; %MA on Sl
Yma_h_o2 = 1/Yx_o2(10) ; %MA on Sh
Yta_c_o2 = 1/Yx_o2(11) ; %TA on Sc
Yta_p_o2 = 1/Yx_o2(12) ; %TA on Sp
Yta_l_o2 = 1/Yx_o2(13) ; %TA on Sl
Yta_h_o2 = 1/Yx_o2(14) ; %TA on Sh
Ymf_c_o2 = 1/Yx_o2(15) ; %MF on Sc
Ymf_p_o2 = 1/Yx_o2(16) ; %MF on Sp
Ymf_l_o2 = 1/Yx_o2(17) ; %MF on Sl
Ymf_h_o2 = 1/Yx_o2(18) ; %MF on Sh
Ymf_lg_o2 = 1/Yx_o2(19) ; %MF on Slg
Ytf_c_o2 = 1/Yx_o2(20) ; %TF on Sc
Ytf_p_o2 = 1/Yx_o2(21) ; %TF onSp
Ytf_l_o2 = 1/Yx_o2(22) ; %TF on Sl
Ytf_h_o2 = 1/Yx_o2(23) ; %TF on Sh
Ytf_lg_o2 = 1/Yx_o2(24) ; %TF on Slg
Yxa_o2 = 1/Yx_o2(25) ; %Xa on NH4

%Inverse of yield of biomass on NH4
Ymb_c_nh4 = 1/Yx_nh4(1) ; %MB on Sc
Ymb_p_nh4 = 1/Yx_nh4(2) ; %MB on Sp
Ymb_l_nh4 = 1/Yx_nh4(3) ; %MB on Sl
Ytb_c_nh4 = 1/Yx_nh4(4) ; %TB on Sc
Ytb_p_nh4 = 1/Yx_nh4(5) ; %TB on Sp
Ytb_l_nh4 = 1/Yx_nh4(6) ; %TB on Sl
Yma_c_nh4 = 1/Yx_nh4(7) ; %MA on Sc
Yma_p_nh4 = 1/Yx_nh4(8) ; %MA on Sp
Yma_l_nh4 = 1/Yx_nh4(9) ; %MA on Sl
Yma_h_nh4 = 1/Yx_nh4(10) ; %MA on Sh
Yta_c_nh4 = 1/Yx_nh4(11) ; %TA on Sc
Yta_p_nh4 = 1/Yx_nh4(12) ; %TA on Sp
Yta_l_nh4 = 1/Yx_nh4(13) ; %TA on Sl
Yta_h_nh4 = 1/Yx_nh4(14) ; %TA on Sh
Ymf_c_nh4 = 1/Yx_nh4(15) ; %MF on Sc
Ymf_p_nh4 = 1/Yx_nh4(16) ; %MF on Sp
Ymf_l_nh4 = 1/Yx_nh4(17) ; %MF on Sl
Ymf_h_nh4 = 1/Yx_nh4(18) ; %MF on Sh
Ymf_lg_nh4 = 1/Yx_nh4(19) ; %MF on Slg
Ytf_c_nh4 = 1/Yx_nh4(20) ; %TF on Sc
Ytf_p_nh4 = 1/Yx_nh4(21) ; %TF onSp
Ytf_l_nh4 = 1/Yx_nh4(22) ; %TF on Sl
Ytf_h_nh4 = 1/Yx_nh4(23) ; %TF on Sh
Ytf_lg_nh4 = 1/Yx_nh4(24) ; %TF on Slg



%limitation by the substrate
fSc = (Sc/W)/(Ks+(Sc/W));
fSp=(Sp/W)/(Ks+(Sp/W));
fSl=(Sl/W)/(Ks+(Sl/W));
fSh=(Sh/W)/(Ks+(Sh/W));
fSlg=(Slg/W)/(Ks+(Slg/W));

%limitation function by temperature
Ti = T ;
Tmax1 = 355; 
Tmin1 =  273  ; 
Topt1 =  322; 
fT1 = ((Ti-Tmax1)*(Ti-Tmin1)^2)/((Topt1-Tmin1)*((Topt1-Tmin1) *(Ti-Topt1)-(Topt1-Tmax1)*(Topt1+Tmin1-2*Ti))) ;

%limitation function by oxygen
fO2 = (O2diss/W)/(kO2+ (O2diss/W)); %heteretroph activities
fO2nit = (O2diss/W)/(kO2nit + (O2diss/W)); %autotroph activity


%Limitation function of ammonium for nitrification
Kn = 10^((0.051*(T-275)-7.158)); %kg/l
fNH4nit = (NH4/W)/(Kn + (NH4/W));

%Limitation of denitrification by NO3
kNO3 =  8.4e-2; %kg/l 0.18 ;
flimNO3 = (NO3/W)/(kNO3+(NO3/W));

%Limitation of denitrification by temperature
if T<286
    fTdenit = exp((((T-275)-11)*log(89)-9*log(2.1))/10);
else
    fTdenit = exp((((T-275)-20)*log(2.1))/10);
end;


%hydrolisis rate
v1 = kh1C*Xmb*(C/((Khs*Xmb)+C));
v2 =kh2P* Xmb*(P/((Khs*Xmb)+P)) ;
v3 = kh3L *Xmb*(L/((Khs*Xmb)+L));
v4 = kh4C*Xtb*(C/((Khs*Xtb)+C));
v5= kh5P*Xtb*(P/((Khs*Xtb)+P));
v6 = kh6L*Xtb*(L/((Khs*Xtb)+L));
v7 = kh7H*Xta*(H/((Khs*Xta)+H));
v8 = kh8CE*Xtf*(CE/((Khs*Xtf)+CE));
v9 = kh9LG*Xtf*(LG/((Khs*Xtf)+LG));
v10 = kh10H*Xma*(H/((Khs*Xma)+H));
v11 = kh11CE*Xmf*(CE/((Khs*Xmf)+CE));
v12 = kh12LG*Xmf*(LG/((Khs*Xmf)+LG));

%Biomass growth reaction
% Mesophilic Bacterias growth
v13 = mmb*Xmb  *fSc *fT1 *fO2;
v14 = mmb*Xmb  * fSp *fT1*fO2 ;
v15 = mmb*Xmb  *fSl *fT1*fO2;
%Thermophilic bacterias growth
v16 = mtb*Xtb  * fSc *fT1*fO2; %*fT2;
v17 = mtb*Xtb *fSp *fT1*fO2;%*fT2;
v18 = mtb*Xtb *fSl *fT1*fO2;%*fT2;
%Mesophilic Actinomycetes growth
v19 = mma*Xma   * fSc *fT1*fO2;
v20 = mma*Xma    *fSp *fT1*fO2;
v21 = mma*Xma  *fSl *fT1*fO2;
v22 = mma*Xma   *fSh *fT1*fO2;
 %Thermophilic actinomycetes growth
v23 = mta*Xta  *fSc *fT1*fO2;%*fT2;
v24 = mta*Xta  *fSp*fT1*fO2;% *fT2;
v25 = mta*Xta  *fSl *fT1*fO2;%*fT2;
v26 = mta*Xta  *fSh*fT1*fO2; %*fT2;
%Mesophilic fungis growth
v27 = mmf*Xmf  *fSc *fT1*fO2;
v28 = mmf*Xmf  * fSp *fT1*fO2;
v29 = mmf*Xmf  * fSl *fT1*fO2;
v30 = mmf*Xmf  *  fSh *fT1*fO2;
v31 = mmf*Xmf  * fSlg *fT1*fO2; 
%Thermophilic fungis growth
v32 = mtf*Xtf  *fSc *fT1*fO2;%*fT2;
v33 = mtf*Xtf *fSp *fT1*fO2;%*fT2;
v34 = mtf*Xtf   *fSl *fT1*fO2;%*fT2; 
v35 = mtf*Xtf    *fSh*fT1*fO2;%*fT2;
v36 = mtf*Xtf   *fSlg *fT1*fO2;%*fT2;

%Heterotroph Microorganisms death
v37 = bmb*Xmb ;
v38 = btb*Xtb ;
v39 = bma*Xma ;
v40 = bta*Xta;
v41 = bmf*Xmf;
v42 = btf*Xtf;

%lysis of microorganisms
v43 = kdec*Xdb;

%Autotroph microorganisms growth
v44 = ma * Xa  * fNH4nit * fO2nit;

%denitrification
pmaxdenit =0.01; %0.042*10 ; %0.3; % ; %kg(N2O+N2)/kgNO3.h

%pN2Odenit = 0.2; %kg(N2O)/kg(N2O+N2)

v45 = pmaxdenit * NO3 * flimNO3 * fTdenit;

%autotroph biomass decay
v46 = ba * Xa;

%Ammoniac emission
v47 = Qair * NH3gaz /((CO2+CH4+N2+NH3gaz+N2O+O2gaz)*TM);

ntot = ((CO2/0.044)+(CH4/0.016)+(NH3gaz/0.017)+(N2O/0.044)+(N2/0.028)+(Wvap/0.018)+(O2gaz/0.032))*TM;  %mol


%Global equations
dCdt= -v1 -v4;
dPdt =   -v5 -v2 + 0.7*(v43+v46); %0.69
dLdt = -v3-v6;
dHdt= -v7-v10;
dCEdt = -v8-v11;
dLGdt= -v9-v12;
dXidt= fi*(v43+v46);
dScdt =  v1 +v4 +v11 +v8 - (1/Yx_s)*(v13+v16+v19+v23+v27+v32);    
dSpdt =  v2  +v5 -(1/Yx_s)*(v14+v17+v20+v24+v28+v33);
dSldt = v3+v6-(1/Yx_s)*(v15+v18+v21+v25+v29+v34);
dShdt = v7+v10-(1/Yx_s)*(v22+v26+v30+v35);
dSlgdt = v9+v12-(1/Yx_s)*(v31+v36);
dXmbdt = v13+ v14 + v15 -v37;
dXtbdt = v16 + v17 + v18 - v38;
dXmadt = v19+v20 + v21 +v22- v39;
dXtadt = v23 + v24 + v25 +v26- v40;
dXmfdt = v27 + v28 + v29 +v30+ v31 - v41;
dXtfdt = v32 + v33 + v34 +v35+v36 - v42;           
dXdbdt = v37+v38+v39+v40+v41+v42-v43;
            
dCO2dt = (Ymb_c_c)*v13+(Ymb_p_c)*v14+(Ymb_l_c)*v15+(Ytb_c_c)*v16+(Ytb_p_c)*v17+(Ytb_l_c)*v18+(Yma_c_c)*v19+(Yma_p_c)*v20+(Yma_l_c)*v21+...
(Yma_h_c)*v22+(Yta_c_c)*v23+(Yta_p_c)*v24+(Yta_l_c)*v25+(Yta_h_c)*v26+(Ymf_c_c)*v27+(Ymf_p_c)*v28+(Ymf_l_c)*v29+...
(Ymf_h_c)*v30+(Ymf_lg_c)*v31+(Ytf_c_c)*v32+(Ytf_p_c)*v33+(Ytf_l_c)*v34+(Ytf_h_c)*v35+(Ytf_lg_c)*v36 + 2.75*(Vmax*((CH4gen/W)/(km+(CH4gen/W)))*((O2diss/(W*32e-3))/(Kch4_O2+(O2diss/(W*32e-3)))));



%Water module         
Pwvap = ((Wvap/0.018)*R*T)/Vgas;  %Pa
Ptot = (ntot*R*T)/Vgas;
klh2O = 1e-4 ; %Petric 2008
Ps = 10^(22.443-(2795/(T))-(1.6798*log(T))); %Petric 2008

dWvapdt = (klh2O/TM) * (Ps - Pwvap) - Qair*Wvap/(rhoair *Vgas*TM);

dWdt =(((Ymb_c_h)*v13+(Ymb_p_h)*v14+(Ymb_l_h)*v15+(Ytb_c_h)*v16+(Ytb_p_h)*v17+(Ytb_l_h)*v18+(Yma_c_h)*v19+(Yma_p_h)*v20+(Yma_l_h)*v21+...
    (Yma_h_h)*v22+(Yta_c_h)*v23+(Yta_p_h)*v24+(Yta_l_h)*v25+(Yta_h_h)*v26+(Ymf_c_h)*v27+(Ymf_p_h)*v28+(Ymf_l_h)*v29+...
    (Ymf_h_h)*v30+(Ymf_lg_h)*v31+(Ytf_c_h)*v32+(Ytf_p_h)*v33+(Ytf_l_h)*v34+(Ytf_h_h)*v35+(Ytf_lg_h)*v36)) - ((klh2O/TM) * (Ps - Pwvap) - Qair*Wvap/(rhoair *Vgas*TM))  + 2.25*Vmax*((CH4gen/W)/(km+(CH4gen/W)))*((O2diss/(W*32e-3))/(Kch4_O2+(O2diss/(W*32e-3)))); %Petric 2005 + 0.842*v45

He =  4.39751e9;%Pa %

%Oxygen module
klO2ref = 180; %h-1; 
klO2 = klO2ref * 1.02^(T-293);
Ho2 = 1.3e-5*exp((12000/8.314)*((1/T)-(1/298)));

O2eq = Ho2*((O2gaz*R*T)/(Vgas*0.032))*W*1e-3*0.032*TM; %Henry law Ceq = H* p, with conversion in order to have kgO2/TM

dO2dissdt = klO2*(O2eq-O2diss) - ((Ymb_c_o2)*v13+(Ymb_p_o2)*v14+(Ymb_l_o2)*v15+(Ytb_c_o2)*v16+(Ytb_p_o2)*v17+(Ytb_l_o2)*v18+(Yma_c_o2)*v19+(Yma_p_o2)*v20+(Yma_l_o2)*v21+...
    (Yma_h_o2)*v22+(Yta_c_o2)*v23+(Yta_p_o2)*v24+(Yta_l_o2)*v25+(Yta_h_o2)*v26+(Ymf_c_o2)*v27+(Ymf_p_o2)*v28+(Ymf_l_o2)*v29+...
    (Ymf_h_o2)*v30+(Ymf_lg_o2)*v31+(Ytf_c_o2)*v32+(Ytf_p_o2)*v33+(Ytf_l_o2)*v34+(Ytf_h_o2)*v35+(Ytf_lg_o2)*v36+(Yxa_o2)*v44+4*Vmax*((CH4gen/W)/(km+(CH4gen/W)))*((O2diss/(W*32e-3))/(Kch4_O2+(O2diss/(W*32e-3)))));

dO2consdt =  ((Ymb_c_o2)*v13+(Ymb_p_o2)*v14+(Ymb_l_o2)*v15+(Ytb_c_o2)*v16+(Ytb_p_o2)*v17+(Ytb_l_o2)*v18+(Yma_c_o2)*v19+(Yma_p_o2)*v20+(Yma_l_o2)*v21+...
    (Yma_h_o2)*v22+(Yta_c_o2)*v23+(Yta_p_o2)*v24+(Yta_l_o2)*v25+(Yta_h_o2)*v26+(Ymf_c_o2)*v27+(Ymf_p_o2)*v28+(Ymf_l_o2)*v29+...
    (Ymf_h_o2)*v30+(Ymf_lg_o2)*v31+(Ytf_c_o2)*v32+(Ytf_p_o2)*v33+(Ytf_l_o2)*v34+(Ytf_h_o2)*v35+(Ytf_lg_o2)*v36+(Yxa_o2)*v44);

dO2gazdt= Qair*0.23/TM - klO2*(O2eq-O2diss) - (Qair*O2gaz)/((CO2+CH4+N2+NH3gaz+N2O+O2gaz)*TM*rhoair); %0.23 is* the mass fraction of O2 in the air, 

dO2outdt = (Qair*O2gaz)/((CO2+CH4+N2+NH3gaz+N2O+O2gaz)*TM*rhoair);


%temperature module
Qbio = hbio *(((v13/Yx_O2mol(1))+(v14/Yx_O2mol(2))+(v15/Yx_O2mol(3))+(v16/Yx_O2mol(4))+(v17/Yx_O2mol(5))+(v18/Yx_O2mol(6))+(v19/Yx_O2mol(7))+(v20/Yx_O2mol(8))+...
    (v21/Yx_O2mol(9))+(v22/Yx_O2mol(10))+(v23/Yx_O2mol(11))+(v24/Yx_O2mol(12))+(v25/Yx_O2mol(13))+(v26/Yx_O2mol(14)))/0.113+((v27/Yx_O2mol(15))+(v28/Yx_O2mol(16))+(v29/Yx_O2mol(17))+...
    (v30/Yx_O2mol(18))+(v31/Yx_O2mol(19))+(v32/Yx_O2mol(20))+(v33/Yx_O2mol(21))+(v34/Yx_O2mol(22))+(v35/Yx_O2mol(23))+(v36/Yx_O2mol(24)))/0.247)* TM ;

Qconv = Ca*( T- Ta)*Qair ;

Qcond = (T - Ta)*U*A ;

dTdt= (Qbio-Qconv-Qcond)/((TM*Cw)+(4*W*TM)) ; %4 is the specific heat of water in kJ/kg.K
                              
%Methane module
dCH4gendt= (Ych4_Sc*(v1+v4+v8+v11 )+Ych4_Sp*(v2+v5 )+Ych4_Sl*(v3+v6 )+Ych4_Sh* (v7+v10 )+Ych4_Slg*(v9+v12 ))*(1/(1+eta*(O2diss/(W*32e-3)))); %kg/kgTM.h


dCH4oxidt = Vmax*((CH4gen/W)/(km+(CH4gen/W)))*((O2diss/(W*32e-3))/(Kch4_O2+(O2diss/(W*32e-3))));


dCH4dt= dCH4gendt - dCH4oxidt ;

%nitrification module
dXadt = v44 - v46;

dNO3dt = (1/Yxa_no3)*v44 - v45;

dN2Odt =0.0710*v45;

dN2dt = 0.181*v45; 


%Volatilization module

klaNH4 = 100*exp(0.03*(T-298)); 

khNH3 = 101325* exp(160.559 - (8621.06/T)-(25.6767*log(T))+(0.035388*T)) ;

NH4eq = 6e-4*((NH3gaz*R*T)/(Vgas*0.017))*W*0.017*TM ;

dNH3gazdt = klaNH4*(NH4-NH4eq) - v47;

dNH3dt = v47;

dNH4dt = -(1/Yxa_nh4)*v44 +(Ymb_p_nh4)*v14-(Ymb_l_nh4)*v15  + (Ytb_p_nh4)*v17 -(Ytb_l_nh4)*v18 -(Yma_c_nh4)*v19+(Yma_p_nh4)*v20-(Yma_l_nh4)*v21-...
(Yma_h_nh4)*v22-(Yta_c_nh4)*v23+(Yta_p_nh4)*v24-(Yta_l_nh4)*v25-(Yta_h_nh4)*v26-(Ymf_c_nh4)*v27+(Ymf_p_c)*v28-(Ymf_l_nh4)*v29-...
(Ymf_h_nh4)*v30-(Ymf_lg_nh4)*v31-(Ytf_c_nh4)*v32+(Ytf_p_c)*v33-(Ytf_l_nh4)*v34-(Ytf_h_nh4)*v35-(Ytf_lg_nh4)*v36- (Ymb_c_nh4)*v13-(Ytb_c_nh4)*v16 - klaNH4*(NH4-NH4eq) ;

dydt = [dCdt,dPdt,dLdt,dHdt,dCEdt,dLGdt,dXidt, dScdt,dSpdt,dSldt,dShdt,dSlgdt, dXmbdt,dXtbdt,dXmadt,dXtadt,dXmfdt,dXtfdt, dXdbdt, dCO2dt, dWdt,dTdt,dCH4gendt, dCH4oxidt,dCH4dt,...
    dXadt, dNO3dt, dN2Odt, dN2dt, dNH3dt, dNH4dt,dWvapdt,dO2dissdt,dO2gazdt,dO2consdt,dO2outdt,dNH3gazdt]'; %,dO2transdt,dO2indt]' %]';