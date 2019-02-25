# -*- coding: utf-8 -*-
"""
Created on Mon Sep 10 20:10:52 2018
@author: pumaflores
"""

import win32com.client
import matplotlib.pyplot as plt
from pylab import *
import numpy as np
import math
from FA_funciones_opendss import *
import pandas as pd
import time

star = time.time()
# =========================================================================== #
#                           OpenDSS_Program!                                  #
# =========================================================================== #

if __name__ == "__main__":
#    objeto = DSS("MASTER_RedeTeste13Barras.dss")                               # stepsize =0.25; number =50 # intervalo de tiempo # points os measurements
#    objeto = DSS("1_Feec_trafo500.dss")                                        # criar objeto de clase dss
    objeto = DSS("1_Feec_7trafos.dss")
    print (u"Version de OpenDSS: " + objeto.versao_DSS() + "\n")
    # Datos de configuracion
    lim_number=1 ;stepsize=0.25; number=2880; LoadMult=1   
    # limitedesimu  # pasodesimu     #comienzodesimu  #factordecarga

# =========================================================================== #
#                   Transformer information!                                  #
# =========================================================================== #  
#    # tensiones y angulos del transformador    
#    objeto.compila_DSS()    
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    VH,VL,ABV=get_GeralInformation(objeto,lim_number,stepsize,number,LoadMult)

# =========================================================================== #
#                           Read_Load_Caracteristic                           #
# =========================================================================== #
    # Read loadshapes
    LoadShape = 'Loadshape_09_18.csv'                                           # fecha de loadshape
    # Analisis de datos
    get_AnalisisDeCarga(LoadShape)
    
## =========================================================================== #
##                   Smart Meter information - tensiones!                                  #
## =========================================================================== #
## Informacion de los medidores conectados a la red, tensiones y corrientes
#    #calculo de perdas del transformador
#    objeto.compila_DSS()    
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    get_TensionesCorrientesPerdidas(objeto)

# =========================================================================== #
#                   Transformers information!                                  #
# =========================================================================== #
#    # Informacion de perdida de transformadores a partir de los medidores
#    objeto.compila_DSS()
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    # Informacion de transfromadores 
#    objeto.get_summary_transformer()
# =========================================================================== #
#                   Transformer - Smart Meter Information!        #
# =========================================================================== #
#     #Informacion de perdida de transformadores a partir de los medidores
#    objeto.compila_DSS()
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#	 # Informacion de medidores
#    objeto.get_summary_energy_meters()
    
# =========================================================================== #
#                   (*)Seleccionar transfromador con menor perdida        #
# =========================================================================== #
##     Informacion de perdida de transformadores a partir de los medidores
#    objeto.compila_DSS()
#    
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    TFloss=[]
#    Tloss,n =objeto.get_summary_energy_meters() 
#    TFloss.append(Tloss[0])
# =========================================================================== #
#                   Seleccionar transfromador con menor perdida - cambio uno por uno  #
# =========================================================================== #
    
#    objeto.compila_DSS()
#    objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    TFloss=[]
#    Tloss,n =objeto.get_summary_energy_meters() 
#    TFloss.append(Tloss[0])    
#    # lista de transformadores
#    print(TFloss)
##     agrega transformador - 
# #    n=0 (500kva) n=1;300kva n=2 ;225kva n=3 ;150kva n=4 ;112kva n=5 ;75kva n=6 ;45kva
#    xx = []
#    for i in range(0,7):
#        objeto.compila_DSS()
#        objeto.get_AddTransformer(i) # cambiar transfromador
#        objeto.solve_DSS_daily(LoadMult,stepsize,number)
#    
#        
# #        objeto.get_summary_energy_meters
# #        objeto.get_summary_transformer()
#        Tloss,n =objeto.get_summary_energy_meters()
#        xx.append(Tloss[0])
#        print(Tloss[0])
#    print(xx)
#    a = np.array(xx)
#    indice_max = int(np.where(a==max(a))[0])
#    indice_min = int(np.where(a==min(a))[0])
#    print('El transfromador que presenta menos perdidas es:'+str(indice_min+1))

# =========================================================================== #
#                   Seleccionar transfromador con menor perdida        #
# =========================================================================== #
##    # transformadores de comparacion
##    #    S   \  Pfe \  Pcu \ Pt
##    Tf1 =[15   , 85	 , 325  , 410 ]
##    Tf2 =[30   , 150 , 545  , 695 ]
#    Tf3 =[45   , 195 , 750  , 945 ]
#    Tf4 =[75   , 295 , 1100 , 1395]
#    Tf5 =[112.5, 390 , 1500 , 1890]
#    Tf6 =[150  , 485 , 1850 , 2335]
#    Tf7 =[225  , 650 , 2610 , 3260]
#    Tf8 =[300  , 810 , 3250 , 4060]
#    Tf9 =[500  , 1143, 6266 , 7409] # valores reales - CC-CA
##    Tf9 =[500  , 1200, 6000 , 7200] # Norma IRAN
#    
##    Trafos = [Tf1,Tf2,Tf3,Tf4,Tf5,Tf6,Tf7,Tf8,Tf9]
#    Trafos = [Tf3,Tf4,Tf5,Tf6,Tf7,Tf8,Tf9]
#    # determina las perdidas en cada trasfromador y compara con otros TF alternos
#    get_BestTransformer1(objeto,LoadMult,stepsize,number,Trafos)

# =========================================================================== #
#                   analisis economico        #
# =========================================================================== # 
    # Determinacion de analisis economico de las perdidas de cada transformador 
    # Perdidas de transformadores nuevos y antiguo

#    # precios de transformadores
#    Tf1 =[15   , 1000, 1500]
#    Tf2 =[30   , 1000, 1500]
#    Tf3 =[45   , 1000, 1500]
#    Tf4 =[75   , 1000, 1500]
#    Tf5 =[112.5, 1000, 1500]
#    Tf6 =[150  , 1000, 1500]
#    Tf7 =[225  , 1000, 1500]
#    Tf8 =[300  , 1000, 1500]
#    Tf9 =[500  , 1000, 1500]
    
#    Tf5=[30	  ,22.69  ]
#    Tf6=[15	  ,45.45  ]
#    Tf4=[45	  ,25.92  ]
#    Tf3=[75	  ,5531.7 ]
#    Tf2=[112.5 ,22.78  ]
#    Tf1=[150	  ,7288.85]

    
    # Custo anual das perda no ferro
# =========================================================================== #
#                    perdida de vida del transfromador        #
# =========================================================================== #  
##    perdida de vida del transfromador de acuerdo en funcion a temperatura del bobinado
##    referencia 110°C
##    TLL = 1
##    print(get_TransformerLossLife)
#
#    Th = 110        # is the winding hottest-spot temperature, °C
##    parametros constantes 
#    A = 9.8e-18
#    B = 15000
#    
#    TLL = A*math.exp(B/(Th+273))
#    print(TLL)

# =========================================================================== # 
#                    caracteristicas de carga
# =========================================================================== #

#    datos = pd.read_excel('meter_2018.xlsx')
#    n = 3*2880
#    datos = datos.ix[n:n+2880]
#    dato = datos['kVA tot mean']
#    date = datos['Date']
#    dato=np.array(dato)
#    ts = pd.Series(dato, index=date)
#    # analise de acordo as fechas 
#    #D: Día natural
#    #B: Día hábil
#    #W: semanal (último día de la semana)
#    #M: mensual (último día del mes)
#    #SM: bimensual (el día 15 y último del mes)
#    #Q: trimestral (último día del trimestre)
#    me=ts.resample('D').mean()
#    ma=ts.resample('B').mean()
#    #ma=ts.resample('D').max()
#    
#    plt.plot(me)#
#    plt.plot(ma)#
#    _=plt.xticks(rotation=45) # rota el indice de x



# =========================================================================== #

end = time.time()

print('\n        =======================================')
print('        | tiempo transcurrido: '+str("%.3f"%(end-star))+' segundos |')
print('        =======================================')


# formato de numero
#numero = raw_input("Dame un numero: ")
#numero = float(numero)
#print "%.2f" % (numero, )
#print "%.*f" % (2, numero)
