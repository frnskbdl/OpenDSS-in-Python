# -*- coding: utf-8 -*-
"""
Created on Tue Oct 23 09:41:42 2018

@author: pumaflores
"""
import win32com.client
import matplotlib.pyplot as plt
import numpy as np
import math
import pandas as pd

# =========================================================================== #
#                           OpenDSS_Functions!                                #
# =========================================================================== #
class DSS():
    def __init__(self, end_modelo_DSS):
        self.end_modelo_DSS = end_modelo_DSS
        # Criar a conexao entre Python e OpenDss
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
        # Criar o objeto DSS
        if self.dssObj.Start(0) == False:
            print ("DSS Failed to Start")
        else:
            # Criar variaveis para as principais interfaces
            self.dssText        = self.dssObj.Text
            self.dssCircuit     = self.dssObj.ActiveCircuit
            self.dssSolution    = self.dssCircuit.Solution
            self.dssLines       = self.dssCircuit.Lines
            self.dssCktElement  = self.dssCircuit.ActiveCktElement
            self.dssBus         = self.dssCircuit.ActiveBus
            self.dssTransformer = self.dssCircuit.Transformers
            self.dssMeters      = self.dssCircuit.Meters
            self.dssMonitors    = self.dssCircuit.Monitors

# =========================================================================== #
    def versao_DSS(self):
        return self.dssObj.Version

    def compila_DSS(self):
        self.dssObj.ClearAll()      # Limpar informacoes de ultimas simulacoes
        self.dssText.Command = "compile " + self.end_modelo_DSS
        
# =========================================================================== #
    def solve_DSS_snapshot(self, multiplicador_carga):# Configuracoes de configuracao
        self.dssText.Command = "Set Mode=SnapShot"
        self.dssText.Command = "Set ControlMode=Static"
        self.dssSolution.LoadMult = multiplicador_carga  # multiplicador de carga
        self.dssSolution.Solve()         # Resolver o fluxo de potencia
    
    def solve_DSS_daily(self, multiplicador_carga,h,number): #need multload,stepsize and number points
        self.dssSolution.LoadMult = multiplicador_carga
        self.dssText.Command = "Set Mode=daily Stepsize="+str(h)+"h number="+str(number)
        self.dssText.Command = "Set ControlMode=Static"
        self.dssSolution.Solve()
        
# =========================================================================== #

    def get_AddLoad(self):
#        print(self.dssText.Command = "? Load.1.kW")
        self.dssText.command = "Load.1.kW=10"
#        print(self.dssText.Command = "? Load.1.kW" )
#        self.dssText.Command = 'New Load.12 Bus1=1.1  Phases=1 Conn=Wye Model=1 daily=Loadshape_06_18 kV=0.22  kW=1 kvar=0'
        
    def get_ChangeTranformer(self,name,S,loadloss,noloadloss):
        self.dssText.Command = "Transformer."+str(name)+".kVA="+str(S)
        self.dssText.Command = "Transformer."+str(name)+".%loadloss="+str(loadloss)
        self.dssText.Command = "Transformer."+str(name)+".%noloadloss="+str(noloadloss)
        
    def get_AddTransformer(self,n):
        t500='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(500,500)     %loadloss=1.2000 %noloadloss=0.24           !%imag=2 xhl=3.88'
        t300='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(300,300)     %loadloss=1.1600 %noloadloss=0.2889         !%imag=2 xhl=3.88'
        t225='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(225,225)     %loadloss=1.0833 %noloadloss=0.2700         !%imag=2 xhl=3.88'
        t150='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(150,150)     %loadloss=1.2333 %noloadloss=0.3233         !%imag=2 xhl=3.88'
        t112='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(112.5,112.5) %loadloss=1.3333 %noloadloss=0.3467         !%imag=2 xhl=3.88'
        t75 ='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(75,75)       %loadloss=1.4667 %noloadloss=0.3933         !%imag=2 xhl=3.88'
        t45 ='New Transformer.500kVA phases=3 windings=2 buses=(SourceBus, 1) Conns=(delta, Wye) kVs=(13.8,0.22) kVAs=(45,45)       %loadloss=1.6667 %noloadloss=0.4333         !%imag=2 xhl=3.88'
        Transformer = [t500,t300,t225,t150,t112,t75,t45]
        
        self.dssText.Command = Transformer[n]
        
    def get_TransformersInformation(self):
        n=self.dssTransformer.Count
        
        return n
        
          
        
#        dssTransformer=dssCircuit.Transformers
#        self.dssTransformer.Name='500kVA'
#        self.dssTransformer.KW  = 300
        
#        self.dssText.Command ='Transformer.500kVA.kW=300'
        
#        self.dssCircuit.SetActiveElement('Transformer.500kVA')
#        self.dssCircuit.Transformer.Name='500kVA'
#        self.dssCircuit.Transformer.kW=float(SizeEntry.get())
        
        
#    DSSCircuit.Generators.PF=float(PFEntry.get())
        
        
        
        
        
# ============================== SHOW .TXT ================================== #
    def get_resultados_potencia(self):
        self.dssText.Command = "Show powers kva elements" # Mostra potencia dos elementos

    def get_resultados_perdas(self):
        self.dssText.Command = "show losses" # Mostra as perdas nos elementos

    def get_resultados_tensoes_linha(self):
        self.dssText.Command = "Show Voltage  LL Nodes"
        
    def get_resultados_medidor(self):
        self.dssText.Command = "Show Meters"
        
# ============================FUNCOES CIRCUIT================================ #
    def get_nome_circuit(self):
        return self.dssCircuit.Name

    def get_potencias_circuit(self):
        p = -self.dssCircuit.TotalPower[0]
        q = -self.dssCircuit.TotalPower[1]
        return p, q

    def get_perda_total_circuit(self):
        p = -self.dssCircuit.Losses[0]
        q = -self.dssCircuit.Losses[1]
        return p, q

    def get_perda_linhas(self):
        p = -self.dssCircuit.LineLosses[0]
        q = -self.dssCircuit.LineLosses[1]
        return p, q


    def get_barras_elemento(self):
        barras = self.dssCktElement.BusNames
        barra1 = barras[0]
        barra2 = barras[1]
        return barra1, barra2

    def get_tensoes_elemento(self):
        return self.dssCktElement.VoltagesMagAng        
    
    def get_AllNodeVmagPUByPhase(self):
        VA    = self.dssCircuit.AllNodeVmagPUByPhase(1)
        VB    = self.dssCircuit.AllNodeVmagPUByPhase(2)
        VC    = self.dssCircuit.AllNodeVmagPUByPhase(3)
        return VA,VB,VC
    
    def get_AllBusVolts(self):
        ABV    = self.dssCircuit.AllBusVolts
#        AVM    = self.dssCircuit.AllBusVmag
        return ABV#,AVM
    
    def ativa_elemento(self, nome_elemento):
        # Ativa elemento pelo seu nome completo Tipo.Nome
        self.dssCircuit.SetActiveElement(nome_elemento)
        # retoma o nome do elemento ativado
        ABV    = self.dssCircuit.AllBusVolts
        xxx = self.dssCktElement.Name
        return ABV
    
    def get_transfromer_caracter(self):
        a=self.dssTransformer.AllNames()
        a = len(a)
        return a
# ============================== MONITORS ================================ #
    def get_MonitorsVoltages(self):
        ln=self.dssMonitors.Count
        M=self.dssMonitors.First
        #matriz de almacenamiento
        VV1=[];VV2=[];VV3=[]
        VA1=[];VA2=[];VA3=[]
        II1=[];II2=[];II3=[]
        IA1=[];IA2=[];IA3=[]
        
        PP1=[];PP2=[];PP3=[]

        l=0;k=0
        for i in range(0,ln):
            mode=self.dssMonitors.Mode
            if mode==1:
                
                print("modo1")
                print(self.dssMonitors.Name)
                P1=self.dssMonitors.Channel(1)
                Q1=self.dssMonitors.Channel(2)
                P2=self.dssMonitors.Channel(3)
                Q2=self.dssMonitors.Channel(4)
                P3=self.dssMonitors.Channel(5)
                Q3=self.dssMonitors.Channel(6)
              
                plt.title('potencia activa en:'+str(self.dssMonitors.Name))
                plt.plot(P1)
                plt.plot(P2)
                plt.plot(P3)
                plt.show()                           
                
                PP1.insert(l,P1)
                PP2.insert(l,P2)
                PP3.insert(l,P3)
                PP = PP1,PP2,PP3
                
                l +=1
                
            else:
                
                print("modo2")
                print(self.dssMonitors.Name)
                V1  =self.dssMonitors.Channel(1)
                AV1 =self.dssMonitors.Channel(2)
                V2  =self.dssMonitors.Channel(3)
                AV2 =self.dssMonitors.Channel(4)
                V3  =self.dssMonitors.Channel(5)
                AV3 =self.dssMonitors.Channel(6)
                
                I1  =self.dssMonitors.Channel(9)
                AI1 =self.dssMonitors.Channel(10)
                I2  =self.dssMonitors.Channel(11)
                AI2 =self.dssMonitors.Channel(12)
                I3  =self.dssMonitors.Channel(13)
                AI3 =self.dssMonitors.Channel(14)
                plt.title('Tension de fase en:'+str(self.dssMonitors.Name))
                plt.plot(V1)
                plt.plot(V2)
                plt.plot(V3)
                plt.show()
                
                plt.title('Corriente de fase en:'+str(self.dssMonitors.Name))
                plt.plot(I1)
                plt.plot(I2)
                plt.plot(I3)
                plt.show()
                
                VV1.insert(k,V1)
                VV2.insert(k,V2)
                VV3.insert(k,V3)
                VV = VV1,VV2,VV3

                VA1.insert(k,V1)
                VA2.insert(k,V2)
                VA3.insert(k,V3)
                VA = VA1,VA2,VA3
                
                II1.insert(k,I1)
                II2.insert(k,I2)
                II3.insert(k,I3)
                II=II1,II2,II3
                
                IA1.insert(k,IA1)
                IA2.insert(k,IA2)
                IA3.insert(k,IA3)
                IA=IA1,IA2,IA3                
                
                k +=1
                
#                VV.append[1,V1]
#                VV.append[2,V2]
#                VV.append[3,V3]
                
            nextM=self.dssMonitors.Next
        
        return  VV,VA,II,IA,PP              
   
    
    def get_Monitors(self):
        nm = self.dssMonitors.Count
        names = self.dssMonitors.AllNames
        ByteStream = self.dssMonitors.ByteStream
        ch1 = self.dssMonitors.Channel(1)
        ch2 = self.dssMonitors.Channel(2)
        ch3 = self.dssMonitors.Channel(3)
        ch4 = self.dssMonitors.Channel(4)
        ch5 = self.dssMonitors.Channel(5)
        ch6 = self.dssMonitors.Channel(6)
        
        plt.plot(ch1)
        plt.plot(ch3)
        plt.plot(ch5)
        plt.show()
        return ch1,ch2,ch3,ch4,ch5,ch6

        
# ==============================PLOT SECTION================================ #
    def plot_profile(self,a):
        VA    = self.dssCircuit.AllNodeVmagPUByPhase(1)
        DistA = self.dssCircuit.AllNodeDistancesByPhase(1)
        VB    = self.dssCircuit.AllNodeVmagPUByPhase(2)
        DistB = self.dssCircuit.AllNodeDistancesByPhase(2)
        VC    = self.dssCircuit.AllNodeVmagPUByPhase(3)
        DistC = self.dssCircuit.AllNodeDistancesByPhase(3)

        plt.subplot(1,1,1)
        plt.title('Phase Voltage Profile - MultLoad ='+str(a))
        plt.plot(DistA,VA,"k*",label = "VA")
        plt.plot(DistB,VB,"b+",label = "VB")
        plt.plot(DistC,VC,"gd",label = "VC")
        plt.ylabel('Voltage (pu)')
        plt.xlabel('Distance from EnergyMeter')
        plt.legend( ('Va', 'Vb', 'Vc'), loc = 'upper right')
        plt.grid()
        plt.show()
       
    def plot_power_loss(self,ListPerdp,ListPerdq,multload):
        plt.subplot(1,1,1)
        plt.title("Transformers' Active and Reactive Power Loss")
        plt.plot(ListPerdp,multload,"b+",label = "multload")
#        plt.plot(ListPerdq,multload,"k*",label = "multload")
        plt.grid()
        plt.ylabel('MulLoad')
        plt.xlabel('Power loss')
        plt.legend( ('Active Loss KW', 'Reactive Loss Kvar'), loc = 'lower right')
        plt.show() 
        
        plt.subplot(1,1,1)
        plt.title("Transformers' Active and Reactive Power Loss")
#        plt.plot(ListPerdp,multload,"b+",label = "multload")
        plt.plot(ListPerdq,multload,"k*",label = "multload")
        plt.grid()
        plt.ylabel('MulLoad')
        plt.xlabel('Power loss')
        plt.legend( ('Active Loss KW', 'Reactive Loss Kvar'), loc = 'lower right')
        plt.show()
        
# ========================SMART METER INFORMATION =========================== #
    def get_NewEnergyMeter(self,name):
        self.dssText.command ='New energymeter.'+str(name)+' element=transformer.'+str(name)+'kVA terminal=1'
        self.dssCircuit.SetActiveElement('energymeter.'+str(name))
        RV = self.dssMeters.RegisterValues
        TL = RV[23]
        return TL
        
    
    
    def get_summary_energy_meters_loss(self):
        n = self.dssMeters.Count
#        self.dssMeters.First
        name        = self.dssMeters.Name
        print(name)
        RV          = self.dssMeters.RegisterValues
        lossEnergy1 =np.array(RV)
        
        TL = lossEnergy1[23]
        return TL
        
        
        
        
    def get_summary_energy_meters(self):
        n = self.dssMeters.Count
        print('============================================================')
        print('           Cantidad de Medidores: '+str(n))
        print('============================================================')
        self.dssMeters.First
        Tloss=[]
        for i in range(0,n):
            name        = self.dssMeters.Name
            RV          = self.dssMeters.RegisterValues
            lossEnergy1 =np.array(RV)
            print('                    Meter Summary N°:'+str(i+1)+'  \n')
            print(u'Meter Name                          : '+str(name))
            print(u'System Energy                       : '+str(lossEnergy1[0]) +' kWh')
            print(u'Total Energy Loss System            : '+str(lossEnergy1[12])+' kWh') # zone loss energy
            print(u'Energy loss in the lines            : '+str(lossEnergy1[22])+' kWh')
            TL = lossEnergy1[23]
            print(u'Energy loss in the transformers     : '+str(TL)+' kWh \n' )
            print(u'% Transformers losses               : '+str(lossEnergy1[23]*100/lossEnergy1[0])+' %  \n' )
            print('============================================================')
            self.dssMeters.Next
            Tloss.append(TL)
        return Tloss,n
    
# ========================TRANSForMERs INFORMATION =========================== #
    def get_summary_transformer(self):
        n = self.dssTransformer.Count
        print('============================================================')
        print('           Cantidad de transformadores: '+str(n))
        print('============================================================')
        self.dssTransformer.First

        for i in range(0,n):
            name = self.dssTransformer.Name
            kva  = self.dssTransformer.kva
            R    = self.dssTransformer.R
            Xhl  = self.dssTransformer.Xhl
            
            print('Nombre     : '+str(name))
            print('Potencia   : '+str(kva))
            print('Resistencia: '+str(R))
            print('Reactancia : '+str(Xhl))
            print('============================================================')
            self.dssTransformer.Next
            
#
#        return n,n1,r1,n2,r2,Xhl
    
# =========================================================================== #
#                           Elementary_Functions                              #
# =========================================================================== #

#-----------------  Convertir cartesianas a polares --------------------------#
def cart2pol(x, y):
    rho = np.sqrt(x**2 + y**2)
    phi = np.arctan2(y, x)
    phi = math.degrees(phi)
    return(rho, phi)
#-----------------  Convertir polares a cartesianas --------------------------#
def pol2cart(rho, phi):
    x = rho * np.cos(phi)
    y = rho * np.sin(phi)
    return(x, y)
    
# =========================================================================== #
#                           Read_Load_Caracteristic                           #
# =========================================================================== #

    
def get_AnalisisDeCarga(LoadShape):
    # se requiere el nombre del archivo loadshape 'fecha'
    datos = pd.read_csv(LoadShape,header=-1)
    dato  = np.array(datos)
    ln    = len(dato)
    Dmax  = max(dato)
    Etotal= sum(dato)/float(4)
    Dmed  = Etotal/(ln/96*24)
    #factor de carga do transformador 
    Fc    = Dmed/Dmax
    print('dias evaluados       : '+str(ln/96))
    print('Demanda Maxima[kW]   : '+str(Dmax) )
    print('Energia total [kWh]  : '+str(Etotal))
    print('Demanda media [kW]   : '+str(Dmed))
    print('Factor de carga trafo: '+str(Fc))
    plt.plot(dato)
    plt.show() 


def get_SumaDeDemanda(Dem1,Dem2,Dem3):
    dato1  = np.array(Dem1)
    dato2  = np.array(Dem2)
    dato3  = np.array(Dem3)
    ln1 = len(Dem1)
    c  = []
    for i in range(0,ln1):
        cc=dato1[i]+dato2[i] +dato3[i]
        c.append(cc)
    datos=c    
    get_AnalisisDeCarga(datos)
    return 

# =========================================================================== #
#                           Transformer_information                           #
# =========================================================================== #
    
def get_TransformerInformation(objeto):
    
    n1,v1,n2,v2 = objeto.get_summary_energy_meters()
    lossEnergy1 =np.array(v1)
    lossEnergy2 =np.array(v2)    
    print('                            Meter 1 Summary                     \n')
    print(u'Activated Meter                     : '+n1) # last meter
#    print(u'Meter Name                          : '+objeto.get_nome_medidor())
    print(u'System Energy                       : '+str(lossEnergy1[0]) +' kWh')
    print(u'Total Energy Loss System            : '+str(lossEnergy1[12])+' kWh') # zone loss energy
    print(u'Energy loss in the lines            : '+str(lossEnergy1[22])+' kWh')
    print(u'Energy loss in the transformers     : '+str(lossEnergy1[23])+' kWh \n' )
    print(u'% Transformers losses               : '+str(lossEnergy1[23]*100/lossEnergy1[0])+' %  \n' )
    print('                            Meter 2 Summary                          \n')
    print(u'Activated Meter                     : '+n2) # last meter
#    print(u'Meter Name                          : '+objeto.get_nome_medidor())
    print(u'System Energy                       : '+str(lossEnergy2[0]) +' kWh')
    print(u'Total Energy Loss System            : '+str(lossEnergy2[12])+' kWh' ) # zone loss energy
    print(u'Energy loss in the lines            : '+str(lossEnergy2[22])+' kWh' )
    print(u'Energy loss in the transformers     : '+str(lossEnergy2[23])+' kWh  \n' )
    print(u'% Transformers losses               : '+str(lossEnergy2[23]*100/lossEnergy2[0])+' %  \n' ) 
    
# =========================================================================== #
#                           Geral information                                 #
# =========================================================================== #
def get_GeralInformation(objeto,lim_number,stepsize,number,LoadMult):
    
    va=[];vb=[];vc=[];VA=[];VB=[];VC=[];
    AA=[];AB=[];AC=[];aa=[];ab=[];ac=[];

    for i in range(0,lim_number):
        print('Flujo N°: '+str(i+1))
        objeto.compila_DSS()     
        objeto.solve_DSS_daily(LoadMult,stepsize,number)    # solution mode

#        magnitud y angulo de tension
        ABV= objeto.get_AllBusVolts()
#        tension y angulos dos barramentos [tenison x angulo]
        ln= len(ABV)
        abv = np.reshape(ABV,(ln/2,2))                                         # reordena la matriz en 2 columnas [magnitud y angulo]
       
#        tensiones en el tiempo
#        tensiones y angulos en alta  tension
        VA.append(abv[0][0]);  VB.append(abv[1][0]);  VC.append(abv[2][0])
        AA.append(abv[0][1]);  AB.append(abv[1][1]);  AC.append(abv[2][1])
#        tesiones y angulos en baja tension
        va.append(abv[3][0]);  vb.append(abv[4][0]);  vc.append(abv[5][0])
        aa.append(abv[3][1]);  ab.append(abv[4][1]);  ac.append(abv[5][1])
        number +=1
    
    print('Tus pinches resultados son ...\n')
        
    # tensiones de alta ##############################    
    ln = len(VA)
#        convertir cartesiano a polinomio
    VVA=[]
    for i in range(0,ln):
        b=cart2pol(VA[i],AA[i])
        VVA.append(b)
#    VVA=np.array(VV)
    print(VVA)
    
    VVB=[]
    for i in range(0,ln):
        b=cart2pol(VB[i],AB[i])
        VVB.append(b)
#    VVB=np.array(VV)
    print(VVB)

    VVC=[]
    for i in range(0,ln):
        b=cart2pol(VC[i],AC[i])
        VVC.append(b)
#    VVC=np.array(VV)
    print(VVC)
    
    
    # tensiones de baja #########################################
    ln = len(VA)
#        convertir cartesiano a polinomio
    VVA=[]
    for i in range(0,ln):
        b=cart2pol(VA[i],AA[i])
        VVA.append(b)
#    VVA=np.array(VV)
    print(VVA)
    
    VVB=[]
    for i in range(0,ln):
        b=cart2pol(VB[i],AB[i])
        VVB.append(b)
#    VVB=np.array(VV)
    print(VVB)

    VVC=[]
    for i in range(0,ln):
        b=cart2pol(VC[i],AC[i])
        VVC.append(b)
#    VVC=np.array(VV)
    print(VVC)
   
    plt.plot(VA)
#    plt.plot(VB)
#    plt.plot(VC)    
#    plt.plot(va)
#    plt.plot(vb)
#    plt.plot(vc)
    plt.grid()
    plt.ylabel('Tensao')
    plt.xlabel('time')
    plt.legend( ('VA', 'VB','VC','va', 'vb','vc'), loc = 'lower right')
    plt.show() 
    
    VH = [VA,VB,VC]
    VL = [va,vb,vc]
    return VH,VL,ABV
# =========================================================================== #
#          Tensiones y corrientes de fase de cada transfromador
# =========================================================================== #
def get_TensionesCorrientesPerdidas(objeto):
    
    VV,VA,II,IA,PP= objeto.get_MonitorsVoltages()
    print('corrientes de transfromador')
    #Corrientes de baja tension
    Iprim_a =II[0][1]
    Iprim_b =II[1][1]
    Iprim_c =II[2][1]

    Isec_a  =II[0][0]
    Isec_b  =II[1][0]
    Isec_c  =II[2][0]
    
    plt.plot(Iprim_a)
    plt.plot(Iprim_b)
    plt.plot(Iprim_c)

    plt.plot(Isec_a)
    plt.plot(Isec_b)
    plt.plot(Isec_c)
    plt.show()

    print(Iprim_a)
    
    # RESISTENCIAS DE ACORDO OPENDSS
    # RESITENCIA EN ALTA TENSION
    r1=4.908865066 #3.3987
    # RESISTENCIA DE BAJA TENSION
    r2=0.001161
    
    #RESISTENCIAS DE ACUERDO A ENSAYOS 
    #AÑO 2016
    # RESISTENCIAS OHMICAS MEDIDAS DO TRANSFORMADOR -2016
    # ALTA TENSION 
    H1H2=2.7
    H1H3=2.7
    H2H3=4.738 #2.7 año2016 - 4.738 año 2018
    # BAJA TENSION
    X1X2=0.91e-3
    X1X3=0.91e-3
    X2X3=0.89e-3

    # FORMULAS DE CALCULO DE RESISTENCIAS DE FASE DO TRANSFORMADOR
    # RESISTENCIA DE BAJA TENSION
    ra = 0.5*(X1X2+X1X3-X2X3)
    rb = 0.5*(X2X3+X1X2-X1X3)
    rc = 0.5*(X2X3+X1X3-X2X3)
    # RESISTENCIA EN ALTA TENSION
    RA = ((-H1H3-H1H2+H2H3)**2-4*H1H3*H1H2)/(2*( H1H3-H1H2-H2H3))
    RB = ((-H1H3-H1H2+H2H3)**2-4*H1H3*H1H2)/(2*(-H1H3+H1H2-H2H3))
    RC = ((-H1H3-H1H2+H2H3)**2-4*H1H3*H1H2)/(2*(-H1H3-H1H2+H2H3))
    
    # CORIENTES DE BAJA TENSION
    ia=np.array(Isec_a); ib=np.array(Isec_b); ic=np.array(Isec_c)
    # CORRIENTES DE ALTA TENSION 
    Ia=np.array(Iprim_a); Ib=np.array(Iprim_b); Ic=np.array(Iprim_c)
    
    # PERDAS DE POTENCIA EN LAS BOBINAS
    # BAJA TENSION
    Pcua=sum(ia**2)*ra/4; Pcub=sum(ib**2)*rb/4; Pcuc=sum(ic**2)*rc/4
    # ALTA TENSION
    PcuA=sum(Ia**2)*RA/4; PcuB=sum(Ib**2)*RB/4; PcuC=sum(Ic**2)*RC/4
    
    #PERDA TOTAL 
    pp = Pcua +Pcub +Pcuc; PP = PcuA +PcuB +PcuC
    
    PTOTAL = (pp+PP)/1000
    print('\n')
    print('Perda total no cobre : '+str(PTOTAL)+'[kWh]')
    
    print('Resistencias en alta tension: ')
    print(RA,RB,RC)

    print('Resistencias en baja tension:')
    print(ra,rb,rc) 

# =========================================================================== #
#             ESCOGER MEJOR TAMAÑO DE TRANSFORMADOR
# =========================================================================== #
    #
def get_BestTransformer1(objeto,LoadMult,stepsize,number,Trafos):
    
#    # comparacion de transfromadores con menor perdida
#    p=get_BestTransformer(objeto,LoadMult,stepsize,number)
    #trafos\ S  \%loadloss\%noladloss   ANEEL
#    Tf45  =[45  , 1.6667 , 0.4333]
#    Tf75  =[75  , 1.4667 , 0.3933]
#    Tf112 =[112 , 1.3333 , 0.3467]
#    Tf150 =[150 , 1.2333 , 0.3233]
#    Tf225 =[225 , 1.1600 , 0.2889]
#    Tf300 =[300 , 1.0833 , 0.2700]
##    Tf500 =[500 , 1.2000 , 0.24  ] #iran
#    Tf500 =[500 , 1.2532 , 0.2286]
#    Trafos = [Tf45,Tf75,Tf112,Tf150,Tf225,Tf300,Tf500]
    
    nt = len(Trafos)                # Numero de transformadores alternos 
    Matriz=[]                       # Matriz de trafos con info OPENDSS
    for i in range(0,nt):
        Matriz.append([])        
    # Transformar perdidas en % - de acuerdo datos necesarios OPENDSS
    for i in range(0,nt):                                                       # 
        S         = Trafos[i][0]
        noloadloss = Trafos[i][1]/float(S)/1000*100
        loadloss   = Trafos[i][2]/float(S)/1000*100
        #Nueva matriz compuesta por S|%loadlosses|%noloadlosses
        Matriz[i].append(S); Matriz[i].append(loadloss); Matriz[i].append(noloadloss)      
#    print(Matriz)
    Trafos=Matriz                   # Matriz de trafos con sus perdidas - OPENDSS                                                
    # Compilar archivo DSS
    objeto.compila_DSS()
    objeto.solve_DSS_daily(LoadMult,stepsize,number)
    
    t = objeto.dssTransformer.Count # transfromadores en la rede DSS
      
    Perdas=[]                       #crear lista de almacenamiento de perdidas  
    for i in range(0,t):            #espacio para cada transfromador de rede
        Perdas.append([])

    l=1   
    # Reemplazo de cada transfromador de rede por transformador alterno
    for k in range(0,nt):           #para cada transformador alterno
        print('Iteracion... '+str(l))
        S = Trafos[k][0]; loadloss= Trafos[k][1]; noloadloss =Trafos[k][2]
        objeto.compila_DSS()
        objeto.solve_DSS_daily(LoadMult,stepsize,number)
        objeto.dssTransformer.First     #primer trafo
        for i in range(0,t):            #para cada transfromador conectado a la red    
            name=objeto.dssTransformer.Name
#            print(name)
            objeto.get_ChangeTranformer(name,S,loadloss,noloadloss)
    #            segundo trafo
            objeto.dssTransformer.Next                                         
        #flujo diario
        objeto.solve_DSS_daily(LoadMult,stepsize,number)
         
        objeto.dssMeters.First
        for i in range(0,t):           
        # mediciones
            name = objeto.dssMeters.Name#nombre del medidor del transformador reemplazado
            RV   = objeto.dssMeters.RegisterValues
            lossEnergy1 =np.array(RV)
            LE = lossEnergy1[23]
#            print(TL1)                  #imprime perdida de trafo cada medidor    
            objeto.dssMeters.Next
            Perdas[i].append(LE)
#            print(name)
            l+=1                        #contador de iteracion
#    print(Perdas)                       #imprime perdidas de cada trafo de rede para cada trafo alterno
    
    #===================== PLOT PERDIDAS=======================================
    objeto.dssTransformer.First         #primer trafo de la red
    for i in range(0,t):
        name=objeto.dssTransformer.Name
        print('==============================================')
        print('Transformador evaluado:'+str(name))
        fig = plt.figure(u'Gráfica de PERIDAS') # Figure
        ax = fig.add_subplot(111) # Axes        
        nombres = ['45','75','112','150','225','300','500']
        if nt==7:                       #plot para N trafos alternos =7
            nombres= ['45','75','112','150','225','300','500']
        else:                           # N trafos alternos =9
            nombres= ['15','30','45','75','112','150','225','300','500']
        datos = Perdas[i]               #cada lista i incluye conjunto de perdidas para cada trafo de rede
        xx = range(len(datos))
        ax.bar(xx, datos, width=0.8, align='center')
        ax.set_xticks(xx)
        ax.set_xticklabels(nombres)
        plt.axhline(min(Perdas[i]), color='r', xmax=1)
        plt.axhspan(min(Perdas[i]),max(Perdas[i]), alpha=0.3, color='y')
        plt.xlabel('Transformers [kVA]')
        plt.ylabel('Energy losses [kWh]')
        plt.title('Transformer Evaluated: '+ str(name))
#        print('success')
        plt.show()
        
        objeto.dssTransformer.Next

# =========================================================================== #
#             rgrafica de carga 
# =========================================================================== #
def get_LoadSummary(self):
    
    datos   = pd.read_excel('meter_2018.xlsx')
    dato = datos['kVA tot mean']
    date = datos['Date']
    dato=np.array(dato)
    ts = pd.Series(dato, index=date)
    # analise de acordo as fechas 
#D: Día natural
#B: Día hábil
#W: semanal (último día de la semana)
#M: mensual (último día del mes)
#SM: bimensual (el día 15 y último del mes)
#Q: trimestral (último día del trimestre)
    me=ts.resample('D').mean()
    ma=ts.resample('B').mean()
    #ma=ts.resample('D').max()
    
    plt.plot(me)#
    plt.plot(ma)#


# =========================================================================== #
#             PERDIDA DE VIDA DEL TRANSFORMADOR
# =========================================================================== #
    #
def get_TransformerLossLife(self):
#    cacula a perda de vida do transformador de acordo con a temperatura pico do bobinado
#    Arrhenius
    Th = 110        # is the winding hottest-spot temperature, °C
#    parametros constantes 
    A = 9.8e-18
    B = 15000
    
    TLL = A*math.exp(B/(Th+273))
    
    # factor de aceleramiento del transfromador 
    return TLL


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    