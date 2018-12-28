import pandas as pd
from numpy import *
import datetime as dt
import time
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import random as rd
import xlsxwriter as xlwt
import os
import monthdelta as md
from bisect import bisect
from tkinter import Tk
from tkinter.filedialog import askopenfilename

#Filtrar warnings
import warnings
warnings.filterwarnings("ignore")




#PARAMETROS INICIALES
##Tk().withdraw() 
##directory = askopenfilename(initialdir = os.path.dirname(os.path.realpath(__file__)), title = "Select file", filetypes = (("xlsx files","*.xlsx"),("all files","*.*"))) #directorio archivo
##fullname=os.path.basename(directory)#nombre archivo con extension
fullname="data_clp.xlsx"
filename=os.path.splitext(fullname)[0] #nombre archivo sin extension
ext=os.path.splitext(fullname)[1] #extension

if(filename not in ['data_clp','data_uf']):
    print("Archivo de datos incorrecto")
    exit() 

excel_data = filename+ext
N=100#int(input("Número de simulaciones? "));


#MOSTRAR HORA DE INICIO
print("Hora inicio: ",time.strftime("%H:%M:%S"))
print("")





#LEER Y PREPROCESAR DATOS
xlsx_df = pd.read_excel(excel_data) #dataframe
codigos_todos=list(xlsx_df["Nemo"]) #todos los instrumentos disponibles

#Extraer columnas (Nemo | fechas)
Fechas=list(xlsx_df) 
Fechas.pop(0) #eliminar "Nemo" de lista

#Todas las TIR
TIR_todos = [0]*len(Fechas) #TIR por fecha (incluye vacios)
for i in range(len(TIR_todos)):
    TIR_todos[i] = list(xlsx_df[Fechas[i]]) #tir asociadas (incluye vacios)

#Pasar datetime a date
for i in range(len(Fechas)):
    Fechas[i]=dt.date(Fechas[i].year,Fechas[i].month,Fechas[i].day)

#TPM, IPC y tasa real
TPM=[0]*len(Fechas)
IPCm=[0]*len(Fechas)
IPCa=[0]*len(Fechas)
TR=[0]*len(Fechas)
t=1/360
for i in range(len(Fechas)):
    TPM[i]=TIR_todos[i][0] #guardar TPM
    IPCm[i]=TIR_todos[i][1] #guardar variacion IPC mensual
    IPCa[i]=TIR_todos[i][2] #guardar variacion IPC 12 meses
    if(filename[-2:]=="uf"):
        TR[i]= ((1+(TPM[i]/100)*t)/(1+(12*IPCm[i]/100)*t)-1)*(100/t) #guardar tasa real (caso uf)
    else:
        TR[i]=TPM[i]
    TIR_todos[i].pop(0); TIR_todos[i].pop(0); TIR_todos[i].pop(0)#quitar TPM e IPC de lista TIR


#Quitar TPM e IPC de lista de codigos
codigos_todos.pop(0);codigos_todos.pop(0);codigos_todos.pop(0)

#Matrices para curvas cero, factores de descuento, parametros y RMSE
T_zero=[] 
for i in range(30*2):
    i=i+1
    T_zero.append(i/2)
zero_SV=[([0]*len(T_zero)) for n in range(len(Fechas))] #Tasas cero cupon
FD_SV=[([0]*len(T_zero)) for n in range(len(Fechas))] #FD cero cupon
param_SV=[0]*len(Fechas) #parametros svensson
RMSE_SV=[0]*len(Fechas) #RES2






#FUNCIONES
#Funcion objetivo (minimizar diferencia de precios)
def fun(p, T, TT, ct, tera, y): #p=[b0,b1,b2,b3,lamb1,lamb2], T=maturity, TT=tiempo cuotas, ct=cuotas, tera=teras y y=precios
    e=0
    for i in range(len(T)):
        sumatoria=0;
        for j in range(len(TT[i])):
            t_sv=p[0]+p[1]*((1-exp(-p[4]*TT[i][j]))/(p[4]*TT[i][j])) + p[2]*(((1-exp(-p[4]*TT[i][j]))/(p[4]*TT[i][j]))-exp(-p[4]*TT[i][j])) + p[3]*(((1-exp(-p[5]*TT[i][j]))/(p[5]*TT[i][j]))-exp(-p[5]*TT[i][j]))
            sumatoria=(ct[i][j]*exp(-t_sv*TT[i][j]/100)) + sumatoria; #composicion continua
        precio=sumatoria/tera[i]*100
        e =(precio-y[i])**2 + e
    return e

#Calcular tir asociada a partir del precio
def calcular_tir(tir, T, TT, ct, tera, y):
    e=0
    for i in range(len(T)):
        sumatoria=0;
        for j in range(len(TT[i])):
            sumatoria=ct[i][j]/((1+tir[i]/100)**TT[i][j]) + sumatoria; #composicion semestral
        precio=sumatoria/tera[i]*100
        e =(precio-y[i])**2 + e
    return e

#Interpolar tasas
def interpolar(Tm,Ri,Rf,Ti,Tf):
    return Ri+(Rf-Ri)*(Tm-Ti)/(Tf-Ti) 






#VALORAR BONOS MERCADO
last_tir=[0]*len(Fechas)
for k in range(len(Fechas)):
    codigos = [] #codigos de la fecha
    TIR = [] #tir de la fecha
    for i in range(len(codigos_todos)): #limpiar vacios
        if (isnan(TIR_todos[k][i])==False):
            codigos.append(codigos_todos[i])
            TIR.append(TIR_todos[k][i])
    last_tir[k]=TIR[-1] #TIR del ultimo instrumento disponible
    
    TIR_SV=[0]*len(codigos) #tir estimadas
    precios=[0]*len(codigos) #precios mercado
    precios_SV=[0]*len(codigos) #precios estimados
    T=[0]*len(codigos) #madureces
    TT=[0]*len(codigos) #tiempo cortes de cupon
    Teras=[0]*len(codigos) #valores tera
    Cuotas=[0]*len(codigos) #cuotas (flujos)

    FechaValoracion = Fechas[k]
    
    for j in range(len(codigos)):
        codigo=codigos[j]
        tir=TIR[j]

        #Fecha de termino
        if (codigo[0:3]=="SWP"):
            FechaTermino=FechaValoracion+md.monthdelta(int(codigo[5:7]))
        else:
            if (codigo=="BTU0451023"):
                DD=15;
            else:
                DD=1;
                
            MM=int(codigo[6:8])
            YY=2000+int(codigo[8:10])
            FechaTermino = dt.date(YY, MM, DD)

            #Cuota BCP/BCU
            TasaEmision = float(codigo[3:6])/1000
            Cuota=TasaEmision/2;
        
        #Madurez de los bonos
        T[j]=((FechaTermino-FechaValoracion).days/365)

        #Definir el plazo en años para el vencimiento
        PlazoAnos=(FechaTermino-FechaValoracion).days/365
        Entero=int(PlazoAnos)
        Resto=PlazoAnos-Entero

        if Resto>0.5:
            Resto=1
        else:
            Resto=0.5
        PlazoAnos=Entero+Resto;
        PlazoSemestre=int(PlazoAnos*2)

        #Crear matriz de datos para valorar bonos
        data_precio=[[0]*2 for n in range(PlazoSemestre)]

        #Columna 1: Fechas
        data_precio[PlazoSemestre-1][0]=FechaTermino
        
        #Ingresar los datos de las fechas en la matriz
        for i in range(PlazoSemestre-1):
            FechaTermino=FechaTermino+md.monthdelta(-6)
            data_precio[PlazoSemestre-i-2][0]=FechaTermino
            
        #Fijar fecha del flujo anterior para el VP TERA
        FechaFlujoAnterior = FechaTermino+md.monthdelta(-6)

        #Ingresar los plazos entre fechas a la matriz
        #cero cupon solo tiene 1 cuota
        if (codigo[0:3]=="SWP"): #swaps (cero cupon)
            TasaEmision=0
            Cuota=tir*T[j] #cuota cupon (menor o igual a 1 año)
            Cuotas[j]=[Cuota+100] #nocional + cupon
            TT[j]=[T[j]] #cero cupon
            TIR[j]=(((100+Cuota)/100)**(1/T[j])-1)*100
            tir=TIR[j]
            sumatoria_precio=(100+Cuota)/((1+tir/100)**T[j])
            
        else:
            #Columna 2: Ingresar flujos
            Cuotas[j]=[0]*PlazoSemestre
            for i in range(PlazoSemestre):
                data_precio[i][1]=Cuota*100
                Cuotas[j][i]=Cuota*100

            #Ingresar el nocional
            data_precio[PlazoSemestre-1][1]+=100
            Cuotas[j][PlazoSemestre-1]+=100
            
            #Calcular plazos de corte de cupon y valor presente del BCP/BCU
            TT[j]=[0]*PlazoSemestre
            sumatoria_precio=0;
            for i in range(PlazoSemestre):
                TT[j][i]=((data_precio[i][0]-FechaValoracion).days/365)
                sumatoria_precio=data_precio[i][1]/(1+(tir/100))**TT[j][i]+sumatoria_precio

        #Calculo tasa TERA y VP TERA
        TERA=(1+TasaEmision)**(365/360)-1;
        ValorTERA=100*(1+TERA)**((FechaValoracion-FechaFlujoAnterior).days/365)
        Teras[j]=ValorTERA

        #Calcular valor bono
        precios[j]=sumatoria_precio/ValorTERA*100
    




    #OPTIMIZAR PRECIOS
    #crear matriz parametros iniciales
    param=[0]*6 #b0, b1, b2, b3, lambda1, lambda2
    #crear matriz de precios calculados
    P_SV_aux=[];
    #crear matriz de error^2
    RES2=[0]*len(codigos);RES2_aux=[0]*len(codigos);
    #crear matriz de suma de error^2
    sum_RES2=100000;sum_RES2_aux=10000;

    #simular N veces
    #parametros iniciales random para cada simulacion (para uf)
    rr=30
    b0=last_tir[k]
    b1=TR[k]-b0
    b2=(2*3.6-(2.3+4.5)-0.00053*b1)/0.37 #(2y(2años)-(y(3meses)+y(10años)-0.00053*b1)/(0.37)
    if(filename[-2:]=="uf"):
        uu=3;dd=1;
    else:
        uu=2;dd=0.5;

    
    for n in range(N):
        #b0, b1, b2, b3, lambda1, lambda2
        param_aux=[rd.uniform(b0-dd,b0+uu), 
                   rd.uniform(b1-2.5,b1+2.5),
                   rd.uniform(b2-5,b2+5),
                   rd.uniform(b2-5,b2+5),
                   rd.uniform(0,1/rr),
                   rd.uniform(0,1/rr)]

        x = array(T)  #madurez
        xx = array(TT) #tiempos de evaluacion
        ct = array(Cuotas) #cuotas
        tera= array(Teras) #valores tera
        y = array(precios)  #solo precios disponibles
        
        p0 = array(param_aux)  # parametros iniciales

        #restricciones de b0+b1=tasa real y lambda1,lambda2>=0      
        cons = ({'type': 'eq','fun': lambda params : params[0]+params[1]-TR[k]})#,{'type': 'ineq','fun': lambda params : params[4]},{'type': 'ineq','fun': lambda params : params[5]}

        bnds = ((last_tir[k]-dd,last_tir[k]+uu),(None,None),(b2-8,b2+8),(b2-8,b2+8),(0,15),(0,15)) 

        p = minimize(fun, p0, args=(x,xx,ct,tera,y),method='SLSQP',constraints=cons,bounds=bnds)

        
        c=[0]*6 #parametros Svensson
        for i in range(len(p.x)):
            c[i]=p.x[i]

        P_SV_aux=[] #lista para guardar valores del precio SV
        for i in range(len(T)):  #calcular los precios con SV
            sumatoria=0
            for j in range(len(TT[i])):
                svv=c[0]+c[1]*((1-exp(-c[4]*TT[i][j]))/(c[4]*TT[i][j]))+c[2]*(((1-exp(-c[4]*TT[i][j]))/(c[4]*TT[i][j]))-exp(-c[4]*TT[i][j]))+c[3]*(((1-exp(-c[5]*TT[i][j]))/(c[5]*TT[i][j]))-exp(-c[5]*TT[i][j]))
                sumatoria=Cuotas[i][j]*exp(-svv*TT[i][j]/100)+sumatoria #continua
            precio=sumatoria/Teras[i]*100
            P_SV_aux.append(precio)

        #calcular residuos^2
        RES2_aux=[0]*len(T);
        for i in range(len(T)):
            RES2_aux[i]=(P_SV_aux[i]-precios[i])**2
        sum_RES2_aux=sum(RES2_aux)
       
        #si tiene mejores resultados, guardar valores
        if (sum_RES2_aux<sum_RES2):
            param=c

            #calcular tir en composicion semestral
            t0 = array([3]*len(T))  # tasas iniciales
            y = array(P_SV_aux) # precios SV
            p = minimize(calcular_tir, t0, args=(x,xx,ct,tera,y),method='SLSQP')
            TIR_SV=p.x
            
            precios_SV=P_SV_aux
            RES2=RES2_aux
            sum_RES2=sum_RES2_aux

            RMSE_TIR=0
            for i in range(len(TIR_SV)):
                RMSE_TIR+=(TIR_SV[i]-TIR[i])**2
                
            print("Cambiaron parámetros en iteración ",n+1)
            print("RMSE TIR: ",sqrt(RMSE_TIR/len(TIR_SV)))
    param_SV[k]=param; #guardar parametros
    RMSE_SV[k]=sqrt(sum_RES2/len(RES2)) #RMSE


    #CURVA CERO CUPON Y FD
    c=param
    for i in range(len(T_zero)):
        #comp 365
        svv=c[0]+c[1]*((1-exp(-c[4]*T_zero[i]))/(c[4]*T_zero[i]))+c[2]*(((1-exp(-c[4]*T_zero[i]))/(c[4]*T_zero[i]))-exp(-c[4]*T_zero[i]))+c[3]*(((1-exp(-c[5]*T_zero[i]))/(c[5]*T_zero[i]))-exp(-c[5]*T_zero[i]))
        FD=exp(-svv*T_zero[i]/100) #factor descuento
        FD_SV[k][i]=FD
        zero_SV[k][i]=((1/FD)**(1/T_zero[i])-1)*100 #tasa cero en comp/365   


    #IMPRIMIR RESULTADOS EN CONSOLA
    print("Fecha: ",Fechas[k])
    print("RMSE TIR: ",sqrt(RMSE_TIR/len(TIR_SV)))
    print("");




#GUARDAR RESULTADOS EN EXCEL
filaname_results=filename+' results (SV) ['+str(Fechas[len(Fechas)-1])+' - '+str(Fechas[0])+'].xlsx'  #donde se guardan los resultados 
workbook = xlwt.Workbook(filaname_results)

#formatos celdas
bold=workbook.add_format({'bold': True}) #titulos negrita
decimales2 = workbook.add_format() 
decimales2.set_num_format('0.00') #formato numero, 2 decimales
decimales3 = workbook.add_format() 
decimales3.set_num_format('0.000') #formato numero, 3 decimales


#Hojas 
wsheet1 = workbook.add_worksheet("Curvas cero")
wsheet1.write(0,0,"Fecha",bold)
wsheet2 = workbook.add_worksheet("FD")
wsheet2.write(0,0,"Fecha",bold)
wsheet3 = workbook.add_worksheet("Betas")
wsheet3.write(0,0,"Fecha",bold)

#datos plazos, curva cero y FD
for i in range(len(Fechas)):
    ii=len(Fechas)-1-i
    fecha_str=Fechas[i].strftime('%d-%m-%Y')
    wsheet1.write(ii+1,0,fecha_str,bold)
    wsheet2.write(ii+1,0,fecha_str,bold)
    wsheet3.write(ii+1,0,fecha_str,bold)
    for j in range(len(T_zero)):
        wsheet1.write(0,j+1,T_zero[j],decimales2)
        wsheet1.write(ii+1,j+1,zero_SV[i][j],decimales3)
        wsheet2.write(0,j+1,T_zero[j],decimales2)
        wsheet2.write(ii+1,j+1,FD_SV[i][j],decimales3)

#datos betas
wsheet3.write(0,1,"b0",bold)
wsheet3.write(0,2,"b1",bold)
wsheet3.write(0,3,"b2",bold)
wsheet3.write(0,4,"b3",bold)
wsheet3.write(0,5,"lambda1",bold)
wsheet3.write(0,6,"lambda2",bold)
wsheet3.write(0,7,"RMSE",bold)
for i in range(len(Fechas)):
    ii=len(Fechas)-1-i
    wsheet3.write(ii+1,1,param_SV[i][0],decimales3)
    wsheet3.write(ii+1,2,param_SV[i][1],decimales3)
    wsheet3.write(ii+1,3,param_SV[i][2],decimales3)
    wsheet3.write(ii+1,4,param_SV[i][3],decimales3)
    wsheet3.write(ii+1,5,param_SV[i][4],decimales3)
    wsheet3.write(ii+1,6,param_SV[i][5],decimales3)
    wsheet3.write(ii+1,7,RMSE_SV[i],decimales3)
    
workbook.close()

print("")
print("")
print("Resultados guardados en archivo: "+filaname_results)


#MOSTRAR HORA DE TÉRMINO
print("")
print("Hora término: ",time.strftime("%H:%M:%S"))

    


