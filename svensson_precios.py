import pandas as pd
from numpy import *
import datetime as dt
import time
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import random as rd
import xlsxwriter as xlwt
from win32com.client import Dispatch
from os import getcwd
import monthdelta as md
from bisect import bisect

#Filtrar warnings
import warnings
warnings.filterwarnings("ignore")

def Optimizar_precio(excel_data,N):
    #ejecutar macro para actualizar datos
    cd=getcwd()+"\\" + excel_data #localizacion del excel
    xl = Dispatch("Excel.Application")
    xl.Visible = False
    wb=xl.Workbooks.Open(cd)
    xl.Run("Macro1") #copiar datos
    xl.Run("Macro2") #ordenar datos por madurez
    xl.Run("Macro3") #limpiar datos sin tasa
    wb.Save()
    wb.Close()
    xl.Quit()

    xlsx_df = pd.read_excel(excel_data) #dataframe

    codigos=list(xlsx_df["Nemo"]) #codigos instrumentos disponibles
    xlsx = list(xlsx_df["Tir BCS"]) #tasas tir
    precios=[0]*len(codigos) #precios mercado
    precios_SV=[0]*len(codigos) #precios estimados
    T=[0]*len(codigos) #madureces
    TT=[0]*len(codigos) #tiempo cortes de cupon
    Teras=[0]*len(codigos) #valores tera
    Cuotas=[0]*len(codigos) #cuotas (flujos)

    #VALORAR BONO ORIGINAL
    for j in range(len(codigos)):
        codigo=codigos[j]
        TIR=xlsx[j]
        
        if (codigo=="BTU0451023"):
            DD=15;
        else:
            DD=1;

        #fecha de termino
        MM=int(codigo[6:8])
        YY=2000+int(codigo[8:10])
        FechaTermino = dt.date(YY, MM, DD)

        #fecha de valoracion hoy
        #FechaValoracion = dt.date.today()
        FechaValoracion = dt.date(2018,9,28) #test datos 28-09-2018
        
        #madurez de los bonos
        T[j]=((FechaTermino-FechaValoracion).days/365)

        #calculamos cuota BCP/BCU
        TasaEmision = float(codigo[3:6])/1000
        Cuota=TasaEmision/2;

        #definimos el plazo en años para el vencimiento
        PlazoAnos=(FechaTermino-FechaValoracion).days/365
        PlazoInicial=PlazoAnos
        PlazoAnos=round(PlazoAnos,2)
        Entero=int(PlazoAnos)
        Resto=PlazoAnos-Entero

        if Resto>0.5:
            Resto=1
        else:
            Resto=0.5
        PlazoAnos=Entero+Resto;
        PlazoSemestre=int(PlazoAnos*2)

        #Creamos matriz de datos
        data_precio=[[0]*4 for n in range(PlazoSemestre)] #matriz de precios 
        data_estimada=[[0]*4 for n in range(PlazoSemestre)] #matriz de precios estimados

        #Fila 1: Fechas
        data_precio[PlazoSemestre-1][0]=FechaTermino
        data_estimada[PlazoSemestre-1][0]=FechaTermino
        
        #Ingresamos los datos de las fechas en la matriz
        for i in range(PlazoSemestre-1):
            FechaTermino=FechaTermino+md.monthdelta(-6)
            data_precio[PlazoSemestre-i-2][0]=FechaTermino
            data_estimada[PlazoSemestre-i-2][0]=FechaTermino
            
        #Fijamos fecha del flujo anterior para el VP TERA
        FechaFlujoAnterior = FechaTermino+md.monthdelta(-6)

        #Ingresamos los plazos entre fechas a la matriz
        #Fila 2: Ingresamos flujos
        if (codigo[0:3]=="SNT"): #0 cupon
            Cuota=TIR/(1/T[j]) #cuota cupon
            Cuotas[j]=[Cuota+100]
            TIR=((100+Cuota)/100)**(1/T[j])-1
            xlsx[j]=TIR*100
            TT[j]=[T[j]]
            
            sumatoria_precio=(100+Cuota)/((1+TIR)**T[j])
            
        else:
            Cuotas[j]=[0]*PlazoSemestre
            for i in range(PlazoSemestre):
                data_precio[i][1]=Cuota*100
                data_estimada[i][1]=Cuota*100
                Cuotas[j][i]=Cuota*100

            #Ingresamos el nocional
            data_precio[PlazoSemestre-1][1]=data_precio[PlazoSemestre-1][1]+100
            data_estimada[PlazoSemestre-1][1]=data_estimada[PlazoSemestre-1][1]+100
            Cuotas[j][PlazoSemestre-1]+=100
            
            #Ingresamos plazos en años a la matriz estimada
            TT[j]=[0]*PlazoSemestre
            for i in range(PlazoSemestre):
                data_estimada[i][2]=((data_estimada[i][0]-FechaValoracion).days/365)
                TT[j][i]=data_estimada[i][2] 
        
            #Calculamos el valor presente del BCP/BCU
            sumatoria_precio=0;
            for i in range(PlazoSemestre):
                sumatoria_precio=data_precio[i][1]/((1+(TIR/100))**((data_precio[i][0]-FechaValoracion).days/365))+sumatoria_precio
                    
        #Calculo tasa TERA
        TERA=(1+TasaEmision)**(365/360)-1;

        #Calculo Valor Presente TERA
        ValorTERA=100*(1+TERA)**((FechaValoracion-FechaFlujoAnterior).days/365)
        Teras[j]=ValorTERA

        #Calcular valor bonos
        Valorar_Bono_precio=sumatoria_precio/ValorTERA*100
        precios[j]=Valorar_Bono_precio
       

    #OPTIMIZAR PRECIOS
    #crear matriz parametros iniciales
    param=[0]*6 #b0, b1, b2, b3, lambda1, lambda2
    #crear matriz de precios calculados
    P_SV=[0]*len(T);P_SV_aux=[];
    #crear matriz de error^2
    RES2=[0]*len(T);RES2_aux=[0]*len(T)
    #crear matriz de suma de error^2
    sum_RES2=100000;sum_RES2_aux=0
    #crear matriz de tasas
    TASAS=[0]*len(T);
    #Tasa politica monetaria
    TPM=2.5
    
          
    #funcion objetivo    
    def fun(p, T, TT, c, tera, y): #p=[b0,b1,b2,b3,lamb1,lamb2], T=maturity, TT=tiempo cuotas, c=cuotas, tera=teras y y=precios
        e=0
        for i in range(len(T)):
            sumatoria=0;
            t_sv=p[0]+p[1]*((1-exp(-p[4]*T[i]))/(p[4]*T[i]))+p[2]*(((1-exp(-p[4]*T[i]))/(p[4]*T[i]))-exp(-p[4]*T[i]))+p[3]*(((1-exp(-p[5]*T[i]))/(p[5]*T[i]))-exp(-p[5]*T[i]))
            for j in range(len(TT[i])):
                #sumatoria=(c[i][j]*exp(-t_sv*TT[i][j]/100)) + sumatoria; #composicion continua
                sumatoria=c[i][j]/((1+t_sv/100)**TT[i][j]) + sumatoria; #composicion semestral
            precio=sumatoria/tera[i]*100
            e =(precio-y[i])**2 + e
        return e

    #simular N veces
    for n in range(N):
        #parametros iniciales random para cada simulacion
        rr=15.0;
        param_aux=[rd.uniform(0,rr),rd.uniform(-rr,rr),rd.uniform(-rr,rr),rd.uniform(-rr,rr),rd.uniform(-rr/5,rr/5),rd.uniform(-rr/10,rr/10)]#b0, b1, b2, b3, lambda1, lambda2

        x = array(T)  #madurez
        xx = array(TT) #tiempos de evaluacion
        c = array(Cuotas) #cuotas
        tera= array(Teras) #valores tera
        y = array(precios)  #solo precios disponibles
        
        p0 = array(param_aux)  # parametros iniciales
        cons = ({'type': 'eq','fun': lambda params : params[0]+params[1] - TPM},{'type': 'ineq','fun': lambda params : params[4]},{'type': 'ineq','fun': lambda params : params[5]}) #restricciones de b0+b1=tpm y lambda1,lambda2>=0
        p = minimize(fun, p0, args=(x,xx,c,tera,y),method='SLSQP',constraints=cons)#options={'maxiter': 10000}

        c=[0]*6
        for i in range(len(p.x)):
            c[i]=p.x[i]

        NN=[] #lista para guardar valores del precio SV
        for i in range(len(T)):  #calcular los precios con SV
            SV=c[0]+c[1]*((1-exp(-c[4]*T[i]))/(c[4]*T[i]))+c[2]*(((1-exp(-c[4]*T[i]))/(c[4]*T[i]))-exp(-c[4]*T[i]))+c[3]*(((1-exp(-c[5]*T[i]))/(c[5]*T[i]))-exp(-c[5]*T[i]))
            sumatoria=0
            for j in range(len(TT[i])):
                #sumatoria=Cuotas[i][j]*exp(-SV*TT[i][j]/100)+sumatoria #continua
                sumatoria=Cuotas[i][j]/((1+SV/100)**TT[i][j])+sumatoria #semestral
            NN.append(sumatoria/Teras[i]*100)
        P_SV_aux=NN

        #calcular residuos^2
        for i in range(len(T)):
            RES2_aux[i]=(P_SV_aux[i]-precios[i])**2
        sum_RES2_aux=sum(RES2_aux)

        #si tiene mejores resultados, guardar valores
        if (sum_RES2_aux<sum_RES2):
            print("Cambiaron parametros en iteracion ",n+1)
            param=c
            P_SV=P_SV_aux
            precios_SV=P_SV
            RES2=RES2_aux
            sum_RES2=sum_RES2_aux
            for i in range(len(T)):
                SV=c[0]+c[1]*((1-exp(-c[4]*T[i]))/(c[4]*T[i]))+c[2]*(((1-exp(-c[4]*T[i]))/(c[4]*T[i]))-exp(-c[4]*T[i]))+c[3]*(((1-exp(-c[5]*T[i]))/(c[5]*T[i]))-exp(-c[5]*T[i]))
                TASAS[i]=SV
        #print("Iteracion ",n+1," completada.")

    

    #COMPARACION RISK AMERICA
    xlsx_df = pd.read_excel("CurvasGobCeroPesos.xlsx") #dataframe
    T_RA=list(xlsx_df["Plazo"]) #Plazos definidos por risk america
    RA=list(xlsx_df["RiskAm"]) #Tasas estimadas por risk america
    TE=[0]*len(T_RA) #Tasas estimadas por modelo NS

    c=param
    #Tasas estimadas en plazos RA
    for i in range(len(T_RA)):
        TE[i]=c[0]+c[1]*((1-exp(-c[4]*T_RA[i]))/(c[4]*T_RA[i]))+c[2]*(((1-exp(-c[4]*T_RA[i]))/(c[4]*T_RA[i]))-exp(-c[4]*T_RA[i]))+c[3]*(((1-exp(-c[5]*T_RA[i]))/(c[5]*T_RA[i]))-exp(-c[5]*T_RA[i]))


    print("")

    for i in range(len(T)):
        print("T: ", T[i])
        print("Precio observado: ", round(precios[i],2))
        print("Precio predicho: ", round(precios_SV[i],2))
        print("Diferencia de precios: ", round(precios[i]-precios_SV[i],2))
        print("Tasa observada: ", round(xlsx[i],2))
        print("Tasa predicha: ", round(TASAS[i],2))
        m=bisect(T_RA,T[i])
        print("Tasa risk america: ", round(RA[m],2))
        print("")

    print("Parametros: ", param[0], param[1], param[2], param[3], param[4], param[5])




    #GUARDAR RESULTADOS EN EXCEL
    workbook = xlwt.Workbook('data_results_precios.xlsx') #donde se guardan los resultados
    data_worksheet = workbook.add_worksheet("data results") 

    #formatos celdas
    bold=workbook.add_format({'bold': True}) #titulos negrita
    decimales = workbook.add_format()
    decimales.set_num_format('0.0000') #solo mostrar 2 decimales

    #Titulos
    data_worksheet.write(0,0,"Plazo",bold)
    data_worksheet.write(0,1,"Tasa RA",bold)
    data_worksheet.write(0,2,"Tasa SV",bold)
    data_worksheet.write(0,3,"Tasa Mercado",bold)
    data_worksheet.write(0,4,"Precio SV0",bold)
    data_worksheet.write(0,5,"Precio Mercado",bold)
    data_worksheet.write(0,6,"Precio SV",bold)
    

    for i in range(len(T_RA)):
        data_worksheet.write(i+1,0,T_RA[i],decimales)
        data_worksheet.write(i+1,1,RA[i],decimales)
        data_worksheet.write(i+1,2,TE[i],decimales)
        psv=100*exp(-T_RA[i]*TE[i]/100)
        data_worksheet.write(i+1,4,psv,decimales)

    for i in range(len(T)):
        data_worksheet.write(bisect(T_RA,T[i]),3,xlsx[i],decimales)
        data_worksheet.write(bisect(T_RA,T[i]),5,precios[i],decimales)
        data_worksheet.write(bisect(T_RA,T[i]),6,precios_SV[i],decimales)
            
    data_worksheet.write(0,8,"b0",bold)
    data_worksheet.write(1,8,"b1",bold)
    data_worksheet.write(2,8,"b2",bold)
    data_worksheet.write(3,8,"b3",bold)
    data_worksheet.write(4,8,"lambda 1",bold)
    data_worksheet.write(5,8,"lambda 2",bold)
    data_worksheet.write(0,9,param[0],decimales)
    data_worksheet.write(1,9,param[1],decimales)
    data_worksheet.write(2,9,param[2],decimales)
    data_worksheet.write(3,9,param[3],decimales)
    data_worksheet.write(4,9,param[4],decimales)
    data_worksheet.write(5,9,param[5],decimales)
    
    data_worksheet.write(0,11,"Fecha valoracion",bold)
    data_worksheet.write(0,12,FechaValoracion.strftime("%d-%m-%Y"))
    
    workbook.close()




    #GRAFICAR
    x1=T_RA 
    xx1=T
    y11=RA
    y12=TE
    y13=xlsx
    
    x2=T
    y21=precios_SV
    y22=precios

    plt.subplot(2, 1, 1)
    plt.plot(x1, y11, '-r', label='Risk America')
    plt.plot(x1, y12, '-b', label='Estimado')
    plt.plot(xx1, y13, '-g', label='Mercado')
    plt.title('Estimaciones de tasas optimizando diferencia de precios')
    plt.legend(loc='lower right')
    plt.ylabel('Tasas')

    plt.subplot(2, 1, 2)
    plt.plot(x2, y21, '-b', label='Estimado')
    plt.plot(x2, y22, '-g', label='Mercado')
    plt.legend(loc='lower right')
    plt.xlabel('Madurez (años)')
    plt.ylabel('Precios')

    plt.show()


excel_data = 'data_test.xlsm'
N=2500#int(input("Numero de simulaciones? ")); #numero de simulaciones

print("Hora inicio: ",time.strftime("%H:%M:%S"))
Optimizar_precio(excel_data,N)
print("Hora termino: ",time.strftime("%H:%M:%S"))
