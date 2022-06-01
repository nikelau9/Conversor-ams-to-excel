# -*- coding: utf-8 -*-
"""
Created on Wed Mar  9 14:12:38 2022

@author: nikel
"""
from tkinter import Tk, StringVar,Y, W, X, PhotoImage, Toplevel, YES, END, Label,ANCHOR, ACTIVE, RIGHT,LEFT, Button, BOTH, TOP, BOTTOM, Frame,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL, Checkbutton, Listbox, MULTIPLE
import xlsxwriter 
import pandas as pd
from pandas import DataFrame
frame1 = Tk()
frame1.config(bg='white')
frame1.geometry('600x100')
frame1.minsize(width=400, height=50)
frame1.title('AMS CNA')
frame1.columnconfigure(0, weight = 1)
frame1.rowconfigure(1, weight= 1)
frame1.columnconfigure(0, weight = 1)
frame1.rowconfigure(2, weight= 1)
frame1.columnconfigure(0, weight = 1)
frame1.rowconfigure(0, weight= 1)


def selecciondemedida():
    dataws.deiconify()
def cerrardataws():
    dataws.withdraw()
    
def indicaI():
    indica3['text']='129I'
def indicaU():
    indica3['text']='236U'
def indicaCa():
    indica3['text']='41Ca'
def indicaCl():
    indica3['text']='37Cl'
    
def abrir_archivo():
    archivo = filedialog.askopenfilename(initialdir ='/', 
                                            title='Selecione archivo', 
                                            filetype=(('ams files', '*.ams*'),('All files', '*.*')))
    indica['text'] = archivo  

def sacardato(lerroa,limiteval):       
    indexagarri=[]
    l=0
    for hizki in lerroa:
        if hizki!='NaN' and l>=limiteval:
            indexagarri.append(l)
        l=int(l+1)
    return indexagarri

def Guardarexcel():
    indica2['text']=filedialog.asksaveasfilename(title='Selecione archivo con los números del carrusel', filetype=(('xlsx files', '*.xlsx*'),('All files', '*.*')))
    indica2['text']=indica2['text']+'.xlsx'
    workbook = xlsxwriter.Workbook(indica2['text']) #LocalizaciÃ³n y nombre del Archivo que se quiere obtener xlsx
    worksheet = workbook.add_worksheet() 
    datos_obtenidos = indica['text']
    fichero_input= open(datos_obtenidos,"r") #LocalizaciÃ³n y nombre del Archivo que quiere leer
    fa=fichero_input.read()
    p1=fa.find('Blocks')
    p2=fa.find('# Analyses')
    documento=fa[p1+7:p2-1]
    documento=documento.split("\n")
    k=1
    j=0
    limiteval=7
    if indica3['text']=='236U':
        sonnumeros=[0 ,2, 3,7, 8, 9, 10, 11, 12, 13,14, 15, 16, 17,18,19, 20, 21,22,23,24,25,28,29] 
        limite=8
    elif indica3['text']=='37Cl':
        sonnumeros=[] 
        limite=4
    elif indica3['text']=='41Ca':
        sonnumeros=[0 ,2, 3,7, 8, 9, 10, 11, 12, 13,14, 15, 16, 17, 20, 21] 
        limite=3
        
    #Seleccionar dependiendo de medida
    variables=[]
    recopilacion={}
    for lerro in documento:
        if k==1:
            cabecero=lerro.split("\t")
            for i in range(len(cabecero)):
                variables.append(cabecero[i])
            k=k+1      
        else:
            j=j+1
            lerro=lerro.split("\t")
            nombre=lerro[1]
            if j==1:
                archivodatosmin=[None]*len(variables)
                for x in range(limiteval):
                    archivodatosmin[x]=lerro[x]
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
            
            elif j==limite:
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
                if nombre in recopilacion:
                    tt=recopilacion[nombre]
                    tt.append(archivodatosmin)
                    recopilacion[nombre]=tt
                else:
                    tt=[archivodatosmin]
                    recopilacion[nombre]=tt
                j=0
            else:
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
    row=0
    column=0        
    for item in variables :   
        worksheet.write(row, column, item)    
        column += 1

    for laginizena in recopilacion:
        archivodedatos=recopilacion[laginizena]
        cnt=1
        for lisra in archivodedatos:
            row+=1
            column=0
            for item in lisra:
                if column in sonnumeros:
                    worksheet.write(row, column, float(item))
                elif column==5:
                    worksheet.write(row, column, int(cnt))
                    cnt+=1
                else:
                    worksheet.write(row, column, item)
                column += 1
    workbook.close() 
    fichero_input.close()

def crearlista():
    archivo=archivo = filedialog.askopenfilename(initialdir ='/', 
                                            title='Selecione archivo', 
                                            filetype=(('Excel files', '*.xlsm*'),('Excel files', '*.xlsx*'),('All files', '*.*')))
    df = pd.read_excel(archivo)
    ll=df.values.tolist()
    lekua=0
    izena=0
    brugakoa=0
    antolaketa={}
    ro=0
    co=0
    for zerrenda in ll:
        co=0
        for elem in zerrenda:
            if elem=='Position':
                lekua=co
                brugakoa=ro
            if elem=='Sample ID':
                izena=co
            co+=1
        ro+=1

    ro=0
    co=0
    for zerrenda in ll:
        if ro>=brugakoa:
            aa='" ' + str(zerrenda[izena]) + ' "'
            antolaketa[aa]=zerrenda[lekua]
        ro+=1
    return antolaketa
def anadircabecero():
    indica2['text']=filedialog.asksaveasfilename(filetype=(('xlsx files', '*.xlsx*'),('All files', '*.*')))
    indica2['text']=indica2['text']+'.xlsx'
    
    workbook = xlsxwriter.Workbook(indica2['text']) #LocalizaciÃ³n y nombre del Archivo que se quiere obtener xlsx
    worksheet = workbook.add_worksheet() 
    datos_obtenidos = indica['text']
    fichero_input= open(datos_obtenidos,"r") #LocalizaciÃ³n y nombre del Archivo que quiere leer
    fa=fichero_input.read()
    p1=fa.find('Blocks')
    p2=fa.find('# Analyses')
    documento=fa[p1+7:p2-1]
    documento=documento.split("\n")
    k=1
    j=0
    limiteval=7
    if indica3['text']=='236U':
        sonnumeros=[0 ,2, 3,7, 8, 9, 10, 11, 12, 13,14, 15, 16, 17,18,19, 20, 21,22,23,24,25,28,29] 
        limite=8
    elif indica3['text']=='37Cl':
        sonnumeros=[] 
        limite=4
    elif indica3['text']=='41Ca':
        sonnumeros=[0 ,2, 3,7, 8, 9, 10, 11, 12, 13,14, 15, 16, 17, 20, 21] 
        limite=3
        
    #Seleccionar dependiendo de medida
    variables=[]
    recopilacion={}
    for lerro in documento:
        if k==1:
            cabecero=lerro.split("\t")
            for i in range(len(cabecero)):
                variables.append(cabecero[i])
            k=k+1      
        else:
            j=j+1
            lerro=lerro.split("\t")
            nombre=lerro[1]
            if j==1:
                archivodatosmin=[None]*len(variables)
                for x in range(limiteval):
                    archivodatosmin[x]=lerro[x]
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
            
            elif j==limite:
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
                if nombre in recopilacion:
                    tt=recopilacion[nombre]
                    tt.append(archivodatosmin)
                    recopilacion[nombre]=tt
                else:
                    tt=[archivodatosmin]
                    recopilacion[nombre]=tt
                j=0
            else:
                index=sacardato(lerro,limiteval)
                for valores in index:
                    archivodatosmin[valores]=lerro[valores]
    row=0
    column=0   
    adhesion=crearlista()
    for elem in adhesion:
        if elem=='" Sample ID "':
            variables.append(adhesion[elem])
        else:
            if elem in recopilacion:
                gehitua=[]
                listareco=recopilacion[elem]
                for posizioagehitu in listareco:
                    posizioagehitua=posizioagehitu
                    posizioagehitua.append(adhesion[elem])
                    gehitua.append(posizioagehitua)
                recopilacion[elem]=gehitua
            
    
    for item in variables :   
        worksheet.write(row, column, item)    
        column += 1

    for laginizena in recopilacion:
        archivodedatos=recopilacion[laginizena]
        cnt=1
        for lisra in archivodedatos:
            row+=1
            column=0
            for item in lisra:
                if column in sonnumeros:
                    worksheet.write(row, column, float(item))
                elif column==5:
                    worksheet.write(row, column, int(cnt))
                    cnt+=1
                else:
                    worksheet.write(row, column, item)
                column += 1
    workbook.close() 
    fichero_input.close()



boton1 = Button(frame1, text= 'Abrir', bg='blue', command= abrir_archivo)
boton1.grid(column = 0, row = 0, sticky='nsew', padx=10, pady=10)

boton2 = Button(frame1, text= 'Seleccionar elemento \n de la medida ', fg='white', bg='black', command= selecciondemedida)
boton2.grid(column = 1, row = 0, sticky='nsew', padx=10, pady=10)
boton4 = Button(frame1, text= 'Guardar en excel \n con la posición\n del carrusel ', fg='white', bg='green2', command= anadircabecero)
boton4.grid(column = 3, row = 0, sticky='nsew', padx=10, pady=10)
boton3 = Button(frame1, text= 'Guardar en Excel', bg='green', command= Guardarexcel)
boton3.grid(column = 2, row = 0, sticky='nsew', padx=10, pady=10)
indica = Label(frame1, fg= 'white', bg='gray26', text= '' , font= ('Arial',10,'bold') )  
indica2 = Label(frame1, fg= 'white', bg='gray26', text= '' , font= ('Arial',10,'bold') )  
dataws = Toplevel() 
dataws.title('Selección de la medida') 
dataws.geometry('50x150')
dataws.withdraw()
indica3 = Label(frame1, fg= 'white', bg='gray26', text= '' , font= ('Arial',10,'bold') )
indica3.grid(column=2, row = 1)
Button(dataws, text='\u00B9'+ '\u00B2'+ '\u2079' + 'I',bg='yellow', command=indicaI).pack(side = TOP,anchor=W, fill=X)
Button(dataws, text="\u00B2\u00B3\u2076U",bg='blue', command=indicaU).pack(side = TOP,anchor=W, fill=X)
Button(dataws, text="\u2074\u00B9Ca",bg='DarkOrchid3', command=indicaCa).pack(side = TOP,anchor=W, fill=X)
Button(dataws, text="\u00B3\u2077Cl",bg='orange', command=indicaCl).pack(side = TOP,anchor=W, fill=X)
Button(dataws, text="Cerrar", bg='red',command=cerrardataws).pack(side = BOTTOM,anchor=W, fill=X)
frame1.mainloop()