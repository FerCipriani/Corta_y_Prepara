from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import *
import pathlib
import os, re
import shutil
from zipfile import ZipFile, ZIP_DEFLATED
from PyPDF2 import PdfFileWriter, PdfFileReader
# auto-py-to-exe desde la consola

def Renombra(PathSel, files):   
#Arranca en el dir, se fija por string a que familia pertenecen,     los renombra segun la matriz de Nombres, Si hay mas de alguna familia los numera. 
       
    Stri= ['desp',  'fac', 'con', 'ori', 'otr']
    Nombr = ['1-OM', '2-FACTURA_COMERCIAL', '3-CONOCIMIENTO_DE_EMBARQUE', '4-CERTIFICADOS_DE_ORIGEN', '5-3ros_ORGANISMOS']
    #Nombr =   ['Despacho', 'Factura', 'Conocimiento', 'CertOrigen', 'Otros']

    files = os.listdir(PathSel) 
    files = [f for f in files if os.path.isfile(f)and f.endswith('.pdf')]
    
    Conta = 0
    #vamos a buscar y renombramos:
    for i in range(len(files)):
        for j in range(len(Stri)):
            Conta = 0
            if files[i].lower().find(Stri[j])!= -1:
                while os.path.exists(Nombr[j] + "-" + str(Conta) + '.pdf'):
                    Conta = Conta + 1
                os.rename(files[i], Nombr[j] + "-" + str(Conta) + '.pdf')
                                    
    return

def Zipea(PathSel):
# Toma como dato el path seleccionado. Hace una lista con los PDFs, toma el nombre del directorio superior para usarlo de nombre en el zip. comprime todo
    
    files = os.listdir(PathSel)
    files = [f for f in files if os.path.isfile(f)and f.endswith('.pdf')] #Filtro solo archivos y que sean .pdf
    
    dirs = PathSel.split(os.sep)

    archivo = dirs[-1] + ' - Legajo.zip'
    Destino = "Q:/Temporales/" + archivo
    
    with ZipFile(archivo, 'w', ZIP_DEFLATED, compresslevel=9) as zipObj:
        for i in files:
            zipObj.write(i)
    zipObj.close()

    shutil.copy(archivo, Destino)
    return    

def PartePDF(PathSel):

    ABorrar = os.listdir(PathSel)
    ABorrar = [f for f in ABorrar if os.path.isfile(f)and f.endswith('.pdf')] #Filtro solo archivos y que sean .pdf

    for Pds in ABorrar:
        pdf_file = open(Pds,'rb')
        pdf_reader = PdfFileReader(pdf_file)
        pageNumbers = pdf_reader.getNumPages()
        #if pageNumbers > 1:
        for i in range (pageNumbers):
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf_reader.getPage(i))
            split_motive = open(Pds[0:-4] + str(i+1) + '.pdf','wb')
            pdf_writer.write(split_motive)
            split_motive.close()
        pdf_file.close()
        os.remove(Pds)


    return

def PesaPDF(PathSel): 

    Peso = os.listdir(PathSel)
    Peso = [f for f in Peso if os.path.isfile(f)and f.endswith('.pdf')] 
    
    xlimite = 307200/1024

    for it in Peso:
        file_size = int(os.path.getsize(it)/1024)
        if file_size >= 300:
            mb.showinfo( message="El archivo: " + it + " Pesa más de 300 KB!")
               
    return
        

'''
        Cuerpo Principal
'''

ventana = Tk()
ventana.title("Dodero & Cia")

color = 'gray89'
ventana.configure(background= color)

Stri= ['desp',  'fac', 'con', 'ori', 'otr']

ancho_ventana = 250
alto_ventana = 80

x_ventana = ventana.winfo_screenwidth() // 2 - ancho_ventana // 2
y_ventana = ventana.winfo_screenheight() // 2 - alto_ventana // 2
 
posicion = str(ancho_ventana) + "x" + str(alto_ventana) + "+" + str(x_ventana) + "+" + str(y_ventana)
ventana.geometry(posicion)

Label(ventana, text="Seleccione el directorio con los PDFs a partir", bg= color,).pack()

path = "Q:/Logistica/Digitalizacion General/"
#Cambiar a path de laburo en "C:/Users/Fer.Cipriani/Documents/Git/"

def carpeta():

    directorio=filedialog.askdirectory(initialdir=path)
    if directorio!="":
        os.chdir(directorio)

        PathSel = os.getcwd() # este es el directorio seleccionado
        #mb.showinfo('Directorio seleccionado:', PathSel)


        dirs = PathSel.split(os.sep)

        # Revisamos la estructura del directorio que sea correcta
        if re.match(r'([C]\s[0-9]{5}\s[-]\s[A-Za-z\s]{3,35})', dirs[-1])== None:
            mb.showinfo( message="El directorio: " + dirs[-1] + '\n' + "No corresponde a un directorio de carpeta!")
            return


        files = os.listdir(PathSel) 
        files = [f for f in files if os.path.isfile(f)and f.endswith('.pdf')] #es un array con los nombres de archivos.
        
        # Verificamos que tengamos PDFs
        if len(files) < 1:
            mb.showerror(message='No se encontraron archivos PDF', title="Error")
            return

        # Revisamos que pertenezcan a las familias
        for i in range(len(files)):
            s=0
            for j in range(len(Stri)):
                if files[i].lower().find(Stri[j])== -1: # Si encuentra
                    s=s+1
                    if s == len(Stri):
                        mb.showinfo( message="El archivo: " + files[i] +'\n'+ "No puede ser asignado a ninguna familia!")
                        return
                           
        msg = '\n'.join(files) # Muestro los archivos en una ventana
        mb.showinfo(message='Los archivos a partir son: '+ '\n' + '\n' + msg, title="Archivos a cortar")

        Renombra(PathSel, files)

        Zipea(PathSel)

        PartePDF(PathSel)

        PesaPDF(PathSel)
        
        if os.path.exists("Q:/Logistica/Digitalizacion General/1-OM-Tapa.pdf")==True:
            Origen = "Q:/Logistica/Digitalizacion General/1-OM-Tapa.pdf"
            Destino = PathSel + os.sep + "1-OM-Tapa.pdf"
            shutil.copy(Origen, Destino)
        else:
            mb.showinfo(message='No encuentro el archivo 1-OM-Tapa.pdf' + '\n' + 'Hay que copiarlo a mano!', title="Falta Tapa!")        
        
          

Button(text="Ok",bg="gray",command=carpeta, height = 1, width = 8).place(x=50,y=30)

def salir():
    ventana.destroy()

Button(text="Salir",bg="gray",command=salir, height = 1, width = 8).place(x=120,y=30)

ventana.mainloop()