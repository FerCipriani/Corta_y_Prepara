#Nueva version 07Ago24
#pip install pyinstalle
#python -m PyInstaller --onefile --windowed CortayPrepara20.py

from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import *
import pathlib
import os, re
import shutil
from zipfile import ZipFile, ZIP_DEFLATED
from PyPDF2 import PdfWriter, PdfReader
import MySQLdb

db=MySQLdb.connect(host="192.168.123.76",user="reportes",passwd="Dodero2022",db="taidatos")
c = db.cursor(MySQLdb.cursors.DictCursor)

# auto-py-to-exe desde la consola

def ConsultaEmpresa(Car):
        
    DFEmp = [] 
    c.execute("SELECT `Carpeta`,`Cliente`FROM `taidatos`.`dir_carpetas` WHERE `Carpeta` = " + Car)
    result_set = c.fetchall()

    for row in result_set:
        DFEmp = row["Cliente"]
    return(DFEmp)

def Renombra(PathSel, files):   
#Arranca en el dir, se fija por string a que familia pertenecen, los renombra segun la matriz de Nombres, Si hay mas de alguna familia los numera. 
  
    Stri = ['desp', 'fac', 'hoja', 'con', 'ori', 'otr']
    Nombr = ['1-OM', '2-FACTURA_COMERCIAL', '2-HOJA_de_VALOR', '3-CONOCIMIENTO_DE_EMBARQUE', '4-CERTIFICADOS_DE_ORIGEN', '5-3ros_ORGANISMOS']
    
    # Filtra los archivos .pdf en el directorio especificado
    files = [f for f in os.listdir(PathSel) if os.path.isfile(os.path.join(PathSel, f)) and f.endswith('.pdf')]
    
    # Diccionario para llevar el conteo de archivos renombrados para cada nuevo nombre
    file_counts = {name: 0 for name in Nombr}
    
    for file in files:
        for j in range(len(Stri)):
            if Stri[j] in file.lower():
                file_counts[Nombr[j]] += 1
                new_name = f"{Nombr[j]}_{file_counts[Nombr[j]]:01d}.pdf"
                os.rename(os.path.join(PathSel, file), os.path.join(PathSel, new_name))
                #print(f"Renamed '{file}' to '{new_name}'")
                break  # Salir del bucle interior si se ha encontrado una coincidencia y renombrado
    return

def Zipea(PathSel):
# Toma como dato el path seleccionado. Hace una lista con los PDFs, toma el nombre del directorio superior para usarlo de nombre en el zip. comprime todo
    
    files = os.listdir(PathSel)
    files = [f for f in files if os.path.isfile(f)and f.endswith('.pdf')] #Filtro solo archivos y que sean .pdf
    
    dirs = PathSel.split(os.sep)

    archivo = dirs[-1] + ' - Legajo.zip'
    Destino = "Q:/Temporales/" + archivo
    #Destino = path + archivo
    
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
        pdf_reader = PdfReader(pdf_file)
        pageNumbers = len(pdf_reader.pages)
        
        for i in range (pageNumbers):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[i])
            split_motive = open(Pds[0:-4] + '_' + str(i+1) + '.pdf','wb')
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

def mostrar_ayuda():
    mensaje = (
        "Con ABRIR se selecciona el directorio donde están los archivos a procesar.\n"
        "\n"
        "El directorio seleccionado tiene que estar formado de la siguiente manera:\n"
        "\n"
        "C - 99999 - Nombre del cliente\n"
        "\n"
        "Se van a procesar los archivos .pdf del directorio que contengan en su nombre 'desp',  'fac', 'hoja', 'con', 'ori' y/ó 'otr'.\n"
        "\n"
        "Si encuentra algún archivo que no tenga eso en el nombre, lo muestra y detiene el proceso.\n"
        "\n"
        "Tambien revisa si hay algún archovo de más de 300 KB que no pueda ser tomado por el F3101.\n"
        "\n"
        "Una vez procesados se van a generar:\n"
        "\n"
        "* Los archivos necesarios para generar el Formulario 3101 con el aplicativo de Alpha2000.\n" 
        "\n"
        "* Los archivos para enviar el legajo al cliente de forma automática.\n" 
        "\n"
        "Con SALIR terminamos el programa."
    )
    mb.showinfo("Ayuda", mensaje)

def salir():
    ventana.destroy() 

def DisLiv(str1, str2):
  str1=str1.lower().replace('.', '').replace(' ', '')
  str2=str2.lower().replace('.', '').replace(' ', '')
  d=dict()
  for i in range(len(str1)+1):
     d[i]=dict()
     d[i][0]=i
  for i in range(len(str2)+1):
     d[0][i] = i
  for i in range(1, len(str1)+1):
     for j in range(1, len(str2)+1):
        d[i][j] = min(d[i][j-1]+1, d[i-1][j]+1, d[i-1][j-1]+(not str1[i-1] == str2[j-1]))
  return d[len(str1)][len(str2)]

'''
Cuerpo Principal

'''

ventana = Tk()
ventana.title("Dodero & Cia")

color = 'gray89'
ventana.configure(background= color)

Stri= ['desp',  'fac', 'hoja', 'con', 'ori', 'otr']

ancho_ventana = 300
alto_ventana = 80

x_ventana = ventana.winfo_screenwidth() // 2 - ancho_ventana // 2
y_ventana = ventana.winfo_screenheight() // 2 - alto_ventana // 2
 
posicion = str(ancho_ventana) + "x" + str(alto_ventana) + "+" + str(x_ventana) + "+" + str(y_ventana)
ventana.geometry(posicion)

Label(ventana, text="Seleccione el directorio con los PDFs a partir", bg= color,).pack()

path = "Q:/Logistica/Digitalizacion General/"
#path = "C:/Users/Fer.Cipriani/Documents/Cortayprepara/"

def carpeta():

    directorio=filedialog.askdirectory(initialdir=path)
    if directorio!="":
        os.chdir(directorio)

        PathSel = os.getcwd() # este es el directorio seleccionado
        dirs = PathSel.split(os.sep)

        # Revisamos la estructura del directorio que sea correcta (ahora dejamos ingresar numeros en el cliente)
        if re.match(r'([C]\s[0-9]{5}\s[-]\s[A-Za-z0-9\s]{3,35})', dirs[-1])== None:
            mb.showinfo( message="El directorio: " + dirs[-1] + '\n' + "No corresponde a un directorio de carpeta!")
            return

        files = os.listdir(PathSel) 
        files = [f for f in files if os.path.isfile(f)and f.endswith('.pdf')] #es un array con los nombres de archivos.
        print(files)

        # Verificamos que tengamos PDFs
        if len(files) < 1:
            mb.showerror(message='No se encontraron archivos PDF', title="Error")
            return

    
        Carpeta = re.search(r'C (\d{5}) -', dirs[-1]).group(1)
        Empre = dirs[-1].split('-')[-1].strip()


        if DisLiv(ConsultaEmpresa(Carpeta),Empre)>3:
            respuesta = mb.askyesno(message = f"Tenemos diferencias:\n\nEsa carpeta corresponde en el TAI a la empresa {ConsultaEmpresa(Carpeta)}\n\nY en el directorio figura: {Empre}\n\nEl error esta en el número de carpeta {Carpeta} o cómo se escribió la empresa en el directorio\n\nContinuar con el error (SI) o cancelar (NO)?", title="Incongruencia!")
            # Comprueba la respuesta
            if respuesta==False:
                return #ventana.destroy() #salir #print("El usuario seleccionó Sí")
        
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


boton_ancho = 7 * 10  # Aproximadamente 10 píxeles por carácter en el botón
boton_alto = 20  # Altura del botón en píxeles
espaciado = 3  # Espaciado entre los botones

# Calcular posiciones de los botones
total_ancho_botones = 3 * boton_ancho + 2 * espaciado
x_inicio = (ancho_ventana - total_ancho_botones) // 2
y_pos = (alto_ventana - boton_alto) // 2

x1 = x_inicio
x2 = x1 + boton_ancho + espaciado
x3 = x2 + boton_ancho + espaciado

Button(text="Abrir",bg="gray",command=carpeta, height = 1, width = 7).place(x=x1,y=y_pos)
Button(text="Salir",bg="gray",command=salir, height = 1, width = 7).place(x=x2,y=y_pos)
Button(text="Ayuda",bg="gray",command=mostrar_ayuda, height = 1, width = 7).place(x=x3,y=y_pos)

ventana.mainloop()