import pandas as pd #para excel y dataframe, manejo y análisis de estructuras de datos
from datetime import datetime #permite manipular fechas y hora
from docxtpl import DocxTemplate #genera documentos a través de plantillas .docx

#doc=DocxTemplate("plantilla_informe.docx") #le pasamos el archivo de Word si está en la misma carpeta
doc=DocxTemplate("C:/Users/user/OneDrive/Documents/plantilla_informe.docx") 
#le pasamos el archivo de Word si tenemos una ruta de la carpeta fuente

'''
En este caso estamos dando una ruta donde va a cargar plantilla_informe.docx y Alumnos.xlsx que tienen la 
información base, el informe generado para cada alumno también será almacenado en esa misma ruta. Como mejora
se podría colocar en el mismo Excel los datos constantes del maestro o leerlos de otro archivo.
'''
# Estos son los datos fijos del maestro,------- 
test="testes"
nombre="Erick Mena"
telefono="(123) 456-7890"
correo="erick@gmail.com"
fecha=datetime.today().strftime("%d/%m/%Y") # obtiene la fecha de hoy con formato d/m/aaaa

constantes={'nombre':nombre,'telefono':telefono,'correo':correo,'fecha':fecha}
# la clave debe coincidir con el nombre usado en el archivo de Word

doc.render(constantes) #reemplaza las variables en el documento

# ---------- USADO SOLO PARA VALIDACION  ------------------------
#doc.save(f"prueba.docx") #se guardan las constantes en el documento y genera el Word 

#df=pd.read_excel('Alumnos.xlsx') # se lee el archivo de Excel y se guarda
df=pd.read_excel('C:/Users/user/OneDrive/Documents/Alumnos.xlsx') # se lee el archivo de Excel y se guarda

for indice, fila in df.iterrows(): # va iterando por cada fila
    contenido={
        'nombre_alumno':fila["Nombre del Alumno"],
        'nota_mat':fila["Mat"],
        'nota_fis':fila["Fis"],
        'nota_qui':fila["Qui"],
    }
    contenido.update(constantes) # se agregan las constantes al diccionario
    doc.render(contenido)
    
    #aqui guarda y genera el archivo de Word en la ruta que se le indique
    doc.save('C:/Users/user/OneDrive/Documents/' + f"Notas_de_{fila['Nombre del Alumno']}.docx") 
    
    #aqui guarda y genera el archivo de Word
    #doc.save(f"Notas_de_{fila['Nombre del Alumno']}.docx") 