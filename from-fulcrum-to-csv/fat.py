__author__ = "Juan Molera Pascual"
__copyright__ = "Alauda Ingeniería SA"
__credits__ = ["Juan Molera Pascual"]
__license__ = "2022"
__version__ = "1.0.10"
__maintainer__ = "Juan Molera Pascual"
__email__ = "jmolera@alaudaingenieria.es"
__status__ = "wip"

'''
Nueva posición para le bucle de fotos
'''

#Librerías
import pandas as pd
import xlwings as xw
import os
import math
import time

#Limpia cmd
os.system("clear")

#Path
cwd=os.getcwd()
print(cwd)

#Método para checkear si falta un dato
def isNaN(string):
    return string != string

#Configuración de Pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

#Lee el csv con los datos de Fulcrum
fulcrum_csv_data=pd.read_csv("prueba_fat_hw_decodificadores.csv")

#Lee el csv que contiene los campos comúnes a todas las pruebas
#Está hardcodeado a mano y sin tildes/letras
fijos_csv_data=pd.read_csv("fijos_fat.csv")

#Lee el csv que contiene los campos variables de las pruebas
#Está hardcodeado a mano y sin tildes/letras
variables_csv_data=pd.read_csv("variables_fat.csv")

#Pega las listas
frames=[fijos_csv_data,variables_csv_data]
result = pd.concat(frames)

#Hay 74 campos en Fulcrum
#Hay 30 campos fijos
#Hay 12 campos variables

#Elimina los campos que genera Fulcrum y no interesan
#Contenidos en la lista de drops
drop_list=list()

k=0
for k in fulcrum_csv_data.columns:
	if k not in result.columns:
		drop_list.append(k)

#Se guardan los datos de csv actualizados
#42 campos
fulcrum_csv_data.drop(columns=drop_list, inplace=True)
#print(len(fulcrum_csv_data.columns))
#print(fulcrum_csv_data.columns)
#print(type(fulcrum_csv_data[fulcrum_csv_data.columns[27]].iloc[0]))
#print(fulcrum_csv_data)

#Se muestran las columnas restantes y con las que trabajamos a partir de aquí
#print(fulcrum_csv_data)

#Se han mantenido solo los datos comúnes del csv
#Esto evidentemente no funciona porque falta los datos específicos de cada prueba
#Lo que se corresponde con el apartado de "Pruebas funcionales"
#Lo ideal sería leer todos los parámetros que se necesitan de la plantilla
#Pero existe el problema de las tildes de Excel vs Fulcrum
#Creo que lo mejor es harcodear un poco esta parte e indicar de donde a donde lee el csv plantilla
#Con dejar las plantillas fijas no debería de ser un problema
#ESTA PARTE NO FUNCIONA HASTA ELIMINAR TILDES DE FULCRUM
#formato_csv_data=pd.read_csv("formato_csv.csv")
#print(formato_csv_data)

for index, row in fulcrum_csv_data.iterrows():
	print(index)
	fulcrum_csv_data[fulcrum_csv_data.columns[35]].iloc[20]
	#Excel
	#Selección de plantilla
	wb = xw.Book("formato_fat.xlsx")
	sheet1 = wb.sheets['Hoja1']
	sheet2 = wb.sheets['Hoja2']
	sheet3 = wb.sheets['Hoja3']

	#Campos de la plantilla
	#Registro de pruebas
	#Número de página:
	sheet1.range('A4').value = 'Página: 1/3'
	sheet2.range('A4').value = 'Página: 2/3'
	sheet3.range('A4').value = 'Página: 3/3'
	#Fecha:
	sheet1.range('B4').value = 'Fecha: '+fulcrum_csv_data[fulcrum_csv_data.columns[1]].iloc[index]
	sheet2.range('B4').value = 'Fecha: '+fulcrum_csv_data[fulcrum_csv_data.columns[1]].iloc[index]
	sheet3.range('B4').value = 'Fecha: '+fulcrum_csv_data[fulcrum_csv_data.columns[1]].iloc[index]
	#Revisión
	sheet1.range('C4').value = 'Revisión: '+fulcrum_csv_data[fulcrum_csv_data.columns[2]].iloc[index]
	sheet2.range('C4').value = 'Revisión: '+fulcrum_csv_data[fulcrum_csv_data.columns[2]].iloc[index]
	sheet3.range('C4').value = 'Revisión: '+fulcrum_csv_data[fulcrum_csv_data.columns[2]].iloc[index]

	#Finalización de la prueba
	#Nombre SICE:
	sheet1.range('B5').value = fulcrum_csv_data[fulcrum_csv_data.columns[32]].iloc[index]
	#Nombre AT 1:
	if fulcrum_csv_data[fulcrum_csv_data.columns[35]].iloc[index]:
		sheet1.range('B6').value = ("Sergio Martín")
	else:
		sheet1.range('B6').value = fulcrum_csv_data[fulcrum_csv_data.columns[34]].iloc[index]
	#Nombre AT 2:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[38]].iloc[index])
	if x:
		sheet1.range('B7').value = ("Ángel Losa")
	else:
		sheet1.range('B7').value = fulcrum_csv_data[fulcrum_csv_data.columns[38]].iloc[index]
	#Nombre MC30:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[43]].iloc[index])
	if x:
		sheet1.range('B8').value = ("Javier Berges")
	else:
		sheet1.range('B8').value = fulcrum_csv_data[fulcrum_csv_data.columns[43]].iloc[index]

	path_photos = r'/Users/jmpas/Desktop/fat/photos/'
	path_firmas = r'/Users/jmpas/Desktop/fat/firmas/'
	
	#Firma SICE:
	sheet1.range('C5').value = fulcrum_csv_data[fulcrum_csv_data.columns[34]].iloc[index]
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[33]].iloc[index])
	if not x:
		sheet1.pictures.add(path_photos+fulcrum_csv_data[fulcrum_csv_data.columns[33]].iloc[index]+'.png', left=300, top=90, scale=0.09)
	#Firma AT 1:
	sheet1.range('C6').value = fulcrum_csv_data[fulcrum_csv_data.columns[37]].iloc[index]
	if fulcrum_csv_data[fulcrum_csv_data.columns[36]].iloc[index]:
		sheet1.pictures.add(path_photos+fulcrum_csv_data[fulcrum_csv_data.columns[36]].iloc[index]+'.png', left=300, top=132, scale=0.1)
	#Firma AT 2:
	sheet1.range('C7').value = fulcrum_csv_data[fulcrum_csv_data.columns[40]].iloc[index]
	if fulcrum_csv_data[fulcrum_csv_data.columns[39]].iloc[index]:
		sheet1.pictures.add(path_photos+fulcrum_csv_data[fulcrum_csv_data.columns[39]].iloc[index]+'.png', left=300, top=173, scale=0.08)
	#Firma MC30:
	sheet1.range('C8').value = fulcrum_csv_data[fulcrum_csv_data.columns[45]].iloc[index]
	sheet1.pictures.add(path_firmas+'berges.png', left=300, top=205, scale=0.1)

	#Registro de pruebas
	#Documento:
	sheet1.range('B10').value = fulcrum_csv_data[fulcrum_csv_data.columns[3]].iloc[index]
	#Código documento:
	sheet1.range('B11').value = fulcrum_csv_data[fulcrum_csv_data.columns[4]].iloc[index]
	#Activo MC30:
	sheet1.range('B12').value = fulcrum_csv_data[fulcrum_csv_data.columns[5]].iloc[index]
	#Tipo de equipo:
	sheet1.range('B13').value = fulcrum_csv_data[fulcrum_csv_data.columns[6]].iloc[index]
	#Dispositivo hardware:
	sheet1.range('B14').value = fulcrum_csv_data[fulcrum_csv_data.columns[7]].iloc[index]
	#Marca y modelo:
	sheet1.range('B15').value = fulcrum_csv_data[fulcrum_csv_data.columns[8]].iloc[index]
	#Capítulo PPT:
	sheet1.range('B16').value = fulcrum_csv_data[fulcrum_csv_data.columns[9]].iloc[index]
	#Referencia presupuesto:
	sheet1.range('B17').value = fulcrum_csv_data[fulcrum_csv_data.columns[10]].iloc[index]

	#Prueba visual (nombre de la sección de Fulcrum)
	#Número de serie:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[11]].iloc[index])
	if not x:
		sheet1.range('B19').value = fulcrum_csv_data[fulcrum_csv_data.columns[11]].iloc[index]
		sheet1.range('C19').value = ("SI")
	else:
		sheet1.range('C19').value = ("NO")
	#MAC:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[12]].iloc[index])
	if not x:
		sheet1.range('B20').value = fulcrum_csv_data[fulcrum_csv_data.columns[12]].iloc[index]
		sheet1.range('C20').value = ("SI")
	else:
		sheet1.range('C20').value = ("NO")
	#IP:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[13]].iloc[index])
	if not x:
		sheet1.range('B21').value = fulcrum_csv_data[fulcrum_csv_data.columns[13]].iloc[index]
		sheet1.range('C21').value = ("SI")
	else:
		sheet1.range('C21').value = ("NO")
	#Version equipo:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[14]].iloc[index])
	if not x:
		sheet1.range('B22').value = fulcrum_csv_data[fulcrum_csv_data.columns[14]].iloc[index]
		sheet1.range('C22').value = ("SI")
	else:
		sheet1.range('C22').value = ("NO")
	#Version FW:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[15]].iloc[index])
	if not x:
		sheet1.range('B23').value = fulcrum_csv_data[fulcrum_csv_data.columns[15]].iloc[index]
		sheet1.range('C23').value = ("SI")
	else:
		sheet1.range('C23').value = ("NO")
	#Marcado CE:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[16]].iloc[index])
	if not x:
		sheet2.pictures.add(path_photos+fulcrum_csv_data[fulcrum_csv_data.columns[16]].iloc[index]+'.jpg', left=140, top=113, scale=0.17)
		sheet1.range('C24').value = ("SI")
		sheet1.range('D24').value = ("Se adjunta fotografía")
	else:
		sheet1.range('C24').value = ("NO")

	#Pruebas funcionales
	#Verificación visual del material acopiado:
	sheet1.range('C26').value = fulcrum_csv_data[fulcrum_csv_data.columns[17]].iloc[index]
	#Observaciones:
	sheet1.range('D26').value = fulcrum_csv_data[fulcrum_csv_data.columns[18]].iloc[index]
	#Comprobación de conexión red eléctrica y alimentador:
	sheet1.range('C27').value = fulcrum_csv_data[fulcrum_csv_data.columns[19]].iloc[index]
	#Observaciones:
	sheet1.range('D27').value = fulcrum_csv_data[fulcrum_csv_data.columns[20]].iloc[index]
	#Comprobación link Ethernet:
	sheet1.range('C28').value = fulcrum_csv_data[fulcrum_csv_data.columns[21]].iloc[index]
	#Observaciones:
	sheet1.range('D28').value = fulcrum_csv_data[fulcrum_csv_data.columns[22]].iloc[index]
	#Comprobación salida de vídeo canal HDMI:
	sheet1.range('C29').value = fulcrum_csv_data[fulcrum_csv_data.columns[23]].iloc[index]
	#Observaciones:
	sheet1.range('D29').value = fulcrum_csv_data[fulcrum_csv_data.columns[24]].iloc[index]
	#Comprobación salida de vídeo canal DP:
	sheet1.range('C30').value = fulcrum_csv_data[fulcrum_csv_data.columns[25]].iloc[index]

	#Observaciones:
	sheet1.range('D30').value = fulcrum_csv_data[fulcrum_csv_data.columns[26]].iloc[index]

	#Fotos pruebas funcionales:
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[27]].iloc[index])
	if not x:
		y_string = fulcrum_csv_data[fulcrum_csv_data.columns[27]].iloc[index]
		y_list = y_string.split(",")
		i = 0
		for i in y_list:
			idx = y_list.index(i)
			if idx==0:
				sheet2.pictures.add(path_photos+y_list[0]+'.jpg', left=140, top=332, scale=0.17)
			
			if idx==1:
				sheet2.pictures.add(path_photos+y_list[1]+'.jpg', left=140, top=530, scale=0.17)
			
			if idx==2:
				sheet3.pictures.add(path_photos+y_list[2]+'.jpg', left=140, top=107, scale=0.17)
			
			if idx==3:
				sheet3.pictures.add(path_photos+y_list[3]+'.jpg', left=140, top=315, scale=0.17)

			if idx==4:
				sheet3.pictures.add(path_photos+y_list[4]+'.jpg', left=140, top=515, scale=0.17)
				
	#Conclusión:
	sheet1.range('A32').value = fulcrum_csv_data[fulcrum_csv_data.columns[29]].iloc[index]
	#Observaciones
	x = False
	x = isNaN(fulcrum_csv_data[fulcrum_csv_data.columns[30]].iloc[index])
	if not x:
		sheet1.range('A33').value = 'Observaciones: '+fulcrum_csv_data[fulcrum_csv_data.columns[30]].iloc[index]
	else:
		sheet1.range('A33').value = 'Observaciones: '


	#Genera y guarda PDF
	filename='ATC_FAT_DEC_%s.pdf'%(fulcrum_csv_data[fulcrum_csv_data.columns[5]].iloc[index])
	#filename='prueba_%d.pdf'%(index)
	#filename = 'ATC_FAT_CPD_DEC'
	wb.to_pdf(filename,include=['Hoja1','Hoja2','Hoja3'])

	#Cierra Excel
	wb.close()





