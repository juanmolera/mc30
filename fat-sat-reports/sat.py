__author__ = "Juan Molera Pascual"
__copyright__ = "Alauda Ingeniería SA"
__credits__ = ["Juan Molera Pascual"]
__license__ = "2022"
__version__ = "1.0.4"
__maintainer__ = "Juan Molera Pascual"
__email__ = "jmolera@alaudaingenieria.es"
__status__ = "wip"

'''
Definitivo
Expandir las hojas en pruebas
'''

#Librerías
import pandas as pd
import xlwings as xw
import os
import csv

#Limpiamos cmd en cada ejecución
os.system("clear")

#Path
cwd=os.getcwd()

#Método para checkear si falta un dato
def isNaN(string):
    return string != string

#Configuración de Pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

#Leemos el csv con los datos de Fulcrum
#NO MODIFICAR EL CSV ORIGINAL
fulcrum_csv_data=pd.read_csv("prueba_sat_hw_decodificadores.csv")

#Cálculo de límites de los datos del csv
campos_fulcrum = 14
limite_inferior_fij = campos_fulcrum - 1
sat_date_rev = 3
fijos = 8
#*******************************************************************
#n1 = int(input('NÚMERO DE PRUEBAS VISUALES: '))
#pruebas_visuales = int(n1)
pruebas_visuales = 4
otros = 6 #hay 2 pruebas fantasma
limite_superior_fij = limite_inferior_fij + sat_date_rev + fijos + otros
#n2 = int(input('NÚMERO DE PRUEBAS FUNCIONALES: '))
#pruebas_funcionales = int(n2*2)
pruebas_funcionales = 12 #ES EL # DE PRUEBASx2
fotos_pruebas = 1
#*******************************************************************
suma = campos_fulcrum + sat_date_rev + fijos + otros
limite_inferior_var = suma - 1
limite_superior_var = limite_inferior_var + pruebas_funcionales + fotos_pruebas

#Lista de campos fijos, se incluyen las pruebas visuales, las conclusiones y las firmas
lista_fijos=list()
for element in enumerate(fulcrum_csv_data):
	if element[0] > limite_inferior_fij:
		lista_fijos.append(element[1])
		if element[0] == limite_superior_fij:
			break

#Lista de pruebas visuales
lista_visuales=list()
limite_inferior_vis = campos_fulcrum + sat_date_rev + fijos - 1
limite_superior_vis = limite_inferior_vis + otros
for element in enumerate(fulcrum_csv_data):
	if element[0] > limite_inferior_vis:
		lista_visuales.append(element[1])
		if element[0] == limite_superior_vis:
			break

#Lista de conclusiones
lista_conclusiones = list()
lista_conclusiones = ['conclusion','observaciones2','foto_del_elemento']

#Lista de firmas
lista_firmas = list()
lista_firmas = ['nombre_tecnico_sice','firma_especialista',
'nombre_asistencia_tecnica','firma_asistencia_tecnica',
'nombre_asistencia_tecnica2','firma_asistencia_tecnica2',
'nombre_direccion_obra','firma_direccion_obra',
'firma_especialista_timestamp','firma_asistencia_tecnica_timestamp',
'firma_asistencia_tecnica2_timestamp','firma_direccion_obra_timestamp']

lista_unida = lista_fijos + lista_conclusiones + lista_firmas

#Creación del csv de campos fijos
with open('fijos_sat.csv', 'w') as csvfile:
    writer = csv.writer(csvfile, delimiter=",")
    writer.writerow(lista_unida)

#Lista de pruebas funcionales
lista_pruebas=list()
for element in enumerate(fulcrum_csv_data):
	if element[0] > limite_inferior_var:
		lista_pruebas.append(element[1])
		if element[0] == limite_superior_var:
			break

#Creación del csv de pruebas funcionales
with open('variables_sat.csv', 'w') as csvfile:
    writer = csv.writer(csvfile, delimiter=",")
    writer.writerow(lista_pruebas)

#Lee el csv que contiene los campos fijos
fijos_csv_data=pd.read_csv("fijos_sat.csv")

#Lee el csv que contiene las pruebas funcionales
variables_csv_data=pd.read_csv("variables_sat.csv")

#Pegamos las listas
frames=[fijos_csv_data,variables_csv_data]
result = pd.concat(frames)

#Eliminamos los campos que genera Fulcrum y no interesan
drop_list=list()
for k in fulcrum_csv_data.columns:
	if k not in result.columns:
		drop_list.append(k)

#Se guardan los datos de csv que nos interesan
fulcrum_csv_data.drop(columns=drop_list, inplace=True)
#print(fulcrum_csv_data.columns)

#Definición de celdas de la plantilla Excel
#ESTA HARDCODEADO PERO SI NO CAMBIAN FIJOS ESTO NO CAMBIA
#*******************************************************************
inicio_num_pruebas_visuales = 19
#*******************************************************************
fin_num_pruebas_visuales = inicio_num_pruebas_visuales + pruebas_visuales - 1
inicio_num_datos_pruebas_fun = fin_num_pruebas_visuales + 2
fin_num_datos_pruebas_fun = int(inicio_num_datos_pruebas_fun + (pruebas_funcionales/2) - 1)
celda_conclusion = fin_num_datos_pruebas_fun + 2
celda_obs_conclusion = celda_conclusion + 1

letra_conclusion = 'A'
letra_datos_pruebas = 'B'
letra_si_no_pruebas = 'C'
letra_obs_pruebas = 'D'

#Listas de celdas para pruebas visuales
lista_celdas_datos_pruebas_vis = list()
lista_celdas_si_no_pruebas_vis = list()
lista_celdas_obs_pruebas_vis = list()
for k in range(pruebas_visuales):
	num_celda = inicio_num_pruebas_visuales + k
	celda_datos = letra_datos_pruebas + str(num_celda)
	lista_celdas_datos_pruebas_vis.append(celda_datos)
	celda_si_no = letra_si_no_pruebas + str(num_celda)
	lista_celdas_si_no_pruebas_vis.append(celda_si_no)
	celda_obs = letra_obs_pruebas + str(num_celda)
	lista_celdas_obs_pruebas_vis.append(celda_obs)

#Listas de celdas para pruebas funcionales
lista_celdas_datos_pruebas_fun = list()
lista_celdas_si_no_pruebas_fun = list()
lista_celdas_obs_pruebas_fun = list()
for k in range(int(pruebas_funcionales/2)):
	num_celda = inicio_num_datos_pruebas_fun + k
	celda_datos = letra_datos_pruebas + str(num_celda)
	lista_celdas_datos_pruebas_fun.append(celda_datos)
	celda_si_no = letra_si_no_pruebas + str(num_celda)
	lista_celdas_si_no_pruebas_fun.append(celda_si_no)
	celda_obs = letra_obs_pruebas + str(num_celda)
	lista_celdas_obs_pruebas_fun.append(celda_obs)

#Lista celdas para conclusión y observaciones
for k in range(fotos_pruebas):
	num_celda_conclusion = num_celda + 2
	celda_conclusion = letra_conclusion + str(num_celda_conclusion)
	lista_celdas_datos_pruebas_fun.append(celda_conclusion)

num_celda_obs = num_celda_conclusion + 1
celda_obs = letra_conclusion + str(num_celda_obs)

lista_celdas_datos_pruebas_fun.append(celda_obs)

#Excel
for index, row in fulcrum_csv_data.iterrows():
	print(index)

	#Selección de plantilla
	wb = xw.Book("formato_sat.xlsx")
	sheet1 = wb.sheets['Hoja1']
	sheet2 = wb.sheets['Hoja2']
	sheet3 = wb.sheets['Hoja3']

	#*******************************************************************
	#Paths imágenes
	path_photos = r'/Users/jmpas/Desktop/sat/photos_sat_decos/'
	path_firmas = r'/Users/jmpas/Desktop/sat/firmas/'
	#*******************************************************************

	#Campos de la plantilla
	#Registro de pruebas
	#Número de página:
	sheet1.range('A4').value = 'Página: 1/3'
	sheet2.range('A4').value = 'Página: 2/3'
	sheet3.range('A4').value = 'Página: 3/3'
	#Fecha:
	sheet1.range('B4').value = 'Fecha: '+fulcrum_csv_data['date'].iloc[index]
	sheet2.range('B4').value = 'Fecha: '+fulcrum_csv_data['date'].iloc[index]
	sheet3.range('B4').value = 'Fecha: '+fulcrum_csv_data['date'].iloc[index]
	#Revisión
	sheet1.range('C4').value = 'Revisión: '+fulcrum_csv_data['revision'].iloc[index]
	sheet2.range('C4').value = 'Revisión: '+fulcrum_csv_data['revision'].iloc[index]
	sheet3.range('C4').value = 'Revisión: '+fulcrum_csv_data['revision'].iloc[index]

	#Finalización de la prueba
	#Nombre SICE:
	x = False
	x = isNaN(fulcrum_csv_data['nombre_tecnico_sice'].iloc[index])
	if x:
		sheet1.range('B5').value = ("Adrián Vicente")
	else:
		sheet1.range('B5').value = fulcrum_csv_data['nombre_tecnico_sice'].iloc[index]
	#Nombre AT 1:
	x = False
	x = isNaN(fulcrum_csv_data['nombre_asistencia_tecnica'].iloc[index])
	if x:
		sheet1.range('B6').value = ("Sergio Martín")
	else:
		sheet1.range('B6').value = fulcrum_csv_data['nombre_asistencia_tecnica'].iloc[index]
	#Nombre AT 2:
	x = False
	x = isNaN(fulcrum_csv_data['nombre_asistencia_tecnica2'].iloc[index])
	if x:
		sheet1.range('B7').value = ("Ángel Losa")
	else:
		sheet1.range('B7').value = fulcrum_csv_data['nombre_asistencia_tecnica2'].iloc[index]
	#Nombre MC30:
	x = False
	x = isNaN(fulcrum_csv_data['nombre_direccion_obra'].iloc[index])
	if x:
		sheet1.range('B8').value = ("Javier Berges")
	else:
		sheet1.range('B8').value = fulcrum_csv_data['nombre_direccion_obra'].iloc[index]
	
	#Firma SICE:
	x = False
	x = isNaN(fulcrum_csv_data['firma_especialista'].iloc[index])
	if not x:
		sheet1.pictures.add(path_photos+fulcrum_csv_data['firma_especialista'].iloc[index]+'.png', left=300, top=90, scale=0.09)
	#Timestamp SICE:
	sheet1.range('C5').value = fulcrum_csv_data['firma_especialista_timestamp'].iloc[index]
	#Firma AT 1:
	x = False
	x = isNaN(fulcrum_csv_data['firma_asistencia_tecnica'].iloc[index])
	if not x:
		sheet1.pictures.add(path_photos+fulcrum_csv_data['firma_asistencia_tecnica'].iloc[index]+'.png', left=300, top=132, scale=0.1)
	#Timestamp AT 1:
	sheet1.range('C6').value = fulcrum_csv_data['firma_asistencia_tecnica_timestamp'].iloc[index]
	#Firma AT 2:
	x = False
	x = isNaN(fulcrum_csv_data['firma_asistencia_tecnica2'].iloc[index])
	if not x:
		sheet1.pictures.add(path_photos+fulcrum_csv_data['firma_asistencia_tecnica2'].iloc[index]+'.png', left=300, top=173, scale=0.08)
	#Timestamp AT 2:
	sheet1.range('C7').value = fulcrum_csv_data['firma_asistencia_tecnica2_timestamp'].iloc[index]
	#Firma MC30:
	sheet1.pictures.add(path_firmas+'berges.png', left=300, top=205, scale=0.1)
	#Timestamp MC30:
	sheet1.range('C8').value = fulcrum_csv_data['firma_direccion_obra_timestamp'].iloc[index]

	#Registro de pruebas
	#Documento:
	sheet1.range('B10').value = fulcrum_csv_data['documento'].iloc[index]
	#Código documento:
	sheet1.range('B11').value = fulcrum_csv_data['codigo_documento'].iloc[index]
	#Activo MC30:
	sheet1.range('B12').value = fulcrum_csv_data['activo_mc30'].iloc[index]
	#Tipo de equipo:
	sheet1.range('B13').value = fulcrum_csv_data['tipo_de_equipo'].iloc[index]
	#Dispositivo hardware:
	sheet1.range('B14').value = fulcrum_csv_data['dispositivo_hardware'].iloc[index]
	#Marca y modelo:
	sheet1.range('B15').value = fulcrum_csv_data['marca_y_modelo'].iloc[index]
	#Capítulo PPT:
	sheet1.range('B16').value = fulcrum_csv_data['capitulo_ppt'].iloc[index]
	#Referencia presupuesto:
	sheet1.range('B17').value = fulcrum_csv_data['referencia_presupuesto'].iloc[index]

	#Prueba visual (nombre de la sección de Fulcrum)
	#Número de serie:
	for k in range(int(otros)):
		if k == 1 or k == 5:
			pass
		else:
			x = False
			x = isNaN(fulcrum_csv_data[lista_visuales[k]].iloc[index])
			h=0
			if k == 0:
				h = 0
			if k == 2:
				h = 1
			if k == 3:
				h = 2
			if k == 4:
				h = 3
			if not x:
				sheet1.range(lista_celdas_datos_pruebas_vis[h]).value = fulcrum_csv_data[lista_visuales[k]].iloc[index]
				sheet1.range(lista_celdas_si_no_pruebas_vis[h]).value = ("SI")
			else:
				sheet1.range(lista_celdas_si_no_pruebas_vis[h]).value = ("NO")
	
	#Pruebas funcionales
	aux = 0
	for k in range(int(pruebas_funcionales/2)):
		if k == 0:
			#Prueba funcional:
			celda_si_no = lista_celdas_si_no_pruebas_fun[k]
			campo_si_no = lista_pruebas[k]
			sheet1.range(celda_si_no).value = fulcrum_csv_data[campo_si_no].iloc[index]
			#Observaciones:
			celda_obs = lista_celdas_obs_pruebas_fun[k]
			campo_obs = lista_pruebas[k+1]
			sheet1.range(celda_obs).value = fulcrum_csv_data[campo_obs].iloc[index]
			aux = k + 1
		else:
			#Prueba funcional:
			celda_si_no = lista_celdas_si_no_pruebas_fun[k]
			campo_si_no = lista_pruebas[aux+1]
			sheet1.range(celda_si_no).value = fulcrum_csv_data[campo_si_no].iloc[index]
			#Observaciones:
			celda_obs = lista_celdas_obs_pruebas_fun[k]
			campo_obs = lista_pruebas[aux+2]
			sheet1.range(celda_obs).value = fulcrum_csv_data[campo_obs].iloc[index]

	#Fotos pruebas funcionales:
	campo_foto_pruebas = len(lista_pruebas)-1
	x = False
	x = isNaN(fulcrum_csv_data[lista_pruebas[campo_foto_pruebas]].iloc[index])
	if not x:
		y_string = fulcrum_csv_data[lista_pruebas[campo_foto_pruebas]].iloc[index]
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
	celda_conclusion = lista_celdas_datos_pruebas_fun[len(lista_celdas_datos_pruebas_fun)-2]
	sheet1.range(celda_conclusion).value = fulcrum_csv_data['conclusion'].iloc[index]
	#Observaciones
	celda_observaciones2 = lista_celdas_datos_pruebas_fun[len(lista_celdas_datos_pruebas_fun)-1]
	x = False
	x = isNaN(fulcrum_csv_data['observaciones2'].iloc[index])
	if not x:
		sheet1.range(celda_observaciones2).value = 'Observaciones: '+fulcrum_csv_data['observaciones2'].iloc[index]
	else:
		sheet1.range(celda_observaciones2).value = 'Observaciones: '
	#Foto del elemento
	x = False
	x = isNaN(fulcrum_csv_data['foto_del_elemento'].iloc[index])
	if not x:
		sheet2.pictures.add(path_photos+fulcrum_csv_data['foto_del_elemento'].iloc[index]+'.jpg', left=140, top=113, scale=0.17)

	#Genera y guarda PDF
	filename='ATC_SAT_DEC_%s.pdf'%(fulcrum_csv_data[fulcrum_csv_data.columns[5]].iloc[index])
	#filename='prueba_%d.pdf'%(index)
	wb.to_pdf(filename,include=['Hoja1','Hoja2','Hoja3'])

	#Cierra Excel
	wb.close()




