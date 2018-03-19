import pandas
import simplekml
import sys
import os

application_path = os.path.dirname(sys.executable)

fileName = ''
#巡覽照片資料夾
for file in os.listdir(application_path):
    if (file.lower()).endswith('xlsx'):
        fileName = file
path = application_path+'/'+fileName
excel = pandas.ExcelFile(path)

sheetNamesList = excel.sheet_names 
#df = pandas.read_excel("屏東_已編輯.xlsx").fillna('')

kml = simplekml.Kml()
doc = kml.newdocument(name = fileName)
layers = []
#for i in range(len(位置集合)):
    #fld.append(doc.newfolder(name = 位置集合[i]))
for sheet_name in sheetNamesList:
	layers.append(doc.newfolder(name= sheet_name))
'''
for item in layers:
	print(type(item))
pass
'''
sharedstyle = simplekml.Style()
sharedstyle.labelstyle.scale = 0.85
sharedstyle.iconstyle.color = 'fffff8f0'
sharedstyle.iconstyle.scale = 1
sharedstyle.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

#dataList = []

pnt = []
for layersItem in layers:
	data = excel.parse(layersItem.name).fillna('')
	for i in range(len(data)):
		point = layersItem.newpoint(name = data.ix[i]['站名'],
							 		description = data.ix[i]['所在鄉鎮']	,	
							 		coords = [(data.ix[i]['E'],data.ix[i]['N'])])
		

kml.save(application_path+"/"+'{}.kml'.format(fileName.split('.')[0]))