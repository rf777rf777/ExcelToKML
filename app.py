import pandas
import simplekml

fileName = '屏東_已編輯'

excel = pandas.ExcelFile('{}.xlsx'.format(fileName))

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
		

kml.save('456{}.kml'.format(fileName))