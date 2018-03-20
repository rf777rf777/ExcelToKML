import pandas
import simplekml
import sys
import os

def findXlsxFilePath(appPath):	
	fileName = ''
	for file in os.listdir(appPath):
		if (file.lower()).endswith('xlsx'):
			fileName = file
	path = appPath + '/{}'.format(fileName)

	return path

def createKML(filePath):
	excel = pandas.ExcelFile(filePath)
	sheetNamesList = excel.sheet_names
	kml = simplekml.Kml()
	doc = kml.newdocument(name = "Test")
	layers = []
	
	for sheet_name in sheetNamesList:
		layers.append(doc.newfolder(name= sheet_name))


	sharedstyle = simplekml.Style()
	sharedstyle.labelstyle.scale = 0.85
	sharedstyle.iconstyle.scale = 1
	sharedstyle.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'

	pnt = []
	complete = 0
	for layersItem in layers:
		data = excel.parse(layersItem.name).fillna('')
		for i in range(len(data)):
			point = layersItem.newpoint(name = data.ix[i]['站名'],
								 		description = data.ix[i]['所在鄉鎮']	,	
								 		coords = [(data.ix[i]['E'],data.ix[i]['N'])])
		complete += 1
		completeInTerminal(complete,len(layers))
	return kml

def completeInTerminal(complete,total):
	done = int(30 * complete / total)
	sys.stdout.write("\rCompleted：[{0}{1}] {2}%".format('#'*done, '-'*(30 - done) , int(complete/total*100)))
	sys.stdout.flush()

def main():
	#if執行exe檔
	if getattr(sys, 'frozen', False):
		filePath = findXlsxFilePath(os.path.dirname(sys.executable))
	#if直接執行.py檔
	elif __file__:
		filePath = findXlsxFilePath(os.path.abspath(os.path.dirname(__file__)))
	try:
		kml = createKML(filePath)		
		kml.save('{}.kml'.format(filePath.split('.')[0]))		
	except Exception as e:
		print('發生例外狀況：{0} \n程式停止！'.format(e))


if __name__ == "__main__":
    main()
    if os.name == 'nt':
    	input("\n<<< 按任意鍵關閉視窗 >>>")