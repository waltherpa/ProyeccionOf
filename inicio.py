from methods.method import prepararData, actualizarArchivo

# filename='7 ofertas.xlsx'
filename='7 ofertas.xlsx'
path='C:/Users/wrojas/Documents/motos/walther/2.- work/7.- datalake/datale files/'
path2='C:/Users/wrojas/Documents/motos/walther/2.- work/14.- flujos/9.- proyeccion motos/'
data= prepararData(file_name=path+filename) # conseguir información y prepararlo

actualizarArchivo(data=data, sede='Honda SM Retail', pathToSave=path2)
actualizarArchivo(data=data, sede='Honda SQ Retail', pathToSave=path2)