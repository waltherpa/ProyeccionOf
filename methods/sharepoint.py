from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import pandas as pd


def loadShareFile(filename):
    authcookie = Office365('https://panamotorssac.sharepoint.com', username='wrojas@panaautos.com.pe', password='Panaautos22').GetCookies()
    site = Site('https://panamotorssac.sharepoint.com/sites/GP_Motos_Honda', version=Version.v365, authcookie=authcookie)
    
    folder = site.Folder('Documentos Compartidos/zProyecciones')
    return folder.get_file(filename)


def saveSharedFile(thefile):
    authcookie = Office365('https://panamotorssac.sharepoint.com', username='wrojas@panaautos.com.pe', password='Panaautos22').GetCookies()
    site = Site('https://panamotorssac.sharepoint.com/sites/GP_Motos_Honda', version=Version.v365, authcookie=authcookie)
    folder = site.Folder('Documentos Compartidos/z_prueba')

    with open(thefile, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, "archivodesharepoint.xlsx")

