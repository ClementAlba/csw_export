# Librarie qui va nous permettre de charger le fichier de configuration JSON
import json
# Librairie qui permet de traiter la réponse XML afin d'en extraire les balises que l'on souhaite
import xml.etree.ElementTree as ET
# Librairie qui permet de créer et d'écrire dans un fichier excel
import xlsxwriter
# Librairie qui permet de faire des requêtes HTTP, on va l'utiliser pour récupérer les fiches des catalogues ISOGEO
import requests

# Le but de ce script est de récupérer toutes les fiches des différents catalogues ISOGEO de la MEL et de mettre en forme certaines
# informations dans des fichiers excel.

# Ouverture et chargement du fichier de configuration json qui sera utilisé via la variable jsonObject
with open(input("Veuillez entrer le chemin de votre fichier de configuration : \n")) as jsonFile:
    jsonObject = json.load(jsonFile)
    jsonFile.close()


# Fonction qui permet de récupérer toutes les fiches des deux catalogues ainsi que leurs informations
def getRecords(url):
    with open(jsonObject.get("fichier").get("request")) as xml:
        # Requête qui va permettre de récupérer les fiches
        response = requests.post(url, data=xml.read(),
                                             headers={"Content-Type": "text/xml"}, verify=False)

    # Traitement de la réponse
    root = ET.fromstring(response.content.decode("utf-8"))
    return root.findall(jsonObject.get("balise_xml").get("record_tag_path"))


# Fonction qui permet de créer les fichiers excel et d'écrire les informations sur les fiches dedans
def writeInExcelFile(records, filePath, url):
    row = 1
    # Création du fichier excel et ajout d'une feuille de travail
    workbook = xlsxwriter.Workbook(filePath)
    worksheet = workbook.add_worksheet()

    # On prépare ici la mise en forme du fichier excel en écrivant "Titre" dans la première cellule et "URL" dans la seconde
    worksheet.write(0, 0, 'Titre', workbook.add_format({'bold': True}))
    worksheet.write(0, 1, 'URL', workbook.add_format({'bold': True}))
    # On défini la taille des colonnes pour que le titre et les urls entrent entièrement dedans
    worksheet.set_column(0, 0, 140)
    worksheet.set_column(1, 1, 120)

    for record in records:
        id = record.find(jsonObject.get("balise_xml").get("identifier"))
        title = record.find(jsonObject.get("balise_xml").get("title"))

        # Écriture de l'id et du titre dans la feuille de travail excel
        worksheet.write(row, 0, title.text)
        worksheet.write(row, 1, url + id.text.split(':')[-1])

        row += 1

    workbook.close()


# Appel des fonctions
# Dans l'appel à getRecords, on utilise deux fois la fonction findJsonObject. On fait ça car il y a deux éléments qui ont le type
# catalogue dans le fichier de config et il est donc nécessaire de les identifier par leur nom pour obtenir l'élément que l'on souhaite.
recordsDig = getRecords(jsonObject.get("catalogue").get("nom_catalogue").get("records"))
writeInExcelFile(recordsDig,
                 jsonObject.get("fichier").get("nom_fichier"),
                 jsonObject.get("catalogue").get("nom_catalogue").get("fiche"))

recordsAgents = getRecords(jsonObject.get("catalogue").get("nom_catalogue").get("records"))
writeInExcelFile(recordsAgents,
                 jsonObject.get("fichier").get("nom_fichier"),
                 jsonObject.get("catalogue").get("nom_catalogue").get("fiche"))