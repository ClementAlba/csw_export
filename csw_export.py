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
try:
    with open(input("Veuillez entrer le chemin de votre fichier de configuration : \n")) as jsonFile:
        jsonObject = json.load(jsonFile)
        jsonFile.close()
except FileNotFoundError:
    print("ERREUR : Le fichier de configuration que vous indiquez n'a pas été trouvé ou n'existe pas. \n")
except json.decoder.JSONDecodeError:
    print("ERREUR : Une erreur a eu lieu lors du décodage du fichier de configuration."
          " Assurez-vous que le fichier ne comporte pas d'erreur de syntaxe et qu'il ne soit pas vide. \n")


# Fonction qui permet de récupérer toutes les fiches des deux catalogues ainsi que leurs informations
def getRecords(url):
    with open(list(filter(lambda elt: elt["type"] == "fichier", jsonObject))[0].get("paths").get("request")) as xml:
        # Requête qui va permettre de récupérer les fiches
        response = requests.post(url, data=xml.read(),
                                             headers={"Content-Type": "text/xml"}, verify=False)
        if response.status_code == 404:
            print("L'url que vous indiquez n'existe pas.")

    # Traitement de la réponse
    root = ET.fromstring(response.content.decode("utf-8"))
    try:
        return root.findall(list(filter(lambda elt: elt["type"] == "balise_xml", jsonObject))[0].get("tags").get("record_tag_path"))
    except TypeError:
        print("ERREUR : Assurez-vous que la clé que vous recherchez existe bel et bien.")

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

    if records:
        for record in records:
            id = record.find(list(filter(lambda elt: elt["type"] == "balise_xml", jsonObject))[0].get("tags").get("identifier"))
            title = record.find(list(filter(lambda elt: elt["type"] == "balise_xml", jsonObject))[0].get("tags").get("title"))

            # Écriture de l'id et du titre dans la feuille de travail excel
            worksheet.write(row, 0, title.text)
            worksheet.write(row, 1, url + id.text.split(':')[-1])

            row += 1
    else:
        print("La variable records est vide.")

    workbook.close()


# Appel des fonctions
try:
    recordsDig = getRecords(list(filter(lambda elt: elt["type"] == "catalogue" and elt["name"] == "dig", jsonObject))[0].get("urls").get("records"))
    writeInExcelFile(recordsDig,
                     list(filter(lambda elt: elt["type"] == "fichier", jsonObject))[0].get("paths").get("excel_dig"),
                     list(filter(lambda elt: elt["type"] == "catalogue" and elt["name"] == "dig", jsonObject))[0].get("urls").get("fiche"))

    recordsAgents = getRecords(list(filter(lambda elt: elt["type"] == "catalogue" and elt["name"] == "agents_mel", jsonObject))[0].get("urls").get("records"))
    writeInExcelFile(recordsAgents,
                     list(filter(lambda elt: elt["type"] == "fichier", jsonObject))[0].get("paths").get("excel_agents"),
                     list(filter(lambda elt: elt["type"] == "catalogue" and elt["name"] == "agents_mel", jsonObject))[0].get("urls").get("fiche"))
except NameError:
    print("ERREUR : Une variable que vous essayez d'utiliser pour l'appel des fonctions n'a pas été définie.")