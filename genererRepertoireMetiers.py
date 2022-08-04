import json
import re
import threading
import tkinter as tk
from tkinter import filedialog as fd
from types import NoneType

import pandas as pd
import requests
from bs4 import BeautifulSoup

JSON_FILE_PATH = "jobsData.json"

EXCEL_SECTOR_HEADER = "N° secteur"
EXCEL_JOB_HEADER = "N° métiers"
EXCEL_SKILL_HEADER = "Numéro de compétences"

EXCEL_DATA_HEADERS = {
    "chemicalRisks": "1. Risques Chimiques",
    "biologicalRisks": "2. Risques Biologiques",
    "equipmentRisks": "3. Risques liés aux machines et aux équipements",
    "fallRisks": "4. Risques de chutes de hauteur et de plain-pied",
    "objectFall": "5. Risques liés aux chutes d’objets",
    "transitRisks": " 6. Risques liés aux déplacements",
    "postureRisks": " 7. Risques liés aux postures contraignantes",
    "motionRisks": "8. Risques liés aux mouvements répétitifs, pressions de contact et chocs",
    "handlingRisks": "9. Risques liés à la manutention",
    "psycologicalRisks": "10. Risques psychosociaux et de violence",
    "noiseRisks": "11. Risques liés au bruit",
    "temperatureRisks": "12. Risques liés à l'Froid et chaleur",
    "vibrationRisks": "13. Risques liés aux vibrations",
    "otherRisks": "14. Autres risques (électrique, explosion, travail en espace clos)",
}


def run():
    '''Run the main script in a new thread'''

    def target():
        entry["state"] = "disabled"
        fileButton["state"] = "disabled"
        startButton["state"] = "disabled"
        start(excelPath.get())
        entry["state"] = "normal"
        fileButton["state"] = "normal"
        startButton["state"] = "normal"

    threading.Thread(target=target).start()


def start(excelPath: str):
    '''Starts the main script'''
    json = {}

    try:
        excel = pd.read_excel(excelPath)

    except FileNotFoundError:
        setMessage("Fichier invalide")
        return

    sectors = getSectors()
    for sector in sectors:
        jobs = {}
        for jobID in getJobIDsOfSector(sector["id"], sector["value"]):
            job = getJob(jobID)
            for skillCode in job["skills"]:
                dataToAdd = getSkillDataFromExcel(
                    excel, sector["value"], job["code"], skillCode)
                job["skills"][skillCode].update(dataToAdd)

            jobs[job["code"]] = job

        json[sector["value"]] = {
            "name": sector["name"],
            "jobs": jobs,
        }

    saveJson(json, JSON_FILE_PATH)
    setMessage("Tout est fini!")


def getSkillDataFromExcel(excel: pd.DataFrame, sector, job, skill):
    '''Returns the corresponding data contained in the excel file'''
    result = {}
    row = excel.loc[(excel[EXCEL_SECTOR_HEADER] == int(sector)) & (
        excel[EXCEL_JOB_HEADER] == int(job)) & (excel[EXCEL_SKILL_HEADER] == int(skill))]

    for header, excelHeader in EXCEL_DATA_HEADERS.items():
        if (row[excelHeader].index.size > 0):
            result[header] = row[excelHeader].get(
                row[excelHeader].index[0], None)

    return result


def getSectors():
    '''Returns all the available sectors.'''
    result = []
    nameRegex = re.compile(r"^\d+ - (\D*)$")
    setMessage("Getting all sectors...")

    page = requests.get(
        "http://www1.education.gouv.qc.ca/sections/metiers/index.asp")
    soup = BeautifulSoup(page.content, "html.parser")

    for input in soup.find_all("input", type="checkbox"):
        result.append(
            {
                "id": input["id"],
                "value": input["value"],
                "name": nameRegex.match(input.find_next_sibling("label").text).group(1),
            }
        )

    return result


def getJobIDsOfSector(id: str, value: str):
    '''Returns all the jobs of a particular sector.'''
    result = []
    hrefRegex = re.compile(r"^index\.asp\?.*id=(\d+)")

    setMessage(f"Getting all jobs of {id}...")
    page = requests.get(
        f"http://www1.education.gouv.qc.ca/sections/metiers/index.asp?page=recherche&action=search&navSeq=1&{id}={value}"
    )
    soup = BeautifulSoup(page.content, "html.parser")

    for job in soup.find_all("a", href=hrefRegex):
        result.append(hrefRegex.match(job["href"]).group(1))

    return result


def getJob(id: str):
    '''Returns a detailed job.'''
    titleRegex = re.compile(r"(\d+) - ([^\t\r\n]*)")

    page = requests.get(
        f"http://www1.education.gouv.qc.ca/sections/metiers/index.asp?page=fiche&id={id}"
    )
    soup = BeautifulSoup(page.content, "html.parser")

    [jobCode, jobName] = soup.find("h2").getText(";", True).split(";")
    result = {"name": jobName, "code": jobCode, "skills": {}}

    for header in soup.find_all("thead"):
        titleSearch = titleRegex.search(header.find("th").text)
        skillCode = titleSearch.group(1)
        skill = titleSearch.group(2)

        lists = header.find_next_sibling("tbody").find_all("ul")

        criteria = []
        for criterion in lists[0].find_all("li"):
            criteria.append(
                cleanUpText(criterion.text))

        tasks = []
        for task in lists[1].find_all("li"):
            tasks.append(cleanUpText(task.text))

        result["skills"][skillCode] = {
            "name": skill, "code": skillCode, "criteria": criteria, "tasks": tasks}

    return result


def cleanUpText(text: str):
    '''Removes unwanted formating chars at the end of [text].'''
    return re.match(r"[^\t\r\n]*", text).group(0)


def cleanUpData(data):
    if isinstance(data, list):
        return [cleanUpData(x) for x in data if x is not None]
    elif isinstance(data, dict):
        return {key: cleanUpData(val) for key, val in data.items() if val is not None}
    else:
        return data


def saveJson(data: dict, path: str):
    '''Saves [json] as a file named [path] formated with an indent of 4.'''
    setMessage("Saving json...")
    with open(path, "w") as file:
        file.write(json.dumps(cleanUpData(data), indent=4))


def setMessage(message: str):
    currentMessage.set(message)


def askExcelPath():
    file = fd.askopenfile(title="Choisir un classeur Excel", filetypes=(
        ("Classeurs Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")))

    if (not isinstance(file, NoneType)):
        excelPath.set(file.name)


# Tkinter initialisation
root = tk.Tk()
root.title("CRCRME - Générer répertoire métiers")
root.geometry("450x140")
root.resizable(False, False)
mainFrame = tk.Frame(root)
mainFrame.pack(padx=20, pady=20)

label = tk.Label(
    mainFrame, text="Entrez le chemin d'accès d'un classeur Excel.")
label.pack()

frame = tk.Frame(mainFrame)
frame.pack()

excelPath = tk.StringVar()
entry = tk.Entry(frame, textvariable=excelPath)
entry.focus()
entry.pack(side="left")

fileButton = tk.Button(frame, text="Parcourir", command=askExcelPath)
fileButton.pack(side="right")

startButton = tk.Button(mainFrame, text="Générer", command=run)
startButton.pack(side="bottom")

currentMessage = tk.StringVar()
messageLabel = tk.Label(mainFrame, textvariable=currentMessage)
messageLabel.pack(side="bottom")

if __name__ == "__main__":
    root.mainloop()
