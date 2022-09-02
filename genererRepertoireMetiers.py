import json
import re
import threading
import tkinter as tk
from tkinter import filedialog as fd

# Imports to install
# > pip install pandas requests bs4 openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup


# Constants
JSON_FILE_PATH = "jobs-data.json"

EXCEL_SECTOR_HEADER = "N° secteur"
EXCEL_JOB_HEADER = "N° métiers"
EXCEL_SKILL_HEADER = "Numéro de compétences"

# Data to change
EXCEL_SST_DATA_HEADERS = {
    # json : excel
    "chemical": "1. Risques Chimiques",
    "biological": "2. Risques Biologiques",
    "equipment": "3. Risques liés aux machines et aux équipements",
    "fall": "4. Risques de chutes de hauteur et de plain-pied",
    "objectFall": "5. Risques liés aux chutes d’objets",
    "transit": " 6. Risques liés aux déplacements",
    "posture": " 7. Risques liés aux postures contraignantes",
    "motion": "8. Risques liés aux mouvements répétitifs, pressions de contact et chocs",
    "handling": "9. Risques liés à la manutention",
    "psycological": "10. Risques psychosociaux et de violence",
    "noise": "11. Risques liés au bruit",
    "temperature": "12. Risques liés à l'Froid et chaleur",
    "vibration": "13. Risques liés aux vibrations",
    "electric": "14.1 Risques électriques",
    "anoxia": "14.2 Risque anoxie et travail en espace clos",
    "fire": "14.3 Risque ATEX,  incendie ou explosion",
    "nanomaterial": "14.4 Risques nanomatériaux ",
}

EXCEL_STAGE_DATA_HEADERS = {
    # json : excel
    "1": "Q1",
    "2": "Q2",
    "3": "Q3",
    "4": "Q4",
    "5": "Q5",
    "6": "Q6",
    "7": "Q7",
    "8": "Q8",
    "9": "Q9",
    "10": "Q10",
    "11": "Q11",
    "12": "Q12",
    "13": "Q13",
    "14": "Q14",
    "15": "Q15",
    "16": "Q16",
    "17": "Q17",
    "18": "Q18",
    "19": "Q19",
    "20": "Q20",
    "21": "Q21",
    "22": "Q22",
    "23": "Q23",
}


# Main functions
def run():
    '''Run the main script in a new thread'''

    def target():
        entrySST["state"] = "disabled"
        entryStage["state"] = "disabled"
        fileButtonSST["state"] = "disabled"
        fileButtonStage["state"] = "disabled"
        startButton["state"] = "disabled"
        start(excelPathSST.get(), excelPathStage.get())
        entrySST["state"] = "normal"
        entryStage["state"] = "normal"
        fileButtonSST["state"] = "normal"
        fileButtonStage["state"] = "normal"
        startButton["state"] = "normal"

    threading.Thread(target=target).start()


def start(excelPathSST: str, excelPathStage: str):
    '''Starts the main script'''
    try:
        excelSST = pd.read_excel(excelPathSST)
    except FileNotFoundError:
        setMessage("Fichier SST invalide")
        return

    try:
        excelStage = pd.read_excel(excelPathStage)
    except FileNotFoundError:
        setMessage("Fichier Stage invalide")
        return

    json = {}
    sectors = getSectors()
    for sector in sectors:
        jobs = {}
        for jobID in getJobIDsOfSector(sector["urlId"], sector["id"]):
            job = getJob(jobID)
            job["questions"] = getStageDataFromExcel(
                excelStage, sector["id"], job["id"])

            for skillIds in job["skills"]:
                job["skills"][skillIds]["risks"] = getSSTDataFromExcel(
                    excelSST, sector["id"], job["id"], skillIds)

            jobs[job["id"]] = job

        json[sector["id"]] = {
            "name": sector["name"],
            "id": sector["id"],
            "jobs": jobs,
        }

    saveJson(json, JSON_FILE_PATH)
    setMessage("Tout est fini!")


# Excel readers
def getSSTDataFromExcel(excel: pd.DataFrame, sector, job, skill):
    '''Returns the corresponding data contained in the SST excel file'''
    result = []
    row = excel.loc[(excel[EXCEL_SECTOR_HEADER] == int(sector)) & (
        excel[EXCEL_JOB_HEADER] == int(job)) & (excel[EXCEL_SKILL_HEADER] == int(skill))]

    for name, excelHeader in EXCEL_SST_DATA_HEADERS.items():
        if (row[excelHeader].index.size > 0 and row[excelHeader].get(row[excelHeader].index[0], "") == "oui"):
            result.append(name)

    return result


def getStageDataFromExcel(excel: pd.DataFrame, sector, job):
    '''Returns the corresponding data contained in the stage excel file'''
    result = []
    row = excel.loc[(excel[EXCEL_SECTOR_HEADER] == int(sector))
                    & (excel[EXCEL_JOB_HEADER] == int(job))]

    for name, excelHeader in EXCEL_STAGE_DATA_HEADERS.items():
        if (row[excelHeader].index.size > 0 and row[excelHeader].get(row[excelHeader].index[0], "") == "Oui"):
            result.append(name)

    return result


# Data fetching
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
                "urlId": input["id"],
                "id": input["value"],
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

    [jobId, jobName] = soup.find("h2").getText(";", True).split(";")
    result = {"name": jobName, "id": jobId, "skills": {}}

    for header in soup.find_all("thead"):
        titleSearch = titleRegex.search(header.find("th").text)
        skillId = titleSearch.group(1)
        skillName = titleSearch.group(2)

        lists = header.find_next_sibling("tbody").find_all("ul")

        criteria = []
        for criterion in lists[0].find_all("li"):
            criteria.append(
                cleanUpText(criterion.text))

        tasks = []
        for task in lists[1].find_all("li"):
            tasks.append(cleanUpText(task.text))

        result["skills"][skillId] = {
            "name": skillName, "id": skillId, "criteria": criteria, "tasks": tasks}

    return result


# String processing
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


# Utils
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

    if (file is not None):
        return file.name
    else:
        return ""


# Tkinter initialisation
root = tk.Tk()
root.title("CRCRME - Générer répertoire métiers")
root.geometry("450x200")
root.resizable(False, False)
mainFrame = tk.Frame(root)
mainFrame.pack(padx=20, pady=20)


tk.Label(mainFrame, text="Entrez le chemin d'accès du classeur Excel contenant les informations SST.").pack()

frame = tk.Frame(mainFrame)
frame.pack()

excelPathSST = tk.StringVar()
entrySST = tk.Entry(frame, textvariable=excelPathSST)
entrySST.focus()
entrySST.pack(side="left")

fileButtonSST = tk.Button(frame, text="Parcourir",
                          command=lambda: excelPathSST.set(askExcelPath()))
fileButtonSST.pack(side="right")


tk.Label(mainFrame, text="Entrez le chemin d'accès du classeur Excel contenant").pack()
tk.Label(mainFrame, text="les questions à poser lors du formulaire de création de stage.").pack()

frame = tk.Frame(mainFrame)
frame.pack()

excelPathStage = tk.StringVar()
entryStage = tk.Entry(frame, textvariable=excelPathStage)
entryStage.pack(side="left")

fileButtonStage = tk.Button(frame, text="Parcourir",
                            command=lambda: excelPathStage.set(askExcelPath()))
fileButtonStage.pack(side="right")


startButton = tk.Button(mainFrame, text="Générer", command=run)
startButton.pack(side="bottom")

currentMessage = tk.StringVar()
messageLabel = tk.Label(mainFrame, textvariable=currentMessage)
messageLabel.pack(side="bottom")


if __name__ == "__main__":
    root.mainloop()
