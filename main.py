from odf.opendocument import load
from odf import text, teletype
from os import system


# TODO: parse args
# TODO: frontend

# Input:
## stopień (strz)
## imie NAZWISKO
## imie ojca (czyj syn?)
## nowy stopień (SEKCYJNY)
## korpus osobowy (instruktorów młodszych ZS)
## nr rozkazu (4/2021)
## data (dd.mm.yyyy)

Strzelcy = [
    [
        "sekc",
        "Jan KOWALSKI",
        "s",
        "Jana",
        "DRUŻYNOWY",
        "instruktorów młodszych ZS",
        "5/70",
        "03.05.2070",
    ],
]

def replace_elem(odt_elem: text.Span) -> None:
    # skopiowane z https://github.com/eea/odfpy/wiki/ReplaceOneTextToAnother#replace-one-text-to-another
    new_item = text.Span()
    new_item.setAttribute("stylename", odt_elem.getAttribute("stylename"))
    new_item.addText(elem)
    odt_elem.parentNode.insertBefore(new_item, odt_elem)
    odt_elem.parentNode.removeChild(odt_elem)


for i in range(0, len(Strzelcy)):
    doc = load("akt_mianowania.odt")
    for odt_elem in doc.getElementsByType(text.Span):
        elem = teletype.extractText(odt_elem)
        if "$rank" in elem:
            elem = elem.replace("$rank", Strzelcy[i][0])
            replace_elem(odt_elem)
        if "$name" in elem:
            elem = elem.replace("$name", Strzelcy[i][1])
            replace_elem(odt_elem)
        if "$gender" in elem:
            elem = elem.replace("$gender", Strzelcy[i][2])
            replace_elem(odt_elem)
        if "$father" in elem:
            elem = elem.replace("$father", Strzelcy[i][3])
            replace_elem(odt_elem)
        if "newrank" in elem:
            elem = elem.replace("newrank", Strzelcy[i][4])
            replace_elem(odt_elem)
        if "$type" in elem:
            elem = elem.replace("$type", Strzelcy[i][5])
            replace_elem(odt_elem)
        if "$order" in elem:
            elem = elem.replace("$order", Strzelcy[i][6])
            replace_elem(odt_elem)
        if "$date1" in elem:
            elem = elem.replace("$date1", Strzelcy[i][7])
            replace_elem(odt_elem)
        if "$date2" in elem:
            elem = elem.replace("$date2", Strzelcy[i][7])
            replace_elem(odt_elem)
    filename = "akt_mianowania_" + Strzelcy[i][1] + ".odt"
    doc.save(filename)
    print("Wygenerowano plik: " + filename)
    print("libreoffice --headless --convert-to pdf " + filename)
    system("libreoffice --headless --convert-to pdf '" + filename + "'")
