import re
from tika import parser

def extract_raw(path):
    "Extract the raw text and the year of the last edit from a file using tika-python"
    parsedfile = parser.from_file(path)

    if parsedfile["status"] != 200: raise SystemExit(
        f"Could not parse {path}: tika returned status {parsedfile['status']}")

    return parsedfile["content"], int(parsedfile["metadata"]["date"][:4])


def extract_organists(raw):
    "Retrieve the organists and datetime from raw and return them as a dictionary"
    
    # only last names are stored, but complete names should be returned 
    organistnames = {} # redacted for privacy, add names here before running

    # use regex to extract date, time and organists from the parsed xlsx.
    # morning and afternoon/evening info is extracted separately
    organists_raw = re.findall(r"(\d{2}\/\d{2}\/\d{2})\t(\d{2}:\d{2})\t(\w+)", raw, re.U)
    organists_raw += re.findall(r"(\d{2}\/\d{2}\/\d{2})(?:\t\d{2}:\d{2}\t\w*\t|\t{3})"
        + r"(\d{2}:\d{2})\t(\w+)", raw, re.U)

    # construct the datetime (ISO format) and store the organist name as its value
    organists = {}
    for organist_raw in organists_raw:
        date, time, name = organist_raw
        date_elements = date.split("/")
        datetime = f"20{'-'.join(date_elements[::-1])}T{time}:00"
        organists[datetime] = organistnames[name]

    return organists


def extract_services(raw, year, organists):
    """
    Retrieve the service information from raw, combine it with the
    organists dictionary and store it as a list of dictionaries
    
    Arguments:
        raw: str, the extracted text from the services pdf
        year: int, the year the first service is in
        organists: dict, return value of extract_organists

    Returns:
        list of dictionaries where each dictionary contains service information
    """

    monthnames = {"JANUARI": "01", "FEBRUARI": "02", "MAART": "03", "APRIL": "04",
    "MEI": "05", "JUNI": "06", "JULI": "07", "AUGUSTUS": "08",
    "SEPTEMBER": "09", "OKTOBER": "10", "NOVEMBER": "11", "DECEMBER": "12"}
    lastmonth = False
    
    # get the relevant section
    next_section = re.findall(r"INHOUD.+?Erediensten.+?\n(\w+?)\b", raw, re.S|re.U)[0]
    raw = raw[raw.find("EREDIENSTEN"):raw.find(next_section.upper())]
    # non-breaking spaces -> spaces, unicode 2010 -> dash
    raw = raw.replace("\xa0", " ").replace("\u2010", "-")

    # get the raw service info with regex
    services_raw = re.findall(r"(\d{1,2} \w+).+?\n(.+?)\n ??\n", raw, re.S|re.U)
    
    services = []
    for service_raw in services_raw:
        # get the day and monthname from the first capture group, then extract
        # collections and other service info from the second one with more regex
        day, monthname = service_raw[0].split()
        collections = re.findall(r"collecte\*?: ?(.+?) ?\n", service_raw[1] + "\n")
        service_items = re.findall(r"(\d{2}.\d{2}) uur\s+(.+?)\n", service_raw[1])

        # if the month is the last month and later becomes the first month, add 1 to year
        month = monthnames[monthname]
        if month == "12": lastmonth = True
        elif month == "01" and lastmonth:
            year += 1
            lastmonth = False

        for service_item in service_items:
            service = {}
            time, voorgangerinfo = service_item

            # construct the datetime (ISO format)
            datetime = f"{year}-{month}-{day.rjust(2, '0')}T{time.replace('.', ':')}:00"
            service["Datetime"] = datetime

            # construct the voorganger dictionary
            voorgangerinfo = voorgangerinfo.split(",")
            if len(voorgangerinfo) > 1:
                name = voorgangerinfo[0].strip()
                place = voorgangerinfo[1].strip()
                if place == "Heilig Avondmaal": place = "plaats"
                service["Voorganger"] = {"Naam": name, "Plaats": place}
            elif len(voorgangerinfo) == 1:
                name = voorgangerinfo[0].strip()
                service["Voorganger"] = {"Naam": name, "Plaats": "plaats"}

            # construct the collections dictionary
            service["Collecten"] = { "Collecte 1": collections[0],
                "Collecte 3": collections[1]}

            # add the organist to the dictionary
            try:
                service["Organist"] = organists[datetime]
            except KeyError:
                service["Organist"] = "organist"

            services.append(service)

    return services
