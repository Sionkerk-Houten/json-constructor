import json
import argparse
from extract import extract_raw, extract_organists, extract_services

# parse commandline arguments
argparser = argparse.ArgumentParser(
    description = "Create a services.json file which can be used by PPTXcreator")
argparser.add_argument("servicespath", help = ".pdf file with service info (kerkbode)")
argparser.add_argument("organistspath", help = ".xlsx file with the organists schedule")
argparser.add_argument("outputpath", help = "Where the output .json should be saved")
args = argparser.parse_args()

# construct the services file
xlsxtext, _ = extract_raw(args.organistspath)
organists = extract_organists(xlsxtext)
pdftext, year = extract_raw(args.servicespath)
services = extract_services(pdftext, year, organists)

# write to outputpath
if args.outputpath[-5:].lower() != ".json":
    args.outputpath += ".json"

with open(args.outputpath, "w") as file:
    json.dump(services, file, indent = 2)
