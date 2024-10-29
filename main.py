from csv import reader
from os import mkdir

from edgarpython.exceptions import InvalidCIK
from edgarpython.secapi import getSubmissionsByCik, getXlsxUrl
from requests import get
from rich.progress import track

def download(url, filename):
    resp = get(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:130.0) Gecko/20100101 Firefox/130.0"
        },
    )
    with open(filename, "wb") as file:
        file.write(resp.content)

with open("sp500.csv", encoding="utf-8") as file:
    csv = reader(file)
    companies = list(csv)[1:]

mkdir("Output")
for company in track(companies):
    mkdir(f"Output/{company[1]}")
    try:
        submissions = getSubmissionsByCik(company[6])
        selected = []
        for submission in submissions:
            if submission.form == "10-K":
                selected.append(submission)
        print(f"Found {len(selected)} 10-K for {company[1]}")
        downloads = []
        missed = 0
        for submission in selected:
            try:
                downloads.append(getXlsxUrl(company[6], submission.accessionNumber))
            except FileNotFoundError:
                missed += 1
                continue
        print(
            f"{len(downloads)} reports to be downloaded for {company[6]} [missed {missed}]"
        )
        total = len(downloads)
        done = 0
        for downloadUrl in downloads:
            download(
                downloadUrl,
                f"Output/{company[1]}/{downloadUrl.split('/')[-2]}.xlsx",
            )
            done += 1
            print(f"Downloaded [{done}/{total}]")

    except InvalidCIK:
        print("Failed for " + company[1])
        continue
