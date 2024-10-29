from typing import List

from requests import get

from edgarpython.exceptions import InvalidCIK
from edgarpython.models import Submission

API_BASE = "https://data.sec.gov"
USERAGENT = "Mozilla/5.0 (X11; Linux x86_64; rv:130.0) Gecko/20100101 Firefox/130.0"
HEADERS = {"User-Agent": USERAGENT}


def getSubmissionsByCik(cik: str) -> List[Submission]:
    reqUrl = API_BASE + f"/submissions/CIK{cik}.json"
    resp = get(reqUrl, timeout=5000, headers=HEADERS)
    match resp.status_code:
        case 404:
            raise InvalidCIK
        case 200:
            pass
        case _:
            raise RuntimeError(
                f"SECAPI call failed with status code {resp.status_code}"
            )
    data = resp.json()
    data = data["filings"]["recent"]
    submissions = []
    for subZiped in zip(data["form"], data["accessionNumber"]):
        submissions.append(Submission(form=subZiped[0], accessionNumber=subZiped[1]))
    return submissions


def getXlsxUrl(cik: str, accessionNumber: str) -> str:
    accessionNumber = accessionNumber.replace("-", "")
    cik = cik.lstrip("0")
    reqUrl = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accessionNumber}/Financial_Report.xlsx"
    resp = get(reqUrl, timeout=2000, headers=HEADERS)
    match resp.status_code:
        case 404:
            raise FileNotFoundError
        case 200:
            pass
        case _:
            raise RuntimeError(
                f"SECAPI call failed with status code {resp.status_code}"
            )
    return reqUrl
