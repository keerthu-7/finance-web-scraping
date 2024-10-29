# Import necessary modules
from csv import reader  # To read CSV files
from os import mkdir  # To create directories
from edgarpython.exceptions import InvalidCIK  # Custom exception for invalid CIKs
from edgarpython.secapi import getSubmissionsByCik, getXlsxUrl  # To fetch SEC submissions and XLSX URLs
from requests import get  # To make HTTP requests for downloading files
from rich.progress import track  # To display a progress bar during iteration

# Function to download a file from a given URL and save it to a specified filename
def download(url, filename):
    resp = get(
        url,
        headers={
            # Standard user-agent string for requests to mimic a browser request
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:130.0) Gecko/20100101 Firefox/130.0"
        },
    )
    # Write the content to the file in binary mode
    with open(filename, "wb") as file:
        file.write(resp.content)

# Open the CSV file containing the S&P 500 companies' data
with open("sp500.csv", encoding="utf-8") as file:
    csv = reader(file)  # Read the CSV content
    companies = list(csv)[1:]  # Skip the header row and convert the rest to a list of companies

# Create the 'Output' directory to store the downloaded reports
mkdir("Output")

# Loop over each company in the list with a progress bar
for company in track(companies):
    # Create a directory for each company based on its name
    mkdir(f"Output/{company[1]}")

    try:
        # Get all the SEC submissions for the company using its CIK number
        submissions = getSubmissionsByCik(company[6])
        selected = []  # List to store relevant submissions (10-K forms)

        # Filter the submissions to include only the ones where form type is '10-K'
        for submission in submissions:
            if submission.form == "10-K":
                selected.append(submission)

        # Print the number of 10-K forms found for the company
        print(f"Found {len(selected)} 10-K for {company[1]}")

        downloads = []  # List to store download URLs for 10-K reports
        missed = 0  # Counter for missed downloads

        # Loop through the selected 10-K submissions to get their download URLs
        for submission in selected:
            try:
                # Get the XLSX URL for each submission
                downloads.append(getXlsxUrl(company[6], submission.accessionNumber))
            except FileNotFoundError:
                # If the file is not found, increase the missed count and skip to the next
                missed += 1
                continue

        # Print the total number of reports to be downloaded and how many were missed
        print(
            f"{len(downloads)} reports to be downloaded for {company[6]} [missed {missed}]"
        )

        total = len(downloads)  # Total number of reports to download
        done = 0  # Counter for successfully downloaded reports

        # Download each report from the URLs collected
        for downloadUrl in downloads:
            # Download the report and save it with the accession number as the filename
            download(
                downloadUrl,
                f"Output/{company[1]}/{downloadUrl.split('/')[-2]}.xlsx",
            )
            done += 1
            # Print the progress of the downloads for the company
            print(f"Downloaded [{done}/{total}]")

    # Handle invalid CIK errors by printing a message and moving to the next company
    except InvalidCIK:
        print("Failed for " + company[1])
        continue
