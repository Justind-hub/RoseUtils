import requests
from bs4 import BeautifulSoup

def run():
    URL  = 'https://github.com/Justind-hub/RoseUtils/blob/main/README.md'
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser") # type: ignore

    entire_div = soup.find(id="readme")
    myver = ""
    with open('README.md', 'r') as f:
        for line in f:
            myver = myver + line

    for line in entire_div: # type: ignore
        if "Version" in str(line):
            for line2 in line:
                if "Version" in str(line2):
                    curver = str(line2)
    curver = curver[14:-4] # type: ignore
    curver = "# RoseUtils\n\n"+curver+"\n"
    return myver == curver