import requests
from bs4 import BeautifulSoup

def run():
    URL  = 'https://raw.githubusercontent.com/Justind-hub/RoseUtils/main/README.md'
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser") # type: ignore

    
    myver = ""
    with open('README.md', 'r') as f:
        for line in f:
            myver = myver + line
    curver = soup.prettify()[:255]
    #for line in entire_div: # type: ignore
    #    if "Version" in str(line):
    #        for line2 in line: # type: ignore
    #            if "Version" in str(line2):
    #                curver = str(line2)
    #curver = curver[14:-4] # type: ignore
    #curver = "# RoseUtils\n\n"+curver+"\n"
    return myver == curver


if __name__ == "__main__":
    run()
    
    '''
    URL  = 'https://raw.githubusercontent.com/Justind-hub/RoseUtils/main/README.md'
    page = requests.get(URL)
    from bs4 import BeautifulSoup

    # Fetch the html file
    response = requests.get(URL)
    

    # Parse the html file
    soup = BeautifulSoup(response.content, 'html.parser')

    # Format the parsed html file
    #strhtm = soup.prettify()
    x = soup.prettify()[:225]
    # Print the first few characters
    print (x)'''