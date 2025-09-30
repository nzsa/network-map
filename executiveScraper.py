from bs4 import BeautifulSoup
import requests
import fyahooImporter as fi
import time
import pandas as pd
import random
from pyvis.network import Network
import webbrowser
import colorsys
from flask import Flask, send_file

import os
from pathlib import Path

DIRECTOR_DATA_PATH = r'C:\Users\Oliver\NZSA\CS - Companies\Directors\NZX Directors.csv'

NAME_MAP = {
    # A
    "Al": "Albert",
    "Alex": "Alexander",
    "Allie": "Alice",
    "Andy": "Andrew",
    "Archie": "Archibald",
    "Art": "Arthur",
    
    # B
    "Barb": "Barbara",
    "Bess": "Elizabeth",
    "Betsy": "Elizabeth",
    "Betty": "Elizabeth",
    "Ben": "Benjamin",
    "Benny": "Benjamin",
    "Bill": "William",
    "Billy": "William",
    "Bob": "Robert",
    "Bobby": "Robert",
    "Brad": "Bradley",
    "Brent": "Brenton",
    "Bryce": "Brycen",
    
    # C
    "Cal": "Calvin",
    "Cam": "Cameron",
    "Carrie": "Caroline",
    "Charlie": "Charles",
    "Chuck": "Charles",
    "Chris": "Christopher",
    "Cindy": "Cynthia",
    "Cliff": "Clifford",
    "Connie": "Constance",
    "Court": "Courtney",
    
    # D
    "Dan": "Daniel",
    "Danny": "Daniel",
    "Dave": "David",
    "Deb": "Deborah",
    "Debbie": "Deborah",
    "Del": "Delbert",
    "Drew": "Andrew",
    "Don": "Donald",
    "Donna": "Donatella",
    "Doug": "Douglas",
    "Dot": "Dorothy",
    "Dottie": "Dorothy",
    
    # E
    "Ed": "Edward",
    "Eddie": "Edward",
    "Ellie": "Eleanor",
    "Elly": "Elizabeth",
    "Elsie": "Elizabeth",
    "Ernie": "Ernest",
    
    # F
    "Frank": "Francis",
    "Frankie": "Francis",
    "Fran": "Frances",
    "Fred": "Frederick",
    "Freddy": "Frederick",
    
    # G
    "Gabe": "Gabriel",
    "Gary": "Garrett",
    "Gerry": "Gerald",
    "Greg": "Gregory",
    "Gus": "Augustus",
    
    # H
    "Hal": "Harold",
    "Hank": "Henry",
    "Harry": "Harold",
    "Hattie": "Harriet",
    "Hettie": "Henrietta",
    "Hazel": "Hazelene",
    "Izzy": "Isabella",
    
    # J
    "Jack": "Jackson",
    "Jake": "Jacob",
    "Jamie": "James",
    "Jan": "Janet",
    "Jenny": "Jennifer",
    "Jess": "Jessica",
    "Jim": "James",
    "Jimmy": "James",
    "Jo": "Joanna",
    "Joe": "Joseph",
    "Joey": "Joseph",
    "Johnny": "John",
    "Josh": "Joshua",
    "Jules": "Julian",
    
    # K
    "Kate": "Katherine",
    "Katie": "Katherine",
    "Kathy": "Katherine",
    "Ken": "Kenneth",
    "Kenny": "Kenneth",
    "Kim": "Kimberly",
    "Kris": "Kristopher",
    
    # L
    "Larry": "Lawrence",
    "Laurie": "Lawrence",
    "Leo": "Leonard",
    "Liz": "Elizabeth",
    "Lizzy": "Elizabeth",
    "Liza": "Elizabeth",
    "Lois": "Louisa",
    "Lou": "Louis",
    "Lori": "Loretta",
    "Lynn": "Lynette",
    
    # M
    "Maddie": "Madeline",
    "Maggie": "Margaret",
    "Maisie": "Margaret",
    "Mandy": "Amanda",
    "Matt": "Matthew",
    "Matty": "Matthew",
    "Meg": "Margaret",
    "Megan": "Margaret",
    "Mike": "Michael",
    "Mikey": "Michael",
    "Mitch": "Mitchell",
    "Molly": "Mary",
    "Monty": "Montgomery",
    
    # N
    "Nate": "Nathan",
    "Ned": "Edward",
    "Nick": "Nicholas",
    "Nicky": "Nicholas",
    "Nell": "Eleanor",
    "Nora": "Eleanor",
    
    # O
    "Oli": "Oliver",
    "Ollie": "Oliver",
    "Orie": "Orville",
    "Ozzie": "Oswald",
    
    # P
    "Pat": "Patrick",
    "Paddy": "Patrick",
    "Patty": "Patricia",
    "Peg": "Margaret",
    "Peggy": "Margaret",
    "Phil": "Philip",
    "Philippa": "Philippa",
    "Pip": "Philippa",
    "Polly": "Mary",
    "Prue": "Prudence",
    
    # R
    "Ray": "Raymond",
    "Rich": "Richard",
    "Richie": "Richard",
    "Rick": "Richard",
    "Ricky": "Richard",
    "Rob": "Robert",
    "Robbie": "Robert",
    "Ron": "Ronald",
    "Ronnie": "Ronald",
    "Rosie": "Rose",
    "Ruthie": "Ruth",
    
    # S
    "Sam": "Samuel",
    "Sammy": "Samuel",
    "Sandy": "Sandra",
    "Sasha": "Alexander",
    "Scottie": "Scott",
    "Shelly": "Michelle",
    "Steph": "Stephanie",
    "Steve": "Steven",
    "Stevie": "Steven",
    "Sue": "Susan",
    "Susie": "Susan",
    
    # T
    "Ted": "Theodore",
    "Teddy": "Theodore",
    "Terri": "Teresa",
    "Tess": "Theresa",
    "Tim": "Timothy",
    "Tom": "Thomas",
    "Tommy": "Thomas",
    "Tony": "Anthony",
    "Trish": "Patricia",
    "Trudy": "Gertrude",
    
    # V
    "Val": "Valerie",
    "Vic": "Victor",
    "Vicky": "Victoria",
    
    # W
    "Walt": "Walter",
    "Will": "William",
    "Willie": "William",
    "Winnie": "Winifred",
    
    # Z
    "Zach": "Zachary",
    "Zack": "Zachary",
    "Zeke": "Ezekiel"
}

css_style = """body, html {
            margin: 0;
            padding: 0;
            height: 100%;
          }
          .info-text {
            position: absolute;
            top: 10px;
            left: 10px;
            z-index: 10;
            padding: 8px 12px;
            color: white;
            font-family: Consolas, Menlo;
          }
          .top-right-button {
            position: absolute;
            top: 10px;
            right: 20px;
            z-index: 10;
            padding: 8px 12px;
            background-color: #217a00;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-family: sans-serif;
          }
          .top-right-button:hover {
            background-color: #217a00;
          }"""

app = Flask(__name__)

def generate_color_map(items):
    """Return random colours for each node"""
    colourDict = {}
    for item in items:
        # Normalize HSV values to 0.0–1.0 range

        h_norm = random.randint(0, 360) / 360.0
        s_norm = random.randint(50, 100) / 100.0
        v_norm = random.randint(50, 100) / 100.0

        # Convert to RGB (returns floats in 0.0–1.0)
        r, g, b = colorsys.hsv_to_rgb(h_norm, s_norm, v_norm)

        # Convert to 0–255 and format as hex
        colourDict[item] = '#{:02x}{:02x}{:02x}'.format(int(r * 255), int(g * 255), int(b * 255))

    return colourDict

def createNetwork(df, csv_path):
    net = Network(height="750px", width="100%", bgcolor="#222222", font_color="white", notebook=False)

    # Add nodes with labels for each person
    for person in df.index:
        net.add_node(person, label=person, group='people')

    #Add invisible nodes with labels for each company
    for company in df.columns:
        net.add_node(company, label=company, group='company', color='rgba(0,0,0,0)')

    companyColours = generate_color_map(df.columns.values)

    #Create a line between people
    for company in df.columns:
        members = df.loc[df[company]].index.tolist()
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                net.add_edge(
                    members[i],
                    members[j],
                    title=company,
                    color=companyColours.get(company, '#CCCCCC'),
                    length=100
                )
            #Connect the company label to each company
            net.add_edge(
                    company,
                    members[i],
                    title=company,
                    color='rgba(0,0,0,0)',
                    length=5
                )

    #Make the company labels large and turn on physics
    net.set_options("""
    {
        "groups": {
            "company": {
                "font": {
                    "size": 52,
                    "bold": true,
                    "color": "white"
                },
                "color": "rgba(0,0,0,0)"
            }
        },
        "physics": {
            "solver": "forceAtlas2Based",
            "forceAtlas2Based": {
                "gravitationalConstant": -50,
                "centralGravity": 0.01,
                "springLength": 20,
                "springConstant": 0.08,
                "damping": 0.4,
                "avoidOverlap": 0
            },
            "stabilization": {
                "iterations": 200
            }
        }
    }""")
    # Save next to CSV
    out_dir = Path(csv_path).parent
    html_path = out_dir / "network_map.html"
    net.write_html(str(html_path))
    return str(html_path)  # absolute path (because out_dir is absolute from DIRECTOR_DATA_PATH)

def scrapeNZX(tickerStrs):
    directorList = []
    companyNum = 1
    for ticker in tickerStrs:
        print(f"\n{ticker}: {companyNum}/{len(tickerStrs)}")
        companyNum += 1
        url = "https://www.nzx.com/companies/" + ticker[:3]
        time.sleep(random.uniform(1, 3))
        try:
            response = requests.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Example: Find elements containing director names (this will vary greatly by website)
            director_names = soup.find_all('div', class_='Grid jfsytL grid-lg-2-3')
            for name_element in director_names:
                #Need to place a known text in between the name and title, to split on
                string = str(name_element)
                string = string.replace("</strong><span>", "</strong>$<span>")
                name_element = BeautifulSoup(string, 'html.parser')
                string = name_element.text
                #Only grab Directors/Chairs
                if 'Director' in string or 'Chair' in string:
                    #Get the name and position on either side of the 'known string'
                    name, position = string.split('$')
                    #Only take the first and last names and lengthen any short names
                    nameList = name.split(' ')
                    if nameList[0] in NAME_MAP:
                        nameList[0] = NAME_MAP[nameList[0]]

                    name = nameList[0] + " " + nameList[-1]
                    print(f"{name}: {position}")
                    directorList.append(pd.DataFrame({'Company': [ticker], 'Name': [name], 'Title': [position]}))

        except requests.exceptions.RequestException as e:
            print(f"Error during web scraping: {e}")
        except Exception as e:
            print(f"Error parsing HTML: {e}")
    
    return directorList

def getTickers(exchange):
    #PNK = OTC1, OQX=OTC2 (Established), OQB=OTC3 (Venture), NZE = NZX, ASX = ASX
    dfStocks = fi.getAllStocks(exchange)
    dfStocks['Length'] = [len(tick) for tick in dfStocks['symbol']]
    dfStocks = dfStocks.loc[dfStocks['Length']==6] # Use only for NZ market, maybe ASX? otherwise it gets all the bonds
    tickerStrs = list(dfStocks['symbol'])
    return tickerStrs

def countConnections(networkDf):
    connectionDict = {}
    for name in networkDf.index:
        relatedCompaniesDf = networkDf.loc[:, networkDf.loc[name] == True]
        connectionDict[name] = relatedCompaniesDf.any(axis=1).sum()-1
    top5ConnectionDict = dict(sorted(connectionDict.items(), key=lambda item: item[1], reverse=True)[:5])
    return top5ConnectionDict

def countIsolatedCompanies(df):
    isolatedCount = 0
    for company in df.columns:
        companyDirectorsDf = df.loc[df[company]==True]
        companyDirectorsDf = companyDirectorsDf.drop(company, axis=1)
        if companyDirectorsDf.values.any() == False:
            isolatedCount += 1 

    return isolatedCount



def insertHTML(fileName, newContent, priorTag):
    with open(fileName, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    tag = soup.head.find(priorTag)

    tag.string = tag.string + newContent

    # Save the modified HTML
    with open(fileName, "w", encoding="utf-8") as f:
        f.write(str(soup))


def removeHTMLTags(fileName, removeTags, classes=None):
    with open(fileName, 'r') as file:
        soup = BeautifulSoup(file, 'html.parser')

    if classes==None:
        for removeTag in removeTags:
            for tag in soup.find_all(removeTag):
                tag.decompose()
    else:
        for removeTag in removeTags:
            for removeClass in classes:
                for tag in soup.find_all(removeTag, class_=removeClass):
                    tag.decompose()

    # Save the cleaned HTML
    with open(fileName, 'w') as file:
        file.write(str(soup))

def insertHTMLTags(fileName, priorTag, afterClose, newTag, newContent=None, priorClass=None, newClass=None, newId=None, newHref=None):
    # Load your HTML
    with open(fileName, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    # Find the target div by class
    target = soup.find(priorTag, class_=priorClass) if priorClass else None
    if not target:
        target = soup.find(priorTag)
    
    newTagObj = soup.new_tag(newTag)
    # Create the new tag
    if newClass != None:
        newTagObj['class'] = newClass
    if newId != None:
        newTagObj['id'] = newId
    if newContent != None:
        newTagObj.string = newContent
    if newHref != None:
        newTagObj['href'] = newHref

    # Insert it after the target div
    if target and afterClose == 0:
        target.insert(0, newTagObj)
    elif target:
        target.insert_after(newTagObj)
    else:
        print("Target div not found.")

    # Save the updated HTML
    with open(fileName, "w", encoding="utf-8") as f:
        f.write(str(soup))

def dfToHTMLText(df, column1, column2):
    
    max_len = max(len(director) for director in df.index)
    busyDirectorStrList = [f"{column1.ljust(max_len+1)} {column2}\n"]
    for director, companyCount in df.items():
        director = director+":"
        busyDirectorStrList.append(f"{director.ljust(max_len+1)} {companyCount}\n")
    return "".join(busyDirectorStrList)

@app.route("/"+DIRECTOR_DATA_PATH)
def download_csv():
    csv_path = DIRECTOR_DATA_PATH

    # Send file to browser
    return send_file(csv_path, mimetype="text/csv", as_attachment=True, download_name="NZX Directors.csv")

def main():
    tickerStrs = getTickers('NZE')
    directorList = scrapeNZX(tickerStrs)
    directorDf = pd.concat(directorList)
    
    numPositions = directorDf.shape[0]
    busiestDirectorsDf = directorDf.groupby('Name').count().sort_values('Company', ascending=False).head(5)['Company']

    directorDf = directorDf.set_index('Name')
    # Create a boolean matrix: rows are names, columns are tickers
    directorNetwork = directorDf.assign(value=True).pivot_table(
        index='Name',
        columns='Company',
        values='value',
        aggfunc='any',
        fill_value=False
    )
    numUniqueDirectors = directorNetwork.shape[0]

    print(busiestDirectorsDf)
    busiestDirectorsHTML = dfToHTMLText(busiestDirectorsDf, "Name", "# Companies")
    mostConnectedDict = countConnections(directorNetwork)
    mostConnectedHTML = dfToHTMLText(pd.Series(mostConnectedDict), "Name", "# Connections")
    isolatedCompanies = countIsolatedCompanies(directorNetwork)

    allInfoText = '# Positions: ' + str(numPositions) + '\n\n' + '# Unique Directors: ' + str(numUniqueDirectors) + '\n\n' + '# Isolated Companies: ' + str(isolatedCompanies) + '\n\n' +busiestDirectorsHTML + '\n\n' + mostConnectedHTML

    #directorNetwork.to_csv('C:\\Users\\Oliver\\NZSA\\CS - Companies\\Directors\\NZX Director Network.csv')
    directorDf.to_csv(DIRECTOR_DATA_PATH)

    fileNameHTML = createNetwork(directorNetwork, DIRECTOR_DATA_PATH)
    removeHTMLTags(fileNameHTML, ['center', 'h1'])
    insertHTML(fileNameHTML, css_style, "style")
    removeHTMLTags(fileNameHTML, ['div'], ['card', 'card-body'])
    insertHTMLTags(fileNameHTML, "body", 0,  "pre", allInfoText, newClass="info-text")
    insertHTMLTags(fileNameHTML, "pre", 1, "a", newHref='/'+DIRECTOR_DATA_PATH, priorClass="info-text")
    insertHTMLTags(fileNameHTML, "a", 0, "button", "Download Data", newClass="top-right-button")
    insertHTMLTags(fileNameHTML, "a", 1, "div", priorClass="top-right-button", newId='mynetwork')
    # Open automatically in browser and print where it went
    webbrowser.open("file://" + os.path.abspath(fileNameHTML).replace("\\", "/"))

main()