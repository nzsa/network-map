# executivescraper.py
# Outputs:
#   - network_map.html
#   - NZX_Directors.csv
#
# These two files are written to the SAME folder as this script (repo root)
# so a deploy job can upload them to your website (e.g. /wp-content/uploads/network-map/).
# The HTML includes a Download button linking to "NZX_Directors.csv" (relative URL).

from bs4 import BeautifulSoup
import requests
import fyahooImporter as fi
import time
import pandas as pd
import random
from pyvis.network import Network
import colorsys
import os
from pathlib import Path

# ---------- CONFIG ----------
# Output directory = folder containing this script (repo root)
OUTPUT_DIR = Path(__file__).parent.resolve()
CSV_PATH = OUTPUT_DIR / "NZX_Directors.csv"          # final CSV name (stable)
HTML_PATH = OUTPUT_DIR / "network_map.html"          # final HTML name (stable)

# If you keep this True locally, the script opens your browser.
# On GitHub Actions, this is automatically suppressed.
OPEN_BROWSER_LOCALLY = True

NAME_MAP = {
    # A
    "Al": "Albert", "Alex": "Alexander", "Allie": "Alice", "Andy": "Andrew",
    "Archie": "Archibald", "Art": "Arthur",
    # B
    "Barb": "Barbara", "Bess": "Elizabeth", "Betsy": "Elizabeth", "Betty": "Elizabeth",
    "Ben": "Benjamin", "Benny": "Benjamin", "Bill": "William", "Billy": "William",
    "Bob": "Robert", "Bobby": "Robert", "Brad": "Bradley", "Brent": "Brenton",
    "Bryce": "Brycen",
    # C
    "Cal": "Calvin", "Cam": "Cameron", "Carrie": "Caroline", "Charlie": "Charles",
    "Chuck": "Charles", "Chris": "Christopher", "Cindy": "Cynthia", "Cliff": "Clifford",
    "Connie": "Constance", "Court": "Courtney",
    # D
    "Dan": "Daniel", "Danny": "Daniel", "Dave": "David", "Deb": "Deborah",
    "Debbie": "Deborah", "Del": "Delbert", "Drew": "Andrew", "Don": "Donald",
    "Donna": "Donatella", "Doug": "Douglas", "Dot": "Dorothy", "Dottie": "Dorothy",
    # E
    "Ed": "Edward", "Eddie": "Edward", "Ellie": "Eleanor", "Elly": "Elizabeth",
    "Elsie": "Elizabeth", "Ernie": "Ernest",
    # F
    "Frank": "Francis", "Frankie": "Francis", "Fran": "Frances",
    "Fred": "Frederick", "Freddy": "Frederick",
    # G
    "Gabe": "Gabriel", "Gary": "Garrett", "Gerry": "Gerald", "Greg": "Gregory",
    "Gus": "Augustus",
    # H
    "Hal": "Harold", "Hank": "Henry", "Harry": "Harold", "Hattie": "Harriet",
    "Hettie": "Henrietta", "Hazel": "Hazelene", "Izzy": "Isabella",
    # J
    "Jack": "Jackson", "Jake": "Jacob", "Jamie": "James", "Jan": "Janet",
    "Jenny": "Jennifer", "Jess": "Jessica", "Jim": "James", "Jimmy": "James",
    "Jo": "Joanna", "Joe": "Joseph", "Joey": "Joseph", "Johnny": "John",
    "Josh": "Joshua", "Jules": "Julian",
    # K
    "Kate": "Katherine", "Katie": "Katherine", "Kathy": "Katherine",
    "Ken": "Kenneth", "Kenny": "Kenneth", "Kim": "Kimberly", "Kris": "Kristopher",
    # L
    "Larry": "Lawrence", "Laurie": "Lawrence", "Leo": "Leonard", "Liz": "Elizabeth",
    "Lizzy": "Elizabeth", "Liza": "Elizabeth", "Lois": "Louisa", "Lou": "Louis",
    "Lori": "Loretta", "Lynn": "Lynette",
    # M
    "Maddie": "Madeline", "Maggie": "Margaret", "Maisie": "Margaret",
    "Mandy": "Amanda", "Matt": "Matthew", "Matty": "Matthew", "Meg": "Margaret",
    "Megan": "Margaret", "Mike": "Michael", "Mikey": "Michael", "Mitch": "Mitchell",
    "Molly": "Mary", "Monty": "Montgomery",
    # N
    "Nate": "Nathan", "Ned": "Edward", "Nick": "Nicholas", "Nicky": "Nicholas",
    "Nell": "Eleanor", "Nora": "Eleanor",
    # O
    "Oli": "Oliver", "Ollie": "Oliver", "Orie": "Orville", "Ozzie": "Oswald",
    # P
    "Pat": "Patrick", "Paddy": "Patrick", "Patty": "Patricia", "Peg": "Margaret",
    "Peggy": "Margaret", "Phil": "Philip", "Philippa": "Philippa", "Pip": "Philippa",
    "Polly": "Mary", "Prue": "Prudence",
    # R
    "Ray": "Raymond", "Rich": "Richard", "Richie": "Richard", "Rick": "Richard",
    "Ricky": "Richard", "Rob": "Robert", "Robbie": "Robert", "Ron": "Ronald",
    "Ronnie": "Ronald", "Rosie": "Rose", "Ruthie": "Ruth",
    # S
    "Sam": "Samuel", "Sammy": "Samuel", "Sandy": "Sandra", "Sasha": "Alexander",
    "Scottie": "Scott", "Shelly": "Michelle", "Steph": "Stephanie",
    "Steve": "Steven", "Stevie": "Steven", "Sue": "Susan", "Susie": "Susan",
    # T
    "Ted": "Theodore", "Teddy": "Theodore", "Terri": "Teresa", "Tess": "Theresa",
    "Tim": "Timothy", "Tom": "Thomas", "Tommy": "Thomas", "Tony": "Anthony",
    "Trish": "Patricia", "Trudy": "Gertrude",
    # V
    "Val": "Valerie", "Vic": "Victor", "Vicky": "Victoria",
    # W
    "Walt": "Walter", "Will": "William", "Willie": "William", "Winnie": "Winifred",
    # Z
    "Zach": "Zachary", "Zack": "Zachary", "Zeke": "Ezekiel"
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

# ---------- HELPERS ----------

def generate_color_map(items):
    """Return random colours for each company label / edge group."""
    colourDict = {}
    for item in items:
        h_norm = random.randint(0, 360) / 360.0
        s_norm = random.randint(50, 100) / 100.0
        v_norm = random.randint(50, 100) / 100.0
        r, g, b = colorsys.hsv_to_rgb(h_norm, s_norm, v_norm)
        colourDict[item] = '#{:02x}{:02x}{:02x}'.format(int(r * 255), int(g * 255), int(b * 255))
    return colourDict

def create_network_html(df, html_path: Path):
    """Build the PyVis network and write to html_path."""
    net = Network(height="750px", width="100%", bgcolor="#222222", font_color="white", notebook=False)

    # People nodes
    for person in df.index:
        net.add_node(person, label=person, group='people')

    # Invisible company label nodes
    for company in df.columns:
        net.add_node(company, label=company, group='company', color='rgba(0,0,0,0)')

    companyColours = generate_color_map(df.columns.values)

    # Edges between people who share a company; invisible edges from company label to members
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
            net.add_edge(
                company,
                members[i],
                title=company,
                color='rgba(0,0,0,0)',
                length=5
            )

    net.set_options("""
    {
        "groups": {
            "company": {
                "font": {"size": 52, "bold": true, "color": "white"},
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
            "stabilization": {"iterations": 200}
        }
    }""")

    net.set_template("""
    <!DOCTYPE html>
    <html>
    <head>{{ head }}</head>
    <body>
    <div id="stats-panel"><!--STATS_MARKER--></div>
    {{ body }}
    </body>
    </html>
    """)

    net.write_html(str(html_path))

def scrape_nzx_directors(tickerStrs):
    directorList = []
    companyNum = 1
    headers = {"User-Agent": "Mozilla/5.0"}  # Be a bit polite
    for ticker in tickerStrs:
        print(f"\n{ticker}: {companyNum}/{len(tickerStrs)}")
        companyNum += 1
        url = "https://www.nzx.com/companies/" + ticker[:3]
        time.sleep(random.uniform(1, 3))
        try:
            response = requests.get(url, headers=headers, timeout=20)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            blocks = soup.find_all('div', class_='Grid jfsytL grid-lg-2-3')
            for block in blocks:
                # Insert $ marker between name and title to split reliably
                string = str(block).replace("</strong><span>", "</strong>$<span>")
                parsed = BeautifulSoup(string, 'html.parser')
                text = parsed.text

                if ('Director' in text) or ('Chair' in text):
                    if '$' not in text:
                        continue  # skip unexpected shapes
                    name, position = text.split('$', 1)
                    parts = name.split()
                    if parts:
                        if parts[0] in NAME_MAP:
                            parts[0] = NAME_MAP[parts[0]]
                        # Use first + last
                        name = f"{parts[0]} {parts[-1]}"
                        print(f"{name}: {position.strip()}")
                        directorList.append(pd.DataFrame({'Company': [ticker], 'Name': [name], 'Title': [position.strip()]}))

        except requests.exceptions.RequestException as e:
            print(f"Error during web scraping: {e}")
        except Exception as e:
            print(f"Error parsing HTML: {e}")
    return directorList

def get_tickers(exchange):
    # PNK = OTC1, OQX=OTC2 (Established), OQB=OTC3 (Venture), NZE = NZX, ASX = ASX
    dfStocks = fi.getAllStocks(exchange)
    dfStocks['Length'] = [len(tick) for tick in dfStocks['symbol']]
    # NZX tickers typically length 6; this filters out bonds, etc.
    dfStocks = dfStocks.loc[dfStocks['Length'] == 6]
    return list(dfStocks['symbol'])

def count_connections(networkDf):
    connectionDict = {}
    for name in networkDf.index:
        relatedCompaniesDf = networkDf.loc[:, networkDf.loc[name] == True]
        connectionDict[name] = relatedCompaniesDf.any(axis=1).sum() - 1
    top5 = dict(sorted(connectionDict.items(), key=lambda item: item[1], reverse=True)[:5])
    return top5

def count_isolated_companies(df):
    isolated = 0
    for company in df.columns:
        companyDirectorsDf = df.loc[df[company] == True]
        # If no director shares any other company, treat as isolated
        companyDirectorsDf = companyDirectorsDf.drop(company, axis=1, errors='ignore')
        if companyDirectorsDf.values.any() == False:
            isolated += 1
    return isolated

def insert_css(fileName, new_css):
    with open(fileName, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")
    style_tag = soup.head.find("style")
    if style_tag and style_tag.string:
        style_tag.string = style_tag.string + new_css
    else:
        # create a <style> if missing
        style = soup.new_tag("style")
        style.string = new_css
        soup.head.append(style)
    with open(fileName, "w", encoding="utf-8") as f:
        f.write(str(soup))

def remove_html_tags(fileName, removeTags, classes=None):
    with open(fileName, 'r', encoding="utf-8") as file:
        soup = BeautifulSoup(file, 'html.parser')
    if classes is None:
        for removeTag in removeTags:
            for tag in soup.find_all(removeTag):
                tag.decompose()
    else:
        for removeTag in removeTags:
            for removeClass in classes:
                for tag in soup.find_all(removeTag, class_=removeClass):
                    tag.decompose()
    with open(fileName, 'w', encoding="utf-8") as file:
        file.write(str(soup))

def insert_html_tag(fileName, priorTag, afterClose, newTag, newContent=None,
                    priorClass=None, newClass=None, newId=None, newHref=None):
    with open(fileName, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")
    target = soup.find(priorTag, class_=priorClass) if priorClass else soup.find(priorTag)
    if not target:
        print("Target tag not found.")
        with open(fileName, "w", encoding="utf-8") as f:
            f.write(str(soup))
        return
    newTagObj = soup.new_tag(newTag)
    if newClass is not None:
        newTagObj['class'] = newClass
    if newId is not None:
        newTagObj['id'] = newId
    if newHref is not None:
        newTagObj['href'] = newHref
    if newContent is not None:
        newTagObj.string = newContent
    if afterClose == 0:
        target.insert(0, newTagObj)
    else:
        target.insert_after(newTagObj)
    with open(fileName, "w", encoding="utf-8") as f:
        f.write(str(soup))

def df_to_pretty_text(series_or_df, column1, column2):
    """Formats a Series like Name -> Count into aligned text block for the overlay."""
    if isinstance(series_or_df, pd.Series):
        idx_iter = series_or_df.index
        vals = series_or_df
    else:
        idx_iter = series_or_df.index
        vals = series_or_df.iloc[:, 0]
    max_len = max(len(str(director)) for director in idx_iter) if len(idx_iter) else len(column1)
    lines = [f"{column1.ljust(max_len+1)} {column2}\n"]
    for director, count in vals.items():
        label = f"{director}:"
        lines.append(f"{label.ljust(max_len+1)} {count}\n")
    return "".join(lines)

# ---------- MAIN ----------

def main():
    tickers = get_tickers('NZE')
    directorList = scrape_nzx_directors(tickers)
    if not directorList:
        raise RuntimeError("No director data scraped; cannot build network.")

    directorDf = pd.concat(directorList, ignore_index=True)

    numPositions = directorDf.shape[0]
    busiestDirectors = directorDf.groupby('Name').count().sort_values('Company', ascending=False).head(5)['Company']

    # Build boolean matrix: rows = director names, cols = companies, True if holds position
    directorDf = directorDf.set_index('Name')
    directorNetwork = directorDf.assign(value=True).pivot_table(
        index='Name',
        columns='Company',
        values='value',
        aggfunc='any',
        fill_value=False
    )
    numUniqueDirectors = directorNetwork.shape[0]

    print(busiestDirectors)
    busiestHTML = df_to_pretty_text(busiestDirectors, "Name", "# Companies")
    mostConnected = count_connections(directorNetwork)
    mostConnectedHTML = df_to_pretty_text(pd.Series(mostConnected), "Name", "# Connections")
    isolatedCompanies = count_isolated_companies(directorNetwork)

    overlayText = (
        '# Positions: ' + str(numPositions) + '\n\n' +
        '# Unique Directors: ' + str(numUniqueDirectors) + '\n\n' +
        '# Isolated Companies: ' + str(isolatedCompanies) + '\n\n' +
        busiestHTML + '\n\n' + mostConnectedHTML
    )

    # 1) Write CSV to repo root
    directorDf.to_csv(CSV_PATH, encoding="utf-8")

    # 2) Build HTML to repo root
    create_network_html(directorNetwork, HTML_PATH)
    with open(HTML_PATH, "r", encoding="utf-8") as f:
    html_text = f.read()

    stats_block = f"""
    <h2>At a glance</h2>
    <p>Total directorships: {numPositions}</p>
    <p>Unique directors: {numUniqueDirectors}</p>
    <h3>Busiest directors (by # boards)</h3>
    {busiestHTML}
    <h3>Most connected directors</h3>
    {mostConnectedHTML}
    <h3>Isolated companies</h3>
    {isolatedCompanies}
    <p><a download href="NZX_Directors.csv">Download full CSV</a></p>
    """

    html_text = html_text.replace("<!--STATS_MARKER-->", stats_block)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html_text)

    # 3) Clean + inject custom CSS and overlay
    remove_html_tags(str(HTML_PATH), ['center', 'h1'])
    insert_css(str(HTML_PATH), css_style)
    remove_html_tags(str(HTML_PATH), ['div'], ['card', 'card-body'])
    insert_html_tag(str(HTML_PATH), "body", 0, "pre", overlayText, newClass="info-text")

    # 4) Insert a Download button that links RELATIVELY to the CSV (works on your website)
    csv_file_url = "NZX_Directors.csv"  # relative to the HTML file location on the server
    insert_html_tag(str(HTML_PATH), "pre", 1, "a", newHref=csv_file_url, priorClass="info-text")
    insert_html_tag(str(HTML_PATH), "a", 0, "button", "Download Data", newClass="top-right-button")
    insert_html_tag(str(HTML_PATH), "a", 1, "div", priorClass="top-right-button", newId='mynetwork')

    print(f"✅ Wrote CSV:  {CSV_PATH}")
    print(f"✅ Wrote HTML: {HTML_PATH}")

    # Open locally, but skip on GitHub Actions
    if OPEN_BROWSER_LOCALLY and os.environ.get("GITHUB_ACTIONS", "").lower() != "true":
        try:
            import webbrowser
            webbrowser.open(HTML_PATH.as_uri())
        except Exception as e:
            print(f"(Info) Could not open browser automatically: {e}")

if __name__ == "__main__":
    main()

