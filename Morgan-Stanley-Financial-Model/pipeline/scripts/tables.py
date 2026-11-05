import sys, re, html
from bs4 import BeautifulSoup
def clean(s):
    s = html.unescape(s).replace('\xa0',' ')
    return re.sub(r'\s+',' ',s).strip()
def parse(path):
    soup = BeautifulSoup(open(path,encoding='utf-8',errors='ignore').read(), 'lxml')
    out=[]
    for ti,tb in enumerate(soup.find_all('table')):
        rows=[]
        for tr in tb.find_all('tr'):
            cells=[clean(td.get_text(' ')) for td in tr.find_all(['td','th'])]
            cells=[c for c in cells if c not in ('','$','%',')')]
            if cells: rows.append(cells)
        if rows: out.append(rows)
    return out
if __name__=='__main__':
    for ti,rows in enumerate(parse(sys.argv[1])):
        print(f'=== TABLE {ti}')
        for r in rows: print(' | '.join(r))
