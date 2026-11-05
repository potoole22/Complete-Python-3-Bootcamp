import json,subprocess,time,os,re,sys
from concurrent.futures import ThreadPoolExecutor
rows=json.load(open('periodic.json'))
UA="FinModel Research research@example.com"
def get(url):
    for i in range(3):
        r=subprocess.run(['curl','-sS','--max-time','30','-A',UA,url],capture_output=True,text=True)
        if r.returncode==0 and r.stdout: return r.stdout
        time.sleep(1+i)
    return ''
def work(row):
    rep,form,fd,acc,prim=row
    a=acc.replace('-','')
    base=f"https://www.sec.gov/Archives/edgar/data/895421/{a}/"
    fs=get(base+'FilingSummary.xml')
    reps=re.findall(r'<Report instance=.*?</Report>',fs,re.S)
    want={}
    for blk in reps:
        fn=re.search(r'<HtmlFileName>(.*?)</HtmlFileName>',blk); sn=re.search(r'<ShortName>(.*?)</ShortName>',blk)
        if not fn or not sn: continue
        fn=fn.group(1); s=sn.group(1).strip().lower()
        if 'balance sheet' in s and 'parenthetical' not in s and 'bs' not in want: want['bs']=fn
        if 'cash flow' in s and 'cf' not in want: want['cf']=fn
        if 'income statement' in s and 'comprehensive' not in s and 'is' not in want: want['is']=fn
    for k,fn in want.items():
        out=f"rfiles/{rep}_{k}.htm"
        if not os.path.exists(out) or os.path.getsize(out)<1000:
            open(out,'w').write(get(base+fn))
    return rep,want
with ThreadPoolExecutor(4) as ex:
    for rep,want in ex.map(work, rows): print(rep,want,flush=True)
