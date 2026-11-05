import json
def labels(t): return [r['label'] for r in t['rows']]
def has(t, s): return any(l.startswith(s) for l in labels(t))
def classify(doc):
    out=[]; seg_order=0
    for t in doc:
        L=labels(t); c='?'
        if has(t,'Brokerage, clearing and exchange fees'): c='cons_is'
        elif has(t,'Advisory revenues') or has(t,'Total sales & trading net revenues'): c='isg_info'
        elif has(t,'Advisory') and has(t,'Compensation and benefits'): c='isg'
        elif has(t,'Transactional') and has(t,'Compensation and benefits'): c='wm'
        elif has(t,'Asset management and related fees') or (has(t,'Performance-based income')): c='im'
        elif has(t,'Asset management') and has(t,'Investments') and has(t,'Compensation and benefits') and not has(t,'Trading'): c='im'
        elif has(t,'Investment banking') and has(t,'Commissions and fees') and has(t,'Compensation and benefits'):
            c=['isg_li','wm_li','im_li'][min(seg_order,2)]; seg_order+=1
        elif any('lient assets' in l for l in L) and not has(t,'Advisory'): c='wm_info'
        elif any('Assets Under Management or Supervision' in l or 'Assets under management or supervision' in l or l.startswith('Total AUM') for l in L): c='im_info'
        elif has(t,'Common Equity Tier 1 capital ratio'): c='capital'
        elif any('Loans and Lending Commitments' in l for l in L): c='loans'
        elif L[:2]==['Institutional Securities','Wealth Management']: c='summary'
        elif has(t,'Earnings per diluted share') or has(t,'Earnings per basic share'): c='metrics'
        elif has(t,'Total assets') or has(t,'Total Assets'): c='bs_info'
        elif any(l.startswith('U.S. Bank') for l in L): c='usbank'
        out.append(c)
    return out
if __name__=='__main__':
    d=json.load(open('data/extracted.json'))
    for k in sorted(d):
        if not k.startswith('fs'): continue
        cl=classify(d[k])
        print(k, ' '.join(f"{c}" for c in cl))
