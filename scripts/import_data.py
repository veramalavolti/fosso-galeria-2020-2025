from pathlib import Path
import json, collections, datetime
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
records=[]
sources=[]
for filename in ['risultatiriogaleria20-23.xlsx','risultatiriogaleria24-25.xlsx']:
    ws=openpyxl.load_workbook(ROOT.parent/filename,read_only=True,data_only=True).active
    count=0
    for rownum,r in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
        if not any(v is not None for v in r): continue
        sample,date,param,unit,sign,value,revision,duration,code,station=r
        value=float(str(value).replace(',','.'))
        assert sign in ('<','=') and code in ('F4.18/0','F4.79/0')
        date=str(date)
        datetime.date.fromisoformat(date[:10])
        records.append(dict(sample=str(sample),date=date[:10],timestamp=date,param=param,unit=unit,sign=sign,value=value,station='S1' if code=='F4.79/0' else 'S2',code=code,source=filename,row=rownum,revision=revision))
        count+=1
    sources.append(dict(file=filename,rows=count))
records.sort(key=lambda r:(r['date'],r['station'],r['param']))
keys=[(r['sample'],r['code'],r['param'],r['revision']) for r in records]
assert len(keys)==len(set(keys)), 'Duplicate records require explicit handling'
meta=dict(measurements=len(records),parameters=len(set(r['param'] for r in records)),samples=len(set((r['sample'],r['code']) for r in records)),below=sum(r['sign']=='<' for r in records),start=min(r['date'] for r in records),end=max(r['date'] for r in records),sources=sources,stations={s:dict(count=sum(r['station']==s for r in records),start=min(r['date'] for r in records if r['station']==s),end=max(r['date'] for r in records if r['station']==s)) for s in ['S1','S2']})
(ROOT/'dist/data.json').write_text(json.dumps(dict(meta=meta,records=records),ensure_ascii=False,separators=(',',':')))
print(json.dumps(meta,ensure_ascii=False,indent=2))
