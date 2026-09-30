import os, re, json, tempfile
from pathlib import Path
import pandas as pd
import requests
from flask import Flask, render_template, request, jsonify

BASE=Path(__file__).resolve().parent
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=80*1024*1024
app.secret_key=os.getenv('SECRET_KEY','dev-only-change-me')

class DataEngine:
    def __init__(self):
        self.sheets={}; self.source='No data loaded'; self.load_demo()
    def load_demo(self):
        p=BASE/'data'/'Demo_Data.xlsx'
        if p.exists(): self.load_file(p, 'Demo supermarket data')
    def load_file(self,path,label=None):
        ext=Path(path).suffix.lower()
        if ext in ('.xlsx','.xls'):
            self.sheets=pd.read_excel(path,sheet_name=None)
        elif ext=='.csv': self.sheets={'Sales':pd.read_csv(path)}
        else: raise ValueError('Use XLSX, XLS or CSV')
        self.source=label or Path(path).name
    def sheet(self,*names):
        for n in names:
            for k,v in self.sheets.items():
                if n.lower() in str(k).lower(): return v.copy()
        return None
    @staticmethod
    def num(s): return pd.to_numeric(s,errors='coerce').fillna(0)
    def sales(self):
        df=self.sheet('Sales')
        if df is None:return {'text':'Sales data is not loaded.'}
        for c in ['Sales','Cost','Gross Profit','Qty']:
            if c in df: df[c]=self.num(df[c])
        if 'Date' in df: df['Date']=pd.to_datetime(df['Date'],errors='coerce')
        total=float(df['Sales'].sum()) if 'Sales' in df else 0
        cost=float(df['Cost'].sum()) if 'Cost' in df else 0
        gp=float(df['Gross Profit'].sum()) if 'Gross Profit' in df else total-cost
        branch=df.groupby('Branch')['Sales'].sum().sort_values(ascending=False).round(2).to_dict() if {'Branch','Sales'}<=set(df.columns) else {}
        monthly={}
        if {'Date','Sales'}<=set(df.columns):
            monthly=df.dropna(subset=['Date']).assign(Month=lambda x:x.Date.dt.to_period('M').astype(str)).groupby('Month')['Sales'].sum().round(2).to_dict()
        return {'total':total,'cost':cost,'gp':gp,'gp_pct':gp/total*100 if total else 0,'branch':branch,'monthly':monthly,'rows':len(df)}
    def inventory(self):
        df=self.sheet('Inventory')
        if df is None:return {'text':'Inventory data is not loaded.','negative':0,'slow':0}
        q=self.num(df['Closing Qty']) if 'Closing Qty' in df else pd.Series(index=df.index,dtype=float)
        sq=self.num(df['Sales Qty']) if 'Sales Qty' in df else pd.Series(index=df.index,dtype=float)
        neg=df[q<0] if len(q) else df.iloc[:0]; slow=df[sq<=2] if len(sq) else df.iloc[:0]
        return {'negative':len(neg),'slow':len(slow),'negative_rows':neg.head(20).fillna('').to_dict('records'),'stock_value':float(self.num(df['Stock Value']).sum()) if 'Stock Value' in df else 0}
    def finance(self):
        df=self.sheet('Bank')
        if df is None:return {'unmatched':0,'difference':0,'text':'Bank data is not loaded.'}
        if 'Difference' in df: bad=df[self.num(df['Difference']).abs()>.009]
        elif 'Recon Status' in df: bad=df[df['Recon Status'].astype(str).str.lower()!='matched']
        else: bad=df.iloc[:0]
        return {'unmatched':len(bad),'difference':float(self.num(bad['Difference']).sum()) if 'Difference' in bad else 0}
    def supplier(self):
        df=self.sheet('Supplier','Payable','GRN','Purchase')
        if df is None:return {'loaded':False,'text':'Supplier data is not loaded.'}
        ac=next((c for c in ['Outstanding','Balance','Payable','Amount Due','Net Amount','Amount'] if c in df),None)
        total=float(self.num(df[ac]).sum()) if ac else 0
        sc=next((c for c in ['Status','Payment Status','GRN Status'] if c in df),None)
        issues=0
        if sc: issues=int((~df[sc].astype(str).str.lower().isin(['paid','matched','completed','closed','ok'])).sum())
        by={}
        supplier_col=next((c for c in ['Supplier','Supplier Name','Vendor','Vendor Name'] if c in df),None)
        if supplier_col and ac: by=df.groupby(supplier_col)[ac].sum().sort_values(ascending=False).round(2).to_dict()
        return {'loaded':True,'total':total,'issues':issues,'by_supplier':by}
    def audit(self):
        s=self.sheet('Sales'); j=self.sheet('Journal'); ds=dj=un=0
        if s is not None:
            keys=[c for c in ['Date','Invoice No','Branch Code','Item Code','Sales'] if c in s]; ds=int(s.duplicated(subset=keys,keep=False).sum()) if keys else int(s.duplicated(keep=False).sum())
        if j is not None:
            keys=[c for c in ['Date','Voucher No','Branch Code','Account Code','Debit','Credit','Reference'] if c in j]; dj=int(j.duplicated(subset=keys,keep=False).sum()) if keys else int(j.duplicated(keep=False).sum())
            if {'Debit','Credit'}<=set(j.columns): un=int(((self.num(j.Debit).abs()+self.num(j.Credit).abs())>=25000).sum())
        return {'duplicate_sales':ds,'duplicate_journal':dj,'unusual':un}
    def facts(self): return {'sales':self.sales(),'inventory':self.inventory(),'finance':self.finance(),'supplier':self.supplier(),'audit':self.audit(),'source':self.source}

engine=DataEngine()

def route(q):
    s=q.lower(); mode='advise' if any(x in s for x in ['improve','recommend','what should','how can','advice']) else 'investigate' if any(x in s for x in ['why','cause','problem']) else 'analyse'
    if any(x in s for x in ['overall','overal','all issue','any issue','full review']): return mode,'Overall'
    if any(x in s for x in ['supplier','payable','grn']): return mode,'Suppliers'
    if any(x in s for x in ['negative stock','inventory','stock','slow moving']): return mode,'Inventory'
    if any(x in s for x in ['bank','cash flow','cashflow','recon']): return mode,'Finance'
    if any(x in s for x in ['duplicate','audit','unusual','journal issue']): return mode,'Audit'
    if any(x in s for x in ['report','dashboard','mtd']): return mode,'Reporting'
    return mode,'Sales'

def deterministic_answer(q,dept,mode,f):
    s=q.lower(); r=f['sales']
    if dept=='Sales':
        if ('highest' in s or 'best' in s) and r['branch']:
            k=max(r['branch'],key=r['branch'].get); return f'Highest sales branch is {k}: SAR {r["branch"][k]:,.2f}.'
        if 'branch' in s and r['branch']: return 'Sales by branch — '+', '.join(f'{k}: SAR {v:,.2f}' for k,v in r['branch'].items())
        if mode in ('advise','investigate') and r['branch'] and r['monthly']:
            best=max(r['branch'],key=r['branch'].get); worst=min(r['branch'],key=r['branch'].get); bm=max(r['monthly'],key=r['monthly'].get); wm=min(r['monthly'],key=r['monthly'].get)
            return f'I analysed the loaded data. {best} is the strongest branch and {worst} is the weakest. {bm} is the strongest month and {wm} is the weakest. GP margin is {r["gp_pct"]:.1f}%. Compare product mix, customer count and average bill value in the weaker branch/period against the stronger one before changing targets or promotions.'
        return f'Sales: SAR {r["total"]:,.2f} | COGS: SAR {r["cost"]:,.2f} | GP: SAR {r["gp"]:,.2f} ({r["gp_pct"]:.1f}%).'
    if dept=='Inventory':
        x=f['inventory']; ans=f'Inventory: {x["negative"]} negative-stock item(s), {x["slow"]} slow-moving item(s), stock value SAR {x["stock_value"]:,.2f}.'
        if 'negative' in s and x.get('negative_rows'): ans+=' First negative-stock record: '+json.dumps(x['negative_rows'][0],default=str)[:450]
        return ans
    if dept=='Finance': return f'Bank reconciliation: {f["finance"]["unmatched"]} unmatched transaction(s); net difference SAR {f["finance"]["difference"]:,.2f}.'
    if dept=='Audit':
        x=f['audit']; return f'Audit: {x["duplicate_sales"]} duplicated sales row(s), {x["duplicate_journal"]} duplicated journal row(s), {x["unusual"]} unusual journal row(s).'
    if dept=='Suppliers':
        x=f['supplier'];
        if not x['loaded']: return x['text']
        if 'supplier' in s and x['by_supplier']: return 'Supplier balances — '+', '.join(f'{k}: SAR {v:,.2f}' for k,v in x['by_supplier'].items())
        return f'Supplier outstanding: SAR {x["total"]:,.2f}; {x["issues"]} open/exception row(s).'
    if dept=='Reporting': return 'Dashboard request recognized. The MTD builder is ready for the model/data connection; use Data Management to load a workbook. Google Drive sync is configured through server environment settings.'
    return deterministic_answer('total sales','Sales','analyse',f)+' '+deterministic_answer('inventory','Inventory','analyse',f)+' '+deterministic_answer('bank','Finance','analyse',f)+' '+deterministic_answer('audit','Audit','analyse',f)+' '+deterministic_answer('supplier','Suppliers','analyse',f)

def llm_answer(q,dept,mode,f,fallback):
    url=os.getenv('LLM_API_URL','').strip(); key=os.getenv('LLM_API_KEY','').strip(); model=os.getenv('LLM_MODEL','').strip()
    if not (url and key and model): return fallback,False
    compact={'sales':f['sales'],'inventory':{k:v for k,v in f['inventory'].items() if k!='negative_rows'},'finance':f['finance'],'supplier':f['supplier'],'audit':f['audit'],'source':f['source']}
    prompt=f'''You are Jaseer, Head of Department for a supermarket group. User request: {q}\nMode: {mode}; department: {dept}.\nVerified Python facts: {json.dumps(compact,default=str)}\nNever invent numbers. Use only verified facts. Answer concisely as a management analyst. If evidence is insufficient, say what data is needed.'''
    try:
        rr=requests.post(url,headers={'Authorization':f'Bearer {key}','Content-Type':'application/json'},json={'model':model,'messages':[{'role':'user','content':prompt}],'temperature':0.2},timeout=45)
        rr.raise_for_status(); data=rr.json(); return data['choices'][0]['message']['content'].strip(),True
    except Exception: return fallback,False

@app.get('/')
def home(): return render_template('index.html',source=engine.source)
@app.get('/health')
def health(): return {'ok':True,'source':engine.source}
@app.post('/api/chat')
def chat():
    q=(request.json or {}).get('message','').strip()
    if not q:return jsonify(error='Empty message'),400
    mode,dept=route(q); facts=engine.facts(); fallback=deterministic_answer(q,dept,mode,facts); ans,used=llm_answer(q,dept,mode,facts,fallback)
    if dept=='Overall': steps=['Jaseer','Understand Request','Sales + Inventory + Finance + Audit + Suppliers','Consolidate','Jaseer Review']
    elif dept=='Reporting': steps=['Jaseer','Understand Request','Sales Agent','Reporting Agent','Build Dashboard','Jaseer Review']
    else: steps=['Jaseer','Understand Request',dept+' Agent','Analyse Data' if mode=='analyse' else ('Investigate Drivers' if mode=='investigate' else 'Prepare Advice'),'Jaseer Review']
    active=['Sales','Inventory','Finance','Audit','Suppliers'] if dept=='Overall' else [dept]
    return jsonify(answer=ans,department=dept,mode=mode,steps=steps,active=active,llm=used,source=engine.source)
@app.post('/api/upload')
def upload():
    f=request.files.get('file')
    if not f:return jsonify(error='No file'),400
    ext=Path(f.filename).suffix.lower()
    if ext not in ('.xlsx','.xls','.csv'): return jsonify(error='Use XLSX, XLS or CSV'),400
    with tempfile.NamedTemporaryFile(suffix=ext,delete=False) as t: f.save(t.name); p=t.name
    try: engine.load_file(p,f.filename); return jsonify(ok=True,source=engine.source,sheets=list(engine.sheets))
    except Exception as e:return jsonify(error=str(e)),400
    finally:
        try:os.unlink(p)
        except:pass
@app.get('/api/status')
def status():
    return jsonify(source=engine.source,llm_configured=bool(os.getenv('LLM_API_URL') and os.getenv('LLM_API_KEY') and os.getenv('LLM_MODEL')),gdrive_configured=bool(os.getenv('GDRIVE_FOLDER_ID')))

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=True)
