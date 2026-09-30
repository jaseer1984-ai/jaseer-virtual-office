import os, re, json, tempfile
from pathlib import Path
import pandas as pd
import requests
from flask import Flask, render_template, request, jsonify

BASE=Path(__file__).resolve().parent
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=250*1024*1024
app.secret_key=os.getenv('SECRET_KEY','dev-only-change-me')

class DataEngine:
    def __init__(self):
        self.sheets={}; self.source='No data loaded'; self.load_demo()
    def load_demo(self):
        p=BASE/'data'/'Demo_Data.xlsx'
        if p.exists(): self.load_file(p, 'Demo supermarket data')
    def read_file(self,path):
        ext=Path(path).suffix.lower()
        if ext in ('.xlsx','.xls'):
            return pd.read_excel(path,sheet_name=None)
        if ext=='.csv':
            return {'Sales':pd.read_csv(path)}
        raise ValueError('Use XLSX, XLS or CSV')
    def load_file(self,path,label=None):
        self.sheets=self.read_file(path)
        self.source=label or Path(path).name
    def load_files(self, files):
        combined={}; names=[]
        for path,label in files:
            book=self.read_file(path); names.append(label)
            for sheet,df in book.items():
                key=str(sheet)
                if key in combined:
                    combined[key]=pd.concat([combined[key],df],ignore_index=True,sort=False)
                else:
                    combined[key]=df.copy()
        self.sheets=combined
        self.source=f'{len(names)} files: '+', '.join(names[:4])+(' …' if len(names)>4 else '')
        return list(combined)

    def sheet(self,*names):
        for n in names:
            for k,v in self.sheets.items():
                if n.lower() in str(k).lower(): return v.copy()
        return None
    @staticmethod
    def num(s): return pd.to_numeric(s,errors='coerce').fillna(0)
    @staticmethod
    def _find_col(df, *aliases):
        norm={re.sub(r'[^a-z0-9]+','',str(c).lower()):c for c in df.columns}
        for a in aliases:
            k=re.sub(r'[^a-z0-9]+','',a.lower())
            if k in norm: return norm[k]
        return None

    def mtd_sales(self):
        """Read the user's real MTD workbooks (numeric branch sheets such as 102/104/107/111)."""
        frames=[]
        for sheet, raw in self.sheets.items():
            branch=str(sheet).strip()
            if not re.fullmatch(r'\d{2,6}', branch):
                continue
            df=raw.copy()
            date=self._find_col(df,'Date')
            target=self._find_col(df,'Sales Target')
            achieved=self._find_col(df,'Sales Achivement','Sales Achievement','Sales Achieved')
            nob_target=self._find_col(df,'NOB Target')
            nob_ach=self._find_col(df,'NOB Achievemnet','NOB Achievement','NOB Achieved')
            abv_target=self._find_col(df,'ABV Target')
            gp_target=self._find_col(df,'GP Target')
            gp_ach=self._find_col(df,'GP Acheivment','GP Achievement')
            if not (date and target and achieved):
                continue
            x=pd.DataFrame({'Date':pd.to_datetime(df[date],errors='coerce'),
                            'Sales Target':self.num(df[target]),
                            'Sales':self.num(df[achieved])})
            x['Branch']=branch
            x['NOB Target']=self.num(df[nob_target]) if nob_target else 0
            x['NOB']=self.num(df[nob_ach]) if nob_ach else 0
            x['ABV Target']=self.num(df[abv_target]) if abv_target else 0
            x['GP Target']=self.num(df[gp_target]) if gp_target else 0
            x['GP %']=self.num(df[gp_ach]) if gp_ach else 0
            x=x.dropna(subset=['Date'])
            # Ignore blank/template rows that contain no sales/target/NOB values.
            x=x[(x['Sales Target']!=0)|(x['Sales']!=0)|(x['NOB']!=0)]
            if not x.empty: frames.append(x)
        if not frames: return None
        allm=pd.concat(frames,ignore_index=True)
        total=float(allm['Sales'].sum()); target=float(allm['Sales Target'].sum())
        nob=float(allm['NOB'].sum()); nob_target=float(allm['NOB Target'].sum())
        branch=allm.groupby('Branch')['Sales'].sum().sort_values(ascending=False).round(2).to_dict()
        branch_target=allm.groupby('Branch')['Sales Target'].sum().round(2).to_dict()
        branch_kpi={}
        for b,v in branch.items():
            bt=float(branch_target.get(b,0) or 0); sub=allm[allm.Branch==b]; bn=float(sub['NOB'].sum())
            branch_kpi[b]={'sales':float(v),'target':bt,'achievement_pct':float(v)/bt*100 if bt else 0,'nob':bn,'abv':float(v)/bn if bn else 0}
        allm['Month']=allm['Date'].dt.to_period('M').astype(str)
        monthly=allm.groupby('Month')['Sales'].sum().round(2).to_dict()
        monthly_target=allm.groupby('Month')['Sales Target'].sum().round(2).to_dict()
        return {'kind':'mtd','total':total,'target':target,'achievement_pct':total/target*100 if target else 0,
                'nob':nob,'nob_target':nob_target,'nob_achievement_pct':nob/nob_target*100 if nob_target else 0,
                'abv':total/nob if nob else 0,'branch':branch,'branch_kpi':branch_kpi,
                'monthly':monthly,'monthly_target':monthly_target,'rows':len(allm),
                'date_from':allm.Date.min().strftime('%Y-%m-%d'),'date_to':allm.Date.max().strftime('%Y-%m-%d')}

    def sales(self):
        # First recognize the real MTD dashboard format.
        mtd=self.mtd_sales()
        if mtd is not None: return mtd
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
        return {'kind':'detail','total':total,'cost':cost,'gp':gp,'gp_pct':gp/total*100 if total else 0,'branch':branch,'monthly':monthly,'rows':len(df)}
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
        if not isinstance(r, dict) or 'text' in r:
            return (r.get('text','Sales data is not available.') if isinstance(r,dict) else 'Sales data is not available.')
        branch=r.get('branch') or {}; monthly=r.get('monthly') or {}
        total=float(r.get('total',0) or 0)
        if r.get('kind')=='mtd':
            target=float(r.get('target',0) or 0); ach=float(r.get('achievement_pct',0) or 0); nob=float(r.get('nob',0) or 0); abv=float(r.get('abv',0) or 0)
            bk=r.get('branch_kpi') or {}
            # Direct branch request, e.g. branch 107 / 107 sales.
            m=re.search(r'(?<!\d)(\d{3,6})(?!\d)',s)
            if m and m.group(1) in bk:
                b=m.group(1); z=bk[b]
                return f'Branch {b}: Sales SAR {z["sales"]:,.2f} | Target SAR {z["target"]:,.2f} | Achievement {z["achievement_pct"]:.1f}% | NOB {z["nob"]:,.0f} | ABV SAR {z["abv"]:,.2f}.'
            if ('highest' in s or 'best' in s) and branch:
                k=max(branch,key=branch.get); z=bk.get(k,{})
                return f'Highest sales branch is {k}: SAR {branch[k]:,.2f} (target SAR {float(z.get("target",0)):,.2f}, achievement {float(z.get("achievement_pct",0)):.1f}%).'
            if ('lowest' in s or 'worst' in s) and branch:
                k=min(branch,key=branch.get); z=bk.get(k,{})
                return f'Lowest sales branch is {k}: SAR {branch[k]:,.2f} (target SAR {float(z.get("target",0)):,.2f}, achievement {float(z.get("achievement_pct",0)):.1f}%).'
            if 'branch' in s and branch:
                return 'Sales by branch — '+', '.join(f'{k}: SAR {v:,.2f} ({float(bk.get(k,{}).get("achievement_pct",0)):.1f}%)' for k,v in branch.items())
            if 'month' in s and monthly:
                return 'Sales by month — '+', '.join(f'{k}: SAR {v:,.2f}' for k,v in monthly.items())
            if 'nob' in s or 'bill' in s: return f'NOB: {nob:,.0f} | Target: {float(r.get("nob_target",0)):,.0f} | Achievement: {float(r.get("nob_achievement_pct",0)):.1f}%.'
            if 'abv' in s or 'average bill' in s: return f'Overall ABV is SAR {abv:,.2f}, calculated from verified sales and NOB.'
            if mode in ('advise','investigate') and branch and monthly:
                best=max(branch,key=branch.get); worst=min(branch,key=branch.get); bm=max(monthly,key=monthly.get); wm=min(monthly,key=monthly.get)
                return f'I analysed the MTD files from {r.get("date_from")} to {r.get("date_to")}. {best} has the highest sales and {worst} the lowest. {bm} is the strongest month and {wm} the weakest. Overall target achievement is {ach:.1f}% and ABV is SAR {abv:,.2f}. Review NOB and ABV gaps branch-by-branch before changing targets or promotions.'
            return f'MTD Sales: SAR {total:,.2f} | Target: SAR {target:,.2f} | Achievement: {ach:.1f}% | NOB: {nob:,.0f} | ABV: SAR {abv:,.2f} | Period: {r.get("date_from")} to {r.get("date_to")}.'
        cost=float(r.get('cost',0) or 0); gp=float(r.get('gp',total-cost) or 0); gp_pct=float(r.get('gp_pct',(gp/total*100 if total else 0)) or 0)
        if ('highest' in s or 'best' in s) and branch:
            k=max(branch,key=branch.get); return f'Highest sales branch is {k}: SAR {branch[k]:,.2f}.'
        if 'branch' in s and branch: return 'Sales by branch — '+', '.join(f'{k}: SAR {v:,.2f}' for k,v in branch.items())
        return f'Sales: SAR {total:,.2f} | COGS: SAR {cost:,.2f} | GP: SAR {gp:,.2f} ({gp_pct:.1f}%).'
    if dept=='Inventory':
        x=f.get('inventory') or {}; ans=f'Inventory: {int(x.get("negative",0) or 0)} negative-stock item(s), {int(x.get("slow",0) or 0)} slow-moving item(s), stock value SAR {float(x.get("stock_value",0) or 0):,.2f}.'
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
    incoming=request.files.getlist('files') or request.files.getlist('file')
    incoming=[f for f in incoming if f and f.filename]
    if not incoming:return jsonify(error='No files selected'),400
    if len(incoming)>30:return jsonify(error='Maximum 30 files per upload'),400
    saved=[]
    try:
        for f in incoming:
            ext=Path(f.filename).suffix.lower()
            if ext not in ('.xlsx','.xls','.csv'):
                return jsonify(error=f'{f.filename}: use XLSX, XLS or CSV'),400
            t=tempfile.NamedTemporaryFile(suffix=ext,delete=False); t.close(); f.save(t.name)
            saved.append((t.name,f.filename))
        sheets=engine.load_files(saved)
        facts=engine.sales(); return jsonify(ok=True,source=engine.source,sheets=sheets,file_count=len(saved),data_type=facts.get('kind','unknown'),sales_rows=facts.get('rows',0),period=[facts.get('date_from'),facts.get('date_to')] if facts.get('kind')=='mtd' else None)
    except Exception as e:
        return jsonify(error=str(e)),400
    finally:
        for p,_ in saved:
            try: os.unlink(p)
            except: pass

@app.errorhandler(Exception)
def json_error(e):
    app.logger.exception('Unhandled application error')
    if request.path.startswith('/api/'):
        return jsonify(error=f'Server error: {type(e).__name__}: {e}'),500
    raise e

@app.get('/api/status')
def status():
    return jsonify(source=engine.source,llm_configured=bool(os.getenv('LLM_API_URL') and os.getenv('LLM_API_KEY') and os.getenv('LLM_MODEL')),gdrive_configured=bool(os.getenv('GDRIVE_FOLDER_ID')))

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=True)
