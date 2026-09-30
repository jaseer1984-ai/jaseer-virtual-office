import os, re, json, tempfile
from pathlib import Path
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from flask import Flask, render_template, request, jsonify, send_file

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


BRANCH_NAMES={'102':'AL KHAIR','104':'NOORA','107':'MAGNUS YAQOOT','109':'HAMRA','111':'MAGNUS EXPRESS','112':'MAGNUS HADA SHAM'}
DASHBOARD_DIR=BASE/'cache'
DASHBOARD_DIR.mkdir(exist_ok=True)
LATEST_DASHBOARD=DASHBOARD_DIR/'Jaseer_MTD_Performance_Dashboard.xlsx'

def _mtd_frame():
    """Normalize all loaded numeric branch sheets into one daily MTD table."""
    frames=[]
    for sheet,raw in engine.sheets.items():
        branch=str(sheet).strip()
        if not re.fullmatch(r'\d{2,6}',branch):
            continue
        df=raw.copy()
        date=engine._find_col(df,'Date'); target=engine._find_col(df,'Sales Target')
        achieved=engine._find_col(df,'Sales Achivement','Sales Achievement','Sales Achieved')
        if not (date and target and achieved): continue
        nt=engine._find_col(df,'NOB Target'); na=engine._find_col(df,'NOB Achievemnet','NOB Achievement','NOB Achieved')
        at=engine._find_col(df,'ABV Target'); aa=engine._find_col(df,'ABV Achievement','ABV Achivement')
        gt=engine._find_col(df,'GP Target','GP Target '); ga=engine._find_col(df,'GP Acheivment','GP Achievement')
        x=pd.DataFrame({
            'Date':pd.to_datetime(df[date],errors='coerce'),'Branch Code':branch,
            'Branch':f'{branch} - {BRANCH_NAMES.get(branch,branch)}',
            'Sales Target':engine.num(df[target]),'Sales Achieved':engine.num(df[achieved]),
            'NOB Target':engine.num(df[nt]) if nt else 0,'NOB Achieved':engine.num(df[na]) if na else 0,
            'ABV Target':engine.num(df[at]) if at else 0,'ABV Source':engine.num(df[aa]) if aa else 0,
            'GP Target':engine.num(df[gt]) if gt else 0,'GP Achievement':engine.num(df[ga]) if ga else 0,
        })
        x=x.dropna(subset=['Date'])
        x=x[(x['Sales Target']!=0)|(x['Sales Achieved']!=0)|(x['NOB Target']!=0)|(x['NOB Achieved']!=0)]
        if not x.empty: frames.append(x)
    if not frames: return pd.DataFrame()
    data=pd.concat(frames,ignore_index=True)
    # Multiple uploaded workbooks may overlap. Keep the latest occurrence for each branch/date.
    data=data.sort_values(['Date','Branch Code']).drop_duplicates(['Date','Branch Code'],keep='last').reset_index(drop=True)
    data['Sales Achievement']=data['Sales Achieved'].div(data['Sales Target'].replace(0,pd.NA)).fillna(0)
    data['NOB Achievement']=data['NOB Achieved'].div(data['NOB Target'].replace(0,pd.NA)).fillna(0)
    data['ABV Actual']=data['Sales Achieved'].div(data['NOB Achieved'].replace(0,pd.NA)).fillna(0)
    data['Gap']=data['Sales Achieved']-data['Sales Target']
    data['Year']=data['Date'].dt.year; data['Month']=data['Date'].dt.to_period('M').astype(str)
    return data

def build_mtd_dashboard(output_path=LATEST_DASHBOARD):
    data=_mtd_frame()
    if data.empty: raise ValueError('No usable MTD branch data is loaded.')
    latest=data['Date'].max(); period=latest.to_period('M')
    current=data[data['Date'].dt.to_period('M')==period].copy()
    if current.empty: raise ValueError('No rows found for the latest MTD period.')
    elapsed=int(current['Date'].max().day); days_in_month=int(period.days_in_month); days_remaining=max(days_in_month-elapsed,0)
    prev_period=period-1
    previous=data[data['Date'].dt.to_period('M')==prev_period].copy()
    previous=previous[previous['Date'].dt.day<=min(elapsed,prev_period.days_in_month)]
    sales=float(current['Sales Achieved'].sum()); target=float(current['Sales Target'].sum()); variance=sales-target
    nob=float(current['NOB Achieved'].sum()); nob_target=float(current['NOB Target'].sum())
    abv=sales/nob if nob else 0; abv_target=(float(current['Sales Target'].sum())/float(current['NOB Target'].sum())) if current['NOB Target'].sum() else 0
    prev_sales=float(previous['Sales Achieved'].sum()); change=(sales/prev_sales-1) if prev_sales else 0
    daily=current.groupby('Date',as_index=False).agg(Target=('Sales Target','sum'),Actual=('Sales Achieved','sum'))
    bg=current.groupby(['Branch Code','Branch'],as_index=False).agg(Target=('Sales Target','sum'),Actual=('Sales Achieved','sum'),NOB_Target=('NOB Target','sum'),NOB_Actual=('NOB Achieved','sum'))
    bg['Achievement']=bg['Actual'].div(bg['Target'].replace(0,pd.NA)).fillna(0); bg['Gap']=bg['Actual']-bg['Target']; bg['NOB Achievement']=bg['NOB_Actual'].div(bg['NOB_Target'].replace(0,pd.NA)).fillna(0)
    best=bg.loc[bg['Achievement'].idxmax()]; worst=bg.loc[bg['Achievement'].idxmin()]
    month_target_daily=target/max(elapsed,1); projected_target=month_target_daily*days_in_month
    forecast=(sales/max(elapsed,1))*days_in_month; forecast_gap=forecast-projected_target
    remaining=max(projected_target-sales,0); required_daily=remaining/days_remaining if days_remaining else 0

    wb=Workbook(); ws=wb.active; ws.title='MTD Dashboard'; detail=wb.create_sheet('Sales Detail'); summary=wb.create_sheet('Branch Summary'); monthly_ws=wb.create_sheet('Monthly Summary')
    navy='14213D'; blue='1F4E78'; teal='0F766E'; light='EAF2F8'; white='FFFFFF'; red='C00000'; amber='F4B183'; green='70AD47'; gray='D9E1F2'
    thin=Side(style='thin',color='D9E2F3')
    for c in range(1,14): ws.column_dimensions[get_column_letter(c)].width=15
    ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=18; ws.column_dimensions['G'].width=24
    ws.sheet_view.showGridLines=False; ws.freeze_panes='A6'
    ws.merge_cells('A1:L1'); ws['A1']='MTD PERFORMANCE DASHBOARD'; ws['A1'].font=Font(size=22,bold=True,color=white); ws['A1'].fill=PatternFill('solid',fgColor=navy); ws['A1'].alignment=Alignment(horizontal='left',vertical='center'); ws.row_dimensions[1].height=34
    ws.merge_cells('A2:L2'); ws['A2']='Branch performance, targets, bills, average bill value and management comparisons'; ws['A2'].font=Font(size=11,color='44546A')
    ws['A4']='REPORT FILTERS'; ws['A4'].font=Font(bold=True,color=white); ws['A4'].fill=PatternFill('solid',fgColor=blue)
    filters=[('A5','Branch','B5','All Branches'),('D5','Year','E5',int(latest.year)),('G5','Quarter','H5',f'Q{((latest.month-1)//3)+1}'),('J5','Month','K5',latest.strftime('%b-%Y'))]
    for lc,lbl,vc,val in filters: ws[lc]=lbl; ws[lc].font=Font(bold=True); ws[vc]=val; ws[vc].fill=PatternFill('solid',fgColor=light)
    ws.merge_cells('A7:L7'); ws['A7']='KEY PERFORMANCE INDICATORS'; ws['A7'].font=Font(bold=True,color=white); ws['A7'].fill=PatternFill('solid',fgColor=blue)
    kpis=[('A9','SALES TARGET',target,'#,##0.00'),('C9','SALES ACHIEVED',sales,'#,##0.00'),('E9','SALES ACHIEVEMENT',sales/target if target else 0,'0.0%'),('G9','SALES VARIANCE',variance,'#,##0.00'),('I9','NOB TARGET',nob_target,'#,##0'),('K9','NOB ACHIEVED',nob,'#,##0'),
          ('A13','NOB ACHIEVEMENT',nob/nob_target if nob_target else 0,'0.0%'),('C13','ABV TARGET',abv_target,'#,##0.00'),('E13','ABV ACTUAL',abv,'#,##0.00'),('G13','PREVIOUS PERIOD SALES',prev_sales,'#,##0.00'),('I13','PERIOD CHANGE',change,'0.0%'),('K13','TOP BRANCH',str(best['Branch']),'@')]
    for cell,label,val,fmt in kpis:
        col=ws[cell].column; row=ws[cell].row; ws[cell]=label; ws[cell].font=Font(bold=True,color='44546A'); v=ws.cell(row+1,col); v.value=val; v.number_format=fmt; v.font=Font(size=14,bold=True,color=navy); v.fill=PatternFill('solid',fgColor='F7F9FC'); v.border=Border(bottom=thin); ws.merge_cells(start_row=row+1,start_column=col,end_row=row+1,end_column=min(col+1,12))
    ws.merge_cells('A17:L17'); ws['A17']='MANAGEMENT INSIGHTS'; ws['A17'].font=Font(bold=True,color=white); ws['A17'].fill=PatternFill('solid',fgColor=blue)
    insights=[
        ('ATTENTION' if sales/target<.9 else 'ON TRACK',f'Sales achievement is {sales/target:.1%}; variance to MTD target is SAR {variance:,.0f}.' if target else 'Sales target is unavailable.'),
        ('ATTENTION' if change<0 else 'POSITIVE',f'Sales changed {change:+.1%} versus the previous comparable period (SAR {prev_sales:,.0f}).' if prev_sales else 'Previous comparable period is unavailable.'),
        ('ATTENTION' if nob/nob_target<.9 else 'ON TRACK',f'Bills reached {nob/nob_target:.1%} of target ({nob:,.0f} achieved vs {nob_target:,.0f} target).' if nob_target else 'NOB target is unavailable.'),
        ('INSIGHT',f'Best branch: {best["Branch"]} at {best["Achievement"]:.1%}; priority branch: {worst["Branch"]} at {worst["Achievement"]:.1%}.')]
    for i,(tag,txt) in enumerate(insights,19): ws[f'A{i}']=tag; ws[f'A{i}'].font=Font(bold=True,color=red if tag=='ATTENTION' else green); ws.merge_cells(start_row=i,start_column=3,end_row=i,end_column=12); ws.cell(i,3).value=txt; ws.cell(i,3).alignment=Alignment(wrap_text=True)
    ws.merge_cells('A25:L25'); ws['A25']='FORECAST & PACE'; ws['A25'].font=Font(bold=True,color=white); ws['A25'].fill=PatternFill('solid',fgColor=blue)
    forecasts=[('A27','MONTH-END FORECAST',forecast),('C27','PROJECTED MONTH TARGET',projected_target),('E27','FORECAST GAP',forecast_gap),('G27','REMAINING TO TARGET',remaining),('I27','REQUIRED DAILY SALES',required_daily),('K27','DAYS REMAINING',days_remaining)]
    for cell,label,val in forecasts: ws[cell]=label; ws[cell].font=Font(bold=True,color='44546A'); c=ws[cell].column; r=ws[cell].row; ws.cell(r+1,c).value=val; ws.cell(r+1,c).number_format='#,##0.00' if 'DAYS' not in label else '0'; ws.cell(r+1,c).font=Font(size=12,bold=True,color=navy)
    ws['A31']='DAILY SALES TREND • SAR'; ws['A31'].font=Font(bold=True,color=white); ws['A31'].fill=PatternFill('solid',fgColor=teal)
    ws['G31']='BRANCH TARGET VS ACHIEVED • SAR'; ws['G31'].font=Font(bold=True,color=white); ws['G31'].fill=PatternFill('solid',fgColor=teal)
    # chart source tables lower on dashboard
    ws['A60']='Date'; ws['B60']='Target'; ws['C60']='Actual'
    for i,row in enumerate(daily.itertuples(index=False),61): ws.cell(i,1).value=row.Date.to_pydatetime(); ws.cell(i,1).number_format='dd-mmm'; ws.cell(i,2).value=float(row.Target); ws.cell(i,3).value=float(row.Actual)
    ws['G60']='Branch'; ws['H60']='Target'; ws['I60']='Actual'; ws['J60']='Achievement'; ws['K60']='NOB Achievement'
    for i,row in enumerate(bg.itertuples(index=False),61): ws.cell(i,7).value=row.Branch; ws.cell(i,8).value=float(row.Target); ws.cell(i,9).value=float(row.Actual); ws.cell(i,10).value=float(row.Achievement); ws.cell(i,11).value=float(row._8 if hasattr(row,'_8') else row[-1])
    lc=LineChart(); lc.title='Daily Sales Trend'; lc.y_axis.title='SAR'; lc.height=8; lc.width=14; lc.add_data(Reference(ws,min_col=2,max_col=3,min_row=60,max_row=60+len(daily)),titles_from_data=True); lc.set_categories(Reference(ws,min_col=1,min_row=61,max_row=60+len(daily))); ws.add_chart(lc,'A33')
    bc=BarChart(); bc.type='col'; bc.style=10; bc.title='Branch Target vs Achieved'; bc.y_axis.title='SAR'; bc.height=8; bc.width=14; bc.add_data(Reference(ws,min_col=8,max_col=9,min_row=60,max_row=60+len(bg)),titles_from_data=True); bc.set_categories(Reference(ws,min_col=7,min_row=61,max_row=60+len(bg))); ws.add_chart(bc,'G33')
    ws['A49']='CURRENT VS PREVIOUS PERIOD'; ws['A49'].font=Font(bold=True,color=white); ws['A49'].fill=PatternFill('solid',fgColor=teal)
    ws['A51']='Previous Period'; ws['B51']=prev_sales; ws['A52']='Selected Period'; ws['B52']=sales; ws['B51'].number_format=ws['B52'].number_format='#,##0.00'
    cc=BarChart(); cc.type='col'; cc.title='Current vs Previous Period'; cc.height=6; cc.width=8; cc.add_data(Reference(ws,min_col=2,min_row=50,max_row=52),titles_from_data=True); cc.set_categories(Reference(ws,min_col=1,min_row=51,max_row=52)); ws.add_chart(cc,'A53')
    ws['G49']='BRANCH PERFORMANCE HEATMAP'; ws['G49'].font=Font(bold=True,color=white); ws['G49'].fill=PatternFill('solid',fgColor=teal)
    ws.append([]) if False else None
    hr=51; ws[f'G{hr}']='BRANCH'; ws[f'H{hr}']='SALES %'; ws[f'I{hr}']='GAP • SAR'; ws[f'J{hr}']='STATUS'
    for j,row in enumerate(bg.sort_values('Achievement',ascending=False).itertuples(index=False),hr+1):
        ws.cell(j,7).value=row.Branch; ws.cell(j,8).value=float(row.Achievement); ws.cell(j,8).number_format='0.0%'; ws.cell(j,9).value=float(row.Gap); ws.cell(j,9).number_format='#,##0.00'; ws.cell(j,10).value='ON TRACK' if row.Achievement>=1 else ('WATCH' if row.Achievement>=.85 else 'CRITICAL')
    if len(bg): ws[f'H{hr+1}:H{hr+len(bg)}'][0][0].parent[f'H{hr+1}:H{hr+len(bg)}']
    ws.conditional_formatting.add(f'H{hr+1}:H{hr+len(bg)}',ColorScaleRule(start_type='num',start_value=0,start_color='F8696B',mid_type='num',mid_value=.9,mid_color='FFEB84',end_type='num',end_value=1.05,end_color='63BE7B'))
    ws['K49']='BILL / NOB ACHIEVEMENT'; ws['K49'].font=Font(bold=True,color=white); ws['K49'].fill=PatternFill('solid',fgColor=teal)
    # Detail sheet
    headers=['Date','Branch Code','Branch','Sales Target','Sales Achieved','Sales Achievement','Gap','NOB Target','NOB Achieved','NOB Achievement','ABV Target','ABV Actual','GP Target','GP Achievement','Year','Month']
    detail.append(headers)
    for row in data[headers].itertuples(index=False,name=None): detail.append([x.to_pydatetime() if isinstance(x,pd.Timestamp) else (float(x) if isinstance(x,(float,int)) and not isinstance(x,bool) else x) for x in row])
    detail.freeze_panes='A2'; detail.auto_filter.ref=detail.dimensions; detail.sheet_view.showGridLines=False
    for c in detail[1]: c.fill=PatternFill('solid',fgColor=navy); c.font=Font(bold=True,color=white)
    for col in ['A']: detail.column_dimensions[col].width=13
    detail.column_dimensions['C'].width=25
    for col in range(4,15): detail.column_dimensions[get_column_letter(col)].width=16
    for r in range(2,detail.max_row+1): detail.cell(r,1).number_format='dd-mmm-yyyy';
    for col in [4,5,7,11,12,13,14]:
        for r in range(2,detail.max_row+1): detail.cell(r,col).number_format='#,##0.00'
    for col in [6,10]:
        for r in range(2,detail.max_row+1): detail.cell(r,col).number_format='0.0%'
    # Branch Summary
    sh=['Branch','Target','Achieved','Achievement %','Gap','NOB Target','NOB Achieved','NOB %','ABV']
    summary.append(sh)
    for row in bg.itertuples(index=False): summary.append([row.Branch,float(row.Target),float(row.Actual),float(row.Achievement),float(row.Gap),float(row.NOB_Target),float(row.NOB_Actual),float(row[-1]),float(row.Actual/row.NOB_Actual if row.NOB_Actual else 0)])
    for c in summary[1]: c.fill=PatternFill('solid',fgColor=navy); c.font=Font(bold=True,color=white)
    summary.freeze_panes='A2'; summary.auto_filter.ref=summary.dimensions; summary.column_dimensions['A'].width=28
    # Monthly Summary
    mg=data.groupby('Month',as_index=False).agg(Target=('Sales Target','sum'),Achieved=('Sales Achieved','sum'),NOB=('NOB Achieved','sum'))
    mg['Achievement']=mg['Achieved'].div(mg['Target'].replace(0,pd.NA)).fillna(0); mg['ABV']=mg['Achieved'].div(mg['NOB'].replace(0,pd.NA)).fillna(0)
    monthly_ws.append(['Month','Target','Achieved','Achievement %','NOB','ABV'])
    for row in mg.itertuples(index=False): monthly_ws.append([row.Month,float(row.Target),float(row.Achieved),float(row.Achievement),float(row.NOB),float(row.ABV)])
    for c in monthly_ws[1]: c.fill=PatternFill('solid',fgColor=navy); c.font=Font(bold=True,color=white)
    monthly_ws.freeze_panes='A2'; monthly_ws.auto_filter.ref=monthly_ws.dimensions
    for shx in [summary,monthly_ws]:
        for col in range(1,shx.max_column+1): shx.column_dimensions[get_column_letter(col)].width=max(14,min(28,max(len(str(shx.cell(r,col).value or '')) for r in range(1,min(shx.max_row,50)+1))+2))
    # Hide chart helper tables but keep the visible dashboard clean.
    for row in range(60,60+max(len(daily),len(bg))+2): ws.row_dimensions[row].hidden=True
    wb.save(output_path)
    return {'path':str(output_path),'period':period.strftime('%b-%Y'),'rows':len(data),'branches':len(bg)}

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
    if dept=='Reporting': return 'MTD dashboard request recognized. Use “make MTD dashboard” and I will build the complete Excel dashboard from the loaded files.'
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
    mode,dept=route(q); facts=engine.facts()
    wants_dashboard=dept=='Reporting' and any(x in q.lower() for x in ['make','create','build','generate','export','download'])
    dashboard_meta=None
    if wants_dashboard:
        try:
            dashboard_meta=build_mtd_dashboard()
            ans=f'Complete MTD Excel dashboard created for {dashboard_meta["period"]} from {dashboard_meta["rows"]} daily branch rows across {dashboard_meta["branches"]} branches.'
            used=False
        except Exception as e:
            return jsonify(error=f'Dashboard creation failed: {e}'),400
    else:
        fallback=deterministic_answer(q,dept,mode,facts); ans,used=llm_answer(q,dept,mode,facts,fallback)
    if dept=='Overall': steps=['Jaseer','Understand Request','Sales + Inventory + Finance + Audit + Suppliers','Consolidate','Jaseer Review']
    elif dept=='Reporting': steps=['Jaseer','Understand Request','Sales Agent','Reporting Agent','Build Dashboard','Jaseer Review']
    else: steps=['Jaseer','Understand Request',dept+' Agent','Analyse Data' if mode=='analyse' else ('Investigate Drivers' if mode=='investigate' else 'Prepare Advice'),'Jaseer Review']
    active=['Sales','Inventory','Finance','Audit','Suppliers'] if dept=='Overall' else [dept]
    return jsonify(answer=ans,department=dept,mode=mode,steps=steps,active=active,llm=used,source=engine.source,download_url=('/api/dashboard/download' if dashboard_meta else None),dashboard_ready=bool(dashboard_meta and LATEST_DASHBOARD.exists()),dashboard_filename=('Jaseer_MTD_Performance_Dashboard.xlsx' if dashboard_meta else None))
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

@app.get('/api/dashboard/download')
def dashboard_download():
    if not LATEST_DASHBOARD.exists():
        return jsonify(error='No dashboard has been created yet.'),404
    resp = send_file(LATEST_DASHBOARD,as_attachment=True,download_name='Jaseer_MTD_Performance_Dashboard.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',conditional=True)
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return resp

@app.errorhandler(Exception)
def json_error(e):
    app.logger.exception('Unhandled application error')
    if request.path.startswith('/api/'):
        return jsonify(error=f'Server error: {type(e).__name__}: {e}'),500
    raise e

@app.get('/api/status')
def status():
    return jsonify(source=engine.source,llm_configured=bool(os.getenv('LLM_API_URL') and os.getenv('LLM_API_KEY') and os.getenv('LLM_MODEL')),gdrive_configured=bool(os.getenv('GDRIVE_FOLDER_ID')),dashboard_ready=LATEST_DASHBOARD.exists(),download_url=('/api/dashboard/download' if LATEST_DASHBOARD.exists() else None))

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=True)
