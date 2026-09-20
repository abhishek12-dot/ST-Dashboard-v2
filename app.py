"""
ST Dashboard v2 - New Supertails Delivery Performance Website
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import tempfile, os, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from process_orders import run_full_pipeline

st.set_page_config(page_title="ST Dashboard v2", page_icon="🚚", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.main-header{font-size:2.2rem;font-weight:700;color:#1a365d;margin-bottom:0}
.sub-header{font-size:1rem;color:#718096;margin-top:0}
div[data-testid="stMetricValue"]{font-size:1.6rem}
.stDownloadButton>button{width:100%}
</style>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🚚 ST Dashboard v2")
    st.caption("Supertails Delivery Performance")
    st.markdown("---")
    mode = st.radio("Report Mode", ["daily","all"], index=0)
    report_date = st.date_input("Report Date", value=date.today())
    st.markdown("---")
    st.markdown("**Upload 3 store CSVs**")
    kalyan_file = st.file_uploader("1. Kalyan Nagar", type=["csv"], key="k")
    varthur_file = st.file_uploader("2. Varthur", type=["csv"], key="v")
    indranagar_file = st.file_uploader("3. Indranagar", type=["csv"], key="i")
    st.markdown("---")
    st.subheader("Rider Count")
    rider_kalyan = st.number_input("Kalyan Nagar", 0, value=12, step=1, key="rk")
    rider_indra = st.number_input("Indranagar", 0, value=6, step=1, key="ri")
    rider_varthur = st.number_input("Varthur", 0, value=6, step=1, key="rv")
    rider_map = {"Kalyan Nagar": rider_kalyan, "Indranagar": rider_indra, "Varthur": rider_varthur}
    st.markdown("---")
    st.code("Kalyan → Supertail_Hyperlocal\nIndranagar → Supertails-Indranagar\nVarthur → Supertails-Varthur")
    st.caption("Final Verdict = Yes only if delay ≥ 2 min")

st.markdown('<p class="main-header">ST Dashboard v2</p>', unsafe_allow_html=True)
st.markdown(f'<p class="sub-header">{datetime.now().strftime("%d %b %Y %H:%M")}</p>', unsafe_allow_html=True)

c1,c2,_ = st.columns([1,1,3])
with c1:
    run = st.button("Generate Reports", type="primary", use_container_width=True)
with c2:
    if st.button("Clear", use_container_width=True):
        st.session_state.pop("result", None)
        st.rerun()

if run:
    if not (kalyan_file and varthur_file and indranagar_file):
        st.error("Upload all 3 CSV files")
        st.stop()
    try:
        with tempfile.TemporaryDirectory() as tmp:
            for p,f in [("k.csv",kalyan_file),("v.csv",varthur_file),("i.csv",indranagar_file)]:
                open(os.path.join(tmp,p),"wb").write(f.getbuffer())
            result = run_full_pipeline(
                os.path.join(tmp,"k.csv"), os.path.join(tmp,"v.csv"), os.path.join(tmp,"i.csv"),
                report_date=datetime.combine(report_date, datetime.min.time()), mode=mode)
        st.session_state["result"] = result
        st.session_state["report_date"] = report_date
        st.rerun()
    except Exception as e:
        st.error(str(e)); st.exception(e); st.stop()

if "result" not in st.session_state:
    st.info("Upload 3 CSVs and click Generate Reports")
    st.stop()

result = st.session_state["result"]
df, breach, pendency, eod = result["processed_data"], result["breach"].copy(), result["pendency"].copy(), result["eod"].copy()
report_dt = st.session_state.get("report_date", date.today())

def add_riders(rdf):
    rdf = rdf.copy()
    rdf["Rider Count"] = rdf["Store"].map(rider_map).fillna(0).astype(int)
    def avg(r):
        if r["Rider Count"]<=0: return 0.0
        for k in ["Delivered","Total Orders","Total Operational Pending"]:
            if k in r and pd.notna(r.get(k)): return round(float(r[k])/r["Rider Count"],1)
        return 0.0
    rdf["Avg Order/Rider"] = rdf.apply(avg, axis=1)
    return rdf

breach, pendency, eod = add_riders(breach), add_riders(pendency), add_riders(eod)
to, td, tb = int(breach["Total Orders"].sum()), int(breach["Delivered"].sum()), int(breach["Total Breach"].sum())
bp = round(tb/td*100,1) if td else 0

k1,k2,k3,k4,k5,k6 = st.columns(6)
k1.metric("Total Orders", to); k2.metric("Delivered", td); k3.metric("Total Breach", tb)
k4.metric("Breach %", f"{bp}%"); k5.metric("On-Time", td-tb); k6.metric("Riders", sum(rider_map.values()))
st.markdown("---")

tabs = st.tabs(["EOD","Pendency","Breach","Breach IDs","Detail","Downloads"])

with tabs[0]:
    st.subheader("EOD Consolidated")
    for _,row in eod.iterrows():
        with st.expander(f"**{row['Store']}**", expanded=True):
            a,b,c,d,e = st.columns(5)
            a.metric("Orders", int(row.get("Total Orders",0)))
            b.metric("Delivered", int(row.get("Delivered",0)))
            c.metric("Breach", int(row.get("Total Breach",0)))
            d.metric("Riders", int(row["Rider Count"]))
            e.metric("Avg/Rider", row["Avg Order/Rider"])
    cols = [c for c in ["Store","Total Orders","Delivered","On Time Delivered","Total Breach","Breach %","Rider Count","Avg Order/Rider"] if c in eod.columns]
    st.dataframe(eod[cols], use_container_width=True, hide_index=True)

with tabs[1]:
    st.subheader("Pendency")
    pref = ["Store","Rider Count","Avg Order/Rider"]
    st.dataframe(pendency[pref+[c for c in pendency.columns if c not in pref]], use_container_width=True, hide_index=True)
    if "Total Operational Pending" in pendency.columns:
        st.bar_chart(pendency.set_index("Store")[["Total Operational Pending"]])

with tabs[2]:
    st.subheader("Breach Report")
    st.caption("Final Verdict = Yes only if delay ≥ 2 min")
    cols = [c for c in ["Store","Total Orders","Delivered","15 Min Orders","30 Min Orders","Other Orders",
                        "Total Breach","15 Min Breach","30 Min Breach","Other Breach","Breach %","Rider Count","Avg Order/Rider"] if c in breach.columns]
    st.dataframe(breach[cols], use_container_width=True, hide_index=True)
    st.bar_chart(breach.set_index("Store")[["Total Breach"]])
    st.bar_chart(breach.set_index("Store")[["Breach %"]])

with tabs[3]:
    st.subheader("Breach Order IDs")
    bo = df[df["Final_Verdict"]=="Yes"].copy()
    if len(bo)==0: st.success("No breaches")
    else:
        st.markdown(f"**{len(bo)} breach(es)**")
        dcols = [c for c in ["CDR ID","Reference ID","Channel","Customer Name","Model","Order_datetime",
                             "ADT(Actual Delivery Time)","Breach_delay","SLA_STATUS","Final_Verdict","Rider Name","Last Failed Remark"] if c in bo.columns]
        st.dataframe(bo[dcols].sort_values("Order_datetime", ascending=False), use_container_width=True, hide_index=True, height=400)
        st.download_button("Download Breach CSV", bo[dcols].to_csv(index=False).encode(), f"breach_{datetime.now():%Y%m%d_%H%M}.csv", "text/csv")

with tabs[4]:
    st.subheader("Order Detail")
    f1,f2,f3,f4 = st.columns(4)
    with f1: ch = st.multiselect("Channel", sorted(df["Channel"].dropna().unique()), default=list(df["Channel"].dropna().unique()))
    with f2: md = st.multiselect("Model", sorted(df["Model"].dropna().unique()), default=list(df["Model"].dropna().unique()))
    with f3: sl = st.multiselect("SLA", sorted(df["SLA_STATUS"].dropna().unique()), default=list(df["SLA_STATUS"].dropna().unique()))
    with f4: fv = st.multiselect("Verdict", sorted(df["Final_Verdict"].dropna().unique()), default=list(df["Final_Verdict"].dropna().unique()))
    filt = df[df["Channel"].isin(ch)&df["Model"].isin(md)&df["SLA_STATUS"].isin(sl)&df["Final_Verdict"].isin(fv)]
    st.markdown(f"Showing {len(filt):,} / {len(df):,}")
    sc = [c for c in ["CDR ID","Reference ID","Channel","Customer Name","Model","Order_datetime","SLA_STATUS","Breach_delay","Final_Verdict","Rider Name"] if c in filt.columns]
    st.dataframe(filt[sc], use_container_width=True, height=450)

with tabs[5]:
    st.subheader("Downloads (Styled Excel)")
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    date_str = report_dt.strftime("%d-%m-%Y")
    thin = Border(left=Side(style='thin',color='B0B0B0'),right=Side(style='thin',color='B0B0B0'),
                  top=Side(style='thin',color='B0B0B0'),bottom=Side(style='thin',color='B0B0B0'))
    hf = PatternFill("solid", fgColor="1F4E79"); hfont = Font(bold=True, color="FFFFFF", size=11)
    tf = PatternFill("solid", fgColor="C65911"); tfont = Font(bold=True, color="FFFFFF", size=13)
    lo = PatternFill("solid", fgColor="FDEBD0"); lb = PatternFill("solid", fgColor="D6EAF8")
    lg = PatternFill("solid", fgColor="D5F5E3"); lr = PatternFill("solid", fgColor="FADBD8")
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")

    def make_breach():
        wb = Workbook(); ws = wb.active; ws.title = "Breach"
        ws.merge_cells('A1:L1'); ws['A1'] = f"Zippee <-> Supertails Delivered/Breach Report {date_str}"
        ws['A1'].fill=tf; ws['A1'].font=tfont; ws['A1'].alignment=ctr; ws.row_dimensions[1].height=28
        headers = ["Stores","Total Orders","Delivered Orders","15 min Orders","30 min Orders","Other orders",
                   "Total Breach","15 Min Breach","30 Min Breach","Other Breach","Breach %","RiderCount"]
        for c,h in enumerate(headers,1):
            cell=ws.cell(2,c,h); cell.fill=hf; cell.font=hfont; cell.alignment=ctr; cell.border=thin
        for i,store in enumerate(["Kalyan Nagar","Indranagar","Varthur"]):
            row=breach[breach["Store"]==store]
            if len(row)==0: continue
            r=row.iloc[0]
            data=[store,int(r.get("Total Orders",0)),int(r.get("Delivered",0)),int(r.get("15 Min Orders",0)),
                  int(r.get("30 Min Orders",0)),int(r.get("Other Orders",0)),int(r.get("Total Breach",0)),
                  int(r.get("15 Min Breach",0)),int(r.get("30 Min Breach",0)),int(r.get("Other Breach",0)),
                  f"{r.get('Breach %',0)}%",int(r.get("Rider Count",0))]
            for c,v in enumerate(data,1):
                cell=ws.cell(3+i,c,v); cell.border=thin; cell.alignment=ctr
                if c in [7,8,9,10]: cell.fill=lo
                elif c==1: cell.fill=lb; cell.font=Font(bold=True)
        for i,w in enumerate([16,13,15,13,13,12,12,13,13,12,10,11],1):
            ws.column_dimensions[get_column_letter(i)].width=w
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

    def make_pendency():
        wb=Workbook(); ws=wb.active; ws.title="Pendency"
        ws.merge_cells('A1:E1'); ws['A1']=f"Zippee <-> Supertails Pendency Report  {date_str}"
        ws['A1'].fill=PatternFill("solid",fgColor="7030A0"); ws['A1'].font=Font(bold=True,color="FFFFFF",size=13); ws['A1'].alignment=ctr
        for c,h in enumerate(["S.no","Key Parameters","Indra Nagar","Kalyan Nagar","Varthur"],1):
            cell=ws.cell(2,c,h); cell.fill=hf; cell.font=hfont; cell.alignment=ctr; cell.border=thin
        for c,n in [(3,"Bharath"),(4,"Dhiva"),(5,"Syeed")]:
            cell=ws.cell(3,c,n); cell.fill=PatternFill("solid",fgColor="2E75B6"); cell.font=Font(bold=True,color="FFFFFF",size=10); cell.alignment=ctr; cell.border=thin
        ws['A3'].border=thin; ws['B3'].border=thin
        pmap={r["Store"]:r for _,r in pendency.iterrows()}
        rows=[(1,"Total Orders",
               int(pmap.get("Indranagar",{}).get("Total Operational Pending",0))+int(pmap.get("Indranagar",{}).get("Cancelled",0)),
               int(pmap.get("Kalyan Nagar",{}).get("Total Operational Pending",0))+int(pmap.get("Kalyan Nagar",{}).get("Cancelled",0)),
               int(pmap.get("Varthur",{}).get("Total Operational Pending",0))+int(pmap.get("Varthur",{}).get("Cancelled",0))),
              (2,"Pending Order",int(pmap.get("Indranagar",{}).get("Total Operational Pending",0)),
               int(pmap.get("Kalyan Nagar",{}).get("Total Operational Pending",0)),int(pmap.get("Varthur",{}).get("Total Operational Pending",0))),
              (3,"Pendency %",f"{pmap.get('Indranagar',{}).get('0-Attempt %',0)}%",
               f"{pmap.get('Kalyan Nagar',{}).get('0-Attempt %',0)}%",f"{pmap.get('Varthur',{}).get('0-Attempt %',0)}%"),
              (4,"Rider Count",rider_indra,rider_kalyan,rider_varthur),
              (5,"Avg Order/ Rider",
               round(float(pmap.get("Indranagar",{}).get("Total Operational Pending",0))/max(rider_indra,1),1),
               round(float(pmap.get("Kalyan Nagar",{}).get("Total Operational Pending",0))/max(rider_kalyan,1),1),
               round(float(pmap.get("Varthur",{}).get("Total Operational Pending",0))/max(rider_varthur,1),1))]
        for i,(sno,param,v1,v2,v3) in enumerate(rows):
            r=4+i
            for c,v in enumerate([sno,param,v1,v2,v3],1):
                cell=ws.cell(r,c,v); cell.border=thin; cell.alignment=ctr
            ws.cell(r,2).alignment=left
            if sno in [3,5]:
                for c in range(3,6): ws.cell(r,c).fill=PatternFill("solid",fgColor="E8DAEF")
        sr=10
        ws.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=2)
        ws.cell(sr,1,"Sub Parameters").fill=hf; ws.cell(sr,1).font=hfont; ws.cell(sr,1).alignment=ctr; ws.cell(sr,1).border=thin; ws.cell(sr,2).border=thin
        for c,n in [(3,"Indra Nagar"),(4,"Kalyan Nagar"),(5,"Varthur")]:
            cell=ws.cell(sr,c,n); cell.fill=hf; cell.font=hfont; cell.alignment=ctr; cell.border=thin
        subs=[(6,"Attempted Delivery","Attempted Delivery"),(7,"Not Dispatched","Not Dispatched"),
              (8,"0 Attempted","0 Attempt"),(9,"Picked Up","Picked Up"),(10,"Intransit","Intransit"),
              (11,"Cancelled","Cancelled"),(12,"Courier Orders",None)]
        for i,(sno,label,key) in enumerate(subs):
            r=sr+1+i
            ws.cell(r,1,sno).border=thin; ws.cell(r,2,label).border=thin
            for c in range(1,6): ws.cell(r,c).alignment=ctr; ws.cell(r,c).border=thin
            ws.cell(r,2).alignment=left
            if key:
                ws.cell(r,3,int(pmap.get("Indranagar",{}).get(key,0)))
                ws.cell(r,4,int(pmap.get("Kalyan Nagar",{}).get(key,0)))
                ws.cell(r,5,int(pmap.get("Varthur",{}).get(key,0)))
            else:
                ws.cell(r,3,0); ws.cell(r,4,0); ws.cell(r,5,0)
        for col,w in [('A',8),('B',22),('C',14),('D',14),('E',12)]: ws.column_dimensions[col].width=w
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

    def make_eod():
        wb=Workbook(); ws=wb.active; ws.title="EOD"
        ws.merge_cells('A1:E1'); ws['A1']=f"Zippee <-> Supertails EOD Report  {date_str}"
        ws['A1'].fill=PatternFill("solid",fgColor="1B4F72"); ws['A1'].font=Font(bold=True,color="FFFFFF",size=13); ws['A1'].alignment=ctr
        for c,h in enumerate(["Stores","Indra Nagar","Kalyan Nagar","Varthur","Total"],1):
            cell=ws.cell(2,c,h); cell.fill=hf; cell.font=hfont; cell.alignment=ctr; cell.border=thin
        bmap={r["Store"]:r for _,r in breach.iterrows()}
        pmap={r["Store"]:r for _,r in pendency.iterrows()}
        def g(s,k,d=0):
            if s in bmap and k in bmap[s] and pd.notna(bmap[s][k]): return bmap[s][k]
            if s in pmap and k in pmap[s] and pd.notna(pmap[s][k]): return pmap[s][k]
            return d
        metrics=[("Total Orders","Total Orders"),("Attempted Delivery","Attempted Delivery"),
                 ("Not Dispatched","Not Dispatched"),("Not Dispatched ( 0 Attempts )","0 Attempt"),
                 ("Pickedup","Picked Up"),("Intransit","Intransit"),("Cancelled","Cancelled"),
                 ("Total Delivered","Delivered"),("15 Mins Orders","15 Min Orders"),
                 ("30 Mins Orders","30 Min Orders"),("Other Orders","Other Orders"),
                 ("On Time Delivered",None),("Total Breach","Total Breach"),
                 ("15 Mins Breach","15 Min Breach"),("30 Mins Breach","30 Min Breach"),
                 ("Other Breach","Other Breach"),("Breach %","Breach %"),
                 ("Planned Rider",None),("Active Rider",None),("Efficiency of Delivered Orders",None)]
        for idx,(label,key) in enumerate(metrics):
            r=3+idx
            ws.cell(r,1,label).border=thin; ws.cell(r,1).alignment=left; ws.cell(r,1).font=Font(bold=True,size=10)
            vals=[]
            for store in ["Indranagar","Kalyan Nagar","Varthur"]:
                if label=="On Time Delivered": val=int(g(store,"Delivered",0))-int(g(store,"Total Breach",0))
                elif label=="Planned Rider": val=rider_map.get(store,0)
                elif label=="Active Rider": val=rider_map.get(store,0)
                elif label=="Efficiency of Delivered Orders":
                    val=round(float(g(store,"Delivered",0))/max(rider_map.get(store,1),1),2)
                elif label=="Breach %": val=f"{g(store,'Breach %',0)}%"
                else:
                    val=g(store,key,0)
                    try:
                        if isinstance(val,(float,np.floating)): val=round(val,1) if not float(val).is_integer() else int(val)
                    except: pass
                vals.append(val)
                cell=ws.cell(r,2+["Indranagar","Kalyan Nagar","Varthur"].index(store),val)
                cell.border=thin; cell.alignment=ctr
            if label in ["Breach %","Efficiency of Delivered Orders"]:
                try:
                    nums=[float(str(v).replace('%','')) for v in vals]
                    total=round(sum(nums)/3,2)
                    if label=="Breach %": total=f"{total}%"
                except: total=""
            elif label in ["Planned Rider","Active Rider"]: total=sum(int(v) for v in vals)
            else:
                try: total=sum(int(v) if not isinstance(v,str) else 0 for v in vals)
                except: total=""
            cell=ws.cell(r,5,total); cell.border=thin; cell.alignment=ctr; cell.font=Font(bold=True)
            if label in ["Total Breach","15 Mins Breach","30 Mins Breach","Other Breach","Breach %"]:
                for c in range(1,6): ws.cell(r,c).fill=lr
            elif label in ["Total Delivered","On Time Delivered"]:
                for c in range(1,6): ws.cell(r,c).fill=lg
            elif label in ["Not Dispatched","Not Dispatched ( 0 Attempts )"]:
                for c in range(1,6): ws.cell(r,c).fill=lo
        for col,w in [('A',32),('B',13),('C',14),('D',12),('E',10)]: ws.column_dimensions[col].width=w
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

    a,b,c = st.columns(3)
    with a:
        st.markdown("**EOD Excel**")
        st.download_button("📥 EOD Excel", make_eod(), f"ST_EOD_{ts}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with b:
        st.markdown("**Pendency Excel**")
        st.download_button("📥 Pendency Excel", make_pendency(), f"ST_Pendency_{ts}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c:
        st.markdown("**Breach Excel**")
        st.download_button("📥 Breach Excel", make_breach(), f"ST_Breach_{ts}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.markdown("---")
    d,e,f = st.columns(3)
    with d: st.download_button("EOD CSV", eod.to_csv(index=False).encode(), f"ST_EOD_{ts}.csv", "text/csv", use_container_width=True)
    with e: st.download_button("Pendency CSV", pendency.to_csv(index=False).encode(), f"ST_Pendency_{ts}.csv", "text/csv", use_container_width=True)
    with f: st.download_button("Breach CSV", breach.to_csv(index=False).encode(), f"ST_Breach_{ts}.csv", "text/csv", use_container_width=True)
    st.download_button("Full Processed CSV", df.to_csv(index=False).encode(), f"ST_Full_{ts}.csv", "text/csv", use_container_width=True)
    st.info("Final Verdict = Yes only if delay ≥ 2 min (matches Excel)")

st.caption(f"ST Dashboard v2 · {len(df):,} orders · Final Verdict ≥ 2 min")
