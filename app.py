
import streamlit as st
import pandas as pd
import sqlite3, hashlib, json, io, os, re, shutil
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
import plotly.express as px
import plotly.graph_objects as go

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "controle_gestion.sqlite3"
EXAMPLES = APP_DIR / "imports_exemples"
DEMO_MODE = True

st.set_page_config(page_title="KMC – Contrôle de gestion territorial", page_icon="🏛️", layout="wide")

st.markdown("""
<style>
:root { --navy:#142b4a; --green:#6b9b55; }
[data-testid="stAppViewContainer"] { background:#f5f7fa; }
[data-testid="stSidebar"] { background:#142b4a; }
[data-testid="stSidebar"] * { color:#f4f7fb !important; }
.kmc-header {background:linear-gradient(115deg,#142b4a,#244b70);padding:22px 26px;border-radius:15px;color:white;margin-bottom:18px}
.kmc-header h1 {color:white;margin:0;font-size:27px}
.kmc-header p {color:#dbe7f1;margin:6px 0 0 0}
.kpi {background:white;border:1px solid #e1e7ee;border-radius:13px;padding:15px 17px;min-height:112px}
.kpi .label {font-size:12px;color:#627184;font-weight:600}
.kpi .value {font-size:25px;color:#142b4a;font-weight:750;margin-top:8px}
.kpi .delta {font-size:12px;color:#637386;margin-top:3px}
.small-note {font-size:12px;color:#6c7888}
div[data-testid="stMetric"] {background:white;border:1px solid #e1e7ee;padding:14px;border-radius:12px}
</style>
""", unsafe_allow_html=True)

def db():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS mandats (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      fingerprint TEXT UNIQUE, exercice_fichier INTEGER, date_mandat TEXT,
      budget TEXT DEFAULT 'Commune', type_flux TEXT DEFAULT 'Dépenses',
      service TEXT, code_service TEXT, antenne TEXT, fonction TEXT,
      article TEXT, libelle_article TEXT, compte TEXT, fournisseur TEXT,
      bordereau TEXT, engagement TEXT, montant_engagement REAL,
      montant_ht REAL, tva REAL, montant_ttc REAL, nom_fichier TEXT,
      date_import TEXT
    );
    CREATE TABLE IF NOT EXISTS imports (
      id INTEGER PRIMARY KEY AUTOINCREMENT, nom_fichier TEXT, exercice INTEGER,
      type_flux TEXT, date_import TEXT, lignes_lues INTEGER, lignes_importees INTEGER,
      doublons INTEGER, rejetees INTEGER, montant_ttc REAL, commentaire TEXT
    );
    CREATE TABLE IF NOT EXISTS actions (
      id INTEGER PRIMARY KEY AUTOINCREMENT, reference TEXT, axe TEXT, description TEXT,
      service TEXT, article TEXT, fournisseur TEXT, responsable TEXT, echeance TEXT,
      economie_prevue REAL DEFAULT 0, cout_induit REAL DEFAULT 0, economie_realisee REAL DEFAULT 0,
      recurrente INTEGER DEFAULT 1, statut TEXT DEFAULT 'À étudier', preuve TEXT, commentaire TEXT
    );
    CREATE TABLE IF NOT EXISTS kpis (
      id INTEGER PRIMARY KEY AUTOINCREMENT, nom TEXT, definition TEXT, formule TEXT,
      unite TEXT, source TEXT, frequence TEXT, seuil_alerte TEXT, actif INTEGER DEFAULT 1
    );
    """)
    n = con.execute("SELECT COUNT(*) FROM kpis").fetchone()[0]
    if n == 0:
        defaults = [
          ("Taux de consommation des crédits","Part des crédits ouverts déjà mandatés","Mandaté / crédits ouverts","%","Budget + grand livre","Mensuelle","Alerte si > 90 % avant fin d'exercice"),
          ("Taux d'engagement","Part des crédits ouverts déjà engagés","Engagé / crédits ouverts","%","Budget + engagements","Mensuelle","Alerte si > 95 %"),
          ("Évolution des dépenses","Évolution à période identique","(N - N-1) / N-1","%","Grands livres","Mensuelle","Alerte si hausse > 5 % et > 1 000 €"),
          ("Coût par repas","Dépenses alimentaires divisées par repas servis","Dépenses alimentation / repas","€/repas","Grand livre + activité","Mensuelle","À paramétrer"),
          ("Énergie par m²","Consommation énergétique par surface","kWh / m²","kWh/m²","Factures + surfaces","Mensuelle","À paramétrer"),
          ("Économie nette","Économie réalisée diminuée du coût induit","Économie réalisée - coût induit","€","Plan d'actions","Mensuelle","Objectif non acquis"),
        ]
        con.executemany("INSERT INTO kpis(nom,definition,formule,unite,source,frequence,seuil_alerte) VALUES(?,?,?,?,?,?,?)", defaults)
    con.commit()
    con.close()

init_db()

# Jeu de démonstration entièrement fictif : aucune donnée communale réelle.
def seed_demo_data():
    con = db()
    existing = con.execute("SELECT COUNT(*) FROM mandats").fetchone()[0]
    if existing == 0:
        import random
        random.seed(20261119)
        services = ["Bâtiments", "Espaces verts", "Voirie", "Éducation", "Culture", "Administration", "Sports"]
        articles = [(60612, "Énergie - électricité"), (60622, "Carburants"), (60631, "Fournitures d'entretien"),
                    (611, "Contrats de prestations"), (615221, "Entretien bâtiments"), (615231, "Entretien voirie"),
                    (623, "Publicité et publications"), (626, "Frais de télécommunications"), (6283, "Nettoyage des locaux")]
        suppliers = ["Fournisseur Démo A", "Services Exemple B", "Énergie Fictive C", "Maintenance Démo D", "Papeterie Exemple E"]
        rows=[]
        for year in (2025, 2026):
            for i in range(110):
                month = random.randint(1, 12 if year == 2025 else 11)
                day = random.randint(1, 25)
                art, label = random.choice(articles)
                ttc = round(random.uniform(180, 7800), 2)
                d={"exercice_fichier":year,"date_mandat":date(year,month,day).isoformat(),"budget":"Commune Démo",
                   "type_flux":"Dépenses","service":random.choice(services),"code_service":None,"antenne":None,
                   "fonction":None,"article":str(art),"libelle_article":label,"compte":None,
                   "fournisseur":random.choice(suppliers),"bordereau":f"BD-{year}-{i+1:03d}",
                   "engagement":f"ENG-{year}-{i+1:04d}","montant_engagement":ttc,"montant_ht":round(ttc/1.2,2),
                   "tva":round(ttc-ttc/1.2,2),"montant_ttc":ttc,"nom_fichier":"JEU_DEMO_FICTIF.xlsx"}
                d["fingerprint"]=fingerprint(d)
                rows.append(d)
        import_rows(rows,"JEU_DEMO_FICTIF.xlsx",2025,"Dépenses")
        # Corrige l'exercice pour les écritures 2026 et ajoute une ligne de journal distincte.
        for d in rows:
            if d["exercice_fichier"] == 2026:
                try:
                    con.execute("UPDATE mandats SET exercice_fichier=2026 WHERE fingerprint=?",(d["fingerprint"],))
                except Exception:
                    pass
        con.commit()
    con.close()


def money(v):
    try: return f"{float(v):,.0f} €".replace(",", " ").replace(".", ",")
    except: return "—"

def parse_date(v):
    if isinstance(v, datetime): return v.date().isoformat()
    if isinstance(v, date): return v.isoformat()
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    for dayfirst in (True, False):
        try: return pd.to_datetime(s, dayfirst=dayfirst, errors="raise").date().isoformat()
        except: pass
    return None

def num(v):
    if isinstance(v, (int,float)): return float(v)
    if v is None: return 0.0
    s = str(v).replace("\u202f","").replace("\xa0","").replace(" ","").replace(",",".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try: return float(s) if s not in ("","-",".") else 0.0
    except: return 0.0

def fingerprint(d):
    keys = ["exercice_fichier","date_mandat","service","article","compte","fournisseur","bordereau","engagement","montant_ht","tva","montant_ttc","type_flux"]
    raw = "|".join(str(d.get(k,"")) for k in keys)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def parse_ciril_xlsx(file_obj, filename, exercice, budget, type_flux):
    wb = load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    current = {"service":None,"code_service":None,"antenne":None,"fonction":None,"article":None,"libelle_article":None,"compte":None}
    parsed=[]; rows_read=0; rejected=0
    for row in ws.iter_rows(values_only=True):
        cells=list(row)+[None]*max(0,18-len(row))
        a0=str(cells[0]).strip() if cells[0] is not None else ""
        if a0.lower().startswith("gestionnaire"):
            current["code_service"]=str(cells[2]).strip() if cells[2] is not None else None
            current["service"]=str(cells[3]).strip() if cells[3] is not None else (str(cells[1]).strip() if cells[1] else None)
            continue
        if a0.lower().startswith("fonction"):
            current["fonction"]=str(cells[1]).strip() if cells[1] is not None else None
            current["antenne"]=str(cells[2]).strip() if cells[2] is not None else None
            continue
        if a0.lower().startswith("nature"):
            nature=str(cells[1]).strip() if cells[1] is not None else ""
            m=re.match(r"\s*(\d{3,5})\s*(.*)",nature)
            current["article"]=m.group(1) if m else None
            current["libelle_article"]=m.group(2).strip() if m else nature
            continue
        if a0.lower().startswith("compte"):
            current["compte"]=a0.split(":",1)[1].strip() if ":" in a0 else (str(cells[1]).strip() if cells[1] else None)
            continue

        # Structure connue des exports fournis : date en F, tiers en D, HT/TVA/TTC en H/I/J.
        dt=parse_date(cells[5])
        supplier=cells[3]
        ttc=cells[9]
        is_mandate = dt is not None and supplier not in (None,"") and isinstance(ttc,(int,float))
        # Repli : recherche d'une date et d'un montant numérique dans la ligne.
        if not is_mandate and dt is not None:
            numeric=[x for x in cells if isinstance(x,(int,float))]
            if supplier not in (None,"") and numeric:
                ttc=numeric[-1]; is_mandate=True
        if is_mandate:
            rows_read += 1
            d={
              "exercice_fichier":int(exercice),"date_mandat":dt,"budget":budget,"type_flux":type_flux,
              "service":current["service"],"code_service":current["code_service"],"antenne":current["antenne"],
              "fonction":current["fonction"],"article":current["article"],"libelle_article":current["libelle_article"],
              "compte":current["compte"],"fournisseur":str(supplier).strip(),"bordereau":str(cells[4] or ""),
              "engagement":str(cells[0] or ""),"montant_engagement":num(cells[1]),
              "montant_ht":num(cells[7]),"tva":num(cells[8]),"montant_ttc":num(ttc),
              "nom_fichier":filename
            }
            d["fingerprint"]=fingerprint(d)
            parsed.append(d)
    wb.close()
    return parsed, rows_read, rejected

def import_rows(rows, filename, exercice, type_flux):
    con=db(); inserted=0; duplicates=0; total=0.0
    for d in rows:
        total += d["montant_ttc"]
        try:
            con.execute("""INSERT INTO mandats(fingerprint,exercice_fichier,date_mandat,budget,type_flux,service,code_service,antenne,
            fonction,article,libelle_article,compte,fournisseur,bordereau,engagement,montant_engagement,montant_ht,tva,montant_ttc,nom_fichier,date_import)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (d["fingerprint"],d["exercice_fichier"],d["date_mandat"],d["budget"],d["type_flux"],d["service"],d["code_service"],
             d["antenne"],d["fonction"],d["article"],d["libelle_article"],d["compte"],d["fournisseur"],d["bordereau"],
             d["engagement"],d["montant_engagement"],d["montant_ht"],d["tva"],d["montant_ttc"],d["nom_fichier"],datetime.now().isoformat(timespec="seconds")))
            inserted += 1
        except sqlite3.IntegrityError:
            duplicates += 1
    con.execute("""INSERT INTO imports(nom_fichier,exercice,type_flux,date_import,lignes_lues,lignes_importees,doublons,rejetees,montant_ttc,commentaire)
    VALUES(?,?,?,?,?,?,?,?,?,?)""",(filename,exercice,type_flux,datetime.now().isoformat(timespec="seconds"),len(rows),inserted,duplicates,0,total,
    "Import des mandats ; dédoublonnage par empreinte de ligne."))
    con.commit(); con.close()
    return inserted,duplicates,total

# Initialisation des données fictives après définition de toutes les fonctions nécessaires.
seed_demo_data()

def load_data():
    con=db()
    df=pd.read_sql_query("SELECT * FROM mandats",con,parse_dates=["date_mandat"])
    con.close()
    if not df.empty:
        df["date_mandat"]=pd.to_datetime(df["date_mandat"],errors="coerce")
        df["annee"]=df["date_mandat"].dt.year
        df["mois"]=df["date_mandat"].dt.to_period("M").astype(str)
        df["article"]=df["article"].fillna("Non renseigné")
        df["service"]=df["service"].fillna("Non renseigné")
        df["fournisseur"]=df["fournisseur"].fillna("Non renseigné")
    return df

def export_excel(df):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as writer:
        df.to_excel(writer,index=False,sheet_name="Données")
    bio.seek(0)
    return bio

# Sidebar
with st.sidebar:
    st.markdown("## 🏛️ KMC")
    st.caption("Contrôle de gestion territorial")
    page=st.radio("Navigation",["Vue générale","Importer Ciril","Analyse des dépenses","Plan d'économies","Indicateurs","Historique des imports","Paramètres & sauvegarde"])
    st.markdown("---")
    st.caption("Saint-Aubin-lès-Elbeuf · M57")
    st.caption("DÉMONSTRATION – données entièrement fictives")

st.markdown('<div class="kmc-header"><h1>KMC – Contrôle de gestion territorial</h1><p>COMMUNE DÉMO · Environnement de démonstration avec données fictives</p></div>',unsafe_allow_html=True)
st.warning("ENVIRONNEMENT DE DÉMONSTRATION : toutes les écritures et tous les fournisseurs affichés sont fictifs. Ne dépose aucun grand livre réel, aucune donnée personnelle ni document de la collectivité.")

df=load_data()

def filtered_data(df):
    if df.empty: return df
    years=sorted([int(x) for x in df["annee"].dropna().unique()])
    c1,c2,c3=st.columns(3)
    with c1: year=st.selectbox("Exercice",["Tous"]+years,index=len(years))
    with c2: budget=st.selectbox("Budget",["Tous"]+sorted(df["budget"].dropna().unique().tolist()))
    with c3: metric=st.selectbox("Montant de référence",["TTC","HT"],index=0)
    out=df.copy()
    if year!="Tous": out=out[out["annee"]==year]
    if budget!="Tous": out=out[out["budget"]==budget]
    amount="montant_ttc" if metric=="TTC" else "montant_ht"
    return out,amount,year,budget

if page=="Vue générale":
    st.subheader("Tableau de bord général")
    if df.empty:
        st.info("Aucune donnée importée. Commence par déposer un grand livre Ciril dans « Importer Ciril ». Les indicateurs seront calculés à partir des écritures importées.")
        st.markdown("**Fonctionnalités prévues dans cette première version :** import des grands livres, comparaison annuelle, analyse par article/service/fournisseur, plan d'économies et sauvegarde.")
    else:
        filtered=filtered_data(df)
        if len(filtered)==4:
            d,amount,year,budget=filtered
            total=d[amount].sum()
            nmand=len(d)
            suppliers=d["fournisseur"].nunique()
            services=d["service"].nunique()
            c=st.columns(4)
            c[0].metric("Dépenses mandatées",money(total))
            c[1].metric("Mandats importés",f"{nmand:,}".replace(","," "))
            c[2].metric("Fournisseurs",suppliers)
            c[3].metric("Services",services)
            st.caption(f"Périmètre : {year if year!='Tous' else 'tous exercices'} · {budget if budget!='Tous' else 'tous budgets'} · montants {amount.upper()}. Les crédits votés et engagements ne sont pas calculés tant que leurs extractions ne sont pas importées.")
            a,b=st.columns(2)
            with a:
                byart=d.groupby("article",dropna=False)[amount].sum().sort_values(ascending=False).head(12).reset_index()
                st.plotly_chart(px.bar(byart,x=amount,y="article",orientation="h",title="Dépenses par article",labels={amount:"Montant","article":"Article"}),use_container_width=True)
            with b:
                bymonth=d.groupby("mois")[amount].sum().reset_index()
                st.plotly_chart(px.line(bymonth,x="mois",y=amount,markers=True,title="Évolution mensuelle"),use_container_width=True)
            a,b=st.columns(2)
            with a:
                byservice=d.groupby("service")[amount].sum().sort_values(ascending=False).head(10).reset_index()
                st.plotly_chart(px.bar(byservice,x=amount,y="service",orientation="h",title="Dépenses par service"),use_container_width=True)
            with b:
                bysup=d.groupby("fournisseur")[amount].sum().sort_values(ascending=False).head(10).reset_index()
                st.plotly_chart(px.bar(bysup,x=amount,y="fournisseur",orientation="h",title="Principaux fournisseurs"),use_container_width=True)
            st.subheader("Alertes de gestion")
            # Compare same period up to latest date for the most recent two years.
            years=sorted(d["annee"].dropna().unique())
            if len(years)>=2:
                y0,y1=years[-2],years[-1]
                maxdate=d[d["annee"]==y1]["date_mandat"].max()
                end_prev=maxdate.replace(year=y0)
                start0=maxdate.replace(year=int(y0),month=1,day=1)
                start1=maxdate.replace(year=int(y1),month=1,day=1)
                end_prev=maxdate.replace(year=int(y0))
                p0=d[(d["annee"]==y0)&(d["date_mandat"]>=start0)&(d["date_mandat"]<=end_prev)]
                p1=d[(d["annee"]==y1)&(d["date_mandat"]>=start1)&(d["date_mandat"]<=maxdate)]
                comp=pd.concat([p0.groupby("article")[amount].sum().rename(str(y0)),p1.groupby("article")[amount].sum().rename(str(y1))],axis=1).fillna(0)
                comp["Écart €"]=comp[str(y1)]-comp[str(y0)]
                comp["Écart %"]=comp.apply(lambda r:r["Écart €"]/r[str(y0)] if r[str(y0)] else None,axis=1)
                alerts=comp[(comp["Écart €"]>1000)&(comp["Écart %"]>0.05)].sort_values("Écart €",ascending=False).head(10)
                if alerts.empty: st.success("Aucune alerte automatique ne dépasse les seuils par défaut.")
                else: st.dataframe(alerts.style.format({str(y0):"{:,.0f} €",str(y1):"{:,.0f} €","Écart €":"{:,.0f} €","Écart %":"{:.1%}"}),use_container_width=True)
            else: st.info("Importe au moins deux exercices pour activer les comparaisons N/N-1.")
            st.download_button("Exporter les données filtrées en Excel",data=export_excel(d),file_name="KMC_donnees_filtrees.xlsx")

elif page=="Importer Ciril":
    st.subheader("Importer les extractions Ciril")
    st.write("Dépose ici un grand livre de dépenses ou de recettes exporté depuis CIVIL Finances. Les lignes déjà importées sont détectées et ignorées.")
    st.info("Les données affichées sont un jeu de démonstration fictif. Pour cette version publique, ne téléverse aucun fichier réel de la collectivité.")
    st.markdown("---")
    upload=st.file_uploader("Import Excel désactivé dans cette démonstration",type=["xlsx","xlsm"],disabled=True)
    c1,c2,c3=st.columns(3)
    with c1: ex=st.number_input("Exercice du fichier",min_value=2000,max_value=2100,value=2026,step=1)
    with c2: bud=st.selectbox("Budget concerné",["Commune","CCAS","Budget annexe"])
    with c3: flux=st.selectbox("Type de flux",["Dépenses","Recettes"])
    if upload:
        if st.button("Analyser le fichier avant import",type="primary"):
            try:
                rows,n,rej=parse_ciril_xlsx(upload,upload.name,int(ex),bud,flux)
                st.session_state["preview_rows"]=rows
                st.session_state["preview_meta"]=(upload.name,int(ex),flux,n,rej)
            except Exception as e: st.error(f"Lecture impossible : {e}")
    if "preview_rows" in st.session_state:
        rows=st.session_state["preview_rows"]; meta=st.session_state["preview_meta"]
        st.subheader("Aperçu avant validation")
        st.write(f"Fichier : **{meta[0]}** · {meta[3]} lignes de mandatement reconnues.")
        if rows:
            st.dataframe(pd.DataFrame(rows).head(30),use_container_width=True)
            st.warning("Vérifie l'aperçu et les montants avant de confirmer l'import. Les écritures seront ajoutées à la base sans écraser les imports précédents.")
            if st.button("Confirmer l'import",type="primary"):
                ins,dup,total=import_rows(rows,meta[0],meta[1],meta[2])
                st.success(f"Import terminé : {ins} nouvelles lignes, {dup} doublons ignorés. Montant TTC des lignes lues : {money(total)}.")
                del st.session_state["preview_rows"]; del st.session_state["preview_meta"]
                st.rerun()
        else: st.error("Aucune ligne de mandatement reconnue. Vérifie le format de l'export.")

elif page=="Analyse des dépenses":
    st.subheader("Analyse détaillée des dépenses")
    if df.empty: st.info("Importe d'abord un grand livre.")
    else:
        d,amount,year,budget=filtered_data(df)
        tabs=st.tabs(["Articles","Services","Fournisseurs","Évolution mensuelle","Écritures"])
        with tabs[0]:
            x=d.groupby(["article","libelle_article"],dropna=False)[amount].sum().reset_index().sort_values(amount,ascending=False)
            st.dataframe(x,use_container_width=True)
        with tabs[1]:
            x=d.groupby("service")[amount].sum().reset_index().sort_values(amount,ascending=False)
            st.dataframe(x,use_container_width=True)
        with tabs[2]:
            x=d.groupby("fournisseur")[amount].sum().reset_index().sort_values(amount,ascending=False)
            st.dataframe(x,use_container_width=True)
        with tabs[3]:
            x=d.groupby("mois")[amount].sum().reset_index()
            st.plotly_chart(px.bar(x,x="mois",y=amount,title="Dépenses mensuelles"),use_container_width=True)
        with tabs[4]:
            c1,c2=st.columns(2)
            with c1: q=st.text_input("Rechercher un fournisseur ou une référence")
            with c2: art=st.selectbox("Article",["Tous"]+sorted(d["article"].astype(str).unique()))
            x=d.copy()
            if q: x=x[x.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)]
            if art!="Tous": x=x[x["article"].astype(str)==art]
            st.dataframe(x.sort_values("date_mandat",ascending=False),use_container_width=True,height=500)
            st.download_button("Exporter les lignes affichées",data=export_excel(x),file_name="KMC_ecritures.xlsx")

elif page=="Plan d'économies":
    st.subheader("Registre des pistes d'économies")
    st.caption("La cible de 175 000 € est un objectif de travail, non une économie acquise.")
    con=db()
    with st.form("new_action"):
        st.markdown("### Ajouter une piste")
        c1,c2=st.columns(2)
        with c1:
            axe=st.text_input("Axe / intitulé",placeholder="Ex. Audit des contrats de maintenance")
            desc=st.text_area("Constat et action envisagée")
            service=st.text_input("Service concerné")
            article=st.text_input("Article comptable")
            fournisseur=st.text_input("Fournisseur")
        with c2:
            responsable=st.text_input("Responsable de l'action")
            echeance=st.date_input("Échéance",value=date.today())
            prev=st.number_input("Économie annuelle prévue (€)",min_value=0.0,step=500.0)
            cout=st.number_input("Coût induit (€)",min_value=0.0,step=100.0)
            real=st.number_input("Économie réalisée (€)",min_value=0.0,step=100.0)
            rec=st.checkbox("Économie récurrente",value=True)
            statut=st.selectbox("Statut",["À étudier","Validée","En cours","Réalisée","Abandonnée"])
        preuve=st.text_input("Preuve / document justificatif")
        commentaire=st.text_area("Commentaire")
        submit=st.form_submit_button("Enregistrer la piste")
        if submit:
            ref=f"KMC-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            con.execute("""INSERT INTO actions(reference,axe,description,service,article,fournisseur,responsable,echeance,economie_prevue,cout_induit,economie_realisee,recurrente,statut,preuve,commentaire)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(ref,axe,desc,service,article,fournisseur,responsable,echeance.isoformat(),prev,cout,real,int(rec),statut,preuve,commentaire))
            con.commit(); st.success("Piste enregistrée.")
    actions=pd.read_sql_query("SELECT * FROM actions ORDER BY id DESC",con); con.close()
    if not actions.empty:
        actions["economie_nette_prevue"]=actions["economie_prevue"]-actions["cout_induit"]
        actions["economie_nette_realisee"]=actions["economie_realisee"]-actions["cout_induit"]
        c=st.columns(3)
        c[0].metric("Économies prévues",money(actions.loc[actions.statut!="Abandonnée","economie_prevue"].sum()))
        c[1].metric("Économies réalisées",money(actions.loc[actions.statut=="Réalisée","economie_realisee"].sum()))
        c[2].metric("Économies nettes réalisées",money(actions.loc[actions.statut=="Réalisée","economie_nette_realisee"].sum()))
        st.dataframe(actions,use_container_width=True)
        st.download_button("Exporter le plan d'économies",data=export_excel(actions),file_name="KMC_plan_economies.xlsx")
    else: st.info("Aucune piste enregistrée. Ajoute les premières actions issues de tes audits.")

elif page=="Indicateurs":
    st.subheader("Catalogue des indicateurs")
    con=db()
    k=pd.read_sql_query("SELECT * FROM kpis ORDER BY id",con)
    st.dataframe(k,use_container_width=True)
    with st.form("new_kpi"):
        st.markdown("### Ajouter un indicateur")
        nom=st.text_input("Nom")
        definition=st.text_area("Définition")
        formule=st.text_input("Formule")
        unite=st.text_input("Unité")
        source=st.text_input("Source des données")
        freq=st.selectbox("Fréquence",["Mensuelle","Trimestrielle","Annuelle"])
        seuil=st.text_input("Seuil d'alerte")
        if st.form_submit_button("Ajouter l'indicateur") and nom:
            con.execute("INSERT INTO kpis(nom,definition,formule,unite,source,frequence,seuil_alerte) VALUES(?,?,?,?,?,?,?)",(nom,definition,formule,unite,source,freq,seuil))
            con.commit(); st.success("Indicateur ajouté.")
    con.close()
    st.info("Les indicateurs dépendant de données d'activité (repas, fréquentation, surfaces, kilomètres, kWh) restent non calculés tant que ces données ne sont pas importées.")

elif page=="Historique des imports":
    st.subheader("Journal des imports")
    con=db()
    imp=pd.read_sql_query("SELECT * FROM imports ORDER BY id DESC",con)
    con.close()
    if imp.empty: st.info("Aucun import enregistré.")
    else:
        st.dataframe(imp,use_container_width=True)
        st.download_button("Exporter le journal",data=export_excel(imp),file_name="KMC_journal_imports.xlsx")
    if not df.empty:
        st.metric("Nombre total de lignes en base",f"{len(df):,}".replace(","," "))
        st.caption("La clé de dédoublonnage empêche l'import répété d'une même ligne selon les champs de l'écriture.")

elif page=="Paramètres & sauvegarde":
    st.subheader("Paramètres et sauvegarde")
    st.write("Version de démonstration hébergée. Les données fictives peuvent être réinitialisées par la plateforme ; ne pas utiliser comme stockage professionnel.")
    st.info("Les extractions Ciril ne sont pas modifiées. Cette application conserve une copie analytique des données importées.")
    con=db()
    c1,c2=st.columns(2)
    with c1:
        st.metric("Écritures importées",len(df))
        st.metric("Budgets enregistrés",df["budget"].nunique() if not df.empty else 0)
    with c2:
        st.metric("Exercices",df["annee"].nunique() if not df.empty else 0)
        st.metric("Taille de la base (Ko)",round(DB_PATH.stat().st_size/1024,1) if DB_PATH.exists() else 0)
    con.close()
    if DB_PATH.exists():
        st.download_button("Télécharger une sauvegarde de la base",data=DB_PATH.read_bytes(),file_name="KMC_sauvegarde.sqlite3",mime="application/octet-stream")
    if not df.empty:
        st.download_button("Exporter toutes les écritures en Excel",data=export_excel(df),file_name="KMC_export_complet.xlsx")
    st.warning("Démonstration uniquement : stockage non garanti, réinitialisation possible et aucune donnée réelle ne doit être importée.")
