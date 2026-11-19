# ============================================================
# GESTION D'INVENTAIRE UNOPS - VERSION FINALE AVEC COLONNE DATE
# ============================================================

import os
import pandas as pd
import gradio as gr
import plotly.express as px
from datetime import datetime

# 1. Configuration du dossier local
DOSSIER = "donnees_inventaire"
os.makedirs(DOSSIER, exist_ok=True)
FICHIER = f"{DOSSIER}/Base_Donnees_Inventaire.xlsx"

# 2. Colonnes des bases de données
COL_ART = ["ID", "Date", "Nom Article", "Nomenclature", "Catégorie", "Projet", "Unités/Pack", "Entrepôt par défaut", "Responsable", "Commentaires"]
COL_ENT = ["ID", "Nom Entrepôt"]
COL_MVT = ["Date", "Type", "ID Article", "Nomenclature", "Nom Article", "Quantité", "Entrepôt", 
           "Entrepôt de destination", "Quantité avant", "Quantité après", "N° Document", "Projet", 
           "Site bénéficiaire", "Responsable", "Observations"]

# 3. Catalogues de Nomenclatures (PARN, PPSB, HOPITAUX)
RAW_HOPITAUX = """
OTF-07	Tabouret
OTF-14	Chariot à instrument
OTF-18	Chariot à pansement
OTF-23	Table d’accouchement
OTF-27	Table de chevet
OTF-28	Lit avec matelas à housse lavable
OTF-33	Lit médicalisé
OTI-06	Boite césarienne
OTI-08	Boite hystérectomie
OTI-09	Boite laparotomie
OTI-11	Boite petite chirurgie
OTI-12	Boite à pansement
OTI-13	Boite à compresse GM
OTI-17	Boite à coton
OTI-19	Kit spéculum vaginal
OTI-20	Spatule
OTI-45	Boite d’accouchement simple
OTI-49	Spatule / forceps
OTM-01	Appareil d’anesthésie complet
OTM-02	Aspirateur chirurgical mobile
OTM-03	Bistouri électrique
OTM-04	Scialytique plafonnier
OTM-05	Table d’opération universelle
OTM-10	Autoclave
OTM-36	Incubateur (couveuse)
OTS-15	Kit d’oxygène
OTS-16	Balance pèse bébé
OTS-21	Lampe d’examen mobile
OTS-22	Sthétoscope de Pinard
OTS-24	Stéthoscope
OTS-25	Tensiomètre à brassard manuel
OTS-26	Ventouse obstétrical manuel
OTS-30	Défibrillateur
OTS-31	AMBU complet adulte
OTS-32	Matériel d’intubation complet
OTS-34	Oxymètre de pouls et saturomètre
OTS-37	AMBU pédiatrique
OTS-38	Aspirateur mucosités électrique
Stabilisateurs	Stabilisateurs
"""

RAW_PARN = """
18	Chariot à pansement
19	Chariot-Brancard
29	Escabeau à 2 marches
38	Lits métalliques avec porte moustiquaire et têtière réglable
50	Paravent
69	Potence à perfusion (Base)
69-1	Potence à perfusion - Pole
82	Table d'accouchement métallique
82-1	Table d'accouchement métallique - Birthing stool
82-2	Table d'accouchement métallique - Yoga ball (set de 2)
83	Table d'examen
84	Table d'examen Gynécologique
86	Table de chevet de malade
8	Balance pèse bébé saltère avec culotte
9	Balance pèse bébé
10	Balance Pèse personne avec toise
44	Mètre ruban
5	Bac de décontamination (lot de 3)
6	Bac plastique 120 litres avec couvercle et à roulette pour la cour
6-1	Bac plastique 120 litres - Accessoires
7	Bac plastique à déchets à pédale (50L gris) 
7-1	Bac plastique à déchets à pédale - 50L rouge
7-2	Bac plastique à déchets à pédale - 50L jaune
7-3	Bac plastique à déchets à pédale - Accessoires
12	Bassin de lit plastique avec manche
64	Pissette plastique 250 ml
64-1	Pissette plastique 500ml
64-2	Pissette plastique 1000ml
72	Réservoir d'eau 60 l muni de robinet
81	Support tubes à essai
87	Tablier plastique tous usages
96	Vessie eau chaude 2 litres
1	Armoire métallique à 2 portes
2	Armoire vestiaire métallique 1 casier
11	Bancs métalliques d'attente (3 seats)
11-1	Metal waiting benches - Arm rest & leg
11-2	Metal waiting benches - Rail
16	Chaise avec accoudoir
17	Chaise visiteur (frame)
17-1	Visitor chair - Back and seat
30	Etagère pour Pharmacie de Gros (Upper)
30-1	Etagère pour Pharmacie de Gros (Lower)
85	Table de bureau (Top Panel)
85-1	Office table - Front Panel
85-2	Office table - Legs
85-3	Office table - Steel Cabinet
88	Tabouret métallique
89	Tabouret tournant
13	Boîte à coton
25	Cupule en inox
34	Haricot inox 500 ml
70	Poubelle à pédale
94	Urinal
14	Boîte à instruments inox (avec couvercle)
15	Boite accouchement complete
20	Ciseau droit pointu 14 cm inox
21	Ciseaux courbe 14 cm inox
22	Ciseaux dauphin à pansement 14 cm inox
23	Ciseaux droits mousse 14 cm inox
24	Ciseaux pour épisiotomie courbe de 14 cm
39	Manche bistouri N° 3
40	Marteau à réflexe
51	Perce membrane
52	Pince à disséquer à griffes 14 cm inox
53	Pince à disséquer sans griffes 14 cm inox
54	Pince de Kocher à griffe 14 cm inox
55	Pince de Kocher à griffes 18 cm inox
56	Pince de Kocher sans griffe 14 cm inox
57	Pince de Kocher sans griffes 18 cm inox
58	Pince de péan ou porte coton 14 cm inox
59	Pince de Pozzi
60	Pince porte aiguille de Mayo Hégar 14 cm inox
61	Pince stérilisateur Cheatle forte
62	Pinces hémostatique inox
63	Pinces hémostatique inox courbe
65	Plateau inox à instruments GM (avec couvercle)
66	Plateau MM (pour instrument AMIU)
73	Sonde vésicale métallique pour femme
74	Speculum auriculaire inox (3 dimensions)
75	Speculum nasal inox
76	Speculum vaginal MM
77	Speculum vaginal GM
90	Tambour inox avec couvercle diam 200 mm
3	Aspiration Manuel Intra – Utérin (AMIU)
26	Doppler
27	Echelle de vision
28	Embu ressussitor
32	Garrots latex
48	Otoscope
49	Manual aspirator (initialment Palette)
79	Stéthoscope biauriculaire
79-1	Spare Eartips for Stethoscopes, standard size
79-2	Classic case for stethoscope
80	Stéthoscope obstétrical de Pinard
91	Tensiomètre anéroïde avec jeu de brassard
92	Thermomètre médical digital
95	Ventouse manuelle (Vacuum extracteur)
95-1	Delivery cups 50 mm
95-2	Delivery cups 60 mm
95-3	Delivery cups 70 mm
31	Extracteur de mucosité
33	Glacière
35	Irrigateur 1,5 litres avec tuyau pince raccord
43	Mesureur d'urine avec anses
67	Poire énéma avec canule
68	Porte vaccin
93	Tubes à essai
71	solar generator (initialement rechaud à petrole)
71-1	lithium battery
71-2	2.5kW Solar Generators - Rail
45	Microscope binoculaire
46	Minuterie 60 mm xls
47	Moustiquaire
41	Matelas mousse avec housse en plastique
4	Bac à sable/Extincteur
42	Mégaphone
37	Lampes solaires
361	Maintenance toolkit
78	Stérilisateur à vapeur + Pièces de rechange
78-1	Désioniseur d'eau ou autre purificateur d'eau
36	Lampe solaire (initialement lampe à alcool)
"""

RAW_PPSB = """
12	Chariot à pansement
13	Chariot-Brancard
23	Escabeau à 2 marches
31	Lits métalliques avec porte moustiquaire et têtière réglable
42	Paravent
61	Potence à perfusion (Base)
61-1	Potence à perfusion - Pole
74	Table d'accouchement métallique
74-1	Table d'accouchement métallique - Birthing stool
74-2	Table d'accouchement métallique - Yoga ball (set de 2)
75	Table d'examen
76	Table d'examen Gynécologique
77	Table de chevet de malade
5	Balance pèse bébé saltère avec culotte
6	Balance pèse bébé
7	Balance Pèse personne avec toise
36	Mètre ruban
2	Bac de décontamination (lot de 3)
3	Bac plastique 120 litres avec couvercle et à roulette pour la cour
3-1	Bac plastique 120 litres - Accessoires
4	Bac plastique à déchets à pédale (50L gris) 
4-1	Bac plastique à déchets à pédale - 50L rouge
4-2	Bac plastique à déchets à pédale - 50L jaune
4-3	Bac plastique à déchets à pédale - Accessoires
8	Bassin de lit plastique avec manche
56	Pissette plastique 250 ml
56-1	Pissette plastique 500ml
56-2	Pissette plastique 1000ml
64	Réservoir d'eau 60 l muni de robinet
73	Support tubes à essai
78	Tablier plastique tous usages
87	Vessie eau chaude 2 litres
88	Armoire métallique à 2 portes
89	Armoire vestiaire métallique 1 casier
901	Bancs métalliques d'attente (3 seats)
901-1	Metal waiting benches - Arm rest & leg
901-2	Metal waiting benches - Rail
91	Chaise avec accoudoir
92	Chaise visiteur (frame)
92-1	Visitor chair - Back and seat
24	Etagère pour Pharmacie de Gros (Upper)
24-1	Etagère pour Pharmacie de Gros (Lower)
94	Table de bureau (Top Panel)
94-1	Office table - Front Panel
94-2	Office table - Legs
94-3	Office table - Steel Cabinet
79	Tabouret métallique
80	Tabouret tournant
9	Boîte à coton
19	Cupule en inox
28	Haricot inox 500 ml
62	Poubelle à pédale
85	Urinal
10	Boîte à instruments inox (avec couvercle)
11	Boite accouchement complete
14	Ciseaux droit pointu 14 cm inox
15	Ciseaux courbe 14 cm inox
16	Ciseaux dauphin à pansement 14 cm inox
17	Ciseaux droits mousse 14 cm inox
18	Ciseaux pour épisiotomie courbe de 14 cm
32	Manche bistouri N° 3
33	Marteau à réflexe
43	Perce membrane
44	Pince à disséquer à griffes 14 cm inox
45	Pince à disséquer sans griffes 14 cm inox
46	Pince de Kocher à griffe 14 cm inox
47	Pince de Kocher à griffes 18 cm inox
48	Pince de Kocher sans griffe 14 cm inox
49	Pince de Kocher sans griffes 18 cm inox
50	Pince de péan ou porte coton 14 cm inox
51	Pince de Pozzi
52	Pince porte aiguille de Mayo Hégar 14 cm inox
53	Pince stérilisateur Cheatle forte
54	Pinces hémostatique inox
55	Pinces hémostatique inox courbe
57	Plateau inox à instruments GM (avec couvercle)
58	Plateau MM (pour instrument AMIU)
65	Sonde vésicale métallique pour femme
66	Speculum auriculaire inox (3 dimensions)
67	Speculum nasal inox
68	Speculum vaginal MM
69	Speculum vaginal GM
81	Tambour inox avec couvercle diam 200 mm
1	Aspiration Manuel Intra – Utérin (AMIU)
20	Doppler
21	Echelle de vision
22	Embu ressussitor
26	Garrots latex
40	Otoscope
41	Manual aspirator (initialment Palette)
71	Stéthoscope biauriculaire
71-1	Spare Eartips for Stethoscopes, standard size
71-2	Classic case for stethoscope
72	Stéthoscope obstétrical de Pinard
82	Tensiomètre anéroïde avec jeu de brassard
83	Thermomètre médical digital
86	Ventouse manuelle (Vacuum extracteur)
86-1	Delivery cups 50 mm
86-2	Delivery cups 60 mm
86-3	Delivery cups 70 mm
25	Extracteur de mucosité
27	Glacière
29	Irrigateur 1,5 litres avec tuyau pince raccord
35	Mesureur d'urine avec anses
59	Poire énéma avec canule
60	Porte vaccin
84	Tubes à essai
63	solar generator (initialement rechaud à petrole)
63-1	lithium battery
63-2	2.5kW Solar Generators - Rail
37	Microscope binoculaire
38	Minuterie 60 mm xls
39	Moustiquaire
34	Matelas mousse avec housse en plastique
90	Bac à sable/Extincteur
93	Mégaphone
922	Lampes solaires
301	Maintenance toolkit
70	Stérilisateur à vapeur + Pièces de rechange
70-1	Désioniseur d'eau ou autre purificateur d'eau
30	Lampe solaire (initialement lampe à alcool)
"""

def extraire_dictionnaire(texte_brut):
    dico = {}
    for ligne in texte_brut.strip().split('\n'):
        if ligne.strip():
            parts = ligne.strip().split(maxsplit=1)
            if len(parts) == 2:
                dico[parts[1].strip()] = parts[0].strip()
    return dico

CATALOGUES = {
    "HÔPITAUX": extraire_dictionnaire(RAW_HOPITAUX),
    "PARN": extraire_dictionnaire(RAW_PARN),
    "PPSB": extraire_dictionnaire(RAW_PPSB)
}

# 4. Gestion des données Excel
def charger():
    if os.path.exists(FICHIER):
        try:
            art = pd.read_excel(FICHIER, sheet_name="Articles").reindex(columns=COL_ART)
            ent = pd.read_excel(FICHIER, sheet_name="Entrepôts").reindex(columns=COL_ENT)
            mvt = pd.read_excel(FICHIER, sheet_name="Mouvements").reindex(columns=COL_MVT)
            return art, ent, mvt
        except Exception:
            pass
    return pd.DataFrame(columns=COL_ART), pd.DataFrame(columns=COL_ENT), pd.DataFrame(columns=COL_MVT)

def calculer_stock(id_article, entrepot, mvt_df):
    stock = 0.0
    m_art = mvt_df[mvt_df["ID Article"].astype(str) == str(id_article)]
    for _, m in m_art.iterrows():
        q = float(m["Quantité"])
        t = m["Type"]
        if t == "Entrée" and str(m["Entrepôt"]) == str(entrepot): stock += q
        elif t == "Sortie" and str(m["Entrepôt"]) == str(entrepot): stock -= q
        elif t == "Transfert":
            if str(m["Entrepôt"]) == str(entrepot): stock -= q
            if str(m["Entrepôt de destination"]) == str(entrepot): stock += q
        elif t == "Ajustement" and str(m["Entrepôt"]) == str(entrepot): stock = q
    return stock

def generer_inventaire_complet(filtre="", art=None, mvt=None):
    if art is None or mvt is None:
        art, _, mvt = charger()
    inventaire_list = []
    
    for _, a in art.iterrows():
        id_art = str(a["ID"])
        try: b_val = float(a.get("Unités/Pack", 0))
        except: b_val = 0.0
        if pd.isna(b_val): b_val = 0.0

        mvt_art = mvt[mvt["ID Article"].astype(str) == id_art]
        entrepots = set(mvt_art["Entrepôt"].dropna().tolist())
        entrepots.update(mvt_art["Entrepôt de destination"].dropna().tolist())
        if pd.notna(a["Entrepôt par défaut"]): entrepots.add(a["Entrepôt par défaut"])
        entrepots.discard("")
        
        for ent_name in entrepots:
            stk = calculer_stock(id_art, ent_name, mvt)
            if stk > 0 or ent_name == a["Entrepôt par défaut"]:
                pack_a = int(stk // b_val) if b_val > 0 else 0
                vrac_c = (stk % b_val) if b_val > 0 else stk
                
                # Récupération de la date (dernière activité ou date de création)
                mvt_ent = mvt_art[(mvt_art["Entrepôt"] == ent_name) | (mvt_art["Entrepôt de destination"] == ent_name)]
                date_aff = str(mvt_ent["Date"].iloc[-1]) if not mvt_ent.empty else str(a.get("Date", "-"))
                if date_aff == "nan": date_aff = "-"

                inventaire_list.append({
                    "ID": id_art,
                    "Date": date_aff,
                    "Nom Article": a["Nom Article"],
                    "Nomenclature": a["Nomenclature"],
                    "Catégorie": a["Catégorie"],
                    "Projet": a["Projet"],
                    "Nbr Pack (A)": pack_a,
                    "Unités/Pack (B)": b_val,
                    "Vrac (C)": vrac_c,
                    "Quantité Totale (D)": stk,
                    "Entrepôt": ent_name
                })
                
    cols = ["ID", "Date", "Nom Article", "Nomenclature", "Catégorie", "Projet", "Nbr Pack (A)", "Unités/Pack (B)", "Vrac (C)", "Quantité Totale (D)", "Entrepôt"]
    df = pd.DataFrame(inventaire_list)[cols] if inventaire_list else pd.DataFrame(columns=cols)
    
    if filtre and str(filtre).strip() != "" and not df.empty:
        f_txt = str(filtre).strip().lower()
        mask = df.apply(lambda row: f_txt in ' '.join(row.values.astype(str)).lower(), axis=1)
        df = df[mask]

    nb_lignes = len(df)
    tot_packs = int(df["Nbr Pack (A)"].sum()) if not df.empty else 0
    tot_vrac = df["Vrac (C)"].sum() if not df.empty else 0
    tot_general = df["Quantité Totale (D)"].sum() if not df.empty else 0

    bandeau_totaux = f"""
    <div style="display:flex; flex-wrap:wrap; gap:15px; background:#f0f9ff; border-left:5px solid #0096D9; padding:12px 18px; border-radius:8px; margin-bottom:10px; font-family:sans-serif;">
        <div><strong>📦 Lignes d'articles :</strong> <span style="color:#0096D9; font-size:16px; font-weight:bold;">{nb_lignes}</span></div>
        <div style="margin-left:20px;"><strong>📁 Total Packs/Cartons (A) :</strong> <span style="color:#0096D9; font-size:16px; font-weight:bold;">{tot_packs:,}</span></div>
        <div style="margin-left:20px;"><strong>🔹 Total Vrac (C) :</strong> <span style="color:#0096D9; font-size:16px; font-weight:bold;">{tot_vrac:,.0f}</span></div>
        <div style="margin-left:20px;"><strong>✅ QUANTITÉ TOTALE (D) :</strong> <span style="color:#0369a1; font-size:18px; font-weight:bold;">{tot_general:,.0f}</span></div>
    </div>
    """

    if not df.empty:
        ligne_total = pd.DataFrame([{
            "ID": "TOTAL",
            "Date": "-",
            "Nom Article": f"TOTAL GÉNÉRAL ({nb_lignes} articles)",
            "Nomenclature": "-",
            "Catégorie": "-",
            "Projet": "-",
            "Nbr Pack (A)": tot_packs,
            "Unités/Pack (B)": "-",
            "Vrac (C)": tot_vrac,
            "Quantité Totale (D)": tot_general,
            "Entrepôt": "-"
        }])
        df_affichage = pd.concat([df, ligne_total], ignore_index=True)
    else:
        df_affichage = df

    return df_affichage, bandeau_totaux

def sauvegarder(art, ent, mvt):
    inv_df, _ = generer_inventaire_complet(art=art, mvt=mvt)
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as writer:
        inv_df.to_excel(writer, sheet_name="Inventaire Actuel", index=False)
        art.to_excel(writer, sheet_name="Articles", index=False)
        ent.to_excel(writer, sheet_name="Entrepôts", index=False)
        mvt.to_excel(writer, sheet_name="Mouvements", index=False)
    return FICHIER

def exporter_excel_filtre(filtre):
    art, ent, mvt = charger()
    inv_df, _ = generer_inventaire_complet(filtre=filtre, art=art, mvt=mvt)
    with pd.ExcelWriter(FICHIER, engine="openpyxl") as writer:
        inv_df.to_excel(writer, sheet_name="Inventaire Actuel", index=False)
        art.to_excel(writer, sheet_name="Articles", index=False)
        ent.to_excel(writer, sheet_name="Entrepôts", index=False)
        mvt.to_excel(writer, sheet_name="Mouvements", index=False)
    return FICHIER

def obtenir_listes(art=None, ent=None, mvt=None):
    if art is None or ent is None or mvt is None:
        art, ent, mvt = charger()
    l_ent = sorted(ent["Nom Entrepôt"].dropna().astype(str).unique().tolist())
    l_art = [f"{r['ID']} - {r['Nomenclature']} - {r['Nom Article']}" for _, r in art.iterrows()]
    l_mvt = [f"{i} | {r['Date']} - {r['Type']} - {r['Nom Article']} ({r['Quantité']})" for i, r in mvt.iterrows()]
    return l_ent, l_art, l_mvt

# 5. Logique Nomenclatures & Calculs
def maj_champs_projet(projet):
    if projet in CATALOGUES:
        return gr.update(choices=list(CATALOGUES[projet].keys()), value=None, interactive=True), gr.update(value="", interactive=False)
    return gr.update(choices=[], value=None, interactive=True), gr.update(value="", interactive=True)

def auto_remplir_nomenclature(projet, nom_article):
    if projet in CATALOGUES and nom_article in CATALOGUES[projet]:
        return gr.update(value=CATALOGUES[projet][nom_article])
    return gr.update() if projet == "UNOPS" else gr.update(value="")

def calculer_total_ui(a, b, c):
    try: return (float(a or 0) * float(b or 0)) + float(c or 0)
    except: return 0

# 6. Actions CRUD (Avec vidage automatique)
def action_ajouter_entrepot(nom):
    art, ent, mvt = charger()
    if not nom or not nom.strip(): return "⚠️ Le nom est vide.", ent, gr.update(), gr.update(), gr.update(), gr.update(), gr.update()
    nom = nom.strip()
    if nom in ent["Nom Entrepôt"].values: return "⚠️ Cet entrepôt existe déjà.", ent, gr.update(), gr.update(), gr.update(), gr.update(), gr.update()
    
    new_id = 1 if ent.empty else int(ent["ID"].max()) + 1
    ent = pd.concat([ent, pd.DataFrame([[new_id, nom]], columns=COL_ENT)], ignore_index=True)
    sauvegarder(art, ent, mvt)
    l_ent, _, _ = obtenir_listes(art, ent, mvt)
    return f"✅ Entrepôt '{nom}' ajouté.", ent, gr.update(value=""), gr.update(choices=l_ent), gr.update(choices=l_ent), gr.update(choices=l_ent), gr.update(choices=l_ent)

def action_ajouter_article(projet, nomenclature, nom_article, categorie, a_pack, a_unit, a_vrac, quantite, entrepot, responsable, commentaires):
    art, ent, mvt = charger()
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    fail_out = (art, inv_df, inv_html) + (gr.update(),)*13
    if not nom_article or not entrepot: return ("⚠️ Nom et Entrepôt obligatoires.",) + fail_out
    
    new_id = 1 if art.empty else int(art["ID"].max()) + 1
    date_jour = datetime.now().strftime("%Y-%m-%d")
    
    art = pd.concat([art, pd.DataFrame([[new_id, date_jour, nom_article, nomenclature, categorie, projet, float(a_unit or 0), entrepot, responsable, commentaires]], columns=COL_ART)], ignore_index=True)
    qte = float(quantite or 0)
    if qte > 0:
        obs = f"Ajout initial - {commentaires}" if commentaires else "Ajout initial"
        mvt = pd.concat([mvt, pd.DataFrame([[date_jour, "Entrée", new_id, nomenclature, nom_article, qte, entrepot, "", 0, qte, "Création initiale", projet, "-", responsable, obs]], columns=COL_MVT)], ignore_index=True)
        
    sauvegarder(art, ent, mvt)
    _, l_art, _ = obtenir_listes(art, ent, mvt)
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    return (
        f"✅ Article '{nom_article}' enregistré.", art, inv_df, inv_html,
        gr.update(choices=l_art), gr.update(choices=l_art),
        gr.update(value="UNOPS"), gr.update(value=None, choices=[]), gr.update(value="", interactive=True), gr.update(value=None),
        gr.update(value=0), gr.update(value=0), gr.update(value=0), gr.update(value=0),
        gr.update(value=None), gr.update(value=""), gr.update(value="")
    )

def action_ajouter_mouvement(date_mvt, type_mvt, selection_article, m_pack, m_unit, m_vrac, quantite, ent_dep, ent_dest, doc, site, resp, obs):
    art, ent, mvt = charger()
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    fail_out = (mvt, inv_df, inv_html) + (gr.update(),)*14
    if not selection_article or not ent_dep: return ("⚠️ Article et Entrepôt obligatoires.",) + fail_out
    
    try:
        id_art = selection_article.split(" - ")[0]
        article_info = art[art["ID"].astype(str) == id_art].iloc[0]
    except: return ("⚠️ Article invalide.",) + fail_out
        
    qte = float(quantite or 0)
    if qte <= 0: return ("⚠️ Quantité doit être > 0.",) + fail_out
    
    stock_actuel = calculer_stock(id_art, ent_dep, mvt)
    if type_mvt in ["Sortie", "Transfert"] and qte > stock_actuel: return (f"⚠️ Stock insuffisant ({stock_actuel} dispo).",) + fail_out
    if type_mvt == "Transfert" and (not ent_dest or ent_dep == ent_dest): return ("⚠️ Entrepôt destination invalide.",) + fail_out

    stock_apres = stock_actuel + qte if type_mvt == "Entrée" else (stock_actuel - qte if type_mvt in ["Sortie", "Transfert"] else qte)

    mvt = pd.concat([mvt, pd.DataFrame([[date_mvt, type_mvt, id_art, article_info["Nomenclature"], article_info["Nom Article"], qte, ent_dep, ent_dest if type_mvt == "Transfert" else "", stock_actuel, stock_apres, doc, article_info["Projet"], site, resp, obs]], columns=COL_MVT)], ignore_index=True)
    sauvegarder(art, ent, mvt)
    _, _, l_mvt = obtenir_listes(art, ent, mvt)
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    
    return (
        "✅ Mouvement validé.", mvt, inv_df, inv_html, gr.update(choices=l_mvt),
        gr.update(value=datetime.now().strftime("%Y-%m-%d")), gr.update(value="Entrée"), gr.update(value=None),
        gr.update(value=0), gr.update(value=0), gr.update(value=0), gr.update(value=0),
        gr.update(value=None), gr.update(value=None), gr.update(value=""), gr.update(value=""), gr.update(value=""), gr.update(value="")
    )

def supprimer_article(selection):
    if not selection: return "⚠️ Sélectionnez un article.", gr.update(), gr.update(), gr.update(), gr.update(), gr.update()
    art, ent, mvt = charger()
    art = art[art["ID"] != int(selection.split(" - ")[0])]
    sauvegarder(art, ent, mvt)
    _, l_art, _ = obtenir_listes(art, ent, mvt)
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    return "🗑️ Article supprimé.", art, inv_df, inv_html, gr.update(choices=l_art, value=None), gr.update(choices=l_art, value=None)

def supprimer_entrepot(nom):
    if not nom: return "⚠️ Sélectionnez un entrepôt.", gr.update(), gr.update(), gr.update(), gr.update(), gr.update()
    art, ent, mvt = charger()
    ent = ent[ent["Nom Entrepôt"] != nom]
    sauvegarder(art, ent, mvt)
    l_ent, _, _ = obtenir_listes(art, ent, mvt)
    return "🗑️ Entrepôt supprimé.", ent, gr.update(choices=l_ent, value=None), gr.update(choices=l_ent, value=None), gr.update(choices=l_ent, value=None), gr.update(choices=l_ent, value=None)

def supprimer_mouvement(selection):
    if not selection: return "⚠️ Sélectionnez un mouvement.", gr.update(), gr.update(), gr.update(), gr.update()
    idx = int(selection.split(" | ")[0])
    art, ent, mvt = charger()
    if idx in mvt.index:
        mvt = mvt.drop(idx)
        sauvegarder(art, ent, mvt)
    _, _, l_mvt = obtenir_listes(art, ent, mvt)
    inv_df, inv_html = generer_inventaire_complet(art=art, mvt=mvt)
    return "🗑️ Mouvement supprimé.", mvt, inv_df, inv_html, gr.update(choices=l_mvt, value=None)

def generer_dashboard():
    art, _, mvt = charger()
    inv_df, _ = generer_inventaire_complet(art=art, mvt=mvt)
    inv_sans_total = inv_df[inv_df["ID"] != "TOTAL"] if not inv_df.empty else inv_df
    total_stock = inv_sans_total["Quantité Totale (D)"].sum() if not inv_sans_total.empty else 0
    entrees = mvt[mvt["Type"] == "Entrée"]["Quantité"].astype(float).sum() if not mvt.empty else 0
    sorties = mvt[mvt["Type"] == "Sortie"]["Quantité"].astype(float).sum() if not mvt.empty else 0
    
    resume = f"""
    ### 📊 Indicateurs Clés
    | Indicateur | Valeur |
    |---|---:|
    | **Articles enregistrés** | {len(art)} |
    | **Stock total (Unités)** | {total_stock:,.0f} |
    | **Total des Entrées** | {entrees:,.0f} |
    | **Total des Sorties** | {sorties:,.0f} |
    """
    if inv_sans_total.empty: return resume, None, None
    fig1 = px.bar(inv_sans_total.groupby("Entrepôt", as_index=False)["Quantité Totale (D)"].sum(), x="Entrepôt", y="Quantité Totale (D)", title="Stock Total par Entrepôt", color_discrete_sequence=["#0096D9"])
    fig2 = px.bar(inv_sans_total.groupby("Projet", as_index=False)["Quantité Totale (D)"].sum(), x="Projet", y="Quantité Totale (D)", title="Stock Alloué par Projet", color_discrete_sequence=["#0096D9"])
    return resume, fig1, fig2

# 7. Interface Graphique
head_html = '<link rel="icon" type="image/png" href="https://upload.wikimedia.org/wikipedia/commons/thumb/e/ee/UN_emblem_blue.svg/150px-UN_emblem_blue.svg.png">'
css_custom = "footer {display: none !important;}"
theme_unops = gr.themes.Soft(primary_hue="blue", secondary_hue="blue").set(button_primary_background_fill="#0096D9", button_primary_background_fill_hover="#007BBA", button_primary_text_color="white")

with gr.Blocks(title="UNOPS - Gestion d'inventaire", theme=theme_unops, head=head_html, css=css_custom) as app:
    gr.HTML('''
    <div style="display:flex;align-items:center;gap:22px;padding:16px 20px; background:#ffffff;border-bottom:4px solid #0096D9;border-radius:10px;margin-bottom:14px;">
      <img src="https://upload.wikimedia.org/wikipedia/commons/5/57/UNOPS_logo_2016_website_blue_304x53.png" alt="Logo UNOPS" style="width:220px;max-width:45%;height:auto;object-fit:contain;">
      <div style="margin-left:auto;text-align:right;color:#334155;">
        <div style="font-size:20px;font-weight:700;">GESTION D'INVENTAIRE INTÉGRÉE</div>
      </div>
    </div>
    ''')

    with gr.Tabs():
        with gr.TabItem("📈 Tableau de Bord"):
            btn_dash = gr.Button("🔄 Actualiser le Tableau de Bord", variant="primary")
            resume_markdown = gr.Markdown()
            with gr.Row():
                graph_entrepot = gr.Plot()
                graph_projet = gr.Plot()
            btn_dash.click(generer_dashboard, outputs=[resume_markdown, graph_entrepot, graph_projet])

        with gr.TabItem("📦 Articles"):
            with gr.Row():
                with gr.Column(scale=1):
                    gr.Markdown("### Nouvel Article")
                    a_projet = gr.Dropdown(["UNOPS", "PARN", "PPSB", "HÔPITAUX"], label="Projet", value="UNOPS")
                    a_nom = gr.Dropdown(choices=[], allow_custom_value=True, label="Nom Article *")
                    a_nomen = gr.Textbox(label="Nomenclature", interactive=True)
                    a_cat = gr.Dropdown(choices=["CSB1", "CSB2", "HÔPITAUX", "AUTRES"], label="Catégorie")
                    
                    a_projet.change(maj_champs_projet, inputs=[a_projet], outputs=[a_nom, a_nomen])
                    a_nom.change(auto_remplir_nomenclature, inputs=[a_projet, a_nom], outputs=[a_nomen])
                    
                    gr.Markdown("**Calcul de la Quantité Initiale :**")
                    with gr.Row():
                        a_pack = gr.Number(label="Nbr Pack/Carton (A)", value=0, minimum=0)
                        a_unit = gr.Number(label="Unités par Pack (B)", value=0, minimum=0)
                    with gr.Row():
                        a_vrac = gr.Number(label="Unités en Vrac (C)", value=0, minimum=0)
                        a_qte = gr.Number(label="Quantité Totale (D)", value=0, interactive=False)
                    
                    gr.on([a_pack.change, a_unit.change, a_vrac.change], calculer_total_ui, inputs=[a_pack, a_unit, a_vrac], outputs=[a_qte])
                    
                    a_ent = gr.Dropdown(choices=[], label="Entrepôt par défaut *")
                    a_resp = gr.Textbox(label="Responsable")
                    a_com = gr.Textbox(label="Commentaires (Optionnel)", lines=2)
                    
                    btn_add_art = gr.Button("Enregistrer l'article", variant="primary")
                    msg_art = gr.Textbox(label="Statut Saisie", interactive=False)
                with gr.Column(scale=2):
                    gr.Markdown("### Base des Articles")
                    vue_articles = gr.Dataframe(charger()[0], interactive=False)
                    with gr.Row():
                        del_art_drop = gr.Dropdown(choices=[], label="Sélectionner l'article à supprimer")
                        btn_del_art = gr.Button("Supprimer", variant="stop")
                    msg_del_art = gr.Textbox(label="Statut Suppression", interactive=False)

        with gr.TabItem("🏭 Entrepôts"):
            with gr.Row():
                with gr.Column(scale=1):
                    gr.Markdown("### Nouvel Entrepôt")
                    e_nom = gr.Textbox(label="Nom de l'Entrepôt *")
                    btn_add_ent = gr.Button("Ajouter", variant="primary")
                    msg_ent = gr.Textbox(label="Statut Saisie", interactive=False)
                with gr.Column(scale=2):
                    gr.Markdown("### Base des Entrepôts")
                    vue_entrepots = gr.Dataframe(charger()[1], interactive=False)
                    with gr.Row():
                        del_ent_drop = gr.Dropdown(choices=[], label="Sélectionner l'entrepôt à supprimer")
                        btn_del_ent = gr.Button("Supprimer", variant="stop")
                    msg_del_ent = gr.Textbox(label="Statut Suppression", interactive=False)

        with gr.TabItem("📊 Inventaire Actuel"):
            with gr.Row():
                search_inv = gr.Textbox(label="🔍 Rechercher (Article, Nomenclature, Projet, Entrepôt...)", scale=3)
                btn_refresh_inv = gr.Button("🔄 Actualiser", variant="primary", scale=1)
                btn_export_excel = gr.DownloadButton("📥 Télécharger Excel", value=FICHIER, variant="secondary", scale=1)
            
            init_inv_df, init_inv_html = generer_inventaire_complet()
            totaux_inv_html = gr.HTML(init_inv_html)
            vue_inventaire_visible = gr.Dataframe(init_inv_df, interactive=False)
            
            def maj_et_preparer_excel(filtre):
                df, html = generer_inventaire_complet(filtre)
                fichier_pret = exporter_excel_filtre(filtre)
                return df, html, fichier_pret

            search_inv.change(maj_et_preparer_excel, inputs=[search_inv], outputs=[vue_inventaire_visible, totaux_inv_html, btn_export_excel])
            btn_refresh_inv.click(maj_et_preparer_excel, inputs=[search_inv], outputs=[vue_inventaire_visible, totaux_inv_html, btn_export_excel])

        with gr.TabItem("🔄 Mouvements"):
            with gr.Row():
                with gr.Column(scale=1):
                    gr.Markdown("### Nouveau Mouvement")
                    m_date = gr.Textbox(label="Date", value=datetime.now().strftime("%Y-%m-%d"))
                    m_type = gr.Dropdown(["Entrée", "Sortie", "Transfert", "Ajustement"], label="Type de Mouvement", value="Entrée")
                    m_art = gr.Dropdown(choices=[], label="Sélectionner l'article *")
                    
                    gr.Markdown("**Calcul de la Quantité :**")
                    with gr.Row():
                        m_pack = gr.Number(label="Nbr Pack/Carton (A)", value=0, minimum=0)
                        m_unit = gr.Number(label="Unités par Pack (B)", value=0, minimum=0)
                    with gr.Row():
                        m_vrac = gr.Number(label="Unités en Vrac (C)", value=0, minimum=0)
                        m_qte = gr.Number(label="Quantité Totale (D)", value=0, interactive=False)
                    
                    gr.on([m_pack.change, m_unit.change, m_vrac.change], calculer_total_ui, inputs=[m_pack, m_unit, m_vrac], outputs=[m_qte])

                    m_ent_dep = gr.Dropdown(choices=[], label="Entrepôt concerné / départ *")
                    m_ent_dest = gr.Dropdown(choices=[], label="Entrepôt de destination (Transfert uniquement)")
                    m_doc = gr.Textbox(label="N° Document")
                    m_site = gr.Textbox(label="Site bénéficiaire")
                    m_resp = gr.Textbox(label="Responsable")
                    m_obs = gr.Textbox(label="Observations")
                    btn_add_mvt = gr.Button("Valider le mouvement", variant="primary")
                    msg_mvt = gr.Textbox(label="Statut Saisie", interactive=False)
                with gr.Column(scale=2):
                    gr.Markdown("### Historique des mouvements")
                    vue_mouvements = gr.Dataframe(charger()[2], interactive=False)
                    with gr.Row():
                        del_mvt_drop = gr.Dropdown(choices=[], label="Sélectionner le mouvement à supprimer")
                        btn_del_mvt = gr.Button("Supprimer", variant="stop")
                    msg_del_mvt = gr.Textbox(label="Statut Suppression", interactive=False)

    # Connexions
    btn_add_ent.click(action_ajouter_entrepot, inputs=[e_nom], outputs=[msg_ent, vue_entrepots, e_nom, del_ent_drop, a_ent, m_ent_dep, m_ent_dest])
    btn_del_ent.click(supprimer_entrepot, inputs=[del_ent_drop], outputs=[msg_del_ent, vue_entrepots, del_ent_drop, a_ent, m_ent_dep, m_ent_dest])
    
    btn_add_art.click(
        action_ajouter_article, 
        inputs=[a_projet, a_nomen, a_nom, a_cat, a_pack, a_unit, a_vrac, a_qte, a_ent, a_resp, a_com], 
        outputs=[msg_art, vue_articles, vue_inventaire_visible, totaux_inv_html, del_art_drop, m_art, a_projet, a_nom, a_nomen, a_cat, a_pack, a_unit, a_vrac, a_qte, a_ent, a_resp, a_com]
    )
    btn_del_art.click(supprimer_article, inputs=[del_art_drop], outputs=[msg_del_art, vue_articles, vue_inventaire_visible, totaux_inv_html, del_art_drop, m_art])
    
    btn_add_mvt.click(
        action_ajouter_mouvement, 
        inputs=[m_date, m_type, m_art, m_pack, m_unit, m_vrac, m_qte, m_ent_dep, m_ent_dest, m_doc, m_site, m_resp, m_obs], 
        outputs=[msg_mvt, vue_mouvements, vue_inventaire_visible, totaux_inv_html, del_mvt_drop, m_date, m_type, m_art, m_pack, m_unit, m_vrac, m_qte, m_ent_dep, m_ent_dest, m_doc, m_site, m_resp, m_obs]
    )
    btn_del_mvt.click(supprimer_mouvement, inputs=[del_mvt_drop], outputs=[msg_del_mvt, vue_mouvements, vue_inventaire_visible, totaux_inv_html, del_mvt_drop])

    def init_app():
        l_ent, l_art, l_mvt = obtenir_listes()
        sauvegarder(*charger())
        inv_df, inv_html = generer_inventaire_complet()
        return generer_dashboard() + (
            gr.update(choices=l_ent), gr.update(choices=l_ent), gr.update(choices=l_ent), gr.update(choices=l_ent),
            gr.update(choices=l_art), gr.update(choices=l_art), gr.update(choices=l_mvt),
            inv_df, inv_html, FICHIER
        )
    
    app.load(init_app, outputs=[resume_markdown, graph_entrepot, graph_projet, a_ent, del_ent_drop, m_ent_dep, m_ent_dest, m_art, del_art_drop, del_mvt_drop, vue_inventaire_visible, totaux_inv_html, btn_export_excel])

msg_connexion = """
<div style="text-align: center; padding: 20px;">
    <img src="https://upload.wikimedia.org/wikipedia/commons/5/57/UNOPS_logo_2016_website_blue_304x53.png" style="max-width: 250px; margin: 0 auto 20px auto;">
    <h2 style="color: #0096D9; font-weight: bold; margin: 0; font-family: sans-serif;">Inventaire UNOPS</h2>
    <p style="color: #64748b; font-size: 14px; margin-top: 10px; font-family: sans-serif;">Veuillez entrer vos identifiants pour accéder au système.</p>
</div>
"""

app.launch(server_name="0.0.0.0", server_port=int(os.environ.get("PORT", 7860)), auth=("admin", "admin"), auth_message=msg_connexion)
