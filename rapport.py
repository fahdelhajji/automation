import os
import traceback
from datetime import date, datetime

import pandas as pd
import win32com.client as win32

# ============================================================
# PARAMETERS
# ============================================================
DOSSIER = r"C:\Automatisation"
FICHIER_EXPORT = os.path.join(DOSSIER, "export.xlsx")
FICHIER_RAPPORT = os.path.join(DOSSIER, f"rapport_{date.today()}.xlsx")
FICHIER_LOG = os.path.join(DOSSIER, "log.txt")

DESTINATAIRES = "fahd.el-hajji@dedalus.com"

MODE_TEST = False  # True = le mail s'affiche sans partir / False = il part vraiment

# Noms des colonnes dans ton Excel (exactement comme écrits)
COL_CLIENT = "Client"
COL_COMMANDE = "Commande"
COL_MONTANT = "Montant"
# ============================================================


def ecrire_log(message):
    print(message)
    with open(FICHIER_LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} - {message}\n")


def lire_donnees():
    """Lit l'Excel et trouve tout seul la ligne des titres."""
    brut = pd.read_excel(FICHIER_EXPORT, header=None)
    attendues = [COL_CLIENT, COL_COMMANDE, COL_MONTANT]

    for i in range(len(brut)):
        valeurs = [str(v).strip() for v in brut.iloc[i]]
        if all(c in valeurs for c in attendues):
            df = brut.iloc[i + 1:].copy()
            df.columns = valeurs
            df = df[[c for c in df.columns if c != "nan"]]
            df = df.dropna(how="all").reset_index(drop=True)
            df[COL_MONTANT] = pd.to_numeric(df[COL_MONTANT], errors="coerce")
            return df

    contenu = brut.head(5).to_string()
    raise ValueError(
        f"Colonnes {attendues} introuvables dans le fichier.\n"
        f"Voici les 5 premieres lignes lues :\n{contenu}"
    )


def calculer_kpi(df):
    kpi_clients = df.groupby(COL_CLIENT).agg(
        total_ventes=(COL_MONTANT, "sum"),
        nb_commandes=(COL_COMMANDE, "count"),
    ).reset_index()
    kpi_clients["panier_moyen"] = (
        kpi_clients["total_ventes"] / kpi_clients["nb_commandes"]
    ).round(2)
    kpi_clients = kpi_clients.sort_values("total_ventes", ascending=False)

    resume = pd.DataFrame({
        "KPI": ["Total des ventes", "Nombre de commandes", "Panier moyen", "Nombre de clients"],
        "Valeur": [
            round(df[COL_MONTANT].sum(), 2),
            df[COL_COMMANDE].count(),
            round(df[COL_MONTANT].mean(), 2),
            df[COL_CLIENT].nunique(),
        ],
    })
    return resume, kpi_clients


def creer_excel(resume, kpi_clients, df):
    with pd.ExcelWriter(FICHIER_RAPPORT, engine="openpyxl") as writer:
        resume.to_excel(writer, sheet_name="Résumé", index=False)
        kpi_clients.to_excel(writer, sheet_name="Par client", index=False)
        df.to_excel(writer, sheet_name="Données brutes", index=False)

        for feuille in writer.sheets.values():
            for colonne in feuille.columns:
                largeur = max(len(str(c.value)) if c.value is not None else 0 for c in colonne)
                feuille.column_dimensions[colonne[0].column_letter].width = largeur + 3


def envoyer_mail(resume):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = DESTINATAIRES
    mail.Subject = f"Rapport KPI du {date.today():%d/%m/%Y}"

    lignes = "\n".join(f"- {r.KPI} : {r.Valeur}" for r in resume.itertuples())
    mail.Body = (
        "Bonjour,\n\n"
        "Veuillez trouver ci-joint le rapport du jour.\n\n"
        f"Chiffres clés :\n{lignes}\n\n"
        "Cordialement"
    )
    mail.Attachments.Add(FICHIER_RAPPORT)

    if MODE_TEST:
        mail.Display()
    else:
        mail.Send()


def main():
    try:
        ecrire_log("VERSION 2 - Début du script")
        df = lire_donnees()
        ecrire_log(f"{len(df)} lignes lues")
        resume, kpi_clients = calculer_kpi(df)
        creer_excel(resume, kpi_clients, df)
        ecrire_log("Excel créé")
        envoyer_mail(resume)
        ecrire_log("Mail envoyé (ou affiché en mode test)")
    except Exception:
        ecrire_log("ERREUR :\n" + traceback.format_exc())


if __name__ == "__main__":
    main()