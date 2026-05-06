import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os
from datetime import datetime

# ─── Fichier Excel de sortie ───────────────────────────────
FICHIER_EXCEL = r"C:\Users\USER\OneDrive - ENSEA\Bureau\SUNU\SUNU_Collecte.xlsx"
def initialiser_excel():
    """Crée le fichier Excel avec les colonnes si il n'existe pas."""
    if not os.path.exists(FICHIER_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Collecte"
        ws.append([
            "Date", "Nom Employé", "Canal",
            "Périmètre", "NPS Score", "Catégorie NPS",
            "CSAT", "CES"
        ])
        wb.save(FICHIER_EXCEL)

def classer_nps(score):
    """Classifie le score NPS."""
    if score >= 9:
        return "Promoteur"
    elif score >= 7:
        return "Passif"
    else:
        return "Détracteur"

def enregistrer():
    """Enregistre la saisie dans Excel."""
    # Récupération des valeurs
    employe = entry_employe.get().strip()
    canal = combo_canal.get()
    perimetre = combo_perimetre.get()
    nps_val = entry_nps.get().strip()
    csat = csat_var.get()
    ces = ces_var.get()

    # Vérifications
    if not employe:
        messagebox.showwarning("Attention", "Veuillez entrer le nom de l'employé.")
        return
    if canal == "":
        messagebox.showwarning("Attention", "Veuillez sélectionner un canal.")
        return
    if perimetre == "":
        messagebox.showwarning("Attention", "Veuillez sélectionner un périmètre.")
        return
    try:
        nps_score = int(nps_val)
        if nps_score < 0 or nps_score > 10:
            raise ValueError
    except ValueError:
        messagebox.showwarning("Attention", "Le score NPS doit être un entier entre 0 et 10.")
        return
    if csat == "":
        messagebox.showwarning("Attention", "Veuillez répondre à la question CSAT.")
        return
    if ces == "":
        messagebox.showwarning("Attention", "Veuillez répondre à la question CES.")
        return

    # Enregistrement
    wb = openpyxl.load_workbook(FICHIER_EXCEL)
    ws = wb["Collecte"]
    ws.append([
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        employe, canal, perimetre,
        nps_score, classer_nps(nps_score),
        csat, ces
    ])
    wb.save(FICHIER_EXCEL)
    messagebox.showinfo("Succès", "Réponse enregistrée avec succès ✅")

    # Réinitialisation
    entry_nps.delete(0, tk.END)
    csat_var.set("")
    ces_var.set("")

# ─── Interface graphique ───────────────────────────────────
initialiser_excel()

root = tk.Tk()
root.title("SUNU Assurances — Saisie Relation Client")
root.geometry("500x520")
root.resizable(False, False)
root.configure(bg="#1A2B4A")

# Titre
tk.Label(root, text="SUNU Business — Saisie Client",
         font=("Helvetica", 14, "bold"),
         bg="#1A2B4A", fg="white").pack(pady=15)

# Cadre principal
frame = tk.Frame(root, bg="white", padx=20, pady=20)
frame.pack(padx=20, pady=5, fill="both", expand=True)

# Nom employé
tk.Label(frame, text="Nom de l'employé :", bg="white", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", pady=6)
entry_employe = tk.Entry(frame, width=30)
entry_employe.grid(row=0, column=1, pady=6)

# Canal
tk.Label(frame, text="Canal :", bg="white", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", pady=6)
combo_canal = ttk.Combobox(frame, values=["Appel", "QR Code"], state="readonly", width=27)
combo_canal.grid(row=1, column=1, pady=6)

# Périmètre
tk.Label(frame, text="Périmètre :", bg="white", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", pady=6)
combo_perimetre = ttk.Combobox(frame, values=["VIE", "IARD"], state="readonly", width=27)
combo_perimetre.grid(row=2, column=1, pady=6)

# NPS
tk.Label(frame, text="Score NPS (0 à 10) :", bg="white", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", pady=6)
entry_nps = tk.Entry(frame, width=30)
entry_nps.grid(row=3, column=1, pady=6)

# CSAT
tk.Label(frame, text="CSAT — Satisfait ?", bg="white", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="w", pady=6)
csat_var = tk.StringVar()
csat_frame = tk.Frame(frame, bg="white")
csat_frame.grid(row=4, column=1, sticky="w")
tk.Radiobutton(csat_frame, text="Oui", variable=csat_var, value="Oui", bg="white").pack(side="left")
tk.Radiobutton(csat_frame, text="Non", variable=csat_var, value="Non", bg="white").pack(side="left")

# CES
tk.Label(frame, text="CES — Effort fourni :", bg="white", font=("Helvetica", 10, "bold")).grid(row=5, column=0, sticky="w", pady=6)
ces_var = tk.StringVar()
ces_frame = tk.Frame(frame, bg="white")
ces_frame.grid(row=5, column=1, sticky="w")
tk.Radiobutton(ces_frame, text="Faible", variable=ces_var, value="Faible", bg="white").pack(side="left")
tk.Radiobutton(ces_frame, text="Modéré", variable=ces_var, value="Modéré", bg="white").pack(side="left")
tk.Radiobutton(ces_frame, text="Élevé", variable=ces_var, value="Élevé", bg="white").pack(side="left")

# Bouton enregistrer
tk.Button(root, text="✅ Enregistrer",
          command=enregistrer,
          bg="#C9A84C", fg="white",
          font=("Helvetica", 12, "bold"),
          width=20).pack(pady=15)

root.mainloop()