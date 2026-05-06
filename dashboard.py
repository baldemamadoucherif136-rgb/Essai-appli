import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# ─── Couleurs SUNU ────────────────────────────────────────
BLEU = "#1A2B4A"
OR = "#C9A84C"
BLANC = "white"

def charger_fichier():
    chemin = filedialog.askopenfilename(
        title="Choisir le fichier de collecte",
        filetypes=[("Fichiers Excel", "*.xlsx")]
    )
    if chemin:
        entry_fichier.delete(0, tk.END)
        entry_fichier.insert(0, chemin)

def calculer_indicateurs(df, perimetre):
    if perimetre != "GLOBAL":
        data = df[df["Périmètre"] == perimetre]
    else:
        data = df

    total = len(data)
    if total == 0:
        return None

    promoteurs = len(data[data["Catégorie NPS"] == "Promoteur"])
    detracteurs = len(data[data["Catégorie NPS"] == "Détracteur"])
    nps = round((promoteurs / total * 100) - (detracteurs / total * 100), 1)

    satisfaits = len(data[data["CSAT"] == "Oui"])
    csat = round(satisfaits / total * 100, 1)

    effort_eleve = len(data[data["CES"] == "Élevé"])
    effort_faible = len(data[data["CES"] == "Faible"])
    ces = round((effort_eleve / total * 100) - (effort_faible / total * 100), 1)

    return {"NPS": nps, "CSAT": csat, "CES": ces, "Total": total}

def afficher_dashboard():
    chemin = entry_fichier.get().strip()
    if not chemin or not os.path.exists(chemin):
        messagebox.showwarning("Attention", "Veuillez sélectionner un fichier Excel valide.")
        return

    df = pd.read_excel(chemin, sheet_name="Collecte")

    # Nettoyage ancien dashboard
    for widget in frame_dashboard.winfo_children():
        widget.destroy()

    # Calcul par périmètre
    perimetres = ["VIE", "IARD", "GLOBAL"]
    resultats = {}
    for p in perimetres:
        res = calculer_indicateurs(df, p)
        if res:
            resultats[p] = res

    if not resultats:
        messagebox.showwarning("Attention", "Aucune donnée disponible.")
        return

    labels = list(resultats.keys())
    nps_vals = [resultats[p]["NPS"] for p in labels]
    csat_vals = [resultats[p]["CSAT"] for p in labels]
    ces_vals = [resultats[p]["CES"] for p in labels]

    # ─── Figure matplotlib ────────────────────────────────
    fig, axes = plt.subplots(1, 3, figsize=(12, 4))
    fig.patch.set_facecolor("#F5F5F5")

    couleurs = ["#2E6DA4", "#C9A84C", "#1A2B4A"]

    # NPS
    bars = axes[0].bar(labels, nps_vals, color=couleurs)
    axes[0].set_title("NPS (%)", fontsize=13, fontweight="bold", color=BLEU)
    axes[0].axhline(y=50, color="red", linestyle="--", label="Seuil 50%")
    axes[0].legend(fontsize=8)
    axes[0].set_ylim(-100, 100)
    for bar, val in zip(bars, nps_vals):
        axes[0].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                     f"{val}%", ha="center", fontsize=10, fontweight="bold")

    # CSAT
    bars = axes[1].bar(labels, csat_vals, color=couleurs)
    axes[1].set_title("CSAT (%)", fontsize=13, fontweight="bold", color=BLEU)
    axes[1].set_ylim(0, 120)
    for bar, val in zip(bars, csat_vals):
        axes[1].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                     f"{val}%", ha="center", fontsize=10, fontweight="bold")

    # CES
    bars = axes[2].bar(labels, ces_vals, color=couleurs)
    axes[2].set_title("CES (%)", fontsize=13, fontweight="bold", color=BLEU)
    axes[2].axhline(y=0, color="red", linestyle="--", label="Seuil 0%")
    axes[2].legend(fontsize=8)
    axes[2].set_ylim(-100, 100)
    for bar, val in zip(bars, ces_vals):
        axes[2].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                     f"{val}%", ha="center", fontsize=10, fontweight="bold")

    plt.tight_layout()

    # Intégration dans tkinter
    canvas = FigureCanvasTkAgg(fig, master=frame_dashboard)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    # Bouton export
    tk.Button(frame_dashboard,
              text="📥 Exporter le dashboard (image)",
              command=lambda: exporter_image(fig),
              bg=OR, fg=BLANC,
              font=("Helvetica", 11, "bold"),
              width=30).pack(pady=10)

def exporter_image(fig):
    chemin = filedialog.asksaveasfilename(
        title="Enregistrer le dashboard",
        defaultextension=".png",
        filetypes=[("Image PNG", "*.png")],
        initialfile="SUNU_Dashboard.png"
    )
    if chemin:
        fig.savefig(chemin, dpi=150, bbox_inches="tight")
        messagebox.showinfo("Succès", f"Dashboard exporté ✅\n{chemin}")

# ─── Interface ────────────────────────────────────────────
root = tk.Tk()
root.title("SUNU Assurances — Dashboard Relation Client")
root.geometry("900x650")
root.resizable(False, False)
root.configure(bg=BLEU)

# Titre
tk.Label(root, text="SUNU Business — Dashboard Relation Client",
         font=("Helvetica", 14, "bold"),
         bg=BLEU, fg=BLANC).pack(pady=15)

# Sélection fichier
frame_fichier = tk.Frame(root, bg=BLANC, padx=15, pady=15)
frame_fichier.pack(padx=20, pady=5, fill="x")

tk.Label(frame_fichier, text="Fichier de collecte :",
         bg=BLANC, font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w")
entry_fichier = tk.Entry(frame_fichier, width=55)
entry_fichier.grid(row=0, column=1, padx=10)
tk.Button(frame_fichier, text="📂 Parcourir",
          command=charger_fichier,
          bg=BLEU, fg=BLANC,
          font=("Helvetica", 10)).grid(row=0, column=2)

# Bouton afficher
tk.Button(root, text="📊 Afficher le Dashboard",
          command=afficher_dashboard,
          bg=OR, fg=BLANC,
          font=("Helvetica", 12, "bold"),
          width=25).pack(pady=10)

# Zone dashboard
frame_dashboard = tk.Frame(root, bg=BLANC)
frame_dashboard.pack(padx=20, pady=5, fill="both", expand=True)

root.mainloop()