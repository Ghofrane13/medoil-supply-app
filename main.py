!pip install streamlit pandas openpyxl -q
import pandas as pd
import numpy as np

def moteur_remplissage_medoil(df_source):
    # --- 1. Nettoyage et Préparation des données Source ---
    df = df_source.copy()

    # S'assurer que les colonnes numériques sont bien lues
    colonnes_num = ['Consommation / mois', 'Coût de revient', 'Cout total annuelle']
    for col in colonnes_num:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    # --- 2. Calculs Supply Chain (Sur la Source) ---

    # A. Analyse ABC (Basée sur le Coût total annuel)
    df = df.sort_values(by='Cout total annuelle', ascending=False)
    df['Cumul_Valeur'] = df['Cout_Revient'].cumsum()
    total_valeur = df['Cout_Revient'].sum()
    df['%_Cumule'] = (df['Cumul_Valeur'] / total_valeur) * 100
    df['ABC_Calcule'] = df['%_Cumule'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))

    # B. Stock de Sécurité Dynamique (Exemple simplifié, à affiner avec l'écart-type)
    # Formule : Z * Conso_Moyenne * Lead_Time (LT pris par défaut ici)
    Z_SCORE = 1.65 # Pour 95%
    DEFAULT_LT_MOIS = 1 # 1 mois de délai par défaut
    df['SS_Calcule'] = Z_SCORE * df['Consommation / mois'] * np.sqrt(DEFAULT_LT_MOIS)

    # C. Point de Commande (Exemple)
    df['ROP_Calcule'] = (df['Consommation / mois'] * DEFAULT_LT_MOIS) + df['SS_Calcule']

    # --- 3. Création du Tableau Cible (Le modèle à remplir) ---

    # On sélectionne les colonnes à garder et on les renomme pour correspondre au modèle cible
    df_cible = df[['Code article', 'Description Article', 'Classe', 'ABC_Calcule', 'SS_Calcule', 'ROP_Calcule']].copy()
    df_cible.columns = ['Code article', 'Description Article', 'Classe', 'ABC', 'Stock de sécurité', 'point de commande']

    # On ajoute les colonnes que l'utilisateur remplira lui-même (vides)
    colonnes_vides = ['Fournisseur', 'Délai de livraison', 'l\'incertitude fournisseur', 'Délai de Sécurité',
                      'Taille de lot', 'MOQ', 'XYZ', 'EOQ', 'Diff MOQ et EOQ', 'Cout de stock de sécurité']
    for col in colonnes_vides:
        df_cible[col] = np.nan # Remplit de vide (Not a Number)

    # --- 4. Réorganisation finale pour correspondre à l'ordre exact de l'image Target ---
    ordre_final = ['Code article', 'Description Article', 'Classe', 'Fournisseur', 'Délai de livraison',
                   'l\'incertitude fournisseur', 'Délai de Sécurité', 'Taille de lot', 'MOQ', 'ABC',
                   'XYZ', 'EOQ', 'Diff MOQ et EOQ', 'Stock de sécurité', 'Cout de stock de sécurité', 'point de commande']
    df_cible = df_cible[ordre_final]

    return df_cible
import streamlit as st
import io

st.title("📦 Med Oil : Remplissage Automatique du Dashboard Supply")
st.write("Cet outil lit votre base de données Source et remplit votre tableau de bord Cible.")

uploaded_source = st.file_uploader("Importer votre Base de Données (Image 1)", type="xlsx")

if uploaded_source:
    df_src = pd.read_excel(uploaded_source)
    st.write("✅ Source chargée. Voici un aperçu :")
    st.dataframe(df_src.head())

    if st.button("▶️ Remplir le Tableau Cible"):
        resultats = moteur_remplissage_medoil(df_src)

        st.write("### 🏆 Tableau Target Rempli (Prêt pour import)")
        # Affichage interactif
        st.dataframe(resultats)

        # Bouton pour télécharger le fichier Excel résultant
        buffer = io.BytesIO()
        resultats.to_excel(buffer, index=False)
        st.download_button(label="📥 Télécharger le fichier Excel final", data=buffer.getvalue(),
                           file_name="target_medoil_rempli.xlsx", mime="application/vnd.ms-excel")
