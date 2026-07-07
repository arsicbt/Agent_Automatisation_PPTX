import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import zipfile

# 1. PARAMÉTRAGE DE L'INTERFACE WEB POUR STREAMLIT
st.set_page_config(page_title="Générateur de Rapports PPTX", page_icon="🚀", layout="centered")

st.title("Générateur Automatique de Rapports PPTX")
st.write("Cette interface permet de générer des présentations PowerPoint personnalisées par lot à partir d'un modèle et d'un tableau de données.")

# Zone d'explications pour l'utilisateur à distance
st.info("""
**Comment tester ?**
1. Téléversez un modèle PowerPoint (.pptx) contenant des balises comme `{{NOM_CLIENT}}`, `{{CHIFFRE_AFFAIRES}}`, ou `{{OBJECTIF}}`.
2. Téléversez votre fichier de données au format CSV.
3. Cliquez sur 'Lancer la génération' pour récupérer vos rapports personnalisés !
""")

# 2. MODULES DE TÉLÉCHARGEMENT (UPLOAD)
uploaded_template = st.file_uploader("1. Choisissez votre fichier Template PPTX", type="pptx")
uploaded_csv = st.file_uploader("2. Choisissez votre fichier de données CSV", type="csv")

# 3. MOTEUR D'INJECTION ET GÉNÉRATION
if uploaded_template and uploaded_csv:
    if st.button("Lancer la génération des rapports"):
        
        # Lecture du tableau CSV grâce à Pandas
        df = pd.read_csv(uploaded_csv)
        
        # Création d'un dossier ZIP virtuel en mémoire pour stocker tous les PPTX
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            
            # La Boucle : On traite le tableau ligne par ligne
            for index, row in df.iterrows():
                
                # On recharge une copie propre du template depuis la mémoire à chaque ligne
                template_bytes = uploaded_template.getvalue()
                prs = Presentation(io.BytesIO(template_bytes))
                
                # Parcours de toutes les slides et de toutes les formes textuelles
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    
                                    # REMPLACEMENT DYNAMIQUE DES BALISES
                                    if "{{NOM_CLIENT}}" in run.text:
                                        run.text = run.text.replace("{{NOM_CLIENT}}", str(row['Client']))
                                    if "{{CHIFFRE_AFFAIRES}}" in run.text:
                                        run.text = run.text.replace("{{CHIFFRE_AFFAIRES}}", f"{row['CA']:,} €")
                                    if "{{OBJECTIF}}" in run.text:
                                        run.text = run.text.replace("{{OBJECTIF}}", f"{row['Objectif']:,} €")
                
                # Sauvegarde du fichier PowerPoint individuel en mémoire
                pptx_buffer = io.BytesIO()
                prs.save(pptx_buffer)
                pptx_buffer.seek(0)
                
                # Ajout du PowerPoint dans le fichier ZIP final
                nom_fichier = f"Rapport_{str(row['Client'])}.pptx".replace(" ", "_")
                zip_file.writestr(nom_fichier, pptx_buffer.getvalue())
        
        # 4. DISPOSITIF DE TÉLÉCHARGEMENT DU RÉSULTAT
        st.success("Tous les rapports ont été générés avec succès !")
        
        st.download_button(
            label="⬇Télécharger le package de rapports (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="tous_les_rapports_generes.zip",
            mime="application/zip"
        )
