import streamlit as st
import pandas as pd
from io import BytesIO

def process_file(template_path):
    try:
        # Lire les feuilles nécessaires
        sites_adresses = pd.read_excel(template_path, sheet_name='Liste des sites avec adresses')
        utilisateurs_clients = pd.read_excel(template_path, sheet_name='Liste des utilisateurs clients')

        # Vérifier les doublons dans les adresses des sites
        sites_adresses_no_duplicates = sites_adresses.drop_duplicates()
        doublons_cgr = sites_adresses_no_duplicates.duplicated(subset=['CGR Chantier'], keep=False)

        # Marquer les doublons dans les adresses des sites
        sites_adresses_no_duplicates['is_duplicate'] = sites_adresses_no_duplicates.duplicated(subset=['CGR Chantier'], keep=False)

        # Vérifier les doublons dans les utilisateurs clients
        utilisateurs_clients_no_duplicates = utilisateurs_clients.drop_duplicates(subset=['Mail'])
        doublons_mail = utilisateurs_clients_no_duplicates.duplicated(subset=['Mail']).sum()

        # Marquer les doublons dans les utilisateurs clients
        utilisateurs_clients_no_duplicates['is_duplicate'] = utilisateurs_clients_no_duplicates.duplicated(subset=['Mail'], keep=False)

        # Fonction pour vérifier et ajouter les colonnes manquantes
        def ensure_columns(df, required_columns):
            for col in required_columns:
                if col not in df.columns:
                    df[col] = None
            return df

        # Identifier les sites avec des informations manquantes
        required_columns_sites = ['Zip/Postal code', 'Ville', 'Adresse', 'Nom du site (CHANTIER)', 'CGR Chantier', 'N° Mag. (facultatif)', 'LOT / REGIONS']
        sites_with_missing_info = sites_adresses_no_duplicates[required_columns_sites].isnull().any(axis=1)
        sites_missing_info = sites_adresses_no_duplicates[sites_with_missing_info]

        # Compter le nombre de valeurs manquantes par site
        sites_adresses_no_duplicates['missing_values'] = sites_adresses_no_duplicates.isnull().sum(axis=1)
        # Trier les sites par nombre de valeurs manquantes en ordre décroissant
        sites_most_missing_info = sites_adresses_no_duplicates.sort_values(by='missing_values', ascending=False)

        # Sauvegarder les sites avec le plus d'informations manquantes dans un nouveau fichier
        sites_most_missing_info_io = BytesIO()
        sites_most_missing_info.to_excel(sites_most_missing_info_io, index=False)
        sites_most_missing_info_io.seek(0)

        # Vérifier l'absence de doublons avant de procéder
        if not doublons_cgr.any() and doublons_mail == 0:
            # Lire le modèle de fichier "Créer contacts en masse"
            creer_contacts = pd.DataFrame()

            # Assurer que toutes les colonnes nécessaires sont présentes
            required_columns_contacts = ['Email', 'Contact_Type__c', 'LastName', 'FirstName', 'region', 'ID Contact', 'Site__c', 'AccountId', 'CGR Chantier']
            creer_contacts = ensure_columns(creer_contacts, required_columns_contacts)
            
            # Transférer les informations vérifiées dans le fichier des contacts
            creer_contacts['Email'] = utilisateurs_clients_no_duplicates['Mail']
            creer_contacts['Contact_Type__c'] = utilisateurs_clients_no_duplicates['Type (Donneurs d\'ordre ou Site)'].replace({'Site': 'portal user site', 'Donneurs d\'ordre': 'portal user'})
            creer_contacts['LastName'] = utilisateurs_clients_no_duplicates['Nom']
            creer_contacts['FirstName'] = utilisateurs_clients_no_duplicates['Prénom']
            creer_contacts['region'] = utilisateurs_clients_no_duplicates['Périmètre des sites']
            creer_contacts['ID Contact'] = utilisateurs_clients_no_duplicates['Nom'] + ' ' + utilisateurs_clients_no_duplicates['Prénom']
            creer_contacts['CGR Chantier'] = utilisateurs_clients_no_duplicates['CGR Chantier'] if 'CGR Chantier' in utilisateurs_clients_no_duplicates.columns else None

            # Compter le nombre de valeurs manquantes par contact
            utilisateurs_clients_no_duplicates['missing_values'] = utilisateurs_clients_no_duplicates.isnull().sum(axis=1)
            # Trier les contacts par nombre de valeurs manquantes en ordre décroissant
            contacts_most_missing_info = utilisateurs_clients_no_duplicates.sort_values(by='missing_values', ascending=False)

            # Sauvegarder les contacts avec le plus d'informations manquantes dans un nouveau fichier
            contacts_most_missing_info_io = BytesIO()
            contacts_most_missing_info.to_excel(contacts_most_missing_info_io, index=False)
            contacts_most_missing_info_io.seek(0)

            # Vérifier que les seules colonnes vides sont 'Site__c' et 'AccountId'
            creer_contacts['Site__c'] = None
            creer_contacts['AccountId'] = None

            # Supprimer toutes les autres colonnes vides
            non_empty_columns_contacts = [col for col in creer_contacts.columns if creer_contacts[col].notna().any() or col in ['Site__c', 'AccountId']]
            creer_contacts = creer_contacts[non_empty_columns_contacts]

            # Sauvegarder le fichier des contacts
            creer_contacts_io = BytesIO()
            creer_contacts.to_excel(creer_contacts_io, index=False)
            creer_contacts_io.seek(0)
            
            # Lire le modèle de fichier "Créer sites en masse"
            creer_sites = pd.DataFrame()

            # Assurer que toutes les colonnes nécessaires sont présentes
            required_columns_sites_final = ['Zip_Postal_code__c', 'City__c', 'Street__c', 'Name', 'Operating_Site__c', 'Country2__c', 'Customer_Site_ID__c', 'CGR Chantier', 'region', 'ID Contact', 'Account__c', 'Accounting_System_Site__c']
            creer_sites = ensure_columns(creer_sites, required_columns_sites_final)
            
            # Transférer les informations vérifiées dans le fichier des sites
            creer_sites['Zip_Postal_code__c'] = sites_adresses_no_duplicates['Zip/Postal code']
            creer_sites['City__c'] = sites_adresses_no_duplicates['Ville']
            creer_sites['Street__c'] = sites_adresses_no_duplicates['Adresse']
            creer_sites['Name'] = sites_adresses_no_duplicates['Nom du site (CHANTIER)']
            creer_sites['Operating_Site__c'] = 'TRUE'
            creer_sites['Country2__c'] = 'FR'
            creer_sites['Customer_Site_ID__c'] = sites_adresses_no_duplicates['N° Mag. (facultatif)']
            creer_sites['CGR Chantier'] = sites_adresses_no_duplicates['CGR Chantier']
            creer_sites['region'] = sites_adresses_no_duplicates['LOT / REGIONS']
            creer_sites['ID Contact'] = sites_adresses_no_duplicates['Nom du compte (sur Salesforce)'] + ' ' + sites_adresses_no_duplicates['Nom du site (CHANTIER)']

            # Supprimer toutes les autres colonnes vides sauf 'Account__c' et 'Accounting_System_Site__c'
            non_empty_columns_sites = [col for col in creer_sites.columns if creer_sites[col].notna().any() or col in ['Account__c', 'Accounting_System_Site__c']]
            creer_sites = creer_sites[non_empty_columns_sites]

            # Sauvegarder le fichier des sites
            creer_sites_io = BytesIO()
            creer_sites.to_excel(creer_sites_io, index=False)
            creer_sites_io.seek(0)

            return (sites_most_missing_info_io, contacts_most_missing_info_io, creer_contacts_io, creer_sites_io)

        else:
            st.warning("Des doublons ont été trouvés ou des données manquent.")
            if doublons_cgr.any():
                st.warning("Doublons trouvés dans CGR Chantier")
            if doublons_mail > 0:
                st.warning(f"Nombre de doublons de mails: {doublons_mail}")
            return None
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        return None

# Interface utilisateur Streamlit
st.title("Traitement de Template de Données")

uploaded_file = st.file_uploader("Sélectionnez le fichier template à traiter", type=["xlsx", "xls"])

if uploaded_file is not None:
    result = process_file(uploaded_file)

    if result is not None:
        sites_most_missing_info_io, contacts_most_missing_info_io, creer_contacts_io, creer_sites_io = result
        
        st.download_button(
            label="Télécharger Sites plus infos manquantes",
            data=sites_most_missing_info_io,
            file_name="Sites_plus_infos_manquantes.xlsx"
        )
        
        st.download_button(
            label="Télécharger Contacts plus infos manquantes",
            data=contacts_most_missing_info_io,
            file_name="Contacts_plus_infos_manquantes.xlsx"
        )
        
        st.download_button(
            label="Télécharger Créer contacts en masse résultat",
            data=creer_contacts_io,
            file_name="Creer_contacts_en_masse_resultat.xlsx"
        )
        
        st.download_button(
            label="Télécharger Créer sites en masse résultat",
            data=creer_sites_io,
            file_name="Creer_sites_en_masse_resultat.xlsx"
        )
