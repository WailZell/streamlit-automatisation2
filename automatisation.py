import streamlit as st
import pandas as pd
from io import BytesIO

def read_excel_sheets(template_path):
    try:
        sites_adresses = pd.read_excel(template_path, sheet_name='Liste des sites avec adresses')
        utilisateurs_clients = pd.read_excel(template_path, sheet_name='Liste des utilisateurs clients')
        return sites_adresses, utilisateurs_clients
    except Exception as e:
        st.error(f"Erreur lors de la lecture des fichiers Excel : {e}")
        return None, None

def check_duplicates_and_missing_values(sites_adresses, utilisateurs_clients):
    sites_adresses_no_duplicates = sites_adresses.drop_duplicates()
    sites_adresses_no_duplicates['is_duplicate'] = sites_adresses_no_duplicates.duplicated(subset=['CGR Chantier'], keep=False)
    
    utilisateurs_clients_no_duplicates = utilisateurs_clients.drop_duplicates(subset=['Mail'])
    utilisateurs_clients_no_duplicates['is_duplicate'] = utilisateurs_clients_no_duplicates.duplicated(subset=['Mail'], keep=False)
    
    sites_missing_info = sites_adresses_no_duplicates[sites_adresses_no_duplicates.isnull().any(axis=1)]
    utilisateurs_missing_info = utilisateurs_clients_no_duplicates[utilisateurs_clients_no_duplicates.isnull().any(axis=1)]

    return sites_adresses_no_duplicates, utilisateurs_clients_no_duplicates, sites_missing_info, utilisateurs_missing_info

def save_to_excel(dataframe, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

def process_file(template_path):
    sites_adresses, utilisateurs_clients = read_excel_sheets(template_path)
    if sites_adresses is None or utilisateurs_clients is None:
        return None
    
    sites_adresses_no_duplicates, utilisateurs_clients_no_duplicates, sites_missing_info, utilisateurs_missing_info = check_duplicates_and_missing_values(sites_adresses, utilisateurs_clients)

    if sites_adresses_no_duplicates['is_duplicate'].any() or utilisateurs_clients_no_duplicates['is_duplicate'].sum() > 0:
        st.warning("Des doublons ont été trouvés ou des données manquent.")
        if sites_adresses_no_duplicates['is_duplicate'].any():
            st.warning("Doublons trouvés dans CGR Chantier")
        if utilisateurs_clients_no_duplicates['is_duplicate'].sum() > 0:
            st.warning(f"Nombre de doublons de mails: {utilisateurs_clients_no_duplicates['is_duplicate'].sum()}")
        return None
    
    creer_contacts = create_contacts_dataframe(utilisateurs_clients_no_duplicates)
    creer_sites = create_sites_dataframe(sites_adresses_no_duplicates)

    return (
        save_to_excel(sites_missing_info, 'Sites Missing Info'),
        save_to_excel(utilisateurs_missing_info, 'Contacts Missing Info'),
        save_to_excel(creer_contacts, 'Creer Contacts Resultat'),
        save_to_excel(creer_sites, 'Creer Sites Resultat')
    )

def create_contacts_dataframe(utilisateurs_clients_no_duplicates):
    creer_contacts = pd.DataFrame()
    required_columns_contacts = ['Email', 'Contact_Type__c', 'LastName', 'FirstName', 'region', 'ID Contact', 'Site__c', 'AccountId', 'CGR Chantier']
    for col in required_columns_contacts:
        if col not in creer_contacts.columns:
            creer_contacts[col] = None

    creer_contacts['Email'] = utilisateurs_clients_no_duplicates['Mail']
    creer_contacts['Contact_Type__c'] = utilisateurs_clients_no_duplicates['Type (Donneurs d\'ordre ou Site)'].replace(
        {'Site': 'portal user site', 'Donneurs d\'ordre': 'portal user'})
    creer_contacts['LastName'] = utilisateurs_clients_no_duplicates['Nom']
    creer_contacts['FirstName'] = utilisateurs_clients_no_duplicates['Prénom']
    creer_contacts['region'] = utilisateurs_clients_no_duplicates['Périmètre des sites']
    creer_contacts['ID Contact'] = utilisateurs_clients_no_duplicates['Nom'] + ' ' + utilisateurs_clients_no_duplicates['Prénom']
    creer_contacts['CGR Chantier'] = utilisateurs_clients_no_duplicates.get('CGR Chantier', None)
    creer_contacts['Site__c'] = None
    creer_contacts['AccountId'] = None

    non_empty_columns_contacts = [col for col in creer_contacts.columns if creer_contacts[col].notna().any() or col in ['Site__c', 'AccountId']]
    return creer_contacts[non_empty_columns_contacts]

def create_sites_dataframe(sites_adresses_no_duplicates):
    creer_sites = pd.DataFrame()
    required_columns_sites = ['Zip_Postal_code__c', 'City__c', 'Street__c', 'Name', 'Operating_Site__c', 'Country2__c', 'Customer_Site_ID__c', 'CGR Chantier', 'region', 'ID Contact', 'Account__c', 'Accounting_System_Site__c']
    for col in required_columns_sites:
        if col not in creer_sites.columns:
            creer_sites[col] = None

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

    non_empty_columns_sites = [col for col in creer_sites.columns if creer_sites[col].notna().any() or col in ['Account__c', 'Accounting_System_Site__c']]
    return creer_sites[non_empty_columns_sites]

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
