# AI-OCR-Intelligent-Optical-Character-Recognition-System-PROJET-PFA-
 
 Plateforme d'Extraction OCR de Bilans ComptablesCe projet est un système intelligent de reconnaissance optique de caractères
 (OCR) conçu pour automatiser l'extraction, la structuration et la gestion des données issues de bilans comptables. 
Il a été développé dans le cadre d'un Projet de Fin d'Année (PFA) réalisé au sein de la BMCI (Groupe BNP Paribas)
 
 🎯 Problématique et Objectif:
 Le traitement des bilans comptables, souvent fournis en PDF natifs ou scannés, est un processus manuel, chronophage et sujet aux erreurs. 
 L'objectif de ce projet est de remédier à ces limites en proposant une plateforme intégrée qui automatise l'ensemble du processus.
 La solution développée permet de :Extraire automatiquement le contenu des PDF.
 Structurer les données en JSON normalisé (Actif, Passif, CPC).
 Convertir les résultats en fichiers Excel.
 Centraliser les informations dans une base de données MySQL via un workflow automatisé.
 
 📸 Captures d'écran de l'application
 Voici un aperçu des fonctionnalités clés de la plateforme :
 1. Interface Utilisateur (Upload)L'interface principale permet à l'utilisateur de sélectionner un client et de téléverser un fichier, en distinguant les PDF natifs des PDF scannés (images).
    <img width="959" height="437" alt="image" src="https://github.com/user-attachments/assets/4ac0bc84-99f5-4930-b201-86352fcf2486" />

 2.Module d'Administration (CRUD Clients)Un tableau de bord sécurisé (accessible après authentification ) permet aux administrateurs de gérer les clients (Créer, Lire, Mettre à jour, Supprimer).
 <img width="944" height="396" alt="image" src="https://github.com/user-attachments/assets/25d0f7a5-82ef-4fb9-8606-c3f9041abeb0" />
 <img width="947" height="430" alt="image" src="https://github.com/user-attachments/assets/3fc199f2-062b-459b-bfea-8dbefb0a37a9" />


 3. Résultat de l'extraction (Fichier Excel)Les données extraites sont structurées et exportées dans un fichier Excel, avec des onglets distincts pour le Bilan Actif, le Bilan Passif et le CPC.
    <img width="1567" height="880" alt="image" src="https://github.com/user-attachments/assets/b1921a16-a1c9-4d50-ace6-e6a771427fe2" />

    
 5. Workflow d'automatisation (n8n)Un workflow n8n automatise le processus de sauvegarde. Il est déclenché depuis l'interface et se charge de lire le fichier de résultat, d'extraire l'ID client et d'insérer les données dans la base MySQL.
<img width="959" height="431" alt="image" src="https://github.com/user-attachments/assets/b3322b99-7dd6-41bb-89f8-fa1276c55a05" />
<img width="553" height="842" alt="image" src="https://github.com/user-attachments/assets/1de6d7a6-0e6a-43cd-b31a-f22e931cbfd1" />
<img width="506" height="807" alt="image" src="https://github.com/user-attachments/assets/f7b5983b-ec82-4b19-8325-d0c7c2948bcc" />
<img width="542" height="842" alt="image" src="https://github.com/user-attachments/assets/7108dc62-e793-4266-bd29-8d87897bf48d" />



 7. Résultat en Base de Données (phpMyAdmin)Les données sont stockées dans la base MySQL et liées à un id_client , assurant la traçabilité et la centralisation de l'information.
 
 ✨ Fonctionnalités Principales:
 Double Prise en Charge (Natif vs Scanné) : Le système détecte automatiquement si un PDF est natif ou scanné pour appliquer le bon pipeline de traitement.
 Extraction Intelligente (LLM) : Utilisation de l'API Google Gemini (LLM) pour analyser sémantiquement les PDF scannés, reconnaître les rubriques comptables et structurer les données, surmontant les limites des OCR classiques.
 Structuration Automatique : Les données extraites sont mappées sur des modèles JSON prédéfinis (Bilan Actif, Bilan Passif, CPC) pour garantir la normalisation.
 Export Excel : Génération de fichiers .xlsx propres et organisés pour une manipulation facile par les équipes financières.
 Module Administrateur Sécurisé : Interface de gestion des clients (CRUD) 28avec authentification et hachage des mots de passe (bcrypt/scrypt) pour la sécurité.
 Automatisation du Workflow : Intégration de n8n pour automatiser la sauvegarde des données extraites du fichier Excel vers la base de données MySQL.
 
 🛠️ Stack Technique:
 Le projet combine plusieurs technologies modernes pour le backend, le frontend, l'IA et l'automatisation :DomaineTechnologieRôle dans le projetBackend Python Langage principal pour la logique métier et le traitement des données.
 Flask Micro-framework web pour créer l'API REST et servir l'interface.
 FrontendReactJS Bibliothèque JavaScript pour construire une interface utilisateur dynamique.
 HTML5 / CSS3 Structure et style des pages web.
 IA & OCRGoogle Gemini (LLM) Analyse sémantique et extraction de données à partir d'images (PDF scannés).
 Base de DonnéesMySQL SGBDR pour stocker de manière centralisée les données extraites et les informations clients.
 XAMPP / phpMyAdmin Environnement de développement local pour la gestion de la base MySQL.
 Data & WorkflowPandas Bibliothèque Python pour le nettoyage, la manipulation des données et la génération des fichiers Excel.
 n8n Outil d'automatisation (workflow automation) pour connecter l'application à la base de données.
 
 ⚙️ Architecture et WorkflowLe pipeline de traitement principal est le suivant:
 Upload : L'utilisateur se connecte, choisit un client et importe un PDF via l'interface Flask/React.
 Détection : Le backend Python détecte si le PDF est natif ou scanné.
 Traitement (Scanné) :Le PDF est converti en images.
 Les images sont envoyées à l'API Gemini (LLM).
 Gemini analyse le contenu et le compare aux templates JSON (Actif, Passif, CPC) pour extraire et structurer les données sémantiquement.
 Traitement (Natif) :Le texte et les tables sont extraits directement du PDF.
 Les données sont nettoyées à l'aide de Pandas.Export : Les données structurées sont converties en un fichier Excel.
 Sauvegarde (Workflow) :L'utilisateur clique sur "Sauvegarder dans la BDD.
 Un webhook déclenche le workflow n8n.
 n8n lit le fichier généré, le nettoie et insère les données dans les tables MySQL (bilan_actif, bilan_passif, bilan_cpc), en les liant à l'id_client approprié.
 
# Installer les dépendances
pip install -r requirements.txt
Configuration Frontend (ReactJS) :Bash# (En supposant un dossier 'frontend')
cd frontend
npm install
Base de Données (MySQL) :Lancez XAMPP (ou votre service MySQL).
Ouvrez phpMyAdmin et créez une nouvelle base de données.
Importez le schéma de la base de données (ex: schema.sql) dans votre BDD.
Automatisation (n8n) :Lancez votre instance n8n.
Importez le fichier workflow.json (s'il est fourni) dans n8n.
Mettez à jour le webhook et les identifiants de la base de données dans les nœuds n8n.
Variables d'environnement :Créez un fichier .env à la racine du backend.
Ajoutez vos clés d'API et identifiants de BDD :# Clé API pour Google Gemini
GEMINI_API_KEY="VOTRE_CLE_API_GEMINI"

# Identifiants BDD MySQL
DB_HOST="localhost"
DB_USER="root"
DB_PASSWORD=""
DB_NAME="votre_nom_de_bdd"
Lancer l'application :Bash# Terminal 1: Lancer le backend Flask
python app.py

# Terminal 2: Lancer le frontend React
cd frontend
npm start
J'espère que ce README vous aidera à bien présenter votre projet !

Auteur : BENALLOU Dalal

