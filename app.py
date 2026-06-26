import streamlit as st
import json, csv, io
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime

st.set_page_config(
    page_title="PL-300 Power BI · Checklist",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Force white background, clean look
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #ffffff; }
[data-testid="stHeader"] { background: #ffffff; }
[data-testid="stSidebar"] { display: none; }
.block-container { padding-top: 2rem; padding-bottom: 4rem; max-width: 680px; }
div[data-testid="stProgress"] > div { background-color: #16a34a; }
</style>
""", unsafe_allow_html=True)

# ─── MODULES DATA ─────────────────────────────────────────────────────────────
MODULES = [
    {
        "id": "m1", "emoji": "🧐", "title": "Introduction à Power BI",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Expliquer ce qu'est la Business Intelligence et pourquoi les entreprises en ont besoin"},
            {"level": "L1", "type": "theory",   "text": "Décrire les trois composants Power BI : Desktop, Service et Mobile — et à quoi sert chacun"},
            {"level": "L1", "type": "theory",   "text": "Savoir quand Power BI est le bon choix et quand un autre outil serait préférable"},
            {"level": "L1", "type": "practice", "text": "Naviguer dans l'interface Desktop et identifier les trois vues principales"},
            {"level": "L1", "type": "practice", "text": "Identifier la différence entre un rapport et un tableau de bord dans le Service"},
            {"level": "L2", "type": "theory",   "text": "Expliquer le rôle du semantic model comme couche centrale entre les données et les rapports dans le Service"},
            {"level": "L2", "type": "theory",   "text": "Comparer Power BI et Tableau : avantages de chacun selon le contexte (intégration M365, flexibilité de visualisation)"},
            {"level": "L2", "type": "theory",   "text": "Comparer Power BI et Looker : modélisation LookML côté serveur vs self-service analytics Power BI"},
            {"level": "L2", "type": "theory",   "text": "Identifier les limites concrètes de Power BI : pas de temps réel natif sans Premium, taille .pbix, pas de co-édition simultanée"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les licences : Pro requis pour partager dans un workspace partagé, Premium Per User vs Premium Capacity"},
            {"level": "L2", "type": "practice", "text": "Localiser dans le Service : Mon Espace de Travail, un workspace partagé, un rapport publié, un semantic model"},
            {"level": "L2", "type": "practice", "text": "Associer précisément chaque rôle métier (analyste, manager, DSI) à son composant Power BI et son niveau d'accès"},
        ]
    },
    {
        "id": "m2", "emoji": "🛠️", "title": "Prérequis & Préparation",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Distinguer OLTP (base opérationnelle, écritures fréquentes) et OLAP (base analytique, lectures agrégées)"},
            {"level": "L1", "type": "theory",   "text": "Expliquer ce qu'est un schéma en étoile et pourquoi Power BI est optimisé pour ce modèle"},
            {"level": "L1", "type": "practice", "text": "Reconnaître un schéma en étoile dans un diagramme : table de faits au centre, dimensions en branches"},
            {"level": "L1", "type": "practice", "text": "Écrire une requête SQL SELECT avec JOIN pour extraire des données de deux tables"},
            {"level": "L2", "type": "theory",   "text": "Décrire ce qu'est un data warehouse et le processus ETL/ELT qui l'alimente"},
            {"level": "L2", "type": "theory",   "text": "Expliquer pourquoi importer des données OLTP normalisées directement dans Power BI est une mauvaise pratique"},
            {"level": "L2", "type": "theory",   "text": "Identifier les concepts SQL essentiels pour Power BI : SELECT, FROM, WHERE, JOIN, GROUP BY, agrégats — et ce que Power Query peut remplacer"},
            {"level": "L2", "type": "practice", "text": "Écrire une requête SQL avec GROUP BY et SUM/COUNT pour pré-agréger avant import"},
            {"level": "L2", "type": "practice", "text": "Distinguer schéma en étoile vs schéma en flocon et expliquer pourquoi le flocon est déconseillé dans Power BI"},
            {"level": "L2", "type": "practice", "text": "Configurer un accès à Google BigQuery ou un équivalent cloud comme source Power BI"},
        ]
    },
    {
        "id": "m3", "emoji": "📂", "title": "Sources de données & Modes de stockage",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Distinguer le mode Import (copie en mémoire, performances max) et DirectQuery (requêtes temps réel, données fraîches)"},
            {"level": "L1", "type": "theory",   "text": "Savoir quand utiliser une passerelle (Gateway) pour connecter le Service à des sources locales"},
            {"level": "L1", "type": "practice", "text": "Connecter Power BI à une source SQL Server via Get Data et choisir Import ou DirectQuery"},
            {"level": "L1", "type": "practice", "text": "Connecter un fichier Excel ou CSV et inspecter les données dans Power Query avant de charger"},
            {"level": "L2", "type": "theory",   "text": "Expliquer le mode Dual (composite model) : une table peut se comporter en Import ou DirectQuery selon le contexte"},
            {"level": "L2", "type": "theory",   "text": "Distinguer Gateway Personal Mode vs Gateway Standard Mode : cas d'usage, limites, support DirectQuery"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les niveaux de confidentialité (Privé, Organisationnel, Public) et leur rôle dans la combinaison de sources"},
            {"level": "L2", "type": "theory",   "text": "Expliquer l'erreur Formula.Firewall : cause, conséquence et deux solutions possibles"},
            {"level": "L2", "type": "practice", "text": "Connecter une source SharePoint Online avec authentification Microsoft 365"},
            {"level": "L2", "type": "practice", "text": "Définir le mode de stockage d'une table dans les propriétés de table (vue Modèle)"},
            {"level": "L2", "type": "practice", "text": "Configurer les niveaux de confidentialité dans Fichier > Options > Confidentialité"},
            {"level": "L2", "type": "practice", "text": "Inspecter les types de colonnes, valeurs nulles et statistiques dans Power Query Editor (Affichage > Aperçu des données)"},
        ]
    },
    {
        "id": "m4", "emoji": "🏗️", "title": "Conception & Implémentation du modèle de données",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Décrire le schéma en étoile : table de faits centrale (mesures + clés) entourée de tables de dimension (attributs)"},
            {"level": "L1", "type": "theory",   "text": "Expliquer la cardinalité Un-à-Plusieurs (1:*) : standard dans un schéma en étoile — une dimension, plusieurs faits"},
            {"level": "L1", "type": "theory",   "text": "Savoir pourquoi une table de dates dédiée est nécessaire pour les calculs temporels"},
            {"level": "L1", "type": "practice", "text": "Créer une relation entre deux tables dans la vue Modèle et vérifier la cardinalité détectée"},
            {"level": "L1", "type": "practice", "text": "Créer une table de dates et la marquer comme table de dates via Outils de table"},
            {"level": "L1", "type": "practice", "text": "Corriger le tri alphabétique des mois avec 'Trier par colonne' (MonthName → MonthNumber)"},
            {"level": "L2", "type": "theory",   "text": "Expliquer pourquoi le schéma en flocon est déconseillé : jointures supplémentaires, DAX plus complexe, performances dégradées"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la cardinalité Plusieurs-à-Plusieurs (*:*) : risques de doublons de calcul et solution via table de pont (bridge table)"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la direction de filtre Single vs Both : Single = filtre de 'un' vers 'plusieurs' uniquement ; Both = bidirectionnel avec risque d'ambiguïté"},
            {"level": "L2", "type": "theory",   "text": "Expliquer quand utiliser Both : cas légitimes (tables de sécurité, comptages inversés) et dangers (filtres circulaires, résultats inattendus avec CALCULATE)"},
            {"level": "L2", "type": "theory",   "text": "Expliquer relations actives vs inactives : une seule relation active possible entre deux tables — les autres sont en pointillés"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les dimensions à rôles multiples : même table de dates reliée via OrderDate ET ShipDate — une active, l'autre inactive"},
            {"level": "L2", "type": "theory",   "text": "Expliquer USERELATIONSHIP() : active temporairement une relation inactive dans CALCULATE — ex : CALCULATE([Ventes], USERELATIONSHIP(DimDate[Date], FactVentes[ShipDate]))"},
            {"level": "L2", "type": "practice", "text": "Corriger une relation *:* en créant une table de pont : valeurs distinctes communes → deux relations 1:*"},
            {"level": "L2", "type": "practice", "text": "Créer deux relations DimDate/FactVentes (OrderDate + ShipDate), désactiver l'une, écrire une mesure avec USERELATIONSHIP()"},
            {"level": "L2", "type": "practice", "text": "Définir l'agrégation par défaut d'une colonne ID à 'Ne pas résumer' pour éviter les sommes d'identifiants"},
            {"level": "L2", "type": "practice", "text": "Créer une hiérarchie Année > Trimestre > Mois > Jour et la tester avec drill-down dans un visuel"},
            {"level": "L2", "type": "practice", "text": "Masquer les colonnes techniques (clés étrangères, colonnes de tri) dans le volet Données"},
            {"level": "L2", "type": "practice", "text": "Organiser les mesures dans des dossiers d'affichage par catégorie (ex : 'Ventes', 'Budget')"},
            {"level": "L2", "type": "practice", "text": "Implémenter une RLS statique : créer un rôle avec filtre DAX, tester avec 'Afficher en tant que rôles'"},
        ]
    },
    {
        "id": "m5", "emoji": "🧮", "title": "Fondamentaux DAX",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Distinguer mesure et colonne calculée : la mesure est dynamique (répond aux filtres), la colonne est fixe (calculée au refresh)"},
            {"level": "L1", "type": "theory",   "text": "Expliquer le contexte de filtre : l'ensemble des filtres actifs qui déterminent ce qu'une mesure calcule dans une cellule"},
            {"level": "L1", "type": "theory",   "text": "Expliquer CALCULATE : la seule fonction DAX qui peut modifier le contexte de filtre"},
            {"level": "L1", "type": "practice", "text": "Écrire les mesures de base : SUM, AVERAGE, COUNT, COUNTROWS, MIN, MAX"},
            {"level": "L1", "type": "practice", "text": "Écrire une mesure avec CALCULATE pour filtrer sur une valeur spécifique"},
            {"level": "L1", "type": "practice", "text": "Écrire une mesure conditionnelle avec IF et une classification avec SWITCH"},
            {"level": "L2", "type": "theory",   "text": "Distinguer mesures implicites vs explicites : les implicites ne se réutilisent pas — toujours préférer les explicites"},
            {"level": "L2", "type": "theory",   "text": "Expliquer le contexte de ligne : présent dans les colonnes calculées et les itérateurs (SUMX), absent dans les mesures seules"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la transition de contexte : CALCULATE dans un itérateur convertit le contexte de ligne en contexte de filtre"},
            {"level": "L2", "type": "theory",   "text": "Expliquer ALL() : supprime les filtres d'une table ou colonne — utile pour le % du total"},
            {"level": "L2", "type": "theory",   "text": "Expliquer FILTER() dans CALCULATE : filtre table dynamique, plus puissant mais plus coûteux que le filtre colonne"},
            {"level": "L2", "type": "theory",   "text": "Expliquer ALLEXCEPT() : supprime tous les filtres sauf les colonnes spécifiées — utile pour les totaux partiels"},
            {"level": "L2", "type": "theory",   "text": "Expliquer SUMX vs SUM : SUM additionne une colonne existante ; SUMX calcule une expression ligne par ligne puis agrège"},
            {"level": "L2", "type": "practice", "text": "Écrire % du Total = DIVIDE([Ventes], CALCULATE([Ventes], ALL(DimProduit)))"},
            {"level": "L2", "type": "practice", "text": "Écrire Marge Totale = SUMX(FactVentes, FactVentes[Qte] * (FactVentes[Prix] - FactVentes[Cout]))"},
            {"level": "L2", "type": "practice", "text": "Écrire une mesure dynamique avec SELECTEDVALUE pour adapter un titre de visuel à la sélection du slicer"},
            {"level": "L2", "type": "practice", "text": "Utiliser COALESCE pour remplacer les valeurs BLANK par zéro"},
            {"level": "L2", "type": "practice", "text": "Utiliser SUMMARIZE pour créer une table virtuelle résumée et la tester dans un visuel"},
        ]
    },
    {
        "id": "m6", "emoji": "⏳", "title": "Time Intelligence & Fonctions statistiques",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Expliquer pourquoi une table de dates marquée est indispensable pour toute fonction Time Intelligence"},
            {"level": "L1", "type": "theory",   "text": "Distinguer YTD (depuis le 1er janvier), QTD (depuis le début du trimestre) et MTD (depuis le 1er du mois)"},
            {"level": "L1", "type": "practice", "text": "Créer une mesure YTD avec TOTALYTD et vérifier l'accumulation progressive dans un graphique"},
            {"level": "L1", "type": "practice", "text": "Créer une mesure Ventes N-1 avec SAMEPERIODLASTYEAR et calculer la variation annuelle"},
            {"level": "L2", "type": "theory",   "text": "Distinguer TOTALYTD et DATESYTD : syntaxe raccourcie vs table de dates utilisable dans CALCULATE avec filtres supplémentaires"},
            {"level": "L2", "type": "theory",   "text": "Expliquer le paramètre d'année fiscale dans TOTALYTD : 3e argument optionnel pour définir la date de fin d'exercice"},
            {"level": "L2", "type": "theory",   "text": "Distinguer PREVIOUSYEAR / PREVIOUSMONTH vs DATEADD : décalage d'une période entière vs décalage d'un nombre d'intervalles défini"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les mesures semi-additives : solde bancaire ou stock — on ne peut pas additionner sur le temps, on utilise LASTDATE ou AVERAGEX"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la moyenne mobile avec DATESINPERIOD : moyenne des N dernières périodes glissantes"},
            {"level": "L2", "type": "practice", "text": "Créer QTD et MTD et vérifier la remise à zéro à chaque début de période"},
            {"level": "L2", "type": "practice", "text": "Calculer Δ vs N-1 (absolu) et % Δ vs N-1 avec DIVIDE pour la robustesse face aux divisions par zéro"},
            {"level": "L2", "type": "practice", "text": "Créer un cumul progressif (Running Total) avec DATESYTD et vérifier que la courbe ne redescend jamais"},
            {"level": "L2", "type": "practice", "text": "Créer une moyenne mobile sur 3 mois avec DATESINPERIOD et comparer à la courbe brute"},
            {"level": "L2", "type": "practice", "text": "Créer un classement avec RANKX(ALL(DimProduit), [Ventes]) et observer le rang dynamique selon les filtres"},
        ]
    },
    {
        "id": "m7", "emoji": "📊", "title": "Créer des rapports",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Associer le bon visuel à une question : Carte (KPI), Courbe (tendance), Histogramme (comparaison), Matrice (croisé), Carte géo (spatial)"},
            {"level": "L1", "type": "theory",   "text": "Distinguer rapport interactif (exploration dynamique) et rapport paginé (impression / export PDF pixel-perfect)"},
            {"level": "L1", "type": "practice", "text": "Créer un visuel Carte, un graphique en courbes et un histogramme avec les bons champs"},
            {"level": "L1", "type": "practice", "text": "Configurer un slicer et un filtre de page pour permettre l'exploration des données"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la différence courbes vs aires empilées — et pourquoi éviter les graphiques en secteurs avec plus de 5 catégories"},
            {"level": "L2", "type": "theory",   "text": "Expliquer quand et comment utiliser la mise en forme conditionnelle : fond coloré, icônes, barres de données dans une matrice"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les calculs visuels (Visual Calculations) : mesures DAX dans le contexte visuel uniquement, sans impacter le modèle"},
            {"level": "L2", "type": "practice", "text": "Créer une matrice avec sous-totaux et vérifier la cohérence des totaux de lignes et colonnes"},
            {"level": "L2", "type": "practice", "text": "Créer un visuel KPI avec valeur, objectif et axe de tendance — vérifier que le statut (↑↓) reflète l'écart"},
            {"level": "L2", "type": "practice", "text": "Appliquer un thème JSON personnalisé pour une palette de couleurs cohérente sur tout le rapport"},
            {"level": "L2", "type": "practice", "text": "Configurer la mise en forme conditionnelle sur une colonne de matrice avec des règles basées sur une mesure"},
            {"level": "L2", "type": "practice", "text": "Configurer la mise en page 16:9 avec image de fond et disposition mobile cohérente"},
            {"level": "L2", "type": "practice", "text": "Importer un visuel personnalisé depuis AppSource et l'intégrer au rapport"},
        ]
    },
    {
        "id": "m8", "emoji": "🧭", "title": "Améliorer les rapports pour l'utilisabilité",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Distinguer Drill-Down (descend dans une hiérarchie au sein du même visuel) et Drill-Through (navigue vers une page dédiée)"},
            {"level": "L1", "type": "theory",   "text": "Distinguer Cross-Filter (filtre les autres visuels — valeurs non liées disparaissent) et Cross-Highlight (valeurs non liées atténuées en gris)"},
            {"level": "L1", "type": "theory",   "text": "Expliquer les signets : ils capturent l'état d'une page (filtres, visibilité) et permettent de l'y revenir en un clic"},
            {"level": "L1", "type": "practice", "text": "Activer le Drill-Down sur un visuel avec hiérarchie et naviguer vers le bas et vers le haut"},
            {"level": "L1", "type": "practice", "text": "Modifier l'interaction entre deux visuels (Filtre, Mise en évidence ou Aucun) via Format > Modifier les interactions"},
            {"level": "L2", "type": "theory",   "text": "Distinguer Drill-Down et Expand All : Drill-Down remplace le niveau pour un membre ; Expand All développe tous les membres simultanément"},
            {"level": "L2", "type": "theory",   "text": "Expliquer quand utiliser 'Aucune interaction' : sur les KPI globaux qui ne doivent pas réagir aux sélections de l'utilisateur"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les trois options de capture d'un signet : Données (filtres), Affichage (visibilité des visuels), Page active — activables indépendamment"},
            {"level": "L2", "type": "theory",   "text": "Expliquer le pattern Toggle avec signets : deux signets + deux boutons pour basculer entre deux états (ex : graphique ↔ tableau)"},
            {"level": "L2", "type": "theory",   "text": "Expliquer l'Analyseur de performances : mesure le temps DAX (moteur), temps d'affichage et rendu — identifier où est le goulot d'étranglement"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la synchronisation des segments : propager la sélection d'un slicer à d'autres pages choisies, visible ou invisible sur chacune"},
            {"level": "L2", "type": "practice", "text": "Configurer une page Drill-Through : créer la page, ajouter le champ dans la zone Drill-through, tester via clic droit > Extraire"},
            {"level": "L2", "type": "practice", "text": "Créer un pattern Toggle : deux visuels superposés, deux signets Affichage, deux boutons associés"},
            {"level": "L2", "type": "practice", "text": "Créer une info-bulle de page personnalisée : page taille Info-bulle + visuels + désignation dans les propriétés du visuel principal"},
            {"level": "L2", "type": "practice", "text": "Synchroniser un slicer sur plusieurs pages via Affichage > Synchroniser les segments"},
            {"level": "L2", "type": "practice", "text": "Utiliser l'Analyseur de performances pour identifier le visuel le plus lent et copier sa requête DAX pour analyse"},
        ]
    },
    {
        "id": "m9", "emoji": "🔐", "title": "Sécurité & Gouvernance",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Décrire les quatre couches de sécurité Power BI : authentification Azure AD, rôles workspace, partage d'éléments, RLS/OLS"},
            {"level": "L1", "type": "theory",   "text": "Distinguer RLS statique (rôle fixe par région) et RLS dynamique (un seul rôle basé sur l'email de l'utilisateur connecté)"},
            {"level": "L1", "type": "practice", "text": "Créer un rôle RLS statique avec un filtre DAX et le tester avec 'Afficher en tant que rôles'"},
            {"level": "L1", "type": "practice", "text": "Assigner un utilisateur à un rôle RLS dans le Service Power BI"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les quatre rôles workspace : Admin, Membre, Contributeur, Lecteur — et la règle du rôle minimum"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la permission Build : donne accès au semantic model pour créer des rapports, indépendamment des rôles workspace"},
            {"level": "L2", "type": "theory",   "text": "Expliquer comment la RLS se propage via les relations : filtre sur DimMagasin → propagé vers FactVentes automatiquement"},
            {"level": "L2", "type": "theory",   "text": "Expliquer l'OLS (Object Level Security) : masque des tables ou colonnes entières — contrairement à RLS qui filtre les lignes"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les étiquettes de confidentialité Microsoft Purview : classification + politiques de protection, héritage vers les rapports"},
            {"level": "L2", "type": "practice", "text": "Implémenter la RLS dynamique : table de mapping Email/Région + relation + LOOKUPVALUE(... USERPRINCIPALNAME()) dans le rôle"},
            {"level": "L2", "type": "practice", "text": "Partager un rapport avec Build permission et vérifier que le destinataire peut créer un nouveau rapport depuis le modèle"},
            {"level": "L2", "type": "practice", "text": "Appliquer une étiquette de confidentialité sur un semantic model et observer la propagation aux rapports"},
        ]
    },
    {
        "id": "m10", "emoji": "🏢", "title": "Espaces de travail & Gestion du contenu",
        "items": [
            {"level": "L1", "type": "theory",   "text": "Expliquer la différence entre Mon Espace de Travail (personnel, non partageable) et un workspace partagé"},
            {"level": "L1", "type": "theory",   "text": "Distinguer rapport, tableau de bord et app Power BI — ce que chacun permet et à qui il est destiné"},
            {"level": "L1", "type": "practice", "text": "Publier un rapport depuis Desktop vers un workspace partagé"},
            {"level": "L1", "type": "practice", "text": "Configurer un refresh planifié pour un semantic model dans le Service"},
            {"level": "L2", "type": "theory",   "text": "Expliquer la bonne pratique Dev/Prod : workspace Développement pour les itérations, workspace Production pour le contenu certifié"},
            {"level": "L2", "type": "theory",   "text": "Décrire les quatre rôles workspace (Admin, Membre, Contributeur, Lecteur) et leurs permissions respectives"},
            {"level": "L2", "type": "theory",   "text": "Distinguer Promouvoir (badge jaune, tout membre) et Certifier (badge vert, administrateur autorisé) un contenu"},
            {"level": "L2", "type": "theory",   "text": "Expliquer pourquoi la passerelle (Gateway) est nécessaire pour les sources on-premises lors du refresh planifié dans le cloud"},
            {"level": "L2", "type": "theory",   "text": "Expliquer les limites du refresh planifié selon la licence : 8 fois/jour en Pro, 48 fois/jour avec Premium"},
            {"level": "L2", "type": "practice", "text": "Créer un workspace, configurer licence et description, inviter des collaborateurs avec les bons rôles"},
            {"level": "L2", "type": "practice", "text": "Créer une app depuis un workspace : sélectionner les rapports, définir les audiences, configurer les permissions par audience"},
            {"level": "L2", "type": "practice", "text": "Créer un tableau de bord en épinglant des visuels depuis plusieurs rapports différents"},
            {"level": "L2", "type": "practice", "text": "Créer une alerte de données sur un tile (seuil + fréquence + notification email)"},
            {"level": "L2", "type": "practice", "text": "Promouvoir un semantic model depuis le Service et observer le badge dans le workspace"},
        ]
    },
]

TOTAL_PAGES = len(MODULES) + 2  # 0=accueil, 1-10=modules, 11=bilan

# ─── STATE ────────────────────────────────────────────────────────────────────
def init():
    if "page" not in st.session_state:
        st.session_state.page = 0
    if "checks" not in st.session_state:
        st.session_state.checks = {}
    if "student" not in st.session_state:
        st.session_state.student = {"name": "", "email": "", "cohort": ""}
    if "show_l2" not in st.session_state:
        st.session_state.show_l2 = {}
    for mod in MODULES:
        if mod["id"] not in st.session_state.show_l2:
            st.session_state.show_l2[mod["id"]] = False
        for i in range(len(mod["items"])):
            k = f"{mod['id']}_{i}"
            if k not in st.session_state.checks:
                st.session_state.checks[k] = False

def go_to(p):
    st.session_state.page = p
    st.rerun()

# ─── STATS ────────────────────────────────────────────────────────────────────
def mod_stats(mod, level=None):
    items = mod["items"]
    items_f = [(i, it) for i, it in enumerate(items)
               if (level is None or it.get("level") == level)]
    total = len(items_f)
    th  = [i for i, it in items_f if it["type"] == "theory"]
    pr  = [i for i, it in items_f if it["type"] == "practice"]
    done   = sum(1 for i, _ in items_f if st.session_state.checks.get(f"{mod['id']}_{i}", False))
    th_d   = sum(1 for i in th if st.session_state.checks.get(f"{mod['id']}_{i}", False))
    pr_d   = sum(1 for i in pr if st.session_state.checks.get(f"{mod['id']}_{i}", False))
    return {"done": done, "total": total,
            "pct": round(done / total * 100) if total else 0,
            "th_done": th_d, "th_total": len(th),
            "pr_done": pr_d, "pr_total": len(pr)}

def global_stats():
    done  = sum(1 for v in st.session_state.checks.values() if v)
    total = sum(len(m["items"]) for m in MODULES)
    th_d  = sum(1 for mod in MODULES for i, it in enumerate(mod["items"])
                if it["type"] == "theory" and st.session_state.checks.get(f"{mod['id']}_{i}", False))
    th_t  = sum(1 for mod in MODULES for it in mod["items"] if it["type"] == "theory")
    pr_d  = sum(1 for mod in MODULES for i, it in enumerate(mod["items"])
                if it["type"] == "practice" and st.session_state.checks.get(f"{mod['id']}_{i}", False))
    pr_t  = sum(1 for mod in MODULES for it in mod["items"] if it["type"] == "practice")
    return {"done": done, "total": total,
            "pct": round(done / total * 100) if total else 0,
            "th_done": th_d, "th_total": th_t,
            "pr_done": pr_d, "pr_total": pr_t}

# ─── CSV ──────────────────────────────────────────────────────────────────────
def make_csv():
    s = st.session_state.student
    rows = []
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    for mod in MODULES:
        ms = mod_stats(mod)
        for i, item in enumerate(mod["items"]):
            rows.append({
                "submitted_at":  ts,
                "student_name":  s.get("name", ""),
                "student_email": s.get("email", ""),
                "cohort":        s.get("cohort", ""),
                "module_id":     mod["id"],
                "module_title":  mod["title"],
                "item_level":    item.get("level", ""),
                "item_type":     item["type"],
                "item_text":     item["text"],
                "checked":       1 if st.session_state.checks.get(f"{mod['id']}_{i}", False) else 0,
                "module_pct":    ms["pct"],
            })
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=rows[0].keys())
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode("utf-8")

# ─── CHARTS ───────────────────────────────────────────────────────────────────
def radar_fig(stats_list):
    short = ["Power BI", "Prérequis", "Sources", "Modèle", "DAX",
             "Time Intel.", "Rapports", "Usabilité", "Sécurité", "Workspaces"]
    vals = [s["pct"] for s in stats_list] + [stats_list[0]["pct"]]
    labs = short + [short[0]]
    fig = go.Figure(go.Scatterpolar(
        r=vals, theta=labs, fill="toself",
        fillcolor="rgba(22,163,74,0.12)",
        line=dict(color="#16a34a", width=2), name="Couverture",
    ))
    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], ticksuffix="%", tickfont_size=9),
            angularaxis=dict(tickfont_size=9),
        ),
        showlegend=False, margin=dict(t=10, b=10, l=20, r=20), height=300,
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
    )
    return fig

def bar_fig(stats_list):
    short = ["Power BI", "Prérequis", "Sources", "Modèle", "DAX",
             "Time Intel.", "Rapports", "Usabilité", "Sécurité", "Workspaces"]
    th = [round(s["th_done"] / s["th_total"] * 100) if s["th_total"] else 0 for s in stats_list]
    pr = [round(s["pr_done"] / s["pr_total"] * 100) if s["pr_total"] else 0 for s in stats_list]
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Théorie",  x=short, y=th, marker_color="#2563eb",
                         text=[f"{v}%" for v in th], textposition="outside"))
    fig.add_trace(go.Bar(name="Pratique", x=short, y=pr, marker_color="#16a34a",
                         text=[f"{v}%" for v in pr], textposition="outside"))
    fig.update_layout(
        barmode="group",
        yaxis=dict(range=[0, 115], ticksuffix="%"),
        xaxis=dict(tickangle=-30, tickfont_size=9),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=30, b=80, l=10, r=10), height=320,
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
    )
    return fig

# ─── INSIGHTS ─────────────────────────────────────────────────────────────────
def build_insights(stats_list):
    g = global_stats()
    pct = g["pct"]
    ins = []
    if pct == 100:
        ins.append(("success", "🎉 Tu as complété 100% de la checklist — tu es prêt(e) pour le PL-300 !"))
    elif pct >= 75:
        ins.append(("success", f"✅ Excellente progression à {pct}% — tu es dans la dernière ligne droite."))
    elif pct >= 40:
        ins.append(("info", f"📈 Tu es à {pct}%. Concentre-toi sur les modules en dessous de 50%."))
    else:
        ins.append(("warning", f"⚠️ Seulement {pct}% complété. Reprends les éléments non cochés avant l'examen."))

    th_pct = round(g["th_done"] / g["th_total"] * 100) if g["th_total"] else 0
    pr_pct = round(g["pr_done"] / g["pr_total"] * 100) if g["pr_total"] else 0
    if th_pct - pr_pct >= 20:
        ins.append(("warning", f"🔵 Tu maîtrises la théorie ({th_pct}%) mais la pratique est en retard ({pr_pct}%). Passe plus de temps sur Power BI Desktop."))
    elif pr_pct - th_pct >= 15:
        ins.append(("info", f"🟢 Tu es solide en pratique ({pr_pct}%) mais les concepts théoriques ({th_pct}%) méritent plus de révision — ils apparaissent dans les QCM."))

    sorted_mods = sorted(stats_list, key=lambda x: x["pct"])
    weak = [m for m in sorted_mods[:3] if m["pct"] < 60]
    if weak:
        titles = ", ".join(f"{m['title']} ({m['pct']}%)" for m in weak)
        ins.append(("warning", f"🎯 Points à travailler en priorité : {titles}"))

    strong = [m for m in sorted_mods if m["pct"] == 100]
    if strong:
        titles = ", ".join(m["title"] for m in strong)
        ins.append(("success", f"💪 Modules maîtrisés à 100% : {titles}"))

    dax = next((s for s in stats_list if s.get("id") == "m5"), None)
    if dax and dax["pct"] < 60:
        ins.append(("warning", "🧮 Les fondamentaux DAX sont sous 60% — ce thème est lourd dans l'examen PL-300. À prioriser absolument."))
    return ins

# ═══════════════════════════════════════════════════════════════════════════════
init()
page = st.session_state.page

# ══ PAGE 0 : Accueil ══════════════════════════════════════════════════════════
if page == 0:
    g = global_stats()

    st.title("📊 Checklist PL-300 Power BI")
    st.caption("Évalue ta maîtrise module par module — deux niveaux de détail — et exporte ton bilan.")

    if g["total"] > 0:
        st.progress(g["pct"] / 100, text=f"Progression globale : {g['pct']}%  ({g['done']}/{g['total']} éléments)")

    st.divider()

    st.subheader("Tes informations")
    c1, c2 = st.columns(2)
    with c1:
        name = st.text_input("Prénom et nom", value=st.session_state.student.get("name", ""),
                             placeholder="Marie Dupont")
    with c2:
        email = st.text_input("Email", value=st.session_state.student.get("email", ""),
                              placeholder="marie@example.com")
    cohort = st.text_input("Promotion", value=st.session_state.student.get("cohort", ""),
                           placeholder="ex: dafs-ft-13")
    st.session_state.student = {"name": name, "email": email, "cohort": cohort}

    st.divider()
    st.subheader("Modules")

    for i, mod in enumerate(MODULES):
        s_l1  = mod_stats(mod, "L1")
        s_l2  = mod_stats(mod, "L2")
        s_all = mod_stats(mod)

        with st.container(border=True):
            col_info, col_btn = st.columns([3, 1])
            with col_info:
                st.markdown(f"**{mod['emoji']} {mod['title']}**")
                st.caption(
                    f"N1 : {s_l1['done']}/{s_l1['total']} · "
                    f"N2 : {s_l2['done']}/{s_l2['total']} · "
                    f"Total : {s_all['pct']}%"
                )
                st.progress(s_all["pct"] / 100)
            with col_btn:
                label = "Commencer" if s_all["pct"] == 0 else f"Continuer"
                if st.button(label, key=f"go_{i}", use_container_width=True):
                    go_to(i + 1)

    st.divider()
    if g["done"] > 0:
        if st.button("📈 Voir mon tableau de bord", use_container_width=True, type="primary"):
            go_to(TOTAL_PAGES - 1)

# ══ PAGES 1-10 : Modules ══════════════════════════════════════════════════════
elif 1 <= page <= len(MODULES):
    mod     = MODULES[page - 1]
    items   = mod["items"]
    show_l2 = st.session_state.show_l2.get(mod["id"], False)
    g       = global_stats()

    # Header
    st.caption(f"Module {page} / {len(MODULES)} — Progression globale : {g['pct']}%")
    st.progress(g["pct"] / 100)

    st.title(f"{mod['emoji']} {mod['title']}")

    # Level switcher
    s_l1 = mod_stats(mod, "L1")
    s_l2 = mod_stats(mod, "L2")

    st.write("**Niveau de détail**")
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        l1_type = "primary" if not show_l2 else "secondary"
        if st.button(f"Niveau 1 — Essentiel  ({s_l1['done']}/{s_l1['total']})",
                     key="lvl1", use_container_width=True, type=l1_type):
            st.session_state.show_l2[mod["id"]] = False
            st.rerun()
    with col_l2:
        l2_type = "primary" if show_l2 else "secondary"
        if st.button(f"Niveau 2 — Approfondi  ({s_l2['done']}/{s_l2['total']})",
                     key="lvl2", use_container_width=True, type=l2_type):
            st.session_state.show_l2[mod["id"]] = True
            st.rerun()

    # Active level items
    active_level = "L2" if show_l2 else "L1"
    active_items = [(i, it) for i, it in enumerate(items) if it.get("level") == active_level]
    s_active     = mod_stats(mod, active_level)

    if show_l2:
        st.info("**Niveau 2 — Approfondi** : détails techniques, cas limites et points d'examen.", icon="🔬")
    else:
        st.info("**Niveau 1 — Essentiel** : les concepts clés à maîtriser en premier.", icon="🎯")

    st.progress(s_active["pct"] / 100,
                text=f"Ce niveau : {s_active['done']}/{s_active['total']} ({s_active['pct']}%)")

    # Theory items
    theory_items   = [(i, it) for i, it in active_items if it["type"] == "theory"]
    practice_items = [(i, it) for i, it in active_items if it["type"] == "practice"]

    if theory_items:
        st.markdown("**🔵 Théorie** — *je comprends ce concept*")
        for i, item in theory_items:
            key = f"{mod['id']}_{i}"
            val = st.checkbox(item["text"], value=st.session_state.checks.get(key, False), key=f"cb_{key}")
            st.session_state.checks[key] = val

    if practice_items:
        st.markdown("**🟢 Pratique** — *je sais le faire dans Power BI*")
        for i, item in practice_items:
            key = f"{mod['id']}_{i}"
            val = st.checkbox(item["text"], value=st.session_state.checks.get(key, False), key=f"cb_{key}")
            st.session_state.checks[key] = val

    st.divider()

    # Navigation
    is_last = (page == len(MODULES))
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("← Précédent", use_container_width=True):
            go_to(page - 1)
    with c2:
        if st.button("🏠 Accueil", use_container_width=True):
            go_to(0)
    with c3:
        next_label = "📈 Voir mon bilan →" if is_last else "Suivant →"
        if st.button(next_label, use_container_width=True, type="primary"):
            go_to(TOTAL_PAGES - 1 if is_last else page + 1)

# ══ PAGE 11 : Dashboard ═══════════════════════════════════════════════════════
elif page == TOTAL_PAGES - 1:
    g = global_stats()
    stats_list = []
    for mod in MODULES:
        s = mod_stats(mod)
        s["id"]    = mod["id"]
        s["title"] = mod["title"]
        s["emoji"] = mod["emoji"]
        stats_list.append(s)

    name_display = st.session_state.student.get("name") or "Étudiant·e"
    st.title(f"📈 Bilan — {name_display}")
    st.caption(f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}")

    # KPIs
    th_pct = round(g["th_done"] / g["th_total"] * 100) if g["th_total"] else 0
    pr_pct = round(g["pr_done"] / g["pr_total"] * 100) if g["pr_total"] else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Progression globale", f"{g['pct']}%", f"{g['done']}/{g['total']} éléments")
    c2.metric("Théorie", f"{th_pct}%", f"{g['th_done']}/{g['th_total']}")
    c3.metric("Pratique", f"{pr_pct}%", f"{g['pr_done']}/{g['pr_total']}")

    st.divider()

    # Charts
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Couverture par module**")
        st.plotly_chart(radar_fig(stats_list), use_container_width=True, config={"displayModeBar": False})
    with col_b:
        st.markdown("**Théorie vs Pratique**")
        st.plotly_chart(bar_fig(stats_list), use_container_width=True, config={"displayModeBar": False})

    # Module table
    st.divider()
    st.markdown("**Détail par module**")
    rows = []
    for s in stats_list:
        rows.append({
            "Module":   f"{s['emoji']} {s['title']}",
            "Score":    f"{s['pct']}%",
            "Théorie":  f"{s['th_done']}/{s['th_total']}",
            "Pratique": f"{s['pr_done']}/{s['pr_total']}",
            "Statut":   "✅ Maîtrisé" if s["pct"] == 100 else ("🔶 En cours" if s["pct"] > 0 else "⬜ Non commencé"),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Insights
    st.divider()
    st.markdown("**💡 Insights**")
    for kind, msg in build_insights(stats_list):
        if kind == "success":
            st.success(msg)
        elif kind == "warning":
            st.warning(msg)
        else:
            st.info(msg)

    # Focus list
    unchecked = [
        {"Module": mod["title"], "Niveau": it.get("level",""), "Type": it["type"].capitalize(), "À maîtriser": it["text"]}
        for mod in MODULES
        for i, it in enumerate(mod["items"])
        if not st.session_state.checks.get(f"{mod['id']}_{i}", False)
    ]
    if unchecked:
        st.divider()
        st.markdown("**🎯 Points prioritaires (20 premiers)**")
        st.dataframe(pd.DataFrame(unchecked[:20]), use_container_width=True, hide_index=True)

    # Export
    st.divider()
    st.markdown("**📤 Exporter mes résultats**")

    if not st.session_state.student.get("name"):
        st.warning("Ajoute ton nom sur la page d'accueil avant d'exporter.")

    notes = st.text_area("Message pour ton formateur (optionnel)",
                         placeholder="ex : j'ai du mal avec la RLS dynamique...",
                         height=80)

    name_safe = (st.session_state.student.get("name", "etudiant") or "etudiant").replace(" ", "_")
    fname     = f"{name_safe}_PL300_{datetime.now().strftime('%Y%m%d_%H%M')}"

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            "⬇️ Télécharger CSV",
            data=make_csv(),
            file_name=f"{fname}.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary",
            help="À envoyer à ton formateur — un fichier par étudiant, fusionnable en un tableau."
        )
    with col_dl2:
        full_json = json.dumps({
            "student":      st.session_state.student,
            "submitted_at": datetime.now().isoformat(),
            "overall":      g,
            "modules":      [{"title": mod["title"], **mod_stats(mod)} for mod in MODULES],
            "notes":        notes,
        }, indent=2, ensure_ascii=False)
        st.download_button(
            "⬇️ Télécharger JSON",
            data=full_json.encode("utf-8"),
            file_name=f"{fname}.json",
            mime="application/json",
            use_container_width=True,
        )

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Retour à l'accueil", use_container_width=True):
            go_to(0)
    with c2:
        if st.button(f"← Reprendre le module {len(MODULES)}", use_container_width=True):
            go_to(len(MODULES))
