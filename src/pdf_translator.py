#!/usr/bin/env python3
"""
PDF Translator - A script to translate technical PDF documents
"""

import os
import re
import sys
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from tqdm import tqdm
from pdf2docx import Converter
import docx
from docx.shared import Pt, RGBColor

class PDFTranslator:
    def __init__(self, dictionary_path=None):
        """Initialize the PDF translator with optional custom dictionary path."""
        self.technical_terms = {
            # General electrical terms
            "Protection": "Protection",
            "Relay": "Relais",
            "Voltage": "Tension",
            "Current": "Courant",
            "Phase": "Phase",
            "Earth": "Terre",
            "Fault": "Défaut",
            "Over current": "Surintensité",
            "Earth fault": "Défaut à la terre",
            "Trip": "Déclenchement",
            "Rating": "Calibre",
            "Setting": "Réglage",
            "Curve": "Courbe",
            "Directional": "Directionnel",
            "Non Directional": "Non Directionnel",
            "Operating time": "Temps de fonctionnement",
            "Grading Margin": "Marge d'échelonnement",
            "CT Ratio": "Rapport TI",
            "VT Ratio": "Rapport TP",
            "Primary": "Primaire",
            "Secondary": "Secondaire",
            "Pickup": "Pickup",
            "Function": "Fonction",
            "Input": "Entrée",
            "Direction": "Direction",
            "Stage": "Étape",
            "Reset": "Réinitialisation",
            "Tripping characteristic": "Caractéristique de déclenchement",
            "Curve type": "Type de courbe",
            "Down stream": "En aval",
            "Grading": "Échelonnement",
            "Required": "Requis",
            "Maximum fault current": "Courant de défaut maximum",
            "Operate time": "Temps de fonctionnement",
            "Fault current": "Courant de défaut",
            "Calculation": "Calcul",
            "Considered": "Considéré",
            "Full load current": "Courant de pleine charge",
            "Date": "Date",
            "Revision": "Révision",
            "Prepared By": "Préparé Par",
            "Checked By": "Vérifié Par",
            "Approved By": "Approuvé Par",
            
            # Technical terms specific to the buscoupler document
            "SETTING CALCULATION DOCUMENT": "DOCUMENT DE CALCUL DE RÉGLAGE",
            "BUSCOUPLER PROTECTION": "PROTECTION DE COUPLEUR DE BARRES",
            "SETTING RECOMMENDATION": "RECOMMANDATION DE RÉGLAGE",
            "CONFIDENTIAL": "CONFIDENTIEL",
            "The information contained in this document is not to be communicated either directly or indirectly to any person not authorised to receive it": 
            "Les informations contenues dans ce document ne doivent être communiquées ni directement ni indirectement à toute personne non autorisée à les recevoir",
            
            "Customer": "Client",
            "Project": "Projet",
            "ALIMENTATION EN ENERGIE ELECTRIQUE EN UN POINT UNIQUE DE 2x30MVA PAR LIGNES PERSONNALISEES ET SECURISEES":
            "POWER SUPPLY AT A SINGLE POINT OF 2x30MVA BY CUSTOMIZED AND SECURED LINES",
            
            "Relay used": "Relais utilisé",
            "Rated voltage": "Tension nominale",
            "CT Ratio": "Rapport TC",
            "CTR-Pri": "CTR-Pri",
            "CTR-Sec": "CTR-Sec",
            "Full load current in Primary": "Courant de pleine charge au primaire",
            "Maximum fault current in Primary (Given)": "Courant de défaut maximal au primaire (donné)",
            
            "Over current protection": "Protection contre les surintensités",
            "Phase TOC-1 (I>1) Non Directional Over current: 51": "Phase TOC-1 (I>1) Surintensité non directionnelle: 51",
            "I>1 Function": "Fonction I>1",
            "I>1 Input": "Entrée I>1",
            "I>1 Direction": "Direction I>1",
            "Stage 1 Over current pickup in primary": "Pickup de surintensité étape 1 au primaire",
            "Stage 1 Over current pickup in secondary (I>1 Current Set)": "Pickup de surintensité étape 1 au secondaire (I>1 Réglage de courant)",
            "Tripping characteristic for I>1": "Caractéristique de déclenchement pour I>1",
            "Curve type (I>1 Curve)": "Type de courbe (I>1 Courbe)",
            "Down stream Operate time (Trafo Downstream)": "Temps de fonctionnement en aval (Trafo en aval)",
            "Grading Margin": "Marge d'échelonnement",
            "Required operating time": "Temps de fonctionnement requis",
            
            "Note: The IDMT curves shall saturate if the fault current is more than 20 times of pickup current. As in this case the actual fault current is 37800 A which is more than 20 times of pickup current hence for calculation purpose we shall consider": 
            "Remarque: Les courbes IDMT saturent si le courant de défaut est supérieur à 20 fois le courant de pickup. Dans ce cas, le courant de défaut réel est de 37800 A, ce qui est supérieur à 20 fois le courant de pickup, donc pour les besoins du calcul, nous considérerons",
            
            "Fault current considered for calculation": "Courant de défaut considéré pour le calcul",
            "Operating time @ TMS = 1": "Temps de fonctionnement @ TMS = 1",
            "TMS (I>1 TMS)": "TMS (I>1 TMS)",
            "I>1 Reset Char": "Caractéristique de réinitialisation I>1",
            "I>1 tReset": "I>1 tRéinitialisation",
            
            "Earth fault protection": "Protection contre les défauts à la terre",
            "EF1- TOC (IN>1) Non Directional Earth fault: 51N": "EF1- TOC (IN>1) Défaut à la terre non directionnel: 51N",
            "IN>1 Function": "Fonction IN>1",
            "IN>1 Direction": "Direction IN>1",
            "Stage 1 Earth fault pickup in primary": "Seuil de déclenchement défaut à la terre étage 1 au primaire",
            "Stage 1 Earth fault pickup in secondary (IN1>1 Current)": "Seuil de déclenchement défaut à la terre étage 1 au secondaire (Courant IN1>1)",
            "Tripping characteristic for IN>1": "Caractéristique de déclenchement pour IN>1",
            "Curve type (IN1>1 Curve)": "Type de courbe (Courbe IN1>1)",
            "LV (Down stream) Operate time (Trafo Downstream)": "Temps de fonctionnement BT (en aval) (Trafo en aval)",
            "FEEDER DETAILS": "DÉTAILS DE L'ALIMENTATION",
            "PU Reference": "Référence PU",
            "Name of the Bay": "Nom de la Travée",
            "Note : PU address to be confirmed at site.": "Remarque : Adresse PU à confirmer sur site.",
            "63kV BUS BAR PROTECTION SETTINGS": "PARAMÈTRES DE PROTECTION DE JEUX DE BARRES 63kV",
            "Substation Information & Inputs:": "Informations et entrées du poste :",
            "Settings and configuration of the P740 Numerical bus bar protection scheme is for 220kV is arrived based on the substation information provided below.": "Les paramètres et la configuration du schéma de protection numérique de jeux de barres P740 pour 220kV sont établis sur la base des informations du poste fournies ci-dessous.",
            "Type of Bus Arrangement & No. of Bays:": "Type d'arrangement de bus et nombre de travées :",
            "Number of independent bars": "Nombre de barres indépendantes",
            "Number of feeders": "Nombre d'alimentations",
            "Fault levels & Load details:": "Niveaux de défaut et détails de charge :",
            "Maximum Fault level (Given)": "Niveau de défaut maximum (Donné)",
            "Minimum Fault level (Given)": "Niveau de défaut minimum (Donné)",
            "Maximum Load level in a feeder (Given)": "Niveau de charge maximum dans une alimentation (Donné)",
            "Minimum Load level in a feeder (Given)": "Niveau de charge minimum dans une alimentation (Donné)",
            "Maximum loading of 1 bar (Given)": "Charge maximale d'une barre (Donnée)",
            "Fault clearing time": "Temps d'élimination du défaut",
            "CT Details:": "Détails TC :",
            "Setting Crieteria - Central Unit (MiCOM P741)": "Critères de réglage - Unité centrale (MiCOM P741)",
            "DIFFERENTIAL ELEMENTS 87BB SETTINGS:": "PARAMÈTRES DES ÉLÉMENTS DIFFÉRENTIELS 87BB :",
            "ID>1 Current Set:": "Réglage de courant ID>1 :",
            "This setting is set for the phase circuitry fault monitoring characteristic for the minimum pickup.": "Ce réglage est défini pour la caractéristique de surveillance des défauts de circuit de phase pour le pickup minimum.",
            "Recommended ID>1": "ID>1 recommandé",
            "Note : If the spill current is more than the recommended value at site it shall be adjusted accordingly.": "Remarque : Si le courant de fuite est supérieur à la valeur recommandée sur site, il sera ajusté en conséquence.",
            "Phase slope k1:": "Pente de phase k1 :",
            "Recommended k1": "k1 recommandé",
            "ID>1 Alarm Timer:": "Temporisateur d'alarme ID>1 :",
            "This timer shall be greater than the longest protection time (such as line, over current, etc...)": "Ce temporisateur doit être supérieur au temps de protection le plus long (comme ligne, surintensité, etc...)",
            "Recommended ID>1 timer": "Temporisateur ID>1 recommandé",
            "Phase slope kCZ:": "Pente de phase kCZ :",
            "This is the Slope angle setting for the check zone biased differential element.": "Il s'agit du réglage de l'angle de pente pour l'élément différentiel polarisé de la zone de contrôle.",
            "Recommended kCZ": "kCZ recommandé",
            "ID>2 Current Set:": "Réglage de courant ID>2 :",
            "Recommended ID>2": "ID>2 recommandé",
            "Phase slope k2:": "Pente de phase k2 :",
            "Recommended k2": "k2 recommandé",
            "IDCZ>2 Current Set:": "Réglage de courant IDCZ>2 :",
            "Recommended IDCZ>2": "IDCZ>2 recommandé",
            "CB FAIL: 50CBF": "DÉFAILLANCE DISJONCTEUR : 50CBF",
            "For different CTRs:": "Pour différents CTRs :",
            "I< Current Set:": "Réglage de courant I< :",
            "CB Fail I< pickup": "Pickup I< de défaillance disjoncteur",
            "CBF Retrip Timer - TBF1 & TBF3:": "Temporisateur de redéclenchement CBF - TBF1 & TBF3 :",
            "A re-trip shall be applied after a time tBF1/tBF3. Here we are issuing the retrip command after 50ms of main trip.": "Un redéclenchement sera appliqué après un temps tBF1/tBF3. Ici, nous émettons la commande de redéclenchement 50ms après le déclenchement principal.",
            "CBF Backtrip Timer - TBF2 & TBF4:": "Temporisateur de déclenchement de secours CBF - TBF2 & TBF4 :",
            "DEAD ZONE PROTECTION:": "PROTECTION DE ZONE MORTE :",
            "I>DZ Current Set: (only Line Bay)": "Réglage de courant I>DZ : (uniquement travée de ligne)",
            "ID>DZ pickup": "Pickup ID>DZ",
            "I>DZ Time Delay:": "Temporisation I>DZ :",
            "Dead Zone delay": "Temporisation de zone morte",
            "CT Supervision:": "Supervision TC :",
            "Error Factor KCE:": "Facteur d'erreur KCE :",
            "Alarm Delay TCE:": "Temporisation d'alarme TCE :",
            "CT Supervision Time delay": "Temporisation de supervision TC",
            "Settings Recommended for CU (P741) & PU (P743)": "Paramètres recommandés pour CU (P741) et PU (P743)",
            "P741 Settings": "Paramètres P741",
            "DIFF BUSBAR PROT": "PROT DIFF JEU DE BARRES",
            "Differential Phase Faults": "Défauts de phase différentiels",
            "CZ Parameters": "Paramètres CZ",
            "Zone Parameters": "Paramètres de zone",
            "Common": "Commun",
            "Phase slope K2": "Pente de phase K2",
            "ID>2 current": "Courant ID>2",
            "ID>1 current": "Courant ID>1",
            "Phase slope K1": "Pente de phase K1",
            "Differential Earth Fault": "Défaut de terre différentiel",
            "BUSBAR OPTIONS": "OPTIONS DE JEU DE BARRES",
            "CZ Circ Flt Mode": "Mode défaut circ CZ",
            "Zx circ flt mode": "Mode défaut circ Zx",
            "Circuitry t reset": "Réinitialisation t circuit",
            "Circ block mode": "Mode blocage circ",
            "CZ PU Error mode": "Mode erreur PU CZ",
            "Zx PU Error mode": "Mode erreur PU Zx",
            "PU Error timer": "Temporisateur erreur PU",
            "PU Error tReset": "tRéinitialisation erreur PU",
            "SEF Block Alarm": "Alarme blocage SEF",
            "Confirm Reset PU": "Confirmer réinitialisation PU",
            "3ph Block-Alarm": "Alarme-blocage 3ph",
            "Delay Trip Status": "État déclenchement retardé",
            "Diff Display Min": "Affichage min diff",
            "P743 Settings": "Paramètres P743",
            "CT RATIOS: (Inputs given)": "RAPPORTS TC : (Entrées données)",
            "Phase CT Primary": "TC phase primaire",
            "Phase CT Sec'y": "TC phase secondaire",
            "RBPh / RBN": "RBPh / RBN",
            "Power Parameters": "Paramètres de puissance",
            "Standard Input": "Entrée standard",
            "Knee Voltage Vk": "Tension de coude Vk",
            "RCT Sec'y": "RCT secondaire",
            "Eff. Burden Ohm": "Charge eff. Ohm",
            "British Standard": "Norme britannique",
            "Note : 1. RBPh/RBN & Effective Burden to be calculated at site & entered in the PU correctly.": "Remarque : 1. RBPh/RBN et charge effective à calculer sur site et à saisir correctement dans le PU.",
            "2. The values of RBPh/RBN, Vk, RCT & Effective Burden values are very important to detect & confirm the CT Saturation condition.": "2. Les valeurs de RBPh/RBN, Vk, RCT et de charge effective sont très importantes pour détecter et confirmer l'état de saturation du TC.",
            "DEAD ZONE PROTECTION: (Only Line Bays)": "PROTECTION DE ZONE MORTE : (Uniquement travées de ligne)",
            "I>DZ Current Set": "Réglage de courant I>DZ",
            "I>DZ Time Delay": "Temporisation I>DZ",
            "Dead Zone Earth": "Terre zone morte",
            "CB FAIL:": "DÉFAILLANCE DISJONCTEUR :",
            "Control By": "Contrôlé par",
            "I< Current Set for": "Réglage de courant I< pour",
            "I> Status": "État I>",
            "Internal Trip:": "Déclenchement interne :",
            "CB Fail Timer 1": "Temporisateur 1 de défaillance disjoncteur",
            "CB Fail Timer 2": "Temporisateur 2 de défaillance disjoncteur",
            "External Trip:": "Déclenchement externe :",
            "CB Fail Timer 3": "Temporisateur 3 de défaillance disjoncteur",
            "CB Fail Timer 4": "Temporisateur 4 de défaillance disjoncteur",
            "SUPERVISION:": "SUPERVISION :",
            "IO Supervision:": "Supervision E/S :",
            "Error Factor Kce": "Facteur d'erreur Kce",
            "Alarm Delay Tce": "Temporisation d'alarme Tce",
            "IO Sup. Blocking": "Blocage de supervision E/S",
            "87BBP & 87BBN": "87BBP et 87BBN",
            "TMS (IN1>1 TMS)": "TMS (TMS IN1>1)",
            "IN1>1 Reset Char": "Caractéristique de réinitialisation IN1>1",
            "IN1>1 tReset": "Réinitialisation IN1>1",
            
            # Transformer protection document terms
            "TRANSFORMER PROTECTION": "PROTECTION DE TRANSFORMATEUR",
            "30MVA TRANSFORMER PROTECTION": "PROTECTION DE TRANSFORMATEUR 30MVA",
            "SETTING CALCULATION DOCUMENT FOR 30 MVA TRANSFORMER PROTECTION": "DOCUMENT DE CALCUL DE RÉGLAGE POUR PROTECTION DE TRANSFORMATEUR 30 MVA",
            "DIFFERENTIAL PROTECTION FOR 30MVA TRAFO": "PROTECTION DIFFÉRENTIELLE POUR TRANSFORMATEUR 30MVA",
            "MICOM P643-87T": "MICOM P643-87T",
            "Transformer Total Capacity in MVA": "Capacité totale du transformateur en MVA",
            "Primary winding": "Enroulement primaire",
            "Secondary winding": "Enroulement secondaire",
            "HV winding": "Enroulement HT",
            "LV winding": "Enroulement BT",
            "HV winding (primary) rated voltage": "Tension nominale de l'enroulement HT (primaire)",
            "LV winding (secondary) rated": "Tension nominale de l'enroulement BT (secondaire)",
            "Primary winding (HV) CTR (Given)": "CTR de l'enroulement primaire (HT) (Donné)",
            "Secondary winding (LV) CTR (Given)": "CTR de l'enroulement secondaire (BT) (Donné)",
            "Transformer % Impedance (Given)": "Impédance du transformateur en % (Donnée)",
            "OLTC Range (Given)": "Plage OLTC (Donnée)",
            "HV winding voltages with OLTC ranges": "Tensions d'enroulement HT avec plages OLTC",
            "Vector Group": "Groupe de vecteurs",
            "HV winding full load current in Primary Amps": "Courant de pleine charge de l'enroulement HT en ampères primaires",
            "HV winding full load current in Secondary Amps": "Courant de pleine charge de l'enroulement HT en ampères secondaires",
            "LV winding full load current in Primary Amps": "Courant de pleine charge de l'enroulement BT en ampères primaires",
            "LV winding full load current in Secondary Amps": "Courant de pleine charge de l'enroulement BT en ampères secondaires",
            "Amplitude matching factor of HV winding": "Facteur d'adaptation d'amplitude de l'enroulement HT",
            "Amplitude matching factor of LV winding": "Facteur d'adaptation d'amplitude de l'enroulement BT",
            "HV winding amplitude corrected current in PU": "Courant corrigé en amplitude de l'enroulement HT en PU",
            "LV winding amplitude corrected current in PU": "Courant corrigé en amplitude de l'enroulement BT en PU",
            "SYSTEM CONFIG": "CONFIGURATION SYSTÈME",
            "Winding config": "Configuration d'enroulement",
            "Winding Type": "Type d'enroulement",
            "HV CT Terminals": "Bornes TC HT",
            "LV CT Terminals": "Bornes TC BT",
            "Ref Power S": "Puissance de référence S",
            "Ref Vector Group": "Groupe de vecteurs de référence",
            "HV Connection": "Connexion HT",
            "HV Grounding": "Mise à la terre HT",
            "HV Nominal": "HT Nominale",
            "HV Rating": "Calibre HT",
            "% Reactance": "% Réactance",
            "LV Vector Group": "Groupe de vecteurs BT",
            "LV Connection": "Connexion BT",
            "LV Grounding": "Mise à la terre BT",
            "LV Nominal": "BT Nominale",
            "LV Rating": "Calibre BT",
            "The Transformer Differential protection relay MiCOM P643-87T has the following parameters for setting:": "Le relais de protection différentielle de transformateur MiCOM P643-87T a les paramètres suivants pour le réglage :",
            "Minimum differential threshold of the low set differential characteristic": "Seuil différentiel minimum de la caractéristique différentielle à seuil bas",
            "First slope setting of the low set differential characteristic": "Réglage de la première pente de la caractéristique différentielle à seuil bas",
            "Second slope setting of the low set differential characteristic": "Réglage de la deuxième pente de la caractéristique différentielle à seuil bas",
            "Bias current threshold for the second slope of the low set differential characteristic": "Seuil de courant de polarisation pour la seconde pente de la caractéristique différentielle à seuil bas",
            "Threshold value of the differential current for deactivation of the inrush stabilization function (harmonic restraint) and of the overfluxing restraint": "Valeur de seuil du courant différentiel pour la désactivation de la fonction de stabilisation d'appel (retenue harmonique) et de la retenue de surflux",
            "Threshold value of the differential current for tripping by the differential protection function independent of restraining variable, harmonic restraint, overfluxing restraint and saturation detector": "Valeur de seuil du courant différentiel pour le déclenchement par la fonction de protection différentielle indépendamment de la variable de retenue, de la retenue harmonique, de la retenue de surflux et du détecteur de saturation",
            "DIFF PROTECTION": "PROTECTION DIFF",
            "Settings": "Réglages",
            "Trans Diff": "Diff Trans",
            "Set Mode": "Mode de réglage",
            "Is1": "Is1",
            "k1": "k1",
            "k2": "k2",
            "Is2": "Is2",
            "Is-HS1": "Is-HS1",
            "Is-HS2": "Is-HS2",
            "Transient Bias": "Polarisation transitoire",
            "Zero Seq filt HV": "Filtre séq zéro HT",
            "Zero Seq filt LV": "Filtre séq zéro BT",
            "Ih (2)%>": "Ih (2)%>",
            "Cross blocking": "Blocage croisé",
            "CT Saturation": "Saturation TC",
            "No Gap": "Pas d'écart",
            "5th harm blocked": "5ème harm bloquée",
            "Ih (5)%>": "Ih (5)%>",
            "Circuitry Fail": "Défaillance de circuit",
            "The slope characteristics are defined to provide a stable operation using normal condition (e.g. tap changing, out of zone fault, branch out of zone fault)": "Les caractéristiques de pente sont définies pour fournir un fonctionnement stable en condition normale (par ex. changement de prise, défaut hors zone, branche hors zone de défaut)",
            "Enabled": "Activé",
            "Amplitude matching factor of LV winding": "Facteur d'adaptation d'amplitude de l'enroulement BT",
            "LeastSignificant BitRefers": "Bit de Poids Faible Réfère",
            "MostSignificant BitRefers": "Bit de Poids Fort Réfère",
            "Conventional": "Conventionnel",
            "HV": "HT",
            "LV": "BT",
            "Advance": "Avancé",
            "The Courant différentiel (Idiff) = 1.118 PU and the Courant de polarisation(Ibias)= 10.636 PU": "Le courant différentiel (Idiff) est de 1,118 PU et le courant de polarisation (Ibias) est de 10,636 PU.",
            "@ +10% Tap": "à +10 % de prise",
            "Primaire Amps at +10% Tap": "Intensité au primaire à +10 % de prise",
            "Secondaire Amps at +10% Tap": "Intensité au secondaire à +10 % de prise",
            "LV winding through fault current in Primary Amps": "Courant de défaut traversant de l'enroulement BT en ampères primaires",
            "LV winding through fault current in Secondary Amps": "Courant de défaut traversant de l'enroulement BT en ampères secondaires",
            "LV winding amplitude corrected current for through fault in PU": "Courant corrigé en amplitude de l'enroulement BT en cas de défaut traversant en PU",
            "The Courant différentiel (Idiff) = 0.916 PU and the Courant de polarisation(Ibias)= 9.619PU": "Le courant différentiel (Idiff) est de 0,916 PU et le courant de polarisation (Ibias) est de 9,619 PU.",
            "Operate value of the differential Protection Fonction as referred to the Secondaire current of the relevant transformer end": "Valeur de fonctionnement de la fonction de protection différentielle rapportée au courant secondaire de l'extrémité du transformateur concerné",
            "Gradient of the Caractéristique de déclenchement of differential Protection in the range": "Gradient de la caractéristique de déclenchement de la protection différentielle dans la plage",
            "The lower slope provide sensitivity for internal faults. Under normal operation steady state magnetizing Courant and the use of tap changers result in unbalanced conditions and hence Courant différentiel. To accommodate these conditions the initial slope, K1 is set to 30%. This Réglage ensures sensitivity to faults while allowing for mismatch when the power transformer is at the limit of its tap range and Rapport TC errors.": "La pente inférieure assure la sensibilité pour les défauts internes. En fonctionnement normal en régime permanent, le courant magnétisant et l'utilisation des changeurs de prises entraînent des conditions déséquilibrées et donc un courant différentiel. Pour s'adapter à ces conditions, la pente initiale K1 est réglée à 30%. Ce réglage assure la sensibilité aux défauts tout en permettant le désappariement lorsque le transformateur de puissance est à la limite de sa plage de prises et les erreurs de rapport TC.",
            "This setting defines the second knee of the tripping characteristic. Above this knee, the gradient is k2": "Ce réglage définit le second coude de la caractéristique de déclenchement. Au-dessus de ce coude, le gradient est k2",
            "If the threshold is set too high, it is possible for the P64x not to trip in the presence of internal faults with transformer saturation)": "Si le seuil est réglé trop haut, il est possible que le P64x ne déclenche pas en présence de défauts internes avec saturation du transformateur",
            "If the threshold is set too low, the P64X can trip in the presence of external faults with transformer saturation": "Si le seuil est réglé trop bas, le P64X peut déclencher en présence de défauts externes avec saturation du transformateur",
            "The stability check is performed for the following cases": "La vérification de stabilité est effectuée pour les cas suivants",
            "Stability of 30MVA Transformer at full load and under operation at -10% Tap": "Stabilité du transformateur 30MVA à pleine charge et en fonctionnement à -10% de prise",
            "Stability of 30MVA Transformer at full load and under operation at +10% Tap": "Stabilité du transformateur 30MVA à pleine charge et en fonctionnement à +10% de prise",
            "Stability of 30MVA Transformer for Through fault on 20kV side when -10% tap at 63kV side": "Stabilité du transformateur 30MVA pour défaut traversant côté 20kV lorsque prise -10% côté 63kV",
            "Stability of 30MVA Transformer for Through fault on 20kV side when +10% tap at 63kV side": "Stabilité du transformateur 30MVA pour défaut traversant côté 20kV lorsque prise +10% côté 63kV",
            "Differential current": "Courant différentiel",
            "Bias Current": "Courant de polarisation",
            "Considering the above referred setting the operating current calculation": "Considérant le réglage mentionné ci-dessus, le calcul du courant de fonctionnement",
            "The differential current (Idiff) = 0.113 PU and the Bias Current (Ibias) = 1.058 PU": "Le courant différentiel (Idiff) = 0,113 PU et le courant de polarisation (Ibias) = 1,058 PU",
            "The required differential current for operation shall be Idiff = k1 * Ibias as per relay Technical manual as Ibias (1.058 PU) is more than Is1/k1(1 PU) and less than Is2 (1.5 PU) as per the above calculation": "Le courant différentiel requis pour le fonctionnement sera Idiff = k1 * Ibias selon le manuel technique du relais car Ibias (1,058 PU) est supérieur à Is1/k1(1 PU) et inférieur à Is2 (1,5 PU) selon le calcul ci-dessus",
            "Differential current required for operation": "Courant différentiel requis pour le fonctionnement",
            "Based on the above calculation, as the Bias current is 1.058 PU and the differential current is 0.113 PU so, the operating point lies below the characteristic curve as the required differential current for operation is 0.212 PU, hence the relay shall remain stable when the transformer is operating at -10% tap at full load condition": "Sur la base du calcul ci-dessus, comme le courant de polarisation est de 1,058 PU et le courant différentiel est de 0,113 PU, le point de fonctionnement se situe en dessous de la courbe caractéristique car le courant différentiel requis pour le fonctionnement est de 0,212 PU, donc le relais restera stable lorsque le transformateur fonctionne à -10% de prise à pleine charge",
            
            # Additional terms from new screenshots
            "Shortcircuit MVA of the Transformer Sref(sc)": "MVA de court-circuit du transformateur Sref(sc)",
            "LV winding through fault current in Primary Amps": "Courant de défaut traversant de l'enroulement BT en ampères primaires",
            "LV winding through fault current in Secondary Amps": "Courant de défaut traversant de l'enroulement BT en ampères secondaires",
            "LV winding amplitude corrected current for through fault in PU": "Courant corrigé en amplitude de l'enroulement BT pour défaut traversant en PU",
            "The required differential current for operation shall be Idiff = Is1 as per relay Technical manual as Ibias (0.956 PU) is less than Is1/k1(1 PU) as per the above calculation.": "Le courant différentiel requis pour le fonctionnement sera Idiff = Is1 selon le manuel technique du relais car Ibias (0,956 PU) est inférieur à Is1/k1(1 PU) selon le calcul ci-dessus.",
            "Based on the above calculation, as the Bias current is 0.956 PU and the differential current is 0.091 PU so, the operating point lies below the characteristic curve as the required differential current for operation is 0.2 PU, hence the relay shall remain stable when the transformer is operating at +10% tap at full load condition.": "Sur la base du calcul ci-dessus, comme le courant de polarisation est de 0,956 PU et le courant différentiel est de 0,091 PU, le point de fonctionnement se situe en dessous de la courbe caractéristique car le courant différentiel requis pour le fonctionnement est de 0,2 PU, donc le relais restera stable lorsque le transformateur fonctionne à +10% de prise à pleine charge.",
            "The required differential current for operation shall be Idiff = k1 * Is2+k2*(Ibias-Is2) as per relay Technical manual as Ibias (10.636 PU) is more than Is2 (1.5 PU) as per the above calculation.": "Le courant différentiel requis pour le fonctionnement sera Idiff = k1 * Is2+k2*(Ibias-Is2) selon le manuel technique du relais car Ibias (10,636 PU) est supérieur à Is2 (1,5 PU) selon le calcul ci-dessus.",
            "Based on the above calculation, as the Bias current is 10.636 PU and the differential current is 1.118 PU so, the operating point lies below the characteristic curve as the required differential current for operation is 7.609 PU, hence the relay shall remain stable when the transformer is operating at -10% tap at through fault condition.": "Sur la base du calcul ci-dessus, comme le courant de polarisation est de 10,636 PU et le courant différentiel est de 1,118 PU, le point de fonctionnement se situe en dessous de la courbe caractéristique car le courant différentiel requis pour le fonctionnement est de 7,609 PU, donc le relais restera stable lorsque le transformateur fonctionne à -10% de prise en condition de défaut traversant.",
            "The required differential current for operation shall be Idiff = k1 * Is2+k2*(Ibias-Is2) as per relay Technical manual as Ibias (9.619 PU) is more than Is2 (1.5 PU) as per the above calculation.": "Le courant différentiel requis pour le fonctionnement sera Idiff = k1 * Is2+k2*(Ibias-Is2) selon le manuel technique du relais car Ibias (9,619 PU) est supérieur à Is2 (1,5 PU) selon le calcul ci-dessus.",
            "Based on the above calculation, as the Bias current is 9.619 PU and the differential current is 0.916 PU so, the operating point lies below the characteristic curve as the required differential current for operation is 6.796 PU, hence the relay shall remain stable when the transformer is operating at +10% tap at through fault condition.": "Sur la base du calcul ci-dessus, comme le courant de polarisation est de 9,619 PU et le courant différentiel est de 0,916 PU, le point de fonctionnement se situe en dessous de la courbe caractéristique car le courant différentiel requis pour le fonctionnement est de 6,796 PU, donc le relais restera stable lorsque le transformateur fonctionne à +10% de prise en condition de défaut traversant.",
            "The differential current required for the relay to operate in Amps": "Le courant différentiel requis pour que le relais fonctionne en ampères",
            
            # Overfluxing protection terms
            "OVERFLUXING PROTECTION (V/F PROTECTION) HV SIDE: 24": "PROTECTION CONTRE LA SURFLUXION (PROTECTION V/F) CÔTÉ HT: 24",
            "HV PT Rating - Primary kV": "Calibre TP HT - Primaire kV",
            "HV PT Rating - secondary V": "Calibre TP HT - secondaire V",
            "Frequency Hz": "Fréquence Hz",
            "Continuous withstand capacity": "Capacité de tenue continue",
            "V/Hz> Alarm status": "État d'alarme V/Hz>",
            "Time delay for V/Hz alarm": "Temporisation pour alarme V/Hz",
            "V/Hz>1 trip status": "État de déclenchement V/Hz>1",
            "V/Hz>1 trip function": "Fonction de déclenchement V/Hz>1",
            "V/Hz>1 trip set": "Réglage de déclenchement V/Hz>1",
            "V/Hz>1 trip delay in secs": "Temporisation de déclenchement V/Hz>1 en secondes",
            "V/Hz>2 trip status": "État de déclenchement V/Hz>2",
            "V/Hz>2 trip set": "Réglage de déclenchement V/Hz>2",
            "V/Hz>2 trip delay in secs": "Temporisation de déclenchement V/Hz>2 en secondes",
            "V/Hz>3 trip status": "État de déclenchement V/Hz>3",
            "V/Hz>3 trip set": "Réglage de déclenchement V/Hz>3",
            "V/Hz>3 trip delay in secs": "Temporisation de déclenchement V/Hz>3 en secondes",
            "V/Hz>4 trip status": "État de déclenchement V/Hz>4",
            "V/Hz>4 trip set": "Réglage de déclenchement V/Hz>4",
            "V/Hz>4 trip delay in secs": "Temporisation de déclenchement V/Hz>4 en secondes",
            "V/Hz>5 trip function": "Fonction de déclenchement V/Hz>5",
            "V/Hz>5 trip set": "Réglage de déclenchement V/Hz>5",
            "V/Hz>5 trip delay in secs": "Temporisation de déclenchement V/Hz>5 en secondes",
            "tPre-Trip Alrm": "tPré-décl. alarme",
            "Note: The above recommended settings are per given inputs.": "Remarque: Les réglages recommandés ci-dessus sont basés sur les données fournies.",
            
            # Earth fault protection terms
            "HV Side (63kV) Standby Earthfault Protection (51NS HV)": "Protection de secours contre les défauts à la terre côté HT (63kV) (51NS HT)",
            "GIVEN DATA:": "DONNÉES FOURNIES:",
            "Relay type used": "Type de relais utilisé",
            "Primary winding (HV) CTR (given)": "CTR de l'enroulement primaire (HT) (donné)",
            "HV winding full load current in Primary Amps": "Courant de pleine charge de l'enroulement HT en ampères primaires",
            "Standby Earth Fault Protection:": "Protection de secours contre les défauts à la terre:",
            "Earth Fault 1": "Défaut à la terre 1",
            "EF 1 Measured": "EF 1, mesuré",
            "IN>1 Direction": "Direction IN>1",
            "Earth Fault Pickup in Primary": "Pickup de défaut à la terre au primaire",
            "Equivalent Pickup Current in Amps": "Courant de pickup équivalent en ampères",
            "NonDirectional": "Non directionnel",
            "IN>1 Delay type": "Type de temporisation IN>1",
            "Time delay in sec": "Temporisation en secondes",
            "LV Side (20kV) Standby Earthfault Protection (51NS LV)": "Protection de secours contre les défauts à la terre côté BT (20kV) (51NS BT)",
            "GIVEN DATA: (LV)": "DONNÉES FOURNIES: (BT)",
            "Transformer Capacity in MVA,LV side": "Capacité du transformateur en MVA, côté BT",
            "% Impedance of the Transformer,Zt": "% d'impédance du transformateur, Zt",
            "LV Voltage rating in kV": "Tension nominale BT en kV",
            "Primary winding (LV) CTR (given)": "CTR de l'enroulement primaire (BT) (donné)",
            "Full Load current 20kV in Amps": "Courant à pleine charge 20kV en ampères",
            "Earth Fault 2": "Défaut à la terre 2",
            "EF 2 Measured": "EF 2, mesuré",
            "IN>2 Direction": "Direction IN>2",
            "IN>2 Delay type": "Type de temporisation IN>2",
            "Note: Standby Earth Fault is independent protection, and this function is recommended with a time delay of 1 s & is expected to operate in last if no other E/F relays are operating and DT characteristics is preferred.": "Remarque: La protection de secours contre les défauts à la terre est une protection indépendante, et cette fonction est recommandée avec une temporisation de 1 s et est prévue pour fonctionner en dernier si aucun autre relais E/F ne fonctionne, et les caractéristiques DT sont préférées.",
            
            # REF protection terms
            "30 MVA TRANSFORMER HV REF PROTECTION": "PROTECTION REF HT DU TRANSFORMATEUR 30 MVA",
            "HV REF PROTECTION: 64R": "PROTECTION REF HT: 64R",
            "Transformer capacity": "Capacité du transformateur",
            "HV Phase side CTR (Given)": "CTR côté phase HT (Donné)",
            "HV-PL CTR - Pri": "CTR HT-PL - Prim",
            "HV-PL CTR - Sec": "CTR HT-PL - Sec",
            "HV Neutral side CTR (Given)": "CTR côté neutre HT (Donné)",
            "HV-N CTR - Pri": "CTR HT-N - Prim",
            "HV-N CTR - Sec": "CTR HT-N - Sec",
            "HV Phase side full load current in Primary Amps": "Courant de pleine charge côté phase HT en ampères primaires",
            "HV Phase side full load current in Secondary Amps": "Courant de pleine charge côté phase HT en ampères secondaires",
            "Amplitude matching factor of HV Phase side": "Facteur d'adaptation d'amplitude du côté phase HT",
            "HV Phase side amplitude corrected current in PU": "Courant corrigé en amplitude du côté phase HT en PU",
            "HV Neutral side full load current in Primary Amps": "Courant de pleine charge côté neutre HT en ampères primaires",
            "HV Neutral side full load current in Secondary Amps": "Courant de pleine charge côté neutre HT en ampères secondaires",
            "Amplitude matching factor of HV Neutral side": "Facteur d'adaptation d'amplitude du côté neutre HT",
            "HV Neutral side amplitude corrected current in PU": "Courant corrigé en amplitude du côté neutre HT en PU",
            "Scaling Factor (K)": "Facteur d'échelle (K)",
            "HV REF Settings :": "Réglages REF HT :",
            "Scaling Factor (K)": "Facteur d'échelle (K)",
            "REF HV Status": "État REF HT",
            "HV Neutral CT": "TC neutre HT",
            "IREF>Is1 HV in Primary": "IREF>Is1 HT au primaire",
            
            # Additional REF terms from new screenshots
            "IREF>Is1 HV in Secondary": "IREF>Is1 HT au secondaire",
            "IREF>Is2 HV": "IREF>Is2 HT",
            "IREF>k1 HV": "IREF>k1 HT",
            "IREF>k2 HV": "IREF>k2 HT",
            "tREF HV": "tREF HT",
            "IH2 REF Block HV": "Blocage IH2 REF HT",
            "IH2 REF Set HV": "Réglage IH2 REF HT",
            "IREF > Is HV": "IREF > Is HT",
            
            "30 MVA TRANSFORMER LV REF PROTECTION": "PROTECTION REF BT DU TRANSFORMATEUR 30 MVA",
            "LV REF PROTECTION: 64R": "PROTECTION REF BT: 64R",
            "LV Phase side CTR": "CTR côté phase BT",
            "LV-PL CTR - Pri": "CTR BT-PL - Prim",
            "LV-PL CTR - Sec": "CTR BT-PL - Sec",
            "LV Neutral side CTR": "CTR côté neutre BT",
            "LV-N CTR - Pri": "CTR BT-N - Prim",
            "LV-N CTR - Sec": "CTR BT-N - Sec",
            "LV Phase side full load current in Primary Amps": "Courant de pleine charge côté phase BT en ampères primaires",
            "LV Phase side full load current in Secondary Amps": "Courant de pleine charge côté phase BT en ampères secondaires",
            "Amplitude matching factor of LV Phase side": "Facteur d'adaptation d'amplitude du côté phase BT",
            "LV Phase side amplitude corrected current in PU": "Courant corrigé en amplitude du côté phase BT en PU",
            "LV Neutral side full load current in Primary Amps": "Courant de pleine charge côté neutre BT en ampères primaires",
            "LV Neutral side full load current in Secondary Amps": "Courant de pleine charge côté neutre BT en ampères secondaires",
            "Amplitude matching factor of LV Neutral side": "Facteur d'adaptation d'amplitude du côté neutre BT",
            "LV Neutral side amplitude corrected current in PU": "Courant corrigé en amplitude du côté neutre BT en PU",
            "LV REF Settings :": "Réglages REF BT :",
            "REF LV Status": "État REF BT",
            "LV Neutral CT": "TC neutre BT",
            "IREF>Is1 LV in Primary": "IREF>Is1 BT au primaire",
            "IREF>Is1 LV in Secondary": "IREF>Is1 BT au secondaire",
            "Minimum settings available in relay": "Réglages minimums disponibles dans le relais",
            "IREF>Is2 LV": "IREF>Is2 BT",
            "IREF>k1 LV": "IREF>k1 BT",
            "IREF>k2 LV": "IREF>k2 BT",
            "tREF LV": "tREF BT",
            "IH2 REF Block LV": "Blocage IH2 REF BT",
            "IH2 REF Set LV": "Réglage IH2 REF BT",
            "IREF > Is LV": "IREF > Is BT",
            
            # HV Overcurrent protection terms
            "30MVA TRANSFORMER HV SIDE (63kV) OVER CURRENT & EARTH FAULT PROTECTION": "PROTECTION CONTRE LES SURINTENSITÉS ET LES DÉFAUTS À LA TERRE CÔTÉ HT (63kV) DU TRANSFORMATEUR 30MVA",
            "Transformer capacity": "Capacité du transformateur",
            "MVA - Rated": "MVA - Nominal",
            "Vr-HV": "Vr-HT",
            "Vr-LV": "Vr-BT",
            "HV winding CT Ratio (Given)": "Rapport TC de l'enroulement HT (Donné)",
            "CTR-Pri(HV)": "CTR-Prim(HT)",
            "CTR-Sec(HV)": "CTR-Sec(HT)",
            "% Impedance of the Transformer": "% d'impédance du transformateur",
            "Vector group": "Groupe de vecteurs",
            "HV winding Full load current in Primary": "Courant de pleine charge de l'enroulement HT au primaire",
            "HV winding Full load current in Secondary": "Courant de pleine charge de l'enroulement HT au secondaire",
            "HV winding through fault current in Primary": "Courant de défaut traversant de l'enroulement HT au primaire",
            "HV winding through fault current in Secondary": "Courant de défaut traversant de l'enroulement HT au secondaire",
            "HV winding Over current protection": "Protection contre les surintensités de l'enroulement HT",
            "Stage - 1 (I>1) Non - Directional Over current: 51": "Étage - 1 (I>1) Surintensité non directionnelle: 51",
            "I>1 Status": "État I>1",
            "I>1 Direction": "Direction I>1",
            "Stage 1 Over current pickup in primary": "Pickup de surintensité étape 1 au primaire",
            "Stage 1 Over current pickup in secondary": "Pickup de surintensité étape 1 au secondaire",
            "Tripping characteristic for I>1": "Caractéristique de déclenchement pour I>1",
            "Curve type": "Type de courbe",
            "LV (Down stream) Operate time (Assumed)": "Temps de fonctionnement BT (en aval) (Supposé)",
            "OT(ds)": "TF(ds)",
            "Grading Margin": "Marge d'échelonnement",
            "T(GM)": "T(ME)",
            "Required operating time": "Temps de fonctionnement requis",
            "treq": "treq",
            "Fault current considered for calculation": "Courant de défaut considéré pour le calcul",
            "Operating time @ TMS = 1": "Temps de fonctionnement @ TMS = 1",
            "t1": "t1",
            
            # Additional overcurrent and earth fault protection terms
            "TMS": "TMS",
            "Stage - 2 (I>2) Non directional Over current: 50": "Étage - 2 (I>2) Surintensité non directionnelle: 50",
            "I>2 Status": "État I>2",
            "I>2 Direction": "Direction I>2",
            "Stage 2 Over current pickup in primary": "Pickup de surintensité étape 2 au primaire",
            "Stage 2 Over current pickup in secondary": "Pickup de surintensité étape 2 au secondaire",
            "Tripping characteristic for I>2": "Caractéristique de déclenchement pour I>2",
            "Operate time delay": "Temporisation de fonctionnement",
            "tI>2": "tI>2",
            "HV winding Earth fault protection": "Protection contre les défauts à la terre de l'enroulement HT",
            "Stage - 1 (IN>1) Non - Directional Earth fault: 51N": "Étage - 1 (IN>1) Défaut à la terre non directionnel: 51N",
            "IN>1 Status": "État IN>1",
            "IN>1 Direction": "Direction IN>1",
            "Stage 1 Earth fault pickup in primary": "Seuil de déclenchement défaut à la terre étage 1 au primaire",
            "Stage 1 Earth fault pickup in secondary": "Pickup de défaut à la terre étape 1 au secondaire",
            "Minimum settings range available in relay": "Plage de réglages minimum disponible dans le relais",
            "Note: The minimum setting range available in the IED is 80mA. So we are recommending 80mA to be adopted.": "Remarque: La plage de réglage minimum disponible dans l'IED est de 80mA. Nous recommandons donc d'adopter 80mA.",
            "Tripping characteristic for IN>1": "Caractéristique de déclenchement pour IN>1",
            "Note: The IDMT curves shall saturate if the fault current is more than 20 times of pickup current. As in this case the actual fault current is 2768.671 A which is more than 20 times of pickup current hence for calculation purpose we shall consider 20*55 = 1100 A": "Remarque: Les courbes IDMT saturent si le courant de défaut est supérieur à 20 fois le courant de pickup. Dans ce cas, le courant de défaut réel est de 2768,671 A, ce qui est supérieur à 20 fois le courant de pickup, donc pour les besoins du calcul, nous considérerons 20*55 = 1100 A",
            "Stage - 2 (IN>2) Non directional Earth fault: 50N": "Étage - 2 (IN>2) Défaut à la terre non directionnel: 50N",
            "IN>2 Status": "État IN>2",
            "IN>2 Direction": "Direction IN>2",
            "Stage 2 Earth fault pickup in primary": "Pickup de défaut à la terre étape 2 au primaire", 
            "Stage 2 Earth fault pickup in secondary": "Pickup de défaut à la terre étape 2 au secondaire",
            "Earth Fault Pickup in Primary": "Pickup de défaut à la terre au primaire",
            "Tripping characteristic for IN>2": "Caractéristique de déclenchement pour IN>2",
            "tIN>2": "tIN>2",
            "100% of Saturated Fault Current": "100% du courant de défaut saturé",
            "TMS (I>1 TMS)": "TMS (I>1 TMS)",
            "I>1 Reset Char": "Caractéristique de réinitialisation I>1",
            "I>1 tReset": "tRéinitialisation I>1",
            "Phase IOC (I>1) Non directional Over current: 50": "IOC de phase (I>1) Surintensité non directionnelle: 50",
            "I>1 Function": "Fonction I>1",
            "I>1 Input": "Entrée I>1",
            "Phasor": "Phaseur",
            "Operate time delay (I>1 Time Delay)": "Temporisation de fonctionnement (Temporisation I>1)",
            "tI>1": "tI>1",
            "EF1- TOC (IN>1) Non - Directional Earth fault: 51N": "EF1- TOC (IN>1) Défaut à la terre non directionnel: 51N",
            "IN1>1 Function": "Fonction IN1>1",
            "IN1>1 Direction": "Direction IN1>1",
            "Stage 1 Earth fault pickup in secondary (IN1>1 Current)": "Pickup de défaut à la terre étape 1 au secondaire (Courant IN1>1)",
            "Tripping characteristic for IN>1": "Caractéristique de déclenchement pour IN>1",
            "Curve type (IN1>1 Curve)": "Type de courbe (Courbe IN1>1)",
            "TMS (IN1>1 TMS)": "TMS (TMS IN1>1)",
            "IN1>1 Reset Char": "Caractéristique de réinitialisation IN1>1",
            "IN1>1 tReset": "tRéinitialisation IN1>1",
            "EF1-IOC (IN>1) Non directional Earth fault: 50N": "EF1-IOC (IN>1) Défaut à la terre non directionnel: 50N",
            "Stage 2 Earth fault pickup in secondary (IN1>1 Current Set)": "Pickup de défaut à la terre étape 2 au secondaire (Réglage de courant IN1>1)",
            "Operate time delay (IN1>1 Time Delay)": "Temporisation de fonctionnement (Temporisation IN1>1)",
            "tIN>1": "tIN>1",
            "EF2-IOC (IN>2) Tank Protection": "EF2-IOC (IN>2) Protection de cuve",
            "CT Ratio primary in A (given)": "Rapport TC primaire en A (donné)",
            "CT Ratio secondary in A": "Rapport TC secondaire en A",
            "IN2>1 Direction": "Direction IN2>1",
            "20% of IFLC (CTR)": "20% de IFLC (CTR)",
            "IN2>1 Pickup in primary": "Pickup IN2>1 au primaire",
            "Pickup in secondary (IN2>2 Current Set)": "Pickup au secondaire (Réglage de courant IN2>2)",
            "Operate time delay (IN2>2 Time Delay)": "Temporisation de fonctionnement (Temporisation IN2>2)",
            "tIN2>1": "tIN2>1",
            
            # Additional technical abbreviations and terms
            "IFLC (HV)-Iref(p)": "IFLC (HT)-Iref(p)",
            "IFLC (HV)-Iref(s)": "IFLC (HT)-Iref(s)",
            "Kamp (HV)": "Kamp (HT)",
            "Iacc (HV)-PU": "Iacc (HT)-PU",
            "IFLC (LV)-Iref(p)": "IFLC (BT)-Iref(p)",
            "IFLC (LV)-Iref(s)": "IFLC (BT)-Iref(s)",
            "Kamp (LV)": "Kamp (BT)",
            "Iacc (LV)": "Iacc (BT)",
            "The Transformer Differential Protection Relais MiCOM P643- 87T has the following parameters for": "Le relais de protection différentielle de transformateur MiCOM P643-87T a les paramètres suivants pour",
            "Grounded": "Mis à la terre",
            
            # Additional buscoupler terms
            "SETTING CALCULATION DOCUMENT FOR": "DOCUMENT DE CALCUL DE RÉGLAGE POUR",
            "63kV BUSCOUPLER PROTECTION": "PROTECTION DE COUPLEUR DE BARRES 63kV",
            
            # Units
            "kV": "kV",
            "A": "A",
            "s": "s",
            "PU": "PU",
            "MVA": "MVA",
            "V/Hz": "V/Hz",
            "sec": "sec",
            "Hz": "Hz",
            "ms": "ms",
            
            # New terms from 63kV PAC-ANCIEN PONT line protection screenshots
            "63kV PAC-ANCIEN PONT LINE PROTECTION": "PROTECTION DE LIGNE PAC-ANCIEN PONT 63kV",
            "DETAILS OF PROTECTED LINE:": "DÉTAILS DE LA LIGNE PROTÉGÉE :",
            "Length of protected line": "Longueur de la ligne protégée",
            "Positive Sequence Resistance per km": "Résistance de séquence positive par km",
            "Positive Sequence Reactance per km": "Réactance de séquence positive par km",
            "Positive sequence impedance per km": "Impédance de séquence positive par km",
            "Zero Sequence Resistance per km": "Résistance de séquence zéro par km",
            "Zero Sequence Reactance per km": "Réactance de séquence zéro par km",
            "Zero Sequence Impedance per km": "Impédance de séquence zéro par km",
            "Positive Sequence Resistance for Protected line": "Résistance de séquence positive pour la ligne protégée",
            "Positive Sequence Reactance for Protected line": "Réactance de séquence positive pour la ligne protégée",
            "Positive Sequence Impedance for Protected line": "Impédance de séquence positive pour la ligne protégée",
            "Positive sequence Impedance angle": "Angle d'impédance de séquence positive",
            "Zero Sequence Resistance for Protected line": "Résistance de séquence zéro pour la ligne protégée",
            "Zero Sequence Reactance for Protected line": "Réactance de séquence zéro pour la ligne protégée",
            "Zero Sequence Impedance for Protected line": "Impédance de séquence zéro pour la ligne protégée",
            "Zero Sequence Impedance Angle": "Angle d'impédance de séquence zéro",
            "DETAILS OF ADJACENT SHORT LINE:": "DÉTAILS DE LA LIGNE ADJACENTE COURTE :",
            "Length of the adjacent line": "Longueur de la ligne adjacente",
            "Pos.Seq. Resistance of Adj. Line per km": "Résistance séq. pos. de la ligne adj. par km",
            "Pos.Seq. Reactance of Adj. Line per km": "Réactance séq. pos. de la ligne adj. par km",
            "Pos.Seq. Impedance of Adj. Line per km": "Impédance séq. pos. de la ligne adj. par km",
            "Positive Sequence Resistance for adj.line": "Résistance de séquence positive pour ligne adj.",
            "Positive Sequence Reactance for adj.line": "Réactance de séquence positive pour ligne adj.",
            "Positive Sequence Impedance for adj.line": "Impédance de séquence positive pour ligne adj.",
            "DETAILS OF ADJACENT LONGEST LINE:": "DÉTAILS DE LA LIGNE ADJACENTE LA PLUS LONGUE :",
            "Length of the adjacent longest line": "Longueur de la ligne adjacente la plus longue",
            "30MVA TRAFO DETAILS:": "DÉTAILS TRAFO 30MVA :",
            "Transformer impedance": "Impédance du transformateur",
            "Transformer rating in MVA": "Puissance nominale du transformateur en MVA",
            "No. of Circuits": "Nombre de circuits",
            "Voltage rating": "Tension nominale",
            "Impedance of Transformer = ((KV)^2/MVA ) x Percentage impedance": "Impédance du transformateur = ((KV)^2/MVA ) x Pourcentage d'impédance",
            "Resultance Impedance when both the Transformers are connected in parallel": "Impédance résultante quand les deux transformateurs sont connectés en parallèle",
            
            # CT & PT Details
            "CT & PT Details": "Détails TC et TP",
            "CT Primary (From Scheme)": "TC Primaire (Du schéma)",
            "CT Secondary": "TC Secondaire",
            "PT Primary": "TP Primaire",
            "PT Secondary": "TP Secondaire",
            "CT/PT Ratio": "Rapport TC/TP",
            
            # Ground Factor
            "GROUND FACTOR": "FACTEUR DE TERRE",
            "Grounding Factor kg = (Z0-Z1)/(3*Z1)": "Facteur de mise à la terre kg = (Z0-Z1)/(3*Z1)",
            "Magnitude of grounding factor Kz": "Magnitude du facteur de mise à la terre Kz",
            "Angle of Grounding factor Kz": "Angle du facteur de mise à la terre Kz",
            "KZ1,KZ2,KZ3....PS1": "KZ1,KZ2,KZ3....PS1",
            "Angle Kz": "Angle Kz",
            "Phase Sequence": "Séquence de phases",
            "Tripping Mode": "Mode de déclenchement",
            "Line Charging Y": "Charge de ligne Y",
            "Standard ABC": "ABC standard",
            "1 and 3 Pole": "1 et 3 pôles",
            
            # Zone Settings
            "ZONE SETTINGS": "RÉGLAGES DE ZONE",
            "DIST Z1": "DIST Z1",
            "DIST Z2": "DIST Z2",
            "DIST Z3": "DIST Z3",
            "DIST Z4": "DIST Z4",
            "100% of protected line +": "100% de la ligne protégée +",
            "120% of protected line +": "120% de la ligne protégée +",
            "of protected line Impedance": "de l'impédance de la ligne protégée",
            "of shortest line Impedance": "de l'impédance de la ligne la plus courte",
            "of longest line Impedance": "de l'impédance de la ligne la plus longue",
            
            # Operating Time Delay
            "OPERATING TIME DELAY": "TEMPORISATION DE FONCTIONNEMENT",
            "ZONE 1 OPERATING TIME (t1)": "TEMPS DE FONCTIONNEMENT ZONE 1 (t1)",
            "ZONE 2 OPERATING TIME (t2)": "TEMPS DE FONCTIONNEMENT ZONE 2 (t2)",
            "ZONE 3 OPERATING TIME (t3)": "TEMPS DE FONCTIONNEMENT ZONE 3 (t3)",
            "ZONE 4 OPERATING TIME (t4)": "TEMPS DE FONCTIONNEMENT ZONE 4 (t4)",
            
            # Distance Setup
            "DISTANCE SETUP": "CONFIGURATION DISTANCE",
            "Setting Mode": "Mode de réglage",
            "Advanced": "Avancé",
            
            # Phase Distance
            "PHASE DISTANCE": "DISTANCE DE PHASE",
            "Phase Chars": "Caractéristiques de phase",
            "Zone 1 Ph Status": "État phase zone 1",
            "Zone 2 Ph Status": "État phase zone 2",
            "Zone 3 Ph Status": "État phase zone 3",
            "Zone 3 Ph Dir": "Direction phase zone 3",
            "Zone P Ph Status": "État phase zone P",
            "Zone 4 Ph Status": "État phase zone 4",
            "Quad": "Quad",
            "Enabled": "Activé",
            "Forward": "Avant",
            "Disabled": "Désactivé",
            
            # Ground Distance
            "GROUND DISTANCE": "DISTANCE DE TERRE",
            "Ground Chars": "Caractéristiques de terre",
            "Zone 1 Gnd Stat": "État terre zone 1",
            "Zone 2 Gnd Stat": "État terre zone 2",
            "Zone 3 Gnd Stat": "État terre zone 3",
            "Zone 3 Gnd Dir": "Direction terre zone 3",
            "Zone P Gnd Stat": "État terre zone P",
            "Zone 4 Gnd Stat": "État terre zone 4",
            
            # Digital Filter and CVT
            "Digital Filter": "Filtre numérique",
            "CVT Filters": "Filtres CVT",
            "Load Blinders": "Limiteurs de charge",
            "Standard": "Standard",
            
            # Delta Directional
            "DELTA DIRECTIONAL": "DELTA DIRECTIONNEL",
            "Dir Status": "État directionnel",
            "Aided Delta Status": "État delta assisté",
            "Dir. Char Angle": "Angle caract. directionnel",
            "Dir. V Fwd": "V directionnel avant",
            "Dir. V Rev": "V directionnel arrière",
            "Dir. I Fwd": "I directionnel avant",
            "Dir. I Rev": "I directionnel arrière",
            "Ph & Gnd": "Ph et Terre",
            "Volts": "Volts",
            "mA": "mA",
            "Note: The Delta directional settings are given as per general practice. However it can be modified according to the site requirements.": "Remarque : Les réglages directionnels delta sont donnés selon la pratique générale. Cependant, ils peuvent être modifiés selon les exigences du site.",
            
            # Scheme Logic
            "SCHEME LOGIC": "LOGIQUE DE SCHÉMA",
            "Basic Scheme Mode": "Mode de schéma de base",
            "Zone1 Tripping": "Déclenchement zone 1",
            "tZ1 Ph. Delay": "Temporisation phase tZ1",
            "tZ1 Gnd. Delay": "Temporisation terre tZ1",
            "Zone2 Tripping": "Déclenchement zone 2",
            "tZ2 Ph. Delay": "Temporisation phase tZ2",
            "tZ2 Gnd. Delay": "Temporisation terre tZ2",
            "Zone3 Tripping": "Déclenchement zone 3",
            "tZ3 Ph. Delay": "Temporisation phase tZ3",
            "tZ3 Gnd. Delay": "Temporisation terre tZ3",
            "Zone P Tripping": "Déclenchement zone P",
            "Zone 4 Tripping": "Déclenchement zone 4",
            "tZ4 Ph. Delay": "Temporisation phase tZ4",
            "tZ4 Gnd. Delay": "Temporisation terre tZ4",
            
            # Aided Scheme
            "Aided Scheme 1": "Schéma assisté 1",
            "Aid 1 Selection": "Sélection aide 1",
            "Aid 1 Distance": "Distance aide 1",
            "Aid 1 Dist. Dly": "Tempor. dist. aide 1",
            "Aid 1 DEF": "DEF aide 1",
            "Aid 1 DEF Delay": "Temporisation DEF aide 1",
            "Aid 1 DEF Trip": "Déclenchement DEF aide 1",
            "Aid 1 Delta": "Delta aide 1",
            "Aid 1 Delta Dly": "Tempor. delta aide 1",
            "Aid1 DeltaTrip": "Déclenchement delta aide 1",
            "tRev. Guard": "Garde tRev.",
            "Send on Trip": "Envoi sur déclenchement",
            "Weak Infeed": "Injection faible",
            "VI V<Thresh": "VI V<Seuil",
            "Aid 2 Selection": "Sélection aide 2",
            "POR": "POR",
            "1 and 3 Pole": "1 et 3 pôles",
            "None": "Aucun",
            "Echo": "Écho",
            "Volt": "Volt",
            
            # Trip on Close
            "TRIP ON CLOSE": "DÉCLENCHEMENT À LA FERMETURE",
            "SOTF Status": "État SOTF",
            "SOTF Delay": "Temporisation SOTF",
            "SOTF Tripping": "Déclenchement SOTF",
            "TOR Status": "État TOR",
            "TOR Tripping": "Déclenchement TOR",
            "TOC Reset Delay": "Temporisation de réinitialisation TOC",
            "TOC Delay": "Temporisation TOC",
            "En Pdead + Pulse": "En Pdead + Pulse",
            "Z1 & Z2": "Z1 et Z2",
            "msec": "msec",
            
            # Z1 Extension and Loss of Load
            "Z1 Extension": "Extension Z1",
            "Z1 Ext Scheme": "Schéma extension Z1",
            "LOSS OF LOAD": "PERTE DE CHARGE",
            "LOL Scheme": "Schéma LOL",
            
            # Zone Reach Settings
            "ZONE REACH SETTINGS": "RÉGLAGES DE PORTÉE DE ZONE",
            "ZONE 1": "ZONE 1",
            "80% of Protected line impedance": "80% de l'impédance de la ligne protégée",
            "ZONE 2": "ZONE 2",
            "100% of the protected line + 50% of shortest line": "100% de la ligne protégée + 50% de la ligne la plus courte",
            "ZONE 3": "ZONE 3",
            "120% of the protected line + 120% of longest line": "120% de la ligne protégée + 120% de la ligne la plus longue",
            "ZONE 4": "ZONE 4",
            "20% of protected line impedance": "20% de l'impédance de la ligne protégée",
            "VALUES": "VALEURS",
            "Ohm": "Ohm",
            
            # Resistive Reach Setting
            "RESISTIVE REACH SETTING:": "RÉGLAGE DE PORTÉE RÉSISTIVE :",
            "Thermal withstand capacity of the line": "Capacité de tenue thermique de la ligne",
            "Amp": "Amp",
            "As per the recommendations of Ramakrishna committee 50% overloading condition is considered and during this over load voltage drop of 15% is considered for calculating the Load Resistance (RL).": "Selon les recommandations du comité Ramakrishna, une condition de surcharge de 50% est considérée et pendant cette surcharge, une chute de tension de 15% est considérée pour calculer la résistance de charge (RL).",
            "Load Resistance (RL) during over load condition in primary Ohms": "Résistance de charge (RL) pendant la condition de surcharge en ohms primaires",
            "Load Resistance (RL) in Secondary": "Résistance de charge (RL) au secondaire",
            "DIST R1,PP(Quad) Ps1 (2*60% of RL)": "DIST R1,PP(Quad) Ps1 (2*60% de RL)",
            "DIST R2,PP(Quad) Ps1 (2*60% of RL)": "DIST R2,PP(Quad) Ps1 (2*60% de RL)",
            "DIST R3 & R4,PP(Quad) Ps1 (2*60% of RL)": "DIST R3 et R4,PP(Quad) Ps1 (2*60% de RL)",
            "Note: The relay algorithm considers the value entered in the relay as +Rph & -Rph i.e +Rph consider 0.5 of the entered value & -Rph consider 0.25 of the entered value. Hence the value to be entered in the relay.": "Remarque : L'algorithme du relais considère la valeur saisie dans le relais comme +Rph et -Rph, c'est-à-dire +Rph considère 0,5 de la valeur saisie et -Rph considère 0,25 de la valeur saisie. D'où la valeur à saisir dans le relais.",
            "DIST R1,PG(Quad) Ps1(80% of RL)": "DIST R1,PG(Quad) Ps1(80% de RL)",
            "DIST R2,PG(Quad) Ps1 (80% of RL)": "DIST R2,PG(Quad) Ps1 (80% de RL)",
            "DIST R3 & R4,PG(Quad) Ps1 (80% of RL)": "DIST R3 et R4,PG(Quad) Ps1 (80% de RL)",
            
            # Distance Elements
            "DIST ELEMENTS": "ÉLÉMENTS DIST",
            "Z1 Ph. Reach": "Portée phase Z1",
            "Z1 Ph. Angle": "Angle phase Z1",
            "R1 Ph. Resistive": "R1 phase résistif",
            "Z1 Tilt Top Line": "Ligne supérieure inclinée Z1",
            "Z1 Sensit. Iph>1": "Sensibilité Z1 Iph>1",
            "Degree": "Degré",
            
            # Additional Zone Distance Elements (Z2, Z3, Z4)
            "Z2 Ph. Reach": "Portée phase Z2",
            "Z2 Ph. Angle": "Angle phase Z2",
            "R2 Ph. Resistive": "R2 phase résistif",
            "Z2 Tilt Top Line": "Ligne supérieure inclinée Z2",
            "Z2 Sensit. Iph>1": "Sensibilité Z2 Iph>1",
            "Z3 Ph. Reach": "Portée phase Z3",
            "Z3 Ph. Angle": "Angle phase Z3",
            "R3 Ph. Resistive": "R3 phase résistif",
            "Z3 Tilt Top Line": "Ligne supérieure inclinée Z3",
            "Z3 Sensit. Iph>1": "Sensibilité Z3 Iph>1",
            "Z4 Ph. Reach": "Portée phase Z4",
            "Z4 Ph. Angle": "Angle phase Z4",
            "R4 Ph. Resistive": "R4 phase résistif",
            "Z4 Tilt Top Line": "Ligne supérieure inclinée Z4",
            "Z4 Sensit. Iph>1": "Sensibilité Z4 Iph>1",
            
            # Ground Distance Elements
            "GROUND DISTANCE": "DISTANCE DE TERRE",
            "Z1 Gnd. Reach": "Portée terre Z1",
            "Z1 Gnd. Angle": "Angle terre Z1",
            "Z1 Dynamic Tilt": "Inclinaison dynamique Z1",
            "kZN1 Res. Comp.": "Comp. rés. kZN1",
            "kZN1 Res. Angle": "Angle rés. kZN1",
            "R1 Gnd. Resistive": "R1 terre résistif",
            "Z1 Sensit. Ignd>1": "Sensibilité Z1 Ignd>1",
            "Z2 Gnd. Reach": "Portée terre Z2",
            "Z2 Gnd. Angle": "Angle terre Z2",
            "Z2 Dynamic Tilt": "Inclinaison dynamique Z2",
            "kZN2 Res. Comp.": "Comp. rés. kZN2",
            "kZN2 Res. Angle": "Angle rés. kZN2",
            "R2 Gnd. Resistive": "R2 terre résistif",
            "Z2 Sensit. Ignd>1": "Sensibilité Z2 Ignd>1",
            "Z3 Gnd. Reach": "Portée terre Z3",
            "Z3 Gnd. Angle": "Angle terre Z3",
            "Z3 Dynamic Tilt": "Inclinaison dynamique Z3",
            "kZN3 Res. Comp.": "Comp. rés. kZN3",
            "kZN3 Res. Angle": "Angle rés. kZN3",
            "R3 Gnd. Resistive": "R3 terre résistif",
            "Z3 Sensit. Ignd>1": "Sensibilité Z3 Ignd>1",
            "Z4 Gnd. Reach": "Portée terre Z4",
            "Z4 Gnd. Angle": "Angle terre Z4",
            "Z4 Dynamic Tilt": "Inclinaison dynamique Z4",
            "kZN4 Res. Comp.": "Comp. rés. kZN4",
            "kZN4 Res. Angle": "Angle rés. kZN4",
            "R4 Gnd. Resistive": "R4 terre résistif",
            "Z4 Sensit. Ignd>1": "Sensibilité Z4 Ignd>1",
            "Mem Volt Dura": "Durée tension mémoire",
            
            # Power Swing Settings
            "POWER SWING SETTING: 68": "RÉGLAGE OSCILLATION DE PUISSANCE : 68",
            "PSB Status": "État PSB",
            "Zone 1 Ph. PSB": "PSB phase zone 1",
            "Zone 2 Ph. PSB": "PSB phase zone 2",
            "Zone 3 Ph. PSB": "PSB phase zone 3",
            "Zone P Ph. PSB": "PSB phase zone P",
            "Zone 4 Ph. PSB": "PSB phase zone 4",
            "Zone 1 Gnd. PSB": "PSB terre zone 1",
            "Zone 2 Gnd. PSB": "PSB terre zone 2",
            "Zone 3 Gnd. PSB": "PSB terre zone 3",
            "Zone P Gnd. PSB": "PSB terre zone P",
            "Zone 4 Gnd. PSB": "PSB terre zone 4",
            "WI Trip PSB": "Déclenchement WI PSB",
            "PSB Unblocking": "Déverrouillage PSB",
            "PSB Unblock delay": "Temporisation déverrouillage PSB",
            "PSB Reset delay": "Temporisation réinitialisation PSB",
            "OST Mode": "Mode OST",
            "Slow Swing": "Oscillation lente",
            "UnBlocking": "Déverrouillage",
            "Blocking": "Verrouillage",
            "Inhibit Trip": "Inhiber déclenchement",
            "The Power swing detection uses an impedance band which surrounds the entire phase fault trip characteristic. Typically the ΔR and ΔX band settings are both set between 10% to 30% of R3ph": "La détection d'oscillation de puissance utilise une bande d'impédance qui entoure toute la caractéristique de déclenchement de défaut de phase. Typiquement, les réglages de bande ΔR et ΔX sont tous deux fixés entre 10% et 30% de R3ph",
            "Considering 20% of the R3ph ΔR and ΔX in Ohms": "Considérant 20% des ΔR et ΔX R3ph en ohms",
            "PSB Z7 in Ohms": "PSB Z7 en ohms",
            "PSB Z8 in Ohms (1.1*Z7)": "PSB Z8 en ohms (1,1*Z7)",
            "PSB Z7' in Ohms": "PSB Z7' en ohms",
            "PSB Z8' in Ohms (1.1*Z7')": "PSB Z8' en ohms (1,1*Z7')",
            "PSB R7 in Ohms": "PSB R7 en ohms",
            "PSB R8 in Ohms (1.1*R7)": "PSB R8 en ohms (1,1*R7)",
            "PSB R7' in Ohms": "PSB R7' en ohms",
            "PSB R8' in Ohms (1.1*R7')": "PSB R8' en ohms (1,1*R7')",
            "Zone 3 Forward reach + ΔR": "Portée avant zone 3 + ΔR",
            "-(Zone 4 (Reverse) reach + ΔR)": "-(Portée zone 4 (inverse) + ΔR)",
            "((R3,PP/2) + ΔR": "((R3,PP/2) + ΔR",
            "-((R3,PP/2) + ΔR)": "-((R3,PP/2) + ΔR)",
            
            # Broken Conductor Protection
            "BROKEN CONDUCTOR PROTECTION: (Note: Only Alarm and not for trip)": "PROTECTION CONDUCTEUR CASSÉ : (Remarque : Alarme seulement et non pour déclenchement)",
            "Broken conductor": "Conducteur cassé",
            "I2/I1 setting": "Réglage I2/I1",
            "I2/I1 Time Delay": "Temporisation I2/I1",
            
            # VT Supervision
            "VT SUPERVISION: VTS": "SUPERVISION TP : VTS",
            "VTS Mode": "Mode VTS",
            "VTS Status": "État VTS",
            "VTS Reset Mode": "Mode réinitialisation VTS",
            "VTS Time Delay": "Temporisation VTS",
            "VTS I> Inhibit": "Inhibition I> VTS",
            "VTS I2> Inhibit": "Inhibition I2> VTS",
            "Inrush Detection": "Détection d'appel",
            "Measured Only": "Mesure seulement",
            "Auto": "Auto",
            
            # Over Voltage Protection
            "OVER VOLTAGE PROTECTION: 59": "PROTECTION SURTENSION : 59",
            "V> Measur't Mode": "Mode mesure V>",
            "V> Operate Mode": "Mode fonctionnement V>",
            "V>1 Function": "Fonction V>1",
            "V>1 Voltage Set": "Réglage tension V>1",
            "V>1 Time Delay": "Temporisation V>1",
            "V>2 Status": "État V>2",
            "V>2 Voltage Set": "Réglage tension V>2",
            "V>2 Time Delay": "Temporisation V>2",
            "V>1 & V>2 Ph-Ph": "V>1 et V>2 Ph-Ph",
            "V>1 & V>2 Any Ph": "V>1 et V>2 toute Ph",
            "DT": "TD",
            "121 V (110%)": "121 V (110%)",
            "154 V (140%)": "154 V (140%)",
            
            # Under Voltage Protection
            "UNDER VOLTAGE PROTECTION: 27": "PROTECTION SOUS-TENSION : 27",
            "V< Measur't Mode": "Mode mesure V<",
            "V<1 Function": "Fonction V<1",
            "V<1 Voltage Set": "Réglage tension V<1",
            "V<1 Time Delay": "Temporisation V<1",
            "V<2 Status": "État V<2",
            "V<2 Voltage Set": "Réglage tension V<2",
            "V<2 Time Delay": "Temporisation V<2",
            "V<1 & V<2 Ph-Ph": "V<1 et V<2 Ph-Ph",
            "V<1 & V<2 Any Ph": "V<1 et V<2 toute Ph",
            "88 V (80%)": "88 V (80%)",
            "77 V (70%)": "77 V (70%)",
            
            # Directional Phase Overcurrent Protection
            "DIRECTIONAL PHASE OVERCURRENT PROTECTION: 67": "PROTECTION SURINTENSITÉ PHASE DIRECTIONNELLE : 67",
            "Direction of Stage 1": "Direction étage 1",
            "Direction of Stage 2": "Direction étage 2",
            "Directional Forward": "Directionnel avant",
            "Non Directional": "Non directionnel",
            "150% of Full load current": "150% du courant de pleine charge",
            "100% of Fault current considered": "100% du courant de défaut considéré",
            "Pickup current in Amps": "Courant de pickup en ampères",
            "Equivalent secondary in Amps": "Secondaire équivalent en ampères",
            "Delay Type": "Type de temporisation",
            "IDMT": "IDMT",
            "Downstream operating time (Zone3 Time)": "Temps de fonctionnement en aval (Temps zone 3)",
            "Grading margin": "Marge d'échelonnement",
            "Required operating time in sec": "Temps de fonctionnement requis en sec",
            "Maximum fault current in Amps (Given)": "Courant de défaut maximum en ampères (donné)",
            "Note: The IDMT curves shall saturate if the fault current is more than 20 times of pickup current.As in this case the actual fault current is 37800 A which is more than 20 times of pickup current hence for calculation purpose we shall consider 20*412.5=8250 A": "Remarque : Les courbes IDMT saturent si le courant de défaut est supérieur à 20 fois le courant de pickup. Dans ce cas, le courant de défaut réel est de 37800 A, ce qui est supérieur à 20 fois le courant de pickup, donc pour les besoins du calcul, nous considérerons 20*412,5=8250 A",
            "Fault current considered for calculations": "Courant de défaut considéré pour les calculs",
            "Operating time t1 in sec @ TMS = 1": "Temps de fonctionnement t1 en sec @ TMS = 1",
            "I>1 tRESET": "Réinitialisation I>1",
            "I>3 tRESET": "Réinitialisation I>3",
            "I> Char Angle": "Angle caract. I>",
            "I> Blocking": "Blocage I>",
            "Time delay in sec": "Temporisation en sec",
            
            # Directional Earth Fault Protection
            "DIRECTIONAL EARTH FAULT PROTECTION: 67N": "PROTECTION DÉFAUT TERRE DIRECTIONNELLE : 67N",
            "20% of CTR": "20% du CTR",
            "Note: The IDMT curves shall saturate if the fault current is more than 20 times of pickup current.As in this case the actual fault current is 37800 A which is more than 20 times of pickup current hence for calculation purpose we shall consider 20*55=1100 A": "Remarque : Les courbes IDMT saturent si le courant de défaut est supérieur à 20 fois le courant de pickup. Dans ce cas, le courant de défaut réel est de 37800 A, ce qui est supérieur à 20 fois le courant de pickup, donc pour les besoins du calcul, nous considérerons 20*55=1100 A",
            "IN> tRESET": "Réinitialisation IN>",
            "IN> Char Angle": "Angle caract. IN>",
            "IN> Blocking": "Blocage IN>",
            "IN> Polarisation": "Polarisation IN>",
            "Zero Sequence": "Séquence zéro",
            "IN> Vnpol Set": "Réglage Vnpol IN>",
            "5.000 V": "5,000 V",
            "IN3> tRESET": "Réinitialisation IN3>",
            "+45°": "+45°",
            "-45°": "-45°",
            "001110": "001110",
            
            # Frequency Protection
            "OVER FREQUENCY: (81O)": "SURFRÉQUENCE : (81O)",
            "F>1 Status": "État F>1",
            "F>1 setting": "Réglage F>1",
            "F>1 Time delay": "Temporisation F>1",
            "F>2 Status": "État F>2",
            "F>2 setting": "Réglage F>2",
            "F>2 Time delay": "Temporisation F>2",
            "F>3 Status": "État F>3",
            "F>4 Status": "État F>4",
            "50.5 Hz": "50,5 Hz",
            "52 Hz": "52 Hz",
            "1 sec": "1 sec",
            "2 sec": "2 sec",
            
            # Under Frequency Protection
            "UNDER FREQUENCY: (81U)": "SOUS-FRÉQUENCE : (81U)",
            "F<1 Status": "État F<1",
            "F<1 setting": "Réglage F<1",
            "F<1 Time delay": "Temporisation F<1",
            "F<2 Status": "État F<2",
            "F<2 setting": "Réglage F<2",
            "F<2 Time delay": "Temporisation F<2",
            "F<3 Status": "État F<3",
            "F<4 Status": "État F<4",
            "49.5 Hz": "49,5 Hz",
            "48 Hz": "48 Hz",
            "Note - The above recommended settings are as per general practice . However the same shall be modified according to site requirement.": "Remarque - Les réglages recommandés ci-dessus sont selon la pratique générale. Cependant, ils doivent être modifiés selon les exigences du site.",
         }

        # Load custom dictionary if provided
        if dictionary_path and os.path.exists(dictionary_path):
            try:
                custom_dict = pd.read_csv(dictionary_path)
                for _, row in custom_dict.iterrows():
                    self.technical_terms[row['english']] = row['french']
                print(f"Loaded {len(custom_dict)} custom translations")
            except Exception as e:
                print(f"Error loading custom dictionary: {str(e)}")

    def process_pdf(self, input_pdf, output_pdf):
        """Process a PDF file and create a translated version."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False

        try:
            # Use the extract_to_excel and then convert back approach as it's more reliable
            # First, create a temporary Excel file
            temp_excel = f"{output_pdf}_temp.xlsx"
            success = self.extract_to_excel(input_pdf, temp_excel)
            
            if not success:
                print("Failed to extract PDF content")
                return False
                
            # Open the Excel file
            try:
                df = pd.read_excel(temp_excel)
                
                # Create a new PDF document
                doc = fitz.open()
                
                # Group by page number
                page_groups = df.groupby("Page")
                
                # Process each page
                for page_num, group in tqdm(page_groups, desc="Creating PDF pages"):
                    # Create a new page in the output document
                    page = doc.new_page(width=595, height=842)  # A4 size
                    
                    # Sort by position (if available) or just by index
                    sorted_group = group.sort_index()
                    
                    # Initialize y position for text
                    y_pos = 50
                    
                    # Add a header
                    page.insert_text((50, y_pos), f"Page {page_num} - Translated Content", fontsize=16)
                    y_pos += 30
                    
                    # Add content
                    for _, row in sorted_group.iterrows():
                        text_type = row["Type"]
                        french_text = row.get("French", "")
                        
                        if pd.notna(french_text) and french_text.strip():
                            # Add type as a subheader
                            page.insert_text((50, y_pos), f"{text_type}:", fontsize=12, color=(0, 0, 0.8))
                            y_pos += 20
                            
                            # Add text with word wrapping (simple approach)
                            words = french_text.split()
                            line = ""
                            for word in words:
                                test_line = line + " " + word if line else word
                                if len(test_line) > 70:  # Approximate line length
                                    page.insert_text((60, y_pos), line)
                                    y_pos += 15
                                    line = word
                                else:
                                    line = test_line
                            
                            if line:  # Add the last line
                                page.insert_text((60, y_pos), line)
                                y_pos += 25
                            
                            # Add some extra space after each text block
                            y_pos += 10
                            
                            # If near the bottom of the page, create a new page
                            if y_pos > 800:
                                page = doc.new_page(width=595, height=842)
                                y_pos = 50
                
                # Save the new PDF
                doc.save(output_pdf)
                doc.close()
                
                # Clean up the temp file
                try:
                    os.remove(temp_excel)
                except:
                    pass
                
                print(f"Successfully created translated PDF: {output_pdf}")
                return True
                
            except Exception as excel_error:
                print(f"Error processing Excel data: {str(excel_error)}")
                return False
            
        except Exception as e:
            print(f"Error processing PDF: {str(e)}")
            return False
    
    def _process_blocks(self, input_page, output_page, blocks):
        """Process text blocks and add translated content to output page."""
        # Create a drawing context for the output page
        shape = output_page.new_shape()
        
        for block in blocks:
            # Process based on block type
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    line_text = ""
                    for span in line["spans"]:
                        # Get the text and its position
                        text = span["text"]
                        bbox = fitz.Rect(span["bbox"])
                        font_size = span["size"]
                        
                        # Translate the text
                        translated_text = self._translate_text(text)
                        
                        # Add translated text to output page - use a safe default font
                        try:
                            output_page.insert_text(
                                bbox.tl,  # top-left point
                                translated_text,
                                fontname="helv",  # Use built-in Helvetica font
                                fontsize=font_size,
                                color=(0, 0, 0)  # black
                            )
                        except Exception as e:
                            print(f"Warning: Could not add text '{translated_text}': {e}")
            
            elif block["type"] == 1:  # Image block
                # Images are handled separately
                pass
                
            # Copy table structures and borders
            if "lines" in block:
                for line in block.get("lines", []):
                    try:
                        p1 = line["p1"]
                        p2 = line["p2"]
                        shape.draw_line((p1[0], p1[1]), (p2[0], p2[1]))
                    except Exception as e:
                        print(f"Warning: Could not draw line: {e}")
        
        # Commit the shapes to the page
        shape.commit()
    
    def _translate_text(self, text):
        """Translate a piece of text from English to French."""
        if not text or text.strip() == "":
            return text
        
        # Check for direct matches in our technical terms dictionary
        if text in self.technical_terms:
            return self.technical_terms[text]
        
        # Normalize text for case-insensitive matching
        normalized_text = text.strip()
        upper_text = normalized_text.upper()
        
        # Check for case-insensitive match
        for eng, fr in self.technical_terms.items():
            if eng.upper() == upper_text:
                return fr
        
        # Try to translate individual terms within the text
        translated = text
        for eng, fr in sorted(self.technical_terms.items(), key=lambda x: len(x[0]), reverse=True):
            # Skip short terms to avoid incorrect replacements
            if len(eng) < 4:
                continue
                
            pattern = r'\b' + re.escape(eng) + r'\b'
            translated = re.sub(pattern, fr, translated, flags=re.IGNORECASE)
        
        return translated

    def extract_to_excel(self, input_pdf, output_excel):
        """Extract text content from PDF to Excel for translation."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False
            
        try:
            # Open the PDF
            doc = fitz.open(input_pdf)
            
            # Create dataframes to store extracted text
            text_data = []
            table_data = []
            
            print(f"Extracting content from PDF: {input_pdf}")
            
            # Extract text from each page
            for page_num, page in enumerate(tqdm(doc, desc="Extracting pages")):
                # Extract regular text first (more reliable)
                try:
                    # Get text blocks
                    page_text = page.get_text("text")
                    if page_text.strip():
                        # Split by lines
                        lines = page_text.split('\n')
                        for line in lines:
                            if line.strip():
                                french_text = self._translate_text(line.strip())
                                text_data.append({
                                    "Page": page_num + 1,
                                    "Type": "Text",
                                    "English": line.strip(),
                                    "French": french_text
                                })
                except Exception as e:
                    print(f"Warning: Error extracting text from page {page_num + 1}: {e}")
                
                # Try to extract tables as a bonus
                try:
                    tables = self._extract_tables(page)
                    for table in tables:
                        for row in table:
                            table_data.append({
                                "Page": page_num + 1,
                                "Type": "Table",
                                "English": " | ".join(row),
                                "French": " | ".join([self._translate_text(cell) for cell in row])
                            })
                except Exception as e:
                    print(f"Warning: Error extracting tables from page {page_num + 1}: {e}")
            
            # Create DataFrame and save to Excel
            df_text = pd.DataFrame(text_data)
            df_table = pd.DataFrame(table_data)
            
            # Combine into one DataFrame
            if not df_text.empty or not df_table.empty:
                if df_text.empty:
                    df_combined = df_table
                elif df_table.empty:
                    df_combined = df_text
                else:
                    df_combined = pd.concat([df_text, df_table])
                
                # Sort by page number
                df_combined = df_combined.sort_values(by=["Page", "Type"])
                
                # Clean up data to avoid Excel corruption
                for col in df_combined.columns:
                    # Replace any characters that might cause Excel issues
                    if df_combined[col].dtype == 'object':  # Only process string columns
                        df_combined[col] = df_combined[col].apply(
                            lambda x: str(x).replace('\0', '').replace('\r', ' ').replace('\x00', '')
                            if pd.notna(x) else x
                        )
                
                # Save to Excel with more compatible options
                try:
                    # First, try saving with engine='openpyxl' and explicit options
                    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='w') as writer:
                        df_combined.to_excel(writer, sheet_name="Translation", index=False)
                    print(f"Successfully saved extracted content to Excel: {output_excel}")
                    return True
                except Exception as e:
                    print(f"Warning: First Excel export attempt failed: {e}")
                    
                    # Try an alternative approach with fewer formatting options
                    try:
                        # Simplify column names
                        df_combined.columns = [str(col).replace(' ', '_') for col in df_combined.columns]
                        
                        # Save with xlsxwriter which can be more reliable
                        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                            df_combined.to_excel(writer, sheet_name="Translation", index=False)
                        print(f"Successfully saved extracted content to Excel (second attempt): {output_excel}")
                        return True
                    except Exception as e2:
                        print(f"Error saving Excel file (second attempt): {e2}")
                        
                        # If both export methods fail, try CSV as last resort
                        try:
                            csv_path = output_excel.replace('.xlsx', '.csv')
                            df_combined.to_csv(csv_path, index=False, encoding='utf-8-sig')
                            print(f"Exported as CSV instead: {csv_path}")
                            return True
                        except Exception as e3:
                            print(f"All export attempts failed: {e3}")
                            return False
            else:
                print("No content found in PDF")
                return False
            
        except Exception as e:
            print(f"Error extracting to Excel: {str(e)}")
            return False
    
    def _extract_tables(self, page):
        """
        Simple heuristic-based table extraction.
        Returns a list of tables, where each table is a list of rows.
        """
        # This is a simplified placeholder for table extraction
        # In a full implementation, you would need more sophisticated table detection
        tables = []
        
        try:
            # Look for tabular structures
            # For simplicity, we're looking for text aligned in columns
            text = page.get_text("dict")
            
            # Group text by y-position (rows)
            rows = {}
            for block in text["blocks"]:
                if block["type"] == 0:  # Text block
                    for line in block["lines"]:
                        y = int(line["bbox"][1])  # top y-coordinate
                        if y not in rows:
                            rows[y] = []
                        
                        for span in line["spans"]:
                            rows[y].append((span["bbox"][0], span["text"]))  # x-position and text
            
            # Sort rows by y-position
            sorted_rows = [rows[y] for y in sorted(rows.keys())]
            
            # Group consecutive rows that might form a table
            current_table = []
            
            for row in sorted_rows:
                # Sort spans by x-position
                sorted_spans = [span[1] for span in sorted(row, key=lambda x: x[0])]
                
                # Heuristic: If the row has multiple text elements, it might be a table row
                if len(sorted_spans) >= 2:
                    current_table.append(sorted_spans)
                elif current_table:
                    # End of table
                    if len(current_table) >= 2:  # At least 2 rows to form a table
                        tables.append(current_table)
                    current_table = []
            
            # Don't forget the last table
            if current_table and len(current_table) >= 2:
                tables.append(current_table)
        
        except Exception as e:
            print(f"Warning: Error in table extraction: {e}")
        
        return tables

    def process_pdf_to_word(self, input_pdf, output_docx):
        """Process a PDF file by converting to Word, translating, and saving as docx."""
        if not os.path.exists(input_pdf):
            print(f"Error: File {input_pdf} does not exist.")
            return False

        try:
            # Create temp files
            temp_docx = f"{output_docx}_original.docx"
            
            # Step 1: Convert PDF to Word
            print(f"Converting PDF to Word: {input_pdf}")
            cv = Converter(input_pdf)
            cv.convert(temp_docx)
            cv.close()
            
            if not os.path.exists(temp_docx):
                print("PDF to Word conversion failed")
                return False
                
            # Step 2: Translate the Word document
            print(f"Translating Word document")
            self._translate_word_document(temp_docx, output_docx)
            
            # Clean up temp file
            try:
                os.remove(temp_docx)
            except:
                pass
                
            print(f"Successfully created translated Word document: {output_docx}")
            return True
            
        except Exception as e:
            print(f"Error in PDF-to-Word processing: {str(e)}")
            return False
    
    def _translate_word_document(self, input_docx, output_docx):
        """Translate text in a Word document while preserving formatting."""
        try:
            # Load the document
            doc = docx.Document(input_docx)
            
            # Process paragraphs
            for para in tqdm(doc.paragraphs, desc="Translating paragraphs"):
                if para.text.strip():
                    # Store original formatting
                    runs_formatting = []
                    for run in para.runs:
                        runs_formatting.append({
                            'bold': run.bold,
                            'italic': run.italic,
                            'underline': run.underline,
                            'font_size': run.font.size,
                            'font_name': run.font.name,
                            'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
                        })
                    
                    # Translate the whole paragraph
                    translated_text = self._translate_text(para.text)
                    
                    # Clear paragraph and add translated text
                    para.clear()
                    para.add_run(translated_text)
                    
                    # If we had multiple runs with different formatting, try to preserve by splitting the text
                    if len(runs_formatting) > 1:
                        # This is a simplified approach - for better results, a more complex algorithm would be needed
                        para.clear()
                        words = translated_text.split()
                        runs_count = len(runs_formatting)
                        words_per_run = max(1, len(words) // runs_count)
                        
                        for i, fmt in enumerate(runs_formatting):
                            start_idx = i * words_per_run
                            end_idx = (i + 1) * words_per_run if i < runs_count - 1 else len(words)
                            if start_idx < len(words):
                                run_text = ' '.join(words[start_idx:end_idx])
                                run = para.add_run(run_text + ' ')
                                run.bold = fmt['bold']
                                run.italic = fmt['italic']
                                run.underline = fmt['underline']
                                if fmt['font_size']:
                                    run.font.size = fmt['font_size']
                                if fmt['font_name']:
                                    run.font.name = fmt['font_name']
                                if fmt['color']:
                                    run.font.color.rgb = fmt['color']
            
            # Process tables
            for table in tqdm(doc.tables, desc="Translating tables"):
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                # Translate the cell text
                                translated_text = self._translate_text(para.text)
                                para.clear()
                                para.add_run(translated_text)
            
            # Save the translated document
            doc.save(output_docx)
            return True
            
        except Exception as e:
            print(f"Error translating Word document: {str(e)}")
            return False

def main():
    """Main function to parse arguments and process PDF files."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Process and translate technical PDF documents.')
    parser.add_argument('input_file', help='Path to the input PDF file')
    parser.add_argument('-o', '--output', help='Path to the output file (default: input_name_translated.pdf/xlsx/docx)')
    parser.add_argument('-e', '--excel', action='store_true', help='Extract content to Excel instead of creating a PDF')
    parser.add_argument('-w', '--word', action='store_true', help='Convert to Word document for better formatting')
    parser.add_argument('-d', '--dictionary', help='Path to a custom translation dictionary CSV file')
    
    args = parser.parse_args()
    
    # Create translator
    translator = PDFTranslator(dictionary_path=args.dictionary)
    
    if args.word:
        # Process with Word conversion
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translated.docx"
        
        success = translator.process_pdf_to_word(args.input_file, output_file)
    elif args.excel:
        # Extract to Excel
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translation.xlsx"
        
        success = translator.extract_to_excel(args.input_file, output_file)
    else:
        # Create translated PDF
        if args.output:
            output_file = args.output
        else:
            basename = os.path.basename(args.input_file)
            name_without_ext = os.path.splitext(basename)[0]
            output_file = f"{name_without_ext}_translated.pdf"
        
        success = translator.process_pdf(args.input_file, output_file)
    
    if success:
        print("Processing completed successfully.")
        sys.exit(0)
    else:
        print("Processing failed.")
        sys.exit(1)

if __name__ == "__main__":
    main() 