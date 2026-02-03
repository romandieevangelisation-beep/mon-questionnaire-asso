import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import json
import re
from supabase import create_client

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Espace Clinique - YSQ-L3 Expert", layout="wide")

# --- CONNEXION SÉCURISÉE (SUPABASE) ---
@st.cache_resource
def init_connection():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
        return None

supabase = init_connection()

# ==============================================================================
# BASE DE DONNÉES CLINIQUES & THÉOLOGIQUES (FUSION TOTALE)
# ==============================================================================
DATA_SCHEMAS = {
    "ED": {
        "titre": "Carence Affective",
        [cite_start]"slogan": "« Personne ne me considère ou ne m'aime vraiment » [cite: 128]",
        # V9 EXPERT (Votre demande de solidité)
        "clinique_expert": "Ce schéma signale un vide émotionnel précoce (alexithymie). Le patient a intégré la croyance que ses besoins de chaleur, d'empathie et de protection ne seront jamais validés.",
        "theologie_expert": "Le mensonge racine est l'orphelinat spirituel. La guérison passe par la doctrine de l'Adoption (Rom 8:15). Dieu est un Père qui s'incline pour nourrir (Osée 11:4).",
        # INFOS DU FICHIER WORD
        "origines": [
            [cite_start]"Le soignant ne répondait pas aux besoins d'affection, d'empathie ou de protection[cite: 136].",
            [cite_start]"Parents froids, absents ou ne reconnaissant pas les signaux de l'enfant [cite: 138-140].",
            "L'enfant ne s'est pas senti 'spécial' ou compris."
        ],
        "symptomes": [
            [cite_start]"Sentiment chronique de vide ou de solitude[cite: 143].",
            [cite_start]"Considérer ses propres besoins comme sans importance[cite: 144].",
            [cite_start]"Devenir dépendant/collant ou au contraire froid/distant[cite: 147].",
            [cite_start]"Ressentiment envers les autres qui 'ne donnent pas assez'[cite: 148]."
        ],
        "mecanisme_titre": "Les 3 Formes de Privation",
        "mecanisme_texte": "1. Privation d'Attention (manque de chaleur).\n2. Privation d'Empathie (manque d'écoute).\n3. [cite_start]Privation de Protection (manque de guidance) [cite: 132-134][cite_start].\nRéaction fréquente : Ne pas dire ce dont on a besoin puis être déçu[cite: 162].",
        "actions_therapeute": [
            [cite_start]"Soyez à l'écoute de vos besoins de protection et d'empathie[cite: 177].",
            [cite_start]"Exprimez vos besoins de manière assertive ('J'ai besoin de...')[cite: 181].",
            [cite_start]"Évitez les partenaires froids qui recréent la privation familière[cite: 180]."
        ],
        "action_pastorale": "Tenez un journal de vos besoins devant Dieu. Osez demander une petite chose simple à un proche sans vous excuser, comme un acte de foi que vous méritez l'amour (Psaume 27:10).",
        "verset": "Psaume 27:10"
    },
    "AB": {
        "titre": "Abandon / Instabilité",
        [cite_start]"slogan": "« Ne me quitte pas » [cite: 2]",
        "clinique_expert": "Perception de l'instabilité fondamentale des liens. Hypervigilance de la perte, alternant entre agrippement anxieux et évitement préventif.",
        "theologie_expert": "L'antidote est la théologie de l'Alliance (Berit). Contrairement aux alliances humaines, celle de Dieu est irrévocable (Hébreux 13:5).",
        "origines": [
            [cite_start]"Décès d'un parent ou départ du foyer dans l'enfance[cite: 8].",
            [cite_start]"Soignant instable (dépression, alcool) ou imprévisible[cite: 10].",
            [cite_start]"Surprotection familiale rendant la séparation angoissante[cite: 13]."
        ],
        "symptomes": [
            [cite_start]"S'accrocher aux gens par peur (agrippement)[cite: 40].",
            [cite_start]"Jalousie excessive et possessivité[cite: 18].",
            [cite_start]"Accusations injustifiées d'infidélité ou d'abandon[cite: 20]."
        ],
        "mecanisme_titre": "Le Cycle de l'Abandon",
        "mecanisme_texte": "1. Anxiété (recherche éperdue). 2. Colère/Désespoir (protestation). 3. [cite_start]Détachement (repli sur soi) [cite: 31-38]. Ce cycle de l'enfance se répète dans les relations adultes.",
        "actions_therapeute": [
            [cite_start]"Repérez votre tendance à dramatiser les séparations[cite: 48].",
            [cite_start]"Apprenez à vous apaiser seul(e) quand l'autre est absent[cite: 51].",
            [cite_start]"Évitez les partenaires instables ou ambivalents[cite: 50]."
        ],
        "action_pastorale": "Pratiquez la 'Solitude Habitée'. Passez 15 min seul(e) en visualisant la présence de Dieu. Rappelez-vous : 'Je ressens de la peur, mais je ne suis pas en danger réel'.",
        "verset": "Hébreux 13:5"
    },
    "MA": {
        "titre": "Méfiance / Abus",
        [cite_start]"slogan": "« Le monde est dangereux » [cite: 55]",
        "clinique_expert": "Attente que l'autre va nuire. Projection d'intentionnalité malveillante. Schéma de survie traumatique post-abus.",
        "theologie_expert": "Dieu est le Refuge (Mahseh). Passer de la suspicion (peur) au discernement (sagesse) sous Sa protection (Psaume 62:8).",
        "origines": [
            [cite_start]"Abus physique, sexuel ou verbal dans l'enfance[cite: 62].",
            [cite_start]"Famille humiliante, sadique ou punitive[cite: 63].",
            [cite_start]"Trahi ou manipulé par une figure de confiance[cite: 67]."
        ],
        "symptomes": [
            [cite_start]"Hypervigilance ('scanner' les menaces)[cite: 69].",
            [cite_start]"Tests de loyauté envers les autres[cite: 117].",
            [cite_start]"Attaquer avant d'être attaqué[cite: 77]."
        ],
        "mecanisme_titre": "Types d'Abus & Méfiance",
        [cite_start]"mecanisme_texte": "Le schéma naît souvent d'abus physiques, sexuels ou verbaux [cite: 81-102]. La personne reste en mode 'survie', s'attendant à ce que toute gentillesse cache un piège.",
        "actions_therapeute": [
            [cite_start]"Faites de petits pas pour faire confiance (test de réalité)[cite: 124].",
            [cite_start]"Fixez des limites claires avec les personnes toxiques[cite: 121].",
            [cite_start]"Développez de la compassion pour l'enfant blessé en vous[cite: 119]."
        ],
        "action_pastorale": "Remplacez la suspicion systématique par la prière : 'Seigneur, donne-moi ton discernement'. Déposez les armes de la défensive à la Croix.",
        "verset": "Psaume 62:8"
    },
    "SI": {
        "titre": "Isolement Social",
        [cite_start]"slogan": "« Je n'ai pas ma place ici » [cite: 251]",
        "clinique_expert": "Sentiment de différence fondamentale ('Alien'). Exclusion du groupe par manque d'appartenance ressentie.",
        "theologie_expert": "En Christ, la 'différence' est une fonction dans le Corps, pas un motif d'exclusion. Réintégration dans la famille de Dieu (Éphésiens 2:19).",
        "origines": [
            [cite_start]"Humiliation ou rejet par les pairs (école, harcèlement)[cite: 258].",
            [cite_start]"Famille différente de la communauté (religion, race, statut)[cite: 259].",
            [cite_start]"Manque de compétences sociales encouragées[cite: 261]."
        ],
        "symptomes": [
            [cite_start]"Se sentir 'imposteur' ou 'inintéressant' en groupe[cite: 267].",
            [cite_start]"Évitement systématique des activités sociales[cite: 265].",
            [cite_start]"Caméléon social pour s'intégrer (perte de soi)[cite: 269]."
        ],
        "mecanisme_titre": "Le Cycle de l'Anxiété Sociale",
        [cite_start]"mecanisme_texte": "Anxiété -> Évitement -> Manque de pratique -> Renforcement de l'inadéquation -> Isolement accru [cite: 274-279].",
        "actions_therapeute": [
            [cite_start]"Exposez-vous progressivement aux situations évitées[cite: 300].",
            [cite_start]"Trouvez votre 'tribu' (intérêts communs)[cite: 304].",
            [cite_start]"Entraînez-vous aux compétences sociales (contact visuel, questions)[cite: 302]."
        ],
        "action_pastorale": "Participez à la vie d'église non pour 'briller' mais pour 'être avec'. Vous êtes membre du Corps : l'œil ne peut dire à la main 'je n'ai pas besoin de toi'.",
        "verset": "Éphésiens 2:19"
    },
    "DS": {
        "titre": "Imperfection / Honte",
        [cite_start]"slogan": "« Je ne vaux rien » [cite: 187]",
        "clinique_expert": "Sentiment d'être intrinsèquement défectueux (Badness). Honte toxique : 'Je SUIS une erreur'.",
        "theologie_expert": "Justification par la foi. Valeur fondée sur le statut en Christ, pas l'état intérieur (Sophonie 3:17).",
        "origines": [
            [cite_start]"Famille critique, humiliante ou punitive[cite: 195].",
            [cite_start]"Rejet ou manque d'amour par un parent[cite: 196].",
            [cite_start]"Comparaison défavorable avec la fratrie[cite: 199]."
        ],
        "symptomes": [
            [cite_start]"Cacher sa vraie personnalité (masque)[cite: 221].",
            [cite_start]"Hypersensibilité à la critique[cite: 207].",
            [cite_start]"Attaquer les autres pour se revaloriser (contre-attaque)[cite: 219]."
        ],
        "mecanisme_titre": "Les 3 Copings de la Honte",
        "mecanisme_texte": "1. Capitulation (autodestruction). 2. Évitement (se cacher). 3. [cite_start]Contre-attaque (narcissisme/critique) [cite: 214-219].",
        "actions_therapeute": [
            [cite_start]"Cessez de vous comparer aux autres[cite: 241].",
            [cite_start]"Dressez une liste de vos qualités réelles[cite: 242].",
            [cite_start]"Célébrez intentionnellement vos succès[cite: 243]."
        ],
        "action_pastorale": "Quand la voix critique attaque, répondez à voix haute : 'Je suis imparfait, mais justifié, lavé et aimé en Christ'. Votre valeur a été fixée à la Croix.",
        "verset": "Sophonie 3:17"
    },
    "FA": {
        "titre": "Échec",
        [cite_start]"slogan": "« Je suis un raté » [cite: 496]",
        "clinique_expert": "Croyance en l'incompétence relative. Évitement des défis pour ne pas confirmer cette croyance.",
        "theologie_expert": "Fin de l'idolâtrie de la réussite. Le succès selon Dieu est la fidélité (2 Corinthiens 12:9).",
        "origines": [
            [cite_start]"Parents très critiques sur les résultats scolaires[cite: 502].",
            [cite_start]"Comparaison défavorable avec les autres enfants[cite: 507].",
            [cite_start]"Manque de limites ou de discipline dans l'enfance[cite: 508]."
        ],
        "symptomes": [
            [cite_start]"Procrastination par peur de l'échec[cite: 513].",
            [cite_start]"Minimiser ses propres réussites ('c'est de la chance')[cite: 519].",
            [cite_start]"Abandonner rapidement une tâche[cite: 516]."
        ],
        "mecanisme_titre": "La Pensée 'Tout ou Rien'",
        "mecanisme_texte": "Vision dichotomique : 'Si je ne suis pas parfait, je suis un échec total'. [cite_start]Cette norme irréaliste condamne à l'échec perçu [cite: 523-531].",
        "actions_therapeute": [
            [cite_start]"Reconnaissez la courbe d'apprentissage normale[cite: 550].",
            [cite_start]"Faites une liste de vos compétences réelles[cite: 552].",
            [cite_start]"Lancez un hobby sans enjeu de performance[cite: 553]."
        ],
        "action_pastorale": "Redéfinissez le succès : pour Dieu, c'est l'amour et l'obéissance. Entreprenez une action en acceptant qu'elle soit 'moyenne' aux yeux du monde, mais faite pour la gloire de Dieu.",
        "verset": "2 Corinthiens 12:9"
    },
    "DI": {
        "titre": "Dépendance / Incompétence",
        [cite_start]"slogan": "« Je n'y arrive pas tout seul » [cite: 309]",
        "clinique_expert": "Régression infantile. Croyance en l'incapacité à survivre seul. Recherche d'une figure parentale.",
        "theologie_expert": "Dieu donne un esprit de force. Dépendance verticale (Dieu) pour autonomie horizontale (hommes) (Phil 4:13).",
        "origines": [
            [cite_start]"Parents surprotecteurs ('je le fais pour toi')[cite: 342].",
            [cite_start]"Parents qui ne laissaient pas prendre de décisions[cite: 345].",
            [cite_start]"Manque de conseils pratiques (négligence)[cite: 351]."
        ],
        "symptomes": [
            [cite_start]"Besoin constant d'être rassuré[cite: 328].",
            [cite_start]"Peur paralysante de prendre une mauvaise décision[cite: 325].",
            [cite_start]"Laisser les autres diriger sa vie[cite: 326]."
        ],
        "mecanisme_titre": "Surprotection vs Négligence",
        [cite_start]"mecanisme_texte": "Soit l'enfant a été étouffé (pas d'autonomie), soit il a été livré à lui-même trop tôt sans guidance (échec appris) [cite: 341-353].",
        "actions_therapeute": [
            [cite_start]"Listez les tâches où vous dépendez des autres[cite: 367].",
            [cite_start]"Célébrez chaque acte d'autonomie, même petit[cite: 369].",
            [cite_start]"Acceptez que les erreurs ne sont pas de l'incompétence[cite: 371]."
        ],
        "action_pastorale": "Prenez une décision quotidienne seul(e) (repas, trajet) en vous confiant au Saint-Esprit qui habite en vous. Vous êtes équipé pour la vie.",
        "verset": "Philippiens 4:13"
    },
    "VU": {
        "titre": "Vulnérabilité au Danger",
        [cite_start]"slogan": "« Une catastrophe arrive » [cite: 377]",
        "clinique_expert": "Anxiété catastrophique. Monde perçu comme dangereux. Hypervigilance (Psaume 91:4).",
        "theologie_expert": "L'anxiété est une tentative d'assumer la Souveraineté de Dieu. Paix par la confiance en la Providence.",
        "origines": [
            [cite_start]"Parent anxieux ou phobique (apprentissage par observation)[cite: 386].",
            [cite_start]"Traumatisme, maladie grave ou décès d'un proche dans l'enfance[cite: 389].",
            [cite_start]"Surprotection parentale concernant les dangers[cite: 387]."
        ],
        "symptomes": [
            [cite_start]"Scénarios catastrophes (santé, argent, agression)[cite: 382].",
            [cite_start]"Vérifications compulsives (corps, portes)[cite: 394].",
            [cite_start]"Rituels superstitieux pour se protéger[cite: 400]."
        ],
        "mecanisme_titre": "Distorsions Cognitives",
        "mecanisme_texte": "1. [cite_start]Catastrophisme (le pire va arriver)[cite: 413].\n2. [cite_start]Surestimation du danger / Sous-estimation des capacités[cite: 419].\n3. [cite_start]Filtrage (ne voir que les risques)[cite: 416].",
        "actions_therapeute": [
            [cite_start]"Analysez la probabilité réelle des catastrophes[cite: 434].",
            [cite_start]"Réduisez les comportements de vérification[cite: 435].",
            [cite_start]"Exposition progressive aux situations craintes[cite: 437]."
        ],
        "action_pastorale": "Faites une 'Diète de l'info' anxiogène. Tenez un carnet de Gratitude notant 3 protections divines par jour. Ancrez-vous dans la sécurité du présent.",
        "verset": "Psaume 91:4"
    },
    "EU": {
        "titre": "Fusion / Personnalité Atrophiée",
        [cite_start]"slogan": "« Je ne peux pas vivre sans toi » [cite: 441]",
        "clinique_expert": "Symbiose émotionnelle. Manque d'individuation. Vie par procuration.",
        "theologie_expert": "Dieu crée des individus distincts. 'Quitter' pour devenir une personne entière (Galates 1:10).",
        "origines": [
            [cite_start]"Parent empêchant l'expression des besoins propres[cite: 450].",
            [cite_start]"Culpabilisation quand l'enfant s'autonomise[cite: 453].",
            [cite_start]"Parent vivant à travers l'enfant[cite: 483]."
        ],
        "symptomes": [
            [cite_start]"Sentiment de vide quand on est seul[cite: 458].",
            [cite_start]"Imiter les émotions/avis de l'autre[cite: 463].",
            [cite_start]"Culpabilité intense à avoir une vie privée[cite: 461]."
        ],
        "mecanisme_titre": "Identité Non-Développée",
        "mecanisme_texte": "La personne ne sait pas qui elle est sans l'autre. Elle se définit par 'nous' plutôt que 'je'. [cite_start]Risque de relations toxiques[cite: 446].",
        "actions_therapeute": [
            [cite_start]"Listez vos différences (goûts, avis) avec l'autre[cite: 487].",
            [cite_start]"Passez du temps seul pour découvrir qui vous êtes[cite: 488].",
            [cite_start]"Fixez des limites (ex: ne pas répondre immédiatement)[cite: 490]."
        ],
        "action_pastorale": "Osez exprimer une opinion différente d'un proche sur un sujet mineur. C'est un acte spirituel d'affirmation de la créature unique que Dieu a faite en vous.",
        "verset": "Galates 1:10"
    },
    "SB": {
        "titre": "Assujettissement",
        [cite_start]"slogan": "« Je dois faire ce que tu veux » [cite: 671]",
        "clinique_expert": "Soumission forcée par peur. Répression des besoins et colère latente (agressivité passive).",
        "theologie_expert": "Serviteur de Dieu affranchi des hommes. La vraie soumission est un choix d'amour (Galates 5:1).",
        "origines": [
            [cite_start]"Parent dominant, contrôlant ou punitif[cite: 681].",
            [cite_start]"Menaces, colère ou retrait d'amour si désaccord[cite: 682].",
            [cite_start]"Rôle de parentification (s'occuper du parent)[cite: 685]."
        ],
        "symptomes": [
            [cite_start]"Peur de dire non[cite: 717].",
            [cite_start]"Sentiment d'être piégé[cite: 690].",
            [cite_start]"Accumulation de colère (ressentiment)[cite: 696]."
        ],
        "mecanisme_titre": "Le Rôle de la Colère",
        "mecanisme_texte": "La soumission crée une dette émotionnelle. [cite_start]La colère refoulée finit par exploser ou devenir des symptômes psychosomatiques [cite: 699-709].",
        "actions_therapeute": [
            [cite_start]"Entraînez-vous à dire 'non'[cite: 719].",
            [cite_start]"Identifiez vos droits et besoins légitimes[cite: 717].",
            [cite_start]"Tolérer l'inconfort de ne pas plaire[cite: 720]."
        ],
        "action_pastorale": "Exercez-vous au 'Non bienveillant'. Refusez une demande cette semaine. Rappelez-vous que vous servez Dieu, pas l'humeur changeante des autres.",
        "verset": "Galates 5:1"
    },
    "SS": {
        "titre": "Abnégation",
        [cite_start]"slogan": "« Je suis le sauveur » [cite: 726]",
        "clinique_expert": "Syndrome du Sauveur. Focalisation sur autrui par culpabilité ou besoin de valorisation.",
        "theologie_expert": "Nous ne sommes pas le Messie. L'intendance de soi est un devoir biblique (Matthieu 22:39).",
        "origines": [
            [cite_start]"Responsabilité excessive d'un proche dans l'enfance[cite: 735].",
            [cite_start]"Valorisée uniquement quand elle donnait ('sois gentil')[cite: 736].",
            [cite_start]"Tempérament naturellement empathique[cite: 738]."
        ],
        "symptomes": [
            [cite_start]"Ne pas savoir recevoir de l'aide[cite: 745].",
            [cite_start]"Épuisement (burnout) et rancœur cachée[cite: 746].",
            [cite_start]"Attiré par les personnes à problèmes[cite: 748]."
        ],
        "mecanisme_titre": "Frontières (Boundaries)",
        "mecanisme_texte": "Difficulté à fixer des limites. Le sacrifice est motivé par la culpabilité, pas par l'amour libre. [cite_start]C'est une forme de codépendance [cite: 761-764].",
        "actions_therapeute": [
            "Équilibrez le donner et le recevoir.",
            "Demandez-vous : 'Je le fais par envie ou par culpabilité ?'.",
            [cite_start]"Redéfinissez l'égoïsme[cite: 780]."
        ],
        "action_pastorale": "Pratiquez le Sabbat : une demi-journée sans 'servir', juste pour être aimé de Dieu sans rien faire. C'est un acte d'humilité : le monde tourne sans vous.",
        "verset": "Matthieu 22:39"
    },
    "EI": {
        "titre": "Inhibition Émotionnelle",
        [cite_start]"slogan": "« Je ne dois pas ressentir » [cite: 896]",
        "clinique_expert": "Sur-contrôle. 'Faux-self' rationnel pour se protéger. Peur de perdre le contrôle.",
        "theologie_expert": "Jésus a pleuré. Les émotions sont des signaux. La vérité implique l'authenticité (Psaume 62:9).",
        "origines": [
            [cite_start]"Émotions moquées, punies ou rejetées dans l'enfance[cite: 907].",
            [cite_start]"Famille puritaine ou très rationnelle ('pleurer c'est faible')[cite: 909].",
            [cite_start]"Peur de ressembler à un parent hystérique[cite: 911]."
        ],
        "symptomes": [
            [cite_start]"Paraître froid, robotique ou distant[cite: 915].",
            [cite_start]"Incapacité à pleurer ou montrer sa joie[cite: 916].",
            [cite_start]"Accent excessif sur la logique[cite: 904]."
        ],
        "mecanisme_titre": "La Roue des Émotions",
        "mecanisme_texte": "Inhibition de la colère, de la joie ou de la vulnérabilité. [cite_start]Rationalisation pour éviter de ressentir la douleur [cite: 924-927].",
        "actions_therapeute": [
            [cite_start]"Utilisez la 'Roue des émotions' pour nommer ce que vous ressentez[cite: 924].",
            [cite_start]"Tenez un journal émotionnel[cite: 940].",
            [cite_start]"Recherchez des expériences émotionnelles (films, musique)[cite: 942]."
        ],
        "action_pastorale": "Priez avec les Psaumes de lamentation. Osez dire 'Je suis triste' ou 'Je suis en colère' à Dieu. Les émotions ne sont pas des péchés, ce sont des informations.",
        "verset": "Psaume 62:9"
    },
    "US": {
        "titre": "Exigences Élevées",
        [cite_start]"slogan": "« Ce n'est jamais assez bien » [cite: 949]",
        "clinique_expert": "Perfectionnisme pathologique. Tyrannie du 'Je dois'. Incapacité à ressentir la satisfaction.",
        "theologie_expert": "Légalisme de la performance. L'Évangile est la fin de la performance pour le salut (Matthieu 11:28).",
        "origines": [
            [cite_start]"Amour parental conditionnel à la réussite[cite: 960].",
            [cite_start]"Parents eux-mêmes perfectionnistes (modèles)[cite: 961].",
            [cite_start]"Critique ou honte en cas d'échec[cite: 963]."
        ],
        "symptomes": [
            [cite_start]"Impossible de se détendre[cite: 965].",
            [cite_start]"Hyper-critique envers soi et les autres[cite: 952].",
            [cite_start]"Symptômes physiques de stress (insomnie, etc.)[cite: 970]."
        ],
        "mecanisme_titre": "Les 3 Types de Normes",
        "mecanisme_texte": "1. Compulsivité (ordre). 2. Orientation réussite (travail). 3. Orientation statut (image). [cite_start]C'est une course sans fin [cite: 973-984].",
        "actions_therapeute": [
            [cite_start]"Essayez de réduire vos exigences de 10%[cite: 999].",
            [cite_start]"Listez les avantages et inconvénients de votre pression[cite: 997].",
            [cite_start]"Forcez-vous à ralentir et faire des pauses[cite: 1002]."
        ],
        "action_pastorale": "Le défi de l'imperfection : laissez volontairement une tâche inachevée (ex: lit mal fait) et observez que Dieu vous aime toujours autant. La grâce suffit.",
        "verset": "Matthieu 11:28"
    },
    "ET": {
        "titre": "Droits Personnels / Grandeur",
        [cite_start]"slogan": "« Les règles ne s'appliquent pas à moi » [cite: 558]",
        "clinique_expert": "Narcissisme et privilège. Manque d'empathie. Compensation d'une infériorité cachée.",
        "theologie_expert": "Le Royaume inversé : le grand est serviteur. Reconnaître sa dépendance brise l'orgueil (Phil 2:3).",
        "origines": [
            [cite_start]"Enfant gâté, sans limites[cite: 565].",
            [cite_start]"Parents n'ayant pas imposé de conséquences[cite: 566].",
            [cite_start]"Compensation d'un sentiment de manque affectif[cite: 569]."
        ],
        "symptomes": [
            [cite_start]"Colère si on ne l'obéit pas immédiatement[cite: 585].",
            [cite_start]"Manque d'empathie pour les besoins d'autrui[cite: 586].",
            [cite_start]"Compétitivité excessive et manipulation[cite: 573]."
        ],
        "mecanisme_titre": "Les 3 Types de Droits",
        "mecanisme_texte": "1. Narcissisme pur (je suis spécial). 2. Dépendance (les autres doivent me servir). 3. [cite_start]Impulsivité (je veux tout, tout de suite) [cite: 581-594].",
        "actions_therapeute": [
            [cite_start]"Mettez-vous à la place des autres (empathie cognitive)[cite: 609].",
            [cite_start]"Demandez un feedback honnête à un ami[cite: 608].",
            [cite_start]"Travaillez la maîtrise des impulsions[cite: 610]."
        ],
        "action_pastorale": "Pratiquez le service anonyme. Faites une bonne action (vaisselle, don) sans le dire et sans attendre de merci. Écoutez les autres sans ramener la conversation à vous.",
        "verset": "Philippiens 2:3"
    },
    "IS": {
        "titre": "Contrôle de soi insuffisant",
        [cite_start]"slogan": "« C'est trop difficile, je m'en fiche » [cite: 617]",
        "clinique_expert": "Impulsivité, principe de plaisir dominant. Difficulté à différer la gratification.",
        "theologie_expert": "La maîtrise de soi est un fruit de l'Esprit. Dire non à la chair pour dire oui à la vie (Prov 25:28).",
        "origines": [
            [cite_start]"Manque de discipline parentale[cite: 625].",
            [cite_start]"Négligence ou stress chronique affectant le cerveau[cite: 624].",
            [cite_start]"Enfant jamais forcé à tolérer la frustration[cite: 626]."
        ],
        "symptomes": [
            [cite_start]"Procrastination chronique[cite: 633].",
            [cite_start]"Addictions ou évitement de l'inconfort[cite: 629].",
            [cite_start]"Difficulté à tenir des engagements[cite: 630]."
        ],
        "mecanisme_titre": "La Stratégie SNAP",
        "mecanisme_texte": "Stop, Notice (Remarquer), Align (Aligner avec valeurs), Process (Agir). [cite_start]Outil pour briser l'impulsion [cite: 638-653].",
        "actions_therapeute": [
            "Utilisez la méthode SNAP.",
            [cite_start]"Fixez des micro-objectifs réalisables[cite: 664].",
            [cite_start]"Enlevez les distractions de l'environnement[cite: 666]."
        ],
        "action_pastorale": "La méthode des 10 minutes : Quand vous voulez abandonner une tâche ou céder à une impulsion, tenez 10 min de plus en priant. C'est un muscle spirituel à exercer.",
        "verset": "Proverbes 25:28"
    },
    "AS": {
        "titre": "Recherche d'approbation",
        [cite_start]"slogan": "« Ma valeur dépend de ton regard » [cite: 785]",
        "clinique_expert": "Estime de soi externalisée. Caméléon social. Perte d'authenticité.",
        "theologie_expert": "Idolâtrie de l'approbation humaine. Seule l'approbation du Père libère (1 Thess 2:4).",
        "origines": [
            [cite_start]"Amour conditionné à la 'bonne conduite' sociale[cite: 792].",
            [cite_start]"Parents soucieux des apparences et du statut[cite: 793].",
            [cite_start]"Manque d'attention comblé par la performance[cite: 794]."
        ],
        "symptomes": [
            [cite_start]"Changer de personnalité selon l'interlocuteur[cite: 800].",
            [cite_start]"Importance excessive du statut/richesse[cite: 798].",
            [cite_start]"Peur panique de déplaire[cite: 801]."
        ],
        "mecanisme_titre": "L'Adaptation Excessive",
        "mecanisme_texte": "Le patient pense : 'Si je suis moi-même, on ne m'aimera pas'. [cite_start]Il développe un 'Faux-Self' pour être validé [cite: 807-811].",
        "actions_therapeute": [
            [cite_start]"Demandez-vous : 'Qu'est-ce que JE veux ?' avant d'agir[cite: 827].",
            [cite_start]"Entraînez-vous à exprimer vos préférences[cite: 829].",
            [cite_start]"Passez du temps seul pour vous retrouver[cite: 830]."
        ],
        "action_pastorale": "Faites le bien en secret (Matthieu 6). Acceptez un compliment par un simple 'Merci' sans vous en nourrir excessivement ni le rejeter. Votre audience est Dieu seul.",
        "verset": "1 Thessaloniciens 2:4"
    },
    "NP": {
        "titre": "Négativité / Pessimisme",
        [cite_start]"slogan": "« Ça va mal finir » [cite: 837]",
        "clinique_expert": "Biais cognitif négatif. Attente anxieuse que tout s'effondre. Positif minimisé.",
        "theologie_expert": "La joie est un combat de la foi. Discipline de l'attention (Phil 4:8).",
        "origines": [
            [cite_start]"Parents pessimistes ou inquiets (modèle)[cite: 845].",
            [cite_start]"Enfance marquée par des difficultés réelles/instabilité[cite: 847].",
            [cite_start]"Découragement de l'autonomie ou de la joie[cite: 846]."
        ],
        "symptomes": [
            [cite_start]"Filtre négatif (ne voit que le problème)[cite: 864].",
            [cite_start]"Incapacité à se réjouir quand tout va bien[cite: 855].",
            [cite_start]"Plaintes chroniques et indécision[cite: 842]."
        ],
        "mecanisme_titre": "Les Distorsions Cognitives",
        "mecanisme_texte": "1. Filtre négatif. 2. Généralisation excessive ('toujours'). 3. Catastrophisme. [cite_start]C'est une protection : 'Si je m'attends au pire, je ne serai pas déçu' [cite: 860-878].",
        "actions_therapeute": [
            [cite_start]"Examinez les preuves : est-ce vraiment toujours négatif ?[cite: 890].",
            [cite_start]"Tenez un journal de gratitude (3 choses/jour)[cite: 889].",
            [cite_start]"Considérez les exceptions à vos prévisions[cite: 892]."
        ],
        "action_pastorale": "Contre la rumination, trouvez un aspect positif pour chaque pensée négative. Louez Dieu pour une petite chose précise chaque matin.",
        "verset": "Lamentations 3:21"
    },
    "PU": {
        "titre": "Punition",
        [cite_start]"slogan": "« Les erreurs doivent être punies » [cite: 1007]",
        "clinique_expert": "Intransigeance. Erreur = châtiment nécessaire. Difficulté à pardonner.",
        "theologie_expert": "Incompréhension de la Croix. Christ a pris la punition (Rom 8:1).",
        "origines": [
            [cite_start]"Punitions sévères, critiques ou humiliations dans l'enfance[cite: 1014].",
            [cite_start]"Parents impitoyables ou moralisateurs[cite: 1015].",
            "Manque de droit à l'erreur."
        ],
        "symptomes": [
            [cite_start]"Rancune tenace[cite: 1022].",
            [cite_start]"Autopunition ou automutilation[cite: 1019].",
            [cite_start]"Jugement sévère des autres et de soi[cite: 1017]."
        ],
        "mecanisme_titre": "Le Cycle de la Rancune",
        "mecanisme_texte": "Standards rigides -> Erreur inévitable -> Colère/Jugement -> Punition. [cite_start]Croyance que la punition 'corrige' le comportement[cite: 1028].",
        "actions_therapeute": [
            [cite_start]"Pratiquez l'auto-compassion[cite: 1052].",
            [cite_start]"Considérez les circonstances atténuantes[cite: 1053].",
            "Pardonnez-vous une erreur passée."
        ],
        "action_pastorale": "Si Jésus a payé, ne cherchez pas à payer encore. Parlez-vous avec la douceur que le Christ utilise pour vous parler.",
        "verset": "Romains 8:1"
    }
}
# --- STRUCTURE DES DOMAINES DE YOUNG ---
YOUNG_DOMAINS_INFO = {
    "Domaine I : Séparation et Rejet": {
        "codes": ["ED", "AB", "MA", "SI", "DS"],
        "besoin": "Besoin de sécurité, de stabilité, d'affection et d'appartenance."
    },
    "Domaine II : Manque d'Autonomie et de Performance": {
        "codes": ["DI", "VU", "EU", "FA"],
        "besoin": "Besoin de compétence, d'identité propre et de confiance en soi."
    },
    "Domaine III : Limites Déficientes": {
        "codes": ["ET", "IS"],
        "besoin": "Besoin de limites réalistes, de respect des autres et d'autodiscipline."
    },
    "Domaine IV : Orientation vers les Autres": {
        "codes": ["SB", "SS", "AS"],
        "besoin": "Besoin de liberté d'expression et d'affirmation de ses besoins."
    },
    "Domaine V : Hypervigilance et Inhibition": {
        "codes": ["NP", "EI", "US", "PU"],
        "besoin": "Besoin de spontanéité, de plaisir et de lâcher-prise."
    }
}

# --- LES 232 QUESTIONS VALIDÉES (OFFICIELLES) ---
YSQ_QUESTIONS = {
    "ED : Carence affective": {
        1: "Je n'ai pas eu quelqu'un pour prendre soin de moi, partager sa vie avec moi, ou se soucier réellement de tout ce qui m'arrivait.",
        2: "Je n'ai pas reçu suffisamment d'affection, de chaleur ou d'amour.",
        3: "Pour l'essentiel, je n'ai eu personne sur qui compter pour recevoir des conseils et un soutien affectif.",
        4: "La plupart du temps, je n'ai eu personne pour me nourrir, m'épauler ou se soucier réellement de moi.",
        5: "J'ai manqué de quelqu'un qui puisse me rassurer physiquement, me serrer dans ses bras ou m'apporter de la tendresse.",
        6: "Pour l'essentiel, les gens n'ont pas été là pour moi, pour me donner de la chaleur et de l'affection.",
        7: "J'ai eu le sentiment que je n'avais personne vers qui me tourner pour recevoir des conseils ou une orientation.",
        8: "Je n'ai pas eu quelqu'un qui m'écoute vraiment, me comprenne ou soit sensible à mes vrais besoins.",
        9: "J'ai rarement eu quelqu'un de fort pour me donner des conseils avisés quand j'étais désemparé(e)."
    },
    "AB : Abandon / Instabilité": {
        10: "Je m'inquiète beaucoup à l'idée que les gens que j'aime vont mourir ou me quitter.",
        11: "Je m'accroche aux gens parce que j'ai peur qu'ils me quittent.",
        12: "Je crains que les gens que j'aime ne trouvent quelqu'un d'autre qu'ils préféreront et ne m'abandonnent.",
        13: "Les gens qui m'ont été proches ont été imprévisibles : un moment disponibles, le moment d'après fâchés ou absents.",
        14: "J'ai tellement besoin des autres que je m'inquiète de les perdre.",
        15: "Je me sens désespéré(e) quand quelqu'un que j'aime s'éloigne de moi, même brièvement.",
        16: "Je tombe souvent amoureux(se) de gens qui ne peuvent pas s'engager avec moi de façon stable.",
        17: "La plupart des gens sont changeants concernant leurs sentiments envers moi.",
        18: "En fin de compte, je serai seul(e).",
        19: "Quand je sens que quelqu'un à qui je tiens s'éloigne de moi, je deviens désespéré(e).",
        20: "Parfois, j'ai tellement peur que les gens me quittent que je les fais fuir.",
        21: "Je ne peux pas compter sur les gens pour être là de façon permanente.",
        22: "Je ne peux pas me laisser aller à être moi-même ou les gens me quitteront.",
        23: "Je suis obsédé(e) par l'idée que mes relations vont se terminer.",
        24: "Je n'ai pas de base affective stable.",
        25: "Je ne peux pas vivre sans quelqu'un qui m'aime.",
        26: "J'ai besoin que les autres me rassurent constamment sur le fait qu'ils ne vont pas me quitter."
    },
    "MA : Méfiance / Abus": {
        27: "J'ai l'impression que les autres vont profiter de moi.",
        28: "Je sens que je dois me protéger des autres.",
        29: "Je pense que si je laisse les gens m'approcher, ils me feront du mal.",
        30: "Si quelqu'un est gentil, je me demande ce qu'il veut vraiment.",
        31: "Je teste les gens pour voir s'ils sont honnêtes et bien intentionnés.",
        32: "Je suis très méfiant(e) vis-à-vis des motifs des autres.",
        33: "Je pense que les gens pensent d'abord à eux-mêmes.",
        34: "J'ai été maltraité(e), abusé(e) ou négligé(e) par des gens importants pour moi.",
        35: "Je me sens souvent trahi(e) par les autres.",
        36: "Je suis sur mes gardes la plupart du temps.",
        37: "Il est très difficile pour moi de faire confiance.",
        38: "Je pense que les gens me feront du mal si j'en laisse l'occasion.",
        39: "Je crains d'être attaqué(e) physiquement ou verbalement par les autres.",
        40: "J'ai l'impression que les gens se moquent de moi ou m'utilisent.",
        41: "Je pense que les gens m'utiliseront à leurs propres fins si je ne me protège pas.",
        42: "Je suis souvent sur la défensive avec les gens.",
        43: "Le monde est globalement un endroit dangereux."
    },
    "SI : Isolement social": {
        44: "Je ne me sens pas à ma place dans les groupes.",
        45: "Je me sens différent(e) des autres.",
        46: "Je me sens isolé(e) du reste du monde.",
        47: "Je n'appartiens à aucun groupe ou communauté spécifique.",
        48: "Je me sens seul(e) même quand je suis avec d'autres gens.",
        49: "Je me sens étranger(ère) partout où je vais.",
        50: "Personne ne me comprend vraiment en profondeur.",
        51: "Je suis en marge de la société.",
        52: "Je me sens ennuyeux(se) ou inintéressant(e) socialement.",
        53: "Je ne sais pas quoi dire dans les situations sociales."
    },
    "DS : Imperfection / Honte": {
        54: "Si les gens me connaissaient vraiment, ils ne m'aimeraient pas.",
        55: "J'ai des secrets personnels que je ne veux pas que les autres découvrent.",
        56: "Je suis fondamentalement défectueux(se) ou imparfait(e).",
        57: "Je ne mérite pas d'être aimé(e).",
        58: "J'ai honte de certains aspects de ma personnalité.",
        59: "Je suis une mauvaise personne au fond de moi.",
        60: "Je cache mes défauts réels aux autres.",
        61: "Je suis indigne de respect.",
        62: "Je me sens humilié(e) par mes échecs ou mes manques.",
        63: "Je suis extrêmement critique envers moi-même.",
        64: "Je me sens coupable d'être qui je suis.",
        65: "Je ne suis pas à la hauteur des autres.",
        66: "Je me dévalorise dès que je fais une erreur.",
        67: "Je crains que mes défauts ne soient exposés au grand jour.",
        68: "Je suis gêné(e) par moi-même."
    },
    "FA : Échec": {
        69: "Je suis moins compétent(e) que les autres dans mon travail ou mes études.",
        70: "J'ai échoué dans presque tout ce que j'ai entrepris.",
        71: "Je ne suis pas aussi intelligent(e) que la plupart de mes pairs.",
        72: "Je n'ai pas de talent particulier qui me distingue.",
        73: "Je ne réussirai jamais rien d'important dans ma vie.",
        74: "Je me sens bête comparé(e) aux autres.",
        75: "Je suis un(e) raté(e).",
        76: "Je ne suis pas capable de travailler aussi efficacement que les autres.",
        77: "Je me sens inférieur(e) professionnellement."
    },
    "DI : Dépendance / Incompétence": {
        78: "Je ne me sens pas capable de me débrouiller seul(e) dans la vie courante.",
        79: "J'ai impérativement besoin de l'aide des autres pour prendre des décisions.",
        80: "Je ne sais pas gérer mes finances ou mes responsabilités sans aide.",
        81: "Je me sens comme un enfant face aux responsabilités d'adulte.",
        82: "J'ai peur de faire des erreurs graves si je n'ai pas de conseils constants.",
        83: "Je ne peux pas survivre sans quelqu'un pour s'occuper de moi.",
        84: "Je ne fais pas confiance à mon propre jugement pour les choses quotidiennes.",
        85: "Je me sens dépassé(e) par les défis de la vie.",
        86: "Je cherche toujours quelqu'un pour me dire quoi faire.",
        87: "Je me sens vulnérable dès que je suis seul(e).",
        88: "Je ne sais pas résoudre les problèmes pratiques courants.",
        89: "Je panique quand je dois affronter un défi inconnu seul(e).",
        90: "Je laisse volontiers les autres prendre les commandes.",
        91: "Je ne suis pas du tout autonome.",
        92: "Je me sens incompétent(e) dans la plupart des domaines techniques."
    },
    "VU : Vulnérabilité": {
        93: "Je ne peux pas m'empêcher de penser qu'une catastrophe va arriver.",
        94: "J'ai peur de tomber subitement malade ou d'avoir une crise cardiaque.",
        95: "J'ai peur d'être agressé(e), volé(e) ou attaqué(e).",
        96: "Je crains de perdre tout mon argent et de finir dans la rue.",
        97: "Je suis obsédé(e) par les mesures de sécurité.",
        98: "Je surveille mon corps excessivement au moindre symptôme.",
        99: "Je crains de devenir fou/folle ou de perdre totalement le contrôle.",
        100: "Je panique facilement face à des risques potentiels.",
        101: "Le monde extérieur est plein de menaces imprévisibles.",
        102: "Je ne me sens jamais totalement en sécurité.",
        103: "Je crains les accidents de transport (avion, voiture).",
        104: "Je suis envahi(e) par l'anxiété la plupart du temps."
    },
    "EU : Fusion / Personnalité atrophiée": {
        105: "Je ne sais pas qui je suis réellement sans les autres.",
        106: "Je suis trop impliqué(e) dans la vie intime de mes proches.",
        107: "Je me sens coupable d'avoir une vie privée cachée à mes proches.",
        108: "Je ne pourrais pas survivre psychologiquement sans mon partenaire ou mes parents.",
        109: "Je n'ai pas d'identité propre séparée de ma famille.",
        110: "Je ressens la douleur des autres comme si c'était la mienne.",
        111: "Je me sens vide ou perdu(e) quand je suis seul(e).",
        112: "Je fusionne émotionnellement avec les gens que j'aime.",
        113: "Je ne sais pas ce que je veux vraiment pour moi-même.",
        114: "Je vis ma vie par procuration à travers les autres.",
        115: "Mes limites personnelles sont floues avec mon entourage."
    },
    "SB : Assujettissement": {
        116: "Je laisse les autres prendre les décisions importantes pour moi.",
        117: "Je n'ose pas dire 'non' aux demandes, même abusives.",
        118: "Je crains les représailles ou le rejet si je ne suis pas d'accord.",
        119: "Je sacrifie systématiquement mes besoins pour éviter le conflit.",
        120: "Je me sens coupable dès que je pense à mes propres intérêts.",
        121: "Je me laisse facilement dominer par des personnalités fortes.",
        122: "Je n'exprime presque jamais mes vrais désirs ou opinions.",
        123: "Je fais tout pour plaire, même au détriment de mes valeurs.",
        124: "Je me sens piégé(e) dans mes relations car je n'ose pas m'affirmer.",
        125: "Je refoule ma colère pour ne pas créer de tensions."
    },
    "SS : Abnégation": {
        126: "Je m'occupe des besoins des autres bien avant les miens.",
        127: "Je suis toujours celui/celle qui écoute et aide les autres.",
        128: "Je donne énormément sans jamais oser demander en retour.",
        129: "Je me sens responsable du bien-être et du bonheur de tout mon entourage.",
        130: "La souffrance des autres m'est insupportable, je dois agir.",
        131: "Je suis le 'sauveteur' attitré de ma famille ou de mes amis.",
        132: "Je m'épuise physiquement pour rendre service.",
        133: "Je néglige ma propre santé pour m'occuper d'autrui.",
        134: "Demander de l'aide me donne l'impression d'être un fardeau.",
        135: "Je suis trop gentil(le) avec les gens, même ceux qui ne le méritent pas.",
        136: "Prendre du temps pour moi me fait me sentir égoïste.",
        137: "Je veux porter les problèmes du monde sur mes épaules.",
        138: "Mes envies passent toujours après celles des autres.",
        139: "Je suis hyper-sensible à la détresse d'autrui.",
        140: "Je ne sais pas recevoir de l'affection ou des cadeaux sans malaise.",
        141: "Tout le monde compte sur moi pour tenir le coup.",
        142: "Je n'existe qu'à travers l'utilité que j'ai pour les autres."
    },
    "EI : Inhibition émotionnelle": {
        143: "Je cache soigneusement mes sentiments profonds.",
        144: "Je ne montre jamais ma colère, même quand elle est justifiée.",
        145: "Les démonstrations d'affection en public me mettent mal à l'aise.",
        146: "Je contrôle rigoureusement tout ce que je ressens.",
        147: "On me dit souvent que je parais froid(e) ou distant(e).",
        148: "Je ne pleure jamais devant les autres.",
        149: "Je privilégie la logique pure au détriment des émotions.",
        150: "J'ai peur de perdre toute crédibilité si j'exprime ma vulnérabilité.",
        151: "Je garde tout mon stress et ma peine à l'intérieur."
    },
    "US : Exigences élevées": {
        152: "Je dois être parfait(e) dans tout ce que je fais.",
        153: "Je ne suis jamais satisfait(e) de mes résultats, je peux faire mieux.",
        154: "Je travaille sans relâche, le repos est une perte de temps.",
        155: "Je suis obsédé(e) par les détails et la perfection.",
        156: "Une seule petite erreur gâche tout mon travail.",
        157: "Je me mets une pression insoutenable pour réussir.",
        158: "Je suis très exigeant(e) envers les autres aussi.",
        159: "Tout doit être parfaitement rangé et organisé.",
        160: "Je ne sais pas m'amuser ou lâcher prise.",
        161: "J'ai toujours l'impression d'être en retard.",
        162: "La médiocrité est inacceptable pour moi.",
        163: "Je sacrifie mon plaisir personnel pour atteindre mes buts.",
        164: "Ce que j'accomplis n'est jamais assez à mes yeux.",
        165: "La compétition est mon moteur permanent.",
        166: "L'échec est une source de honte profonde.",
        167: "Mes standards sont beaucoup plus élevés que ceux des autres."
    },
    "ET : Droits personnels / Grandeur": {
        168: "Je suis quelqu'un de spécial et je mérite des privilèges.",
        169: "Je ne supporte pas d'attendre ou qu'on me dise 'non'.",
        170: "Je mérite un traitement de faveur par rapport aux autres.",
        171: "Les règles ordinaires ne s'appliquent pas vraiment à moi.",
        172: "Mes besoins passent avant ceux de n'importe qui d'autre.",
        173: "Je me sens supérieur(e) à la moyenne des gens.",
        174: "Je m'énerve violemment si je ne contrôle pas la situation.",
        175: "J'utilise les autres pour obtenir ce que je veux.",
        176: "Je n'ai pas à me justifier de mes actes.",
        177: "Je suis destiné(e) à une réussite exceptionnelle.",
        178: "Les autres devraient m'admirer pour ce que je suis."
    },
    "IS : Contrôle de soi insuffisant": {
        179: "Je n'ai aucune autodiscipline.",
        180: "Je suis incapable de finir une tâche ennuyeuse.",
        181: "Je cède instantanément à mes impulsions.",
        182: "Je ne supporte pas la frustration ou l'attente.",
        183: "Je m'ennuie très vite et je change tout le temps d'avis.",
        184: "Je vis au-dessus de mes moyens sans réfléchir.",
        185: "J'agis avant de penser aux conséquences.",
        186: "Je fuis les responsabilités trop pesantes.",
        187: "Ma vie est chaotique et désorganisée.",
        188: "Je perds mon sang-froid pour des broutilles.",
        189: "Je remets tout au lendemain par paresse.",
        190: "J'ai des addictions ou des excès que je ne contrôle pas.",
        191: "Je suis incapable de suivre un plan à long terme.",
        192: "Mes émotions dictent ma conduite du moment.",
        193: "Je cherche le plaisir immédiat sans regarder l'avenir."
    },
    "AS : Recherche d'approbation": {
        194: "L'opinion des autres définit mon estime de moi.",
        195: "Je change de personnalité selon la personne en face de moi.",
        196: "J'ai besoin d'être constamment complimenté(e).",
        197: "Je fais tout pour être le centre de l'attention.",
        198: "Une critique peut me briser pour plusieurs jours.",
        199: "Je soigne mon image de façon excessive.",
        200: "Je veux à tout prix être célèbre ou influent(e).",
        201: "Je suis incapable de choisir sans demander l'avis d'autrui.",
        202: "Je veux plaire à tout le monde, même aux gens que je n'aime pas.",
        203: "Je me sens vide si personne ne me remarque.",
        204: "Je cherche désespérément des signes de reconnaissance.",
        205: "Ma valeur dépend de mon statut social.",
        206: "Je flatte les gens pour être bien accepté(e).",
        207: "Le prestige est ma motivation principale."
    },
    "NP : Négativité / Pessimisme": {
        208: "Je m'attends toujours au pire scénario possible.",
        209: "La vie est injuste et cruelle.",
        210: "Rien ne sert de se réjouir, le malheur arrive toujours.",
        211: "Je rumine mes échecs passés sans cesse.",
        212: "Je vois les défauts avant les qualités.",
        213: "Je suis convaincu(e) que la chance n'existe pas pour moi.",
        214: "Le monde va de mal en pis.",
        215: "Je suis très cynique vis-à-vis des intentions humaines.",
        216: "Je m'inquiète pour des choses qui n'arriveront probablement jamais.",
        217: "L'avenir me semble sombre et angoissant.",
        218: "Je me plains souvent de mes difficultés.",
        219: "Je suis incapable de voir le bon côté des choses.",
        220: "Le bonheur est une illusion éphémère.",
        221: "Je décourage les autres avec mon réalisme négatif.",
        222: "Je me focalise sur ce qui manque plutôt que sur ce que j'ai."
    },
    "PU : Punition": {
        223: "On doit payer sévèrement pour ses fautes.",
        224: "Je ne me pardonne aucune erreur.",
        225: "Je garde rancune très longtemps.",
        226: "La vengeance est une forme de justice nécessaire.",
        227: "Je suis impitoyable avec les gens qui se trompent.",
        228: "Je mérite d'être puni(e) quand je n'atteins pas mes buts.",
        229: "La faiblesse humaine ne mérite pas d'indulgence.",
        230: "Je m'insulte intérieurement dès que je fais une gaffe.",
        231: "La justice doit être dure pour être efficace.",
        232: "Je ne crois pas à la deuxième chance."
    }
}

# --- FONCTIONS SUPABASE ---
def save_patient_data(nom, email, reponses_dict):
    if not supabase: return False
    data = {"nom": nom, "email": email, "reponses_json": json.dumps(reponses_dict), "created_at": datetime.now().isoformat()}
    try:
        supabase.table("patients_ysq").insert(data).execute()
        return True
    except: return False

def load_all_patients():
    if not supabase: return pd.DataFrame()
    try:
        response = supabase.table("patients_ysq").select("*").order("created_at", desc=True).execute()
        return pd.DataFrame(response.data)
    except: return pd.DataFrame()

def delete_patient(patient_id):
    if not supabase: return False
    try:
        supabase.table("patients_ysq").delete().eq("id", patient_id).execute()
        return True
    except: return False

# --- INTERFACE ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3050/3050525.png", width=80)
st.sidebar.title("Navigation")
mode = st.sidebar.radio("Aller vers :", ["Espace Patient", "Espace Thérapeute"])

# ==============================================================================
# 1. ESPACE PATIENT (PUBLIC)
# ==============================================================================
if mode == "Espace Patient":
    st.header("Bienvenue dans votre questionnaire des Schémas (YSQ-L3)")
    st.markdown("---")
    
    st.info("""
    ### 💡 Guide pour remplir ce questionnaire
    
    Ce questionnaire est un outil précieux pour mieux comprendre votre fonctionnement émotionnel. Il ne s'agit pas d'un examen, mais d'une "photographie" de vos ressentis.
    
    **Comment répondre ?**
    1.  **Soyez spontané(e) :** Ne réfléchissez pas trop longtemps. Votre première impression est souvent la plus juste.
    2.  **Visez la globalité :** Répondez en fonction de ce que vous ressentez **la plupart du temps** dans votre vie, et pas seulement aujourd'hui.
    
    **L'échelle de notation :**
    * **1** : Ceci est **complètement faux** pour moi.
    * **2** : C'est **faux dans l'ensemble**, cela ne me ressemble pas vraiment.
    * **3** : C'est **plutôt vrai** que faux.
    * **4** : C'est **modérément vrai**, cela me correspond assez souvent.
    * **5** : C'est **vrai dans l'ensemble**, cela me décrit bien.
    * **6** : Ceci me décrit **parfaitement**, c'est tout à fait moi.
    
    *Vos réponses sont strictement confidentielles et seront analysées uniquement par votre thérapeute.*
    """)
    
    with st.form("form_patient", clear_on_submit=False):
        c1, c2 = st.columns(2)
        nom = c1.text_input("Votre Nom et Prénom *")
        email = c2.text_input("Votre Email *")
        
        reponses = {}
        st.divider()
        
        # Affichage par séries neutres (Série 1, Série 2...) pour éviter les biais
        for i, (domaine, q_dict) in enumerate(YSQ_QUESTIONS.items()):
            with st.container():
                st.markdown(f"#### 📝 Série {i+1}")
                for q_num, q_text in q_dict.items():
                    st.write(f"**{q_num}.** {q_text}")
                    reponses[f"Q{q_num}"] = st.pills(
                        f"Choix Q{q_num}",
                        options=[1, 2, 3, 4, 5, 6],
                        selection_mode="single",
                        label_visibility="collapsed",
                        key=f"q_{q_num}"
                    )
                    st.caption("")
            st.divider()
        
        submitted = st.form_submit_button("Envoyer mes résultats au thérapeute", type="primary")
        
        if submitted:
            missing = [k for k, v in reponses.items() if v is None]
            if not nom or not email:
                st.error("⚠️ Oups ! Vous avez oublié de remplir votre **Nom** ou votre **Email** en haut du formulaire.")
            elif missing:
                st.warning(f"⚠️ Il manque des réponses à **{len(missing)} questions**. Merci de vérifier les séries incomplètes.")
            else:
                with st.spinner("Envoi sécurisé en cours..."):
                    if save_patient_data(nom, email, reponses):
                        st.success("✅ Vos réponses ont été bien reçues et enregistrées ! Merci.")
                        st.balloons()

# ==============================================================================
# 2. ESPACE THÉRAPEUTE (ADMIN)
# ==============================================================================
elif mode == "Espace Thérapeute":
    st.sidebar.divider()
    pwd_input = st.sidebar.text_input("Mot de passe Admin", type="password")
    
    if "ADMIN_PASSWORD" in st.secrets and pwd_input == st.secrets["ADMIN_PASSWORD"]:
        st.header("🔒 Tableau de Bord Clinique Expert")
        
        df = load_all_patients()
        
        if df.empty:
            st.info("Aucun dossier pour le moment.")
        else:
            st.markdown("### Gestion des Dossiers")
            st.dataframe(df[["created_at", "nom", "email"]], use_container_width=True)
            
            st.divider()
            
            c_select, c_action = st.columns([3, 1])
            with c_select:
                patient_options = {f"{row['nom']} ({row['created_at'][:16]})": row['id'] for index, row in df.iterrows()}
                selected_label = st.selectbox("Sélectionner un dossier :", list(patient_options.keys()))
                selected_id = patient_options[selected_label]
            
            with c_action:
                st.write("") 
                st.write("") 
                if st.button("🗑️ Supprimer", type="primary"):
                    if delete_patient(selected_id):
                        st.success("Dossier supprimé.")
                        st.rerun()
            
            st.markdown("---")
            
            # --- BOUTONS D'ACTION ---
            col_analyse, col_raw = st.columns(2)
            patient_data = df[df["id"] == selected_id].iloc[0]
            reponses_dict = json.loads(patient_data["reponses_json"])

            # 1. FICHIER RÉPONSES BRUTES
            def generate_raw_responses():
                doc = Document()
                doc.add_heading(f"Détail des Réponses : {patient_data['nom']}", 0)
                doc.add_paragraph(f"Date : {patient_data['created_at'][:10]}")
                for i, (domaine, q_dict) in enumerate(YSQ_QUESTIONS.items()):
                    doc.add_heading(domaine, level=2)
                    for q_num, q_text in q_dict.items():
                        score = reponses_dict.get(f"Q{q_num}", "-")
                        p = doc.add_paragraph()
                        p.add_run(f"Q{q_num}. ").bold = True
                        p.add_run(f"{q_text} : ")
                        runner = p.add_run(f"[{score}/6]")
                        runner.bold = True
                        if score != "-" and int(score) >= 5: runner.font.color.rgb = RGBColor(255, 0, 0)
                out = BytesIO()
                doc.save(out)
                return out.getvalue()
            
            with col_raw:
                st.download_button("📄 Télécharger les Réponses Brutes", generate_raw_responses(), f"Reponses_{patient_data['nom']}.docx")

            # 2. ANALYSE EXPERTE
            if st.button("📊 Lancer l'Analyse Clinique Complète"):
                resultats = []
                active_codes = []
                
                for domaine, q_dict in YSQ_QUESTIONS.items():
                    code = domaine.split(" : ")[0]
                    nom_sch = domaine.split(" : ")[1]
                    scores = [reponses_dict.get(f"Q{k}", 1) or 1 for k in q_dict.keys()]
                    moy = sum(scores) / len(scores)
                    sev = len([x for x in scores if x >= 5])
                    pct = (sev / len(scores)) * 100
                    etoile = "⭐" if sev > 0 else ""
                    if etoile: active_codes.append(code)
                    
                    niv = "🟢 Faible"
                    if moy > 3.5: niv = "🔴 IMPORTANT"
                    elif moy >= 2.5: niv = "🟡 Moyen"
                    
                    resultats.append({
                        "Code": code, "Schéma": f"{nom_sch} {etoile}", "Moyenne": round(moy, 2),
                        "% Sévérité": f"{round(pct, 1)}%", "Niveau": niv
                    })
                
                df_res = pd.DataFrame(resultats)
                
                c1, c2 = st.columns(2)
                with c1: st.table(df_res)
                with c2:
                    fig_radar = px.line_polar(df_res, r='Moyenne', theta='Code', line_close=True, range_r=[0,6])
                    fig_radar.update_traces(fill='toself', line_color='blue')
                    st.plotly_chart(fig_radar)
                    
                    df_res["Color"] = df_res["Moyenne"].apply(lambda x: "red" if x > 3.5 else ("orange" if x >= 2.5 else "green"))
                    fig_bar = px.bar(df_res, x='Code', y='Moyenne', range_y=[0,6], color="Color", color_discrete_map={"red": "#d32f2f", "orange": "#f57c00", "green": "#388e3c"})
                    fig_bar.update_layout(showlegend=False)
                    st.plotly_chart(fig_bar)

               def generate_word_expert(patient_data, df_res, active_codes):
    doc = Document()
    doc.add_heading(f"Bilan Psychométrique : {patient_data['nom']}", 0)
    doc.add_paragraph(f"Date : {patient_data['created_at'][:10]}")

    # Section 1: Analyse Visuelle
    doc.add_heading('1. Synthèse des Résultats', level=1)
    # (Code insertion graphique - voir partie précédente)
    
    # Section 2: Tableau
    doc.add_heading('2. Tableau de Synthèse', level=1)
    table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
    hdr = table.rows[0].cells; hdr[0].text="Schéma"; hdr[1].text="Score"; hdr[2].text="Niveau"
    for _, row in df_res.iterrows():
        cells = table.add_row().cells
        cells[0].text = row['Schéma']; cells[1].text = str(row['Moyenne'])
        cells[2].text = row['Niveau']

    # Section 3: Analyse Détaillée (COEUR DU SUJET)
    doc.add_heading('3. Analyse Approfondie & Plan d\'Action', level=1)
    
    if active_codes:
        for domain_name, domain_info in YOUNG_DOMAINS_INFO.items():
            match = [c for c in domain_info["codes"] if c in active_codes]
            if match:
                doc.add_heading(domain_name, level=2)
                p_besoin = doc.add_paragraph(domain_info["besoin"]); p_besoin.italic = True
                
                for c in match:
                    inf = DATA_SCHEMAS[c]
                    # En-tête Schéma
                    p = doc.add_paragraph(); p.add_run(f"\n🔹 {inf['titre']}").bold = True
                    p.add_run(f" - {inf['slogan']}").italic = True
                    
                    # 1. Analyse Expert
                    doc.add_paragraph("Analyse Clinique (Expert) :").bold = True
                    doc.add_paragraph(inf['clinique_expert'])
                    doc.add_paragraph("Perspective Théologique :").bold = True
                    doc.add_paragraph(inf['theologie_expert'])
                    
                    # 2. Origines & Symptômes (NOUVEAU - DU FICHIER)
                    doc.add_paragraph("Origines Possibles :").bold = True
                    for o in inf['origines']: doc.add_paragraph(f"- {o}", style='List Bullet')
                    doc.add_paragraph("Signes au Quotidien :").bold = True
                    for s in inf['symptomes']: doc.add_paragraph(f"- {s}", style='List Bullet')
                    
                    # 3. Mécanisme
                    doc.add_paragraph(f"Mécanisme Clé : {inf['mecanisme_titre']}").bold = True
                    doc.add_paragraph(inf['mecanisme_texte'])
                    
                    # 4. Plan d'Action (Fusion)
                    doc.add_paragraph("👉 Plan d'Action Intégratif :").bold = True
                    doc.add_paragraph("Stratégies Thérapeutiques :").italic = True
                    for act in inf['actions_therapeute']: doc.add_paragraph(f"• {act}")
                    doc.add_paragraph("Conseil Pastoral :").italic = True
                    doc.add_paragraph(inf['action_pastorale'])
                    
                    # Verset
                    p_v = doc.add_paragraph(); p_v.add_run("Verset d'ancrage : ").bold = True
                    p_v.add_run(inf['verset']).italic = True
                    doc.add_paragraph("-" * 30)
    else: doc.add_paragraph("Aucun schéma significatif.")
    
    out = BytesIO(); doc.save(out); return out.getvalue()

                st.download_button("📥 Télécharger le Rapport Expert (Complet)", generate_word_expert(), f"Bilan_Expert_{patient_data['nom']}.docx")

    elif pwd_input:
        st.error("Mot de passe incorrect.")

# --- FOOTER ---
st.markdown("---")
with st.expander("🔒 Confidentialité et Protection des Données"):
    st.markdown("""
    **Engagement de confidentialité :**
    * Les données recueillies via ce formulaire sont strictement confidentielles.
    * Seul votre thérapeute a accès aux résultats détaillés.
    * Conformément à la loi, vous pouvez demander à tout moment la suppression intégrale de vos réponses.
    """)
    st.caption("Outil clinique YSQ-L3. Usage professionnel uniquement pour relation d'aide chrétienne- La Barque 2026.")
