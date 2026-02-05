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
from cryptography.fernet import Fernet

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Espace Clinique - YSQ-L3 Intégral", layout="wide")

@st.cache_resource
def init_connection():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except: return None
supabase = init_connection()

# --- 2. SÉCURITÉ & DONNÉES ---
def get_cipher(): return Fernet(st.secrets["ENCRYPTION_KEY"].encode())

def decrypt_reponses(encrypted_json):
    """Lecture hybride : tente de déchiffrer, sinon lit en clair"""
    if not encrypted_json: return {}
    try:
        return json.loads(get_cipher().decrypt(encrypted_json.encode()).decode())
    except:
        try: return json.loads(encrypted_json)
        except: return {}

def save_patient_data(nom, email, reponses_dict):
    if not supabase: return False
    data = {"nom": nom, "email": email, "reponses_json": json.dumps(reponses_dict), "created_at": datetime.now().isoformat()}
    try:
        try: data["reponses_json"] = get_cipher().encrypt(json.dumps(reponses_dict).encode()).decode()
        except: pass
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

# --- 3. BASE DE CONNAISSANCES COMPLETE (AVEC VERSETS ENTIERS) ---
DATA_SCHEMAS = {
    "ED": {
        "titre": "Carence Affective",
        "slogan": "« Personne ne me considère ou ne m'aime vraiment »",
        "clinique_expert": "Ce schéma signale un vide émotionnel précoce. Le patient a intégré la croyance que ses besoins de chaleur, d'empathie et de protection ne seront jamais validés par autrui. Il y a souvent une 'alexithymie' (difficulté à nommer ses besoins) par résignation acquise.",
        "theologie_expert": "Le mensonge racine est l'orphelinat spirituel. La guérison passe par la doctrine de l'Adoption (Romains 8:15). Dieu n'est pas un observateur distant mais un Père qui s'incline pour nourrir (Osée 11:4). Le travail consiste à réapprendre à 'recevoir' sans dissociation.",
        "origines": ["Le soignant ne répondait pas aux besoins d'affection ou de protection.", "Parents froids ou absents émotionnellement.", "L'enfant n'a pas été 'vu' ou compris dans sa singularité."],
        "symptomes": ["Sentiment chronique de vide ou de solitude.", "Considérer ses propres besoins comme sans importance.", "Devenir dépendant ou au contraire contre-dépendant (froid).", "Ressentiment envers les autres qui 'ne donnent pas assez'."],
        "mecanisme_titre": "Les 3 Formes de Privation",
        "mecanisme_texte": "1. Privation d'Attention (manque de chaleur).\n2. Privation d'Empathie (manque d'écoute).\n3. Privation de Protection (manque de guidance).",
        "actions_therapeute": ["Soyez à l'écoute de vos besoins d'attention et de protection.", "Exprimez vos besoins de manière assertive ('J'ai besoin de...').", "Évitez les partenaires froids qui recréent la privation familière."],
        "action_pastorale": "Tenez un journal de vos besoins devant Dieu. Osez demander une petite chose simple à un proche sans vous excuser, comme un acte de foi que vous méritez l'amour.",
        "verset": "Psaume 27:10 - « Car mon père et ma mère m'abandonnent, mais l'Éternel me recueillera. »"
    },
    "AB": {
        "titre": "Abandon / Instabilité",
        "slogan": "« Ne me quitte pas »",
        "clinique_expert": "Perception de l'instabilité fondamentale des liens. Le patient vit dans l'hypervigilance de la perte, alternant entre agrippement anxieux et évitement préventif. La permanence de l'objet (l'autre) n'est pas acquise émotionnellement.",
        "theologie_expert": "L'antidote est la théologie de l'Alliance (Berit). Contrairement aux alliances humaines brisées, l'Alliance divine est unilatérale et irrévocable, fondée sur la fidélité de Dieu et non la performance humaine. Dieu est le Rocher (stabilité ontologique).",
        "origines": ["Décès d'un parent ou départ du foyer dans l'enfance.", "Soignant instable (dépression, alcool) ou imprévisible.", "Surprotection familiale rendant la séparation angoissante."],
        "symptomes": ["S'accrocher aux gens par peur (agrippement).", "Jalousie excessive et possessivité.", "Accusations injustifiées d'infidélité ou d'abandon."],
        "mecanisme_titre": "Le Cycle de l'Abandon",
        "mecanisme_texte": "1. Anxiété (recherche éperdue). 2. Colère/Désespoir (protestation). 3. Détachement (repli sur soi). Ce cycle de l'enfance se répète dans les relations adultes.",
        "actions_therapeute": ["Repérez votre tendance à dramatiser les séparations.", "Apprenez à vous apaiser seul(e) quand l'autre est absent.", "Évitez les partenaires instables ou ambivalents."],
        "action_pastorale": "Pratiquez la 'Solitude Habitée'. Passez 15 min seul(e) en visualisant la présence de Dieu. Rappelez-vous : 'Je ressens de la peur, mais je ne suis pas en danger réel'.",
        "verset": "Hébreux 13:5 - « Je ne te délaisserai point, et je ne t'abandonnerai point. »"
    },
    "MA": {
        "titre": "Méfiance / Abus",
        "slogan": "« Le monde est dangereux »",
        "clinique_expert": "Attente que l'autre va nuire, manipuler ou trahir. Le patient projette une intentionnalité malveillante sur autrui. C'est un schéma de survie traumatique où la confiance est vécue comme une mise en danger.",
        "theologie_expert": "Le monde est déchu, mais Dieu est le Refuge (Mahseh). La guérison demande de renoncer à l'auto-protection cynique pour accepter la protection de Dieu. C'est le passage de la suspicion (peur) au discernement (sagesse) sous le regard de Dieu.",
        "origines": ["Abus physique, sexuel ou verbal dans l'enfance.", "Famille humiliante, sadique ou punitive.", "Trahi ou manipulé par une figure de confiance."],
        "symptomes": ["Hypervigilance ('scanner' les menaces).", "Tests de loyauté envers les autres.", "Attaquer avant d'être attaqué."],
        "mecanisme_titre": "Types d'Abus & Méfiance",
        "mecanisme_texte": "Le schéma naît souvent d'abus physiques, sexuels ou verbaux. La personne reste en mode 'survie', s'attendant à ce que toute gentillesse cache un piège.",
        "actions_therapeute": ["Faites de petits pas pour faire confiance (test de réalité).", "Fixez des limites claires avec les personnes toxiques.", "Développez de la compassion pour l'enfant blessé en vous."],
        "action_pastorale": "Remplacez la suspicion systématique par la prière : 'Seigneur, donne-moi ton discernement'. Déposez les armes de la défensive à la Croix.",
        "verset": "Psaume 62:8 - « En tout temps, peuples, confiez-vous en lui, répandez vos cœurs en sa présence ! Dieu est notre refuge. »"
    },
    "SI": {
        "titre": "Isolement Social",
        "slogan": "« Je n'ai pas ma place ici »",
        "clinique_expert": "Sentiment de différence fondamentale ('Je suis un extraterrestre'). Exclusion du groupe, non par rejet actif, mais par manque d'appartenance ressentie. Le patient se vit comme fondamentalement inadapté au lien social.",
        "theologie_expert":"L'homme a été créé pour la communion. En Christ, la 'différence' n'est plus un motif d'exclusion mais une fonction dans le Corps (1 Corinthiens 12). La rédemption inclut la réintégration dans la famille de Dieu, brisant la malédiction de l'errance.",
        "origines": ["Humiliation ou rejet par les pairs (école).", "Famille différente de la communauté.", "Manque de compétences sociales encouragées."],
        "symptomes": ["Se sentir 'imposteur' ou 'inintéressant' en groupe.", "Évitement systématique des activités sociales.", "Caméléon social pour s'intégrer (perte de soi)."],
        "mecanisme_titre": "Le Cycle de l'Anxiété Sociale",
        "mecanisme_texte": "Anxiété -> Évitement -> Manque de pratique -> Renforcement de l'inadéquation -> Isolement accru.",
        "actions_therapeute": ["Exposez-vous progressivement aux situations évitées.", "Trouvez votre 'tribu' (intérêts communs).", "Entraînez-vous aux compétences sociales."],
        "action_pastorale": "Participez à la vie d'église non pour 'briller' mais pour 'être avec'. Vous êtes membre du Corps : l'œil ne peut dire à la main 'je n'ai pas besoin de toi'.",
        "verset": "Éphésiens 2:19 - « Vous n'êtes plus des étrangers, ni des gens du dehors; mais vous êtes concitoyens des saints. »"
    },
    "DS": {
        "titre": "Imperfection / Honte",
        "slogan": "« Je ne vaux rien »",
        "clinique_expert": "Sentiment d'être intrinsèquement défectueux (Badness). La honte est ici toxique : ce n'est pas 'j'ai fait une erreur' (culpabilité), mais 'je SUIS une erreur'. Cela entraîne une hypersensibilité à la critique et des stratégies de masque.",
        "theologie_expert": "C'est le cœur de l'Évangile : la Justification. Christ a pris notre honte à la croix. Nous sommes déclarés justes non par notre mérite ou par qui nous sommes, mais par l'imputation de sa justice, a cause de qui est Dieu. La valeur du patient ne dépend plus de son 'état', mais de son 'statut' en Christ.",
        "origines": ["Famille critique, humiliante ou punitive.", "Rejet ou manque d'amour par un parent.", "Comparaison défavorable avec la fratrie."],
        "symptomes": ["Cacher sa vraie personnalité (masque).", "Hypersensibilité à la critique.", "Attaquer les autres pour se revaloriser."],
        "mecanisme_titre": "Les 3 Copings de la Honte",
        "mecanisme_texte": "1. Capitulation (autodestruction). 2. Évitement (se cacher). 3. Contre-attaque (narcissisme/critique).",
        "actions_therapeute": ["Cessez de vous comparer aux autres.", "Dressez une liste de vos qualités réelles.", "Acceptez les compliments sans les minimiser."],
        "action_pastorale": "Quand la voix critique attaque, répondez à voix haute : 'Je suis imparfait, mais justifié, lavé et aimé en Christ'. Votre valeur a été fixée à la Croix.",
        "verset": "Sophonie 3:17 - « Il fera de toi le sujet de sa joie... il se réjouira à ton sujet avec des cris de joie. »"
    },
    "FA": {
        "titre": "Échec",
        "slogan": "« Je suis un raté »",
        "clinique_expert": "Croyance en l'incompétence relative aux pairs. Le patient s'identifie à ses échecs scolaires ou professionnels. Il y a souvent un évitement des défis pour ne pas confirmer cette croyance (prophétie auto-réalisatrice).",
        "theologie_expert": "L'idolâtrie de la réussite sociale est brisée. Dieu appelle souvent 'les choses faibles pour confondre les fortes'. Le succès selon le Royaume est la fidélité, pas le résultat visible. La dignité du travail est restaurée comme service, non comme prouesse.",
        "origines": ["Parents très critiques sur les résultats.", "Comparaison défavorable avec les autres enfants.", "Manque de limites ou de discipline dans l'enfance."],
        "symptomes": ["Procrastination par peur de l'échec.", "Minimiser ses propres réussites.", "Abandonner rapidement une tâche."],
        "mecanisme_titre": "La Pensée 'Tout ou Rien'",
        "mecanisme_texte": "Vision dichotomique : 'Si je ne suis pas le meilleur, je suis un échec total'. Cette norme irréaliste condamne à l'échec perçu.",
        "actions_therapeute": ["Reconnaissez la courbe d'apprentissage normale.", "Faites une liste de vos compétences réelles.", "Lancez un hobby sans enjeu de performance."],
        "action_pastorale": "Redéfinissez le succès : pour Dieu, c'est l'amour et l'obéissance. Entreprenez une action en acceptant qu'elle soit 'moyenne' aux yeux du monde, mais faite pour la gloire de Dieu.",
        "verset": "2 Corinthiens 12:9 - « Ma grâce te suffit, car ma puissance s'accomplit dans la faiblesse. »"
    },
    "DI": {
        "titre": "Dépendance / Incompétence",
        "slogan": "« Je n'y arrive pas tout seul »",
        "clinique_expert": "Croyance en l'incapacité à survivre seul. Le patient régresse dans une posture infantile, cherchant une 'figure parentale' pour assumer ses responsabilités. Manque de confiance dans son propre jugement et compétences (Self-Efficacy faible).",
        "theologie_expert": "Dieu nous a donné un esprit de force et de sagesse (2 Tim 1:7). La dépendance saine est verticale (envers Dieu), ce qui permet une autonomie horizontale (envers les hommes). L'Esprit Saint est le 'Paraclet' qui capacite le croyant à marcher.",
        "origines": ["Parents surprotecteurs ('je le fais pour toi').", "Parents qui ne laissaient pas prendre de décisions.", "Manque de conseils pratiques (négligence)."],
        "symptomes": ["Besoin constant d'être rassuré.", "Peur paralysante de prendre une mauvaise décision.", "Laisser les autres diriger sa vie."],
        "mecanisme_titre": "Surprotection vs Négligence",
        "mecanisme_texte": "Soit l'enfant a été étouffé (pas d'autonomie), soit il a été livré à lui-même trop tôt sans guidance (échec appris).",
        "actions_therapeute": ["Listez les tâches où vous dépendez des autres.", "Prenez des petites décisions seul et assumez le résultat.", "Célébrez chaque acte d'autonomie."],
        "action_pastorale": "Prenez une décision quotidienne seul(e) (repas, trajet) en vous confiant au Saint-Esprit qui habite en vous. Vous êtes équipé pour la vie.",
        "verset": "Philippiens 4:13 - « Je puis tout par celui qui me fortifie. »"
    },
    "VU": {
        "titre": "Vulnérabilité au Danger",
        "slogan": "« Une catastrophe arrive »",
        "clinique_expert": "Anxiété catastrophique. Le monde est perçu comme un lieu de dangers imminents (maladie, ruine) qu'on ne peut ni prévoir ni contrôler. Hypervigilance constante du système nerveux.",
        "theologie_expert": "Le problème racine est le contrôle. L'anxiété est une tentative d'assumer la Souveraineté de Dieu. La paix vient non pas de la sécurité totale (impossible), mais de la confiance en la Providence divine qui tient les temps et les circonstances.",
        "origines": ["Parent anxieux ou phobique.", "Traumatisme ou maladie grave dans l'enfance.", "Environnement insécure."],
        "symptomes": ["Scénarios catastrophes.", "Vérifications compulsives.", "Rituels superstitieux."],
        "mecanisme_titre": "Distorsions Cognitives",
        "mecanisme_texte": "1. Catastrophisme (le pire va arriver).\n2. Surestimation du danger / Sous-estimation de ses capacités.\n3. Superstition (pensée magique).",
        "actions_therapeute": ["Analysez la probabilité réelle des catastrophes.", "Réduisez les comportements de vérification.", "Exposition progressive aux situations craintes."],
        "action_pastorale": "Faites une 'Diète de l'info' anxiogène. Tenez un carnet de Gratitude notant 3 protections divines par jour. Ancrez-vous dans la sécurité du présent.",
        "verset": "Psaume 91:4 - « Il te couvrira de ses plumes, et tu trouveras un refuge sous ses ailes. »"
    },
    "EU": {
        "titre": "Fusion / Personnalité Atrophiée",
        "slogan": "« Je ne peux pas vivre sans toi »",
        "clinique_expert": "Symbiose émotionnelle. Le patient n'a pas achevé son processus d'individuation. Il vit par procuration, absorbant les émotions de l'autre. Sentiment de vide existentiel sans la figure d'attachement.",
        "theologie_expert": "Dieu a créé des individus distincts responsables de leurs propres âmes. La fusion est une forme d'idolâtrie relationnelle. Christ appelle à le suivre, ce qui nécessite parfois de 'quitter' (émotionnellement) père et mère pour devenir une personne entière.",
        "origines": ["Parent envahissant ne respectant pas les frontières.", "Culpabilisation quand l'enfant s'autonomise.", "Parent vivant à travers l'enfant."],
        "symptomes": ["Sentiment de vide quand on est seul.", "Imiter les émotions/avis de l'autre.", "Culpabilité intense à avoir une vie privée."],
        "mecanisme_titre": "Identité Non-Développée",
        "mecanisme_texte": "La personne ne sait pas qui elle est sans l'autre. Elle se définit par 'nous' plutôt que 'je'. Risque de relations toxiques.",
        "actions_therapeute": ["Listez vos préférences personnelles (goûts, avis) distincts de l'autre.", "Passez du temps seul pour découvrir qui vous êtes.", "Fixez des limites."],
        "action_pastorale": "Osez exprimer une opinion différente d'un proche sur un sujet mineur. C'est un acte spirituel d'affirmation de la créature unique que Dieu a faite en vous.",
        "verset": "Galates 1:10 - « Est-ce la faveur des hommes que je désire, ou celle de Dieu ? »"
    },
    "SB": {
        "titre": "Assujettissement",
        "slogan": "« Je dois faire ce que tu veux »",
        "clinique_expert": "Soumission forcée pour éviter la colère ou l'abandon. Le patient réprime ses besoins et accumule une colère latente (agressivité passive). Il ne se sent pas le 'droit' d'avoir des limites.",
        "theologie_expert": "La crainte de l'homme est un piège. Le chrétien est serviteur de Dieu, ce qui l'affranchit de l'esclavage des hommes. La vraie soumission est un choix libre d'amour (agapé), pas une contrainte de peur (phobos). Dire 'non' est parfois un acte spirituel.",
        "origines": ["Parent dominant, contrôlant ou punitif.", "Menaces si désaccord.", "Rôle de parentification."],
        "symptomes": ["Peur de dire non.", "Sentiment d'être piégé.", "Accumulation de colère (ressentiment)."],
        "mecanisme_titre": "Le Rôle de la Colère Refoulée",
        "mecanisme_texte": "La soumission crée une dette émotionnelle. La colère refoulée finit par exploser ou devenir des symptômes psychosomatiques.",
        "actions_therapeute": ["Entraînez-vous à dire 'non' sur des petites choses.", "Identifiez vos droits et besoins légitimes.", "Tolérer l'inconfort de ne pas plaire."],
        "action_pastorale": "Exercez-vous au 'Non bienveillant'. Refusez une demande cette semaine. Rappelez-vous que vous servez Dieu, pas l'humeur changeante des autres.",
        "verset": "Galates 5:1 - « C'est pour la liberté que Christ nous a affranchis. »"
    },
    "SS": {
        "titre": "Abnégation",
        "slogan": "« Je suis le sauveur »",
        "clinique_expert": "Le syndrome du Sauveur. Focalisation excessive sur les besoins d'autrui au détriment des siens, motivée par la culpabilité ou le besoin de valorisation narcissique ('Je suis utile donc je suis').",
        "theologie_expert": "Nous ne sommes pas le Messie. Vouloir sauver tout le monde est une limite que seul Dieu peut franchir. L'intendance (gérance) de son propre corps et de son âme est un devoir biblique. L'amour du prochain implique de s'aimer soi-même correctement.",
        "origines": ["Responsabilité excessive d'un proche dans l'enfance.", "Valorisée uniquement quand elle donnait.", "Tempérament hyper-empathique."],
        "symptomes": ["Ne pas savoir recevoir de l'aide.", "Épuisement (burnout).", "Attiré par les personnes à problèmes."],
        "mecanisme_titre": "Frontières (Boundaries)",
        "mecanisme_texte": "Difficulté à fixer des limites. Le sacrifice est motivé par la culpabilité, pas par l'amour libre. C'est une forme de codépendance.",
        "actions_therapeute": ["Équilibrez le donner et le recevoir.", "Demandez-vous : 'Je le fais par envie ou par culpabilité ?'.", "Acceptez que les autres gèrent leurs problèmes."],
        "action_pastorale": "Pratiquez le Sabbat : une demi-journée sans 'servir', juste pour être aimé de Dieu sans rien faire. C'est un acte d'humilité : le monde tourne sans vous.",
        "verset": "Matthieu 22:39 - « Tu aimeras ton prochain comme toi-même. »"
    },
    "EI": {
        "titre": "Inhibition Émotionnelle",
        "slogan": "« Je ne dois pas ressentir »",
        "clinique_expert": "Sur-contrôle des affects. La spontanéité est jugée dangereuse ou honteuse. Le patient présente un 'faux-self' rationnel et froid pour se protéger de la vulnérabilité.",
        "theologie_expert": "Jésus a pleuré, a ressenti la colère et l'angoisse. Les émotions sont créées par Dieu comme des signaux. Les réprimer, c'est vivre dans le mensonge intérieur. La vérité (aletheia) implique l'authenticité émotionnelle devant Dieu (Psaumes de lamentation).",
        "origines": ["Émotions moquées ou punies.", "Famille puritaine ou très rationnelle.", "Peur de ressembler à un parent hystérique."],
        "symptomes": ["Paraître froid ou distant.", "Incapacité à pleurer ou montrer sa joie.", "Accent excessif sur la logique."],
        "mecanisme_titre": "La Roue des Émotions",
        "mecanisme_texte": "Inhibition de la colère, de la joie ou de la vulnérabilité. Tendance à rationaliser pour éviter de ressentir la douleur.",
        "actions_therapeute": ["Utilisez la 'Roue des émotions' pour nommer ce que vous ressentez.", "Tenez un journal émotionnel.", "Recherchez des expériences émotionnelles."],
        "action_pastorale": "Priez avec les Psaumes de lamentation. Osez dire 'Je suis triste' ou 'Je suis en colère' à Dieu. Les émotions ne sont pas des péchés, ce sont des informations.",
        "verset": "Psaume 62:9 - « Répandez votre cœur en sa présence ! Dieu est notre refuge. »"
    },
    "US": {
        "titre": "Exigences Élevées",
        "slogan": "« Ce n'est jamais assez bien »",
        "clinique_expert": "Perfectionnisme pathologique. La valeur personnelle est conditionnelle à la performance. Tyrannie du 'Je dois'. Incapacité à ressentir la satisfaction ou le repos.",
        "theologie_expert": "C'est une forme de légalisme : chercher à se justifier par les œuvres. L'Évangile est la fin de la performance pour le salut. Dieu a institué le Sabbat (repos) pour rappeler que le monde tourne sans nos efforts. La Grâce est l'acceptation de l'imperfection.",
        "origines": ["Amour conditionnel à la réussite.", "Parents perfectionnistes.", "Critique ou honte en cas d'échec."],
        "symptomes": ["Impossible de se détendre.", "Hyper-critique envers soi et les autres.", "Symptômes physiques de stress."],
        "mecanisme_titre": "Les 3 Types de Normes",
        "mecanisme_texte": "1. Compulsivité (ordre). 2. Orientation réussite (travail). 3. Orientation statut (image). C'est une course sans fin.",
        "actions_therapeute": ["Essayez de réduire vos exigences de 10%.", "Listez les avantages et inconvénients de votre pression.", "Forcez-vous à ralentir."],
        "action_pastorale": "Le défi de l'imperfection : laissez volontairement une tâche inachevée (ex: lit mal fait) et observez que Dieu vous aime toujours autant. La grâce suffit.",
        "verset": "Matthieu 11:28 - « Venez à moi, vous tous qui êtes fatigués et chargés, et je vous donnerai du repos. »"
    },
    "ET": {
        "titre": "Droits Personnels / Grandeur",
        "slogan": "« Les règles ne s'appliquent pas à moi »",
        "clinique_expert": "Narcissisme et sentiment de privilège. Le patient refuse les limites communes, manque d'empathie et tolère mal la frustration. C'est souvent une compensation d'un sentiment d'infériorité caché.",
        "theologie_expert": "L'orgueil précède la chute. Le Royaume de Dieu est un 'monde à l'envers' où le plus grand est le serviteur. Reconnaître sa dépendance totale à la grâce de Dieu est le seul remède à l'inflation de l'ego. L'autre n'est pas un outil, mais un porteur de l'image de Dieu.",
        "origines": ["Enfant gâté, sans limites.", "Parents n'ayant pas imposé de conséquences.", "Compensation d'un sentiment de manque."],
        "symptomes": ["Colère si on ne l'obéit pas immédiatement.", "Manque d'empathie.", "Compétitivité excessive."],
        "mecanisme_titre": "Les 3 Types de Droits",
        "mecanisme_texte": "1. Narcissisme pur (je suis spécial). 2. Dépendance (les autres doivent me servir). 3. Impulsivité (je veux tout, tout de suite).",
        "actions_therapeute": ["Mettez-vous à la place des autres (empathie cognitive).", "Demandez un feedback honnête.", "Respectez les règles communes."],
        "action_pastorale": "Pratiquez le service anonyme. Faites une bonne action (vaisselle, don) sans le dire et sans attendre de merci. Écoutez les autres sans ramener la conversation à vous.",
        "verset": "Philippiens 2:3 - « Regardez les autres comme étant au-dessus de vous-mêmes. »"
    },
    "IS": {
        "titre": "Contrôle de soi insuffisant",
        "slogan": "« C'est trop difficile, je m'en fiche »",
        "clinique_expert": "Impulsivité et intolérance à la frustration. Le principe de plaisir domine le principe de réalité. Difficulté à différer la gratification pour un but à long terme.",
        "theologie_expert": "La maîtrise de soi est un fruit de l'Esprit (Galates 5). Ce n'est pas une simple volonté humaine, mais une discipline spirituelle. C'est apprendre à dire 'non' à la chair pour dire 'oui' à la vie. La sagesse biblique valorise la construction patiente.",
        "origines": ["Manque de discipline parentale.", "Négligence ou stress chronique.", "Enfant jamais forcé à tolérer la frustration."],
        "symptomes": ["Procrastination chronique.", "Addictions.", "Évitement systématique de l'inconfort."],
        "mecanisme_titre": "La Stratégie SNAP",
        "mecanisme_texte": "Stop, Notice (Remarquer), Align (Aligner avec valeurs), Process (Agir). Outil pour briser l'impulsion.",
        "actions_therapeute": ["Utilisez la méthode SNAP.", "Fixez des micro-objectifs réalisables.", "Enlevez les distractions."],
        "action_pastorale": "La méthode des 10 minutes : Quand vous voulez abandonner une tâche ou céder à une impulsion, tenez 10 min de plus en priant. C'est un muscle spirituel à exercer.",
        "verset": "Proverbes 25:28 - « Comme une ville forcée et sans murailles, ainsi est l'homme qui n'est pas maître de lui-même. »"
    },
    "AS": {
        "titre": "Recherche d'approbation",
        "slogan": "« Ma valeur dépend de ton regard »",
        "clinique_expert": "Le 'Caméléon'. L'estime de soi est externalisée : elle dépend entièrement du regard de l'autre. Le patient perd son authenticité pour s'adapter aux attentes supposées de l'entourage.",
        "theologie_expert": "La crainte de l'homme est un piège. C'est de l'idolâtrie de l'approbation. Le chrétien vit 'Coram Deo' (devant la face de Dieu). Seule l'approbation du Père ('Tu es mon fils bien-aimé') peut saturer ce besoin et libérer de la tyrannie du regard d'autrui.",
        "origines": ["Amour conditionné à la 'bonne conduite'.", "Parents soucieux des apparences.", "Manque d'attention."],
        "symptomes": ["Changer de personnalité.", "Importance excessive du statut.", "Peur panique de déplaire."],
        "mecanisme_titre": "L'Adaptation Excessive",
        "mecanisme_texte": "Le patient pense : 'Si je suis moi-même, on ne m'aimera pas'. Il développe un 'Faux-Self' pour être validé.",
        "actions_therapeute": ["Demandez-vous : 'Qu'est-ce que JE veux ?'.", "Entraînez-vous à exprimer vos préférences.", "Passez du temps seul."],
        "action_pastorale": "Faites le bien en secret (Matthieu 6). Acceptez un compliment par un simple 'Merci' sans vous en nourrir excessivement ni le rejeter. Votre audience est Dieu seul.",
        "verset": "1 Thessaloniciens 2:4 - « Nous parlons, non pour plaire aux hommes, mais pour plaire à Dieu. »"
    },
    "NP": {
        "titre": "Négativité / Pessimisme",
        "slogan": "« Ça va mal finir »",
        "clinique_expert": "Biais cognitif de focalisation sur le négatif. Attente anxieuse que 'tout va s'effondrer'. Le positif est minimisé ou considéré comme suspect. Souvent lié à une anxiété chronique.",
        "theologie_expert": "Bien que le mal soit réel, la résignation est un déni de la bonté de Dieu et de l'Espérance. La 'joie' biblique est un combat de la foi, une discipline de l'attention (Phil 4:8) pour reconnaître la grâce commune et la providence au milieu des épreuves.",
        "origines": ["Parents pessimistes ou inquiets.", "Enfance marquée par des difficultés.", "Découragement de l'autonomie."],
        "symptomes": ["Filtre négatif.", "Incapacité à se réjouir.", "Plaintes chroniques."],
        "mecanisme_titre": "Les Distorsions Cognitives",
        "mecanisme_texte": "1. Filtre négatif. 2. Généralisation excessive ('toujours'). 3. Catastrophisme. C'est une protection : 'Si je m'attends au pire, je ne serai pas déçu'.",
        "actions_therapeute": ["Examinez les preuves.", "Tenez un journal de gratitude.", "Considérez les exceptions."],
        "action_pastorale": "Contre la rumination, trouvez un aspect positif pour chaque pensée négative. Louez Dieu pour une petite chose précise chaque matin.",
        "verset": "Lamentations 3:21 - « Voici ce que je veux repasser en mon cœur, ce qui me donnera de l'espérance. »"
    },
    "PU": {
        "titre": "Punition",
        "slogan": "« Les erreurs doivent être punies »",
        "clinique_expert": "Intransigeance et dureté. Croyance que l'erreur mérite châtiment. Difficulté à pardonner (à soi et aux autres). Tendance au jugement moralisateur.",
        "theologie_expert": "C'est une incompréhension de la Croix. Christ a pris la punition. Il n'y a plus de condamnation (Rom 8:1). Maintenir une attitude punitive, c'est nier la suffisance du sacrifice de Jésus. Nous sommes appelés à être des canaux de la miséricorde que nous avons reçue.",
        "origines": ["Punitions sévères dans l'enfance.", "Parents impitoyables.", "Manque de droit à l'erreur."],
        "symptomes": ["Rancune tenace.", "Autopunition.", "Jugement sévère des autres."],
        "mecanisme_titre": "Le Cycle de la Rancune",
        "mecanisme_texte": "Standards rigides -> Erreur inévitable -> Colère/Jugement -> Punition. Croyance que la punition 'corrige' le comportement.",
        "actions_therapeute": ["Pratiquez l'auto-compassion.", "Considérez les circonstances atténuantes.", "Pardonnez-vous une erreur passée."],
        "action_pastorale": "Si Jésus a payé, ne cherchez pas à payer encore. Parlez-vous avec la douceur que le Christ utilise pour vous parler.",
        "verset": "Romains 8:1 - « Il n'y a donc maintenant aucune condamnation pour ceux qui sont en Jésus-Christ. »"
    }
}

YOUNG_DOMAINS_INFO = {
    "Domaine I : Séparation et Rejet": {"codes": ["ED", "AB", "MA", "SI", "DS"], "besoin": "Besoin de sécurité, de stabilité, d'affection et d'appartenance."},
    "Domaine II : Manque d'Autonomie": {"codes": ["DI", "VU", "EU", "FA"], "besoin": "Besoin de compétence, d'identité propre et de confiance en soi."},
    "Domaine III : Limites Déficientes": {"codes": ["ET", "IS"], "besoin": "Besoin de limites réalistes, de respect des autres et d'autodiscipline."},
    "Domaine IV : Orientation vers les Autres": {"codes": ["SB", "SS", "AS"], "besoin": "Besoin de liberté d'expression et d'affirmation de ses besoins."},
    "Domaine V : Hypervigilance et Inhibition": {"codes": ["NP", "EI", "US", "PU"], "besoin": "Besoin de spontanéité, de plaisir et de lâcher-prise."}
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
# --- 4. INTERFACE ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3050/3050525.png", width=80)
st.sidebar.title("Navigation")
mode = st.sidebar.radio("Aller vers :", ["Espace Patient", "Espace Thérapeute"])

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
        
        for i, (domaine, q_dict) in enumerate(YSQ_QUESTIONS.items()):
            with st.container():
                st.markdown(f"#### 📝 Série {i+1}")
                for q_num, q_text in q_dict.items():
                    st.write(f"**{q_num}.** {q_text}")
                    reponses[f"Q{q_num}"] = st.pills(f"Choix Q{q_num}", options=[1, 2, 3, 4, 5, 6], selection_mode="single", label_visibility="collapsed", key=f"q_{q_num}")
                    st.caption("")
            st.divider()
        
        if st.form_submit_button("Envoyer mes résultats", type="primary"):
            missing = [k for k, v in reponses.items() if v is None]
            if not nom or not email: st.error("⚠️ Merci de remplir votre nom et email.")
            elif missing: st.warning(f"⚠️ Il manque {len(missing)} réponse(s).")
            else:
                if save_patient_data(nom, email, reponses):
                    st.success("✅ Vos réponses ont été transmises."); st.balloons()

elif mode == "Espace Thérapeute":
    st.sidebar.divider()
    pwd_input = st.sidebar.text_input("Mot de passe Admin", type="password")
    
    if pwd_input == st.secrets["ADMIN_PASSWORD"]:
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
                sel_label = st.selectbox("Sélectionner un dossier :", list(patient_options.keys()))
                sel_id = patient_options[sel_label]
            with c_action:
                st.write(""); st.write("")
                if st.button("🗑️ Supprimer", type="primary"):
                    if delete_patient(sel_id): st.success("Supprimé."); st.rerun()
            
            st.markdown("---")
            col_raw, col_an = st.columns(2)
            pat_data = df[df["id"] == sel_id].iloc[0]
            reponses_dict = decrypt_reponses(pat_data["reponses_json"])

            # 1. FICHIER BRUT
            def gen_raw():
                doc = Document()
                doc.add_heading(f"Réponses Brutes : {pat_data['nom']}", 0)
                for i, (dom, q_d) in enumerate(YSQ_QUESTIONS.items()):
                    doc.add_heading(dom, 2)
                    for q, t in q_d.items():
                        s = reponses_dict.get(f"Q{q}", "-")
                        p = doc.add_paragraph(); p.add_run(f"Q{q}. ").bold=True; p.add_run(f"{t} : [{s}/6]")
                        if s != "-" and str(s).isdigit() and int(s) >= 5: p.runs[-1].font.color.rgb = RGBColor(255, 0, 0)
                out = BytesIO(); doc.save(out); return out.getvalue()
            
            with col_raw: st.download_button("📄 Télécharger Réponses Brutes", gen_raw(), f"Reponses_{pat_data['nom']}.docx")

           # 2. ANALYSE EXPERTE & SÉLECTION
            if st.button("📊 Lancer l'Analyse Clinique"):
                
                # --- A. CALCUL DES SCORES ---
                resultats = []
                # Liste pour la pré-sélection automatique (ceux > 2.5 ou avec des 5/6)
                pre_selection = []
                
                for domaine, q_dict in YSQ_QUESTIONS.items():
                    code = domaine.split(" : ")[0]
                    nom_sch = domaine.split(" : ")[1]
                    scores = [int(reponses_dict.get(f"Q{k}", 1) or 1) for k in q_dict.keys()]
                    
                    moy = sum(scores) / len(scores)
                    sev = len([x for x in scores if x >= 5])
                    pct = (sev / len(scores)) * 100
                    
                    # Logique des étoiles
                    etoile = " ⭐" if sev > 0 else ""
                    
                    # Logique des niveaux
                    if moy > 3.5:
                        niv = "🔴 IMPORTANT"
                        pre_selection.append(code) # On pré-sélectionne
                    elif moy >= 2.5:
                        niv = "🟡 Moyen"
                        pre_selection.append(code) # On pré-sélectionne aussi les moyens
                    else:
                        niv = "🟢 Faible"
                        # On ajoute aussi si faible mais avec au moins une réponse sévère (5 ou 6)
                        if sev > 0: pre_selection.append(code)

                    resultats.append({
                        "Code": code,
                        "Schéma": f"{nom_sch}{etoile}",
                        "Moyenne": round(moy, 2),
                        "% Sévérité": f"{round(pct, 1)}%",
                        "Niveau": niv
                    })
                
                df_res = pd.DataFrame(resultats)
                
                # --- B. AFFICHAGE TABLEAU & GRAPHIQUES ---
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("### 📋 Synthèse des Scores")
                    st.table(df_res)
                with c2:
                    st.markdown("### 📈 Visualisation")
                    # Radar
                    fig_radar = px.line_polar(df_res, r='Moyenne', theta='Code', line_close=True, range_r=[0,6])
                    fig_radar.update_traces(fill='toself', line_color='blue')
                    st.plotly_chart(fig_radar)
                    
                    # Barres (Dégradé)
                    df_res["Color"] = df_res["Moyenne"].apply(lambda x: "red" if x > 3.5 else ("orange" if x >= 2.5 else "green"))
                    fig_bar = px.bar(df_res, x='Code', y='Moyenne', range_y=[0,6], 
                                     color="Color", 
                                     color_discrete_map={"red": "#d32f2f", "orange": "#f57c00", "green": "#388e3c"})
                    fig_bar.update_layout(showlegend=False)
                    st.plotly_chart(fig_bar)

                # --- C. SÉLECTION MANUELLE POUR LE RAPPORT (NOUVEAU !) ---
                st.divider()
                st.markdown("### 📝 Personnalisation du Rapport")
                st.info("Le logiciel a pré-sélectionné les schémas Moyens et Importants. Vous pouvez ajuster cette liste avant de générer le Word.")
                
                # Le widget multiselect
                codes_disponibles = df_res["Code"].tolist()
                selection_finale = st.multiselect(
                    "Sélectionnez les schémas à inclure dans le dossier patient :",
                    options=codes_disponibles,
                    default=pre_selection # Pré-rempli avec notre logique intelligente
                )

                # Fonction de génération (utilise maintenant 'selection_finale')
                def gen_expert():
                    doc = Document()
                    doc.add_heading(f"Bilan Psychométrique : {pat_data['nom']}", 0)
                    doc.add_paragraph(f"Date : {pat_data['created_at'][:10]}")
                    
                    doc.add_heading('1. Synthèse Visuelle', 1)
                    try: 
                        doc.add_picture(BytesIO(fig_radar.to_image(format="png", engine="kaleido")), width=Inches(4.5))
                        doc.add_picture(BytesIO(fig_bar.to_image(format="png", engine="kaleido")), width=Inches(4.5))
                    except: doc.add_paragraph("[Graphiques indisponibles]")

                    doc.add_heading('2. Tableau des Scores', 1)
                    tbl = doc.add_table(rows=1, cols=4); tbl.style = 'Table Grid'
                    h = tbl.rows[0].cells; h[0].text="Code"; h[1].text="Schéma"; h[2].text="Score"; h[3].text="Niveau"
                    for _, r in df_res.iterrows():
                        row = tbl.add_row().cells
                        row[0].text = str(r['Code']); row[1].text = str(r['Schéma']); row[2].text = str(r['Moyenne'])
                        run = row[3].paragraphs[0].add_run(r['Niveau']); run.bold = True
                        if "IMPORTANT" in r['Niveau']: run.font.color.rgb = RGBColor(255, 0, 0)
                        elif "Moyen" in r['Niveau']: run.font.color.rgb = RGBColor(255, 140, 0)
                        else: run.font.color.rgb = RGBColor(0, 128, 0)

                    doc.add_heading('3. Analyse Intégrale', 1)
                    
                    # On utilise la sélection manuelle ici !
                    if selection_finale:
                        for d_name, d_info in YOUNG_DOMAINS_INFO.items():
                            # On ne garde que ceux qui sont dans la sélection finale
                            match = [c for c in d_info["codes"] if c in selection_finale]
                            if match:
                                doc.add_heading(d_name, 2)
                                doc.add_paragraph(d_info["besoin"]).italic = True
                                for c in match:
                                    inf = DATA_SCHEMAS.get(c, {})
                                    if not inf: continue
                                    
                                    p = doc.add_paragraph(); p.add_run(f"\n🔹 {inf['titre']}").bold = True
                                    p.add_run(f" - {inf['slogan']}").italic = True
                                    p.add_run(f" (Score: {df_res.loc[df_res['Code'] == c, 'Moyenne'].values[0]})")
                                    
                                    doc.add_paragraph("🧠 Analyse Clinique (Expert) :").bold = True
                                    doc.add_paragraph(inf['clinique_expert'])
                                    
                                    doc.add_paragraph("✝️ Perspective Théologique :").bold = True
                                    doc.add_paragraph(inf['theologie_expert'])
                                    
                                    doc.add_paragraph("🌱 Origines & Développement :").bold = True
                                    for o in inf['origines']: doc.add_paragraph(f"- {o}", style='List Bullet')
                                    
                                    doc.add_paragraph("⚠️ Signes au Quotidien :").bold = True
                                    for s in inf['symptomes']: doc.add_paragraph(f"- {s}", style='List Bullet')
                                    
                                    doc.add_paragraph(f"⚙️ Mécanisme Clé : {inf['mecanisme_titre']}").bold = True
                                    doc.add_paragraph(inf['mecanisme_texte'])
                                    
                                    doc.add_paragraph("👉 Plan d'Action Intégratif :").bold = True
                                    doc.add_paragraph("🛠️ Stratégies Thérapeutiques :").italic = True
                                    for act in inf['actions_therapeute']: doc.add_paragraph(f"• {act}")
                                    
                                    doc.add_paragraph("🙏 Conseil Pastoral :").italic = True
                                    doc.add_paragraph(inf['action_pastorale'])
                                    
                                    p_v = doc.add_paragraph(); p_v.add_run("📖 Verset d'ancrage : ").bold = True
                                    p_v.add_run(inf['verset']).italic = True
                                    doc.add_paragraph("-" * 30)
                    else: doc.add_paragraph("Aucun schéma sélectionné.")
                    
                    out = BytesIO(); doc.save(out); return out.getvalue()

                st.download_button("📥 Télécharger Rapport Expert", gen_expert(), f"Bilan_{pat_data['nom']}.docx")
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
