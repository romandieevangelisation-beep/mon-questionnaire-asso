import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from docx import Document
from docx.shared import Inches
from io import BytesIO
import json
from supabase import create_client

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Espace Clinique - Questionnaire YSQ-L3", layout="wide")

# --- CONNEXION SÉCURISÉE (SUPABASE) ---
@st.cache_resource
def init_connection():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except KeyError:
        st.error("Erreur critique : Les secrets Supabase ne sont pas configurés dans Streamlit Cloud.")
        return None

supabase = init_connection()

# --- FONCTIONS DE GESTION DES DONNÉES ---
def save_patient_data(nom, email, reponses_dict):
    if not supabase: return False
    data = {
        "nom": nom,
        "email": email,
        "reponses_json": json.dumps(reponses_dict),
        "created_at": datetime.now().isoformat()
    }
    try:
        supabase.table("patients_ysq").insert(data).execute()
        return True
    except Exception as e:
        st.error(f"Erreur d'enregistrement : {e}")
        return False

def load_all_patients():
    if not supabase: return pd.DataFrame()
    try:
        # On récupère aussi l'ID pour pouvoir supprimer
        response = supabase.table("patients_ysq").select("*").order("created_at", desc=True).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        return pd.DataFrame()

def delete_patient(patient_id):
    """Supprime un patient de la base de données via son ID"""
    if not supabase: return False
    try:
        supabase.table("patients_ysq").delete().eq("id", patient_id).execute()
        return True
    except Exception as e:
        st.error(f"Erreur de suppression : {e}")
        return False

# --- DÉFINITIONS EXPERTES & PASTORALES ---
INTERPRETATIONS_EXPERTES = {
    "ED": {
        "titre": "Carence Affective",
        "desc": "Sentiment profond que vos besoins de sécurité, d'affection et d'écoute ne seront jamais comblés. Vous avez appris à survivre sans demander, pensant que personne ne se soucie vraiment de vous.",
        "verset": "Psaume 27:10",
        "biblique": "Si les figures parentales humaines ont pu faillir, Dieu se révèle comme le Père parfait qui 'recueille'. La guérison passe par l'apprentissage de la réception : oser croire que vous êtes digne de soin et laisser le Père céleste combler ce vide émotionnel par Son Esprit."
    },
    "AB": {
        "titre": "Abandon / Instabilité",
        "desc": "Peur viscérale que les relations soient fragiles et que vous finissiez seul(e). Cela crée une anxiété chronique et une tendance à 's'agripper' aux autres ou à tester leur fidélité.",
        "verset": "Hébreux 13:5",
        "biblique": "L'insécurité relationnelle se soigne par la certitude de l'Alliance. Dieu a promis : 'Je ne te délaisserai point'. Contrairement aux humains, Sa fidélité est immuable. Ancrez votre sécurité dans ce lien indestructible pour apaiser votre anxiété relationnelle."
    },
    "MA": {
        "titre": "Méfiance / Abus",
        "desc": "Attente que les autres vont intentionnellement vous blesser, vous manipuler ou vous trahir. Vous restez sur vos gardes, percevant le monde comme un lieu hostile.",
        "verset": "Psaume 56:4-5",
        "biblique": "Dieu est le seul refuge totalement sûr. La guérison spirituelle implique de déposer les armes de la défensive aux pieds de Christ, pour réapprendre progressivement le discernement (sagesse) plutôt que la suspicion (peur), sous Sa protection."
    },
    "SI": {
        "titre": "Isolement Social",
        "desc": "Sentiment de ne pas appartenir, d'être différent, 'à part' du reste de l'humanité. Une solitude qui n'est pas physique, mais existentielle.",
        "verset": "Éphésiens 2:19",
        "biblique": "En Christ, vous n'êtes plus étranger. Vous avez une citoyenneté céleste et une place légitime dans la famille de Dieu. L'Église est le lieu où votre différence est accueillie comme une richesse membre du Corps, et non comme une anomalie."
    },
    "DS": {
        "titre": "Imperfection / Honte",
        "desc": "Croyance profonde d'être intérieurement défectueux, indigne d'amour ou 'mauvais'. Crainte constante que si l'on vous voit tel que vous êtes, on vous rejettera.",
        "verset": "Romains 8:1 / Sophonie 3:17",
        "biblique": "La honte dit 'je suis une erreur', Dieu dit 'tu es ma créature merveilleuse'. La Justification par la foi signifie que vous êtes déclaré juste et pur grâce à Christ. Votre valeur ne dépend pas de votre perfection, mais de Son amour inconditionnel."
    },
    "FA": {
        "titre": "Échec",
        "desc": "Sentiment d'être incompétent par rapport à vos pairs, de ne pas avoir de talent ou d'intelligence. Croyance que l'échec est inévitable.",
        "verset": "2 Corinthiens 12:9",
        "biblique": "Le Royaume de Dieu renverse les valeurs du monde : la puissance s'accomplit dans la faiblesse. Votre identité ne se trouve pas dans la performance sociale ou professionnelle, mais dans votre filiation divine. Dieu utilise souvent ce que le monde considère comme faible."
    },
    "DI": {
        "titre": "Dépendance / Incompétence",
        "desc": "Sentiment d'être incapable de gérer le quotidien sans l'aide d'autrui. Manque de confiance dans son propre jugement.",
        "verset": "Philippiens 4:13",
        "biblique": "Dieu ne nous a pas donné un esprit de timidité, mais de force. La dépendance saine est envers Dieu, qui vous donne la sagesse. L'Esprit Saint habite en vous pour vous guider et vous rendre capable de prendre vos propres responsabilités d'adulte."
    },
    "VU": {
        "titre": "Vulnérabilité au danger",
        "desc": "Peur exagérée et constante qu'une catastrophe (maladie, ruine, agression) soit imminente. État d'hyper-vigilance.",
        "verset": "Psaume 91:4",
        "biblique": "L'antidote à la peur n'est pas le contrôle, mais la confiance en la Souveraineté de Dieu. Il est votre abri. Apprendre à remettre l'avenir, que vous ne pouvez contrôler, entre les mains de Celui qui tient l'univers."
    },
    "EU": {
        "titre": "Fusion / Personnalité Atrophiée",
        "desc": "Manque d'individualité, relation symbiotique avec un parent ou un conjoint. Sentiment de ne pas exister sans l'autre.",
        "verset": "Galates 1:10",
        "biblique": "Vous êtes une création unique voulue par Dieu, distincte de vos parents ou conjoint. Christ vous appelle à le suivre Lui, ce qui nécessite de devenir une personne à part entière. L'amour vrai laisse l'autre exister."
    },
    "SB": {
        "titre": "Assujettissement",
        "desc": "Vous vous soumettez aux autres par peur des conséquences (colère, rejet). Vous taisez vos propres besoins et émotions.",
        "verset": "Galates 5:1",
        "biblique": "'C'est pour la liberté que Christ nous a affranchis'. Dieu ne veut pas d'une soumission basée sur la peur, mais d'un service basé sur l'amour. Vous avez le droit d'exister, d'avoir des limites et de dire non, car vous êtes serviteur de Dieu avant d'être serviteur des hommes."
    },
    "SS": {
        "titre": "Abnégation",
        "desc": "Sacrifice excessif de vos besoins pour satisfaire ceux des autres, souvent par culpabilité. Le syndrome du 'Sauveur'.",
        "verset": "Matthieu 22:39",
        "biblique": "Le commandement est d'aimer son prochain 'comme soi-même', pas 'au lieu de soi-même'. Votre corps et votre âme sont le temple du Saint-Esprit ; en prendre soin n'est pas de l'égoïsme, c'est de la bonne gestion pour pouvoir servir durablement."
    },
    "EI": {
        "titre": "Inhibition Émotionnelle",
        "desc": "Verrouillage des émotions et de la spontanéité par peur de la honte ou de perdre le contrôle. Apparence froide ou trop rationnelle.",
        "verset": "Psaume 62:9",
        "biblique": "'Répandez votre cœur devant Lui'. Les Psaumes nous montrent une vie émotionnelle riche devant Dieu (joie, colère, tristesse). Jésus a pleuré. La maturité spirituelle inclut la vérité émotionnelle, car la vérité libère."
    },
    "US": {
        "titre": "Exigences Élevées",
        "desc": "Tyrannie du 'je dois'. Perfectionnisme, règles rigides et impossibilité de se détendre. La valeur personnelle dépend de la performance.",
        "verset": "Éphésiens 2:8-9",
        "biblique": "C'est par la grâce que vous êtes sauvés, non par les œuvres. Dieu n'est pas un patron exigeant qui attend la performance, mais un Père qui donne le repos. Accepter la grâce, c'est accepter d'être imparfait et pourtant pleinement aimé."
    },
    "ET": {
        "titre": "Droits Personnels / Grandeur",
        "desc": "Sentiment de supériorité, d'avoir des droits spéciaux, manque d'empathie pour les besoins d'autrui. Intolérance à la frustration.",
        "verset": "Philippiens 2:3-5",
        "biblique": "L'attitude de Christ, Roi de gloire qui s'est fait serviteur, est le remède. La vraie grandeur dans le Royaume est l'humilité et le service. Reconnaître que tout ce que nous avons est un don immérité de Dieu brise l'orgueil."
    },
    "IS": {
        "titre": "Contrôle de soi insuffisant",
        "desc": "Difficulté à tolérer la frustration, impulsivité, difficulté à tenir ses engagements ou à discipliner ses émotions.",
        "verset": "Galates 5:22",
        "biblique": "La maîtrise de soi est un fruit de l'Esprit, pas juste un effort de volonté. C'est Dieu qui produit en nous le vouloir et le faire. La croissance spirituelle implique d'apprendre à différer la gratification immédiate pour une joie supérieure en Dieu."
    },
    "AS": {
        "titre": "Recherche d'approbation",
        "desc": "Votre estime de soi dépend entièrement du regard, de l'attention et de la validation des autres. Caméléon social.",
        "verset": "Jean 5:44",
        "biblique": "Comment pouvez-vous croire, vous qui tirez votre gloire les uns des autres ? La guérison vient quand le regard de Dieu ('Tu es mon enfant bien-aimé') devient plus lourd et plus réel que le regard des hommes. Vous vivez pour l'Audience d'Un seul."
    },
    "NP": {
        "titre": "Négativité / Pessimisme",
        "desc": "Focalisation obsessionnelle sur le négatif (douleur, mort, perte) en minimisant le positif. Attente constante que tout tourne mal.",
        "verset": "Lamentations 3:21-23",
        "biblique": "La foi chrétienne est foncièrement réaliste mais porteuse d'Espérance. 'Voici ce que je veux repasser en mon cœur, ce qui me donnera de l'espérance : les bontés de l'Éternel ne sont pas épuisées'. C'est une discipline spirituelle que d'entraîner son regard à voir la grâce active."
    },
    "PU": {
        "titre": "Punition",
        "desc": "Croyance que les erreurs (les siennes ou celles des autres) méritent une punition sévère. Difficulté à pardonner, intolérance.",
        "verset": "Romains 8:1",
        "biblique": "Il n'y a plus de condamnation. À la croix, Christ a pris la punition que nous méritions. Si Dieu ne nous traite pas selon nos fautes, nous sommes appelés à relâcher ce désir de justice punitive envers nous-mêmes et envers les autres. C'est la voie de la Miséricorde."
    }
}

# --- LES 232 QUESTIONS (STANDARD CLINIQUE YSQ-L3) ---
YSQ_QUESTIONS = {
    "ED : Carence affective": {
        1: "Je n'ai pas eu quelqu'un pour prendre soin de moi, partager sa vie avec moi, ou se soucier réellement de tout ce qui m'arrivait.",
        2: "Je n'ai pas reçu suffisamment d'affection et de chaleur ou d'amour.",
        3: "Pour l'essentiel, je n'ai eu personne sur qui compter pour recevoir des conseils et un soutien affectif.",
        4: "La plupart du temps, je n'ai eu personne pour me nourrir, m'épauler ou se soucier de tout ce qui m'arrivait.",
        5: "La plupart du temps, je n'ai eu personne pour me nourrir, m'épauler ou se soucier de tout ce qui m'arrivait.",
        6: "Pour l'essentiel, les gens n'ont pas été là pour moi, pour me donner de la chaleur, de la tendresse et de l'affection.",
        7: "J'ai eu le sentiment que je n'avais personne vers qui me tourner pour recevoir des conseils ou une orientation.",
        8: "Je n'ai pas eu quelqu'un qui m'écoute vraiment, me comprenne ou soit sensible à mes vrais besoins et à mes sentiments.",
        9: "J'ai rarement eu quelqu'un de fort pour me donner des conseils avisés ou me dire quoi faire quand j'étais désemparé(e)."
    },
    "AB : Abandon / Instabilité": {
        10: "Je m'inquiète beaucoup à l'idée que les gens que j'aime vont mourir ou me quitter.",
        11: "Je m'accroche aux gens parce que j'ai peur qu'ils me quittent.",
        12: "Je crains que les gens que j'aime ne trouvent quelqu'un d'autre qu'ils préféreront et ne m'abandonnent.",
        13: "Les gens qui m'ont été proches ont toujours été imprévisibles ; un moment disponibles, le moment d'après fâchés ou absents.",
        14: "J'ai tellement besoin des autres que je m'inquiète de les perdre.",
        15: "Je me sens désespéré(e) quand quelqu'un que j'aime s'éloigne de moi, même brièvement.",
        16: "Je tombe amoureux(se) de gens qui ne peuvent pas s'engager avec moi de façon stable.",
        17: "La plupart des gens sont imprévisibles concernant leurs sentiments envers moi.",
        18: "En fin de compte, je serai seul(e).",
        19: "Quand je sens que quelqu'un à qui je tiens s'éloigne de moi, je deviens désespéré(e).",
        20: "Parfois, j'ai tellement peur que les gens me quittent que je les fais fuir.",
        21: "Je ne peux pas compter sur les gens qui me soutiennent pour être là de façon permanente.",
        22: "Je ne peux pas me laisser aller à être moi-même ou les gens me quitteront.",
        23: "Je suis obsédé(e) par l'idée que mes relations vont se terminer.",
        24: "Je n'ai pas de base stable affectivement.",
        25: "Je ne peux pas vivre sans quelqu'un qui m'aime.",
        26: "J'ai besoin que les autres me rassurent constamment sur le fait qu'ils ne vont pas me quitter."
    },
    "MA : Méfiance / Abus": {
        27: "J'ai l'impression que les autres vont profiter de moi.",
        28: "Je sens que je dois me protéger des autres.",
        29: "Je pense que si je laisse les gens m'approcher, ils me feront du mal.",
        30: "Si quelqu'un est gentil, je me demande ce qu'il veut.",
        31: "Je teste les gens pour voir s'ils sont honnêtes.",
        32: "Je suis très méfiant(e) vis-à-vis des motifs des autres.",
        33: "Je pense que les gens pensent d'abord à eux-mêmes.",
        34: "J'ai été maltraité(e), abusé(e) ou négligé(e) par des gens importants pour moi.",
        35: "Je me sens souvent trahi(e) par les autres.",
        36: "Je suis sur mes gardes la plupart du temps.",
        37: "Je ne peux faire confiance à personne.",
        38: "Je pense que les gens me feront du mal si j'en laisse l'occasion.",
        39: "Je crains d'être attaqué(e) physiquement ou verbalement par les autres.",
        40: "J'ai l'impression que les gens se moquent de moi derrière mon dos.",
        41: "Je pense que les gens m'utiliseront à leurs propres fins.",
        42: "Je suis souvent sur la défensive avec les gens.",
        43: "Le monde est un endroit dangereux."
    },
    "SI : Isolement social": {
        44: "Je ne me sens pas à ma place dans les groupes.",
        45: "Je me sens différent(e) des autres.",
        46: "Je me sens isolé(e) des autres.",
        47: "Je n'appartiens à aucun groupe ou communauté.",
        48: "Je me sens seul(e) même quand je suis avec des gens.",
        49: "Je me sens étranger(ère) partout.",
        50: "Personne ne me comprend vraiment.",
        51: "Je suis en marge de la société.",
        52: "Je me sens ennuyeux(se) ou inintéressant(e) dans les situations sociales.",
        53: "Je ne sais pas quoi dire dans les situations sociales."
    },
    "DS : Imperfection / Honte": {
        54: "Si les gens me connaissaient vraiment, ils ne m'aimeraient pas.",
        55: "J'ai des secrets que je ne veux pas que les autres connaissent.",
        56: "Je suis fondamentalement défectueux(se) ou imparfait(e).",
        57: "Je ne mérite pas d'être aimé(e).",
        58: "J'ai honte de moi.",
        59: "Je suis une mauvaise personne.",
        60: "Je cache mes défauts aux autres.",
        61: "Je suis indigne de respect.",
        62: "Je me sens humilié(e) par mes échecs ou mes défauts.",
        63: "Je suis très critique envers moi-même.",
        64: "Je me sens coupable d'être qui je suis.",
        65: "Je ne suis pas à la hauteur.",
        66: "Je me dévalorise souvent.",
        67: "Je crains que mes défauts ne soient exposés.",
        68: "Je suis gêné(e) par moi-même."
    },
    "FA : Échec": {
        69: "Je suis moins compétent(e) que les autres dans le domaine du travail (ou scolaire).",
        70: "J'ai échoué dans tout ce que j'ai entrepris.",
        71: "Je ne suis pas aussi intelligent(e) que la plupart des gens.",
        72: "Je n'ai pas de talent particulier.",
        73: "Je ne réussirai jamais rien d'important.",
        74: "Je me sens bête comparé(e) aux autres.",
        75: "Je suis un(e) raté(e).",
        76: "Je ne suis pas capable de travailler aussi bien que les autres.",
        77: "Je me sens inférieur(e) professionnellement aux autres."
    },
    "DI : Dépendance / Incompétence": {
        78: "Je ne me sens pas capable de me débrouiller seul(e) dans la vie quotidienne.",
        79: "J'ai besoin de l'aide des autres pour prendre des décisions.",
        80: "Je ne sais pas gérer ma vie quotidienne (finances, réparations, etc.) sans aide.",
        81: "Je me sens comme un enfant quand il s'agit de responsabilités d'adulte.",
        82: "J'ai peur de faire des erreurs graves si je n'ai pas de conseils.",
        83: "Je ne peux pas survivre sans quelqu'un pour s'occuper de moi.",
        84: "Je ne fais pas confiance à mon propre jugement.",
        85: "Je me sens dépassé(e) par les responsabilités de la vie.",
        86: "Je cherche toujours quelqu'un pour me dire quoi faire.",
        87: "Je me sens vulnérable face aux défis de la vie.",
        88: "Je ne sais pas résoudre les problèmes courants.",
        89: "Je panique quand je suis seul(e) face à un défi.",
        90: "Je laisse les autres prendre les commandes de ma vie.",
        91: "Je ne suis pas autonome.",
        92: "Je me sens incompétent(e) dans la plupart des domaines."
    },
    "VU : Vulnérabilité": {
        93: "Je ne peux pas m'empêcher de penser qu'une catastrophe va arriver.",
        94: "J'ai peur de tomber malade gravement ou d'avoir une attaque.",
        95: "J'ai peur d'être agressé(e) ou volé(e).",
        96: "Je crains de perdre tout mon argent et de devenir pauvre.",
        97: "Je suis obsédé(e) par la sécurité.",
        98: "Je surveille mon corps excessivement pour détecter des maladies.",
        99: "Je crains de devenir fou/folle ou de perdre le contrôle.",
        100: "Je panique facilement face à des dangers potentiels.",
        101: "Le monde est plein de dangers imprévisibles.",
        102: "Je ne me sens jamais en sécurité.",
        103: "Je crains les accidents (avion, voiture, ascenseur).",
        104: "Je suis très anxieux(se) la plupart du temps."
    },
    "EU : Fusion / Personnalité atrophiée": {
        105: "Je ne sais pas qui je suis sans les autres.",
        106: "Je suis trop impliqué(e) dans la vie de mes proches.",
        107: "Je me sens coupable d'avoir des secrets pour mes proches.",
        108: "Je ne peux pas vivre sans mon partenaire ou mes parents.",
        109: "Je n'ai pas d'identité propre séparée de mes proches.",
        110: "Je ressens les émotions des autres comme si c'étaient les miennes.",
        111: "Je me sens vide quand je suis seul(e).",
        112: "Je fusionne avec les gens que j'aime.",
        113: "Je ne sais pas ce que je veux vraiment pour moi-même.",
        114: "Je vis ma vie à travers les autres.",
        115: "Je n'ai pas de limites claires entre moi et les autres."
    },
    "SB : Assujettissement": {
        116: "Je laisse les autres prendre les décisions à ma place.",
        117: "Je n'ose pas dire non aux demandes des autres.",
        118: "Je crains la colère ou le rejet des autres si je ne suis pas d'accord.",
        119: "Je sacrifie mes besoins pour éviter les conflits.",
        120: "Je me sens coupable de penser à moi.",
        121: "Je me laisse dominer par les autres.",
        122: "Je n'exprime pas mes vrais désirs ou opinions.",
        123: "Je fais beaucoup de choses pour plaire aux autres, même si je ne veux pas.",
        124: "Je me sens piégé(e) dans des relations où je dois céder.",
        125: "Je refoule ma colère pour ne pas faire d'histoires."
    },
    "SS : Abnégation": {
        126: "Je m'occupe des besoins des autres plus que des miens.",
        127: "Je suis celui/celle vers qui on se tourne pour écouter les problèmes.",
        128: "Je donne beaucoup aux autres et je reçois peu en retour.",
        129: "Je me sens responsable du bien-être et du bonheur des autres.",
        130: "Je ne peux pas supporter de voir les autres souffrir sans rien faire.",
        131: "Je suis l'oreille attentive de tout le monde.",
        132: "Je m'épuise pour aider les autres.",
        133: "Je néglige ma santé ou mes intérêts pour aider les autres.",
        134: "Je ne demande jamais d'aide pour moi-même.",
        135: "Je suis 'trop bon(ne)' avec les autres.",
        136: "Je me sens égoïste si je prends du temps pour moi.",
        137: "Je veux sauver les autres de leurs problèmes.",
        138: "Mes besoins passent toujours en dernier.",
        139: "Je suis très empathique à la douleur des autres.",
        140: "Je ne sais pas recevoir de l'aide ou de l'affection.",
        141: "Je suis le pilier de la famille ou du groupe.",
        142: "Je me sens utile seulement quand j'aide quelqu'un."
    },
    "EI : Inhibition émotionnelle": {
        143: "Je cache mes sentiments aux autres.",
        144: "Je ne montre pas ma colère ou mon irritation.",
        145: "Je suis gêné(e) par les effusions d'émotions ou d'affection.",
        146: "Je contrôle tout ce que je ressens.",
        147: "Je parais froid(e) ou distant(e) aux yeux des autres.",
        148: "Je ne pleure jamais devant les autres.",
        149: "Je suis très rationnel(le) et logique, je n'aime pas les émotions.",
        150: "J'ai peur de perdre le contrôle si je laisse sortir mes émotions.",
        151: "Je garde tout à l'intérieur."
    },
    "US : Exigences élevées": {
        152: "Je dois être le/la meilleur(e) dans ce que je fais.",
        153: "Je ne suis jamais satisfait(e) de mes accomplissements.",
        154: "Je travaille tout le temps, je ne prends pas de vacances.",
        155: "Je suis perfectionniste, tout doit être impeccable.",
        156: "Les erreurs sont inacceptables pour moi.",
        157: "Je me mets une pression énorme pour réussir.",
        158: "Je suis très critique envers les autres s'ils ne sont pas performants.",
        159: "Il faut que tout soit en ordre et organisé.",
        160: "Je ne sais pas me détendre ou jouer.",
        161: "Je cours toujours après le temps.",
        162: "La médiocrité m'insupporte.",
        163: "Je sacrifie ma vie personnelle et ma santé pour réussir.",
        164: "Je dois toujours faire plus, ce n'est jamais assez.",
        165: "Je suis très compétitif(ve).",
        166: "L'échec est une catastrophe pour moi.",
        167: "Je suis esclave de mes devoirs et responsabilités."
    },
    "ET : Droits personnels / Grandeur": {
        168: "Je suis une personne spéciale et je mérite mieux que les autres.",
        169: "Les règles ne s'appliquent pas à moi comme aux autres.",
        170: "Je mérite un traitement de faveur.",
        171: "Je ne supporte pas qu'on me dise non ou qu'on me limite.",
        172: "Je fais passer mes besoins avant ceux des autres.",
        173: "Je me sens supérieur(e) aux autres.",
        174: "Je m'énerve si je n'obtiens pas ce que je veux immédiatement.",
        175: "Je manipule les autres pour obtenir ce que je veux.",
        176: "Je n'ai pas à m'excuser ou à me justifier.",
        177: "Je suis destiné(e) à de grandes choses et les gens ordinaires ne me comprennent pas.",
        178: "Les autres sont là pour me servir ou m'admirer."
    },
    "IS : Contrôle de soi insuffisant": {
        179: "Je ne sais pas me discipliner pour atteindre mes objectifs.",
        180: "Je suis impulsif(ve), j'agis sur un coup de tête.",
        181: "Je ne finis pas ce que je commence.",
        182: "Je cède facilement à mes envies ou impulsions.",
        183: "Je m'ennuie très vite.",
        184: "Je ne supporte pas la frustration ou l'attente.",
        185: "J'agis sans réfléchir aux conséquences.",
        186: "J'ai du mal à me concentrer sur des tâches ennuyeuses.",
        187: "Je suis désorganisé(e) et chaotique.",
        188: "Je perds mon calme et je m'énerve facilement.",
        189: "Je procrastine beaucoup (je remets au lendemain).",
        190: "Je dépense sans compter ou je mange/bois trop.",
        191: "Je change souvent d'avis ou de direction.",
        192: "Je suis instable émotionnellement.",
        193: "Je vis au jour le jour sans plan d'avenir."
    },
    "AS : Recherche d'approbation": {
        194: "L'avis des autres est plus important que le mien.",
        195: "Je change mon comportement pour plaire aux autres.",
        196: "J'ai besoin qu'on m'admire et qu'on me complimente.",
        197: "Je veux être le centre de l'attention.",
        198: "Je ne supporte pas la critique ou la désapprobation.",
        199: "Je soigne mon image et mon apparence excessivement.",
        200: "Je veux être célèbre ou reconnu(e) socialement.",
        201: "Je ne sais pas prendre de décision sans l'approbation des autres.",
        202: "Je veux que tout le monde m'aime.",
        203: "Je suis très sensible au rejet ou à l'indifférence.",
        204: "Je cherche constamment les signes d'appréciation.",
        205: "Je me sens vide si on ne me remarque pas.",
        206: "Je flatte les gens pour être bien vu(e).",
        207: "La réussite sociale et le statut sont ma priorité."
    },
    "NP : Négativité / Pessimisme": {
        208: "Je vois toujours le mauvais côté des choses.",
        209: "Je m'attends toujours au pire.",
        210: "La vie est une vallée de larmes et de souffrance.",
        211: "Je ne suis pas chanceux(se), tout tourne mal pour moi.",
        212: "Je m'inquiète pour tout ce qui pourrait mal tourner.",
        213: "Je suis cynique et désabusé(e).",
        214: "Le bonheur ne dure jamais, quelque chose de mauvais va arriver.",
        215: "Je rumine mes erreurs et mes échecs passés.",
        216: "L'avenir est sombre et sans espoir.",
        217: "Je suis amer(ère) à cause de ce que j'ai vécu.",
        218: "Je me plains souvent de mes problèmes.",
        219: "Rien ne va jamais comme je veux.",
        220: "Je suis méfiant(e) quand tout va bien, ça cache quelque chose.",
        221: "Je décourage les autres avec mon pessimisme.",
        222: "Je vois les défauts partout, chez moi et chez les autres."
    },
    "PU : Punition": {
        223: "On doit payer pour ses fautes, il n'y a pas d'excuse.",
        224: "Je suis très dur(e) avec moi-même quand j'échoue ou je fais une erreur.",
        225: "Je ne pardonne pas facilement, je garde rancune.",
        226: "Je suis rancunier(ère) et je veux me venger.",
        227: "Il faut punir les gens qui font mal, c'est la justice.",
        228: "Je mérite d'être puni(e) quand je fais une erreur.",
        229: "Je ne supporte pas l'indulgence ou la faiblesse.",
        230: "Je me mets en colère contre mes erreurs et je m'insulte.",
        231: "La justice doit être sévère et inflexible.",
        232: "Je ne m'accorde pas le droit à l'erreur."
    }
}

# --- INTERFACE ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3050/3050525.png", width=80)
st.sidebar.title("Navigation")
mode = st.sidebar.radio("Aller vers :", ["Espace Patient", "Espace Thérapeute"])

# 1. ESPACE PATIENT
if mode == "Espace Patient":
    st.header("🌱 Questionnaire des Schémas (YSQ-L3)")
    st.markdown("---")
    
    st.info("""
    ### 📋 Instructions Importantes
    
    Ce questionnaire contient **232 affirmations**. Veuillez lire chacune d'elles et évaluer à quel point elle vous décrit bien.
    Il n'y a pas de bonne ou de mauvaise réponse. Soyez le plus sincère possible.
    
    **Échelle de réponse :**
    * **1** = Entièrement **FAUX** de moi
    * **2** = L'essentiel est **FAUX** de moi
    * **3** = Plutôt **VRAI** que faux
    * **4** = Modérément **VRAI** de moi
    * **5** = L'essentiel est **VRAI** de moi
    * **6** = Me décrit **PARFAITEMENT**
    """)
    
    with st.form("form_patient", clear_on_submit=False):
        c1, c2 = st.columns(2)
        nom = c1.text_input("Votre Nom et Prénom *")
        email = c2.text_input("Votre Email *")
        
        reponses = {}
        st.divider()
        
        # MODIFICATION : Titres neutres pour éviter le biais
        for i, (domaine, q_dict) in enumerate(YSQ_QUESTIONS.items()):
            st.markdown(f"#### 📝 Série {i+1}") 
            for q_num, q_text in q_dict.items():
                st.write(f"**{q_num}.** {q_text}")
                reponses[f"Q{q_num}"] = st.slider(
                    f"Rép. Q{q_num}", 1, 6, 1, key=f"q_{q_num}", label_visibility="collapsed"
                )
            st.markdown("---")
        
        submitted = st.form_submit_button("Envoyer mes résultats au thérapeute", type="primary")
        
        if submitted:
            if not nom or not email:
                st.error("⚠️ Oups ! Vous avez oublié de remplir votre **Nom** ou votre **Email**. Veuillez remonter en haut pour compléter ces champs, vos réponses ont été conservées.")
            else:
                with st.spinner("Envoi sécurisé en cours..."):
                    if save_patient_data(nom, email, reponses):
                        st.success("✅ Vos réponses ont été bien reçues et enregistrées ! Merci.")
                        st.balloons()

# 2. ESPACE THÉRAPEUTE
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
            
            # SÉLECTION ET SUPPRESSION
            c_select, c_action = st.columns([3, 1])
            with c_select:
                patient_options = {f"{row['nom']} ({row['created_at'][:16]})": row['id'] for index, row in df.iterrows()}
                selected_label = st.selectbox("Sélectionner un dossier à analyser ou supprimer :", list(patient_options.keys()))
                selected_id = patient_options[selected_label]
            
            with c_action:
                st.write("") 
                st.write("") 
                if st.button("🗑️ Supprimer ce dossier", type="primary"):
                    if delete_patient(selected_id):
                        st.success("Dossier supprimé.")
                        st.rerun()
            
            st.markdown("---")

            if st.button("📊 Générer le Bilan Complet pour ce patient"):
                patient_data = df[df["id"] == selected_id].iloc[0]
                reponses_dict = json.loads(patient_data["reponses_json"])
                
                # Calculs
                resultats = []
                active_schemas_codes = []
                
                for domaine, q_dict in YSQ_QUESTIONS.items():
                    code = domaine.split(" : ")[0]
                    nom_sch = domaine.split(" : ")[1]
                    scores = [reponses_dict.get(f"Q{k}", 1) for k in q_dict.keys()]
                    
                    if scores:
                        moy = sum(scores) / len(scores)
                        sev = len([x for x in scores if x >= 5])
                        pct = (sev / len(scores)) * 100
                        etoile = "⭐" if sev > 0 else ""
                        if etoile: active_schemas_codes.append(code)
                        
                        niveau = "🟢 Faible"
                        if moy > 3.5: niveau = "🔴 IMPORTANT"
                        elif moy >= 2.5: niveau = "🟡 Moyen"
                        
                        resultats.append({
                            "Code": code,
                            "Schéma": f"{nom_sch} {etoile}",
                            "Moyenne": round(moy, 2),
                            "% Sévérité": f"{round(pct, 1)}%",
                            "Niveau": niveau
                        })
                
                df_res = pd.DataFrame(resultats)
                
                # Affichage
                c1, c2 = st.columns(2)
                with c1:
                    st.table(df_res)
                with c2:
                    fig_radar = px.line_polar(df_res, r='Moyenne', theta='Code', line_close=True, range_r=[0,6])
                    fig_radar.update_traces(fill='toself', line_color='red')
                    st.plotly_chart(fig_radar)
                    
                    fig_bar = px.bar(df_res, x='Code', y='Moyenne', range_y=[0,6], color='Moyenne', color_continuous_scale='Reds')
                    st.plotly_chart(fig_bar)

                # Export Word
                def generate_word_expert():
                    doc = Document()
                    doc.add_heading(f"Bilan Psychométrique : {patient_data['nom']}", 0)
                    doc.add_paragraph(f"Date : {patient_data['created_at'][:10]}")
                    
                    doc.add_heading('1. Visualisation', level=1)
                    try:
                        img_radar = fig_radar.to_image(format="png", engine="kaleido")
                        doc.add_picture(BytesIO(img_radar), width=Inches(5))
                        img_bar = fig_bar.to_image(format="png", engine="kaleido")
                        doc.add_picture(BytesIO(img_bar), width=Inches(5))
                    except Exception as e:
                        doc.add_paragraph(f"[Graphiques non disponibles : {str(e)}]")

                    doc.add_heading('2. Scores Détaillés', level=1)
                    table = doc.add_table(rows=1, cols=4)
                    table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text = "Code"
                    hdr[1].text = "Schéma"
                    hdr[2].text = "Moyenne"
                    hdr[3].text = "Niveau"
                    for _, r in df_res.iterrows():
                        row = table.add_row().cells
                        row[0].text = str(r["Code"])
                        row[1].text = str(r["Schéma"])
                        row[2].text = str(r["Moyenne"])
                        row[3].text = str(r["Niveau"])

                    doc.add_heading('3. Analyse Clinique & Pastorale', level=1)
                    doc.add_paragraph("Ce rapport intègre une approche de relation d'aide chrétienne, reliant les schémas émotionnels aux vérités bibliques pour la restauration de l'identité.")
                    
                    if active_schemas_codes:
                        for code in active_schemas_codes:
                            info = INTERPRETATIONS_EXPERTES.get(code)
                            if info:
                                p = doc.add_paragraph()
                                p.add_run(f"\n{info['titre']} ({code})").bold = True
                                p.add_run(f" - Moyenne: {df_res.loc[df_res['Code'] == code, 'Moyenne'].values[0]}")
                                doc.add_paragraph(f"Diagnostic : {info['desc']}")
                                p_bib = doc.add_paragraph()
                                p_bib.add_run("Piste Pastorale : ").bold = True
                                p_bib.add_run(info['biblique'])
                                p_verset = doc.add_paragraph()
                                p_verset.add_run(f"Verset clé : {info['verset']}").italic = True
                                doc.add_paragraph("-" * 30)
                    else:
                        doc.add_paragraph("Aucun schéma significatif détecté (scores < 5).")

                    out = BytesIO()
                    doc.save(out)
                    return out.getvalue()

                st.download_button(
                    "📥 Télécharger le Rapport Expert (Word)",
                    generate_word_expert(),
                    f"Bilan_Pastoral_{patient_data['nom']}.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    elif pwd_input:
        st.error("Mot de passe incorrect.")
