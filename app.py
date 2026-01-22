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
        response = supabase.table("patients_ysq").select("*").order("created_at", desc=True).execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        return pd.DataFrame()

# --- DÉFINITIONS EXPERTES & BIBLIQUES ---
INTERPRETATIONS_EXPERTES = {
    "ED": {
        "titre": "Carence Affective",
        "desc": "Ce schéma réfère au sentiment que vos besoins émotionnels de sécurité, d'affection et d'empathie ne seront jamais comblés par les autres. Vous pouvez vous sentir invisible ou émotionnellement seul(e).",
        "verset": "Psaume 27:10",
        "biblique": "Même si les humains faillissent, Dieu est le Père parfait qui comble votre vide affectif."
    },
    "AB": {
        "titre": "Abandon / Instabilité",
        "desc": "La perception que vos proches sont instables et qu'ils finiront inévitablement par vous quitter (par décès, départ ou rejet). Cela crée une anxiété chronique dans les relations.",
        "verset": "Hébreux 13:5",
        "biblique": "L'alliance avec Dieu est éternelle. Il a promis : 'Je ne te délaisserai point'."
    },
    "MA": {
        "titre": "Méfiance / Abus",
        "desc": "L'attente que les autres vont vous blesser, vous abuser, vous humilier ou vous mentir. Vous percevez le monde comme dangereux et les gens comme malveillants.",
        "verset": "Psaume 56:4",
        "biblique": "Dieu est votre refuge et votre bouclier. En Lui, la confiance peut être restaurée sans crainte."
    },
    "SI": {
        "titre": "Isolement Social",
        "desc": "Le sentiment d'être différent des autres, de ne pas appartenir à la 'famille humaine' ou à un groupe. Sentiment de solitude sociale chronique.",
        "verset": "Éphésiens 2:19",
        "biblique": "Vous n'êtes plus des étrangers, mais concitoyens des saints et membres de la famille de Dieu."
    },
    "DS": {
        "titre": "Imperfection / Honte",
        "desc": "Le sentiment d'être intérieurement défectueux, mauvais, ou sans valeur. Crainte que si les autres voient qui vous êtes vraiment, ils vous rejetteront.",
        "verset": "Psaume 139:14",
        "biblique": "Vous êtes une créature merveilleuse. Votre valeur dépend de l'amour de Dieu, pas de vos performances."
    },
    "FA": {
        "titre": "Échec",
        "desc": "La croyance que vous avez échoué, que vous échouerez inévitablement, ou que vous êtes fondamentalement moins compétent que vos pairs.",
        "verset": "Philippiens 4:13",
        "biblique": "Le succès selon Dieu n'est pas la réussite sociale, mais la fidélité. Sa force s'accomplit dans votre faiblesse."
    },
    "DI": {
        "titre": "Dépendance / Incompétence",
        "desc": "La conviction d'être incapable de gérer vos responsabilités quotidiennes sans une aide considérable d'autrui. Sentiment d'impuissance.",
        "verset": "2 Timothée 1:7",
        "biblique": "Dieu ne nous a pas donné un esprit de timidité, mais de force, d'amour et de sagesse."
    },
    "VU": {
        "titre": "Vulnérabilité au danger",
        "desc": "Peur exagérée qu'une catastrophe (médicale, financière, criminelle) puisse survenir à tout moment et que vous ne pourrez pas y faire face.",
        "verset": "Psaume 91:4",
        "biblique": "Il vous couvrira de ses plumes. Sa fidélité est une armure et un bouclier contre la peur."
    },
    "EU": {
        "titre": "Fusion / Personnalité Atrophiée",
        "desc": "Attachement émotionnel excessif à des proches, au détriment de votre propre individualité. Sentiment de ne pas pouvoir exister ou être heureux sans l'autre.",
        "verset": "Galates 5:1",
        "biblique": "C'est pour la liberté que Christ nous a affranchis. Vous avez une identité propre en Lui."
    },
    "SB": {
        "titre": "Assujettissement",
        "desc": "Soumission excessive au contrôle des autres pour éviter la colère, les représailles ou l'abandon. Vous supprimez vos propres besoins.",
        "verset": "Actes 5:29",
        "biblique": "Il faut obéir à Dieu plutôt qu'aux hommes. Vous êtes serviteur de Dieu, pas esclave des hommes."
    },
    "SS": {
        "titre": "Abnégation",
        "desc": "Concentration excessive sur les besoins des autres au détriment des vôtres, souvent pour éviter de se sentir coupable. Le syndrome du 'Sauveur'.",
        "verset": "Matthieu 22:39",
        "biblique": "Tu aimeras ton prochain 'comme toi-même'. Prendre soin de soi est nécessaire pour bien servir."
    },
    "EI": {
        "titre": "Inhibition Émotionnelle",
        "desc": "Inhibition excessive des actions, des sentiments ou de la communication spontanée, souvent pour éviter la désapprobation ou la honte.",
        "verset": "Jean 8:36",
        "biblique": "Si le Fils vous affranchit, vous serez réellement libres. Libres de ressentir et d'être vrais."
    },
    "US": {
        "titre": "Exigences Élevées",
        "desc": "La croyance que l'on doit s'efforcer d'atteindre des normes internes très élevées de comportement et de performance, souvent pour éviter la critique.",
        "verset": "Matthieu 11:28",
        "biblique": "Venez à moi, vous tous qui êtes fatigués et chargés, et je vous donnerai du repos."
    },
    "ET": {
        "titre": "Droits Personnels / Grandeur",
        "desc": "La croyance que vous êtes supérieur aux autres, que vous avez des droits spéciaux, ou que vous n'avez pas à suivre les règles de réciprocité sociale.",
        "verset": "Philippiens 2:3",
        "biblique": "Que l'humilité vous fasse regarder les autres comme étant au-dessus de vous-mêmes."
    },
    "IS": {
        "titre": "Contrôle de soi insuffisant",
        "desc": "Difficulté persistante à exercer un contrôle de soi suffisant et une tolérance à la frustration pour atteindre ses objectifs personnels.",
        "verset": "Galates 5:22",
        "biblique": "Le fruit de l'Esprit est... la maîtrise de soi."
    },
    "AS": {
        "titre": "Recherche d'approbation",
        "desc": "Recherche excessive de l'attention, de l'estime et de l'approbation des autres, au détriment du développement d'un sentiment de soi sûr et authentique.",
        "verset": "1 Thessaloniciens 2:4",
        "biblique": "Nous parlons, non pour plaire aux hommes, mais pour plaire à Dieu qui sonde nos cœurs."
    },
    "NP": {
        "titre": "Négativité / Pessimisme",
        "desc": "Focalisation constante sur les aspects négatifs de la vie (douleur, mort, perte, conflit) tout en minimisant les aspects positifs.",
        "verset": "Philippiens 4:8",
        "biblique": "Que tout ce qui est vrai, honorable, juste... soit l'objet de vos pensées."
    },
    "PU": {
        "titre": "Punition",
        "desc": "La croyance que les gens (y compris soi-même) doivent être sévèrement punis pour leurs erreurs. Tendance à la colère, à l'intolérance et à l'impatience.",
        "verset": "Romains 8:1",
        "biblique": "Il n'y a donc maintenant aucune condamnation pour ceux qui sont en Jésus-Christ."
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
    
    with st.form("form_patient", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nom = c1.text_input("Votre Nom et Prénom")
        email = c2.text_input("Votre Email")
        
        reponses = {}
        st.divider()
        
        for domaine, q_dict in YSQ_QUESTIONS.items():
            st.subheader(f"🔹 {domaine}")
            for q_num, q_text in q_dict.items():
                st.write(f"**{q_num}.** {q_text}")
                reponses[f"Q{q_num}"] = st.slider(
                    f"Rép. Q{q_num}", 1, 6, 1, key=f"q_{q_num}", label_visibility="collapsed"
                )
            st.markdown("---")
        
        submitted = st.form_submit_button("Envoyer mes résultats au thérapeute", type="primary")
        
        if submitted:
            if not nom:
                st.warning("⚠️ Merci d'indiquer votre nom avant d'envoyer.")
            else:
                with st.spinner("Envoi sécurisé en cours..."):
                    if save_patient_data(nom, email, reponses):
                        st.success("✅ Vos réponses ont été bien reçues ! Merci.")
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
            st.dataframe(df[["created_at", "nom", "email"]], use_container_width=True)
            st.divider()
            patient_select = st.selectbox("Sélectionner un patient :", df["nom"].unique())
            
            if st.button("📊 Générer le Bilan Complet"):
                patient_data = df[df["nom"] == patient_select].iloc[0]
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

                # Export Word avec Graphiques
                def generate_word_expert():
                    doc = Document()
                    doc.add_heading(f"Bilan Psychométrique : {patient_data['nom']}", 0)
                    doc.add_paragraph(f"Date : {patient_data['created_at'][:10]}")
                    
                    # 1. Graphiques (Tentative d'insertion)
                    doc.add_heading('1. Visualisation', level=1)
                    try:
                        # Radar
                        img_radar = fig_radar.to_image(format="png")
                        doc.add_picture(BytesIO(img_radar), width=Inches(5))
                        # Barres
                        img_bar = fig_bar.to_image(format="png")
                        doc.add_picture(BytesIO(img_bar), width=Inches(5))
                    except:
                        doc.add_paragraph("[Note: Installez 'kaleido' dans requirements.txt pour voir les graphiques ici]")

                    # 2. Tableau
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

                    # 3. Analyse Expert & Spirituelle
                    doc.add_heading('3. Analyse Clinique & Spirituelle', level=1)
                    if active_schemas_codes:
                        for code in active_schemas_codes:
                            info = INTERPRETATIONS_EXPERTES.get(code)
                            if info:
                                p = doc.add_paragraph()
                                runner = p.add_run(f"{info['titre']} ({code})")
                                runner.bold = True
                                runner.font.size = Inches(0.16)
                                
                                doc.add_paragraph(f"Analyse : {info['desc']}")
                                doc.add_paragraph(f"Promesse Biblique : '{info['biblique']}' ({info['verset']})").italic = True
                                doc.add_paragraph("-" * 30)
                    else:
                        doc.add_paragraph("Aucun schéma significatif détecté.")

                    out = BytesIO()
                    doc.save(out)
                    return out.getvalue()

                st.download_button(
                    "📥 Télécharger le Rapport Expert (Word)",
                    generate_word_expert(),
                    f"Bilan_Expert_{patient_data['nom']}.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    elif pwd_input:
        st.error("Mot de passe incorrect.")
