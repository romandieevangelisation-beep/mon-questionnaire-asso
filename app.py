import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
import json
import re
from supabase import create_client

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Espace Clinique - YSQ-L3 Officiel", layout="wide")

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

# --- DÉFINITIONS EXPERTES & PASTORALES ---
INTERPRETATIONS_EXPERTES = {
    "ED": {"titre": "Carence Affective", "desc": "Sentiment que vos besoins de sécurité, d'affection et d'écoute ne seront jamais comblés.", "verset": "Psaume 27:10", "biblique": "Dieu se révèle comme le Père parfait qui 'recueille'. Guérir, c'est oser croire que vous êtes digne de Son soin personnel."},
    "AB": {"titre": "Abandon / Instabilité", "desc": "Peur viscérale que les relations soient fragiles et que vous finissiez seul(e).", "verset": "Hébreux 13:5", "biblique": "L'Alliance avec Dieu est immuable. Il a promis : 'Je ne te délaisserai point'. Votre sécurité est ancrée en Lui."},
    "MA": {"titre": "Méfiance / Abus", "desc": "Attente que les autres vont intentionnellement vous blesser ou vous trahir.", "verset": "Psaume 56:4-5", "biblique": "Dieu est votre refuge sûr. La guérison permet de troquer la suspicion (peur) contre le discernement (sagesse) sous Sa garde."},
    "SI": {"titre": "Isolement Social", "desc": "Sentiment de ne pas appartenir, d'être différent, 'à part' du reste de l'humanité.", "verset": "Éphésiens 2:19", "biblique": "En Christ, vous n'êtes plus étranger. Vous avez une place légitime et vitale dans la famille de Dieu, le Corps de Christ."},
    "DS": {"titre": "Imperfection / Honte", "desc": "Croyance profonde d'être intérieurement défectueux ou indigne d'amour.", "verset": "Sophonie 3:17", "biblique": "La honte dit 'je suis une erreur', Dieu dit 'tu es ma créature merveilleuse'. Votre valeur est scellée par Son amour, pas vos manques."},
    "FA": {"titre": "Échec", "desc": "Sentiment d'être incompétent, croyance que l'échec est inévitable.", "verset": "2 Corinthiens 12:9", "biblique": "La puissance de Dieu s'accomplit dans la faiblesse. Votre succès est défini par votre fidélité, pas par la performance sociale."},
    "DI": {"titre": "Dépendance / Incompétence", "desc": "Sentiment d'être incapable de gérer le quotidien sans l'aide d'autrui.", "verset": "Philippiens 4:13", "biblique": "Dieu vous a donné un esprit de force. L'Esprit Saint vous rend capable de marcher dans vos responsabilités avec sagesse."},
    "VU": {"titre": "Vulnérabilité au danger", "desc": "Peur constante qu'une catastrophe soit imminente. Hyper-vigilance.", "verset": "Psaume 91:4", "biblique": "L'antidote à la peur est la confiance en la Souveraineté de Dieu. Il est votre abri contre toute terreur."},
    "EU": {"titre": "Fusion / Personnalité Atrophiée", "desc": "Manque d'individualité, sentiment de ne pas exister sans l'autre.", "verset": "Galates 1:10", "biblique": "Vous êtes une création unique. Christ vous appelle à une liberté qui permet d'aimer sans se perdre dans l'autre."},
    "SB": {"titre": "Assujettissement", "desc": "Soumission aux autres par peur des conséquences. Vous taisez vos besoins.", "verset": "Galates 5:1", "biblique": "Christ vous a libéré pour l'amour, pas pour l'esclavage de la peur. Vous servez Dieu en premier, avec des limites saines."},
    "SS": {"titre": "Abnégation", "desc": "Sacrifice excessif de vos besoins (Syndrome du Sauveur).", "verset": "Matthieu 22:39", "biblique": "Aimer son prochain 'comme soi-même'. Honorer votre propre vie, c'est honorer le temple du Saint-Esprit."},
    "EI": {"titre": "Inhibition Émotionnelle", "desc": "Verrouillage des émotions et de la spontanéité par peur de la honte.", "verset": "Psaume 62:9", "biblique": "Répandez votre cœur devant Lui. La vérité émotionnelle est une étape vers la liberté que Jésus nous offre."},
    "US": {"titre": "Exigences Élevées", "desc": "Perfectionnisme, règles rigides et impossibilité de se détendre.", "verset": "Matthieu 11:28", "biblique": "Dieu n'est pas un tyran mais un Père qui donne le repos. La grâce remplace le perfectionnisme étouffant."},
    "ET": {"titre": "Droits Personnels / Grandeur", "desc": "Sentiment de supériorité, d'avoir des droits spéciaux.", "verset": "Philippiens 2:3", "biblique": "La vraie grandeur réside dans l'humilité de Christ. Servir les autres est le chemin de la joie véritable."},
    "IS": {"titre": "Contrôle de soi insuffisant", "desc": "Difficulté à tolérer la frustration, impulsivité.", "verset": "Galates 5:22", "biblique": "La maîtrise de soi est un fruit de l'Esprit. Il fortifie votre volonté pour des choix qui honorent Dieu."},
    "AS": {"titre": "Recherche d'approbation", "desc": "Votre estime dépend entièrement de la validation des autres.", "verset": "1 Thessaloniciens 2:4", "biblique": "Vous vivez pour plaire à Dieu qui sonde les cœurs. Son approbation est votre seul fondement sûr."},
    "NP": {"titre": "Négativité / Pessimisme", "desc": "Focalisation sur le négatif en minimisant le positif.", "verset": "Philippiens 4:8", "biblique": "La foi entraîne le regard à voir la grâce. Fixez vos pensées sur ce qui est digne de louange."},
    "PU": {"titre": "Punition", "desc": "Croyance que les erreurs méritent une punition sévère. Rancune.", "verset": "Romains 8:1", "biblique": "La condamnation a été portée par Christ. Pratiquer la miséricorde envers soi et les autres est la voie de la paix."}
}

# --- LES 232 QUESTIONS RÉVISÉES (YSQ-L3) ---
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
        85: "Je me sens dépassé(e) par les responsabilités de la vie.",
        86: "Je cherche toujours quelqu'un pour me dire quoi faire.",
        87: "Je me sens vulnérable dès que je suis seul(e).",
        88: "Je ne sais pas résoudre les problèmes pratiques courants.",
        89: "Je panique quand je sors seul(e).",
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
        155: "Je suis perfectionniste, tout doit être impeccable.",
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
        184: "Je vis au jour le jour sans plan financier.",
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
    
    options_reponse = [1, 2, 3, 4, 5, 6]

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
                        options=options_reponse,
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

            if st.button("📊 Générer le Bilan Complet"):
                patient_data = df[df["id"] == selected_id].iloc[0]
                reponses_dict = json.loads(patient_data["reponses_json"])
                
                # Calculs
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
                        "Code": code,
                        "Schéma": f"{nom_sch} {etoile}",
                        "Moyenne": round(moy, 2),
                        "% Sévérité": f"{round(pct, 1)}%",
                        "Niveau": niv
                    })
                
                df_res = pd.DataFrame(resultats)
                
                c1, c2 = st.columns(2)
                with c1: st.table(df_res)
                with c2:
                    fig_radar = px.line_polar(df_res, r='Moyenne', theta='Code', line_close=True, range_r=[0,6])
                    fig_radar.update_traces(fill='toself', line_color='blue')
                    st.plotly_chart(fig_radar)
                    
                    fig_bar = px.bar(df_res, x='Code', y='Moyenne', range_y=[0,6], color='Moyenne', color_continuous_scale='Blues')
                    st.plotly_chart(fig_bar)

                def generate_word_expert():
                    doc = Document()
                    doc.add_heading(f"Bilan Psychométrique : {patient_data['nom']}", 0)
                    
                    doc.add_heading('1. Visualisation', level=1)
                    try:
                        img_radar = fig_radar.to_image(format="png", engine="kaleido")
                        doc.add_picture(BytesIO(img_radar), width=Inches(5))
                    except: doc.add_paragraph("[Graphique indisponible]")

                    doc.add_heading('2. Analyse par Domaine', level=1)
                    if active_codes:
                        for domain_name, domain_info in YOUNG_DOMAINS_INFO.items():
                            match = [c for c in domain_info["codes"] if c in active_codes]
                            if match:
                                doc.add_heading(domain_name, level=2)
                                doc.add_paragraph(domain_info["besoin"]).italic = True
                                for c in match:
                                    info = INTERPRETATIONS_EXPERTES[c]
                                    doc.add_paragraph(f"🔹 {info['titre']}").bold = True
                                    doc.add_paragraph(f"Diagnostic : {info['desc']}")
                                    doc.add_paragraph(f"Piste Pastorale : {info['biblique']}").bold = True
                                    doc.add_paragraph(f"Verset : {info['verset']}").italic = True
                    
                    out = BytesIO()
                    doc.save(out)
                    return out.getvalue()

                st.download_button("📥 Télécharger le Rapport Expert", generate_word_expert(), f"Bilan_{patient_data['nom']}.docx")

    elif pwd_input:
        st.error("Mot de passe incorrect.")

# --- FOOTER ---
st.markdown("---")
with st.expander("🔒 Confidentialité et Protection des Données"):
    st.markdown("""
    **Engagement de confidentialité :**
    * Les données recueillies sont strictement confidentielles.
    * Seul votre thérapeute a accès aux résultats détaillés.
    * Vous pouvez demander la suppression de vos données à tout moment.
    """)
    st.caption("Application YSQ-L3 révisée. pour usage ne relation d'aide chrétienne © 2026 la Barque.")
