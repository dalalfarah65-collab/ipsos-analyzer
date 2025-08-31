# utils.py

class LLMBridge:
    """Classe fictive pour communication avec un modèle de langage."""
    def __init__(self):
        pass

    def ask(self, prompt: str) -> str:
        return f"Réponse fictive pour : {prompt}"


def extract_questions(text: str):
    """Extrait les phrases qui se terminent par un point d'interrogation."""
    return [q.strip() for q in text.split("\n") if q.strip().endswith("?")]


def read_docx_text(file):
    """Lit le texte brut d'un fichier DOCX."""
    try:
        import docx
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception:
        return "Erreur lors de la lecture du fichier DOCX."


def analyze_group(text: str):
    """Analyse très simple du texte (fictive)."""
    return {
        "summary": "Résumé fictif.",
        "keywords": ["mot1", "mot2", "mot3"],
        "word_count": len(text.split())
    }


def detect_language(text: str):
    """Détecte la langue du texte (fictive)."""
    if any(word in text.lower() for word in ["le", "la", "et"]):
        return "français"
    elif any(word in text.lower() for word in ["the", "and", "is"]):
        return "anglais"
    else:
        return "inconnu"
