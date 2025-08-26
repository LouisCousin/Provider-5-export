from __future__ import annotations

"""Module d'analyse de documents pour l'import."""

from typing import Any, Dict, Optional, Tuple

import docx
import fitz  # PyMuPDF


def analyser_docx(file_stream) -> Tuple[str, Optional[Dict[str, Any]]]:
    """Tente d'extraire le texte et un template de style simplifié d'un DOCX.

    Retourne ``(texte_brut, styles)`` où ``styles`` est ``None`` si l'extraction
    de style échoue.
    """
    try:
        file_stream.seek(0)
        document = docx.Document(file_stream)
        full_text = "\n".join(para.text for para in document.paragraphs)

        styles: Optional[Dict[str, Any]] = None
        try:
            first_para = document.paragraphs[0] if document.paragraphs else None
            first_run = first_para.runs[0] if first_para and first_para.runs else None
            if first_run:
                font_name = first_run.font.name or "Calibri"
                font_size = first_run.font.size.pt if first_run.font.size else 11
                styles = {
                    "prompt": {
                        "font_name": font_name,
                        "font_size": font_size,
                        "is_bold": bool(first_run.bold),
                    },
                    "response": {
                        "font_name": font_name,
                        "font_size": font_size,
                        "is_bold": False,
                    },
                }
        except Exception:
            styles = None

        return full_text, styles
    except Exception:
        try:
            file_stream.seek(0)
            document = docx.Document(file_stream)
            full_text = "\n".join(para.text for para in document.paragraphs)
        except Exception:
            full_text = ""
        return full_text, None


def analyser_pdf(file_stream) -> Tuple[str, None]:
    """Extrait le contenu textuel brut d'un PDF."""
    file_stream.seek(0)
    with fitz.open(stream=file_stream.read(), filetype="pdf") as doc:
        full_text = "".join(page.get_text() for page in doc)
    return full_text, None


def analyser_document(fichier) -> Tuple[str, Optional[Dict[str, Any]]]:
    """Analyse un fichier importé et choisit la méthode appropriée."""
    filename = fichier.name.lower()
    if filename.endswith(".docx"):
        return analyser_docx(fichier)
    if filename.endswith(".pdf"):
        return analyser_pdf(fichier)
    return "", None
