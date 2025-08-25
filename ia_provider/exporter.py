"""Outils d'exportation des résultats de batch en document DOCX."""

from __future__ import annotations

import io
import json
from dataclasses import asdict, is_dataclass
from typing import Any, Dict, List

import markdown as md
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


def _apply_style(run, style: Dict[str, Any]) -> None:
    """Applique un style basique à un run du document."""

    if font_name := style.get("font_name"):
        run.font.name = font_name
        # Nécessaire pour une compatibilité complète avec Word
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if size := style.get("font_size"):
        run.font.size = Pt(size)
    if color := style.get("font_color_rgb"):
        run.font.color.rgb = RGBColor(*color)
    run.bold = style.get("is_bold", False)
    run.italic = style.get("is_italic", False)


def _add_inline(paragraph, node, style: Dict[str, Any]) -> None:
    """Ajoute récursivement les noeuds inline à un paragraphe."""

    if node.text:
        run = paragraph.add_run(node.text)
        _apply_style(run, style)

    for child in list(node):
        if child.tag in {"strong", "b"}:
            run = paragraph.add_run(child.text or "")
            _apply_style(run, {**style, "is_bold": True})
        elif child.tag in {"em", "i"}:
            run = paragraph.add_run(child.text or "")
            _apply_style(run, {**style, "is_italic": True})
        elif child.tag == "code":
            run = paragraph.add_run(child.text or "")
            _apply_style(run, style)
            run.font.name = "Consolas"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
        elif child.tag == "a":
            # Ajout simplifié : texte du lien suivi de l'URL entre parenthèses
            text = child.text or ""
            href = child.get("href", "")
            run = paragraph.add_run(text)
            _apply_style(run, style)
            if href and href != text:
                tail_run = paragraph.add_run(f" ({href})")
                _apply_style(tail_run, style)
        else:
            _add_inline(paragraph, child, style)

        if child.tail:
            tail = paragraph.add_run(child.tail)
            _apply_style(tail, style)


def _process_element(doc, elem, style: Dict[str, Any], list_style: str | None = None) -> None:
    """Traite les éléments HTML convertis depuis le Markdown."""

    tag = elem.tag
    if tag in {"p", "li"}:
        paragraph = doc.add_paragraph(style=list_style) if list_style else doc.add_paragraph()
        _add_inline(paragraph, elem, style)
        for child in list(elem):
            if child.tag in {"ul", "ol"}:
                _process_element(
                    doc,
                    child,
                    style,
                    "List Bullet" if child.tag == "ul" else "List Number",
                )
    elif tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
        level = int(tag[1])
        paragraph = doc.add_heading(level=level)
        _add_inline(paragraph, elem, style)
    elif tag == "ul":
        for li in elem.findall("li"):
            _process_element(doc, li, style, "List Bullet")
    elif tag == "ol":
        for li in elem.findall("li"):
            _process_element(doc, li, style, "List Number")
    elif tag == "pre":
        code_text = "".join(elem.itertext()).strip()
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(code_text)
        _apply_style(run, style)
        run.font.name = "Consolas"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
    elif tag == "table":
        rows = elem.findall("tr")
        if rows:
            cols = len(rows[0].findall("th")) or len(rows[0].findall("td"))
            table = doc.add_table(rows=len(rows), cols=cols)
            for r_idx, row in enumerate(rows):
                cells = row.findall("th") or row.findall("td")
                for c_idx, cell in enumerate(cells):
                    paragraph = table.cell(r_idx, c_idx).paragraphs[0]
                    _add_inline(paragraph, cell, style)
    else:
        if elem.text:
            paragraph = doc.add_paragraph()
            _add_inline(paragraph, elem, style)


def _add_markdown(doc: Document, text: str, style: Dict[str, Any]) -> None:
    """Convertit un texte Markdown et l'ajoute au document."""

    if not text:
        return

    md_converter = md.Markdown(extensions=["fenced_code", "tables"])
    html = md_converter.convert(text)

    # Encapsuler dans un élément racine pour le parsing XML
    from xml.etree import ElementTree as ET

    root = ET.fromstring(f"<root>{html}</root>")
    for elem in list(root):
        _process_element(doc, elem, style)


def generer_export_docx(resultats: List[Any], styles: Dict[str, Dict[str, Any]]) -> io.BytesIO:
    """Génère un document DOCX à partir d'une liste de résultats de batch.

    Chaque résultat doit contenir au minimum les champs ``status``,
    ``prompt_text`` et ``clean_response`` (ou ``response``).
    """

    document = Document()

    succeeded: List[Dict[str, Any]] = []
    failed: List[Dict[str, Any]] = []

    for res in resultats:
        data = asdict(res) if is_dataclass(res) else dict(res)
        if data.get("status") == "succeeded":
            succeeded.append(data)
        else:
            failed.append(data)

    # Section principale : prompts et réponses
    for item in succeeded:
        prompt_text = item.get("prompt_text", "")
        response_text = item.get("clean_response") or item.get("response", "")

        para = document.add_paragraph()
        run = para.add_run(prompt_text)
        _apply_style(run, styles.get("prompt", {}))

        _add_markdown(document, response_text, styles.get("response", {}))
        document.add_paragraph()

    # Section annexe pour les erreurs
    if failed:
        document.add_page_break()
        document.add_heading("Annexe - Requêtes échouées", level=1)
        for item in failed:
            title = item.get("prompt_text") or item.get("custom_id")
            document.add_paragraph(title, style="List Bullet")
            err = item.get("error")
            if isinstance(err, dict):
                err = json.dumps(err, ensure_ascii=False, indent=2)
            document.add_paragraph(str(err))

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output

