#!/usr/bin/env python3
# coding: utf-8
"""
Read .docx, .txt, or .pdf files and analyze if they were written by AI
Author: LincolnLandForensics
Version: 0.0.3
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>> #

import os
import re
import sys
import math
import hashlib
import argparse

import unicodedata

from collections import Counter
from typing import Dict, Any, List
from docx import Document
from PyPDF2 import PdfReader
from docx.opc import coreprops

# colors
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
from colorama import Fore, Back, Style
print(Back.BLACK)
color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Main           >>>>>>>>>>>>>>>>>>>>>>>>>> #

def main():
    parser = argparse.ArgumentParser(description="Analyze DOCX or PDF for AI authorship indicators")
    parser.add_argument('-I', '--input', help='Input DOCX, TXT or PDF file', required=False)
    parser.add_argument('-a', '--analyzeAI', help='Analyze the input file for AI-like characteristics', action='store_true')
    parser.add_argument('-O', '--output', help='Write cleaned text to this file', required=False)

    args = parser.parse_args()

    input_file = args.input if args.input else "sample.docx"


    # Replace last '.' with '_cleaned.'
    parts = input_file.rsplit('.', 1)
    output_file2 = f"{parts[0]}_cleaned.{parts[1]}" if len(parts) == 2 else input_file + "_cleaned"

    output_file = args.output if args.output else output_file2

    if args.analyzeAI:
        if os.path.exists(input_file):
            msg_blurb_square(f"Reading {input_file}", color_green)

            result = analyze_document(input_file, highlight=True, clean=True)

            if "highlighted_text" in result:
                print("\n--- Highlighted Text Preview ---\n")
                print(result["highlighted_text"][:500])

            # if args.output and "cleaned_text" in result:
            if "cleaned_text" in result:
            
                with open(output_file, "w", encoding="utf-8") as f:
                # with open(args.output, "w", encoding="utf-8") as f:
                    # f.write(result["cleaned_text"])
                    write_to_docx(output_file, result["cleaned_text"])

                msg_blurb_square(f"Cleaned output written to {args.output}", color_green)

            print_pretty_summary(result)

        else:
            msg_blurb_square(f"{input_file} does not exist", color_red)
            sys.exit(1)
    else:
        usage()

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Core Logic     >>>>>>>>>>>>>>>>>>>>>>>>>> #

def analyze_document(q: str, highlight: bool = False, clean: bool = False) -> Dict[str, Any]:
    if q.lower().endswith(".docx") or q.lower().endswith(".doc"):
        text = extract_text_from_docx(q)
        metadata = extract_docx_metadata(q)

    elif q.lower().endswith(".pdf"):
        text = extract_text_from_pdf(q)
        metadata = extract_pdf_metadata(q)

    elif q.lower().endswith((
    ".txt", ".text", ".asc", ".nfo", ".log", ".csv", ".md",
    ".xml", ".json", ".ini", ".yaml", ".yml", ".tex", ".cfg", ".dat"
    )):

        with open(q, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
        metadata = {
            "filename": q,
            "length": len(text),
            "lines": text.count("\n") + 1,
            "encoding": "utf-8"
        }

    else:
        raise ValueError("Unsupported file type. Only DOCX, DOC, PDF, or TXT files are allowed.")

    hidden_chars = detect_hidden_characters(text)
    score = compute_ai_likelihood(text, hidden_chars)
    # filehash = file_hash(q)

    result = {
        "file": q,
        "hidden_characters": hidden_chars,
        "ai_likelihood_score": score,
        # "file_hash": filehash,
        "metadata": metadata,
    }

    if highlight:
        result["highlighted_text"] = highlight_hidden_characters(text, hidden_chars)
    if clean:
        result["cleaned_text"] = clean_text(text, hidden_chars)

    return result

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Utilities      >>>>>>>>>>>>>>>>>>>>>>>>>> #

def entropy_score(text: str) -> float:
    """
    Estimates Shannon entropy of the input text.
    Higher entropy suggests more randomness; lower entropy suggests repetition or predictability.
    """
    if not text:
        return 0.0

    # Normalize text: lowercase and remove whitespace
    normalized = text.lower().replace(" ", "")
    
    # Count character frequencies
    freq = Counter(normalized)
    total = sum(freq.values())

    # Compute entropy
    entropy = -sum((count / total) * math.log2(count / total) for count in freq.values())
    # print(f' temp entopy score = {entropy}')    # temp
    
    print(f'{color_blue}Entropy score: {round(entropy, 4)}{color_reset}')
    return round(entropy, 4)


def extract_text_from_docx(input_file: str) -> str:
    ext = os.path.splitext(input_file)[1].lower()

    if ext == ".docx":
        try:
            doc = Document(input_file)
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        except Exception as e:
            return (f"{color_red}Error reading .docx file: {e}{color_reset}")
            try:
                import textract
                text = textract.process(input_file).decode("utf-8")
                return "\n".join([line.strip() for line in text.splitlines() if line.strip()])
            except Exception as e:
                return f"Error reading .doc file: {e}"


    elif ext == ".doc":
        print(f'I havent written the .doc reader yet')
        # try:
            # import textract
            # text = textract.process(input_file).decode("utf-8")
            # return "\n".join([line.strip() for line in text.splitlines() if line.strip()])
        # except Exception as e:
            # return f"Error reading .doc file: {e}"

    else:
        return "Unsupported file format. Please provide a .docx or .pdf"


def extract_text_from_docxOLD(input_file: str) -> str:
    doc = Document(input_file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def extract_text_from_pdf(input_file: str) -> str:
    reader = PdfReader(input_file)
    return "\n".join([page.extract_text() or "" for page in reader.pages])

def extract_docx_metadata(file_path: str) -> Dict[str, Any]:
    doc = Document(file_path)
    props = doc.core_properties

    metadata = {
        "title": props.title,
        "subject": props.subject,
        "author": props.author,
        "category": props.category,
        "comments": props.comments,
        "keywords": props.keywords,
        "content_status": props.content_status,
        "identifier": props.identifier,
        "language": props.language,
        "revision": props.revision,
        "version": props.version,
        "created": str(props.created) if props.created else None,
        "modified": str(props.modified) if props.modified else None,
        "last_printed": str(props.last_printed) if props.last_printed else None,
        "last_modified_by": props.last_modified_by,
    }

    # Filter out empty or None values
    return {k: v for k, v in metadata.items() if v not in [None, ""]}

def extract_pdf_metadata(file_path: str) -> Dict[str, Any]:
    reader = PdfReader(file_path)
    return dict(reader.metadata or {})

def file_hash(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()

def detect_hidden_characters(text: str) -> Dict[str, Any]:
    suspicious_chars = {
        "ZERO_WIDTH_SPACE": "\u200b",
        "ZERO_WIDTH_NON_JOINER": "\u200c",
        "ZERO_WIDTH_JOINER": "\u200d",
        "NON_BREAKING_SPACE": "\u00a0",
        "SOFT_HYPHEN": "\u00ad",
        "LEFT_TO_RIGHT_MARK": "\u200e",
        "RIGHT_TO_LEFT_MARK": "\u200f",
    }

    findings = {}
    for name, char in suspicious_chars.items():
        count = text.count(char)
        if count > 0:
            findings[name] = {"char": char, "count": count}

    unusual = [ch for ch in text if ord(ch) > 126 and not unicodedata.category(ch).startswith("Z")]
    if unusual:
        findings["UNUSUAL_UNICODE"] = {
            "char": None,
            "count": len(unusual),
            "examples": Counter(unusual).most_common(5)
        }

    return findings

def highlight_hidden_characters(text: str, findings: Dict[str, Any]) -> str:
    highlighted = text
    for name, data in findings.items():
        if name == "UNUSUAL_UNICODE":
            continue
        char = data["char"]
        highlighted = highlighted.replace(char, f"[{name}]")
    return highlighted

def clean_text(text: str, findings: Dict[str, Any]) -> str:
    cleaned = (
        text.replace("’", "'")
             .replace("—", ", ")
             .replace("“", "\"")
             .replace("”", "\"")
             .replace("️", "")
             .replace("️", "")
             .replace("⚠", "")
             .replace("🧠", "")
             .replace("☠", "")
             .replace("🧭", "")
)
    
    for name, data in findings.items():
        if name == "UNUSUAL_UNICODE":
            continue
        char = data["char"]
        cleaned = cleaned.replace(char, "")
    return cleaned

def compute_ai_likelihood(text: str, hidden_char_findings: Dict[str, Any]) -> float:
    score = 0.0

    if hidden_char_findings:
        score += min(30, sum(v["count"] for v in hidden_char_findings.values() if "count" in v) * 5)

    sentences = re.split(r"[.!?]", text)
    sentences = [s.strip() for s in sentences if s.strip()]
    avg_len = sum(len(s.split()) for s in sentences) / max(1, len(sentences))
    if avg_len > 25:
        score += 20

    words = re.findall(r"\w+", text.lower())
    freq = Counter(words)
    repeated = [w for w, c in freq.items() if c > 10 and len(w) > 6]
    if repeated:
        score += 15

    # Passive voice detection (naive)
    passive_hits = len(re.findall(r"\b(is|was|were|been|being|be)\b\s+\w+ed\b", text))
    if passive_hits > 5:
        score += 10

    # Placeholder for entropy/burstiness
    score += entropy_score(text)

    return min(100.0, score)

def generate_summary(result: Dict[str, Any]) -> str:
    lines = [f"# AI Authorship Analysis for `{result['file']}`"]
    lines.append(f"**AI-Likelihood Score**: `{result['ai_likelihood_score']}`")
    lines.append(f"**File Hash**: `{result['file_hash']}`")
    lines.append("## Metadata:")
    for k, v in result["metadata"].items():
        lines.append(f"- **{k}**: {v}")
    lines.append("## Hidden Characters Detected:")
    for k, v in result["hidden_characters"].items():
        lines.append(f"- **{k}**: {v['count']}")
    return "\n".join(lines)

def msg_blurb_square(msg, color):
    border = f"+{'-' * (len(msg) + 2)}+"
    print(f"{color}{border}\n| {msg} |\n{border}{color_reset}")

def print_pretty_summary(result: Dict[str, Any]):
    print("\n" + "=" * 60)
    print("📄 AI Authorship Analysis Summary")
    print("=" * 60)

    print(f"\n📊 AI-Likelihood Score:\n  {result['ai_likelihood_score']} (Moderate indicators based on hidden characters, entropy and repetition)")

    print("\n🕵️ Hidden Characters Detected:")
    print(f"{'Type':<25} {'Unicode':<10} {'Count':<6} {'Examples / Notes'}")
    print("-" * 60)
    for k, v in result["hidden_characters"].items():
        char = repr(v["char"]) if v["char"] else "—"
        count = v["count"]
        if k == "UNUSUAL_UNICODE":
            examples = ", ".join([f"{repr(ch)}×{n}" for ch, n in v["examples"]])
        else:
            examples = ""
        print(f"{k:<25} {char:<10} {str(count):<6} {examples}")

    print("\n🧾 Document Metadata:")
    print(f"{'Field':<20} {'Value'}")
    print("-" * 40)
    for k, v in result["metadata"].items():
        print(f"{k:<20} {v}")
    # print(f"\n🔐 File Hash (SHA256):\n  {result['file_hash']}")

        
    print("\n" + "=" * 60 + "\n")

def write_to_docx(output_path: str, text: str) -> None:
    doc = Document()
    for line in text.splitlines():
        if line.strip():  # Avoid empty lines
            doc.add_paragraph(line.strip())
    doc.save(output_path)

def usage():
    print(f"Usage: {sys.argv[0]} -a [-I sample.docx] [-O sample_updated.docx]")
    print("Example:")
    print(f"    {sys.argv[0]} -a")
    print(f"    {sys.argv[0]} -a -I sample.docx -O sample_cleaned.docx")
    print(f"    {sys.argv[0]} -a -I sample_AI_written.txt  -O sample_AI_Cleaned.docx")
    print(f"    {sys.argv[0]} -a -I sample_AI_written.txt  -a -O sample_AI_Cleaned.docx")


if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
0.1.2 - original version

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>


'''

'''

