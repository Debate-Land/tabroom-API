from anthropic import Anthropic
import json
from typing import List, TypedDict, Any
import re
import os
from docx import Document
from docx.oxml.ns import qn
from pdf2docx import Converter
import tempfile
from bs4 import BeautifulSoup
from shared.const import API_BASE
import html
import time
import requests

client = Anthropic(api_key=os.environ['ANTHROPIC_KEY'])

class Card(TypedDict):
    author: str
    start: str
    end: str

class CardWithHTML(TypedDict):
    author: str
    url: str
    html_content: str

class MetadataEntry(TypedDict):
    author: str
    url: str

def extract_formatted_text(file_path: str) -> str:
    """Extract formatted text from Word document, preserving highlighting and structure."""
    doc = Document(file_path if file_path.endswith(".docx") else convert_pdf_to_docx(file_path))
    html_content: List[str] = []

    for paragraph in doc.paragraphs:
        para_html: List[str] = []
        for run in paragraph.runs:
            text = html.escape(run.text)
            if run.bold:
                text = f'<strong>{text}</strong>'
            if run.italic:
                text = f'<em>{text}</em>'
            if run.underline:
                text = f'<u>{text}</u>'
            if run.font.highlight_color:
                highlight_color = run.font.highlight_color
                text = f'<mark style="background-color: {highlight_color};">{text}</mark>'
            if run.font.color.rgb:
                color = run.font.color.rgb
                text = f'<span style="color: rgb({color[0]},{color[1]},{color[2]});">{text}</span>'

            # Check for hyperlinks
            if run._element.rPr is not None:
                if run._element.rPr.xpath("./w:rStyle[@w:val='Hyperlink']"):
                    for link in run._element.xpath(".//w:hyperlink"):
                        if link.get(qn("r:id")):
                            href = doc.part.rels[link.get(qn("r:id"))].target_ref
                            text = f'<a href="{href}">{text}</a>'

            para_html.append(text)

        html_content.append(f'<p>{"".join(para_html)}</p>')

    return '\n'.join(html_content)

def convert_pdf_to_docx(pdf_path: str) -> str:
    """Convert PDF to DOCX and return the path to the new file"""
    # Create a temporary file with .docx extension
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
        docx_path = tmp_file.name

    # Convert PDF to DOCX
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

    return docx_path

def flexible_match(target: str, text: str, threshold: float = 0.8) -> bool:
    """Perform a flexible match between target and text."""
    target_words = target.lower().split()
    text_words = text.lower().split()

    matches = sum(1 for word in target_words if word in text_words)
    return matches / len(target_words) >= threshold

def validate_and_clean_cards(cards_data: Any) -> List[Card]:
    """Clean and validate the extracted cards."""
    cleaned_cards: List[Card] = []

    if not isinstance(cards_data, list):
        return []

    for card in cards_data:
        if isinstance(card, dict):
            author = card.get('author', '').strip()
            start = card.get('start', '').strip()
            end = card.get('end', '').strip()

            if author and start and end:
                cleaned_cards.append({
                    'author': author,
                    'start': start,
                    'end': end
                })

    return cleaned_cards

def identify_card_boundaries(html_content: str) -> List[Card]:
    """Use Anthropic API to identify card boundaries in HTML content."""
    soup = BeautifulSoup(html_content, 'html.parser')
    text = soup.get_text()

    print(f"Extracted text length: {len(text)} characters")
    print("Sending request to Anthropic API...")

    prompt = f"""Human: Analyze the following text and identify each distinct evidence card. For each card, extract:
1. The author's name and year (e.g., "Massey '17")
2. The exact start of the card content WHICH INCLUDES THE AUTHOR NAME
3. The exact end of the card content

Rules:
- Each card should have a distinct author and year.
- Ensure you capture ALL cards in the text.
- The start should be the beginning of the card, including the author name.
- The end should be the last few words of the card content.
- MAKE SURE TO INCLUDE ALL CARDS IN THE DOCUMENT AND SOME DOCUMENTS CAN BE 100s OF CARDS LONG.
- MAKE SURE THE JSON IS exactly formatting NEVER HAS TEXT OUTSIDE THE JSON

Format your response as a JSON array of objects, like this:
[
    {{
        "author": "Author Name 'YY",
        "start": "Author Name 'YY [exact start of card content]",
        "end": "[exact end of card content]"
    }},
    // ... (all cards should be listed here)
]

Here's the text to analyze:

{text}
"""

    response_content = client.messages.create(
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": prompt + "\n\nText to analyze:\n" + text
        }],
        model="claude-3-5-sonnet-20241022",
        stream=False
    ).content[0].text

    match = re.search(r'\[\s*\{.*?\}\s*\]', response_content, re.DOTALL)
    if match:
        json_str = match.group(0)
        try:
            cards_data = json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"JSON decoding error: {e}")
            return []
        cleaned_cards = validate_and_clean_cards(cards_data)
        if cleaned_cards:
            return cleaned_cards

    print("Invalid or empty response from Anthropic. Retrying...")
    time.sleep(1)  # Wait a second before retrying

    return []

def extract_card_html(html_content: str, cards: List[Card]) -> List[CardWithHTML]:
    """Extract HTML content for each card."""
    cards_with_html: List[CardWithHTML] = []
    soup = BeautifulSoup(html_content, 'html.parser')

    for card in cards:
        try:
            start_text = html.unescape(card['start'])
            end_text = html.unescape(card['end'])

            # Find the start element
            start_elem = None
            for elem in soup.find_all(['p', 'div', 'span']):
                if flexible_match(start_text, elem.get_text()):
                    start_elem = elem
                    break

            if not start_elem:
                print(f"Couldn't find start for card: {card['author']}")
                continue

            # Find the end element
            end_elem = None
            current = start_elem
            while current:
                if flexible_match(end_text, current.get_text()):
                    end_elem = current
                    break
                current = current.next_element

            if not end_elem:
                print(f"Couldn't find end for card: {card['author']}")
                continue

            # Extract all HTML content between start and end elements
            full_content = []
            current = start_elem
            while current:
                full_content.append(str(current))
                if current == end_elem:
                    break
                current = current.next_sibling

            full_html = ''.join(full_content)

            # Extract URL
            url_match = re.search(r'https?://\S+', full_html)
            url = url_match.group(0) if url_match else ''

            cards_with_html.append({
                'author': card['author'],
                'url': url,
                'html_content': full_html
            })

        except Exception as e:
            print(f"Error processing card: {card['author']}")
            print(f"Error details: {str(e)}")
            continue

    return cards_with_html

def clean_card_content(html_content: str, author: str) -> str:
    soup = BeautifulSoup(html_content, 'html.parser')

    # Find the first occurrence of the author name
    author_elem = soup.find(string=re.compile(re.escape(author)))
    if author_elem and author_elem.parent:
        # Remove any preceding siblings
        for elem in list(author_elem.parent.previous_siblings):
            elem.decompose()

        # Remove duplication in the first element
        first_elem = author_elem.parent
        text = first_elem.get_text()
        parts = text.split(author)
        if len(parts) > 2:
            new_text = author + ''.join(parts[2:])
            first_elem.string = new_text

    return str(soup)

def process_document(file_path: str, output_dir: str) -> None:
    """Process document and save results with HTML content."""
    os.makedirs(output_dir, exist_ok=True)

    # Extract formatted text
    html_content = extract_formatted_text(file_path)
    print(f"Extracted HTML content length: {len(html_content)} characters")

    # Identify card boundaries
    cards = identify_card_boundaries(html_content)
    print(f"Identified {len(cards)} cards")

    for i, card in enumerate(cards):
        print(f"Card {i+1}:")
        print(f"  Author: {card['author']}")
        print(f"  Start: {card['start']}")
        print(f"  End: {card['end']}")

    # Extract HTML content for each card
    cards_with_html = extract_card_html(html_content, cards)
    print(f"Processed document. Found {len(cards_with_html)} cards with HTML content.")

    # Write cards to individual HTML files
    for i, card in enumerate(cards_with_html):
        cleaned_content = clean_card_content(card['html_content'], card['author'])
        file_name = f"{card['author'].replace(' ', '_')}_{i+1}.html"
        output_path = os.path.join(output_dir, file_name)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"<html><body>{cleaned_content}</body></html>")
        print(f"Created HTML file: {file_name}")
        res = requests.post(f"{API_BASE}/cards", json={
            'html': f"<html><body>{cleaned_content}</body></html>",
            'text': BeautifulSoup(cleaned_content, 'html.parser').get_text(" "),
            'author': card['author'],
            'url': card['url']
        })

    # Write metadata to JSON file
    metadata: List[MetadataEntry] = [{'author': card['author'], 'url': card['url']} for card in cards_with_html]
    json_path = os.path.join(output_dir, 'metadata.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, indent=2)
    print(f"Created metadata file: metadata.json")

def process_directory(input_dir: str, output_dir: str) -> None:
    """Process all .docx and .pdf files in the directory and its subdirectories."""
    os.makedirs(output_dir, exist_ok=True)

    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith(('.docx', '.pdf')):
                file_path = os.path.join(root, file)
                print(f"Processing file: {file_path}")
                process_document(file_path, output_dir)
