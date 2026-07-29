import re
import os
import pandas as pd
from typing import Dict, Tuple, Optional

try:
    import openai
except Exception:
    openai = None

# ---------- Regex patterns ----------
RE_POSTAL_5 = re.compile(r'\b\d{5}\b')
RE_POSTAL_4 = re.compile(r'\b\d{4}\b')
RE_FIRST_NUMBER_TOKEN = re.compile(r'\b\d+[A-Za-zА-Яа-я\-\/]*\b')

RE_BLOCK = re.compile(r'\b(?:бл\.?|бл|block)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)', re.IGNORECASE)
RE_ENTRANCE = re.compile(r'\b(?:вх\.?|вх|entr\.?|entrance)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)', re.IGNORECASE)
RE_FLOOR = re.compile(r'\b(?:ет\.?|ет|floor)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)', re.IGNORECASE)
RE_APT = re.compile(r'\b(?:ап\.?|ап|апартамент|apt|apartment)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)', re.IGNORECASE)

# ---------- Parsing helpers ----------
def split_city_and_rest(combined: str) -> Tuple[str, str]:
    if not isinstance(combined, str):
        return '', ''
    s = combined.strip()
    if '(' in s:
        city = s.split('(', 1)[0].strip()
        rest = s[len(city):].strip()
        return city, rest
    if ',' in s:
        city = s.split(',', 1)[0].strip()
        rest = s[len(city)+1:].strip()
        return city, rest
    return s, ''

def extract_street_from_combined(combined: str, district_present: bool, ast_name: Optional[str]=None) -> str:
    if not isinstance(combined, str):
        return ''
    if ast_name:
        return ast_name.strip()
    parts = [p.strip() for p in combined.split(',')]
    if not district_present:
        if len(parts) >= 3:
            return parts[-1]
        if len(parts) == 2:
            return parts[1]
    for p in parts:
        if re.search(r'\b(ул\.?|улица|street|str|бул\.?|булевард|бул)\b', p, re.IGNORECASE):
            return p
    return combined

def extract_postal_code(combined: str) -> Optional[str]:
    if not isinstance(combined, str):
        return None
    m = RE_POSTAL_5.search(combined)
    if m:
        return m.group(0)
    m = RE_POSTAL_4.search(combined)
    if m:
        return m.group(0)
    return None

def extract_first_number_token(combined: str) -> Optional[str]:
    if not isinstance(combined, str):
        return None
    m = RE_FIRST_NUMBER_TOKEN.search(combined)
    if m:
        return m.group(0)
    return None

def extract_block_entr_floor_apt(combined: str) -> Dict[str, Optional[str]]:
    return {
        'block': (lambda m: m.group(1).strip() if m else None)(RE_BLOCK.search(combined)),
        'entrance': (lambda m: m.group(1).strip() if m else None)(RE_ENTRANCE.search(combined)),
        'floor': (lambda m: m.group(1).strip() if m else None)(RE_FLOOR.search(combined)),
        'apartment': (lambda m: m.group(1).strip() if m else None)(RE_APT.search(combined)),
    }

# ---------- AI verification ----------
def ai_verify_parts(parts: Dict[str, str], original: str, openai_api_key: Optional[str]=None, model: str='gpt-4o-mini'):
    if openai is None:
        return {'ai_available': False, 'verified': None, 'suggestions': None, 'confidence': None}

    if not openai_api_key:
        openai_api_key = os.getenv('OPENAI_API_KEY')

    if not openai_api_key:
        return {'ai_available': False, 'verified': None, 'suggestions': None, 'confidence': None}

    openai.api_key = openai_api_key

    prompt = (
        "You are an address parsing assistant. Given the original address and parsed parts, "
        "verify whether each parsed part is correct. If incorrect, provide corrected value. "
        "Return ONLY a JSON object with keys: verified (true/false), confidence (0..1), corrections (object). "
        f"Original: '''{original}'''. Parsed: {parts}."
    )

    try:
        resp = openai.ChatCompletion.create(
            model=model,
            messages=[
                {"role": "system", "content": "You output only JSON."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.0,
            max_tokens=400
        )
        text = resp['choices'][0]['message']['content'].strip()
        import json
        parsed = json.loads(text)
        return {
            'ai_available': True,
            'verified': parsed.get('verified'),
            'suggestions': parsed.get('corrections'),
            'confidence': parsed.get('confidence')
        }
    except Exception as e:
        return {'ai_available': True, 'error': str(e), 'verified': None, 'suggestions': None, 'confidence': None}

# ---------- Main parsing ----------
def parse_address_row(combined: str, ast_name: Optional[str]=None, district_present: bool=False):
    city, rest = split_city_and_rest(combined or '')
    street = extract_street_from_combined(combined or '', district_present, ast_name)
    postal = extract_postal_code(combined or '')
    street_number = extract_first_number_token(combined or '')
    extras = extract_block_entr_floor_apt(combined or '')

    return {
        'full_address': combined,
        'city': city or None,
        'street': street or None,
        'postal_code': postal,
        'street_number': street_number,
        'block': extras.get('block'),
        'entrance': extras.get('entrance'),
        'floor': extras.get('floor'),
        'apartment': extras.get('apartment'),
    }

def process_excel(input_path: str, output_path: str, sheet_name: str, ai_verify: bool=False, openai_api_key: Optional[str]=None):
    df = pd.read_excel(input_path, sheet_name=sheet_name, engine='openpyxl')

    if df.shape[1] < 6:
        raise ValueError("Input file must have at least 6 columns; address expected in column F.")

    address_col = df.columns[5]
    results = []

    for idx, row in df.iterrows():
        combined = row[address_col]
        parsed = parse_address_row(combined)

        # Build full street name including number
        full_street = parsed['street']
        if parsed['street_number']:
            full_street = f"{parsed['street']} {parsed['street_number']}"

        ai_result = None
        if ai_verify:
            ai_result = ai_verify_parts(
                {
                    'city': parsed['city'],
                    'street': parsed['street'],
                    'postal_code': parsed['postal_code'],
                    'street_number': parsed['street_number'],
                    'block': parsed['block'],
                    'entrance': parsed['entrance'],
                    'floor': parsed['floor'],
                    'apartment': parsed['apartment']
                },
                original=str(combined or ''),
                openai_api_key=openai_api_key
            )

        out = {
            'City of': parsed['city'],
            'Region of': None,
            'District of': None,
            'Street of': full_street,
            'Postal Code': parsed['postal_code'],
            'Street Number': parsed['street_number'],
            'Block': parsed['block'],
            'Entrance': parsed['entrance'],
            'Floor': parsed['floor'],
            'Apartment': parsed['apartment'],
            'AI Verified': ai_result.get('verified') if ai_result else None,
            'AI Suggestion': ai_result.get('suggestions') if ai_result else None,
            'AI Confidence': ai_result.get('confidence') if ai_result else None
        }

        results.append(out)

    parsed_df = pd.DataFrame(results)
    out_df = pd.concat([df.reset_index(drop=True), parsed_df.reset_index(drop=True)], axis=1)
    out_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"Processed {len(df)} rows. Output written to {output_path}")

# ---------- Defaults ----------
DEFAULT_INPUT = r"\\virgo\Work\OpControl\2016\Vlad\Address Check\Address check_SQL_A1.xlsx"
DEFAULT_OUTPUT = r"\\virgo\Work\OpControl\2016\Vlad\Address Check\Address_check_parsed.xlsx"
DEFAULT_SHEET = "1_Addresses_All_-6"

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Parse addresses and optionally verify with AI.")
    parser.add_argument('--input', default=DEFAULT_INPUT)
    parser.add_argument('--output', default=DEFAULT_OUTPUT)
    parser.add_argument('--sheet', default=DEFAULT_SHEET)
    parser.add_argument('--ai', action='store_true')
    parser.add_argument('--key', default=None)
    args = parser.parse_args()

    process_excel(
        args.input,
        args.output,
        sheet_name=args.sheet,
        ai_verify=args.ai,
        openai_api_key=args.key
    )
