from __future__ import annotations
import re
from pathlib import Path

_STOPWORDS: set[str] = set()

_INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*]+')
_SPLIT_RE = re.compile(r'[\s._,\-–—]+')
_DIGIT_RE = re.compile(r'\d')


def _tokenize(filename: str) -> list[str]:
    base = Path(str(filename)).stem
    text = base.lower()
    text = _INVALID_CHARS_RE.sub(' ', text)
    text = re.sub(r'[\(\)\[\]\{\}]+', ' ', text)
    tokens_raw = [t for t in _SPLIT_RE.split(text) if t]
    out = []
    for tok in tokens_raw:
        tok = tok.strip().strip('.')
        if not tok:
            continue
        if tok in _STOPWORDS:
            continue
        if len(tok) < 2:
            continue
        out.append(tok)
    return out


def derive_common_label(filename_a: str, filename_b: str) -> str:
    tokens_a = _tokenize(filename_a)
    tokens_b_list = _tokenize(filename_b)
    tokens_b = set(tokens_b_list)

    common = [t for t in tokens_a if t in tokens_b]

    if not common:
        fuzzy = []
        for a in tokens_a:
            best = ''
            for b in tokens_b_list:
                if len(a) >= 4 and (a in b or b in a):
                    cand = a if len(a) >= len(b) else b
                    if len(cand) > len(best):
                        best = cand
            if best:
                fuzzy.append(best)
        common = fuzzy

    if not common:
        return ''

    label = ' '.join(common)
    label = re.sub(r'\s+', ' ', label).strip().rstrip(' .')
    return label


def filename_match_score(filename_a: str, filename_b: str) -> int:
    tokens_a = _tokenize(filename_a)
    tokens_b_list = _tokenize(filename_b)
    tokens_b = set(tokens_b_list)

    direct = [t for t in tokens_a if t in tokens_b]
    if direct:
        return sum(len(t) for t in direct)

    score = 0
    for a in tokens_a:
        best = 0
        for b in tokens_b_list:
            if len(a) >= 4 and len(b) >= 4 and (a in b or b in a):
                best = max(best, min(len(a), len(b)))
        score += best
    return score


def make_sverka_filename(filename_a: str, filename_b: str, prefix: str = 'Сверка') -> str:
    label = derive_common_label(filename_a, filename_b)
    if not label:
        label = 'Без названия'
    label = _INVALID_CHARS_RE.sub(' ', label)
    label = re.sub(r'\s+', ' ', label).strip().rstrip(' .')
    if not label:
        label = 'Без названия'
    return f'{prefix}_{label}.xlsx'
