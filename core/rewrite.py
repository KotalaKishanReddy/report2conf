"""Rewrite engine: API LLM with deterministic extractive fallback.

Lossless means every section, number, table value, and citation survives
in some form. Eloquent means short academic sentences. The fallback keeps
sentences containing digits first, then the lead sentences of each paragraph.
"""
import os
import re

WORDS_PER_PAGE = 850
TABLE_WORD_EQUIV = 120


def target_words(max_pages: int, n_tables: int, n_images: int) -> int:
    budget = max_pages * WORDS_PER_PAGE
    budget -= (n_tables * TABLE_WORD_EQUIV + n_images * 60)
    return max(400, budget)


def _sentences(text: str):
    parts = re.split(r"(?<=[.!?])\s+", text.strip())
    return [s.strip() for s in parts if s.strip()]


def extractive_compress(text: str, max_words: int) -> str:
    if len(text.split()) <= max_words:
        return text
    sents = _sentences(text)
    if not sents:
        return text[: max_words * 7]
    with_digits = [s for s in sents if re.search(r"\d", s)]
    rest = [s for s in sents if s not in with_digits]
    # interleave: lead sentence, digit sentences, then rest in order
    ordered = ([sents[0]] if sents else []) + with_digits + rest
    seen, out, count = set(), [], 0
    for s in ordered:
        if s in seen:
            continue
        seen.add(s)
        w = len(s.split())
        if count + w > max_words and out:
            break
        out.append(s)
        count += w
    return " ".join(out)


def llm_rewrite(text: str, instruction: str, model: str | None = None) -> str | None:
    """Return rewritten text via OpenAI-compatible API, or None if unavailable."""
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key or not text.strip():
        return None
    try:
        from openai import OpenAI

        client = OpenAI(
            api_key=api_key,
            base_url=os.environ.get("OPENAI_BASE_URL") or None,
        )
        resp = client.chat.completions.create(
            model=model or os.environ.get("REPORT2CONF_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": instruction},
                {"role": "user", "content": text[:12000]},
            ],
            temperature=0.3,
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return None


SYSTEM_INSTRUCTION = (
    "Rewrite the following technical report section in concise, eloquent IEEE "
    "conference prose. Preserve every fact, number, unit, method name, and "
    "citation. Do not invent results. Keep the same section heading ideas. "
    "Shorten by removing repetition, not information."
)


def rewrite_sections(sections: list, word_budget: int, use_llm: bool = True,
                     model: str | None = None) -> tuple[list, dict]:
    """Rewrite each section to a proportional share of word_budget."""
    total = sum(len(" ".join(s["paras"]).split()) for s in sections) or 1
    out, coverage = [], {"llm_used": 0, "fallback_used": 0, "sections": []}
    for s in sections:
        body = "\n".join(s["paras"])
        share = max(80, int(word_budget * len(body.split()) / total))
        new_body = None
        if use_llm:
            new_body = llm_rewrite(f"Section: {s['heading']}\n\n{body}",
                                   SYSTEM_INSTRUCTION, model)
        if new_body:
            coverage["llm_used"] += 1
        else:
            new_body = extractive_compress(body, share)
            coverage["fallback_used"] += 1
        out.append({"heading": s["heading"], "level": s["level"], "text": new_body,
                    "orig_words": len(body.split()),
                    "new_words": len(new_body.split())})
        coverage["sections"].append(
            {"heading": s["heading"], "orig_words": len(body.split()),
             "new_words": len(new_body.split())})
    return out, coverage
