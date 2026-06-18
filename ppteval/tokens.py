"""Token counting via tiktoken (free/OSS, offline once the encoding is cached).

Used for the `out_tokens` metric and `tokens_per_slide`. We deliberately treat
the count as a cross-converter *proxy* for LLM cost, not an exact per-model figure.
"""
from __future__ import annotations

from functools import lru_cache

from .config import TOKENIZER


@lru_cache(maxsize=4)
def _encoder(name: str):
    import tiktoken

    try:
        return tiktoken.get_encoding(name)
    except Exception:
        # Fallback that ships with tiktoken and needs no network.
        return tiktoken.get_encoding("cl100k_base")


def count_tokens(text: str, name: str | None = None) -> int:
    if not text:
        return 0
    enc = _encoder(name or TOKENIZER)
    # disallowed_special=() => never raise on literal "<|endoftext|>"-style text.
    return len(enc.encode(text, disallowed_special=()))
