"""Generate pinyin variants from manually provided pinyin aliases."""

import re
from typing import List, Set


def get_pinyin_alias_variants(alias: str) -> List[str]:
    """Return lowercase pinyin variants from a manually provided pinyin alias.

    Example:
    "li zhuoran" -> lizhuoran, li zhuoran, li_zhuoran, li-zhuoran,
                    zhuoranli, zhuoran li, zhuoran_li, zhuoran-li
    """
    alias = alias.strip().lower()
    if not alias:
        return []

    parts = [p for p in re.split(r"[\s_\-]+", alias) if p]
    if not parts:
        return []

    if len(parts) == 1:
        return [parts[0]]

    variants: Set[str] = set()
    separators = ["", " ", "_", "-"]

    def add_joined(items: List[str]) -> None:
        items = [x for x in items if x]
        if not items:
            return
        for sep in separators:
            variants.add(sep.join(items))

    surname_parts = parts[:1]
    given_parts = parts[1:]

    surname_joined = "".join(surname_parts)
    given_joined = "".join(given_parts)

    # surname-first: li zhuoran / lizhuoran / li_zhuoran / li-zhuoran
    add_joined([surname_joined, given_joined])

    # given-first: zhuoran li / zhuoranli / zhuoran_li / zhuoran-li
    add_joined([given_joined, surname_joined])

    # fully split given name: li zhuo ran / zhuo ran li
    if len(given_parts) > 1:
        add_joined([surname_joined] + given_parts)
        add_joined(given_parts + [surname_joined])
        add_joined(given_parts)

    # given name only: zhuoran
    if len(given_joined) >= 4:
        variants.add(given_joined)

    return sorted(variants, key=len, reverse=True)


def normalise_for_match(s: str) -> str:
    """Lowercase + strip non-alphanumeric (except for keeping content recognisable).

    Used when matching variants against filenames or PDF text. We keep underscores
    and hyphens because variants include them.
    """
    return s.lower()
