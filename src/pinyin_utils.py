"""Generate pinyin variants for matching Chinese names against filenames/text.

Handles forms like: zhangsan, zhang_san, ZhangSan, sanzhang (given-first)
"""
from typing import List, Set

from pypinyin import lazy_pinyin


def get_pinyin_variants(chinese_name: str) -> List[str]:
    """Return a list of lowercase pinyin variants used for substring matching.

    For "李卓然" (li/zhuo/ran), produces variants like:
      lizhuoran, li_zhuoran, li-zhuoran, zhuoranli, zhuoran_li, zhuoran-li,
      lizr, lizhr (initials-style)

    For 2-character names, surname-first and given-first are generated.
    For names with >=2 syllables in the given name, the given name is also
    handled both as one block and with internal separators.
    """
    if not chinese_name:
        return []

    # lazy_pinyin: "李卓然" -> ["li", "zhuo", "ran"]
    syllables = [s.lower() for s in lazy_pinyin(chinese_name)]
    if not syllables:
        return []

    if len(syllables) == 1:
        # Single-syllable name (rare), just return it
        return [syllables[0]]

    # # Treat first syllable as surname, the rest as given name.
    # # (Compound surnames like "欧阳" are uncommon and would still match
    # # via the given-name half; we don't try to detect them here.)
    # surname = syllables[0]
    # given_syllables = syllables[1:]
    # given_joined = "".join(given_syllables)

    # variants: Set[str] = set()

    # # Surname-first orderings
    # variants.add(surname + given_joined)              # lizhuoran
    # variants.add(surname + "_" + given_joined)        # li_zhuoran
    # variants.add(surname + "-" + given_joined)        # li-zhuoran
    # if len(given_syllables) > 1:
    #     variants.add(surname + "_" + "_".join(given_syllables))  # li_zhuo_ran
    #     variants.add(surname + "-" + "-".join(given_syllables))

    # # Given-name-first orderings (some students write name-then-surname)
    # variants.add(given_joined + surname)              # zhuoranli
    # variants.add(given_joined + "_" + surname)        # zhuoran_li
    # variants.add(given_joined + "-" + surname)        # zhuoran-li
    # if len(given_syllables) > 1:
    #     variants.add("_".join(given_syllables) + "_" + surname)
    #     variants.add("-".join(given_syllables) + "-" + surname)

    # return sorted(variants, key=len, reverse=True)  # longer first to prefer specific matches
    
    variants: Set[str] = set()

    # Separators commonly used in filenames:
    #   ZhuoranLi, Zhuoran Li, Zhuoran_Li, Zhuoran-Li
    separators = ["", " ", "_", "-"]

    def add_joined(parts: List[str]) -> None:
        """Add variants by joining name parts with common separators."""
        parts = [p for p in parts if p]
        if not parts:
            return
        for sep in separators:
            variants.add(sep.join(parts))

    def add_name_orders(surname_parts: List[str], given_parts: List[str]) -> None:
        """Add surname-first, given-first, and given-only variants."""
        surname_joined = "".join(surname_parts)      # li
        given_joined = "".join(given_parts)          # zhuoran

        # Basic blocks:
        #   li + zhuoran
        #   zhuoran + li
        add_joined([surname_joined, given_joined])   # lizhuoran, li zhuoran, li_zhuoran, li-zhuoran
        add_joined([given_joined, surname_joined])   # zhuoranli, zhuoran li, zhuoran_li, zhuoran-li

        # Fully split given name:
        #   li zhuo ran
        #   zhuo ran li
        if len(given_parts) > 1:
            add_joined([surname_joined] + given_parts)   # li zhuo ran, li_zhuo_ran, etc.
            add_joined(given_parts + [surname_joined])   # zhuo ran li, zhuo_ran_li, etc.

        # Given name only:
        #   zhuoran
        # Useful when filename is only "Report Zhuoran_0427.pdf".
        if len(given_joined) >= 4:
            variants.add(given_joined)

        # Given name split:
        #   zhuo ran, zhuo_ran, zhuo-ran
        # Do not add individual syllables like "zhuo" or "ran";
        # they are too ambiguous.
        if len(given_parts) > 1:
            add_joined(given_parts)

        # Surname only is usually too ambiguous, especially "li".
        # Also matcher.py skips variants shorter than 4 chars anyway,
        # so we intentionally do not add surname-only variants.

    # Normal case: first syllable is surname, rest is given name.
    # 李卓然 -> surname = li, given = zhuo ran
    if len(syllables) >= 2:
        add_name_orders(syllables[:1], syllables[1:])

    # Optional compound-surname support:
    # 欧阳娜娜 -> also try surname = ou yang, given = na na
    # This does not affect 李卓然, but helps names with 2-character surnames.
    if len(syllables) >= 3:
        add_name_orders(syllables[:2], syllables[2:])

    return sorted(variants, key=len, reverse=True)


def normalise_for_match(s: str) -> str:
    """Lowercase + strip non-alphanumeric (except for keeping content recognisable).

    Used when matching variants against filenames or PDF text. We keep underscores
    and hyphens because variants include them.
    """
    return s.lower()
