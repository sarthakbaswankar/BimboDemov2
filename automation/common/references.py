"""Reference naming (NF-, SB INV-, SP/OP, NAC, DISC-) and automatic -1/-2 suffixing."""
import re


def norm_ref(ref):
    """Comparable form of a reference: upper case, single spaces, no leading zeros
    in numbers, so 'NF-026866321582' and 'NF-26866321582' are the same reference."""
    s = re.sub(r"\s+", " ", str(ref).upper().strip())
    if s.isdigit():
        return s.lstrip("0") or "0"
    m = re.match(r"^(NF-|SB INV-)0*(\d+)(.*)$", s)
    if m:
        return m.group(1) + m.group(2) + m.group(3)
    return s


class RefRegistry:
    """Hands out references that exist neither in Oracle nor earlier in this run."""

    def __init__(self, existing_refs, style="number", work_date=None):
        self.existing = set(existing_refs)
        self.created = set()
        self.style = style
        self.work_date = work_date

    def exists(self, ref):
        return norm_ref(ref) in self.existing

    def _taken(self, ref):
        n = norm_ref(ref)
        return n in self.existing or n in self.created

    def make(self, base):
        """Return (reference, suffixed?)."""
        cand, n = base, 0
        while self._taken(cand):
            n += 1
            if self.style == "date" and self.work_date:
                tag = self.work_date.strftime("%m%d%y")
                cand = f"{base}-{tag}" if n == 1 else f"{base}-{tag}-{n - 1}"
            else:
                cand = f"{base}-{n}"
        self.created.add(norm_ref(cand))
        return cand, n > 0
