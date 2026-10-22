"""Invoice-number normalisation. Replaces the R12 helper column and VLOOKUPs."""
import re


def clean(value):
    """Cell value -> trimmed text, without losing digits (12345.0 -> '12345')."""
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            value = int(value)
    return re.sub(r"\s+", " ", str(value).strip())


def significant(text):
    """Digits without leading zeros ('026866321582' -> '26866321582')."""
    return text.lstrip("0") or "0"


def metro_key(value):
    """Metro: numeric references compare with leading zeros stripped."""
    s = clean(value).upper()
    return significant(s) if s.isdigit() else s


def r12(value):
    """Costco: last 12 digits, left-padded with 0 to 12 (Excel =RIGHT(A2,12) fixed)."""
    s = clean(value).upper()
    return s[-12:].zfill(12) if s.isdigit() else s


def oracle_invoice_text(trx):
    """Oracle invoice numbers are 12 digits; the export drops a leading 0 on 11-digit ones."""
    s = clean(trx)
    return s.zfill(12) if s.isdigit() and len(s) == 11 else s
