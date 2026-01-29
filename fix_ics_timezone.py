import re
from datetime import datetime, timedelta

import requests

OUTLOOK_ICS_URL = r"https://outlook.office365.com/owa/calendar/89ebda66f073440ba078a3c877245d6f@amsterdam.nl/c00c9d1d76bd4c1b8d3230c0f82f3b74390987971231848833/S-1-8-4130157284-2221344195-4234538342-997209183/reachcalendar.ics"
OUTPUT_FILE = "fixed_calendar.ics"
TZID = "Europe/Amsterdam"

# Minimal VTIMEZONE (werkt voor Google/Outlook; niet perfect historisch, maar voldoende voor agenda’s)
VTIMEZONE_BLOCK = f"""BEGIN:VTIMEZONE
TZID:{TZID}
BEGIN:STANDARD
DTSTART:19701025T030000
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19700329T020000
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
END:DAYLIGHT
END:VTIMEZONE
"""

def unfold_ics(text: str) -> str:
    # RFC 5545: regels die beginnen met spatie/tab zijn een voortzetting
    return re.sub(r"\r?\n[ \t]", "", text)

def fold_ics(text: str) -> str:
    # Optioneel: vouw lange regels (75 octets). Voor nu laten we het simpel: geen folding.
    # Google slikt lange regels meestal prima.
    return text

def last_sunday(year: int, month: int) -> datetime:
    # Vind de laatste zondag van een maand
    d = datetime(year, month, 31)
    while d.weekday() != 6:  # maandag=0 ... zondag=6
        d -= timedelta(days=1)
    return d

def is_dst_europe_amsterdam(dt_local: datetime) -> bool:
    """
    DST (zomertijd) EU-regel: laatste zondag maart 01:00 UTC tot laatste zondag oktober 01:00 UTC.
    In lokale tijd: start rond 02:00 → 03:00, einde rond 03:00 → 02:00.
    We gebruiken een praktische benadering:
    - zomertijd tussen laatste zondag maart 02:00 (lokale tijd) en laatste zondag oktober 03:00 (lokale tijd).
    """
    start = last_sunday(dt_local.year, 3).replace(hour=2, minute=0, second=0)
    end = last_sunday(dt_local.year, 10).replace(hour=3, minute=0, second=0)
    return start <= dt_local < end

def utc_z_to_local_eu_amsterdam(value: str) -> str:
    """
    Converteer 'YYYYMMDDTHHMMSSZ' (UTC) naar lokale tijd Europe/Amsterdam en geef terug zonder 'Z'.
    """
    dt_utc = datetime.strptime(value, "%Y%m%dT%H%M%SZ")
    # Eerst “ruw” naar CET (+1), daarna corrigeren naar CEST (+2) als het dst is.
    dt_local = dt_utc + timedelta(hours=1)
    if is_dst_europe_amsterdam(dt_local):
        dt_local = dt_utc + timedelta(hours=2)
    return dt_local.strftime("%Y%m%dT%H%M%S")

def ensure_calendar_headers(lines: list[str]) -> list[str]:
    """
    Zorg dat X-WR-TIMEZONE in VCALENDAR staat (Google gebruikt dit).
    """
    out = []
    inserted = False
    for line in lines:
        out.append(line)
        if not inserted and line.strip() == "BEGIN:VCALENDAR":
            out.append(f"X-WR-TIMEZONE:{TZID}")
            inserted = True
    return out

def ensure_vtimezone(lines: list[str]) -> list[str]:
    if any(l.strip() == "BEGIN:VTIMEZONE" for l in lines):
        return lines
    out = []
    inserted = False
    for line in lines:
        out.append(line)
        if not inserted and line.strip() == "BEGIN:VCALENDAR":
            out.append(VTIMEZONE_BLOCK.strip())
            inserted = True
    return out

def fix_dt_line(line: str) -> str:
    """
    - Laat VALUE=DATE (all-day) met rust
    - Als TZID ontbreekt: voeg TZID toe
    - Als waarde eindigt op Z: converteer UTC→local en verwijder Z
    """
    if not (line.startswith("DTSTART") or line.startswith("DTEND")):
        return line

    # all-day event: DTSTART;VALUE=DATE:20260129
    if "VALUE=DATE" in line:
        return line

    # Splits "KEY;params:VALUE"
    m = re.match(r"^(DTSTART|DTEND)([^:]*)\:(.*)$", line)
    if not m:
        return line

    key, params, value = m.groups()

    # UTC Z? converteer naar lokale tijd
    if value.endswith("Z") and re.match(r"^\d{8}T\d{6}Z$", value):
        value = utc_z_to_local_eu_amsterdam(value)

    # Voeg TZID toe als die nog niet aanwezig is
    if "TZID=" not in params:
        params = f"{params};TZID={TZID}"

    return f"{key}{params}:{value}"

def main():
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(OUTLOOK_ICS_URL, headers=headers, timeout=30)
    r.raise_for_status()

    raw = unfold_ics(r.text)
    lines = raw.splitlines()

    # Fix DTSTART/DTEND regels
    fixed = [fix_dt_line(l) for l in lines]

    # Voeg calendar-level timezone header toe
    fixed = ensure_calendar_headers(fixed)

    # Voeg VTIMEZONE toe (als die ontbreekt)
    fixed = ensure_vtimezone(fixed)

    final_text = fold_ics("\r\n".join(fixed) + "\r\n")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(final_text)

    print(f"✅ Klaar: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
