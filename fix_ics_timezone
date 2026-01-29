import re
import requests

OUTLOOK_ICS_URL = r"https://outlook.office365.com/owa/calendar/89ebda66f073440ba078a3c877245d6f@amsterdam.nl/c00c9d1d76bd4c1b8d3230c0f82f3b74390987971231848833/S-1-8-4130157284-2221344195-4234538342-997209183/reachcalendar.ics"
OUTPUT_FILE = "fixed_calendar.ics"
TZID = "Europe/Amsterdam"

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
    return re.sub(r"\r?\n[ \t]", "", text)

def fix_dt_line(line: str) -> str:
    if line.startswith("DTSTART") or line.startswith("DTEND"):
        if "VALUE=DATE" in line or "TZID=" in line:
            return line
        line = line.replace("DTSTART:", f"DTSTART;TZID={TZID}:", 1)
        line = line.replace("DTEND:",   f"DTEND;TZID={TZID}:",   1)
    return line

def main():
    r = requests.get(OUTLOOK_ICS_URL, timeout=30)
    r.raise_for_status()

    raw = unfold_ics(r.text)
    lines = raw.splitlines()
    fixed = []

    has_vtimezone = any(l.strip() == "BEGIN:VTIMEZONE" for l in lines)

    for l in lines:
        fixed.append(fix_dt_line(l))

    if not has_vtimezone:
        out = []
        inserted = False
        for l in fixed:
            out.append(l)
            if not inserted and l.strip() == "BEGIN:VCALENDAR":
                out.append(VTIMEZONE_BLOCK.strip())
                inserted = True
        fixed = out

    final_text = "\r\n".join(fixed) + "\r\n"
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(final_text)

if __name__ == "__main__":
    main()
