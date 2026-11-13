"""Lokale uren- en factuurcontrole voor de Berco-weekbestanden."""
import re
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO

import pandas as pd
import pdfplumber
from openpyxl import load_workbook


def cents(value):
    return int((Decimal(str(value)) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def pdf_lines(page):
    lines = []
    for word in sorted(page.extract_words(), key=lambda w: (w["top"], w["x0"])):
        if not lines or abs(lines[-1][0]["top"] - word["top"]) > 3:
            lines.append([])
        lines[-1].append(word)
    return [sorted(line, key=lambda w: w["x0"]) for line in lines]


CLOCK_CODES = {
    "01": "Gewerkte uren",
    "02": "Dokter",
    "03": "Ziek",
    "04": "Vakantie",
    "05": "Te weinig gewerkt",
    "09": "Uren overwerk 150%",
    "12": "Cursus / opleiding",
    "te": "Te laat",
}


def legend_columns(frame):
    aliases = {"Dokter/Tandarts/Specialist": "Dokter", "Verlof af": "Vakantie"}
    return frame.rename(columns={old: new for old, new in aliases.items() if new not in frame})


def clock_pdf(data):
    rows = []
    employer = name = personnel = None
    with pdfplumber.open(BytesIO(data)) as doc:
        for number, page in enumerate(doc.pages, 1):
            has_total = False
            before = len(rows)
            positions = None
            dummy_numbers = ["Gewerkt", "Ziek", "Verlof af", "Overwerk 150%", "Totaal uren", "Uren DP"]
            for words in pdf_lines(page):
                line = " ".join(w["text"] for w in words)
                if line.startswith("Werkgever:"):
                    page_employer = line.partition(":")[2].strip()
                    if page_employer != employer:
                        name = personnel = None
                    employer = page_employer
                elif match := re.fullmatch(r"(\d{6,})\s+(.+)", line):
                    personnel, name = match.groups()
                elif line.startswith("Datum "):
                    labels = [w["text"].casefold() for w in words]
                    coded = "01" in labels
                    if coded:
                        codes = labels[labels.index("grootste") + 1:labels.index("totaal")]
                        if set(codes) != set(CLOCK_CODES) or len(codes) != len(CLOCK_CODES):
                            raise ValueError(f"Onbekende of ontbrekende uurcode op pagina {number}.")
                        numbers = [CLOCK_CODES[c] for c in codes] + ["Totaal uren", "Uren DP"]
                        headers = ["kleinste", "grootste", *codes, "totaal", "uren"]
                        # Kloktijden zijn links uitgelijnd. Uurwaarden staan rechts onder de codes.
                        positions = [words[labels.index(h)]["x0" if i < 2 else "x1"]
                                     for i, h in enumerate(headers)]
                    else:
                        numbers = dummy_numbers
                        if not all(label in line for label in numbers):
                            raise ValueError(f"Onbekende kolomindeling op pagina {number}.")
                        positions = [next(w["x0"] for w in words if w["text"] == label.split()[0])
                                     for label in ["Kleinste", "Grootste", *numbers]]
                elif line.startswith("Totaal ") and positions:
                    has_total = True
                    name = personnel = None
                elif re.search(r"\bPagina \d+ van \d+$", line, re.I):
                    continue
                elif re.match(r"^\d{2}-\d{2}-\d{4}\b", line):
                    if not (name and employer and positions):
                        raise ValueError(f"Onvolledige tabel op pagina {number}.")
                    fields = [[] for _ in positions]
                    for w in words[1:]:
                        if coded:
                            edge = "x0" if ":" in w["text"] or "-" in w["text"].strip("-") else "x1"
                            choices = range(2) if edge == "x0" else range(2, len(positions))
                            index = min(choices, key=lambda i: abs(positions[i] - w[edge]))
                            fields[index].append(w["text"])
                            continue
                        index = max((i for i, x in enumerate(positions) if w["x0"] >= x - 3), default=-1)
                        if index < 0:
                            raise ValueError(f"Onbekende kolom op pagina {number}.")
                        fields[index].append(w["text"])
                    values = [" ".join(f) for f in fields]
                    hours = [float(v.replace(",", ".")) if v else 0 for v in values[2:]]
                    row = dict(zip(numbers, hours))
                    row.update({"Werkgever": employer, "Persnr": personnel, "Werknemer": name,
                                "Datum": line.split()[0], "Kloktijd van": values[0], "Kloktijd tot": values[1]})
                    rows.append(row)
            if len(rows) == before and not has_total:
                raise ValueError(f"Geen herkenbare klokregels op pagina {number}.")
    if not rows:
        raise ValueError("De klokwerkstaat bevat geen dagregistraties.")
    frame = pd.DataFrame(rows)
    frame = frame.rename(columns={"Gewerkt": "Gewerkte uren", "Overwerk 150%": "Uren overwerk 150%"})
    frame = legend_columns(frame)
    columns = ["Werkgever", "Persnr", "Werknemer", "Datum", "Kloktijd van", "Kloktijd tot",
                  "Gewerkte uren", "Ziek", "Vakantie", "Uren overwerk 150%", "Totaal uren", "Uren DP"]
    return frame[columns + [label for label in CLOCK_CODES.values() if label in frame and label not in columns]]


def expected(data, company):
    sheet = load_workbook(BytesIO(data), data_only=True)[company]
    rows = list(sheet.values)
    offset = 0 if company == "BCC" else 1
    week = f"{int(rows[4][4+offset])}{int(rows[4][3+offset]):02}"
    people = {}
    agreements = []
    for row in rows[8:]:
        if not str(row[1]).strip().isdigit() or not isinstance(row[2+offset], str):
            continue
        label = f"{row[1]} · {row[3+offset]} {row[2+offset]}"
        amount = row[23+offset]
        if label in people or not isinstance(amount, (float, int)):
            raise ValueError(f"Controleer rij van {label} en sla Excel na herberekenen op.")
        people[label] = cents(amount)
        agreements.append([label, *row[19+offset:23+offset]])
    if not people:
        raise ValueError("Geen medewerkers gevonden in dit Synergie-sjabloon.")
    return week, people, pd.DataFrame(agreements, columns=["Medewerker", "Fase", "Tarief 100% (€)", "Tarief 114% (€)", "Tarief 150% (€)"])


def invoice(data, company, week):
    pattern = r"(-?[\d.]+,\d{2})"
    entry = re.compile(r"(.+?)\s+" + pattern + r"\s+(?:" + pattern + r"\s+)?" + pattern + r"\s+" + pattern)
    entries = []
    worker = period = None
    text_pages = []
    with pdfplumber.open(BytesIO(data)) as doc:
        for p in doc.pages:
            text = p.extract_text() or ""
            if not text.strip():
                raise ValueError("Deze PDF bevat een scan. Gebruik een PDF met selecteerbare tekst.")
            text_pages.append(text)
    company_text = "car carpets" if company == "BCC" else "truck components"
    if company_text not in " ".join(text_pages).lower():
        raise ValueError("De bedrijfsnaam in de factuur komt niet overeen met het gekozen tabblad.")
    for text in text_pages:
        for line in text.splitlines():
            person = re.match(r"^(?:De heer|Mevrouw)\s+(.+)", line, re.I)
            if person:
                worker, period = person[1].strip(), None
                continue
            wk = re.search(r"\bWK:\s*(\d{6})", line, re.I)
            if wk:
                period = wk[1]
                if period != week:
                    raise ValueError(f"De factuur bevat week {period}. Excel bevat week {week}.")
                line = line[wk.end():].strip()
            m = entry.fullmatch(line)
            if m and worker and period:
                label, quantity, percentage, rate, amount = m.groups()
                entries.append({"Factuurnaam": worker, "Omschrijving": label, "Aantal": quantity,
                                "Percentage": percentage or "", "Tarief": rate, "Bedrag": amount,
                                "cent": cents(amount.replace(".", "").replace(",", "."))})
            elif worker and re.search(r"\d,\d{2}\b", line):
                raise ValueError("Onbekende bedragregel. Controleer de PDF-indeling: " + line)
    if not entries:
        raise ValueError("Geen herkenbare factuurregels. Controleer de originele PDF-indeling.")
    return pd.DataFrame(entries)


def excel_bytes(frame):
    output = BytesIO()
    frame.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()


def normal_name(value):
    return " ".join(value.casefold().split())


def synergie_form(st, frame, template):
    from datetime import date, timedelta
    from hashlib import sha256

    required = {"Werknemer", "Datum", "Werkgever", "Gewerkte uren", "Uren overwerk 150%"}
    if missing := required.difference(frame.columns):
        st.warning("Voor het Synergie-bestand ontbreken: " + ", ".join(sorted(missing)))
        return
    book = load_workbook(BytesIO(template))
    data = frame.copy()
    data["Datum"] = pd.to_datetime(data["Datum"], dayfirst=True, format="mixed", errors="raise").dt.date
    weeks = {d.isocalendar()[:2] for d in data["Datum"]}
    if len(weeks) != 1:
        raise ValueError("Kies registraties van één kalenderweek.")
    year, week = weeks.pop()
    monday = date.fromisocalendar(year, week, 1)
    st.caption(f"Week {week} van {year} · Controleer de koppelingen en de daguren voor beide bedrijven.")
    basis = st.radio("Hoe staan overuren in ‘Gewerkte uren’ in deze bron?",
                     ["Overuren staan apart", "Overuren zijn inbegrepen"], horizontal=True)
    grid = []
    destinations = []
    for company, employer, start, shift in [("BCC", "car carpets", 35, 0), ("BTC", "truck components", 33, 1)]:
        sheet = book[company]
        if sheet.cell(6, start).value != "Maandag" or sheet.cell(6, start + 15).value != "Zaterdag":
            raise ValueError(f"De dagkolommen in {company} hebben een andere indeling.")
        sheet.cell(5, 4 + shift).value = week
        sheet.cell(5, 5 + shift).value = year
        yc = sheet.cell(5, 5 + shift).coordinate
        wc = sheet.cell(5, 4 + shift).coordinate
        sheet.cell(5, 6 + shift).value = f"=DATE({yc},1,4)-WEEKDAY(DATE({yc},1,4),2)+1+7*({wc}-1)"
        subset = data[data["Werkgever"].astype(str).str.contains(employer, case=False, regex=False)]
        group_cols = ["Werknemer"] + (["Persnr"] if "Persnr" in subset else [])
        groups = {}
        for key, records in subset.groupby(group_cols, dropna=False, sort=False):
            parts = key if isinstance(key, tuple) else (key,)
            groups[" · ".join(map(str, parts))] = (str(parts[0]), records)
        used = set()
        with st.expander(f"Medewerkers koppelen · {company}", expanded=False):
            for row in range(9, sheet.max_row + 1):
                payroll = sheet.cell(row, 2).value
                surname = sheet.cell(row, 3 + shift).value
                first = sheet.cell(row, 4 + shift).value
                if not str(payroll).isdigit() or not isinstance(surname, str) or not isinstance(first, str):
                    continue
                label = f"{payroll} · {first.strip()} {surname.strip()}"
                aliases = {normal_name(first + " " + surname), normal_name(surname + " " + first)}
                initials = ".".join(x[0] for x in re.findall(r"\w+", surname)) + "."
                aliases.add(normal_name(first.strip() + " " + initials))
                matches = [k for k, (name, _) in groups.items()
                           if any(normal_name(n) in aliases for n in name.split(" / "))]
                options = ["Kies klokregistratie", "Geen uren deze week", *groups]
                selected = st.selectbox(label, options, index=options.index(matches[0]) if len(matches)==1 else 0,
                                        key=f"synergie_{company}_{payroll}_{year}_{week}")
                if selected == options[0]:
                    continue
                if selected in groups and selected in used:
                    raise ValueError(f"Dezelfde klokregistratie is tweemaal gekoppeld in {company}.")
                used.add(selected)
                records = groups[selected][1] if selected in groups else subset.iloc[:0]
                if any(day.weekday()==6 for day in records["Datum"]):
                    raise ValueError(f"Zondagregistratie bij {label}. Het sjabloon heeft geen zondagkolom.")
                for day in range(6):
                    date_value = monday + timedelta(days=day)
                    daily = records[records["Datum"] == date_value]
                    def total(column):
                        return pd.to_numeric(daily[column].astype("string").str.replace(",", ".", regex=False),
                                             errors="raise").sum() if column in daily else 0
                    regular = float(total("Gewerkte uren"))
                    overtime = float(total("Uren overwerk 150%"))
                    if basis == "Overuren zijn inbegrepen":
                        regular -= overtime
                    col = start + day * 3
                    fill = sheet.cell(row,col).fill
                    ploeg = day < 5 and fill.patternType == "solid" and fill.fgColor.type == "theme" and fill.fgColor.theme == 1
                    extras = ", ".join(f"{k}: {total(k):g}" for k in CLOCK_CODES.values()
                                       if k not in ("Gewerkte uren","Uren overwerk 150%") and total(k))
                    grid.append([company,label,date_value,0.0 if ploeg else regular,regular if ploeg else 0.0,overtime,extras])
                    destinations.append((company,row,start,day))
            roster = sum(str(sheet.cell(r,2).value).isdigit() and isinstance(sheet.cell(r,3+shift).value,str)
                         for r in range(9,sheet.max_row+1))
            unresolved = roster - len({r for c,r,_,_ in destinations if c==company})
            if unresolved:
                st.info(f"{unresolved} medewerker(s) vragen nog een keuze.")
            unused = sorted(set(groups)-used)
            if unused:
                st.caption("Klokregistraties buiten de huidige koppeling: " + ", ".join(unused))
        if unresolved:
            st.info(f"Rond de koppelingen voor {company} af om het bestand te maken.")
            return
    if not grid:
        raise ValueError("Geen medewerkers gevonden in het sjabloon.")
    preview = pd.DataFrame(grid, columns=["Bedrijf","Medewerker","Datum","100%","114%","150%","Overige uurcodes"])
    st.caption("Controleer de verdeling per dag. De zwarte invoervakken in dit sjabloon bepalen de voorgestelde ploeguren. Overige uurcodes staan ter beoordeling naast de uren. Controleer ook de bestaande toeslagen in Excel.")
    fingerprint = sha256(template + preview.to_csv(index=False).encode()).hexdigest()[:16]
    edited = st.data_editor(preview, hide_index=True, disabled=["Bedrijf","Medewerker","Datum","Overige uurcodes"],
                            key="daguren_"+fingerprint)
    for (company,row,start,day), values in zip(destinations, edited[["100%","114%","150%"]].itertuples(index=False,name=None)):
        if any(pd.isna(v) or v < 0 for v in values):
            st.warning("Vul voor alle drie de uursoorten een positief getal of nul in.")
            return
        if day == 5 and any(values[:2]):
            st.warning("Op zaterdag heeft het sjabloon alleen een kolom voor 150%. Controleer de zaterdaguren.")
            return
        sheet = book[company]
        for index,value in enumerate(values if day<5 else values[2:]):
            cell = sheet.cell(row,start+day*3+index)
            if cell.data_type == "f":
                raise ValueError(f"Onverwachte formule in invoervak {company} {cell.coordinate}.")
            cell.value = float(value) if value else None
        if day == 5:
            sheet.cell(row,start+17).value = f"={sheet.cell(row,start+16).coordinate}"
    book.calculation.fullCalcOnLoad = True
    book.calculation.calcMode = "auto"
    out = BytesIO()
    book.save(out)
    st.download_button("Download ingevuld Synergie-bestand", out.getvalue(), f"Synergie_{year}-W{week:02}.xlsx")
    st.caption("Open het bestand in Excel, controleer de toeslagen en bedragen en sla het op voordat je het naar Synergie stuurt.")


def main():
    import streamlit as st
    st.set_page_config(page_title="Berco | Urencontrole", layout="wide")
    st.markdown("""
    <style>
    .block-container {max-width:1280px;padding-top:3.5rem;padding-bottom:3rem}
    .berco-header {background:#173344;color:white;border-left:6px solid #51b89a;
        border-radius:12px;padding:24px 32px 28px;margin-bottom:20px;font-family:"Segoe UI",sans-serif}
    .berco-brand {color:#a6e4d2;font-size:.85rem;font-weight:700;letter-spacing:.14em}
    .berco-header h1 {color:white;font-family:inherit;font-size:clamp(2rem,4vw,2.75rem);
        line-height:1.15;margin:8px 0;padding:0}
    .berco-header p {color:#e0ebef;margin:0;font-size:1rem}
    @media(max-width:640px){.berco-header{padding:20px}}
    </style>
    <header class="berco-header"><div class="berco-brand">BERCO</div>
    <h1>Urencontrole</h1><p>Controleer de geregistreerde uren per medewerker en werkdag.</p></header>
    """, unsafe_allow_html=True)
    clock, bills, staff = st.tabs(["Urencontrole", "Factuurcontrole", "Medewerkers info"])
    with staff:
        st.write("Hier wordt in de toekomst een database van gemaakt.")
    with clock:
        with st.container(border=True):
            st.subheader("Weekbestand uploaden")
            source = st.file_uploader("Selecteer de klokexport (.xlsx) of klokwerkstaat (.pdf)", type=["xlsx", "pdf"])
        if source:
            try:
                if source.name.lower().endswith(".pdf"):
                    frame = clock_pdf(source.getvalue())
                else:
                    with pd.ExcelFile(BytesIO(source.getvalue())) as sheets:
                        frame = legend_columns(sheets.parse(st.selectbox("Kies een tabblad", sheets.sheet_names)))
                st.caption(f"{len(frame)} registraties")
                st.subheader("Controle van de uren")
                labels = list(CLOCK_CODES.values())
                columns = ["Werknemer", "Datum", *labels, "Totaal uren", "Uren DP"]
                if any(column in frame for column in labels):
                    for label in labels:
                        if label not in frame:
                            frame[label] = pd.NA
                    columns = [column for column in columns if column in frame]
                    metadata = [col for col in frame if col not in columns]
                    frame = frame[metadata + columns]
                    headers = {label: f"{code.capitalize()} · {label}"
                               for code, label in CLOCK_CODES.items()}
                    st.dataframe(frame[columns], column_config=headers, hide_index=True)
                    st.caption("Een streepje bij een uurcode betekent dat de bron geen waarde bevat.")
                else:
                    st.info("Dit tabblad bevat geen herkenbare urenkolommen.")
                with st.expander("Alle gegevens en kolomnamen"):
                    st.dataframe(frame, column_config={label: f"{code.capitalize()} · {label}" for code, label in CLOCK_CODES.items()}, hide_index=True)
                    st.write(frame.columns.tolist())
                st.download_button("Download registraties als Excel", excel_bytes(frame), "Klokregistraties.xlsx")
                st.subheader("Weekbestand voor Synergie maken")
                template = st.file_uploader("Synergie-sjabloon met medewerkers, fase en tarieven (BCC en BTC)", type="xlsx")
                if template:
                    synergie_form(st, frame, template.getvalue())
            except Exception as exc:
                st.error(str(exc))
    with bills:
        company = st.radio("Bedrijf", ["BCC", "BTC"], horizontal=True,
                           format_func=lambda x: {"BCC": "Car Carpets", "BTC": "Truck Components"}[x])
        book = st.file_uploader("Weekbestand dat Berco naar Synergie heeft gestuurd", type="xlsx")
        bill = st.file_uploader("Factuur van Synergie", type="pdf")
        if not book:
            return
        try:
            week, people, agreements = expected(book.getvalue(), company)
            st.caption(f"Week {week[-2:]} van {week[:4]} · {len(people)} medewerkers in Excel")
            with st.expander("Fase en tarieven uit het verstuurde weekbestand"):
                st.dataframe(agreements, hide_index=True)
            if not bill:
                return
            lines = invoice(bill.getvalue(), company, week)
            st.subheader("Gelezen factuurregels")
            st.dataframe(lines.drop(columns="cent"), hide_index=True)
            st.subheader("Gefactureerde uren per medewerker")
            categories = {"100,00": "Normale uren 100%", "114,00": "Ploeguren 114%",
                          "150,00": "Overuren 150%"}
            hours = lines.loc[lines["Percentage"].isin(categories)].copy()
            if hours.empty:
                st.info("Geen uren met percentage 100%, 114% of 150% gevonden.")
            else:
                hours["Uursoort"] = hours["Percentage"].map(categories)
                hours["Uren"] = pd.to_numeric(hours["Aantal"].str.replace(".", "", regex=False)
                                              .str.replace(",", ".", regex=False))
                totals = hours.pivot_table(index="Factuurnaam", columns="Uursoort", values="Uren",
                                           aggfunc="sum", sort=False, fill_value=0)
                totals = totals.reindex(columns=list(categories.values()), fill_value=0)
                totals["Totaal uren"] = totals.sum(axis=1)
                st.dataframe(totals.rename_axis(columns=None).reset_index(), hide_index=True)
            st.caption("Koppel elke factuurnaam aan de medewerker in Excel. Kies ‘Alleen op factuur’ bij een ontbrekende medewerker.")
            used = set()
            results = []
            for name, group in lines.groupby("Factuurnaam", sort=False):
                options = ["Kies medewerker", "Alleen op factuur", *people]
                matches = [person for person in people
                           if normal_name(person.partition(" · ")[2]) == normal_name(name)]
                default = options.index(matches[0]) if len(matches) == 1 else 0
                person = st.selectbox(name, options, index=default,
                                      key=(company + week + name))
                actual = int(group["cent"].sum())
                if person == "Kies medewerker":
                    st.info("Kies een medewerker om verder te gaan.")
                    return
                if person == "Alleen op factuur":
                    results.append([name, None, actual/100, None, person])
                    continue
                if person in used:
                    st.error("Dezelfde Excel-medewerker is twee keer gekoppeld. Controleer de namen.")
                    return
                used.add(person)
                delta = actual - people[person]
                status = "Gelijk" if delta == 0 else ("Te veel gefactureerd" if delta > 0 else "Te weinig gefactureerd")
                results.append([person, people[person]/100, actual/100, delta/100, status])
            for person in people.keys() - used:
                results.append([person, people[person]/100, None, None, "Alleen in Excel"])
            report = pd.DataFrame(results, columns=["Medewerker", "Verwacht (€)", "Factuur (€)", "Verschil (€)", "Status"])
            report = report.merge(agreements, on="Medewerker", how="left", validate="many_to_one")
            st.subheader("Vergelijking van weekbedragen")
            st.dataframe(report, hide_index=True)
            st.download_button("Download controle", excel_bytes(report), f"Controle_{company}_{week}.xlsx")
        except Exception as exc:
            st.error(str(exc))


if __name__ == "__main__":
    main()
