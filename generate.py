from pathlib import Path
import json
import html
import pandas as pd


# ============================================================
# KONFIGURATION
# ============================================================

EXCEL_DATEI = Path("data/rezepte.xlsx")
AUSGABE_DATEI = AUSGABE_DATEI = Path("index.html")

REZEPT_SHEET = "Rezepet"
ZUTATEN_SHEET = "Zutaten_basis"

KATEGORIEN = [
    "Gemüse",
    "Obst",
    "Protein",
    "Kohlenhydrate",
    "Sonstiges",
]

WOCHENTAGE = [
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
]


# ============================================================
# EXCEL EINLESEN
# ============================================================

def lade_daten():
    print(f"Lese Excel-Datei: {EXCEL_DATEI}")

    rezepte_df = pd.read_excel(
        EXCEL_DATEI,
        sheet_name=REZEPT_SHEET
    )

    zutaten_df = pd.read_excel(
        EXCEL_DATEI,
        sheet_name=ZUTATEN_SHEET
    )

    return rezepte_df, zutaten_df


# ============================================================
# DATEN VALIDIEREN
# ============================================================

def validiere_daten(rezepte_df, zutaten_df):

    rezept_spalten = {
        "Rezept-ID",
        "Rezept"
    }

    zutaten_spalten = {
        "Rezept-ID",
        "Menge",
        "Einheit",
        "Zutaten",
        "Zutaten-Kategorie"
    }

    fehlend_rezepte = rezept_spalten - set(rezepte_df.columns)
    fehlend_zutaten = zutaten_spalten - set(zutaten_df.columns)

    if fehlend_rezepte:
        raise ValueError(
            f"Fehlende Spalten im Tabellenblatt "
            f"'{REZEPT_SHEET}': {fehlend_rezepte}"
        )

    if fehlend_zutaten:
        raise ValueError(
            f"Fehlende Spalten im Tabellenblatt "
            f"'{ZUTATEN_SHEET}': {fehlend_zutaten}"
        )

    print("✓ Excel-Struktur ist korrekt.")


# ============================================================
# DATENMODELL ERSTELLEN
# ============================================================

def erstelle_rezept_daten(rezepte_df, zutaten_df):

    # Leere Werte behandeln
    rezepte_df = rezepte_df.copy()
    zutaten_df = zutaten_df.copy()

    rezepte_df["Rezept-ID"] = rezepte_df["Rezept-ID"].astype(str)
    zutaten_df["Rezept-ID"] = zutaten_df["Rezept-ID"].astype(str)

    zutaten_by_rezept = {}

    for _, row in zutaten_df.iterrows():

        rezept_id = row["Rezept-ID"]

        zutat = {
            "menge": row["Menge"] if pd.notna(row["Menge"]) else "",
            "einheit": (
                str(row["Einheit"]).strip()
                if pd.notna(row["Einheit"])
                else ""
            ),
            "name": (
                str(row["Zutaten"]).strip()
                if pd.notna(row["Zutaten"])
                else ""
            ),
            "kategorie": (
                str(row["Zutaten-Kategorie"]).strip()
                if pd.notna(row["Zutaten-Kategorie"])
                else "Sonstiges"
            ),
        }

        if rezept_id not in zutaten_by_rezept:
            zutaten_by_rezept[rezept_id] = []

        zutaten_by_rezept[rezept_id].append(zutat)

    rezepte = []

    for _, row in rezepte_df.iterrows():

        rezept_id = str(row["Rezept-ID"])

        rezept_name = (
            str(row["Rezept"]).strip()
            if pd.notna(row["Rezept"])
            else "Unbekanntes Rezept"
        )

        rezept = {
            "id": rezept_id,
            "name": rezept_name,
            "zutaten": zutaten_by_rezept.get(rezept_id, [])
        }

        rezepte.append(rezept)

    return rezepte


# ============================================================
# HTML ERZEUGEN
# ============================================================

def erstelle_html(rezepte):

    # Python → JSON
    rezepte_json = json.dumps(
        rezepte,
        ensure_ascii=False
    )

    kategorien_json = json.dumps(
        KATEGORIEN,
        ensure_ascii=False
    )

    wochen_tage_json = json.dumps(
        WOCHENTAGE,
        ensure_ascii=False
    )

    html_datei = f"""<!DOCTYPE html>
<html lang="de">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>Mein Wochenplan</title>

<style>

* {{
    box-sizing: border-box;
}}

body {{
    margin: 0;
    font-family:
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        Roboto,
        Arial,
        sans-serif;

    background: #f4f5f7;
    color: #222;
}}

.container {{
    max-width: 1200px;
    margin: auto;
    padding: 30px 20px 60px;
}}

header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 20px;
    margin-bottom: 30px;
}}

h1 {{
    margin: 0;
    font-size: 32px;
}}

button {{
    border: none;
    border-radius: 10px;
    padding: 13px 20px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    background: #222;
    color: white;
}}

button:hover {{
    opacity: 0.85;
}}

.wochenplan {{
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 12px;
    margin-bottom: 40px;
}}

.tag {{
    background: white;
    border-radius: 14px;
    padding: 16px;
    min-height: 140px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}}

.tag-name {{
    font-size: 13px;
    font-weight: 700;
    color: #777;
    text-transform: uppercase;
    margin-bottom: 15px;
}}

.gericht {{
    font-size: 17px;
    font-weight: 700;
    line-height: 1.3;
}}

.einkaufsliste {{
    background: white;
    border-radius: 16px;
    padding: 25px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}}

.einkaufsliste h2 {{
    margin-top: 0;
}}

.kategorie {{
    margin-top: 25px;
}}

.kategorie h3 {{
    margin-bottom: 10px;
}}

.zutat {{
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 9px 0;
    border-bottom: 1px solid #eee;
}}

.zutat input {{
    width: 18px;
    height: 18px;
}}

.zutat.erledigt span {{
    text-decoration: line-through;
    color: #999;
}}

.menge {{
    margin-left: auto;
    color: #555;
    font-weight: 500;
}}

.leer {{
    color: #999;
}}

@media (max-width: 900px) {{

    .wochenplan {{
        grid-template-columns: repeat(2, 1fr);
    }}

}}

@media (max-width: 550px) {{

    .container {{
        padding: 20px 12px 40px;
    }}

    header {{
        flex-direction: column;
        align-items: stretch;
    }}

    h1 {{
        font-size: 26px;
    }}

    button {{
        width: 100%;
    }}

    .wochenplan {{
        grid-template-columns: 1fr;
    }}

}}

</style>

</head>

<body>

<div class="container">

<header>

    <div>
        <h1>🍴 Mein Wochenplan</h1>
    </div>

    <button onclick="neueWoche()">
        🔀 Neue Woche
    </button>

</header>


<section>

    <div
        id="wochenplan"
        class="wochenplan"
    ></div>

</section>


<section class="einkaufsliste">

    <h2>🛒 Einkaufsliste</h2>

    <div id="einkaufsliste"></div>

</section>

</div>


<script>

// ============================================================
// DATEN AUS PYTHON
// ============================================================

const REZEPTE = {rezepte_json};

const KATEGORIEN = {kategorien_json};

const WOCHENTAGE = {wochen_tage_json};


// ============================================================
// HILFSFUNKTION
// ============================================================

function zufall(min, max) {{

    return Math.floor(
        Math.random() * (max - min + 1)
    ) + min;

}}


// ============================================================
// 7 EINZIGARTIGE REZEPTE AUSWÄHLEN
// ============================================================

function zieheSiebenRezepte() {{

    if (REZEPTE.length < 7) {{

        alert(
            "Es müssen mindestens 7 Rezepte vorhanden sein."
        );

        return [];

    }}

    const kopie = [...REZEPTE];

    // Fisher-Yates Shuffle

    for (
        let i = kopie.length - 1;
        i > 0;
        i--
    ) {{

        const j = zufall(0, i);

        [
            kopie[i],
            kopie[j]
        ] = [
            kopie[j],
            kopie[i]
        ];

    }}

    return kopie.slice(0, 7);

}}


// ============================================================
// WOCHENPLAN ERZEUGEN
// ============================================================

function neueWoche() {{

    const auswahl = zieheSiebenRezepte();

    if (auswahl.length !== 7) {{
        return;
    }}

    const wochenplan = [];

    for (let i = 0; i < 7; i++) {{

        wochenplan.push({{
            tag: WOCHENTAGE[i],
            rezept: auswahl[i]
        }});

    }}

    zeigeWochenplan(wochenplan);

    erstelleEinkaufsliste(wochenplan);

}}


// ============================================================
// WOCHENPLAN DARSTELLEN
// ============================================================

function zeigeWochenplan(wochenplan) {{

    const container =
        document.getElementById("wochenplan");

    container.innerHTML = "";

    wochenplan.forEach(tag => {{

        const element =
            document.createElement("div");

        element.className = "tag";

        element.innerHTML = `
            <div class="tag-name">
                ${{tag.tag}}
            </div>

            <div class="gericht">
                ${{escapeHtml(tag.rezept.name)}}
            </div>
        `;

        container.appendChild(element);

    }});

}}


// ============================================================
// EINKAUFSLISTE
// ============================================================

function erstelleEinkaufsliste(wochenplan) {{

    const einkaufsliste = {{}};

    KATEGORIEN.forEach(kategorie => {{
        einkaufsliste[kategorie] = {{}};
    }});


    // Zutaten aller Gerichte sammeln

    wochenplan.forEach(tag => {{

        tag.rezept.zutaten.forEach(zutat => {{

            let kategorie = zutat.kategorie;

            if (!KATEGORIEN.includes(kategorie)) {{
                kategorie = "Sonstiges";
            }}

            const name = zutat.name;
            const einheit = zutat.einheit;

            if (!name) {{
                return;
            }}

            const schluessel =
                name.toLowerCase() + "|" + einheit.toLowerCase();


            if (!einkaufsliste[kategorie][schluessel]) {{

                einkaufsliste[kategorie][schluessel] = {{
                    name: name,
                    menge: 0,
                    einheit: einheit
                }};

            }}


            const menge =
                parseFloat(
                    String(zutat.menge)
                        .replace(",", ".")
                );


            if (!isNaN(menge)) {{

                einkaufsliste[kategorie][schluessel].menge
                    += menge;

            }}

        }});

    }});


    zeigeEinkaufsliste(einkaufsliste);

}}


// ============================================================
// EINKAUFSLISTE DARSTELLEN
// ============================================================

function zeigeEinkaufsliste(einkaufsliste) {{

    const container =
        document.getElementById("einkaufsliste");

    container.innerHTML = "";


    KATEGORIEN.forEach(kategorie => {{

        const zutaten =
            Object.values(
                einkaufsliste[kategorie]
            );


        const bereich =
            document.createElement("div");

        bereich.className = "kategorie";


        const titel =
            document.createElement("h3");

        titel.textContent =
            kategorie;


        bereich.appendChild(titel);


        if (zutaten.length === 0) {{

            const leer =
                document.createElement("div");

            leer.className = "leer";

            leer.textContent =
                "Keine Zutaten";

            bereich.appendChild(leer);

        }} else {{

            zutaten.forEach((zutat, index) => {{

                const element =
                    document.createElement("label");

                element.className = "zutat";

                const checkbox =
                    document.createElement("input");

                checkbox.type = "checkbox";


                const text =
                    document.createElement("span");

                text.textContent =
                    zutat.name;


                const menge =
                    document.createElement("span");

                menge.className = "menge";

                menge.textContent =
                    formatiereMenge(
                        zutat.menge
                    ) +
                    " " +
                    zutat.einheit;


                checkbox.addEventListener(
                    "change",
                    function() {{

                        element.classList.toggle(
                            "erledigt",
                            checkbox.checked
                        );

                    }}
                );


                element.appendChild(checkbox);
                element.appendChild(text);
                element.appendChild(menge);

                bereich.appendChild(element);

            }});

        }}


        container.appendChild(bereich);

    }});

}}


// ============================================================
// MENGE FORMATIEREN
// ============================================================

function formatiereMenge(menge) {{

    if (Number.isInteger(menge)) {{
        return menge;
    }}

    return Number(menge.toFixed(2));
}}


// ============================================================
// HTML SICHER AUSGEBEN
// ============================================================

function escapeHtml(text) {{

    const div =
        document.createElement("div");

    div.textContent = text;

    return div.innerHTML;

}}


// ============================================================
// ERSTE WOCHE AUTOMATISCH ERZEUGEN
// ============================================================

neueWoche();

</script>

</body>

</html>
"""

    return html_datei


# ============================================================
# MAIN
# ============================================================

def main():

    AUSGABE_DATEI.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    rezepte_df, zutaten_df = lade_daten()

    validiere_daten(
        rezepte_df,
        zutaten_df
    )

    rezepte = erstelle_rezept_daten(
        rezepte_df,
        zutaten_df
    )

    print(
        f"✓ {len(rezepte)} Rezepte geladen."
    )

    html_inhalt = erstelle_html(rezepte)

    AUSGABE_DATEI.write_text(
        html_inhalt,
        encoding="utf-8"
    )

    print()
    print(
        f"✓ HTML erstellt: {AUSGABE_DATEI}"
    )
    print()
    print(
        "Du kannst die HTML-Datei jetzt "
        "direkt im Browser öffnen oder verschicken."
    )


if __name__ == "__main__":
    main()