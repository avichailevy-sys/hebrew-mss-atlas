# -*- coding: utf-8 -*-
"""build_atlas.py - generate atlas_data.json from catalogue_14c.xlsx"""
import pandas as pd
import json
import unicodedata

def nfc(s):
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s

INPUT  = "catalogue_14c.xlsx"
OUTPUT = "atlas_data.json"

GEO = {
  "Rome (Italy)": [41.9028, 12.4964],
  "Bologna (Italy)": [44.4949, 11.3426],
  "Venice (Italy)": [45.4408, 12.3155],
  "Padua (Italy)": [45.4064, 11.8768],
  "Florence (Italy)": [43.7696, 11.2558],
  "Naples (Italy)": [40.8518, 14.2681],
  "Palermo (Italy)": [38.1157, 13.3615],
  "Milan (Italy)": [45.4642, 9.1900],
  "Turin (Italy)": [45.0703, 7.6869],
  "Pisa (Italy)": [43.7228, 10.4017],
  "Ferrara (Italy)": [44.8378, 11.6196],
  "Mantua (Italy)": [45.1564, 10.7914],
  "Verona (Italy)": [45.4384, 10.9916],
  "Ancona (Italy)": [43.6158, 13.5189],
  "Rimini (Italy)": [44.0678, 12.5695],
  "Perugia (Italy)": [43.1107, 12.3908],
  "Siena (Italy)": [43.3186, 11.3306],
  "Tivoli (Italy)": [41.9633, 12.7950],
  "Assisi (Italy)": [43.0707, 12.6196],
  "Modena (Italy)": [44.6471, 10.9252],
  "Cesena (Italy)": [44.1391, 12.2431],
  "Forl\u00ec (Italy)": [44.2226, 12.0408],
  "Faenza (Italy)": [44.2858, 11.8854],
  "Imola (Italy)": [44.3534, 11.7148],
  "Montepulciano (Italy)": [43.0942, 11.7860],
  "Tolentino (Italy)": [43.2086, 13.2828],
  "Gubbio (Italy)": [43.3499, 12.5754],
  "Osimo (Italy)": [43.4855, 13.4830],
  "Sacile (Italy)": [45.9551, 12.5005],
  "Mineo (Italy)": [37.2647, 14.6921],
  "Gerace (Italy)": [38.2700, 16.2206],
  "Geraci Siculo (Italy)": [37.8721, 14.1453],
  "Cividate Camuno (Italy)": [45.9783, 10.2861],
  "Pescia (Italy)": [43.9024, 10.6885],
  "Frascati (Italy)": [41.8095, 12.6810],
  "Amandola (Italy)": [42.9789, 13.3608],
  "Bevagna (Italy)": [42.9342, 12.6086],
  "Spello (Italy)": [42.9908, 12.6700],
  "Macerata (Italy)": [43.3007, 13.4530],
  "Camerino (Italy)": [43.1364, 13.0686],
  "Chieti (Italy)": [42.3517, 14.1675],
  "Fano (Italy)": [43.8430, 13.0186],
  "Civitanova Alta (Italy)": [43.3036, 13.6919],
  "Reggio di Calabria (Italy)": [38.1147, 15.6505],
  "Cento (Italy)": [44.7244, 11.2872],
  "Bertinoro (Italy)": [44.1503, 12.1331],
  "Castel Bolognese (Italy)": [44.3194, 11.8025],
  "Lugo (Italy)": [44.4203, 11.9094],
  "Recanati (Italy)": [43.4011, 13.5503],
  "Pavia (Italy)": [45.1847, 9.1582],
  "Voghera (Italy)": [44.9919, 9.0078],
  "Otranto (Italy)": [40.1450, 18.4920],
  "Treviglio (Italy)": [45.5217, 9.5917],
  "Castelfranco Veneto (Italy)": [45.6722, 11.9281],
  "Legnano (Italy)": [45.5928, 8.9136],
  "Alessandria (Italy)": [44.9133, 8.6153],
  "Visso (Italy)": [42.9300, 13.0894],
  "Narni (Italy)": [42.5183, 12.5158],
  "Rieti (Italy)": [42.4042, 12.8567],
  "Ostra (Marche, Italy)": [43.7547, 13.1606],
  "Viterbo (Italy)": [42.4196, 12.1077],
  "Urbino (Italy)": [43.7264, 12.6364],
  "Ascoli Piceno (Italy)": [42.8520, 13.5750],
  "Pesaro (Italy)": [43.9098, 12.9131],
  "Monticelli (Florence, Italy)": [43.7696, 11.2558],
  "Foligno (Italy)": [42.9558, 12.7036],
  "Toledo (Spain)": [39.8628, -4.0273],
  "Barcelona (Spain)": [41.3851, 2.1734],
  "Seville (Spain)": [37.3891, -5.9845],
  "Zaragoza (Spain)": [41.6488, -0.8891],
  "Burgos (Spain)": [42.3439, -3.6969],
  "Lerida (Spain)": [41.6176, 0.6200],
  "Murcia (Spain)": [37.9922, -1.1307],
  "Soria (Spain)": [41.7665, -2.4796],
  "Gerona (Spain)": [41.9794, 2.8214],
  "Cervera (Spain)": [41.6700, 1.2722],
  "Camprod\u00f3n (Spain)": [42.3128, 2.3625],
  "Calatayud (Spain)": [41.3536, -1.6440],
  "Huesca (Spain)": [42.1401, -0.4087],
  "Majorca (Spain)": [39.5696, 2.6502],
  "Pamplona (Spain)": [42.8125, -1.6458],
  "Castell\u00f3 d'Emp\u00faries (Spain)": [42.2576, 3.0750],
  "Albalate de Cinca (Spain)": [41.7269, 0.1500],
  "Alcolea de Cinca (Spain)": [41.7256, 0.1183],
  "Trujillo (C\u00e1ceres, Spain)": [39.4584, -5.8821],
  "Falset (Spain)": [41.1474, 0.8200],
  "Solsona (Spain)": [41.9941, 1.5169],
  "Milagro (Spain)": [42.2475, -1.6044],
  "La Almunia de Do\u00f1a Godina (Spain)": [41.4811, -1.3739],
  "Paris (France)": [48.8566, 2.3522],
  "Avignon (France)": [43.9493, 4.8055],
  "Tarascon (Bouches-du-Rh\u00f4ne, France)": [43.8055, 4.6608],
  "Marseille (France)": [43.2965, 5.3698],
  "Montpellier (France)": [43.6108, 3.8767],
  "Laon (France)": [49.5641, 3.6242],
  "Evreux (France)": [49.0241, 1.1508],
  "Carpentras (France)": [44.0552, 5.0479],
  "Arles (France)": [43.6764, 4.6304],
  "Lunel (France)": [43.6781, 4.1372],
  "Beaucaire (Gard, France)": [43.8077, 4.6433],
  "Bagnols-sur-C\u00e8ze (France)": [44.1611, 4.6184],
  "Bray-sur-Seine (France)": [48.4172, 3.2403],
  "Chinon (France)": [47.1675, 0.2417],
  "Clisson (France)": [47.0867, -1.2839],
  "Capestang (France)": [43.3261, 3.0392],
  "Colmar (France)": [48.0794, 7.3585],
  "Besan\u00e7on (France)": [47.2378, 6.0241],
  "Orange (France)": [44.1378, 4.8094],
  "Salon-de-Provence (France)": [43.6403, 5.0978],
  "Villefort (Loz\u00e8re, France : Canton)": [44.4392, 3.9272],
  "Lalbenque (France)": [44.3406, 1.5417],
  "Vienne (France)": [45.5256, 4.8744],
  "Vermenton (France : Canton)": [47.6644, 3.7250],
  "Trets (France)": [43.4477, 5.6849],
  "Provence (France)": [43.5297, 5.4474],
  "Graisivaudan Valley (France)": [45.2050, 5.9080],
  "Mainz (Germany)": [50.0000, 8.2711],
  "Worms (Germany)": [49.6336, 8.3594],
  "Cologne (Germany)": [50.9375, 6.9603],
  "Nuremberg (Germany)": [49.4521, 11.0767],
  "Frankfurt am Main (Germany)": [50.1109, 8.6821],
  "Munich (Germany)": [48.1351, 11.5820],
  "Berlin (Germany)": [52.5200, 13.4050],
  "Augsburg (Germany)": [48.3705, 10.8978],
  "Plauen (Germany)": [50.4949, 12.1382],
  "Ingolstadt (Germany)": [48.7665, 11.4258],
  "Andernach (Germany)": [50.4404, 7.4036],
  "Rothenburg ob der Tauber (Germany)": [49.3804, 10.1864],
  "Hammelburg (Germany)": [50.1175, 9.8911],
  "Marburg (Germany)": [50.8021, 8.7669],
  "Konstanz (Germany)": [47.6779, 9.1732],
  "Horn (Germany)": [51.8848, 8.9438],
  "Aue (Saxony, Germany)": [50.5871, 12.7016],
  "G\u00f6rlitz (G\u00f6rlitz, Germany)": [51.1556, 14.9686],
  "M\u00fcnster in Westfalen (Germany)": [51.9607, 7.6261],
  "Straubing (Germany)": [48.8825, 12.5790],
  "F\u00fcrth (Bavaria, Germany)": [49.4775, 10.9886],
  "Bad Kreuznach (Germany)": [49.8447, 7.8669],
  "Bayerisch-Schwaben (Germany : Regierungsbezirk)": [48.3705, 10.8978],
  "Vienna (Austria)": [48.2082, 16.3738],
  "Krems an der Donau (Austria)": [48.4097, 15.6147],
  "D\u00fcrnstein (Austria)": [48.3956, 15.5161],
  "Zurich (Switzerland)": [47.3769, 8.5417],
  "Cheb (Czech Republic)": [50.0793, 12.3700],
  "Krakow (Poland)": [50.0647, 19.9450],
  "Brussels (Belgium)": [50.8503, 4.3517],
  "Louvain (Belgium)": [50.8798, 4.7005],
  "London (England)": [51.5074, -0.1278],
  "Selham (England)": [50.9694, -0.7164],
  "Vlor\u00eb (Albania)": [40.4686, 19.4914],
  "Mistra (Greece)": [37.0732, 22.3650],
  "Thebes (Greece)": [38.3217, 23.3247],
  "Thessalonik\u0113 (Greece)": [40.6401, 22.9444],
  "Chania (Greece)": [35.5138, 24.0180],
  "\u0112rakleion (Greece)": [35.3387, 25.1442],
  "Kerkyra (Greece)": [39.6243, 19.9217],
  "Serrai (Greece)": [41.0856, 23.5478],
  "Rhodes (Greece)": [36.4341, 28.2176],
  "Crete (Greece)": [35.2401, 24.8093],
  "Istanbul (Turkey)": [41.0082, 28.9784],
  "Edirne (Turkey)": [41.6772, 26.5559],
  "\u0130zmir (Turkey)": [38.4192, 27.1287],
  "Mardin (Turkey)": [37.3120, 40.7350],
  "Manisa (Turkey)": [38.6191, 27.4289],
  "Feodosii\ufe20a\ufe21 (Ukraine)": [45.0319, 35.3825],
  "Chufut-Kale (Ukraine)": [44.7406, 33.9214],
  "I\ufe20E\ufe21vpatorii\ufe20a\ufe21 (Ukraine)": [45.1992, 33.3673],
  "Cairo (Egypt)": [30.0444, 31.2357],
  "Alexandria (Egypt)": [31.2001, 29.9187],
  "Mit Damsis (Egypt)": [31.0700, 31.5500],
  "Jerusalem (Israel)": [31.7683, 35.2137],
  "Safed (Israel)": [32.9646, 35.4960],
  "Acre (Israel)": [32.9281, 35.0818],
  "Nablus (West Bank)": [32.2226, 35.2622],
  "Gaza (Gaza Strip)": [31.5017, 34.4668],
  "Damascus (Syria)": [33.5138, 36.2765],
  "Aleppo (Syria)": [36.2021, 37.1343],
  "Masyaf (Syria)": [35.0667, 36.3417],
  "Baghdad (Iraq)": [33.3152, 44.3661],
  "\u02bb\u0100nah (Iraq)": [34.4666, 41.9494],
  "Ba\u02bblabakk (Lebanon)": [34.0042, 36.2075],
  "San'a (Yemen)": [15.3694, 44.1910],
  "Al-Tawilah (Yemen)": [15.5681, 43.8869],
  "Tanam (Yemen)": [14.8000, 43.7000],
  "\u1e62a\u02bbdah (Yemen)": [16.9402, 43.7634],
  "Thil\u0101 (Yemen)": [15.5728, 43.9778],
  "Tabr\u012bz (Iran)": [38.0962, 46.2738],
  "I\u1e63fah\u0101n (Iran)": [32.6546, 51.6680],
  "Tunis (Tunisia)": [36.8065, 10.1815],
  "Kalaat es Senam (Tunisia)": [35.7167, 8.3500],
  "Mostaganem (Algeria)": [35.9314, 0.0892],
  "Constantine (Algeria)": [36.3650, 6.6147],
  "F\u00e8s (Morocco)": [34.0181, -5.0078],
  "Santiago do Cac\u00e9m (Portugal)": [38.0167, -8.7000],
  "Set\u00fabal (Portugal)": [38.5244, -8.8882],
  "Tripoli (Lebanon)": [34.4332, 35.8497],
  "Hillah (Iraq)": [32.4664, 44.4170],
  "Sagunto (Spain)": [39.6764, -0.2766],
  "Dresden (Germany)": [51.0504, 13.7373],
  "Louhans (France)": [46.6311, 5.2197],
  "Tallard (France)": [44.4633, 6.0531],
  "Seligenstadt (Hesse, Germany)": [50.0419, 8.9742],
  "Einbeck (Germany)": [51.8197, 9.8722],
  "Corbeil (France)": [48.6147, 2.4795],
  "Hranice (Olomoucky kraj, Czech Republic)": [49.5483, 17.7339],
  "Bourges (France)": [47.0810, 2.3988],
  "Strasbourg (France)": [48.5734, 7.7521],
  "Tudela (Spain)": [42.0658, -1.6044],
  "Troyes (France)": [48.2973, 4.0744],
  "Norcia (Italy)": [42.7929, 13.0939],
  "Mogendorf (Germany)": [50.5236, 7.7011],
  "Prague (Czech Republic)": [50.0755, 14.4378],
  "Solutr\u00e9-Pouilly (France)": [46.2980, 4.7197],
  "Wernigerode (Germany)": [51.8348, 10.7858],
  "Bursa (Turkey)": [40.1828, 29.0665],
  "Aybak (Afghanistan)": [36.2680, 68.0150],
  "Sant'a (Yemen)": [15.3694, 44.1910],
  "Marseille (France)": [43.2965, 5.3698],
  "Pisa (Italy)": [43.7228, 10.4017],
  "Coburg (Germany)": [50.2620, 10.9619],
  "Blaubeuren (Germany)": [48.4147, 9.7831],
  "Leszno (Poland)": [51.8419, 16.5750],
  "Pamplona (Spain)": [42.8125, -1.6458],
  "Guarda (Guarda, Portugal)": [40.5380, -7.2664],
  "Torres Vedras (Portugal)": [39.0918, -9.2596],
  "Tarquinia (Italy)": [42.2553, 11.7569],
  "Oria (Italy)": [40.5000, 17.6403],
  "Pisa (Italy) ||| Perugia (Italy) ||| Ia\u015fi (Romania) ||| Istanbul (Turkey)": [43.7228, 10.4017],
  "Tallinn (Estonia)": [59.4370, 24.7536],
  "Livorno (Italy)": [43.5485, 10.3106],
  "Poligny (France)": [46.8390, 5.7081],
}

# Normalize all GEO keys to NFC so they match catalog data
GEO = {nfc(k): v for k, v in GEO.items()}

REGION_ONLY = {
  "Spain", "Italy", "Italy, Northern", "Italy, Central",
  "Egypt", "Austria", "Belgium", "Portugal", "France",
  "Germany", "Greece", "Israel", "Yemen (Republic)",
  "Bohemia (Czech Republic)",
}

def first_place(s):
    if pd.isna(s): return None
    return nfc(s.split(" ||| ")[0].strip())

def first_val(s):
    if pd.isna(s): return ""
    return str(s).split(" ||| ")[0].strip()

def main():
    df = pd.read_excel(INPUT)
    df["_first"] = df["place_of_writing"].apply(first_place)
    records = []
    unmatched = {}
    region_only_count = 0

    for _, row in df.iterrows():
        first = row["_first"]
        if first in REGION_ONLY:
            region_only_count += 1
            continue
        if first not in GEO:
            unmatched[first] = unmatched.get(first, 0) + 1
            continue
        lat, lon = GEO[first]
        records.append({
            "id": str(row["id"]) if not pd.isna(row["id"]) else "",
            "title": str(row["title"]) if not pd.isna(row["title"]) else "",
            "year": int(row["year_start"]) if not pd.isna(row["year_start"]) else None,
            "date_text": first_val(row["date_text"]),
            "language": first_val(row["language"]),
            "place": first,
            "lat": lat, "lon": lon,
            "holding_library": first_val(row["holding_library"]),
            "shelfmark": first_val(row["shelfmark"]),
            "genre": first_val(row["genre"]),
        })

    by_place = {}
    for r in records:
        key = r["place"]
        if key not in by_place:
            by_place[key] = {"place": key, "lat": r["lat"], "lon": r["lon"],
                             "count": 0, "manuscripts": []}
        by_place[key]["count"] += 1
        by_place[key]["manuscripts"].append({
            "id": r["id"], "year": r["year"], "title": r["title"],
            "date_text": r["date_text"], "language": r["language"],
            "holding_library": r["holding_library"],
            "shelfmark": r["shelfmark"], "genre": r["genre"],
        })

    for s in by_place.values():
        s["manuscripts"].sort(key=lambda m: (m["year"] is None, m["year"] or 0))

    places = sorted(by_place.values(), key=lambda s: -s["count"])

    out = {
        "meta": {
            "title": "Hebrew Manuscripts of the 14th Century",
            "subtitle": "Place of writing, from the NLI Ktiv catalog",
            "total_records_in_subset": int(len(df)),
            "rendered": len(records),
            "unique_places": len(places),
            "region_only_excluded": region_only_count,
            "unmatched_excluded": sum(unmatched.values()),
        },
        "places": places,
    }
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"Wrote {OUTPUT}")
    print(f"  Records rendered:    {len(records)} / {len(df)}")
    print(f"  Unique places:       {len(places)}")
    print(f"  Region-only skipped: {region_only_count}")
    print(f"  Unmatched skipped:   {sum(unmatched.values())}")
    if unmatched:
        print("\n  Unmatched (top 20):")
        for k, v in sorted(unmatched.items(), key=lambda kv: -kv[1])[:20]:
            print(f"    {v:3d}  {k}")

if __name__ == "__main__":
    main()
