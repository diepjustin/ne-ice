#!/usr/bin/env python3
"""
Build ne-ice/data.json from the Deportation Data Project "Arrests with detentions"
export, filtered to the seven Nebraska ICE detention facilities.

Input  (not committed): joined-arrests-detention-stays_filtered_*.xlsx
Output (committed):     data.json  -- small, pre-aggregated, no individual rows

Re-pull the input from https://app.deportationdata.org
  dataset = "Arrests with detentions"
  filter  = detention_facility_codes_all contains any of
            OMAHOLD CASSCNE DAKOTNE LINCONE NEMCCOI PHELPNE SARPYNE  (match: any)
  download = Excel (.xlsx)

Then:  python3 build.py path/to/that.xlsx
"""

import sys, json, glob, datetime as dt
from collections import Counter
import pandas as pd

# ---------------------------------------------------------------------------
# The seven facilities the filter targets, for the methodology note.
FACILITIES = {
    "OMAHOLD": "Omaha ERO hold room",
    "CASSCNE": "Cass County Jail (Plattsmouth)",
    "DAKOTNE": "Dakota County Jail (Dakota City)",
    "LINCONE": "Lincoln County Jail (North Platte)",
    "NEMCCOI": "McCook IGSA (Red Willow County)",
    "PHELPNE": "Phelps County Jail (Holdrege)",
    "SARPYNE": "Sarpy County Jail (Papillion)",
}

STATE_ABBR = {
    "NEBRASKA": "NE", "IOWA": "IA", "SOUTH DAKOTA": "SD", "MINNESOTA": "MN",
    "NORTH DAKOTA": "ND", "MISSOURI": "MO", "KANSAS": "KS", "WISCONSIN": "WI",
    "ILLINOIS": "IL", "MICHIGAN": "MI", "NEW YORK": "NY", "NORTH CAROLINA": "NC",
    "TEXAS": "TX", "OHIO": "OH", "COLORADO": "CO", "CALIFORNIA": "CA",
    "CONNECTICUT": "CT",
}

# Approximate town centroids, keyed "City, ST". Hand-built to cover every
# apprehension city with 3+ records (~91% of rows); the rest fall back to a
# state marker. Coordinates are good to a few hundred metres -- enough for a
# dot map, not for anything finer.
GEO = {
    # --- Nebraska ---
    "Omaha, NE": [41.2565, -95.9345],
    "Lincoln, NE": [40.8136, -96.7026],
    "Grand Island, NE": [40.9264, -98.3420],
    "Papillion, NE": [41.1544, -96.0422],
    "Bellevue, NE": [41.1370, -95.9145],
    "La Vista, NE": [41.1839, -96.0311],
    "Fremont, NE": [41.4333, -96.4981],
    "Columbus, NE": [41.4297, -97.3684],
    "Kearney, NE": [40.6994, -99.0817],
    "Hastings, NE": [40.5861, -98.3900],
    "North Platte, NE": [41.1239, -100.7654],
    "Norfolk, NE": [42.0281, -97.4170],
    "Schuyler, NE": [41.4472, -97.0587],
    "Lexington, NE": [40.7808, -99.7415],
    "Gering, NE": [41.8266, -103.6607],
    "Scottsbluff, NE": [41.8666, -103.6672],
    "South Sioux City, NE": [42.4739, -96.4136],
    "Dakota City, NE": [42.4142, -96.4200],
    "Madison, NE": [41.8297, -97.4548],
    "Wahoo, NE": [41.2114, -96.6203],
    "David City, NE": [41.2528, -97.1298],
    "York, NE": [40.8681, -97.5925],
    "Wilber, NE": [40.4820, -96.9628],
    "Crete, NE": [40.6264, -96.9614],
    "Beatrice, NE": [40.2681, -96.7470],
    "Blair, NE": [41.5442, -96.1253],
    "Plattsmouth, NE": [41.0111, -95.8822],
    "Nebraska City, NE": [40.6767, -95.8592],
    "Pender, NE": [42.1136, -96.7078],
    "Ponca, NE": [42.5647, -96.7059],
    "Wakefield, NE": [42.2694, -96.8642],
    "Neligh, NE": [42.1289, -98.0303],
    "O'Neill, NE": [42.4578, -98.6478],
    "Clearwater, NE": [42.1739, -98.1859],
    "Elkhorn, NE": [41.2889, -96.2358],
    "Gretna, NE": [41.1394, -96.2400],
    "Waverly, NE": [40.9186, -96.5286],
    "Seward, NE": [40.9069, -97.0995],
    "Aurora, NE": [40.8680, -98.0028],
    "Central City, NE": [41.1156, -98.0003],
    "Doniphan, NE": [40.7708, -98.3695],
    "Wood River, NE": [40.8205, -98.6015],
    "Waco, NE": [40.8622, -97.4523],
    "Clay Center, NE": [40.5225, -98.0559],
    "Minden, NE": [40.4972, -98.9484],
    "Ord, NE": [41.6033, -98.9265],
    "St. Paul, NE": [41.2147, -98.4581],
    "St. Libory, NE": [41.0894, -98.3595],
    "Holdrege, NE": [40.4403, -99.3701],
    "Cozad, NE": [40.8611, -99.9876],
    "Ogallala, NE": [41.1281, -101.7196],
    "Imperial, NE": [40.5169, -101.6435],
    "Trenton, NE": [40.1739, -101.0146],
    "McCook, NE": [40.2019, -100.6254],
    "Maxwell, NE": [41.0777, -100.5290],
    "Kimball, NE": [41.2359, -103.6627],
    "Sidney, NE": [41.1428, -102.9779],
    "Falls City, NE": [40.0611, -95.6022],
    "Tecumseh, NE": [40.3697, -96.1961],
    "Lancaster, NE": [40.8136, -96.7026],   # county name used as place
    # --- Iowa ---
    "Des Moines, IA": [41.5910, -93.6037],
    "West Des Moines, IA": [41.5772, -93.7113],
    "Clive, IA": [41.6030, -93.7244],
    "Grimes, IA": [41.6889, -93.7911],
    "Waukee, IA": [41.6114, -93.8858],
    "Ankeny, IA": [41.7317, -93.6002],
    "Adel, IA": [41.6139, -94.0208],
    "Van Meter, IA": [41.5314, -93.9541],
    "Perry, IA": [41.8402, -94.1069],
    "Indianola, IA": [41.3581, -93.5574],
    "New Virginia, IA": [41.1836, -93.7247],
    "Greenfield, IA": [41.3053, -94.4608],
    "Creston, IA": [41.0586, -94.3614],
    "Osceola, IA": [41.0347, -93.7655],
    "Mitchellville, IA": [41.6672, -93.3591],
    "Newton, IA": [41.6997, -93.0479],
    "Boone, IA": [42.0597, -93.8802],
    "Ames, IA": [42.0347, -93.6199],
    "Nevada, IA": [42.0225, -93.4524],
    "Marshalltown, IA": [42.0494, -92.9080],
    "Eldora, IA": [42.3608, -93.0988],
    "Grundy Center, IA": [42.3614, -92.7688],
    "Toledo, IA": [41.9944, -92.5779],
    "Cedar Rapids, IA": [41.9779, -91.6656],
    "Vinton, IA": [42.1683, -92.0246],
    "Coralville, IA": [41.6764, -91.5802],
    "Iowa City, IA": [41.6611, -91.5302],
    "Tipton, IA": [41.7708, -91.1268],
    "Washington, IA": [41.2999, -91.6935],
    "Wapello, IA": [41.1808, -91.1874],
    "Marengo, IA": [41.7980, -92.0704],
    "Monticello, IA": [42.2372, -91.1874],
    "Anamosa, IA": [42.1083, -91.2846],
    "Manchester, IA": [42.4847, -91.4557],
    "Independence, IA": [42.4694, -91.8890],
    "Waterloo, IA": [42.4928, -92.3426],
    "Waverly, IA": [42.7258, -92.4755],
    "Charles City, IA": [43.0666, -92.6727],
    "Osage, IA": [43.2841, -92.8121],
    "Mason City, IA": [43.1536, -93.2010],
    "Clarion, IA": [42.7327, -93.7327],
    "Hampton, IA": [42.7419, -93.2024],
    "Webster City, IA": [42.4692, -93.8158],
    "Fort Dodge, IA": [42.4975, -94.1680],
    "Humboldt, IA": [42.7194, -94.2152],
    "Jefferson, IA": [42.0153, -94.3775],
    "Carroll, IA": [42.0658, -94.8669],
    "Denison, IA": [42.0142, -95.3555],
    "Harlan, IA": [41.6530, -95.3255],
    "Onawa, IA": [42.0269, -96.0969],
    "Sidney, IA": [40.7492, -95.6486],
    "Red Oak, IA": [41.0097, -95.2255],
    "Glenwood, IA": [41.0464, -95.7433],
    "Council Bluffs, IA": [41.2619, -95.8608],
    "Sioux City, IA": [42.4999, -96.4003],
    "Salix, IA": [42.3197, -96.2461],
    "Le Mars, IA": [42.7944, -96.1656],
    "Orange City, IA": [43.0089, -96.0597],
    "Sioux Center, IA": [43.0797, -96.1764],
    "Primghar, IA": [43.0872, -95.6272],
    "Sibley, IA": [43.4008, -95.7522],
    "Rock Rapids, IA": [43.4272, -96.1764],
    "Storm Lake, IA": [42.6461, -95.2097],
    "Cherokee, IA": [42.7494, -95.5522],
    "Spencer, IA": [43.1414, -95.1444],
    "Spirit Lake, IA": [43.4222, -95.1022],
    "Estherville, IA": [43.4019, -94.8322],
    "Emmetsburg, IA": [43.1130, -94.6802],
    "Forest City, IA": [43.2622, -93.6374],
    "New Hampton, IA": [43.0591, -92.3179],
    "Cresco, IA": [43.3811, -92.1141],
    "Decorah, IA": [43.3033, -91.7859],
    "West Union, IA": [42.9622, -91.8085],
    "Waukon, IA": [43.2669, -91.4796],
    "St. Olaf, IA": [42.9169, -91.3835],
    "Dubuque, IA": [42.5006, -90.6646],
    "Clinton, IA": [41.8445, -90.1887],
    "Davenport, IA": [41.5236, -90.5776],
    "Muscatine, IA": [41.4245, -91.0432],
    "Burlington, IA": [40.8075, -91.1129],
    "Mount Pleasant, IA": [40.9636, -91.5579],
    "Ottumwa, IA": [41.0199, -92.4113],
    "Oskaloosa, IA": [41.2964, -92.6446],
    "Montrose, IA": [40.5303, -91.4093],
    "Webster, IA": [41.4286, -92.1466],
    "Iowa, IA": None,   # state-only free text
    # --- South Dakota ---
    "Sioux Falls, SD": [43.5460, -96.7313],
    "Brookings, SD": [44.3114, -96.7984],
    "Watertown, SD": [44.8994, -97.1147],
    "Milbank, SD": [45.2180, -96.6356],
    "Madison, SD": [44.0064, -97.1139],
    "Salem, SD": [43.7241, -97.3892],
    "Mitchell, SD": [43.7094, -98.0298],
    "Huron, SD": [44.3633, -98.2142],
    "Aberdeen, SD": [45.4647, -98.4865],
    "Yankton, SD": [42.8711, -97.3973],
    "Vermillion, SD": [42.7794, -96.9292],
    "Elk Point, SD": [42.6836, -96.6836],
    "Pierre, SD": [44.3683, -100.3510],
    "Rapid City, SD": [44.0805, -103.2310],
    "Box Elder, SD": [44.1097, -103.0699],
    "Sturgis, SD": [44.4097, -103.5091],
    "Spearfish, SD": [44.4908, -103.8594],
    # --- Minnesota ---
    "Minneapolis, MN": [44.9778, -93.2650],
    "Saint Paul, MN": [44.9537, -93.0900],
    "St. Paul, MN": [44.9537, -93.0900],
    "Bloomington, MN": [44.8408, -93.2983],
    "Richfield, MN": [44.8833, -93.2830],
    "Hopkins, MN": [44.9250, -93.4627],
    "Savage, MN": [44.7791, -93.3363],
    "Shakopee, MN": [44.7974, -93.5269],
    "Chaska, MN": [44.7894, -93.6022],
    "Burnsville, MN": [44.7677, -93.2777],
    "Fort Snelling, MN": [44.8930, -93.1811],
    "New Brighton, MN": [45.0655, -93.2016],
    "Brooklyn Park, MN": [45.0941, -93.3563],
    "Lino Lakes, MN": [45.1608, -93.0891],
    "Stillwater, MN": [45.0561, -92.8060],
    "Hastings, MN": [44.7433, -92.8524],
    "Faribault, MN": [44.2949, -93.2688],
    "Worthington, MN": [43.6197, -95.5964],
    "Pipestone, MN": [44.0000, -96.3175],
    "Rochester, MN": [44.0121, -92.4802],
    "Sandstone, MN": [46.1319, -92.8657],
    "Moose Lake, MN": [46.4536, -92.7616],
    # --- North Dakota ---
    "Bismarck, ND": [46.8083, -100.7837],
    "Minot, ND": [48.2330, -101.2923],
    "Grand Forks, ND": [47.9253, -97.0329],
    "Watford City, ND": [47.8022, -103.2830],
}

# Free-text spelling variants -> canonical "City, ST" key in GEO.
CITY_FIXES = {
    "Hastangs, NE": "Hastings, NE",
    "Wilbur, NE": "Wilber, NE",
    "Oneill, NE": "O'Neill, NE",
    "Mccook, NE": "McCook, NE",
    "Toleda, IA": "Toledo, IA",
    "Pringhar, IA": "Primghar, IA",
    "Ornage City, IA": "Orange City, IA",
    "Mt Pleasant, IA": "Mount Pleasant, IA",
    "Mt. Pleasant, IA": "Mount Pleasant, IA",
    "Saint Olaf, IA": "St. Olaf, IA",
    "Ft Snelling, MN": "Fort Snelling, MN",
    "Ft. Snelling, MN": "Fort Snelling, MN",
    "St Paul, MN": "Saint Paul, MN",
}

STATE_MARKER = {   # fallback point when the city is unknown/blank
    "NE": [41.5, -99.9], "IA": [42.0, -93.5], "SD": [44.4, -100.2],
    "MN": [46.0, -94.3], "ND": [47.5, -100.5], "MO": [38.5, -92.5],
    "KS": [38.5, -98.0], "WI": [44.5, -89.5], "IL": [40.0, -89.0],
    "MI": [44.3, -85.6], "TX": [30.5, -98.5], "LA": [31.0, -92.3],
    "AZ": [33.7, -111.9], "MS": [32.7, -89.7], "WA": [47.4, -120.3],
    "CO": [39.2, -105.6], "HI": [20.8, -156.3], "NY": [42.9, -75.5],
}

# The seven Nebraska facilities the filter targets, with a location for the
# routes. OMAHOLD is the ERO processing hold room almost every stay passes
# through; the other six are county-jail beds.
NE_FAC = {
    "OMAHOLD": {"label": "Omaha ERO hold room",            "lat": 41.2565, "lon": -95.9345},
    "DAKOTNE": {"label": "Dakota County Jail, Dakota City", "lat": 42.4141, "lon": -96.4206},
    "NEMCCOI": {"label": "McCook IGSA",                     "lat": 40.2003, "lon": -100.6254},
    "PHELPNE": {"label": "Phelps County Jail, Holdrege",    "lat": 40.4403, "lon": -99.3701},
    "LINCONE": {"label": "Lincoln County Jail, North Platte","lat": 41.1400, "lon": -100.7601},
    "SARPYNE": {"label": "Sarpy County Jail, Papillion",    "lat": 41.1500, "lon": -96.0300},
    "CASSCNE": {"label": "Cass County Jail, Plattsmouth",   "lat": 41.0111, "lon": -95.8819},
}
NE_CODES = list(NE_FAC)

# Out-of-state ICE facilities people are transferred to, keyed by the exact
# detention_facility_last name. Short label + coordinates.
EXIT_FAC = {
    "ALEXANDRIA STAGING FACILITY":        {"label": "Alexandria staging, LA",       "lat": 31.3949, "lon": -92.5279},
    "PORT ISABEL SPC":                    {"label": "Port Isabel, TX",              "lat": 26.0731, "lon": -97.4108},
    "PINE PRAIRIE ICE PROCESSING CENTER": {"label": "Pine Prairie, LA",             "lat": 30.7861, "lon": -92.4218},
    "CENTRAL LOUISIANA ICE PROC CTR":     {"label": "Jena / CLIPC, LA",             "lat": 31.6869, "lon": -92.1332},
    "ERO EL PASO CAMP EAST MONTANA":      {"label": "El Paso camp, TX",             "lat": 31.7920, "lon": -106.3210},
    "WINN CORRECTIONAL CENTER":           {"label": "Winnfield, LA",                "lat": 31.9251, "lon": -92.6471},
    "JACKSON PARISH CORRECTIONAL CENTER": {"label": "Jonesboro, LA",                "lat": 32.2418, "lon": -92.7188},
    "SOUTH LOUISIANA ICE PROC CTR":       {"label": "Basile, LA",                   "lat": 30.4880, "lon": -92.5960},
    "ADAMS COUNTY CORRECTIONAL CENTER":   {"label": "Natchez, MS",                  "lat": 31.5604, "lon": -91.4032},
    "POTTAWATTAMIE COUNTY JAIL":          {"label": "Council Bluffs, IA",           "lat": 41.2619, "lon": -95.8608},
    "FLORENCE SPC":                       {"label": "Florence, AZ",                 "lat": 33.0314, "lon": -111.3873},
    "FLORENCE STAGING FACILITY":          {"label": "Florence staging, AZ",         "lat": 33.0314, "lon": -111.3873},
    "RICHWOOD COR CENTER":                {"label": "Monroe / Richwood, LA",        "lat": 32.4515, "lon": -92.0846},
    "AZ REM OP COORD CENTER (AROCC)":     {"label": "Mesa / AROCC, AZ",             "lat": 33.4152, "lon": -111.8315},
    "EL VALLE DETENTION FACILITY":        {"label": "Raymondville, TX",             "lat": 26.4788, "lon": -97.7828},
    "NW ICE PROCESSING CTR":              {"label": "Tacoma, WA",                   "lat": 47.2637, "lon": -122.4290},
    "URSULA CENTRALIZED PROCESSING CNTR": {"label": "McAllen / Ursula, TX",         "lat": 26.2159, "lon": -98.2300},
    "COASTAL BEND DET. FACILITY":         {"label": "Robstown, TX",                 "lat": 27.7903, "lon": -97.6689},
    "River Correctional Center":          {"label": "Ferriday, LA",                 "lat": 31.6302, "lon": -91.5518},
    "HARLINGEN HOLD ROOM":                {"label": "Harlingen, TX",                "lat": 26.1906, "lon": -97.6961},
    "SHERBURNE COUNTY JAIL":              {"label": "Elk River, MN",                "lat": 45.3033, "lon": -93.5674},
    "EL PASO SPC":                        {"label": "El Paso SPC, TX",              "lat": 31.7920, "lon": -106.3210},
    "SOUTH TEXAS ICE PROCESSING CENTER":  {"label": "Pearsall, TX",                 "lat": 28.8906, "lon": -99.0958},
    "BISHOP HENRY WHIPPLE FED BLDG":      {"label": "Fort Snelling, MN",            "lat": 44.8930, "lon": -93.1811},
    "ELOY FED CTR FACILITY (CORE CIVIC)": {"label": "Eloy, AZ",                     "lat": 32.7595, "lon": -111.5487},
}

# Country centroids for the removal legs (top destinations only).
COUNTRY_PT = {
    "Mexico": [23.6, -102.5], "Guatemala": [15.7, -90.2], "Honduras": [14.8, -86.9],
    "El Salvador": [13.8, -88.9], "Nicaragua": [12.9, -85.2], "Venezuela": [7.1, -66.1],
    "Ecuador": [-1.4, -78.4], "Colombia": [4.1, -73.0], "Cuba": [21.5, -79.5],
    "Peru": [-9.2, -75.0], "Brazil": [-10.8, -52.9], "Dominican Republic": [18.7, -70.2],
    "South Sudan": [7.3, 30.3], "Liberia": [6.4, -9.4], "India": [22.0, 79.0],
    "Burma": [21.9, 95.9], "Laos": [18.2, 103.9], "Vietnam": [16.2, 107.8],
    "China, Peoples Republic Of": [35.0, 103.0],
    "Micronesia, Federated States Of": [6.9, 158.2],
}

TOP_N_COUNTRY = 12
TOP_N_METHOD = 8
MIN_ARC = 3          # drop route legs thinner than this


def primary_ne(codes_all):
    """The Nebraska facility a stay is attributed to: the first non-Omaha NE
    facility in the (chronological) codes list, else the Omaha hold room."""
    if not isinstance(codes_all, str):
        return None
    codes = [c.strip() for c in codes_all.split(";")]
    ne = [c for c in codes if c in NE_FAC]
    if not ne:
        return None
    non_oma = [c for c in ne if c != "OMAHOLD"]
    return non_oma[0] if non_oma else "OMAHOLD"


def norm_city(raw_city, raw_state):
    city = (str(raw_city).strip() if pd.notna(raw_city) else "")
    st_full = (str(raw_state).strip().upper() if pd.notna(raw_state) else "")
    st = STATE_ABBR.get(st_full, "")
    if not city or city.lower() in ("nan", "unknown", "none", "n/a"):
        return None, st
    key = f"{city.title()}, {st}" if st else city.title()
    key = CITY_FIXES.get(key, key)
    return key, st


def counter_to_list(counter, key_name, top=None, other_label="Other"):
    items = counter.most_common()
    if top and len(items) > top:
        head = items[:top]
        tail_total = sum(n for _, n in items[top:])
        out = [{key_name: k, "n": n} for k, n in head]
        out.append({key_name: other_label, "n": tail_total})
        return out
    return [{key_name: k, "n": n} for k, n in items]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        path = args[0]
    else:
        hits = sorted(glob.glob("joined-arrests-detention-stays_filtered_*.xlsx"))
        if not hits:
            sys.exit("no input xlsx found; pass the path as an argument")
        path = hits[-1]
    print(f"reading {path}")

    df = pd.read_excel(path, sheet_name="data")
    info = pd.read_excel(path, sheet_name="info")
    n_before = None
    for _, r in info.iterrows():
        if str(r["field"]).startswith("Rows (before"):
            n_before = int(str(r["value"]).replace(",", ""))

    n_rows = len(df)
    n_people = int(df["unique_identifier"].nunique())

    # ---- arrest geography ----
    place_counts = Counter()
    state_counts = Counter()
    city_known = 0
    state_known = 0
    unmapped = Counter()
    for _, r in df.iterrows():
        key, st = norm_city(r.get("apprehension_city"), r.get("apprehension_state_filled_in"))
        if st:
            state_known += 1
            state_counts[st] += 1
        if key is not None:
            city_known += 1
            place_counts[key] += 1

    # Only towns we have a real coordinate for go on the map. Everything else
    # (mostly one-off free-text spellings) still counts in arrest_state below,
    # so the map never invents a location it does not have.
    arrest_places = []
    mapped_n = 0
    for key, n in place_counts.most_common():
        coord = GEO.get(key, "MISS")
        if coord in ("MISS", None):
            if coord == "MISS" and n >= 3:
                unmapped[key] = n
            continue
        city, _, stc = key.rpartition(", ")
        mapped_n += n
        arrest_places.append({
            "city": city or key, "state": stc, "n": n,
            "lat": round(coord[0], 4), "lon": round(coord[1], 4),
        })

    if unmapped:
        print("\nWARNING: apprehension cities with 3+ records and no coordinates:")
        for k, n in unmapped.most_common():
            print(f"  {n:5}  {k}")
        print("  -> add them to GEO in build.py to put them on the map\n")

    # ---- who ----
    citizenship = Counter(df["citizenship_country"].dropna().astype(str).str.title())

    # ---- how ----
    method_simple = Counter(df["apprehension_method_simple"].dropna().astype(str))
    method_detail = Counter(df["apprehension_method"].dropna().astype(str))
    program = Counter(df["final_program"].dropna().astype(str))

    crim_raw = Counter(df["book_in_criminality"].dropna().astype(str))
    crim_map = {
        "1 Convicted Criminal": "Convicted of a crime",
        "2 Pending Criminal Charges": "Criminal charge pending",
        "3 Other Immigration Violator": "No criminal charge or conviction",
    }
    criminality = Counter()
    for k, n in crim_raw.items():
        criminality[crim_map.get(k, k)] += n

    # ---- when ----
    ad = pd.to_datetime(df["apprehension_date"], errors="coerce")
    monthly = Counter(ad.dropna().dt.strftime("%Y-%m"))
    monthly_list = [{"month": m, "n": monthly[m]} for m in sorted(monthly)]
    arrest_min = ad.min()
    arrest_max = ad.max()

    # ---- what happens next ----
    out_state = Counter(df["detention_facility_state_last"].dropna().astype(str))
    removed_mask = df["stay_release_reason"].astype(str).str.strip().str.lower() == "removed"
    removed = df[removed_mask]
    removal_country = Counter(removed["departure_country"].dropna().astype(str).str.title())

    # ---- routes: arrest -> Nebraska facility -> exit facility -> country ----
    arrivals = Counter()      # (town_key, ne_code)
    transfers = Counter()     # (ne_code, exit_key)
    removals = Counter()      # (origin_key, country)   origin = exit_key or ne_code
    ne_totals = Counter()
    exit_pts = {}             # exit_key -> {label,lat,lon}
    routed_in = routed_out = 0

    def exit_key_for(row):
        name = str(row.get("detention_facility_last") or "").strip()
        if name in EXIT_FAC:
            f = EXIT_FAC[name]
            exit_pts[name] = {"label": f["label"], "lat": f["lat"], "lon": f["lon"]}
            return name
        st = STATE_ABBR.get(str(row.get("detention_facility_state_last") or "").strip().upper(), "")
        if st in STATE_MARKER:
            k = f"{st} (other facility)"
            exit_pts[k] = {"label": k, "lat": STATE_MARKER[st][0], "lon": STATE_MARKER[st][1]}
            return k
        return None

    for _, r in df.iterrows():
        ne = primary_ne(r.get("detention_facility_codes_all"))
        if ne is None:
            continue
        ne_totals[ne] += 1
        tkey, _ = norm_city(r.get("apprehension_city"), r.get("apprehension_state_filled_in"))
        if tkey in GEO and GEO[tkey]:
            arrivals[(tkey, ne)] += 1
            routed_in += 1
        last_is_ne = r.get("detention_facility_code_last") in NE_FAC
        ekey = None if last_is_ne else exit_key_for(r)
        if ekey:
            transfers[(ne, ekey)] += 1
            routed_out += 1
        if removed_mask.loc[r.name]:
            country = str(r.get("departure_country") or "").title()
            if country in COUNTRY_PT:
                removals[(ekey or ne, country)] += 1

    def pt_for(key):
        if key in NE_FAC:
            f = NE_FAC[key]; return f["lat"], f["lon"], f["label"]
        if key in exit_pts:
            f = exit_pts[key]; return f["lat"], f["lon"], f["label"]
        c = GEO.get(key)
        if c:
            city, _, st = key.rpartition(", ")
            return c[0], c[1], f"{city}, {st}" if st else key
        return None

    def arcs(counter, kind):
        out = []
        for (a, b), n in counter.most_common():
            if n < MIN_ARC:
                continue
            pa = pt_for(a)
            pb = (COUNTRY_PT[b][0], COUNTRY_PT[b][1], b) if kind == "removal" else pt_for(b)
            if not pa or not pb:
                continue
            out.append({
                "from": [round(pa[0], 2), round(pa[1], 2)], "from_label": pa[2],
                "to": [round(pb[0], 2), round(pb[1], 2)], "to_label": pb[2],
                "n": n,
            })
        return out

    ne_nodes = [
        {"code": c, "label": NE_FAC[c]["label"], "lat": NE_FAC[c]["lat"],
         "lon": NE_FAC[c]["lon"], "n": ne_totals[c]}
        for c in sorted(ne_totals, key=lambda x: -ne_totals[x])
    ]
    exit_nodes = []
    _exit_n = Counter()
    for (ne, ek), n in transfers.items():
        _exit_n[ek] += n
    for ek, n in _exit_n.most_common():
        p = exit_pts[ek]
        exit_nodes.append({"label": p["label"], "lat": p["lat"], "lon": p["lon"], "n": n})
    country_nodes = []
    _cty_n = Counter()
    for (o, c), n in removals.items():
        _cty_n[c] += n
    for c, n in _cty_n.most_common():
        country_nodes.append({"country": c, "lat": COUNTRY_PT[c][0], "lon": COUNTRY_PT[c][1], "n": n})

    data = {
        "meta": {
            "generated": dt.date.today().isoformat(),
            "source": "Deportation Data Project - ICE 'Arrests with detentions', released 21 Aug 2026",
            "dataset_url": "https://deportationdata.org/",
            "facilities": FACILITIES,
            "n_stays": n_rows,
            "n_people": n_people,
            "n_before_filter": n_before,
            "arrest_date_min": arrest_min.date().isoformat() if pd.notna(arrest_min) else None,
            "arrest_date_max": arrest_max.date().isoformat() if pd.notna(arrest_max) else None,
            "pct_city_known": round(100 * city_known / n_rows, 1),
            "pct_state_known": round(100 * state_known / n_rows, 1),
            "pct_on_map": round(100 * mapped_n / n_rows, 1),
            "pct_routed_in": round(100 * routed_in / n_rows, 1),
            "pct_routed_out": round(100 * routed_out / n_rows, 1),
            "n_removed": int(removed_mask.sum()),
        },
        "citizenship": counter_to_list(citizenship, "country", top=TOP_N_COUNTRY),
        "arrest_places": arrest_places,
        "arrest_state": counter_to_list(state_counts, "state"),
        "method": counter_to_list(method_simple, "label"),
        "method_detail": counter_to_list(method_detail, "label", top=TOP_N_METHOD),
        "program": counter_to_list(program, "label", top=TOP_N_METHOD),
        "criminality": counter_to_list(criminality, "label"),
        "monthly": monthly_list,
        "transfer_out_state": counter_to_list(out_state, "state", top=10),
        "removal_country": counter_to_list(removal_country, "country", top=TOP_N_COUNTRY),
        "routes": {
            "arrivals": arcs(arrivals, "arrival"),
            "transfers": arcs(transfers, "transfer"),
            "removals": arcs(removals, "removal"),
        },
        "nodes": {
            "ne_facilities": ne_nodes,
            "exit_facilities": exit_nodes,
            "countries": country_nodes,
        },
    }

    with open("data.json", "w") as f:
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)
    r = data["routes"]
    print(f"wrote data.json  ({n_rows} stays, {n_people} people, "
          f"{len(arrest_places)} arrest towns; arcs: "
          f"{len(r['arrivals'])} in / {len(r['transfers'])} transfer / {len(r['removals'])} removal)")


if __name__ == "__main__":
    main()
