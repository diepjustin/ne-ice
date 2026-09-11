# ne-ice

Static page at `https://diepjustin.github.io/ne-ice/` showing where ICE arrests
the people who pass through Nebraska's seven immigration-detention facilities,
what countries they come from, how the arrests happen, and where they go next.

## Files

| file | role |
| --- | --- |
| `index.html` | the whole page — inline CSS/JS, Leaflet from CDN, fetches `data.json` |
| `data.json` | pre-aggregated counts, ~25 KB, **committed** — the only file the page loads |
| `build.py` | reads the source `.xlsx`, writes `data.json` |
| `*.xlsx` | source exports — **git-ignored**, not needed at runtime |

## Data source

Deportation Data Project, processed **"Arrests with detentions"** dataset
(ICE records via FOIA). <https://deportationdata.org/>

Each row is one continuous detention stay, matched to the arrest that began it.
The filter keeps a stay if its facility list (`detention_facility_codes_all`)
contains any of the seven Nebraska codes:

```
OMAHOLD  Omaha ERO hold room
CASSCNE  Cass County Jail (Plattsmouth)
DAKOTNE  Dakota County Jail (Dakota City)
LINCONE  Lincoln County Jail (North Platte)
NEMCCOI  McCook IGSA (Red Willow County)
PHELPNE  Phelps County Jail (Holdrege)
SARPYNE  Sarpy County Jail (Papillion)
```

## Refreshing

1. Open the DDP explorer: <https://app.deportationdata.org/?agency=ice&dataset=joined-arrests-detention-stays>
2. Explore tab -> **Filter by** -> column `Detention facility codes all`, comparison
   `contains`. Add one filter per code above; set the group rule to **any**.
3. Download as **Excel (.xlsx)** into this folder.
4. `python3 build.py joined-arrests-detention-stays_filtered_*.xlsx`
5. Commit the updated `data.json`. `build.py` prints a warning for any arrest city
   with 3+ records it could not geocode — add those to the `GEO` dict.

Requires `pandas` and `openpyxl`.

## Caveats (also stated on the page)

- **Arrest location is free text** typed by ICE staff. Spellings vary; names are
  normalized and geocoded to a town centroid. ~91% of stays land on the map; the
  rest have a blank or unrecognised city.
- **The arrest is not where the person lives.** A custodial arrest — the majority
  here — is recorded at the jail or state prison holding the person, via the
  Criminal Alien Program.
- A stay counts if *any* stint touched a Nebraska facility, so some arrests and
  some "moved next" facilities are in neighbouring states.
- Counts are **detention stays, not unique people**. A small number of people
  appear twice after a release and re-arrest.
- The DDP "detention stays" dataset (no arrest join) has ~7,800 matching stays;
  this joined dataset has ~7,400 because it requires a matched arrest record.
