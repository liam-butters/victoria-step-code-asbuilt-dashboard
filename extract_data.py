"""
Rebuilds dashboard/data.json from the "Reports" tab of "step_code_reports.xlsx"
(the sibling folder's as-built step code workbook for Victoria).

This is a separate, much smaller dataset than the NRCan Database Analysis
dashboard (bc-energy-step-code-dashboard). It exists to check whether that
dashboard, filtered to Victoria, gives a representative view of this
independently-sourced as-built data. See method.html for the column mapping
and caveats.

Output is aggregated into (municipality, year) cells - counts, sums, and
category histograms only, no per-home records - matching the main
dashboard's own privacy approach.

Usage:
    python extract_data.py ["path to step_code_reports.xlsx"]
"""
import datetime
import json
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

SHEET_NAME = 'Reports'
DEFAULT_WORKBOOK = Path(__file__).parent.parent / 'step_code_reports.xlsx'
OUT_PATH = Path(__file__).parent / 'data.json'

# Only as-built reports have the "As-Built" columns populated; the handful of
# "Pre-Construction" rows in this workbook are projects not yet built.
REPORT_STAGE_FIELD = 'Report Stage'
WANTED_REPORT_STAGE = 'As-Built'

MUNICIPALITY_FIELD = 'Municipality / District'
YEAR_FIELD = 'Year'
BUILDING_TYPE_FIELD = 'Building Type'
STEP_ACHIEVED_FIELD = 'Step As-Built'
FLOOR_AREA_FIELD = 'Conditioned Floor Area (m2)'
# A row can represent a whole block of identical attached units (e.g. a row
# of townhomes modeled once in HOT2000 and applied to all of them) rather
# than a single house - see Dwelling Units below.
DWELLING_UNITS_FIELD = 'Dwelling Units'

# Conditioned Floor Area and GHGs are reported as totals for the whole row
# (i.e. for all of its Dwelling Units combined), not per-home values - e.g. a
# 6-unit rowhouse block's floor area is the sum of all 6 units' areas, and
# its GHGs are the combined total for all 6. Every other metric (MEUI, TEDI,
# ACH, NLA, NLR, GHG intensity, % improvements) is already a per-home
# intensity/ratio, since it comes from the single HOT2000 model that stands
# for every identical unit in the block.
BLOCK_TOTAL_KEYS = {'floor_area', 'ghg'}

# Numeric fields taken directly from an "As-Built" column, mapped onto the
# same metric keys the main NRCan dashboard uses (see method.html).
NUMERIC_FIELDS = {
    'MEUI As-Built': 'meui',
    'TEDI As-Built': 'tedi',
    'ACH@50Pa As-Built': 'ach',
    'NLA@10Pa As-Built': 'nla',
    'NLR As-Built': 'nlr',
    FLOOR_AREA_FIELD: 'floor_area',
    # Peak Thermal Load (W/m2) is already floor-area-normalized, unlike the
    # main dashboard's raw EGHDESHTLOSS - this is the closest as-built
    # equivalent to that dashboard's "design heat loss per floor area".
    'F PTL As-Built (W/m2)': 'design_heat_loss_per_area',
}

# GHGs are *not* taken from "ZC Total GHG Result" - that column is unreliable
# for reports pulled from the Zero Carbon framework (it has produced values
# three orders of magnitude too high on some rows; see git history). Instead,
# GHGs are calculated directly from this workbook's as-built fuel
# consumption, the same way the underlying compliance reports do it:
# GHG (kg CO2e/yr) = Electricity (GJ/yr) x ELEC_GHG_FACTOR + Natural Gas (GJ/yr) x GAS_GHG_FACTOR.
ELEC_FIELD = 'F Electricity As-Built (GJ/yr)'
GAS_FIELD = 'F Natural Gas As-Built (GJ/yr)'
ELEC_GHG_FACTOR = 3.06   # kg CO2e per GJ of electricity (BC Hydro grid intensity)
GAS_GHG_FACTOR = 50      # kg CO2e per GJ of natural gas

MIN_PLAUSIBLE_YEAR = 1990
MAX_PLAUSIBLE_YEAR = datetime.date.today().year + 1

MUNICIPALITY_FIXES = {
    'victoria': 'Victoria',
    'city of victoria': 'Victoria',
    'victoira': 'Victoria',
    'vcitora': 'Victoria',
    'victoria /victoria': 'Victoria',
    'victoria / south vancouver island': 'Victoria',
    'oak bay': 'Oak Bay',
    'saanich': 'Saanich',
    'nanaimo': 'Nanaimo',
}

BUILDING_TYPE_FIXES = {
    'single detached': 'Single Detached',
    'single detached w/ secondary suite': 'Single Detached w/Secondary Suite',
    'laneway house': 'Laneway House',
    'duplex)': 'Double/Semi-detached (non-MURB)',
    'select one': 'Unknown',
    'detached': 'Single Detached',
}

# A handful of Building Type values end in a stray U+20AC ("€") - almost
# certainly a mis-decoded character from how this workbook's data was
# originally entered/exported - with no other visible effect. Strip it
# before applying the fixes above.
STRAY_CHARS = '€'


def to_float(raw):
    if raw is None or raw == '':
        return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        return None


def to_float_fuel(raw):
    # This workbook's fuel-breakdown ("F ...") columns use "-" for "no
    # consumption of this fuel", distinct from a genuinely blank/missing row
    # (where every "F ..." column is blank, meaning the compliance report had
    # no fuel breakdown at all).
    if isinstance(raw, str) and raw.strip() == '-':
        return 0.0
    return to_float(raw)


def normalize_municipality(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(str(raw).strip().split())
    if not name:
        return 'Unknown'
    return MUNICIPALITY_FIXES.get(name.lower(), name)


def normalize_building_type(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(str(raw).strip().split())
    name = name.rstrip(STRAY_CHARS).strip()
    if not name:
        return 'Unknown'
    return BUILDING_TYPE_FIXES.get(name.lower(), name)


def normalize_year(raw):
    y = to_float(raw)
    if y is None:
        return None
    y = int(y)
    if y < MIN_PLAUSIBLE_YEAR or y > MAX_PLAUSIBLE_YEAR:
        return None
    return y


def normalize_step_achieved(raw):
    v = to_float(raw)
    if v is None:
        return 'Unknown'
    return str(int(v)) if v == int(v) else str(v)


def load_workbook_with_lock_retry(path):
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / f'_{path.name}'
        shutil.copy(path, tmp)
        try:
            return openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        finally:
            tmp.unlink(missing_ok=True)


def main():
    wb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK
    print(f'Reading {wb_path.name} ...')
    wb = load_workbook_with_lock_retry(wb_path)
    ws = wb[SHEET_NAME]

    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    col = {name: i for i, name in enumerate(header) if name}

    missing = [n for n in [REPORT_STAGE_FIELD, MUNICIPALITY_FIELD, YEAR_FIELD, BUILDING_TYPE_FIELD,
                            STEP_ACHIEVED_FIELD, ELEC_FIELD, GAS_FIELD, DWELLING_UNITS_FIELD]
               + list(NUMERIC_FIELDS) if n not in col]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')

    # Aggregated into (municipality, year) cells - counts/sums/category
    # histograms only, no per-home records - matching the main NRCan
    # dashboard's own privacy approach of shipping no record-level data.
    cells = {}
    total_in_sheet = 0
    total_kept = 0
    skipped_not_as_built = 0
    muni_counts = {}

    for r in rows:
        if r is None or all(v is None for v in r):
            continue
        total_in_sheet += 1

        def get(name):
            return r[col[name]]

        stage = get(REPORT_STAGE_FIELD)
        stage = stage.strip() if isinstance(stage, str) else stage
        if stage != WANTED_REPORT_STAGE:
            skipped_not_as_built += 1
            continue

        municipality = normalize_municipality(get(MUNICIPALITY_FIELD))
        year = normalize_year(get(YEAR_FIELD))
        year_key = year if year is not None else 'Unknown'
        building_type = normalize_building_type(get(BUILDING_TYPE_FIELD))
        step_achieved = normalize_step_achieved(get(STEP_ACHIEVED_FIELD))
        # Matches the main dashboard's two-bucket BUILDINGTYPE convention
        # ("House" vs "Multi-unit: whole building"), so the two datasets'
        # building-type mixes can be compared directly despite this sheet's
        # much more granular labels.
        house_type_group = 'Multi-unit: whole building' if 'MURB' in building_type.upper() else 'House'

        # Default to 1 unit when Dwelling Units is blank/invalid.
        units = to_float(get(DWELLING_UNITS_FIELD))
        if not units or units < 1:
            units = 1.0

        muni_counts[municipality] = muni_counts.get(municipality, 0) + units
        total_kept += 1

        num = {}
        for field, key in NUMERIC_FIELDS.items():
            fv = to_float(get(field))
            if fv is not None:
                num[key] = fv

        floor_area = num.get('floor_area')
        elec = to_float_fuel(get(ELEC_FIELD))
        gas = to_float_fuel(get(GAS_FIELD))
        if elec is not None or gas is not None:
            ghg = (elec or 0.0) * ELEC_GHG_FACTOR + (gas or 0.0) * GAS_GHG_FACTOR
            if floor_area:
                num['ghg'] = ghg
                num['ghg_intensity'] = ghg / floor_area

        cell = cells.setdefault((municipality, year_key), {
            'municipality': municipality,
            'year': year_key,
            'count': 0,
            'num': {},
            'cat': {'step_achieved': {}, 'house_type_group': {}, 'building_type': {}},
        })
        # 'count' and every metric are weighted by Dwelling Units, to match
        # the NRCan baseline's convention of counting individual homes. For
        # the two BLOCK_TOTAL_KEYS, the row's value is already a total across
        # `units` homes, so only the denominator is weighted (giving the
        # correct per-home average); every other metric is already a
        # per-home value, so both its sum and denominator are weighted.
        cell['count'] += units
        for key, v in num.items():
            pair = cell['num'].setdefault(key, [0.0, 0])
            pair[0] += v if key in BLOCK_TOTAL_KEYS else v * units
            pair[1] += units
        for cat_key, value in (('step_achieved', step_achieved), ('house_type_group', house_type_group), ('building_type', building_type)):
            hist = cell['cat'][cat_key]
            hist[value] = hist.get(value, 0) + units

    print(f'{total_in_sheet} data rows in sheet, {skipped_not_as_built} skipped (not "{WANTED_REPORT_STAGE}" report stage), '
          f'{total_kept} kept')
    print('Municipalities:', dict(sorted(muni_counts.items(), key=lambda kv: -kv[1])))

    out = {
        'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'source_file': wb_path.name,
        'total_rows': total_kept,
        'cells': list(cells.values()),
    }
    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {OUT_PATH} ({OUT_PATH.stat().st_size / 1000:.1f} KB)')


if __name__ == '__main__':
    main()
