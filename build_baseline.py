"""
Computes the comparison baseline for this dashboard: the main BC Energy Step
Code Dashboard (NRCan Database Analysis), filtered to Victoria, collapsed to
the metrics this smaller as-built dataset can also provide.

Reads the main dashboard's already-built data.json (run its own
extract_data.py first if it needs refreshing) rather than the workbook, since
that cube is already the aggregated source of truth.

Builds both year modes (year built / entry date) rather than picking one, so
the dashboard can offer a live dropdown between them instead of baking in a
choice here - and keeps every metric broken out by year (not just an overall
average), so the dashboard can match the NRCan comparison to whichever years
this spreadsheet's own data actually covers, rather than comparing against
every year NRCan has. Also carries each year's raw per-home values (from the
main dashboard's own 'median_export', scoped to just Victoria/Saanich - see
its extract_data.py) so a median can be computed client-side alongside the
mean.

Usage:
    python build_baseline.py ["path to main dashboard's data.json"]
"""
import json
import sys
from pathlib import Path

DEFAULT_MAIN_DATA = (Path(__file__).parent.parent.parent.parent.parent
                      / '2025-16 - NRCan Database Analysis' / 'dashboard' / 'data.json')
OUT_PATH = Path(__file__).parent / 'baseline.json'

CITY = 'Victoria'
YEAR_MODES = ('built', 'entry')

# Metric keys shared with this dashboard's extract_data.py (see method.html).
NUMERIC_KEYS = ['meui', 'tedi', 'ach', 'nla', 'nlr', 'floor_area', 'ghg', 'ghg_intensity', 'design_heat_loss_per_area']
CATEGORY_KEYS = {'step_achieved': 'step_achieved', 'building_type': 'building_type', 'el_achieved': 'el_achieved'}


def build_mode(data, mode):
    cells = [c for c in data['year_modes'][mode]['cells'] if c['city'] == CITY]

    total_count = sum(c['count'] for c in cells)
    sums = {k: [0.0, 0] for k in NUMERIC_KEYS}
    cats = {k: {} for k in CATEGORY_KEYS}
    by_year = {}

    for cell in cells:
        for k in NUMERIC_KEYS:
            v = cell['num'].get(k)
            if v:
                sums[k][0] += v[0]
                sums[k][1] += v[1]
        for out_key, cell_key in CATEGORY_KEYS.items():
            hist = cell['cat'].get(cell_key)
            if hist:
                for cat, n in hist.items():
                    cats[out_key][cat] = cats[out_key].get(cat, 0) + n
        year = cell['year']
        if isinstance(year, int):
            acc = by_year.setdefault(year, {'count': 0, 'num': {k: [0.0, 0] for k in NUMERIC_KEYS}})
            acc['count'] += cell['count']
            for k in NUMERIC_KEYS:
                v = cell['num'].get(k)
                if v:
                    acc['num'][k][0] += v[0]
                    acc['num'][k][1] += v[1]

    overall_avg = {k: (sums[k][0] / sums[k][1] if sums[k][1] else None) for k in NUMERIC_KEYS}
    overall_n = {k: sums[k][1] for k in NUMERIC_KEYS}

    median_raw = data.get('median_export', {}).get(mode, {})
    years_out = []
    for year in sorted(by_year):
        acc = by_year[year]
        row = {'year': year, 'count': acc['count'], 'num': acc['num']}
        median = median_raw.get(f'{CITY}|{year}')
        if median:
            row['median'] = median
        years_out.append(row)

    return {
        'count': total_count,
        'avg': overall_avg,
        'n': overall_n,
        'cat': cats,
        'years': years_out,
    }


def main():
    main_data_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_MAIN_DATA
    print(f'Reading {main_data_path} ...')
    data = json.loads(main_data_path.read_text(encoding='utf-8'))

    year_modes_out = {mode: build_mode(data, mode) for mode in YEAR_MODES}

    out = {
        'generated_at': data.get('generated_at'),
        'source_file': data.get('source_file'),
        'city': CITY,
        'year_mode_default': 'built',
        'year_modes': year_modes_out,
    }
    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    for mode in YEAR_MODES:
        m = year_modes_out[mode]
        print(f"{mode}: {m['count']} Victoria homes, {len(m['years'])} years, avg {m['avg']}")


if __name__ == '__main__':
    main()
