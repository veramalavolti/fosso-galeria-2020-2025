from pathlib import Path
import json, datetime
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SOURCE = 'risultati FOSSO GALERIA 1 E 2 FINALI.xlsx'
NH4_AS_N = 14.0067/18.0385

# Parametri monitorati e rispettivi valori soglia, nell'ordine fornito.
# Chiave: nome nel file Excel → (nome visualizzato, soglia).
# label: soglia come indicata; converted: stessa soglia nell'unità dei dati (None se coincide);
# min/max: valori numerici nell'unità dei dati usati per il confronto; extra: soglie secondarie (solo grafiche);
# star: parametro contrassegnato con * nella tabella delle soglie.
PARAMS = {
    'OSSIGENO DISCIOLTO': ('OSSIGENO DISCIOLTO', dict(label='≥ 70% / ≥ 5 mg/L', min=5, star=True, note="L'archivio riporta solo la concentrazione in mg/L: il confronto automatico usa ≥ 5 mg/L. La saturazione ≥ 70% non è verificabile con questi dati.")),
    'ACIDO PERFLUOROTTANSULFONICO (PFOS)': ('ACIDO PERFLUOROOTTANSOLFONICO (PFOS)', dict(label='0,00065 µg/L', max=0.00065, star=True)),
    'AMPA': ('AMPA', dict(label='0,1 µg/L', max=0.1, star=True)),
    'Arsenico (Disciolto)': ('ARSENICO DISCIOLTO', dict(label='0,02 mg/L', converted='20 µg/L', max=20)),
    'AZOTO AMMONIACALE': ('AZOTO AMMONIACALE', dict(label='2 mg/L NH₄', converted='≈ 1,55 mg/L (N)', max=round(2*NH4_AS_N, 4), note='Soglia espressa come ione ammonio (NH₄), convertita in azoto (N = NH₄ × 14,007 / 18,039) per il confronto con i dati.')),
    'AZOTO NITRICO': ('AZOTO NITRICO', dict(label='20 mg/L (N)', max=20, star=True)),
    'BENZENE': ('BENZENE', dict(label='0,001 mg/L', converted='1 µg/L', max=1)),
    'BENZO [a] PIRENE': ('BENZO [a] PIRENE', dict(label='0,00001 mg/L', converted='0,01 µg/L', max=0.01)),
    'BOD5': ('BOD5', dict(label='20 mg/L (O₂)', max=20)),
    'Cadmio (Disciolto)': ('CADMIO DISCIOLTO', dict(label='0,005 mg/L', converted='5 µg/L', max=5)),
    'CIPERMETRINA': ('CIPERMETRINA', dict(label='0,0001 mg/L', converted='0,1 µg/L', max=0.1)),
    'COD': ('COD', dict(label='100 mg/L (O₂)', max=100)),
    'Cromo (Totale)': ('CROMO TOTALE', dict(label='0,1 mg/L', converted='100 µg/L', max=100)),
    "CONDUCIBILITA' ELETTRICA": ("CONDUCIBILITA' ELETTRICA", dict(label='3.000 µS/cm (2.500 per uso potabile)', max=3000, extra=[dict(label='2.500 · uso potabile', max=2500)])),
    'DDT TOTALE': ('DDT TOTALE', dict(label='0,0001 mg/L', converted='0,1 µg/L', max=0.1)),
    'DOC - Carbonio Organico Disciolto': ('DOC (CARBONIO ORGANICO DISCIOLTO)', dict(label='Senza variazioni anomale', star=True, note='Soglia qualitativa: non è previsto un confronto numerico automatico. Valuta l’andamento nel grafico.')),
    'DUREZZA': ('DUREZZA', dict(label='15–50 °F (consigliato)', converted='150–500 mg/L (CaCO₃)', min=150, max=500, note='Intervallo consigliato. Conversione: 1 °F = 10 mg/L di CaCO₃.')),
    'ESCHERICHIA COLI': ('ESCHERICHIA COLI', dict(label='10 UFC/100 mL (max 100)', max=10, extra=[dict(label='100 · valore massimo', max=100)])),
    'FOSFORO TOTALE': ('FOSFORO TOTALE', dict(label='2 mg/L (P)', converted='2.000 µg/L (P)', max=2000)),
    'GLIFOSATE': ('GLIFOSATE', dict(label='0,05 mg/L', converted='50 µg/L', max=50)),
    'IDROCARBURI POLICICLICI AROMATICI': ('IDROCARBURI POLICICLICI AROMATICI', dict(label='0,01 mg/L', converted='10 µg/L', max=10)),
    'Mercurio (Disciolto)': ('MERCURIO DISCIOLTO', dict(label='0,001 mg/L', converted='1 µg/L', max=1)),
    'Nichel (Disciolto)': ('NICHEL DISCIOLTO', dict(label='0,2 mg/L', converted='200 µg/L', max=200)),
    'ORTOFOSFATO': ('ORTOFOSFATO', dict(label='0,5 mg/L (P)', converted='500 µg/L (P)', max=500, star=True)),
    'pH': ('pH', dict(label='6 – 9,5 (6,5 – 9,5 per uso potabile)', min=6, max=9.5, extra=[dict(label='6,5 · uso potabile', min=6.5)])),
    'Piombo (Disciolto)': ('PIOMBO DISCIOLTO', dict(label='0,1 mg/L', converted='100 µg/L', max=100)),
    "SALINITA'": ("SALINITA'", dict(label='1.500 mg/L', converted='≈ 1,5 PSU', max=1.5, star=True, note='I dati sono espressi in PSU: conversione approssimata 1 PSU ≈ 1 g/L (1.000 mg/L) di sali disciolti.')),
    'TEMPERATURA ACQUA': ('TEMPERATURA ACQUA', dict(label='≤ 30 °C / Δ < 3 °C', max=30, star=True, note='Il confronto automatico usa ≤ 30 °C; il criterio Δ < 3 °C non è valutato automaticamente.')),
    'TETRACLOROETILENE': ('TETRACLOROETILENE', dict(label='0,01 mg/L', converted='10 µg/L', max=10)),
    'TERBUTILAZINA': ('TERBUTILAZINA', dict(label='0,05 mg/L', converted='50 µg/L', max=50)),
}

# Categorie usate per raggruppare i parametri nella dashboard.
GROUPS = {
    'Parametri chimico-fisici': ['OSSIGENO DISCIOLTO', 'pH', 'TEMPERATURA ACQUA', "CONDUCIBILITA' ELETTRICA", "SALINITA'", 'DUREZZA'],
    'Carico organico e nutrienti': ['BOD5', 'COD', 'DOC (CARBONIO ORGANICO DISCIOLTO)', 'AZOTO AMMONIACALE', 'AZOTO NITRICO', 'FOSFORO TOTALE', 'ORTOFOSFATO'],
    'Metalli': ['ARSENICO DISCIOLTO', 'CADMIO DISCIOLTO', 'CROMO TOTALE', 'MERCURIO DISCIOLTO', 'NICHEL DISCIOLTO', 'PIOMBO DISCIOLTO'],
    'Microinquinanti organici': ['ACIDO PERFLUOROOTTANSOLFONICO (PFOS)', 'BENZENE', 'BENZO [a] PIRENE', 'IDROCARBURI POLICICLICI AROMATICI', 'TETRACLOROETILENE'],
    'Pesticidi': ['AMPA', 'CIPERMETRINA', 'DDT TOTALE', 'GLIFOSATE', 'TERBUTILAZINA'],
    'Microbiologia': ['ESCHERICHIA COLI'],
}
GROUP_OF = {name: group for group, names in GROUPS.items() for name in names}
assert sorted(GROUP_OF) == sorted(name for name, _ in PARAMS.values()), 'Every parameter needs exactly one group'

ws = openpyxl.load_workbook(ROOT/'data'/SOURCE, read_only=True, data_only=True).active
records = []
for rownum, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
    if not any(v is not None for v in r): continue
    sample, date, param, unit, sign, value, revision, duration, code, station = r
    value = float(str(value).replace(',', '.'))
    assert sign in ('<', '=') and code in ('F4.18/0', 'F4.79/0') and param in PARAMS, (rownum, param)
    date = str(date)
    datetime.date.fromisoformat(date[:10])
    records.append(dict(sample=str(sample), date=date[:10], timestamp=date, param=PARAMS[param][0], original=param, unit=unit, sign=sign, value=value, station='S1' if code == 'F4.79/0' else 'S2', code=code, source=SOURCE, row=rownum, revision=revision))
records.sort(key=lambda r: (r['date'], r['station'], r['param']))
keys = [(r['sample'], r['code'], r['param'], r['revision']) for r in records]
assert len(keys) == len(set(keys)), 'Duplicate records require explicit handling'
assert {r['original'] for r in records} == set(PARAMS), 'Every listed parameter must be present in the source'
thresholds = {name: dict(n=n, group=GROUP_OF[name], **t) for n, (name, t) in enumerate(PARAMS.values(), 1)}
meta = dict(measurements=len(records), parameters=len(set(r['param'] for r in records)), samples=len(set((r['sample'], r['code']) for r in records)), below=sum(r['sign'] == '<' for r in records), start=min(r['date'] for r in records), end=max(r['date'] for r in records), sources=[dict(file=SOURCE, rows=len(records))], stations={s: dict(count=sum(r['station'] == s for r in records), start=min(r['date'] for r in records if r['station'] == s), end=max(r['date'] for r in records if r['station'] == s)) for s in ['S1', 'S2']})
(ROOT/'dist/data.json').write_text(json.dumps(dict(meta=meta, groups=list(GROUPS), thresholds=thresholds, records=records), ensure_ascii=False, separators=(',', ':')))
print(json.dumps(meta, ensure_ascii=False, indent=2))
