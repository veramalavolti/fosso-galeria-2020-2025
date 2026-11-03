# fosso-galeria-2020-2025
Dashboard A SUD per il monitoraggio delle acque del Fosso Galeria · dati 2020–2025

Dati: `data/risultati FOSSO GALERIA 1 E 2 FINALI.xlsx` (2.077 misurazioni, 30 parametri, 103 campioni).
I valori soglia dei 30 parametri sono definiti in `scripts/import_data.py`, che rigenera `dist/data.json`:

    python3 scripts/import_data.py   # richiede openpyxl
    npm test
