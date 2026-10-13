# Facturieren

Kleine webapp voor de maandelijkse facturatie op basis van de "VMs export 01-MM-JJJJ.xlsx".

- **Upload** de maandexport via de webpagina (slepen of kiezen). De maand komt uit de bestandsnaam.
  Exports worden in de container bewaard (`/data`); de app kijkt zelf nergens naar een map.
- **Klanten** leg je één keer vast met een filter op kolom C (`klant`), bijv. `Stadsbeheer` of
  `Paralax/Rostarcas`. Het filter matcht op hele mapnamen, dus `GGZ` telt `GGZ GV` niet mee.
- **Overzicht** per klant: aantal VM's, CPU's, memory, tier1-single/dual, tier3-dual, mgm en totale storage,
  met het verschil t.o.v. de vorige geüploade maand. Teksten als `225.00` worden automatisch getallen
  (geen zoek-en-vervang meer nodig).
- **Wijzigingen** per klant: nieuwe VM's, verwijderde VM's en VM's met andere CPU/memory/storage.
- **Download Excel** met het overzicht en per klant een tabblad met de VM's.
- "Niet toegewezen" toont wat in de export staat maar onder geen enkele klant valt.

## Starten

```bash
docker compose up -d --build
```

Open daarna http://localhost:8000. Data staat in `./data` (exports + `klanten.json`).

Zonder compose:

```bash
docker build -t facturieren .
docker run -d -p 8000:8000 -v facturieren-data:/data --name facturieren facturieren
```

### Wachtwoord

Zet `APP_PASSWORD` (en eventueel `APP_USER`, standaard `admin`) om de app achter basic auth te zetten.
Doe dit zeker als de container niet alleen op je eigen machine draait.

## Let op

Klantdata (`data/`, `*.xlsx`) staat in `.gitignore` en hoort niet in deze repo.
