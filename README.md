# Facturieren

Kleine webapp voor de maandelijkse facturatie op basis van de "VMs export 01-MM-JJJJ.xlsx".

- **Upload** de maandexport via de webpagina (slepen of kiezen), als kale `.csv`
  (bijv. `Server_Chargeback_V3-2026-10-01.csv`) of als `.xlsx`. De maand komt uit de bestandsnaam.
  Exports worden in de container bewaard (`/data`); de app kijkt zelf nergens naar een map.
- **Klanten** leg je één keer vast met een filter op kolom C (`klant`), bijv. `Stadsbeheer` of
  `Paralax/Rostarcas`. Het filter matcht op hele mapnamen, dus `GGZ` telt `GGZ GV` niet mee.
- **Overzicht** per klant: aantal VM's, CPU's, memory, tier1-single/dual, tier3-dual, mgm en totale storage,
  met het verschil t.o.v. de vorige geüploade maand. Teksten als `225.00` worden automatisch getallen
  (geen zoek-en-vervang meer nodig).
- **Wijzigingen** per klant: nieuwe VM's, verwijderde VM's en VM's met andere CPU/memory/storage.
- **Download xlsx** per maand: de export als nette xlsx (getallen als getal, opgemaakte kop, filter,
  vastgezette eerste rij), als `VMs export 01-MM-JJJJ.xlsx`. Een CSV kan ook los worden omgezet zonder op te slaan.
- **Download Excel** met het overzicht en per klant een tabblad met de VM's.
- "Niet toegewezen" toont wat in de export staat maar onder geen enkele klant valt.

## Starten

### Docker Swarm (stack)

Het image wordt door GitHub Actions gebouwd. `ghcr.io/richrdj/facturieren:latest` wijst altijd naar de
laatste release; elke release heeft ook een eigen tag, bijv. `:1.0.0`. Zie [Releases](https://github.com/RichrdJ/facturieren/releases).

```bash
docker stack deploy -c docker-stack.yml facturieren
```

Open daarna http://<swarm-host>:9898. De data staat in het volume `facturieren_data`.
Updaten naar de nieuwste versie:

```bash
docker service update --image ghcr.io/richrdj/facturieren:latest --force facturieren_app
```

Heb je meer dan één node, zet dan de `placement`-constraint in `docker-stack.yml` aan, zodat de
service op de node blijft waar het volume staat.

### Docker Compose (lokaal)

```bash
docker compose up -d --build
```

Open daarna http://localhost:9898. Data staat in `./data` (exports + `klanten.json`).

## Inloggen

Zet een wachtwoord met `APP_PASSWORD` in `docker-stack.yml` (of `docker-compose.yml`) en deploy opnieuw.
De app toont dan een inlogpagina.

- Een browser logt één keer in en blijft 30 dagen ingelogd. Uitloggen kan rechtsboven.
- Wachtwoord veranderd? Dan moet iedereen opnieuw inloggen.
- Na 5 verkeerde pogingen moet je even wachten, en steeds langer bij nieuwe fouten.
- Draait de app via https (bijv. achter een reverse proxy), zet dan ook `COOKIE_SECURE=true`.
- Zonder `APP_PASSWORD` werkt alles zonder inloggen.

## Nieuwe release maken

Schrijf de notes in `docs/releases/vX.Y.Z.md`, commit, en push een tag:

```bash
git tag v1.1.0 && git push origin v1.1.0
```

GitHub Actions bouwt dan het image (`:1.1.0`, `:1.1`, `:latest`) en maakt de release aan.

## Let op

Klantdata (`data/`, `*.xlsx`) staat in `.gitignore` en hoort niet in deze repo.
