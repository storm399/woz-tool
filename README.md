# WOZ-waardeloket tool

Webtool die WOZ-waarden ophaalt via de publieke Kadaster-API.

- Eén adres opzoeken met autocomplete (PDOK Locatieserver).
- Excel uploaden met adressen, ingevuld Excel terugkrijgen.

## Lokaal draaien

```bash
pip install -r requirements.txt
python3 app.py
```

Open <http://127.0.0.1:5005>.

## Deploy op Render.com

Deze repo bevat een `render.yaml` voor één-klik deploy.

1. Maak een gratis account op [render.com](https://render.com).
2. **New +** → **Blueprint** → koppel deze GitHub-repo.
3. Render leest `render.yaml` en zet de service op (free tier).
4. Na ~3 minuten staat de tool op `https://woz-tool-xxxx.onrender.com`.

Free tier slaapt na 15 min inactiviteit; eerste request daarna duurt ~30s.

## API-limieten

WOZ-waardeloket: 60 requests/minuut, 5000/dag per IP.
