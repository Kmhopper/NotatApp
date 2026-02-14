# Notat Overlay (Windows)

Liten notatapp for rask notering over andre vinduer.

## Funksjoner
- Global hurtigtast `|` for vis/skjul av notatvindu.
- Gjennomsiktig og alltid øverst (`always-on-top`) notatflate.
- Autosave til `data/session.json`.
- Auto-fangst av tekst fra clipboard (marker tekst + `Ctrl + C`).
- Auto-fangst bevarer fet skrift når clipboard inneholder rik tekst (HTML/RTF fra Word/nettleser).
- Fet tekst (bold) med knapp eller `Ctrl + B` (toggle skrivemodus uten markering).
- Stavekontroll (norsk + engelsk) med markering av ord som ser feil ut.
- Bilde fra utklippstavle:
  - ta skjermklipp med `Win + Shift + S`
  - lim inn i notatet med `Ctrl + V`
- Eksport til `Word (.docx)` og `PDF`.

## Kom i gang
### Enklest (anbefalt)
1. Kjør `setup.bat` (installerer alt som trengs i `.venv`).
2. Kjør `run.bat` for å starte appen.

### Manuelt (PowerShell)
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Bruk
- Skriv notater i tekstfeltet.
- Trykk `Auto-fangst: PÅ` for å fange markert tekst fra andre vinduer via `Ctrl + C`.
- Marker tekst og bruk `Bold`/`Ctrl + B` for fet skrift.
- Uten markering toggler `Ctrl + B` bold skrivemodus for tekst du skriver videre.
- Stavekontroll kan slås av/på med knappen `Stavekontroll: PÅ/AV`.
- Sett inn skjermklipp med `Ctrl + V` (lagres som bilde og settes inn som token i teksten).
- Eksporter med knappene `Eksporter Word` eller `Eksporter PDF`.
- Skjul/vise appen med `|`.

## Viktig
- Denne versjonen er laget for Windows.
- Hvis hurtigtasten ikke virker, brukes den trolig av en annen app.
