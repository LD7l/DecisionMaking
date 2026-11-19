# DecisionMaking

Dit project bevat een Python-versie van het voorbeeld voor het opstellen en
beoordelen van een verpleegkundig maandrooster. Er zijn geen externe Python-
packages nodig.

## Uitvoeren

Zet de invoerbestanden in de map `files` en voer vanuit de projectmap uit:

```powershell
python main.py
```

Standaard wordt afdeling A gebruikt. Een andere afdeling of invoermap kan
worden opgegeven:

```powershell
python main.py --department B
python main.py --department C --data-dir data
```

De invoermap moet voor de gekozen afdeling deze bestanden bevatten:

- `Shift_system_dpt_<afdeling>.txt`
- `Cyclic_roster_dpt_<afdeling>.txt`
- `Personnel_dpt_<afdeling>.txt`
- `Constraints_dpt_<afdeling>.txt`

Het rooster gebruikt 28 dagen. De bestanden volgen dezelfde whitespace-
gescheiden gegevensindeling als het oorspronkelijke C++-voorbeeld. Als de
bestanden ontbreken of onvolledig zijn, meldt het programma welk bestand of
gegeven nodig is.

Het programma schrijft `Monthly_Roster_dpt_<afdeling>.txt` en
`Violations_dpt_<afdeling>.txt` in de invoermap. Het rooster wordt, net als in
het oorspronkelijke voorbeeld, gelijkgesteld aan het cyclische rooster.
