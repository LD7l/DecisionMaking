# Invoerbestanden (formaat zoals main.cpp ze leest)

Shiftcodes in onze bestanden: 0 = vrij, 1 = Early (06:00), 2 = Late (15:00), 3 = Night (21:00).

## Shift_system_dpt_X.txt
```
3   9        aantal shiften, lengte (uur)
6            starturen, één per regel
15
21
10           minimale bezetting per shift, zelfde volgorde (E, L, N)
11
7
```
Dienst A: bezetting = Case C-eis ZONDER 25%-buffer (10/11/7). De buffer zit al in het
aantal lijnen (Case D); verlof wordt in het maandrooster expliciet gemodelleerd.
Variant met buffer (13/14/9) = sensitiviteitsanalyse.

## Constraints_dpt_X.txt
```
0   20       min / max aantal shiften over 28 dagen (main.cpp vermenigvuldigt met tewerkstelling)
1   6        min / max opeenvolgende werkdagen
1   6        min / max opeenvolgende dagen zelfde shift: E
1   6                                                     L
1   6                                                     N
0   20       min / max aantal shiften per type: E
0   20                                             L
0   20                                             N
NO           identieke weekends (Case D: non-identical)
```
Casetekst p. 25: max 6 werkdagen per week, max 1 shift per dag, geen N -> E/L.
"Max 6 opeenvolgend" is strenger dan "max 6 per week" (dat laat 12 op rij toe over twee
weken heen); de template kent enkel opeenvolgende dagen. 1 shift/dag zit in het formaat zelf.
N -> E/L en type I-minimum controleert de template niet: dat moet in onze eigen evaluatie.

### Shiftopvolging (zelf implementeren in procedure() en evaluatie)
Verboden overgangen (shift vandaag -> shift morgen):
- N -> E (0u rust) en N -> L (9u rust): casetekst p. 25
- L -> E (6u rust): eigen toevoeging, 11u-rustregel (verantwoorden in rapport; sensitiviteit: met/zonder)
Het formaat van Constraints kan dit niet uitdrukken. Het cyclische rooster bevat deze overgangen
nooit (elke Case D-lijn = vaste shift); ze kunnen enkel ontstaan in het maandrooster.

## Personnel_dpt_X.txt
Per nurse één regel: id, 28 dagen x (E, D, L, N, F) voorkeur (0-9, 100 = verlof), tewerkstelling, type (1/2).

## Cyclic_roster_dpt_X.txt
Eerste regel: aantal nurses. Daarna per nurse 28 shiftcodes (0-3). Volgorde = volgorde in Personnel.

Originele voorbeeldbestanden van de prof: `examples/prof_files/`.
