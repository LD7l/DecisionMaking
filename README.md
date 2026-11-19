# Case E – Shift scheduling (groep 11)

Decision Making for Business – Sint-Maria Olivia Hospital, tactische fase.
Gebaseerd op de C++-template van de prof (`src/main.cpp`, origineel in `examples/Template_Code.zip`).

## Structuur
```
data/      Excel van groep 11 (bron voor de invoerbestanden)
files/     invoer- en uitvoer-txt per dienst (pad zoals de template verwacht)
src/       main.cpp
results/   output van experimenten (λ-grid, sensitiviteit)
```

## Bouwen en draaien (Windows, VS Code)
1. Compiler: installeer MSYS2 → `pacman -S mingw-w64-ucrt-x86_64-gcc` → voeg `C:\msys64\ucrt64\bin` toe aan PATH.
   Controle in een nieuwe terminal: `g++ --version`.
2. VS Code-extensie: *C/C++* (Microsoft).
3. `Ctrl+Shift+B` bouwt `case_e.exe`. Draaien: Terminal → Run Task → *Run case_e*.
   Het programma moet starten vanuit de hoofdmap, anders vindt het `files/` niet.

## Controle
Met de voorbeeldbestanden in `files/` moet `files/Violations_dpt_A.txt` een
voorkeurscore van 25928 geven.
