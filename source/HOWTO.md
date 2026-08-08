# Akými príkazmi sa zo zdrojových .qmd súborov vytvorí náhľad?

V koreňovom priečinku projektu stačí spustiť:
`quarto preview`

Otvorí náhľad v prehliadači a pri uložení zmien v .qmd súboroch ho automaticky obnoví.

Ak chceš iba jednorazovo vytvoriť statický náhľad do priečinka .preview:
`quarto render --to html --output-dir .preview`

Potom otvor:
`Start-Process .preview\index.html`
