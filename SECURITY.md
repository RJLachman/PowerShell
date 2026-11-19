# Security-beleid

## Ondersteunde versies

Alleen de laatste versie op de branch `main` (en de meest recente release) wordt ondersteund.

| Versie | Ondersteund |
|---|---|
| 1.x | ✅ |

## Een kwetsbaarheid melden

Meld beveiligingsproblemen **niet** via een publieke issue.

1. Ga naar het tabblad **Security** van deze repository.
2. Kies **Advisories** → **Report a vulnerability** (GitHub private vulnerability reporting).
3. Beschrijf het probleem, de betrokken scripts en, indien mogelijk, stappen om het te reproduceren.

Je ontvangt zo snel mogelijk een reactie. Na een fix wordt het advisory gepubliceerd en in de [CHANGELOG](CHANGELOG.md) vermeld.

## Geen secrets in deze repository

- De scripts bevatten **geen** wachtwoorden, tokens, certificaten of andere secrets.
- Authenticatie verloopt altijd interactief (bijv. `Connect-MgGraph`, `Connect-IPPSSession`) of via de context van de ingelogde beheerder.
- Organisatiespecifieke waarden zijn vervangen door voorbeeldwaarden zoals `contoso.nl` en `contoso.local`.

Zie je toch iets dat op een secret of gevoelige informatie lijkt? Meld het dan via de stappen hierboven.
