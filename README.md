# Foto do znalostí

PWA pro iPhone: vyfotí dokument, vizitku, ceduli nebo tabuli a odešle ji přímo do OneDrive
(`Znalosti/Inbox`). Do Fotek iPhonu se nic neukládá. Fotka čeká jen v úložišti aplikace,
dokud ji OneDrive nepřijme, pak se smaže. Na PC ji Claude Code zpracuje do Markdown
znalostní báze `C:\Data\OneDrive\Znalosti` (pravidla: `Znalosti\CLAUDE.md`), u vizitek vytvoří `.vcf`.

Aplikace nemá žádný server. Přihlášení k Microsoftu (OAuth + PKCE) i nahrávání přes Microsoft Graph
běží přímo v telefonu.

## Nasazení (jednorázově)

### 1. Registrace aplikace v Microsoft Entra
1. https://entra.microsoft.com → **App registrations** → **New registration**
   (účet, pod kterým máte Azure – stejně jako u MOPANI).
2. Name: `Foto do znalosti`. Supported account types: **Personal Microsoft accounts only**.
3. Redirect URI: platforma **Single-page application (SPA)**,
   URI `https://vagon74.github.io/foto-znalosti/` (včetně lomítka na konci).
   Pro místní test můžete přidat i `http://localhost:5520/`.
4. Po vytvoření zkopírujte **Application (client) ID** a vložte ho do `config.js` (`clientId`).
5. Žádný client secret ani další oprávnění nejsou potřeba (`Files.ReadWrite` se odsouhlasí při prvním přihlášení).

### 2. GitHub Pages
Repozitář `vagon74/foto-znalosti` (veřejný, Pages z větve `main`, kořen). Neobsahuje žádná tajemství.

### 3. iPhone
Safari → `https://vagon74.github.io/foto-znalosti/` → Přihlásit → Sdílet → **Přidat na plochu**.
Dál spouštějte jen z ikony na ploše.

## Používání
- Zvolte typ (nebo nechte *Automaticky*) → **Vyfotit** → případně **Další strana** → poznámka → **Odeslat**.
- Bez signálu se fotky uloží do fronty v aplikaci a odešlou se samy po připojení
  (nebo po dalším otevření aplikace).
- Přihlášení vydrží cca 24 h (omezení Microsoftu pro aplikace bez serveru). Pak stačí ťuknout na
  **Přihlásit**; neodeslané fotky mezitím čekají ve frontě.

## Zpracování na PC
V Claude Code otevřete složku `C:\Data\OneDrive\Znalosti` a napište „zpracuj Inbox“,
nebo nechte běžet naplánovanou úlohu. Výstup: `Poznamky/`, `INDEX.md`, `Lhuty.md`, `Kontakty/*.vcf`.

**Vizitka do iPhonu:** v aplikaci OneDrive otevřete `Znalosti/Kontakty/<jméno>.vcf` → Sdílet → Kontakty
(nebo `vse-kontakty.vcf` pro všechny najednou).

## Vývoj
`npx serve foto-znalosti -l 5520` (konfigurace `foto-znalosti` v `.claude/launch.json`);
`http://localhost:5520/?demo` = náhled bez přihlášení, nahrávání se jen simuluje.
Při změně souborů zvyšte `VERSION` v `app.js` a `CACHE` v `sw.js`.
