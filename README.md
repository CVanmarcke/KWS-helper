# KWS-helper
Collectie van Autohotkey scripts om de radiologie workflow te verbeteren.

## Installatie
Download eerst de zip versie van [Autohotkey hier](https://www.autohotkey.com/download/), en pak dit uit op een plaats die je later terugvindt.
Download vervolgens de bestanden van deze repo [hier](https://github.com/CVanmarcke/KWS-helper/archive/refs/heads/main.zip), en pak deze uit zodat de inhoud in dezelfde folder als het `AutohotkeyU64.exe` bestand terechtkomt.
Start autohotkey via `AutohotkeyU64.exe`

## Hotkeys veranderden
TODO
De hotkeys zijn gedefinieerd in `AutohotkeyU64.ahk`. Zie de [Autohotkey tutorial](https://www.autohotkey.com/docs/Tutorial.htm#s2) voor een korte handleiding om hoe deze te veranderen.

## Mogelijke functies
Functienaam | Standaard Hotkey | Beschrijving 
--- | --- | --- 
copyLastReport_KWS() | `F7`<br />`-i- (speechkit)` | Plakt de inhoud van het voorgaande verslag in het huidige, corrigeert de inhoud en data, voegt "in vergelijking met" toe, en doet enkele kleinere aanpassingen. 
cleanReport_KWS() | `F9`<br />`F3 (speechkit)` | Kuist het huidige verslag op door meerdere kleine aanpassingen te doen: oa. zet een punt achter elke zin, corrigeert de hoofdletter die achter een dubbelpunt wordt gezet, zet streepjes voor elke lijn, sorteert alles met een punt vooraan onder het verslag (en respecteerd subheadings bij bijv thorax-abdomen of conclusies), flair->FLAIR, bekend->gekend, besluit->conclusie, D11-D12 -> Th11-Th12, ... Deze instellingen kunnen worden aangepast onder de functie `cleanreport(inputtext)`
validateAndClosePt_KWS() | `INS (speechkit, 2x duwen)` | Valideert het verslag door tweemaal op de knop te duwen. Voegt bijkomend dit toe aan de logfile: de patiënt kan opnieuw geopend worden door eender waar `openlastpt` te typen. 
saveAndClosePt_KWS() | `EOL (speechkit, 2x duwen)` | Slaat het verslag op en sluit het vervolgens, door tweemaal op de knop te duwen. Voegt bijkomend dit toe aan de logfile: de patiënt kan opnieuw geopend worden door eender waar `openlastpt` te typen.  
heightLossGUI() | `hoogteverlies`<br />`wervelfx` | Typ de shortcut om een klein venster te krijgen waar je de hoogte van 2 wervels kan typen, waarna automatisch het hoogteverlies (in mm en %) geplakt zal worden. 
openLastPtInLog_KWS() | `openlastpt` | Typ de shortcut eender waar om de laatste patiënt gesloten via `validateAndClosePt_KWS` of `saveAndClosePt_KWS` opnieuw te openen. 
openEAD_KWS() | `Ctrl-O` | TODO
pedAbdomenTemplate() | `pedabdomen` | Typ de shortcut om een venster te openen waar de lengte van de lever, milt en nieren in ingevuld kunnen worden, waarna automatisch een verslag met ingevulde standaarddeviaties geplakt zal worden. Origineel gemaakt door Johannes Devos.
MoveLineUp() | `Ctrl-↑`<br />`rewind (speechkit)` | Verplaatst de huidige lijn één lijn naar boven (wisseld de huidige lijn met de lijn er net boven).
MoveLineDown() | `Ctrl-↓`<br />`forward (speechkit)` | Verplaatst de huidige lijn één lijn naar onder (wisseld de huidige lijn met de lijn er net onder).
deleteLine() | `Ctrl-d` | Verwijderd de huidige lijn. 
auto_scroll() | `Ctrl-&` of `Ctrl-é` | Start automatisch met scrollen in het venster waar de muis in staat (Enterprise of IMPAX). Vertragen met `&` en versnellen met `é`, omkeren van richting met `space`. Gemaakt door Johannes Devos.

## Speechkit knoppen
Knop | functie
--- | --- 
`EOL (2x duwen)` | saveAndClosePt_KWS() 
`-i-` | copyLastReport_KWS() 
`INS (2x duwen)` | validateAndClosePt_KWS() 
`record` | Niet aangepast: start/stop opnemen. 
`rewind` | MoveLineUp() 
`forward` | MoveLineDown() 
`play/pause` | `Backspace` 
`F1` | copyLastReport_KWS() 
`F2` | / 
`F3` | cleanReport_KWS() 
`F4` | Ctrl-F8 (start speechkit in venster)
`back button (hold)` | Simuleert de Ctrl knop: kan gebruikt worden om in enterprise te zoomen en pannen. 

## Ideeën, bigfixes en aanpassingen
Liefst een berichtje op Whatsapp of UZ Leuven mail.

Voor bugfixes graag de error noteren en op welke lijn de error voorkwam, samen met een beschrijving van hoe het wordt uitgelokt.
