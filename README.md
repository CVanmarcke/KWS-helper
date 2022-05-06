# KWS-helper
Collectie van Autohotkey scripts om de radiologie workflow te verbeteren.

## Installatie
Download eerst de zip versie van [Autohotkey hier](https://www.autohotkey.com/download/), en pak dit uit op een plaats die je later terugvindt.
Download vervolgens de bestanden van deze repo [hier](https://github.com/CVanmarcke/KWS-helper/archive/refs/heads/main.zip), en pak deze uit zodat de inhoud in dezelfde folder als het `AutohotkeyU64.exe` bestand terechtkomt.
Start autohotkey via `AutohotkeyU64.exe`

## Hotkeys veranderden
TODO

## Mogelijke functies
TODO
Functienaam | Standaard Hotkey | Beschrijving 
--- | --- | --- 
copyLastReport_KWS() | `F7`<br />`-i- (speechkit)` | Plakt de inhoud van het voorgaande verslag in het huidige, corrigeert de inhoud en data, voegt "in vergelijking met" toe, en doet enkele kleinere aanpassingen. 
cleanReport_KWS() | `F9`<br />`F3 (speechkit)` | Kuist het huidige verslag op door meerdere kleine aanpassingen te doen: oa. zet een punt achter elke zin, corrigeert de hoofdletter die achter een dubbelpunt wordt gezet, zet streepjes voor elke lijn, sorteert alles met een punt vooraan onder het verslag (en respecteerd subheadings bij bijv thorax-abdomen of conclusies), flair->FLAIR, bekend->gekend, besluit->conclusie, D11-D12 -> Th11-Th12, ... Deze instellingen kunnen worden aangepast onder de functie `cleanreport(inputtext)`
validateAndClosePt_KWS() | `INS (speechkit, 2x duwen)` | Valideert het verslag door tweemaal op de knop te duwen. Voegt bijkomend dit toe aan de logfile: de patiënt kan opnieuw geopend worden door eender waar `openlastpt` te typen. 
saveAndClosePt_KWS() | `EOL (speechkit, 2x duwen)` | Slaat het verslag op en sluit het vervolgens, door tweemaal op de knop te duwen. Voegt bijkomend dit toe aan de logfile: de patiënt kan opnieuw geopend worden door eender waar `openlastpt` te typen.  
heightLossGUI() | `hoogteverlies`<br />`wervelfx` | Typ de shortcut om een klein venster te krijgen waar je de hoogte van 2 wervels kan typen, waarna automatisch het hoogteverlies (in mm en %) geplakt zal worden. 
openLastPtInLog_KWS() | `openlastpt` | Typ de shortcut eender waar om de laatste patiënt gesloten via `validateAndClosePt_KWS` of `saveAndClosePt_KWS` opnieuw te openen. 
openEAD_KWS() | `Ctrl-O` | TODO
pedAbdomenTemplate() | `pedabdomen` | Typ de shortcut om een venster te openen waar de lengte van de lever, milt en nieren in ingevuld kunnen worden, waarna automatisch een verslag met ingevulde standaarddeviaties geplakt zal worden. 
MoveLineUp() | `Ctrl-↑`<br />`rewind (speechkit)` | Verplaatst de huidige lijn één lijn naar boven (wisseld de huidige lijn met de lijn er net boven).
MoveLineDown() | `Ctrl-↓`<br />`forward (speechkit)` | Verplaatst de huidige lijn één lijn naar onder (wisseld de huidige lijn met de lijn er net onder).
deleteLine() | `Ctrl-d` | Verwijderd de huidige lijn. 

## Speechkit knoppen
Knop | functie
--- | --- 
`INS (2x duwen)` | TODO 
`EOL (2x duwen)` | TODO 
`-i-` | TODO 
`record` | TODO 
`rewind` | TODO 
`forward` | TODO 
`play/pause` | TODO 
`F1` | TODO 
`F2` | TODO 
`F3` | TODO 
`F4` | TODO 
`back button (hold)` | simuleert de Ctrl knop: kan gebruikt worden om in enterprise te zoomen en pannen. 

