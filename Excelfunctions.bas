Attribute VB_Name = "Excelfuncties"

'Deze declaratie is een timer. Kan worden aangeroepen met de call Sleep(miliseconden)

Option Explicit


'-------------------------'
'Auteur: Siebe Bosch      '
'Hydroconsult             '
'Lulofsstraat 55, unit 47 '
'2521 AL Den Haag         '
'siebe@hydroconsult.nl    '
'0617682689               '
'-------------------------'

'VERSIE 4.96
    'bovenstaande regel is bedoeld om het declareren van elke variabele verplicht te maken
'Beschikbare functies en routines:

'getallenreeksen:
'COUNT_UNIQUE                - telt het aantal unieke waarden in een bereik
'GOALSEEKRANGE               - voert een goalseek uit op een range
'COUNTDISTINCTVALUESINRANGE  - telt het aantal unieke waarden in een gegeven bereik
'INTERPOLATE                 - Interpoleert tussen twee xy-punten, blockinterpolation is optioneel
'EXTRAPOLATE                 - Extrapoleert lineair vanaf twee xy-punten
'FITLINEAR_A                 - berekent a in y = ax + b gegeven twee coordinaten
'FITLINEAR_B                 - berekent b in y = ax + b gegeven twee coordinaten
'INTERPOLATEGAPSINRANGE      - interpoleert alle gaten in een gegeven bereik
'INTERPOLATENODATAVALUESINRANGE - interpoleert alle nodata-values in een gegeven bereik
'INTERPOLATEFROMRANGE        - interpoleert een gegeven X in een range van X-waarden en een range van Y-waarden
'INTERPOLATERANGEFROMRANGE   - interpoleert voor een hele reeks getallen (in een range) uit een gegeven range van X en Y-waarden
'INTERPOLATEFROMRANGEPLUS    - interpoleert een gegeven X in een range van X-waarden en een range van Y-waarden, gegeven een ID in een range met ID's
'KLEINSTEKWADRATENMETHODE    - geeft het kleinstekwadratenverschil tussen een gemeten en berekende reeks
'ISSTRINGARRAYEMPTY          - test of een array van strings leeg is
'SORTARRAY                   - Sorteert een array met getallen in oplopende volgorde
'HEAPSORT                    - Creeert een array met de indexnummers van de gesorteerde input-array
'SORTCOLLECTIONBYKEY         - Creeert een array met de indexnummers van een op key gesorteerde collection van objecten
'COLLECTIONGETIDX            - geeft het indexnummer terug van een gegeveen waarde in een collectie
'COLLECTIONCONTAINS          - geeft een boolean terug met het antwoord of een collectie een gegeven waarde bevat
'RANDOM                      - creeert een random integer tussen een gespecificeerd minimum en maximum
'RANDOMDOUBLE                - creeert een random double tussen een gespecificeerd minimum en maximum
'MAXIMUM                     - geeft het maximum van twee getallen terug
'FROMWORKSHEET               - haalt data uit een range van een werkblad en stopt die in een array
'ARRAYVARIANTTOWORKSHEET     - schrijft data uit een array van het type variant naar een werkblad
'ARRAYDATETOWORKSHEET        - schrijft data uit een array van het type date naar een werkblad
'ARRAYSINGLETOWORKSHEET      - schrijft data uit een array van het type single naar een werkblad
'TIMESERIES2ARRAYS           - leest een tijdreeks van het werkblad in twee arrays
'USERSELECTRANGE             - laat de gebruiker een range op het werkblad selecteren
'RANGEADDRESSFROMRC          - definieert een bereik op basis van rij- en kolomnummers
'GETCOLUMNNAME               - geeft de kolomnaam terug, gegeven een celadres (bijv. "D4")
'SHIFTCELLADDRESS            - voert een shift van een aantal kolommen en rijen uit op een gegeven celadres
'RANGECOLIDXFROMMAXVAL       - geeft het kolomnummer terug dat hoort bij de hoogste waarde uit een bereik
'ASSIGNVALUEBYMONTH          - geeft een waarde terug, afhankelijk van de maand van een gegeven datum
'CLASSIFYNUMBERBYCLASS       - classificeert een getal naar een gegeven ondergrens, bovengrens en increment
'DESIGNTWOPARTTEMPORALPATTERNS  - genereert een serie tijdsafhankelijke patronen, gegeven een duur (tijdstappen) en stapgrootte (percentiel)
'DESIGNTEMPORALBLOCKPATTERNS - genereert een serie blokpatronen, gegeven een duur (tijdstappen) en blockgrootte
'OBSERVATIONSTOWEBVIEWER     - exporteert een werkblad met gemeten tijdreeksen naar een javascript-bestand met JSON als input voor de webviewer van Hydroconsult
'POINTOBJECTSTOWEBVIEWER     - exporteert een werkblad met objecten naar een javascript-bestand met JSON als input voor de webviewer van Hydroconsult

'graphics:
'SVGCOMPONENTTRAPEZIUMSHARP  - genereert een trapeziumvorm en schrijft het in SVG-formaat (zonder headers)

'collections:
'MAXFROMCOLLECTION           - retourneert de maximumwaarde uit een collectie
'MINFROMCOLLECTION           - retourneert de minimumwaarde uit een collectie
'AVGFROMCOLLECTION           - retourneert de gemiddelde waarde uit een collectie

'kansrekening:
'GEVCUM                      - berekent de cumulatieve kansdichtheid volgens de Gegeneraliseerde Extremewaardenverdeling (GEV)
'GENPARETOCDF                - berekent de cumulatieve kansdichtheid volgens de Gegeneraliseerde Pareto-kansverdeling (GENPAR)
'EXP2PCUM                    - berekent de cumulatieve kansdichtheid volgens de tweepunts-exponentiele verdeling (EXP2P)
'CONDWEIBULLCDF              - berekent de cumulatieve kansdichtheid volgens de conditionele weibull-verdeling (CONDWEIBULL)
'GENPARETOCDF                - berekent de cumulatieve kansdichtheid volgens de gegeneraliseerde pareto-verdeling
'BEREKENSTOCHASTVOLUMEKLASSE       - berekent de aangepaste frequentie wanneer bepaalde stochasten voor NeerslagVolume zijn uitgeschakeld
'BEREKENSTOCHASTPATROONKLASSE      - berekent de aangepaste frequentie wanneer bepaalde stochasten NeerslagPatroon zijn uitgeschakeld
'HERH2KLASSEFREQ             - berekent de frequentie van een klasse gegeven herhalingstijd van vorige, huidige en volgende klasse en duur
'HERHFROMSTOCHASTICRESULT    - berekent de waterhoogte behorende bij een gegeven herhalingtijd op basis van de uitkomsten van een stochastenanalyse
'KLASSEFREQUENTIEUITHERHALINGSTIJD  - berekent de frequentie van een klasse gegeven de onder- en bovengrens van volumes uitgedrukt in herhalingstijd
'KLASSEKANSUITOVERSCHRIJDINGSKANSEN   - berekent de kans voor een klasse, gegeven omringende overschrijdingskansen
'CLASSIFYDURATIONS           - classificeert gebeurtenissen naar hun duur door te zoeken naar de duur van een overschrijding

'meetkundig:
'OPPERVLAKAFGEPLATTECIRKEL   - berkent het oppervlak van een afgeplatte cirkel
'NATTEOMTREKAFGEPLATTECIRKEL - berekent de contactomtrek van een afgeplatte cirkel
'ELLIPSBREEDTE               - berekent bij verschillende hoogtes de breedte van een ellips
'ARCSIN                      - berekent de inverse sinus
'ARCCOS                      - berekent de inverse cosinus
'ARCTAN                      - berekent de inverse tangens
'ROTATEPOINT                 - roteert een xy-coordinaat rond een vastgelegd nulpunt en verplaatst het
'D2R                         - graden naar radialen
'R2G                         - radialen naar graden
'LINEANGLEDEGREES            - berekent de hoek van een lijn tussen twee xy-co-ordinaten
'POINTDISTANCE               - berekent de afstand tussen twee x,y coordinaten
'POINTINPOLYGON              - berekent of een gegeven punt binnen een polygoon ligt
'NEARESTPOINT                - zoekt gegeven een XY-coordinaat het dichtstbijzijnde punt uit een range
'POOLCOORDINAATX             - geeft X terug gegeven een hoek alpha (ten opzichte van noord-as) en lengte L
'POOLCOORDINAATY             - geeft Y terug gegeven een hoek alpha (ten opzichte van noord-as) en lengte L
'PYTHAGORAS                  - geeft lengte schuine kant terug
'PYTHAGORAS_INVERSE          - geeft lengte van een rechte kant terug
'TRAPEZIUMAREA               - geeft het oppervlak van een trapezium (bestaande uit een rechthoek met daarbovenop een driehoek) terug.

'wiskundig:
'MILEAGEONEUP                - verhoogt de 'kilometerstand' in een array met één
'MEETSCONDITION              - checkt of een waarde voldoet aan een bepaalde condition (bijv. ">= -0.52")
'RMSE                        - berekent de Root Mean Square Error voor twee ranges

'objecten in Excel
'CLEARCOMBOBOX               - verwijdert alle items uit een combobox
'GetShapeByNameFromWorksheet - vraagt een shape van een gegeven werkblad op, op basis van zijn naam.

'grafieken
'MAKESCATTERCHART
'MAKECHART
'EXPORTCHART                 - exporteert een grafiek naar een .png-bestand in dezelfde folder als de applicatie.

'datum- en tijdfuncties
'DAYSINMONTH                 - Geeft het aantal dagen in de maand van een gespecificeerde datum terug
'DAYSINMONTH2                - Geeft het aantal dagen in de maand van een gespecificeerd maandnummer
'ISLEAPYEAR                  - Geeft terug of een gegeven jaar een schrikkeljaar is (TRUE/FALSE)
'KWARTAAL                    - Geeft het kwartaalnummer van een datum terug
'ZOMERWINTERHALFJAAR         - Geeft het halfjaar van een datum terug
'METEOROLOGISCHSEIZOEN       - Geeft het meteorologisch seizoen terug waarin een opgegeven datum ligt
'METEOROLOGISCHHALFJAAR      - Geeft het meteorologische halfjaar terug van een opgegeven datum
'HYDROLOGISCHSEIZOEN         - Geeft het hydrologisch seizoen terug waarin een opgegeven datum ligt
'DOUBLE2DATETIMESTRING       - transformeert een getal (double) naar een datum/tijd-string
'DATEEXISTS                  - controleert of een opgegeven combinatie van dag, maand en jaar valide is
'DAYNUMBER                   - geeft het dagnummer in het jaar van een gegeven datum terug
'CALCDAYQUARTER              - geeft datum + eerste uur binnen het zesuurswindow van een gegeven datumtijd terug
'DATEFROMSTRING              - maakt een datum aan op basis van een string
'DATE2TEXT                   - maakt een string aan op basis van een datum
'TIMEFROMSTRING              - maakt een tijd aan op basis van een string
'DATEANDTIMEFROMSTRINGS      - maak een datum en tijd aan op basis van twee strings
'DECADE                      - berekent het decadenummer gegeven een datum
'SOBEKTIMETABLESTRING        - maakt een tijdstabel voor SOBEK in de Control.def
'AvgExceedenceValueContiguous - bepaalt de overschrijdingswaarde die hoort bij een gemiddeld eens per jaar overschrijding met een gegeven duur

'werkbladfuncties
'HOR_ZOEKEN_DOUBLE           - laat de gebruiker zoeken op bass van waarden in de eerste TWEE rijen
'VERT_HORIZ_ZOEKEN           - Geef kolomnaam en rijnaam op, en krijg de inhoud van de bijbehorende cel terug
'VERT_ZOEKEN_DOUBLE          - laat de gebruiker zoeken op basis van waarden in de eerste TWEE kolommen
'VERT_ZOEKEN_TRIPLE          - laat de gebruiker zoeken op basis van waarden in de eerste DRIE kolommen
'VERT_ZOEKEN_QUADRUPLE       - laat de gebruiker zoeken op basis van waarden in de eerste VIER kolommen
'VERT_ZOEKEN_MIN             - Geeft de minimumwaarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_MAX             - Geeft de maximumwaarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_MODUS           - Geeft de meest voorkomende waarde terug uit een range waarin eenzelfde ID meerdere Parent1n voorkomt
'VERT_ZOEKEN_SOM             - Sommeert alle waarden uit kolom Y die gevonden worden bij een opgegeven zoekterm in kolom X
'VERT_ZOEKEN_NEARESTXY                  - Lookup in een range met X,Y en Waarde, waarbij de waarde van het meest dichtbijzijnde object wordt teruggegeven
'FindColumnInRange           - geeft de kolomindex van een range terug, gegeven een waarde die gezocht wordt. optioneel geeft hij een lege terug indien niet gevonden
'FindRowInRange              - geeft de rijindex van een range terug, gegeven een waarde die gezocht wordt. optioneel geeft hij een lege terug indien niet gevonden
'AVERAGEFROMRANGE            - geeft de gemiddelde waarde uit een range terug
'MINFROMRANGE                - geeft de laatste waarde uit een range terug
'MAXFROMRANGE                - geeft de kleinste waarde uit een range terug
'FIRSTFROMRANGE              - geeft de eerste waarde uit een range terug
'LASTFROMRANGE               - geeft de laatste waarde uit een range terug
'MOSTCOMMONFROMRANGE         - Geeft de meest voorkomende waarde terug uit een range
'GEWOGEN_GEMIDDELDE          - Geeft van een reeks voor elke ID de gewogen gemiddelde waarde terug op basis van bijv. meerdere waarde- en oppervlakteparen
'VERT_ZOEKEN_GROOTSTEAANDEELHOUDER      - Geeft terug voor welke 'aandeelhouder' de som van de waarden het grootst is, gegeven een objectID
'HEADERBYMAXIMUMVALUE        - Geeft voor een gegeven range de titel terug die staat boven de kolom met de grootste waarde
'WORKSHEETEXISTS             - Returns 'true if a worksheet exists
'SUMRANGE                    - Geeft de som van de inhoud van een range op een werkblad terug
'FRACTIONOFDAYSUM            - Geeft voor een bepaalde cel voor een bepaalde datum/tijd de fractie van de totale dagsom terug
'ISRANGEASCENDING            - Checkt of een opgegeven range een oplopende volgorde heeft
'MINYFROMXYRANGE             - Geeft de minimum Y-waarde terug uit een range met daarin X en Y waarden. Optioneel zoekrange beperken van Xmin tot Xmax
'MAXYFROMXYRANGE             - Geeft de maximum Y-waarde terug uit een range met daarin X en Y waarden. Optioneel zoekrange beperken van Xmin tot Xmax
'CONCATENATEALGEBRAIC        - Veegt termen uit een reeks samen tot een algebraische formule, bijv "X + Y + Z"
'CONCATENATEWITHDELIMITER    - Veegt waarden uit een reeks van cellen samen tot een string, met een gegeven delimiter
'ADDWORKSHEET                - Voegt een nieuw werkblad toe aan het huidige werkboek.
'FINDCOLUMNONWORKSHEET       - Zoekt het kolomnummer voor een gegeven header op een werkblad
'UNPIVOTMULTIHEADER          - ontvlecht een tabel met meerdere headers in zowel rij als kolom en zet hem klaar voor pivot-doeleinden
'UNPIVOTTABLE                - ontvlecht een tabel met één kolomheader en één rijheader en schrijft hem in COLHEADER, ROWHEADER, VALUE formaat
'UNPIVOT                     - converteert een 2D-tabel naar een Header1, Header2, Waarde-tabel voor pivot-doeleinden
'UNPIVOT2CSV                 - converteert een 2D-tabel naar een Header1, Header2, Waarde-tabel in een .csv-bestand
'TIMESERIES2CSV              - schrijft een tijdreeks uit een range naar een CSV-bestand met opgegeven datumformatting
'RANGE2CSV                   - schrijft de gegevens uit een range naar een csv file.
'TIMESERIESFROMCSV           - leest tijdreeks uit een CSV en schrijft die naar een opgegeven werkblad
'GOALSEEKTRIPLE              - zoekt het optimum voor een cel waarvan de waarde een functie is van drie variabelen
'GOALSEEKDOUBLE              - zoekt het optimum voor een cel waarvan de waarde een functie is van twee variabelen
'COLUMN_NUMBER               - zoekt het kolomnummer uit een reeks, gegeven een gezochte celinhoud
'PRINTARRAY                  - schrijft een array naar het werkblad
'RANGEVERTASCENDING           - checkt of een range in vertikale richting oploopt
'VALUEFROMCELLADDRESS        - geeft de waarde uit een cel terug gegeven zijn rij- en kolomnummer

'werkbladroutines voor ranges (hiervoor moet je wel een knop inbouwen)
'AGGREGERENNAARUREN          - aggregeert een tijdreeks met waarden naar hele uren
'COUNTSEQUENTIALEXCEEDANCES  - telt het aantal achtereenvolgende overschrijdingen van een drempelwaarde in een gegeven range (betaande uit een kolom)
'AGGREGEREN                  - aggregeert een tijdreeks door een vast aantal rijen over te slaan
'AGGREGATEFROMRANGE          - aggregeert een gegeven kolom uit een range op basis van waarden in een andere kolom en een opgegeven aggregatiemethode
'AGGREGATERANGECONDITIONALLY - aggregeert een range op basis van een geselecteerde kolom en een specificatie van de aggregatiemethode per kolom, maar met een voorwaarde voor de waarden uit een andere kolom
'COLUMNFROMRANGE             - geeft een kolom uit een range terug als range. Houdt ook rekening met multi-area ranges!
'CONDITIONALSUBRANGE         - geeft een subrange uit een range terug, waarvoor aan een gegeven voorwaarde wordt voldaan
'GETASCIIGRIDVALUES          - geeft voor gegeven X en Y coordinaten de bijbehorende waarde uit een ASCIIGRID terug
'GETROWCOLFROMASCIIGRID      - geeft voor gegeven grid-dimensies het rij- en kolomnummer behorende bij een X- en Y-coordinaat terug
'RANGEWITHHEADER2THREECOLRANGE - converteert een reeks met header en daaronder X en Y naar een reeks met drie kolommen: ID, X, Y
'WEAVETABLESBLOCKINTERPOLATION - weeft twee tabellen met datum/waarde ineen, werkend met blokinterpolatie. Handig voor gemaalactiviteiten
'TRUNCATERANGEBYEMPTYROWS      - kapt een gegeven range af op lege rijen

'eenheidsconversies
'CELCIUS2KELVIN              - converteert graden celcius naar kelvin
'KELVIN2CELCIUS              - converteert graden kelvin naar celcius
'FORMATROMAN                 - converteert een getal (integer) naar Romeins formaat
'LSHA2MMPD                   - converteert liter/seconde/ha naar mm/d
'MMPD2LSHA                   - converteert mm/d naar liter/seconde/ha
'M3PS2MMPD                   - converteert m3/s naar mm/d
'MMPD2M3PS                   - converteert mm/d naar m3/s
'MMPU2M3PS                   - converteert mm/u naar m3/s
'M3PS2MMPU                   - converteert m3/s naar mm/u

'geografisch
'RD2LATLONG                  - converteert een coordinaat in RD naar LAT/LONG
'RD2LAT                      - converteert een coordinaat in RD naar LAT
'RD2LON                      - converteert een coordinaat in RD naar LONG
'RD2WGS84                    - converteert een coordinaat in RD naar WGS84 (LAT/LONG)
'WGS842RD                    - converteert een WGS84-coordinaat (LAT/LONG) naar RD
'WGS842X                     - converteert een WGS84-coordinaat (LAT/LONG) naar RD, X-coordinaat
'WGS842Y                     - converteert een WGS84-coordinaat (LAT/LONG) naar RD, Y-coordinaat
'RD2BESSEL                   - van RD naar besselfunctie
'BESSEL2WGS84                - van Besselfunctie naar latlong
'WGS84DEG2DECIMAL            - converteert latlong in graden naar decimalen
'WGS84DEG2LATDECIMAL         - converteert latlong in graden naar latitude in decimalen
'WGS84DEG2LONDECIMAL         - converteert latlong in graden naar longitude in decimalen
'WGS842BESSEL                - van latlong naar besselfunctie
'BESSEL2RD                   - van Besselfunctie naar RD
'ExtractCoordinatesFromWKT   - haalt coordinaten uit een WKT-string (QGis export topology)

'electronica
'OHM                         - berekent de benodigde weerstand, gegeven het voltage en gewenste amperage


'hydrologisch
'GETIJDEN_SINUS              - Berekent de waterstand van een getijdenslag voor elk gewenst tijdstip
'YZ2TABULATED                - Converteert een YZ-profiel naar tabelprofiel
'YZ2TABULATEDINVERTED        - Converteert een omgekeerd YZ-profiel (bijv.boogbrug) naar tabelprofiel
'QSTUW                       - Berekent het debiet over een rechthoekige stuw
'QTHOMSONWEIR                - Berekent het debiet over een Thomson-meetstuw met een V van 90 graden
'QVNOTCHWEIR                 - Berekent het debiet over een V-Notch weir, als functie van de V-hoek, gebruikmakend van Kindsvater-Shen
'QUNIVERSALWEIR              - Berekent het debiet over een universal weir (YZ)
'QABUTMENTBRIDGE             - berekent het debiet door een landhoofdbrug
'BREEDTE_STUW                - berekent de gewenste breedte van een stuw, gegeven de afvoer in mm/d en overstortende straal
'WEIRSUBMERGED               - Berekent of een stuw verdronken is
'QHEVEL                      - Berekent het debiet door een hevel
'QDUIKERRECHTHOEK            - Berekent het debiet door een rechthoekige duiker
'QDUIKER                     - Berekent het debiet door een duiker
'QORIFICE                    - Berekent het debiet door een schuif met gegeven dH, breedte en openingshoogte
'DHGEVULDERONDEDUIKER        - Berekent het verval over een ronde duiker die geheel gevuld is met water
'DHRONDEDUIKER               - Berekent het verval over een ronde duiker die al dan niet geheel gevuld is
'WIDTHORIFICE                - Berekent de benodigde breedte van een niet-verdronken onderlaat gegeven een gevraagd debiet en drempelhoogte
'COUNTSPUIPERIODES           - berekent het aantal spuiperiodes wat verwacht mag worden in een gegeven periode
'GETTIDALWAVEMINUTES         - berekent de duur van één getijdenslag, uitgedrukt in minuten
'GETAVGMINFROMTIDE           - haalt de gemiddelde laagwaterstand uit een getijdenreeks
'GETAVGMAXFROMTIDE           - haalt de gemiddelde hoogwaterstand uit een getijdenreeks
'WINDRICHTING                - geeft de windrichting terug (N, NO, O, ZO, etc.) als functie van de hoek in graden. Optioneel in graden (0,45,90,135,180,225,270,315,360)
'TIDALMINMAXFROMSERIES       - haalt per getijdenslag de hoogste en laagste waterstand binnen en schrijft deze naar een range
'TIDALLOWSFROMSERIES         - haalt per getijdenslag de laagste waterstand binnen en schrijft deze naar een range
'LGN5TONBW                   - converteert LGN code naar de benodigde landgebruikscode voor NBW-toetsing (zoals afgeleid voor waterschap Noorderzijlvest)
'LGN2SOBEK                   - converteert LGN code naar landgebruiksnummer in SOBEK (1=grass, 2=potatoes etc.)
'ERNSTRecord                 - schrijft een ERNST-record weg
'BOD2CAPSIM                  - converteert bodemcode (letter + cijfer) naar CAPSIM bodemtypenummer voor SOBEK
'GHGGLG2GT                   - converteert een gegeven GHG en GLG (m - maaiveld) naar Grondwatertrap
'HYDROZOMERWINTER            - berekent of een datum in de hydrologische zomer/winter valt
'EVAPDEBRUINKEIJMAN          - openwaterverdamping volgens de bruin-keijman
'EVAPMAKKINK                 - referentiegewasverdamping volgens Makkink
'MAKKINK2OPENWATER           - converteert makkinkverdamping naar openwaterverdamping
'OPENWATEREVAPFACTOR         - berekent gegeven de datum de 'gewasfactor' openwaterverdamping terug
'EVAPDAY2HOUR                - deaggregeert etmaalverdampingssom naar uurwaarden
'HOURLYEVAPORATIONFRACTION   - bepaalt, gegeven het uur van de dag, de fractie van de etmaalverdampingssom op basis sinus van 6 tot 18)
'HydraulicRadius             - berekent de hydraulische straal
'Manning2Chezy                      - Converteert n_manning naar chezy ruwheid
'Chezy2Manning                      - converteert chezy naar n_manning ruwheid
'Chezy                              - berekent Q voor een gegeven Chezy-waarde en bodemverhang
'MaatgevendeAfvoer                  - berekent de maatgevende afvoer op basis van neerslagintensiteit en oppervlak
'NEERSLAGPATROON                    - berekent het type neerslagpatroon volgens STOWA 2004 (Nieuwe neerslagstatistiek voor waterbeheerders)
'WETTEDPERIMETERFROMYZPROFILE       - berekent de natte omtrek voor een gegeven waterhoogte in een YZ-profiel
'WettedAreaFromYZProfile            - berekent de natte doorsnede voor een gegeven waterhoogte in een YZ-profiel

'statistiek
'CORRELATIONBYWINDOW                - berekent de correlatiecoefficient r tussen twee ranges met gegeven window size
'EXPONENTIELEVERDELINGCDF           - berekent de cumulatieve voor de Exponentiele verdeling
'GUMBELFIT                          - fit de gumbel-kansverdeling aan een gegeven reeks van maxima, gebruikmakend van de Maximum Likelihood
'GUMBELVERDELINGSFUNCTIE            - berekent de ONDERschrijdingskans van een bepaalde parameterwaarde op basis van opgegeven GUMBEL-parameters en parameterwaarde X
'GUMBELINVERSE                      - berekent de parameterwaarde die hoort bij een gegeven ONDERschrijdinskans volgens de GUMBEL kansverdeling type I
'GENPARETOINVERSE                   - berekent (iteratief) de inverse van de Generalized Pareto CDF
'GENPARETOCDF                       - berekent de cumulatieve onderschrijdingskans volgens Gen. Pareto
'GEVVERDELINGSFUNCTIE               - berekent de ONDERschrijdingskans van een bepaalde parameterwaarde op basis van een opgegeven GEV-kansverdeling
'GEVINVERSE                         - berekent de parameterwaarde die hoort bij een gegeven ONDERschrijdingskans volgens de GEV-kansverdeling
'CALCNEERSLAGSTATS                  - berekent de statistische parameters (GEV-kansdichtheidsfunctie) van een bui, gegeven gebiedsoppervlak en neerslagduur
'CALCHERHALINGSTIJD                 - berekent gegeven neerslagduur (uren), gebiedsoppervlak (km2) en volume (mm) de herhalingstijd
'CALCNEERSLAGVOLUME                 - berekent gegeven herhalingstijd, neerslagduur en gebiedsoppervlak het bijbehorende neerslagvolume
'PRECIPITATIONAREAREDUCTION         - reduceert het neerslagvolume in een reeks als functie van duur, herhalingstijd en gebiedsoppervlak
'ANNUALMAXIMUMPRECIPITATIONEVENTS   - extraheert uit een neerslagreeks en een gegeven neerslagduur de jaarmaxima, zowel voor zomer, winter als jaarrond
'PLOTTINGPOSITIONFROMANNUALMAXIMA   - berekent de plotting position (herhalingstijd in jaren) voor een getal adhv een lijst met jaarmaxima
'IDENTIFYPRECIPITATIONEVENTSPOT     - extraheert uit een neerslagreeks gegeven een POT-waarde en neerslagduur de neerslagvolumes
'CLASSIFYEVENTS                     - schrijft rangnummers weg voor gebeurtenissen met gegeven duur die binnen een klasse > X en < Y vallen
'RANKNUMBEROFEXCEEDANCES            - schrijft de rangnummers van overschrijdingen van een vooraf opgegeven drempel naar een nieuwe kolom
'POTANALYSISSUM                     - indexeert de extreemste gebeurtenissen met vooraf opgegeven duur door hun som en indexnummers naar naastgelegen kolommen te schrijven
'POTANALYSISMAX                     - indexeert de extreemste gebeurtenissen met vooraf opgegeven duur door hun maximum en indexnummers naar naastgelegen kolommen te schrijven
'CALCULATEEXTREMEEVENTS             - extraheert uit een neerslagreeks alle zwaarste buien gegeven neerslagduur
'NASH_SUTCLIFFE                     - berekent de nash-sutcliffe-coëfficiënt voor twee reeksen
'NASH_SUTCLIFFE_FAST
'FILTERBASEFLOW                     - filtert baseflow of interflow uit een opgegeven reeks met afvoeren
'HOOGHOUDT_q                        - berekent de stationaire q voor een situatie met twee drains volgens de formule van Hooghoudt
'HOOGHOUDT_L                        - berekent de drainafstand tussen twee drains volgens de formule van Hooghoudt, gegeven q en opbolling
'YZHYDRAULICPROPERTIES              - berekent de hydraulische eigenschappen van een gegeven YZ-profiel: A, P en R as functie van de diepte
'XYZTOYZPROFILE                     - genereert een YZ-profiel op basis van gegeven XYZ-coordinaten

'sobek
'READSPECIFICHISRESULTS      - leest resultaten uit een HIS-file met vooraf opgegeven locatie en parameter
'READHISLOCPARTIM            - leest de Locaties, Parameters en Tijdstappen in van een HIS-file
'GETNODESTATSFROMSOBEK       - haalt voor alle objecten in calcpnt.his x,y,min,max,avg,first en last op
'MERGESTORAGETABLES          - voegt twee bergingstabellen samen (hoogte/oppervlak). Beide tabellen moeten een collection of clsLevelAreaPair zijn
'INTERPOLATEFROMSTORAGETABLE - interpoleert uit een bergingstabel (hoogte/oppervlak) invoer:hoogte, uitvoer:oppervlak. tabel moet collection of clsLevelAreaPair zijn
'PARSESOBEKFILE              - parst een sobek inputfile en schrijft het resultaat naar een vooraf opgegeven locatie
'PARSESOBEKTABLE             - parst een sobek tabel en geeft een array met de resultaten terug
'PARSEBYSINGLECHAR           - parst een string op basis van 1 karakter per keer
'MAKESOBEKTARGETLEVELTABLE   - maakt een tabel met zomer- en winterstreefpeilen
'READBUIFILE                 - leest een .bui file van SOBEK in en schrijft de data naar het werkblad
'WRITEBUIFILE                - leest neerslagdata van het werkblad en schrijft een .bui file voor SOBEK weg
'WRITERKSFILE                - leest neerslagdata van het werkblad en schrijft een .rks file voor SOBEK weg
'WRITEPRNFILE                - leest een tijdtabel van het werkblad en schrijft een .prn file voor SOBEK weg
'WRITEPRNFILES               - leest een tijdtabel van het werkblad en schrijft meerdere .prn files voor SOBEK weg
'WRITERRBOUNDARYDATA         - schrijft bound3b.3b en bound3b.tbl
'GETDELWAQID                 - Genereert het DELWAQ-ID gegeven het segmentnummer
'IDFROMSTRING                - extraheert een ID uit een string, gegeven een prefix en/of een afbreekstring
'REMOVEPOSTFIX               - verwijdert een postfix uit een string
'WRITESTOCHASTXMLFILE        - schrijft locaties en bijbehorende herhalingstijden en waterhoogtes weg in XML zodat de toetsingstool ze kan inlezen
'REPLACEDATESINSETTINGSDAT   - vervangt de start- en einddatum van een simulatie in het bestand settings.dat
'REPLACEDATESINDELFT3BINI    - vervangt de start- en einddatum van een simulatie in het bestand delft_3b.ini
'WriteUnpaved3BRecord        - schrijft een unpaved.3b record
'TrapeziumRangeToYZProfiles  - creëert dwarsprofielinformatie in YZ-formaat op basis van data in trapezium-kentallen (inclusief plasberm)

'overige modellen
'WRITEWAGMODINPUT            - schrijft een .dat file voor het wageningenmodel, met neerslag, verdamping en gemeten afvoeren (optioneel)
'READWAGMODOUTPUT            - leest een .00P-bestand van het Wageningen model in en schrijft de inhoud naar het actieve werkblad.
'WRITEPCRASTERXYZ            - schrijft een .xyz file ten behoeve van PCRASTER, die op zijn beurt weer een inundatiegrid kan opstellen

'meteofucties
'MAKKINKAVG                  - geeft voor een gegeven dag in het jaar de meerjarig gemiddelde potentiele gewasverdamping volgens Makkink terug
'DAYSTOHOURS                 - disaggregeert etmaalwaarden naar uurwaarden. Opties "none" (voor bijv. temperatuur) en "divide"
'EVAPDAYTOHOUR               - disaggregeert etmaalverdampingssommen tot uursommen, gebaseerd op een sinusoide
'NEERSLAGTEKORT              - berekent het neerslagtekort op basis van een tijdstap met neerslag, verdamping en het tekort van de vorige tijdstap
'HIRLAMTRANSLATE             - converteert HIRLAM-voorspellingsrasters met neerslag

'stringbewerkingen
'PARSESTRING                 - parst een string op basis van een te specificeren deelstring
'TEXTSNIPPET                 - deelt een string op in drie delen, gegeven twee karakterposities
'MULTIPARSE                  - parst in een keer het n'de element uit een string
'PARSENUMERIC                - parst net zo lang een karakter tot het volgende niet langer numeriek is
'BNASTRING                   - creeert een string voor een BNA-file. Vraagt om ID en X- en Y-coordinaat
'WAGMODSTASTRING             - creeert een meteo-string voor de .STA-file van het Wageningenmodel
'WALRUSDATSTRING             - creeert een meteo-string voor de .DAT-file van het WALRUS-model
'VERWIJDERDAGNAAMUITDATUM    - verwijdert de naam van de dag uit een string
'MAKEXMLTOKEN                - maakt van een tokenID en de waarde een tokenID="waarde" string
'STRINGPOSITIE               - geeft het positienummer van de eerstvoorkomende string van een opgegeven type op
'REPLACESTRING               - vervangt een opgegeven deelstring van een string door een andere string, dus niet op basis van positie
'REPLACESTRINGINALLFILES     - vervangt een string in alle files in de huidige directory, eventueel incl. subdirectories
'DOUBLEIDSINSTRINGCOLLECTION - checkt of een collectie met strings dubbele waarden bevat (boolean)
'TRIMUSINGCUSTOMSTRING       - voert een VBA.Trim uit met een opgegeven karakter ipv standaard de spatie
'UnifyString                 - uniformeert een string door te VBA.Trimmen en altijd de uppercase te gebruiken. Te gebruiken als Key in collections
'ISBANKNUMBER                - herkent of een string een bankrekeningnummer is
'MATCHWILDCARD               - checkt of een gegeven ID matcht met een gegeven structuur met wildcards
'CONCATENATECOMBINATIONS     - maakt een lijst met alle unieke stringcombinaties door te permuteren
'GetAttributeValueFromString - zoekt de waarde voor een gegeven attribuut op in een string. Bijv X in "X=508309.9 y = 4786842.744 Z=1368.026"

'importeren van bestanden
'READHMCZDATA                               - Leest waterstanden van het Hydro Meteo Centrum (ASCII formaat) in
'READASCIIGRID                              - Leest een Arc/Info grid in
'WriteASCIIGridFromEquation                 - Schrijft een Arc/Info grid op basis van z = ax + by + c
'WriteASCIIGridFromMultipleEquations        - Schrijft een Arc/Info grid op basis van meerdere z = ax + by + c vergelijkingen
'WRITEASCIIGRID                             - Schrijft een Arc/Info grid
'ASCII2XYZ                                  - converteert een Arc/Info grid naar een bestand met XYZ-waardes
'READMT940                                  - leest een MT940-file in, dat rekeningoverzichten bevat (o.a. ABN-AMRO)
'READENTIRETEXTFILE                         - leest de volledige inhoud van een tekstbestand naar het geheugen
'SetFilenameLengthByInsertingZeroes         - voegt nullen in in bestandsnamen om een gewenste lengte van de bestandsnaam te bewerkstelligen

'GIS-bewerkingen
'JoinNodes                   - maakt een nieuwe knoopID aan voor meerdere xy-knopen als ze dicht genoeg bijeen liggen
'FindNearestObjectInRange    - zoekt het ID van het dichstbijzijnde object uit een lijst (bijv. Meteo-stations) op basis van XY-coordinaten

'bestanden
'OPENSINGLEFILE              - open file dialog box
'LISTFILESINFOLDER           - produceert een collection van alle bestanden in een directory
'DIRECTORYEXISTS             - geeft terug of een directory bestaat
'CONTAINSKEY                 - geeft terug of een gegeven key onderdeel uitmaakt van een collection (WEKT NIET!!!!)
'CONTAINSKEY_BYOBJECTID      - geeft terug of een gegeven ID onderdeel uitmaakt van een collection met objecten die een element ID hebben
'DELETESHAPEFILE             - verwijdert een shapefile inclusief zijn bijbehorende bestanden (shx, dbf, shp)
'MOVEFILE                    - verplaatst een bestand van dir1 naar dir2
'DIRECTORYCOPY               - kopieert een directory incl. subdirs en inhoud naar een andere dir.
'FOLDERBROWSER               - presenteert een folder browser dialog
'REPLACEINFILE               - vervangt een opgegeven string overal in een tekstbestand
'REPLACEINSTRING             - vervangt een opgegeven string overal in een string

'Binaire functies
'Binary to Hex               - BinToHex(BinNum As String)
'Binary to Octal             - BinToOct(BinNum As String)
'Binary to Decimal           - BinToDec(BinNum As String)
'Hex to Binary               - HexToBin(HexNum As String)
'Octal to Binary             - OctToBin(OctNum As String)
'Decimal to Binary           - DecToBin(DecNum As String)


'overig
'RUNDOEVENTS                 - voert de optie doEvents uit voor een opgegeven aantal seconden, zodat andere processen even de ruimte krijgen
'SLEEP                       - laat de uitvoering van de macro een gespecificeerd aantal miliseconden wachten
'SHELLANDWAIT                - voert executables via de command line uit en wacht tot ze klaar zijn
'FINANCIELECATEGORIE         - rubriceert op basis van omschrijving uitgaven en inkomsten
'FILEEXISTS                  - controleert of een bestand bestaat.
'IB2011                      - berekent ruwweg de inkomstenbelasting voor 2011 op basis van opgegeven bruto inkomen


Public Enum enmKVDParameter
    dispersiecoefficient = 0
    locatieparameter = 1
    schaalparameter = 2
    vormparameter = 3
    waarde = 4
    terugkeertijd = 5
End Enum


Public Enum enmAggregateMethod
  Average = 1
  Most = 2
  Smallest = 3
  Largest = 4
  first = 5
  Last = 6
  sum = 7
End Enum

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400
Public Const pi As Variant = 3.141592

Public Function Interpolate(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant, X3 As Variant, Optional BlockInterpolate As Boolean = False) As Variant
'v1.78: MADE A DISTINCTION BETWEEN DATES AND NUMBERS
Dim Y3 As Variant 'de geïnterpoleerde waarde die we straks in de cel gaan zetten
If (IsDate(X1) And IsDate(X2) And IsDate(X3)) Then
    If X3 = X1 Then
        Interpolate = VAL(Y1)
        Exit Function
    ElseIf X3 = X2 Then
      Interpolate = VAL(Y2)
        Exit Function
    ElseIf X3 < Minimum(X1, X2) Then
        Y3 = -999
        Exit Function
    ElseIf X3 > Maximum(X1, X2) Then
        Y3 = -999
        Exit Function
    Else
        If BlockInterpolate = True Then
            Interpolate = VAL(Y1)
        Else
            Interpolate = VAL(Y1) + (VAL(Y2) - VAL(Y1)) / (X2 - X1) * (X3 - X1)
        End If
    End If
Else
    If VAL(X3) = VAL(X1) Then
        Interpolate = Y1
        Exit Function
    ElseIf VAL(X3) = VAL(X2) Then
      Interpolate = Y2
        Exit Function
    ElseIf X3 < Minimum(VAL(X1), VAL(X2)) Then
        Y3 = -999
        Exit Function
    ElseIf X3 > Maximum(VAL(X1), VAL(X2)) Then
        Y3 = -999
        Exit Function
    Else
        If BlockInterpolate = True Then
            Interpolate = VAL(Y1)
        Else
            Interpolate = VAL(Y1) + (VAL(Y2) - VAL(Y1)) / (VAL(X2) - VAL(X1)) * (VAL(X3) - VAL(X1))
        End If
    End If
End If


End Function

Public Function Count_Unique(MyRange As Range) As Integer
    Dim myCollection As Collection
    Dim r As Integer, c As Integer
    Set myCollection = New Collection
    For r = 1 To MyRange.Rows.Count
        For c = 1 To MyRange.Columns.Count
            If CollectionContainsKey(myCollection, MyRange.Cells(r, c)) = False Then
                Call myCollection.Add(MyRange.Cells(r, c), MyRange.Cells(r, c))
            End If
        Next
    Next
    Count_Unique = myCollection.Count
End Function

Public Function CollectionContainsKey(col As Collection, key As Variant)
Dim obj As Variant
On Error GoTo err
    CollectionContainsKey = True
    obj = col(key)
    Exit Function
err:
    CollectionContainsKey = False
End Function

Public Function GoalSeekRange(GoalCellsRange As Range, GoalValuesRange As Range, AdjustRange As Range) As Boolean
    If GoalCellsRange.Rows.Count <> GoalValuesRange.Rows.Count Or GoalCellsRange.Rows.Count <> AdjustRange.Rows.Count Then
        MsgBox ("Error: number of rows must be equal for all ranges that are passed to function GoalSeekRange.")
        GoalSeekRange = False
    ElseIf GoalCellsRange.Columns.Count > 1 Or GoalValuesRange.Columns.Count > 1 Or AdjustRange.Columns.Count > 1 Then
        MsgBox ("Error: each range passed to the function GoalSeekRange must only consist of one column.")
        GoalSeekRange = False
    Else
        Dim i As Integer
        For i = 1 To GoalCellsRange.Rows.Count
            GoalCellsRange.Cells(i, 1).GoalSeek Goal:=GoalValuesRange.Cells(i, 1), ChangingCell:=AdjustRange.Cells(i, 1)
        Next
    End If
End Function

Public Function CountDistinctValuesInRange(MyRange As Range) As Integer
    Dim myList() As Object
    Dim Initialized As Boolean
    Dim Found As Boolean
    Dim r As Integer
    Dim c As Integer
    Dim i As Integer
    For r = 1 To MyRange.Rows.Count
        For c = 1 To MyRange.Columns.Count
        
            If MyRange.Cells(r, c) <> "" Then
                
                Found = False
                
                If Initialized = False Then
                    ReDim myList(1)
                    Set myList(1) = MyRange.Cells(r, c)
                    Initialized = True
                Else
                    For i = 1 To UBound(myList)
                        If myList(i) = MyRange.Cells(r, c) Then
                            Found = True
                            Exit For
                        End If
                    Next
                                    
                    If Found = False Then
                        ReDim Preserve myList(UBound(myList) + 1)
                        Set myList(UBound(myList)) = MyRange.Cells(r, c)
                    End If
                
                End If
            End If
        Next
    Next
    CountDistinctValuesInRange = UBound(myList)
End Function

Public Function Extrapolate(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant, X3 As Variant) As Variant
'extrapolates linearly

Dim Y3 As Variant, Rico As Variant
If X3 > X2 Then
  Rico = (Y2 - Y1) / (X2 - X1)
  Extrapolate = Y2 + (X3 - X2) * Rico
ElseIf X3 < X1 Then
  Rico = (Y2 - Y1) / (X2 - X1)
  Extrapolate = Y1 - (X1 - X3) * Rico
Else
  Extrapolate = -999
End If

End Function

Public Function FitLinear_a(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Variant
  'creates a straight line between two XY-co-ordinates and returns a (from y = ax + b)
  FitLinear_a = (Y2 - Y1) / (X2 - X1)
End Function

Public Function FitLinear_b(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Variant
  Dim a As Variant
  a = (Y2 - Y1) / (X2 - X1)
  FitLinear_b = Y2 - a * X2
End Function

Public Function INTERPOLATEGAPSINRANGE(MyRange As Range, XCol As Integer, ValCol As Integer)
    Dim r As Long, r2 As Long
    Dim PrevX As Variant, NextX As Variant, x As Variant
    Dim PrevVal As Variant, NextVal As Variant
    
    If MyRange.Cells(1, ValCol) = "" Then
        MsgBox ("Error interpolating range: first and last cell cannot be empty.")
    ElseIf MyRange.Cells(MyRange.Rows.Count, ValCol) = "" Then
        MsgBox ("Error interpolating range: first and last cell cannot be empty.")
    Else
        For r = 2 To MyRange.Rows.Count - 1
            If MyRange.Cells(r, ValCol) = "" Then
                x = MyRange.Cells(r, XCol)
                PrevX = MyRange.Cells(r - 1, XCol)
                PrevVal = MyRange.Cells(r - 1, ValCol)
                For r2 = r + 1 To MyRange.Rows.Count
                    If MyRange.Cells(r2, ValCol) <> "" Then
                        NextX = MyRange.Cells(r2, XCol)
                        NextVal = MyRange.Cells(r2, ValCol)
                        Exit For
                    End If
                Next
                MyRange.Cells(r, ValCol) = Interpolate(PrevX, PrevVal, NextX, NextVal, x)
                MyRange.Cells(r, ValCol).Interior.ColorIndex = 37
            End If
        Next
    End If
    
End Function

Public Function FillNodataValuesInRangeByLastValue(XYRange As Range, XCol As Integer, YCol As Integer, NodataValue As Variant) As Boolean
    'this function fills up nodata-values in a given range by the previous value found
    Dim r As Long
    Dim r2 As Long
    For r = 1 To XYRange.Rows.Count
        If XYRange.Cells(r, YCol) = NodataValue Then
            For r2 = r - 1 To 1 Step -1
              If Not XYRange.Cells(r2, YCol) = NodataValue Then
                XYRange.Cells(r, YCol) = XYRange.Cells(r2, YCol)
                XYRange.Cells(r, YCol).Interior.Color = vbYellow
                Exit For
              End If
            Next
        End If
    Next
End Function

Public Function InterpolateNodataValuesInRange(XYRange As Range, XCol As Integer, YCol As Integer, NodataValue As Variant, Optional ByVal InterpolateWhenAscending As Boolean = True, Optional ByVal InterpolateWhenDescending As Boolean = True) As Boolean
    Dim r As Long
    Dim r2 As Long
    Dim X1 As Variant, Y1 As Variant
    Dim X2 As Variant, Y2 As Variant
    Dim PrevFound As Boolean
    Dim NextFound As Boolean
    
    For r = 1 To XYRange.Rows.Count
        If XYRange.Cells(r, YCol) = NodataValue Then
            PrevFound = False
            NextFound = False
            'walk back until a valid cell is found
            For r2 = r - 1 To 1 Step -1
                If Not XYRange.Cells(r2, YCol) = NodataValue Then
                    PrevFound = True
                    X1 = XYRange.Cells(r2, XCol)
                    Y1 = XYRange.Cells(r2, YCol)
                    Exit For
                End If
            Next
            'walk forward until a valid cell is found
            For r2 = r + 1 To XYRange.Rows.Count
                If Not XYRange.Cells(r2, YCol) = NodataValue Then
                    NextFound = True
                    X2 = XYRange.Cells(r2, XCol)
                    Y2 = XYRange.Cells(r2, YCol)
                    Exit For
                End If
            Next
            If PrevFound And NextFound Then
                If Y2 >= Y1 Then
                    If InterpolateWhenAscending Then
                        XYRange.Cells(r, YCol) = Interpolate(X1, Y1, X2, Y2, XYRange.Cells(r, XCol))
                        XYRange.Cells(r, YCol).Interior.Color = vbYellow
                    End If
                ElseIf Y2 <= Y1 Then
                    If InterpolateWhenDescending Then
                        XYRange.Cells(r, YCol) = Interpolate(X1, Y1, X2, Y2, XYRange.Cells(r, XCol))
                        XYRange.Cells(r, YCol).Interior.Color = vbYellow
                    End If
                End If
            End If
                        
        End If
    Next
End Function

Public Function InterpolateFromRange(x As Variant, MyRange As Range, XColIdx As Integer, YColIdx As Integer, Optional ExtrapolateBelow As Boolean = True, Optional ExtrapolateAbove As Boolean = True, Optional BlockInterpolation As Boolean = False, Optional CheckIfAscending As Boolean = True) As Variant
  Dim r As Long, r2 As Long, startr As Long, stepsize As Long
    If CheckIfAscending Then
    If Not IsRangeAscending(MyRange) Then
      InterpolateFromRange = "Error: column containing X values must be ascending."
      Exit Function
    End If
 End If

If x < MyRange.Cells(1, XColIdx).value Then
  If ExtrapolateBelow = True Then
    InterpolateFromRange = MyRange(1, YColIdx).value
    Exit Function
  Else
    InterpolateFromRange = 0
    Exit Function
  End If
ElseIf x >= MyRange.Cells(MyRange.Rows.Count, XColIdx).value Then
  If ExtrapolateAbove = True Then
    InterpolateFromRange = MyRange.Cells(MyRange.Rows.Count, YColIdx).value
    Exit Function
  Else
    InterpolateFromRange = 0
    Exit Function
  End If
ElseIf MyRange.Rows.Count > 1 Then
    If MyRange.Count > 100000 Then
    stepsize = 10000
  ElseIf MyRange.Rows.Count > 10000 Then
    stepsize = 1000
  ElseIf MyRange.Rows.Count > 1000 Then
    stepsize = 100
  ElseIf MyRange.Rows.Count > 100 Then
    stepsize = 10
  Else
    stepsize = 1
  End If

 'find the appropriate block
  For r = stepsize + 1 To MyRange.Rows.Count Step stepsize
    If MyRange(r, XColIdx).value > x Then
      'apparently our value is located in the previous block. So step back to the beginning of that block and proceed
      startr = r - stepsize
      For r2 = startr To MyRange.Rows.Count
        If x >= MyRange.Cells(r2, XColIdx).value And x <= MyRange.Cells(r2 + 1, XColIdx).value Then
          InterpolateFromRange = Interpolate(MyRange.Cells(r2, XColIdx).value, MyRange.Cells(r2, YColIdx).value, MyRange.Cells(r2 + 1, XColIdx).value, MyRange.Cells(r2 + 1, YColIdx).value, x, BlockInterpolation)
          Exit Function
        End If
      Next
    End If
  Next
  'if we end up here, our value must be located somewhere in the last block. walk backwards to find it
  For r = MyRange.Rows.Count To 2 Step -1
    If x >= MyRange.Cells(r - 1, XColIdx).value And x <= MyRange.Cells(r, XColIdx).value Then
        InterpolateFromRange = Interpolate(MyRange.Cells(r - 1, XColIdx).value, MyRange.Cells(r - 1, YColIdx).value, MyRange.Cells(r, XColIdx).value, MyRange.Cells(r, YColIdx).value, x, BlockInterpolation)
        Exit Function
    End If
  Next
  
Else
  InterpolateFromRange = "Error: outside range."
End If

End Function


Public Function InterpolateRangeFromRange(XYRange As Range, ResultsRange As Range, Optional BlockInterpolation As Boolean = False) As Variant
  
  Dim i As Long, j As Long
  Dim r As Long, c As Long
  Dim XLookup As Variant, CurX As Variant, NextX As Variant
  
  'first some checks:
  If XYRange.Columns.Count <> 2 Then
    MsgBox ("Input range must consist of two columns: one containing X-values; one containing Y-values")
  ElseIf ResultsRange.Columns.Count <> 2 Then
    MsgBox ("Results range must consist of two columns: one containing X-values; one for the computed Y-values")
  ElseIf RANGEVERTASCENDING(XYRange) = False Then
    MsgBox ("Input range must be ascending.")
  Else
  
    'read the input range
    Dim XYData As Variant
    ReDim XYData(XYRange.Rows.Count, XYRange.Columns.Count)
    XYData = XYRange
  
    'read the output range
    Dim results As Variant
    ReDim results(ResultsRange.Rows.Count, ResultsRange.Columns.Count)
    results = ResultsRange
    
    For r = 1 To UBound(results, 1)
      XLookup = results(r, 1)
      
      If XLookup < XYData(1, 1) Then
        results(r, 2) = XYData(1, 2)
      ElseIf XLookup > XYData(UBound(XYData, 1), 1) Then
        results(r, 2) = XYData(UBound(XYData, 1), 2)
      Else
        For i = 1 To UBound(XYData, 1)
          CurX = XYData(i, 1)
          NextX = XYData(i + 1, 1)
          
          If CurX <= XLookup And NextX >= XLookup Then
            results(r, 2) = Interpolate(XYData(i, 1), XYData(i, 2), XYData(i + 1, 1), XYData(i + 1, 2), XLookup, BlockInterpolation)
            Exit For
          End If
        Next
      End If
    Next
  End If
  
  Call PrintArray(results, ResultsRange)

End Function

Public Function InterpolateFromRangePlus(id As String, x As Variant, IDRange As Range, XRange As Range, YRange As Range, Optional ExtrapolateBelow As Boolean = True, Optional ExtrapolateAbove As Boolean = True, Optional BlockInterpolation As Boolean = False, Optional CheckIfAscending As Boolean = True) As Variant
  Dim r As Long, r2 As Long, startr As Long, stepsize As Long
  If IDRange.Count <> XRange.Count Then
    InterpolateFromRangePlus = "Error: ID and X range must be of equal size."
    Exit Function
  ElseIf XRange.Count <> YRange.Count Then
    InterpolateFromRangePlus = "Error: X and Y range must be of equal size."
    Exit Function
  ElseIf XRange.Columns.Count > 1 Then
    InterpolateFromRangePlus = "Error: column for X values must consist of one column."
    Exit Function
  ElseIf YRange.Columns.Count > 1 Then
    InterpolateFromRangePlus = "Error: column for Y values must consit of one column."
    Exit Function
  End If
  
  Dim startRow As Long, endRow As Long
  Dim startfound As Boolean
    
  'first find the start- and endrow for the given ID
  For r = 1 To IDRange.Count
    If UCase(Trim(IDRange(r, 1))) = UCase(Trim(id)) And startfound = False Then
      startfound = True
      startRow = r
    ElseIf startfound = True And IDRange(r, 1) <> id Then
      endRow = r - 1
      Exit For
    End If
  Next
  
  If x <= XRange(startRow, 1).value Then
    If ExtrapolateBelow = True Then
      InterpolateFromRangePlus = YRange(startRow, 1).value
      Exit Function
    Else
      InterpolateFromRangePlus = 0
      Exit Function
    End If
  ElseIf x >= XRange(endRow, 1).value Then
    If ExtrapolateAbove = True Then
      InterpolateFromRangePlus = YRange(endRow, 1).value
      Exit Function
    Else
      InterpolateFromRangePlus = 0
      Exit Function
    End If
  ElseIf (endRow - startRow) > 1 Then
    For r = startRow To endRow
      If XRange(r, 1).value > x Then
        InterpolateFromRangePlus = Interpolate(XRange(r - 1, 1).value, YRange(r - 1, 1).value, XRange(r, 1).value, YRange(r, 1).value, x, BlockInterpolation)
        Exit Function
      End If
    Next
  Else
    InterpolateFromRangePlus = "Error: outside range."
 End If

End Function

Public Function KleinsteKwadratenMethode(GemetenDatum As Range, GemetenWaarden, BerekendDatum As Range, BerekendWaarden As Range) As Variant

'deze functie berekent het kleinstekwadratenverschil tussen een berekende en gemeten reeks
Dim d As Variant, V As Variant
Dim d1 As Variant, v1 As Variant
Dim D2 As Variant, v2 As Variant
Dim v3 As Variant
Dim sum As Variant
Dim r As Long, c As Long
Dim r2 As Long, c2 As Long

sum = 0
For r = 1 To GemetenDatum.Rows.Count
  d = GemetenDatum.Cells(r, 1)
  V = GemetenWaarden.Cells(r, 1)
  
  If d >= BerekendDatum.Cells(1, 1) And d <= BerekendDatum.Cells(BerekendDatum.Rows.Count, 1) Then
  
    For r2 = 1 To BerekendDatum.Rows.Count - 1
      d1 = BerekendDatum.Cells(r2, 1)
      v1 = BerekendWaarden.Cells(r2, 1)
      D2 = BerekendDatum.Cells(r2 + 1, 1)
      v2 = BerekendWaarden.Cells(r2 + 1, 1)
      
      If d1 <= d And D2 >= d Then
        v3 = Interpolate(d1, v1, D2, v2, d)
        sum = sum + (v3 - V) ^ 2
        Exit For
      End If
    Next
  
  End If
Next
KleinsteKwadratenMethode = sum

End Function

Function IsStringArrayEmpty(anArray() As String)

Dim i As Integer
On Error Resume Next
i = UBound(anArray, 1)
If err.Number = 0 Then
    IsStringArrayEmpty = False
Else
    IsStringArrayEmpty = True
End If

End Function

Public Function GETARRAYSORTIDX(myArr() As Variant) As Long()
    Dim IdxArr() As Long, DoneArr() As Boolean, i As Long
    ReDim IdxArr(LBound(myArr), UBound(myArr()))
    ReDim DoneArr(LBound(myArr()), UBound(myArr()))
        
    For i = LBound(myArr()) To UBound(myArr())
      IdxArr(i) = GETMAXIDXFROMARRAY(myArr, DoneArr)
      DoneArr(i) = True
    Next
    GETARRAYSORTIDX = IdxArr

End Function

Public Function GETMAXIDXFROMARRAY(ByRef myArr() As Variant, ByRef DoneArr() As Boolean) As Long
  Dim i As Long, myMax As Variant, myMaxIdx As Long
  myMax = -999999999
  For i = LBound(myArr()) To UBound(myArr())
    If myArr(i) >= myMax And DoneArr(i) = False Then
      myMax = myArr(i)
      myMaxIdx = i
    End If
  Next
  GETMAXIDXFROMARRAY = myMaxIdx
End Function


' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function SortCollection(ByVal c As Collection) As Collection
   Dim n As Long: n = c.Count
   If n = 0 Then Set SortCollection = New Collection: Exit Function
   ReDim Index(0 To n - 1) As Long                    ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(i) = i + 1: Next         ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                     ' generate ordered heap
      Heapify c, Index, i, n
      Next
   For m = n To 2 Step -1                             ' sort the index array
      Exchange Index, 0, m - 1                        ' move highest element to top
      Heapify c, Index, 0, m - 1
      Next
   Dim c2 As New Collection
   For i = 0 To n - 1: c2.Add c.Item(Index(i)): Next  ' fill output collection
   Set SortCollection = c2
End Function
   
' Heapsort routine.
' Returns a sorted Index array for the Keys array.
' Author: Christian d'Heureuse (www.source-code.biz)
Public Function HeapSort(Keys) As Long()
   Dim Base As Long: Base = LBound(Keys)                    ' array index base
   Dim n As Long: n = UBound(Keys) - LBound(Keys) + 1       ' array size
   Dim Index() As Long
   ReDim Index(Base To Base + n - 1) As Long                ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(Base + i) = Base + i: Next     ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                           ' generate ordered heap
      Heapify Keys, Index, i, n
      Next
   For m = n To 2 Step -1
      Exchange Index, 0, m - 1                              ' move highest element to top
      Heapify Keys, Index, 0, m - 1
   Next
   HeapSort = Index
End Function

Public Function SortCollectionOfLongByKey(myCollection As Collection) As Long()
  'in order to sort a collection of items by its key we'll first create an array that contains all keys
  'then we'll sort that array using Christian d'Heureuse's Heapsort-routine, which we'll return
  'this means that the function will return an array that contains the index numbers for the sorted keys
  
  'IMPORTANT: IN VBA it is NOT possible to retrieve the actual key. Therefore make sure you also store the key
  'as an element of the object within the collection!
  
  Dim SortMe() As Variant, i As Long
  ReDim SortMe(1 To myCollection.Count)
    
  For i = 1 To myCollection.Count
    SortMe(i) = myCollection.Item(i).key
  Next
  
  SortCollectionByKey = HeapSort(SortMe)

End Function



Public Function Random(lowerbound As Integer, upperbound As Integer) As Integer
  'geeft een random getal terug tussen twee gespecificeerde boundarywaaren (hele getallen)
  Randomize
  Random = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Public Function RandomDouble(lowerbound As Variant, upperbound As Variant) As Variant
  
  'creeer een random integer tussen 0 en 32000
  Dim myRnd As Integer
  myRnd = Random(0, 32000)
  
  'transformeer deze terug naar een waarde tussen min en max
  RandomDouble = lowerbound + myRnd / 32000 * (upperbound - lowerbound)

End Function



Public Function Maximum(val1 As Variant, val2 As Variant) As Variant
  If val1 > val2 Then
    Maximum = val1
  Else
    Maximum = val2
  End If
End Function

Public Function Minimum(val1 As Variant, val2 As Variant) As Variant
  If val1 < val2 Then
    Minimum = val1
  Else
    Minimum = val2
  End If
End Function

Public Function ARRAYFROMWORKSHEET(sheetName As String, startRow As Long, StartCol As Long, EndCol As Long) As Variant()
  'Author: Siebe Bosch
  'Date : 1-9-2013
  'Description: extracts data from a worksheet and puts it into an array
  Dim curSheet As String, i As Long, j As Long, r As Long, c As Long
  Dim endRow As Long
  Dim myArray() As Variant
  curSheet = ActiveSheet.Name
  Worksheets(sheetName).Activate
  r = startRow - 1
  
  'find the last record
  While Not ActiveSheet.Cells(r + 1, StartCol) = ""
    r = r + 1
  Wend
  endRow = r
  ReDim myArray(1 To endRow - startRow + 1, 1 To EndCol - StartCol + 1)
      
  r = 0
  For i = startRow To endRow
    c = 0
    r = r + 1
    For j = StartCol To EndCol
      c = c + 1
      myArray(r, c) = ActiveSheet.Cells(i, j)
    Next
  Next

  Worksheets(curSheet).Activate
  ARRAYFROMWORKSHEET = myArray

End Function

Public Sub ARRAYVARIANTTOWORKSHEET(sheetName As String, myArray() As Variant, startRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(sheetName).Activate
    
  'write the data to the worksheet
  r = startRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    For j = 1 To UBound(myArray, 2)
      c = c + 1
      ActiveSheet.Cells(r, c) = myArray(i, j)
    Next
  Next
  
  Worksheets(curSheet).Activate

End Sub

Public Sub ARRAYDATETOWORKSHEET(sheetName As String, myArray() As Date, startRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(sheetName).Activate
    
  'write the data to the worksheet
  r = startRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    ActiveSheet.Cells(r, c) = myArray(i)
  Next
  
  Worksheets(curSheet).Activate

End Sub

Public Sub ARRAYSINGLETOWORKSHEET(sheetName As String, myArray() As Single, startRow As Long, StartCol As Long)
  
  Dim curSheet As String, Header As String, r As Long, c As Long, i As Long, j As Long
  curSheet = ActiveSheet.Name
  Worksheets(sheetName).Activate
    
  'write the data to the worksheet
  r = startRow - 1
  c = StartCol
  For i = 1 To UBound(myArray, 1)
    r = r + 1
    c = StartCol
    ActiveSheet.Cells(r, c) = myArray(i)
  Next
  
  Worksheets(curSheet).Activate

End Sub


Public Sub TIMESERIES2ARRAYS(MyRange As Range, ByRef dates() As Date, ByRef Vals() As Single)
  Dim r As Long
  ReDim dates(1 To MyRange.Rows.Count)
  ReDim Vals(1 To MyRange.Rows.Count)
  
  For r = 1 To MyRange.Rows.Count
    dates(r) = MyRange.Cells(r, 1)
    Vals(r) = MyRange.Cells(r, 2)
  Next

End Sub

Public Function USERSELECTRANGE() As Range
  Set USERSELECTRANGE = Application.InputBox(Prompt:="Please Select Range", Title:="Range Select", Type:=8)
  Call Application.GoTo(USERSELECTRANGE)
End Function

Public Function GetColumnName(ByVal CellAddress As String) As String
    'returns the column name for a given cell address. Courtesy of ChatGPT
    Dim i As Integer
    Dim ColumnName As String
    
    ' Loop through each character in the cell address
    For i = 1 To Len(CellAddress)
        If IsNumeric(Mid(CellAddress, i, 1)) Then
            Exit For
        Else
            ColumnName = ColumnName & Mid(CellAddress, i, 1)
        End If
    Next i
    
    ' Return the column name (letter) of the cell
    GetColumnName = ColumnName
End Function


Public Function ShiftCellAddress(ByVal CellAddress As String, ByVal ColumnShift As Integer, ByVal RowShift As Integer) As String
    'shifts a given cell address by a number of rows and columns. Courtesy of ChatGPT
    Dim OriginalCell As Range
    Dim NewCell As Range
    
    ' Get the original cell based on the input cell address
    Set OriginalCell = Range(CellAddress)
    
    ' Calculate the new cell address by shifting the column and row
    Set NewCell = OriginalCell.Offset(RowShift, ColumnShift)
    
    ' Return the address of the new cell
    ShiftCellAddress = NewCell.Address(False, False)
End Function


Public Function RANGEADDRESSFROMRC(ByRef sheetName As String, R1 As Long, C1 As Long, r2 As Long, c2 As Long) As Range
  Dim MySheet As Worksheet
  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = sheetName Then
      Set RANGEADDRESSFROMRC = Worksheets(i).Range(Cells(R1, C1).Address, Cells(r2, c2).Address)
    End If
  Next
End Function

Public Function RANGECOLIDXFROMMAXVAL(MyRange As Range) As Long
  'Returns the column indexnumber for the largest value in a given range
  'note: this is not the worksheet column number but the column number from within the range!
  Dim myMax As Variant, myVal As Variant, myCol As Long
  Dim r As Long, c As Long
  myMax = -99999999999999#
  
  For c = 1 To MyRange.Columns.Count
    For r = 1 To MyRange.Rows.Count
      myVal = MyRange.Cells(r, c).value
      If myVal >= myMax Then
        myCol = c
        myMax = myVal
      End If
    Next
  Next
  RANGECOLIDXFROMMAXVAL = myCol

End Function

Public Function ASSIGNVALUEBYMONTH(myDate As Date, Jan As Variant, Feb As Variant, Mar As Variant, Apr As Variant, May As Variant, Jun As Variant, Jul As Variant, Aug As Variant, Sep As Variant, Oct As Variant, Nov As Variant, Dec As Variant) As Variant
  Dim myMonth As Integer, myVal As Variant
  myMonth = month(myDate)
  Select Case myMonth
    Case Is = 1
      myVal = Jan
    Case Is = 2
      myVal = Feb
    Case Is = 3
      myVal = Mar
    Case Is = 4
      myVal = Apr
    Case Is = 5
      myVal = May
    Case Is = 6
      myVal = Jun
    Case Is = 7
      myVal = Jul
    Case Is = 8
      myVal = Aug
    Case Is = 9
      myVal = Sep
    Case Is = 10
      myVal = Oct
    Case Is = 11
      myVal = Nov
    Case Is = 12
      myVal = Dec
  End Select
  ASSIGNVALUEBYMONTH = myVal
End Function

Public Function CLASSIFYNUMBERBYCLASS(myVal As Variant, ByVal minVal As Variant, ByVal maxVal As Variant, ByVal classSize As Variant) As String
  Dim Done As Boolean, curVal As Variant
  curVal = minVal
  
  If myVal < curVal Then
    CLASSIFYNUMBERBYCLASS = "< " & curVal
    Done = True
    Exit Function
  End If
  
  While Not Done
    curVal = curVal + classSize
    If myVal < curVal Then
      CLASSIFYNUMBERBYCLASS = curVal - classSize & " - " & curVal
      Done = True
      Exit Function
    End If
  Wend
  
  If Done = False Then
      CLASSIFYNUMBERBYCLASS = curVal - classSize & " - " & curVal
  End If
  
End Function

Public Sub DESIGNTEMPORALBLOCKPATTERNS(nTimesteps As Integer, nBlocks As Integer, PercentageIncrement As Integer, startRow As Integer, endRow As Integer)
    Dim PercentageSum As Variant
    Dim Mileage() As Integer, i As Integer, sum As Integer
    Dim r As Integer, c As Integer
    r = startRow
    c = StartCol

    If Math.Round(nTimesteps / nBlocks, 0) <> nTimesteps / nBlocks Then
        MsgBox ("Error: choose another number of blocks.")
    Else
        Dim nBlockSteps As Integer, Percentage As Integer
        nBlockSteps = nTimesteps / nBlocks
        ReDim Mileage(1 To nBlocks)
                        
        While MileageOneUp(0, 100, Mileage)
            
            sum = 0
            For i = 1 To Mileage.Count
                sum = sum + Mileage(i)
            Next
            
            If sum = 100 Then
                c = c + 1
                r = startRow
                For iblock = 1 To nBlocks
                    For iStep = 1 To nBlockSteps
                        r = r + 1
                        ActiveSheet.Cells(r, StartCol) = "Fraction"
                        ActiveSheet.Cells(r, c) = Mileage(iblock)
                    Next
                Next
                
            End If
            
        Wend
                        
                    
    End If
    
    
End Sub

Public Sub POINTOBJECTSTOWEBVIEWER(IDCol As Integer, XCol As Integer, YCol As Integer, DescriptionCol As Integer, ResultsPath As String, VariableName As String)
    Dim r As Long, c As Long, i As Integer, fn As Integer
    Dim id As String, VALUENAME As String, VAL As String, x As Double, y As Double
    Dim lat As Double, lon As Double
    Dim introstr As String, propstr As String, geostr As String
    fn = FreeFile
    Open ResultsPath For Output As #fn

    Print #fn, "let " & VariableName & " = {"
    Print #fn, vbTab & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "FeatureCollection" & Chr(34) & ","
    Print #fn, vbTab & Chr(34) & "crs" & Chr(34) & ": { " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "name" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": { " & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & "urn:ogc:def:crs:EPSG::3857" & Chr(34) & " } },"
    Print #fn, vbTab & Chr(34) & "features" & Chr(34) & ": ["

    r = 1
    While Not ActiveSheet.Cells(r + 1, IDCol) = ""
        r = r + 1
        id = ActiveSheet.Cells(r, IDCol)
        x = ActiveSheet.Cells(r, XCol)
        y = ActiveSheet.Cells(r, YCol)
        lat = RD2LAT(x, y)
        lon = RD2LON(x, y)
                
        introstr = vbTab & vbTab & "{ " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Feature" & Chr(34) & ", "
        propstr = Chr(34) & "properties" & Chr(34) & ": { " & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & id & Chr(34)
        geostr = Chr(34) & "geometry" & Chr(34) & ": { " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Point" & Chr(34) & ", " & Chr(34) & "coordinates" & Chr(34) & ": [ " & lon & "," & lat & "] }"
        
        propstr = propstr & ", " & Chr(34) & "description" & Chr(34) & ":" & Chr(34) & VBA.Replace(ActiveSheet.Cells(r, DescriptionCol), Chr(34), "_") & Chr(34)
                
        'For i = 1 To ValueCols.Count
        '    VALUENAME = ActiveSheet.Cells(1, ValueCols(i))
        '    If ActiveSheet.Cells(r, ValueCols.Item(i)) <> "" And VBA.IsNumeric(ActiveSheet.Cells(r, ValueCols.Item(i))) Then
        '        propstr = propstr & ", " & Chr(34) & VALUENAME & Chr(34) & ":" & ActiveSheet.Cells(r, ValueCols.Item(i))
        '    Else
        '       VAL = VBA.Replace(ActiveSheet.Cells(r, ValueCols.Item(i)), Chr(34), "_")                'attributes cannot contain double quotes so replace them by underscores
        '       propstr = propstr & ", " & Chr(34) & VALUENAME & Chr(34) & ":" & Chr(34) & VAL & Chr(34)
        '    End If
        'Next
        
        propstr = propstr & "}, "
        If ActiveSheet.Cells(r + 1, IDCol) = "" Then
            geostr = geostr & "}"
        Else
            geostr = geostr & "},"
        End If
        Print #fn, introstr & propstr & geostr
        
    Wend
        
    Print #fn, vbTab & "]"
    Print #fn, "}"
    
    Close (fn)
    

End Sub

Public Sub OBSERVATIONSTOWEBVIEWER(FromDate As Date, ToDate As Date, IDRow As Integer, AliasRow As Integer, ModelIDRow As Integer, ParameterRow As Integer, NodataValue As Double, nDecimals As Integer, OnlyWriteValueWhenChanged As Boolean, ResultsPath As String)
    'assume the first colummn contains date/time
    Dim r As Long, c As Long
    Dim id As String, Alias As String, ModelID As String
    Dim CurDate As Date, curVal As Double
    Dim valuesstr As String, datesstr As String
    Dim datestr As String, valuestr As String
    Dim Parameter As String         'can be h or Q
    Dim lastVal As Double
    Dim fn As Long
    
    fn = FreeFile
    Open ResultsPath For Output As #fn
    
    Print #fn, "let measurements = {"
    Print #fn, "    " & Chr(34) & "locations" & Chr(34) & ":["
    
    r = 0
    c = 1
    While Not ActiveSheet.Cells(1, c + 1) = ""
        c = c + 1
        r = Maximum(IDRow, AliasRow)           'set the starting position for the row counter
        r = Maximum(r, ModelIDRow)
        r = Maximum(r, ParameterRow)
        
        'a new location
        lastVal = NodataValue                       'initialize the last value as the nodata-value
        id = ActiveSheet.Cells(IDRow, c)
        Alias = ActiveSheet.Cells(AliasRow, c)
        ModelID = ActiveSheet.Cells(ModelIDRow, c)
        Parameter = ActiveSheet.Cells(ParameterRow, c)
                        
        Print #fn, vbTab & vbTab & "{"
        Print #fn, vbTab & vbTab & vbTab & Chr(34) & "ID" & Chr(34) & ":" & Chr(34) & id & Chr(34) & ","
        Print #fn, vbTab & vbTab & vbTab & Chr(34) & "Alias" & Chr(34) & ":" & Chr(34) & Alias & Chr(34) & ","
        Print #fn, vbTab & vbTab & vbTab & Chr(34) & "ModelID" & Chr(34) & ":" & Chr(34) & ModelID & Chr(34) & ","
        Print #fn, vbTab & vbTab & vbTab & Chr(34) & Parameter & Chr(34) & ": {"
        datesstr = vbTab & vbTab & vbTab & vbTab & Chr(34) & "dates" & Chr(34) & ": ["
        valuesstr = vbTab & vbTab & vbTab & vbTab & Chr(34) & "values" & Chr(34) & ": ["
        
        While Not ActiveSheet.Cells(r + 1, 1) = ""
            r = r + 1
            CurDate = ActiveSheet.Cells(r, 1)
            If CurDate >= FromDate And CurDate <= ToDate Then
                If Not ActiveSheet.Cells(r, c) = "" Then
                    curVal = Math.Round(ActiveSheet.Cells(r, c), nDecimals)
                    If Not curVal = NodataValue Then
                        If OnlyWriteValueWhenChanged = False Or curVal <> lastVal Then
                            datestr = Chr(34) & Application.WorksheetFunction.Text(CurDate, "yyyy-mm-ddThh:MM:ss.000Z") & Chr(34)
                            datesstr = datesstr & datestr & ","
                            valuesstr = valuesstr & curVal & ","
                            lastVal = curVal
                        End If
                    End If
                End If
            End If
        Wend
        
        Print #fn, Left(datesstr, Len(datesstr) - 1) & "],"
        Print #fn, Left(valuesstr, Len(valuesstr) - 1) & "]"
                
        Print #fn, vbTab & vbTab & vbTab & "}"
        If Not ActiveSheet.Cells(ModelIDRow, c + 1) = "" Then
            Print #fn, vbTab & vbTab & "},"
        Else
            Print #fn, vbTab & vbTab & "}"
        End If
        
    Wend
    
    Print #fn, vbTab & "]"
    Print #fn, "}"
    
    Close (fn)
    
    

End Sub

Public Sub DESIGNTWOPARTTEMPORALPATTERNS(nTimesteps As Integer, PercentageIncrement As Integer, TimestepIncrement As Integer, startRow As Integer, StartCol As Integer)
    'this function designs temporal patterns that consist of two lineair sections: start to mid and mid to end
    'hereby the start and end do NOT necessarily have to be 0
    'we'll start by incrementing the centerpoint in time
    
    '...............´;&%/.I%%&?í´...................................................'
    '...........,»%%%%%%/.I%%%%%%%%%&?í,............................................
    '.......^=%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%=;'.....................................
    '.....,''''''''''''',.I%%%%%%%%%%%%%%%%%%%%%%%%*/'..............................
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/^.......................
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%I»~´...............
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&?~´........
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%&=í,.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%=.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    '.....=%%%%%%%%%%%%%/.I%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%?.
    Dim ts As Integer
    Dim LeftTSFraction As Variant, RightTSFraction As Variant
    Dim p1Done As Boolean, p2Done As Boolean, p3Done As Boolean
    Dim p1 As Variant, p2 As Variant, p3 As Variant
    Dim r As Integer, c As Integer
    Dim SecondPartRemaining As Variant
    Dim patnum As Integer, Check As Variant
    
    c = StartCol
    
    For ts = TimestepIncrement To (nTimesteps - TimestepIncrement) Step TimestepIncrement
        'now we can iterate through all possibilities for the first point
            
        LeftTSFraction = ts / nTimesteps
        RightTSFraction = (nTimesteps - ts) / nTimesteps
            
        p1Done = False
        While Not p1Done
            
            'now that p1 is known we can iterate through p2 (the center point)
            p2Done = False
            p2 = 0
            While Not p2Done
                            
              'now that p2 is known we can calculate p3 (right point)
              'each part consists of a rectangle + a triangle on top
              Dim FirstPart As Variant, SecondPart As Variant
              FirstPart = TrapeziumArea(p1, p2, LeftTSFraction)
              SecondPart = 1 - FirstPart
              If SecondPart < 0 Then
                'the current combination of p1 with all possible values for p2 has been fully investigated, so move on
                p2Done = True
                p1Done = True
              Else
                'first decide if p2 is the hightest or lowest for our second part
                If p2 * RightTSFraction <= SecondPart Then
                    'we now know that p3 must lie higher than p2
                    SecondPartRemaining = SecondPart - RightTSFraction * p2
                    '(p3 - p2) * RightTSFraction = SecondPartRemaining
                    'p3 - p2 = SecondPartRemaining / RightTSFraction
                    'p3 = p2 + SecondPartRemaining / RightTSFraction
                    p3 = p2 + SecondPartRemaining / RightTSFraction * 2
                Else
                    'we now know that p3 lies lower than p2
                    '(p2 * RightTSFraction) + (p3 - p2) * RightTSFraction = SecondPart
                    'RightTSFraction * (p2 + (p3 - p2)/2) = SecondPart
                    '(p2 + (p3 - p2)/2) = SecondPart/RightTSFraction
                    '(p3 - p2)/2 = SecondPart/RightTSFraction - p2
                    'p3/2 - p2/2 = SecondPart/RightTSFraction - p2
                    'p3/2 = SecondPart/RightTSFraction - p2 + p2/2
                    'p3/2 = SecondPart/RightTSFraction - p2/2
                    'p3 = 2 * (SecondPart/RightTSFraction - p2/2)
                    p3 = 2 * (SecondPart / RightTSFraction - p2 / 2)
                End If
                
                'final check!
                Check = TrapeziumArea(p1, p2, LeftTSFraction) + TrapeziumArea(p2, p3, RightTSFraction)
                If Round(Check, 2) <> 1 Then Stop
                
                If p3 < 0 Then
                    'for the current p2 we have reached the limit regarding p3
                    p2Done = True
                Else
                    'write the result!
                    patnum = patnum + 1
                    ActiveSheet.Cells(startRow, StartCol) = "ID"
                    ActiveSheet.Cells(startRow, c) = "pattern" & patnum
                    ActiveSheet.Cells(startRow + 1, StartCol) = "timestep peak"
                    ActiveSheet.Cells(startRow + 1, c) = ts
                    ActiveSheet.Cells(startRow + 2, StartCol) = "fractionFirst"
                    ActiveSheet.Cells(startRow + 2, c) = p1
                    ActiveSheet.Cells(startRow + 3, StartCol) = "fractionPeak"
                    ActiveSheet.Cells(startRow + 3, c) = p2
                    ActiveSheet.Cells(startRow + 4, StartCol) = "fractionLast"
                    ActiveSheet.Cells(startRow + 4, c) = p3
                    ActiveSheet.Cells(startRow + 5, StartCol) = "Checksum"
                    ActiveSheet.Cells(startRow + 5, c) = Check
                    c = c + 1
                End If
                
                
              End If
              p2 = p2 + PercentageIncrement / 100
            Wend
            
          p1 = p1 + PercentageIncrement / 100
      Wend
        
    
    Next

End Sub

Public Function TrapeziumArea(Y1 As Variant, Y2 As Variant, Width As Variant) As Variant
    TrapeziumArea = (Minimum(Y1, Y2) + (Maximum(Y1, Y2) - Minimum(Y1, Y2)) / 2) * Width
End Function

Public Function USERSELECTCELL() As Range
  Set USERSELECTCELL = Application.InputBox(Prompt:="Select first Data Cell", Title:="Cell Select", Type:=8)
  Call Application.GoTo(USERSELECTCELL)
End Function

Public Function SVGICONHEADER(Optional ByVal ScaleFactor As Variant = 1) As String
    'in de basis is een SVG-icon 20 x 20 pixels. Met de ScaleFactor kun je hem optioneel groter maken
    SVGICONHEADER = "<svg width='" & 20 * ScaleFactor & "' height='" & ScaleFactor * 20 & "' xmlns='http://www.w3.org/2000/svg'>"
End Function

Public Function SVGICONFOOTER() As String
    SVGICONFOOTER = "</svg>"
End Function

Public Function SVGTRAPEZIUMSHARP(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGTRAPEZIUMSHARP = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGTRAPEZIUMBLUNT(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGTRAPEZIUMBLUNT = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMBLUNT(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGCULVERT(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGCULVERT = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMBLUNT(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPECIRCLE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGABUTMENTBRIDGE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGABUTMENTBRIDGE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEABUTMENTBRIDGE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGPILLARBRIDGE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGPILLARBRIDGE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEPILLARBRIDGE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGRECTANGULARWEIR(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGRECTANGULARWEIR = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPERECTANGULARWEIR(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGORIFICE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGORIFICE = SVGICONHEADER(ScaleFactor) & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGSHAPEORIFICE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth) & SVGICONFOOTER
End Function

Public Function SVGPUMP(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    SVGSTRING = SVGICONHEADER(ScaleFactor)
    SVGSTRING = SVGSTRING & SVGSHAPECIRCLE(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGSTRING = SVGSTRING & SVGSHAPEPUMP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGPUMP = SVGSTRING & SVGICONFOOTER
End Function

Public Function SVGFISH(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    SVGSTRING = SVGICONHEADER(ScaleFactor)
    SVGSTRING = SVGSTRING & SVGSHAPETRAPEZIUMSHARP(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGSTRING = SVGSTRING & SVGSHAPEFISH(ScaleFactor, FillColor, FillOpacity, StrokeColor, StrokeWidth)
    SVGFISH = SVGSTRING & SVGICONFOOTER
End Function

Public Function GetMeteobaseOrderTypeFromWebRequest(myRequest As String) As String
    'this function returns the type of data our user has requested from Meteobase
    If InStr(1, myRequest, ".zip", vbTextCompare) > 0 Then
        If InStr(1, myRequest, "uploads", vbTextCompare) > 0 Then
            'this is a data upload by a user
            GetMeteobaseOrderTypeFromWebRequest = ""
        ElseIf InStr(1, myRequest, "reeksen", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Toetsingsreeksen"
        ElseIf InStr(1, myRequest, "Toetsingsreeksen", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Toetsingsreeksen"
        ElseIf InStr(1, myRequest, "Stochastentabellen", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Stochasten"
        ElseIf InStr(1, myRequest, "wwwroot", vbTextCompare) > 0 Then
            'this is a webrequest header. Skip it
            GetMeteobaseOrderTypeFromWebRequest = ""
        ElseIf InStr(1, myRequest, "ASCII", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Rasterdata to ASCII"
        ElseIf InStr(1, myRequest, "MODFLOW", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Rasterdata for MODFLOW"
        ElseIf InStr(1, myRequest, "SIMGRO", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Rasterdata for SIMGRO"
        ElseIf InStr(1, myRequest, "HDF5", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Rasterdata to HDF5"
        ElseIf InStr(1, myRequest, "Regios", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Klimaatregio's"
        ElseIf InStr(1, myRequest, "Herhalingstijd", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Herhalingstijd bui"
        ElseIf InStr(1, myRequest, "Neerslagreductie", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Macro neerslagreductie"
        ElseIf InStr(1, myRequest, "Stedelijke", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Stedelijke events"
        Else
            'any other zip-files will 99% of the time consist of rasterdata aggregated by polygon
            GetMeteobaseOrderTypeFromWebRequest = "Rasterdata by Polygon"
        End If
    ElseIf InStr(1, myRequest, ".xls", vbTextCompare) > 0 Then
        If InStr(1, myRequest, "etmaalstation", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Etmaalstations"
        ElseIf InStr(1, myRequest, "uurstation", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Uurstations"
        ElseIf InStr(1, myRequest, "neerslagstatistiek", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Neerslagstatistiek"
        ElseIf InStr(1, myRequest, "2030", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Toetsingsreeksen"
        ElseIf InStr(1, myRequest, "2050", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Toetsingsreeksen"
        ElseIf InStr(1, myRequest, "2085", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Toetsingsreeksen"
        ElseIf InStr(1, myRequest, "regenduurlijn", vbTextCompare) > 0 Then
            GetMeteobaseOrderTypeFromWebRequest = "Regenduurlijnen-app"
        Else
            GetMeteobaseOrderTypeFromWebRequest = "Excel overig"
        End If
    ElseIf InStr(1, myRequest, ".7z", vbTextCompare) > 0 Then
        GetMeteobaseOrderTypeFromWebRequest = "SatDATA 3.0"
    ElseIf InStr(1, myRequest, ".pdf", vbTextCompare) > 0 Then
        GetMeteobaseOrderTypeFromWebRequest = "Documentatie (pdf)"
    ElseIf InStr(1, myRequest, "/wiwb/work_dir", vbTextCompare) > 0 Then
        GetMeteobaseOrderTypeFromWebRequest = "Regenduurlijnen-app"
    Else
        GetMeteobaseOrderTypeFromWebRequest = ""
    End If
    
End Function

Public Function SVGSHAPEPUMP(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim Schoep1 As String, Schoep2 As String, Schoep3 As String
    Dim SVGSTRING As String
    'definitie van een Bézier curve in SVG: C X1 Y1 X2 Y2 x y waarbij: X1,Y1 = handle van het startpunt, X2,Y2 = handle van het eindpunt, x,y = eindpunt
    
    Dim i As Integer
    Dim X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant, Xstart As Variant, Ystart As Variant, Xend As Variant, Yend As Variant
        
    For i = 0 To 5
        Xstart = 10 * ScaleFactor
        Ystart = 5 * ScaleFactor
        Xend = (10 + 3) * ScaleFactor
        Yend = Ystart
        X1 = Xstart + 0.5 * ScaleFactor
        Y1 = Ystart + 0.5 * ScaleFactor
        X2 = Xend - 0.5 * ScaleFactor
        Y2 = Yend + 0.5 * ScaleFactor
        
        'rotate the endpoint and the handles
        Call RotatePoint(Xend, Yend, Xstart, Ystart, 360 / 6 * i, Xend, Yend)
        Call RotatePoint(X1, Y1, Xstart, Ystart, 360 / 5 * i, X1, Y1)
        Call RotatePoint(X2, Y1, Xstart, Ystart, 360 / 5 * i, X2, Y2)
        
        SVGSTRING = SVGSTRING & "<path d='M " & Xstart & " " & Ystart & " C " & X1 & " " & Y1 & " " & X2 & " " & Y2 & " " & Xend & " " & Yend & " ' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "' fill='transparent'/>"
    Next
    
    SVGSHAPEPUMP = SVGSTRING
End Function



Public Function SVGSHAPEFISH(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    Dim SVGSTRING As String
    Dim Xstart As Variant, Xend As Variant, Ystart As Variant, Yend As Variant, X1 As Variant, X2 As Variant, Y1 As Variant, Y2 As Variant
    
    'buik
    Xstart = 5 * ScaleFactor
    Ystart = (5 + 2) * ScaleFactor
    X1 = (5 + 1) * ScaleFactor
    Y1 = 5 * ScaleFactor
    Xend = 15 * ScaleFactor
    Yend = 5 * ScaleFactor
    X2 = (15 - 2) * ScaleFactor
    Y2 = (5 - 4) * ScaleFactor
    SVGSTRING = "<path d='M " & Xstart & " " & Ystart & " C " & X1 & " " & Y1 & " " & X2 & " " & Y2 & " " & Xend & " " & Yend
    
    'rug
    Xstart = 5 * ScaleFactor
    Ystart = (5 - 2) * ScaleFactor
    X1 = (5 + 1) * ScaleFactor
    Y1 = 5 * ScaleFactor
    Xend = 15 * ScaleFactor
    Yend = 5 * ScaleFactor
    X2 = (15 - 2) * ScaleFactor
    Y2 = (5 + 4) * ScaleFactor
    SVGSTRING = SVGSTRING & " C " & X2 & " " & Y2 & " " & X1 & " " & Y1 & " " & Xstart & " " & Ystart & " Z'" & " stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "'/>"
    
    SVGSHAPEFISH = SVGSTRING
End Function


Public Function SVGSHAPEORIFICE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'we schematize an orifice using a shape and a rectangle
    Dim SVGSTRING As String
    SVGSTRING = "<path d='M 0 0 l " & 7 * ScaleFactor & " " & 0 & " l 0 " & 8 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & -8 * ScaleFactor & " l " & 7 * ScaleFactor & " 0 l " & -5 * ScaleFactor & " " & 10 * ScaleFactor & " l " & -10 * ScaleFactor & " 0 " & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
    SVGSHAPEORIFICE = SVGSTRING & "<path d='M " & 7 * ScaleFactor & " 0 l 0 " & 4 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & -4 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPERECTANGULARWEIR(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'we schematize a rectangular weir using a shape and two vertical stripes
    SVGSHAPERECTANGULARWEIR = "<path d='M " & 2 * ScaleFactor & " " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l " & 10 * ScaleFactor & " 0 l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " l " & -5 * ScaleFactor & " 0 l 0 " & 3 * ScaleFactor & " l " & -6 * ScaleFactor & " 0 l 0 " & -3 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 7 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>" & "<path d='M " & 13 * ScaleFactor & " " & 5 * ScaleFactor & " l 0 " & 5 * ScaleFactor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPEPILLARBRIDGE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPEPILLARBRIDGE = "<path d='M 0 0 l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l 0 " & -4 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l 0 " & 8 * ScaleFactor & " l " & 2 * ScaleFactor & " 0 " & " l 0 " & -8 * ScaleFactor & " l " & 6 * ScaleFactor & " 0 l " & " 0 " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPEABUTMENTBRIDGE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPEABUTMENTBRIDGE = "<path d='M 0 0 l " & 3 * ScaleFactor & " " & 6 * ScaleFactor & " l 0 " & -4 * ScaleFactor & " l " & 14 * ScaleFactor & " " & 0 & " l " & " 0 " & 4 * ScaleFactor & " l " & 3 * ScaleFactor & " " & -6 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPECIRCLE(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    SVGSHAPECIRCLE = "<circle cx='" & 10 * ScaleFactor & "' cy='" & 5 * ScaleFactor & "' r='" & 3 * ScaleFactor & "' fill-opacity='0.5' fill='kleur' stroke='kleur' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPETRAPEZIUMSHARP(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPETRAPEZIUMSHARP = "<path d='M 0 0 l " & 5 * ScaleFactor & " " & 10 * ScaleFactor & " l " & 10 * ScaleFactor & " 0 " & " l " & 5 * ScaleFactor & " " & -10 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function SVGSHAPETRAPEZIUMBLUNT(Optional ByVal ScaleFactor As Variant = 1, Optional ByVal FillColor As String = "green", Optional ByVal FillOpacity As Variant = 1, Optional ByVal StrokeColor As String = "black", Optional ByVal StrokeWidth As Integer = 1) As String
    'a typical trapezium would be width 20 and height 10
    SVGSHAPETRAPEZIUMBLUNT = "<path d='M " & 1 * ScaleFactor & " 0 l 0 " & 2 * ScaleFactor & " l " & 4 * ScaleFactor & " " & 8 * ScaleFactor & " l " & 10 * ScaleFactor & " " & 0 & " l " & 4 * ScaleFactor & " " & -8 * ScaleFactor & " l " & 0 & " " & -2 * ScaleFactor & " Z' fill-opacity='" & FillOpacity & "' fill='" & FillColor & "' stroke='" & StrokeColor & "' stroke-width='" & StrokeWidth & "'/>"
End Function

Public Function MAXFROMCOLLECTION(myColl As Collection) As Variant
  Dim VAL As Variant, max As Variant, i As Long
  max = -999999999999#
  For i = 1 To myColl.Count
    VAL = myColl(i)
    If VAL > max Then max = VAL
  Next
  MAXFROMCOLLECTION = max
End Function

Public Function MINFROMCOLLECTION(myColl As Collection) As Variant
  Dim VAL As Variant, Min As Variant, i As Long
  Min = 999999999999#
  For i = 1 To myColl.Count
    VAL = myColl(i)
    If VAL < Min Then Min = VAL
  Next
  MINFROMCOLLECTION = Min
End Function

Public Function AVGFROMCOLLECTION(myColl As Collection) As Variant
  Dim VAL As Variant, sum As Variant, i As Long
  For i = 1 To myColl.Count
    VAL = myColl(i)
    sum = sum + VAL
  Next
  AVGFROMCOLLECTION = sum / myColl.Count
End Function


Public Function GENPARETOPDF(mu As Variant, Sigma As Variant, kappa As Variant, x As Variant) As Variant
  'calculates the cumulative probability density according to the Generalized Pareto probability distribution
  Dim par As Variant
  par = (x - mu) / Sigma

    GENPARETOPDF = 1 / Sigma * (1 + kappa * par) ^ -(1 / kappa + 1)
   
End Function

Public Function GENPARETOCDF(mu As Variant, Sigma As Variant, kappa As Variant, x As Variant) As Variant
  'calculates the cumulative probability density according to the Generalized Pareto probability distribution
  Dim par As Variant
  par = (x - mu) / Sigma

  If kappa = 0 Then
    GENPARETOCDF = 1 - Exp(-par)
  Else
    GENPARETOCDF = 1 - (1 + kappa * par) ^ (-1 / kappa)
  End If
   
End Function

Public Function CONDWEIBULLCDF(alpha As Variant, beta As Variant, gamma As Variant, x As Variant) As Variant
  'calculates the cumulative probability density according to the Conditional Weibull probability distribution
  CONDWEIBULLCDF = 1 - Math.Exp(-((x - gamma) / beta) ^ alpha)
End Function

Function BerekenStochastVolumeKlasse(rKlasseFreq As Range, rStochastInUse As Range, rCurCell As Range) As Variant

Dim rCell As Range
Dim vResult
Dim CurRow As Long, i As Long, n As Long, r As Long, R1 As Long
Dim startRow As Long, endRow As Long
Dim Inuse() As Boolean, Freq() As Variant
Dim Done As Boolean, rad As Integer
Dim rLow As Integer, rHigh As Integer

CurRow = rCurCell.row
n = rKlasseFreq.Count
startRow = rKlasseFreq.row
endRow = startRow + n - 1
ReDim Inuse(startRow To endRow)
ReDim Freq(startRow To endRow)

'Kijk welke stochasten in gebruik zijn
r = startRow - 1
For Each rCell In rStochastInUse
  r = r + 1
  If rCell.value = "a" Then
    Inuse(r) = True
  Else
    Inuse(r) = False
  End If
Next

'inventariseer voor iedere klasse de frequentie
r = startRow - 1
For Each rCell In rKlasseFreq
  r = r + 1
  Freq(r) = rCell.value
Next

'doorloop de range en zoek bij inactieve cellen naar de dichtstbijzijnde actieve broeders
r = startRow - 1
For Each rCell In rKlasseFreq
  r = r + 1
  
  If Inuse(r) = False Then
    rLow = 0
    rHigh = 0
    For R1 = r - 1 To startRow Step -1
      If Inuse(R1) Then
        rLow = R1
        Exit For
      End If
    Next
      
    For R1 = r + 1 To endRow
      If Inuse(R1) Then
        rHigh = R1
        Exit For
      End If
    Next
  
    'herverdeel de frequentie van de inactieve klasse
    If rLow = 0 Then
      Freq(rHigh) = Freq(rHigh) + Freq(r)
      Freq(r) = 0
    ElseIf rHigh = 0 Then
      Freq(rLow) = Freq(rLow) + Freq(r)
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) = Math.Abs(r - rLow) Then
      'divide equally
      Freq(rHigh) = Freq(rHigh) + Freq(r) / 2
      Freq(rLow) = Freq(rLow) + Freq(r) / 2
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) > Math.Abs(r - rLow) Then
      'low is nearest so assign all frequency to that one
      Freq(rLow) = Freq(rLow) + Freq(r)
      Freq(r) = 0
    ElseIf Math.Abs(rHigh - r) < Math.Abs(rLow - r) Then
      Freq(rHigh) = Freq(rHigh) + Freq(r)
      Freq(r) = 0
    End If
  End If
Next

If Freq(CurRow) = 0 Then
  BerekenStochastVolumeKlasse = ""
Else
  BerekenStochastVolumeKlasse = Freq(CurRow)
End If


End Function

Function BerekenStochastPatroonKlasse(rPatroonNaam As Range, rPatroonKans As Range, rStochastInUse As Range, rCurCell As Range) As Variant

Dim Inuse(1 To 7) As Boolean
Dim Kans(1 To 7) As Variant
Dim Naam(1 To 7) As String

Dim StartCol As Integer
Dim c As Integer, C1 As Integer
Dim rCell As Range
Dim pSum As Variant
Dim curCol As Variant, CurIdx As Integer

curCol = rCurCell.column
StartCol = rPatroonNaam.column
CurIdx = curCol - StartCol + 1

'inventariseer de namen
c = 0
For Each rCell In rPatroonNaam
  c = c + 1
  Naam(c) = VBA.UCase(rCell)
Next

'inventariseer de kansen
c = 0
For Each rCell In rPatroonKans
  c = c + 1
  Kans(c) = rCell
Next

'inventariseer het gebruik
c = 0
For Each rCell In rStochastInUse
  c = c + 1
  If rCell = "a" Then
    Inuse(c) = True
  Else
    Inuse(c) = False
  End If
Next

'bereken de som van kansen van de actieve patronen
pSum = 0
For c = 1 To 7
  If Inuse(c) Then pSum = pSum + Kans(c)
Next

'herverdeel de ongebruikte kansen naar rato over alle actieve patronen
For c = 1 To 7
  If Inuse(c) Then
    Kans(c) = Kans(c) / pSum
  Else
    Kans(c) = 0
  End If
Next


If Kans(CurIdx) = 0 Then
  BerekenStochastPatroonKlasse = ""
Else
  BerekenStochastPatroonKlasse = Kans(CurIdx)
End If


End Function

Public Function HERH2KLASSEFREQ(PrevH As Variant, curH As Variant, nextH As Variant, DurationHours As Integer) As Variant
  'computes the frequency of a class given its return period AND the return period of the previous and next class
  'the result is based on the average return period between the surrounding classes
  Dim ExceedanceFrequencyLower As Variant, ExceedanceFrequencyUpper As Variant
  
  If Not IsNumeric(PrevH) Then PrevH = 0
  If Not IsNumeric(nextH) Then nextH = 0
  If Not IsNumeric(curH) Then curH = 0
  
  If curH = 0 Then
    'invalid return period!
    ExceedanceFrequencyLower = 0
    ExceedanceFrequencyUpper = 0
  ElseIf PrevH = 0 Then
    'this is the first class!
    ExceedanceFrequencyLower = 365.25 * 24 / DurationHours
    ExceedanceFrequencyUpper = 1 / ((curH + nextH) / 2)
  ElseIf nextH = 0 Then
    'this is the last class!
    ExceedanceFrequencyLower = 1 / ((PrevH + curH) / 2)
    ExceedanceFrequencyUpper = 0
  Else
    ExceedanceFrequencyLower = 1 / ((PrevH + curH) / 2)
    ExceedanceFrequencyUpper = 1 / ((curH + nextH) / 2)
  End If
  HERH2KLASSEFREQ = ExceedanceFrequencyLower - ExceedanceFrequencyUpper
  
End Function

Public Function HERHFROMSTOCHASTICRESULT(HERH As Variant, WLEventNumRange As Range, WLValueRange As Range, FreqEventNumRange As Range, FreqValueRange As Range) As Variant
  'this function computes the exceedance level for a given return period.
  'it expects two ranges with resp. event numbers and corresponding water levels,
  'and two ranges with resp. event numbers and corresponding frequencies
  Dim rWL As Long, rFreq As Long
  
  Dim WLValues() As Variant, WLEventNums() As Integer, WLSortedIdx() As Long
  Dim FreqValues() As Variant, FreqEventNums() As Integer
  Dim WLSorted() As Variant, Herhalingstijd() As Variant
  Dim FreqSum As Variant, i As Long, j As Long
  Dim myEventNum As Integer, myWL As Variant, myFreq As Variant
  
  'input
  ReDim WLValues(1 To WLValueRange.Rows.Count)
  ReDim WLEventNums(1 To WLEventNumRange.Rows.Count)
  ReDim FreqValues(1 To FreqValueRange.Rows.Count)
  ReDim FreqEventNums(1 To FreqEventNumRange.Rows.Count)
  
  'output
  ReDim WLSorted(1 To WLValueRange.Rows.Count)
  ReDim Herhalingstijd(1 To WLValueRange.Rows.Count)
  
  If WLValueRange.Rows.Count <> WLEventNumRange.Rows.Count Then
    MsgBox ("Error: number of rows in water level range must be equal to that in the event number range.")
  ElseIf FreqValueRange.Rows.Count <> FreqEventNumRange.Rows.Count Then
    MsgBox ("Error: number of rows in frequency value range must be equal to that in the event number range.")
  Else
  
    'read the water levels
    For rWL = 1 To WLEventNumRange.Rows.Count
      WLValues(rWL) = WLValueRange.Cells(rWL, 1)
      WLEventNums(rWL) = WLEventNumRange.Cells(rWL, 1)
    Next
    
    'read the frequencies
    For rFreq = 1 To FreqEventNumRange.Rows.Count
      FreqValues(rFreq) = FreqValueRange.Cells(rFreq, 1)
      FreqEventNums(rFreq) = FreqEventNumRange.Cells(rFreq, 1)
    Next
    
    'create an array with the index number for the water levels in ascending order
    WLSortedIdx = HeapSort(WLValues)
    
    'walk through the water levels in descending order
    For i = UBound(WLSortedIdx) To 1 Step -1
      myEventNum = WLEventNums(WLSortedIdx(i))
      myWL = WLValues(WLSortedIdx(i))
      
      'find the frequency corresponding with this event
      For j = 1 To UBound(FreqEventNums)
        If FreqEventNums(j) = myEventNum Then
          myFreq = FreqValues(j)
          Exit For
        End If
      Next
      
      FreqSum = FreqSum + myFreq
      WLSorted(i) = myWL
      Herhalingstijd(i) = 1 / FreqSum
    Next
    
  End If
  
  'interpolate between the two surrounding Return Periods.
  For i = 1 To UBound(WLSorted) - 1
    If Herhalingstijd(i) <= HERH And Herhalingstijd(i + 1) >= HERH Then
      HERHFROMSTOCHASTICRESULT = Interpolate(Herhalingstijd(i), WLSorted(i), Herhalingstijd(i + 1), WLSorted(i + 1), HERH)
      Exit Function
    End If
  Next
  
End Function

Public Function KLASSEFREQUENTIEUITHERHALINGSTIJD(FrequentieSom As Variant, HerhOndergrens As Variant, HerhBovengrens As Variant, VolgendeHerh As Variant) As Variant
  If Not IsNumeric(HerhOndergrens) Or HerhOndergrens = "" Then
    'bereken klassefrequentie voor de onderste klasse
    KLASSEFREQUENTIEUITHERHALINGSTIJD = FrequentieSom - 1 / HerhBovengrens
  ElseIf Not IsNumeric(VolgendeHerh) Or VolgendeHerh = "" Then
    'er is geen volgende klasse, dus bereken hier het restant van de frequenties
    KLASSEFREQUENTIEUITHERHALINGSTIJD = 1 / HerhOndergrens
  Else
    KLASSEFREQUENTIEUITHERHALINGSTIJD = (1 / HerhOndergrens) - (1 / HerhBovengrens)
  End If
End Function

Public Function KLASSEKANSUITOVERSCHRIJDINGSKANSEN(Vorige As Variant, Huidige As Variant, Volgende As Variant) As Variant
  If Not IsNumeric(Vorige) Or Vorige = "" Then
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = 1 - Huidige
  ElseIf Not IsNumeric(Volgende) Or Volgende = "" Then
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = Vorige
  Else
    KLASSEKANSUITOVERSCHRIJDINGSKANSEN = Vorige - Huidige
  End If
End Function

Public Sub CLASSIFYDURATIONS(ValuesRange As Range, threshold As Variant, ResultsRow As Integer, ResultsCol As Integer)
  'deze routine onderzoekt welke duur (aantal tijdstappen) gebeurtenissen in een reeks hebben
  'argumenten: het bereik waarin de getallen staan en de drempelwaarde waarboven een gebeurtenis wordt 'gedetecteerd'
  
  Dim i As Long, j As Long, values() As Variant, Inuse() As Boolean, Durations() As Integer
  Dim n As Integer, maxn As Integer
  ReDim values(1 To ValuesRange.Rows.Count)
  ReDim Inuse(1 To ValuesRange.Rows.Count)
  ReDim Durations(1 To ValuesRange.Rows.Count)
  For i = 1 To ValuesRange.Rows.Count
    values(i) = ValuesRange.Cells(i, 1)
    Inuse(i) = False
  Next
  
  Dim Index() As Long
  Index = HeapSort(values)
  
  For i = UBound(Index) To 1 Step -1
    If values(Index(i)) > threshold And Inuse(Index(i)) = False Then
      n = 1
      Inuse(Index(i)) = True
      'move backwards to find the start of the event
      For j = Index(i) - 1 To 1 Step -1
        If Inuse(j) = True Then Exit For
        If values(j) <= threshold Then Exit For
        n = n + 1
        Inuse(j) = True
      Next
      'move forwards to find the end of the event
      For j = Index(i) + 1 To ValuesRange.Rows.Count
        If Inuse(j) = True Then Exit For
        If values(j) <= threshold Then Exit For
        n = n + 1
        Inuse(j) = True
      Next
      
      'keep track of the longest event found
      If n > maxn Then maxn = n
      
      'we've found an event and identified its duration. Store it in a histogram
      Durations(n) = Durations(n) + 1
    
    End If
  Next
  
  ReDim Preserve Durations(1 To maxn)
  
  'write the histogram to the results sheet
  For i = 1 To UBound(Durations)
    ActiveSheet.Cells(ResultsRow + i - 1, ResultsCol) = i
    ActiveSheet.Cells(ResultsRow + i - 1, ResultsCol + 1) = Durations(i)
  Next
    
  
End Sub

Private Sub Heapify(Keys, Index() As Long, ByVal i1 As Long, ByVal n As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   Dim Base As Long: Base = LBound(Index)
   Dim nDiv2 As Long: nDiv2 = n \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim k As Long: k = 2 * i + 1
      If k + 1 < n Then
         If Keys(Index(Base + k)) < Keys(Index(Base + k + 1)) Then k = k + 1
         End If
      If Keys(Index(Base + i)) >= Keys(Index(Base + k)) Then Exit Do
      Exchange Index, i, k
      i = k
      Loop
   End Sub

Private Sub Exchange(a() As Long, ByVal i As Long, ByVal j As Long)
   Dim Base As Long: Base = LBound(a)
   Dim Temp As Long: Temp = a(Base + i)
   a(Base + i) = a(Base + j)
   a(Base + j) = Temp
   End Sub

Private Function GenerateArrayWithRandomValues()
   Dim n As Long: n = 1 + Rnd * 100
   ReDim a(0 To n - 1) As Long
   Dim i As Long
   For i = LBound(a) To UBound(a)
      a(i) = Rnd * 1000
      Next
   GenerateArrayWithRandomValues = a
   End Function

Private Sub VerifyIndexIsSorted(Keys, Index)
   Dim i As Long
   For i = LBound(Index) To UBound(Index) - 1
      If Keys(Index(i)) > Keys(Index(i + 1)) Then
         err.Raise vbObjectError, , "Index array is not sorted!"
         End If
      Next
   End Sub


Public Function OPPERVLAKAFGEPLATTECIRKEL(r As Variant, Y_center As Variant, Y_snede As Variant) As Variant
  'R = straal, Y_center = hoogte VBA.Middelpunt cirkel, Y_snede = hoogte waar de cirkel is afgesneden
  Dim O_cirkel As Variant, O_taartpunt As Variant, O_driehoek As Variant
  Dim Hoogte As Variant, Breedte As Variant, Hoek As Variant, pi As Variant
  
  pi = 3.141592
  Hoogte = Y_snede - Y_center
  O_cirkel = pi * r ^ 2
  
  If Hoogte >= r Then
    'volledig gevulde cirkel
    OPPERVLAKAFGEPLATTECIRKEL = O_cirkel
  ElseIf Hoogte <= -1 * r Then 'lege cirkel
    OPPERVLAKAFGEPLATTECIRKEL = 0
  Else
    'de taartpunt die eruit wordt geknipt
    Breedte = Sqr(r ^ 2 - Hoogte ^ 2) 'pythagoras
    Hoek = 2 * ArcCos(Hoogte / r)
    O_taartpunt = Hoek / (2 * pi) * O_cirkel
    
    'de driehoek die weer moet worden toegevoegd
    O_driehoek = 2 * Hoogte * Breedte / 2
    
    OPPERVLAKAFGEPLATTECIRKEL = O_cirkel - O_taartpunt + O_driehoek
  End If
  
End Function

Public Function RotatePoint(ByVal Xold As Variant, ByVal Yold As Variant, ByVal XOrigin As Variant, ByVal YOrigin As Variant, ByVal degrees As Variant, ByRef Xnew As Variant, ByRef Ynew As Variant) As Boolean
 Dim r As Variant, theta As Variant, dY As Variant, dX As Variant, Direction As Variant
 'roteert een punt ten opzichte van zijn oorsprong
  
 dY = (Yold - YOrigin)
 dX = (Xold - XOrigin)
 r = Sqr(dX ^ 2 + dY ^ 2)
 
 If dX = 0 Then dX = 0.00000000000001
 theta = Math.Atn(dY / dX)
    
 Xnew = r * Math.Cos(theta - DEG2RAD(degrees)) + XOrigin
 Ynew = r * Math.Sin(theta - DEG2RAD(degrees)) + YOrigin
 RotatePoint = True
End Function
  
Public Function DEG2RAD(ByVal Angle As Variant) As Variant
  'graden naar radialen
  DEG2RAD = Angle / 180 * pi
End Function

Public Function RAD2DEG(ByVal Angle As Variant) As Variant
  'radialen naar graden
  RAD2DEG = Angle * 180 / pi
End Function

Public Function LINEANGLEDEGREES(ByVal X1 As Variant, ByVal Y1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant) As Variant
  'berekent de hoek van een lijn tussen twee xy co-ordinaten
  Dim dX As Variant, dY As Variant
  
  dX = VBA.Abs(X2 - X1)
  dY = VBA.Abs(Y2 - Y1)
  
  If dX = 0 Then
    If dY = 0 Then
      LINEANGLEDEGREES = 0
    ElseIf Y2 > Y1 Then
      LINEANGLEDEGREES = 0
    ElseIf Y2 < Y1 Then
      LINEANGLEDEGREES = 180
    End If
  ElseIf dY = 0 Then
    If X2 > X1 Then
      LINEANGLEDEGREES = 90
    ElseIf X2 < X1 Then
      LINEANGLEDEGREES = 270
    End If
  Else
    If X2 > X1 And Y2 > Y1 Then 'eerste kwadrant
      LINEANGLEDEGREES = R2D(VBA.Atn(dX / dY))
    ElseIf X2 > X1 And Y2 < Y1 Then 'tweede kwadrant
      LINEANGLEDEGREES = 90 + R2D(VBA.Atn(dY / dX))
    ElseIf X2 < X1 And Y2 < Y1 Then 'derde kwadrant
      LINEANGLEDEGREES = 180 + R2D(VBA.Atn(dX / dY))
    Else 'vierde kwadrant
      LINEANGLEDEGREES = 270 + R2D(VBA.Atn(dX / dY))
    End If
  End If
  
End Function

Public Function PointDistance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Variant
  PointDistance = VBA.Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Public Function PointInPolygon(ByVal x As Variant, ByVal y As Variant, VerticesX As Collection, VerticesY As Collection) As Boolean
Dim pt As Integer
Dim total_angle As Variant

  'Add up the angles between the point in question and adjacent points on the polygon taken in order.
  'If the total of all the angles is 2 * PI or -2 * PI, then the point is inside the polygon.
  'If the total is zero, the point is outside. You can verify this intuitively with some simple examples using squares or triangles.

    ' Get the angle between the point and the
    ' first and last vertices.
    total_angle = GetAngle(VerticesX(VerticesX.Count), VerticesY(VerticesY.Count), x, y, VerticesX(1), VerticesY(1))
    
    ' Add the angles from the point to each other pair of vertices.
    For pt = 1 To VerticesX.Count - 1
      total_angle = total_angle + GetAngle(VerticesX(pt), VerticesY(pt), x, y, VerticesX(pt + 1), VerticesY(pt + 1))
    Next pt

    ' The total angle should be 2 * PI or -2 * PI if
    ' the point is in the polygon and close to zero
    ' if the point is outside the polygon.
    PointInPolygon = (Abs(total_angle) > pi)
End Function

Public Function NearestPoint(ByVal x As Variant, ByVal y As Variant, ByVal MyRange As Range, ByVal XCol As Integer, ByVal YCol As Integer, ByVal ReturnCol As Integer, HasHeader As Boolean)

  Dim r As Long, minDist As Variant, myDist As Variant
  Dim startRow As Integer
  Dim myX As Variant, myY As Variant, minID As String
  minDist = 99999999
  
  If HasHeader Then
    startRow = 2
  Else
    startRow = 1
  End If
  
  For r = startRow To MyRange.Rows.Count
    myX = MyRange.Cells(r, XCol)
    myY = MyRange.Cells(r, YCol)
    myDist = Math.Sqr((myX - x) ^ 2 + (myY - y) ^ 2)
    If myDist < minDist Then
      minDist = myDist
      minID = MyRange.Cells(r, ReturnCol)
    End If
  Next

  NearestPoint = minID

End Function

Public Function PoolCoordinaatX(ByVal alpha As Variant, Length As Variant) As Variant
  'geeft de x-coordinaat terug, gegeven poolcoordinaat (alpha, lengte).
  'Let op: de hoek alhpa is gedefinieerd vanaf de vertikale as, NIET vanaf de horizontale!
  Dim rad As Variant
  rad = D2R(alpha)
  PoolCoordinaatX = Sin(rad) * Length
End Function

Public Function PoolCoordinaatY(ByVal alpha As Variant, Length As Variant) As Variant
  'geeft de y-coordinaat terug, gegeven poolcoordinaat (alpha, lengte).
  'Let op: de hoek alhpa is gedefinieerd vanaf de vertikale as, NIET vanaf de horizontale!
  Dim rad As Variant
  rad = D2R(alpha)
  PoolCoordinaatY = Cos(rad) * Length
End Function

Public Function PYTHAGORAS(ByVal a As Variant, b As Variant) As Variant
  PYTHAGORAS = Math.Sqr(a ^ 2 + b ^ 2)
End Function

Public Function PYTHAGORAS_INVERSE(ByVal a As Variant, c As Variant) As Variant
  'c = schuine zijde, a = rechte zijde
  'a^2 + b^2 = c ^2
  'b^2 = c^2 - a ^2
  'b = sqr(c^2 - a^2)
  PYTHAGORAS_INVERSE = Math.Sqr(c ^ 2 - a ^ 2)
End Function

Public Function Afstand(Naar As String) As Double
        If Naar = "HDSR" Then
            Afstand = 80.7
        ElseIf Naar = "WSHD" Then
            Afstand = 41
        ElseIf Naar = "ValleiEnVeluwe" Then
            Afstand = 153
        ElseIf Naar = "Wetterskip" Then
            Afstand = 193
        ElseIf Naar = "Noorderzijlvest" Then
            Afstand = 243
        ElseIf Naar = "HunzeEnAas" Then
            Afstand = 261
        ElseIf Naar = "Scheldestromen" Then
            Afstand = 158
        ElseIf Naar = "Rijnland" Then
            Afstand = 20.1
        ElseIf Naar = "Delfland" Then
            Afstand = 11.2
        ElseIf Naar = "Rivierenland" Then
            Afstand = 106
        ElseIf Naar = "Larenstein" Then
            Afstand = 133
        ElseIf Naar = "ViForis" Then
            Afstand = 200
        ElseIf Naar = "RijnEnIJssel" Then
            Afstand = 161
        ElseIf Naar = "HWH" Then
            Afstand = 95.9
        ElseIf Naar = "STOWA" Then
            Afstand = 96.9
        ElseIf Naar = "LWRO" Then
            Afstand = 108
        ElseIf Naar = "Waterschap Rijn en IJssel" Then
            Afstand = 155
        ElseIf Naar = "WUR" Then
            Afstand = 107
        ElseIf Naar = "Viforis" Then
            Afstand = 190
        ElseIf Naar = "WL" Then
            Afstand = 190
        ElseIf Naar = "HWH" Then
            Afstand = 90.3
        ElseIf Naar = "STOWA" Then
            Afstand = 90.3
        End If
    
End Function


Public Function MileageOneUp(startNum As Integer, endNum As Integer, ByRef myArray() As Integer) As Boolean
  'werkt als een kilometerteller. Als het hectometergetal boven z'n maximum komt, springt hij terug naar nul
  'en gaat het getalletje ervoor een omhoog et cetera. Produceert TRUE bij succes
  'produceert FALSE als hij aan z'n eind is gekomen en niet verder kan ophogen
  Dim nElements As Integer
  nElements = UBound(myArray)
  Dim i As Long, j As Long
  Dim Done As Boolean
  Done = True
  
  '---------------------------------------------------------------------
  'for the very first run, the following is crucial:
  'if necessary initialize the array to be equal to its start number
  'EXCEPT for the last number. That should start one lower than startnum
  '---------------------------------------------------------------------
  For i = 1 To nElements - 1
    If myArray(i) < startNum Then myArray(i) = startNum
  Next
  If myArray(nElements) < startNum - 1 Then myArray(nElements) = startNum - 1
  '---------------------------------------------------------------------

  'also check if the counter is already at its maxumum. if so, quit
  For i = 1 To nElements
    If myArray(i) < endNum Then Done = False
    Exit For
  Next
  If Done Then
    MileageOneUp = False
    Exit Function
  End If

  'there is still some room for moving further
  For i = nElements To 1 Step -1
    If myArray(i) < endNum Then
      myArray(i) = myArray(i) + 1
      MileageOneUp = True
      Exit Function
    ElseIf myArray(i) = endNum And i = 1 Then
      MileageOneUp = False
      Exit Function
    ElseIf myArray(i) = endNum And i > 1 Then
      myArray(i) = startNum 'reset de waarde naar de basisstand en draai de voorgaande een omhoog
      For j = i - 1 To 1 Step -1
        If myArray(j) < endNum Then
          myArray(j) = myArray(j) + 1
          MileageOneUp = True
          Exit Function
        Else
          myArray(j) = startNum 'reset ook deze voorgaande waarde naar de basisstand en ga weer door naar degene hiervoor
        End If
      Next
    End If
  Next

End Function

Public Function MeetsCondition(ByVal myVal As Variant, ByVal Condition As String) As Boolean
  Dim Operator As String, Operand As Variant
  
  'tests a value to a certain conditions
  Condition = VBA.Trim(Condition)
  
  'if no condition specified, exit straight away. Always true
  If Condition = "" Then
    MeetsCondition = True
    Exit Function
  End If
  
  'check validity of the condition string
  If InStr(1, Condition, " ") <= 0 Then
    MsgBox ("Error: condition must contain a space between operator and operand: " & Condition)
    End
  End If
  
  'parse the string to retrieve operator and operand
  Operator = ParseString(Condition)
  Operand = Condition
  
  'perform the check
  Select Case Operator
    Case Is = "<"
      If myVal < Operand Then MeetsCondition = True
    Case Is = "<="
      If myVal <= Operand Then MeetsCondition = True
    Case Is = ">"
      If myVal > Operand Then MeetsCondition = True
    Case Is = ">="
      If myVal >= Operand Then MeetsCondition = True
    Case Is = "<>"
      If myVal <> Operand Then MeetsCondition = True
    Case Is = "="
      If myVal = Operand Then MeetsCondition = True
    Case Else
      MsgBox ("Error: operand not (yet) supported in condition: " & Operand & " " & Operator)
      End
  End Select

End Function

Public Function RMSE(Range1 As Range, Range2 As Range) As Variant
    Dim i As Long, myRMSE As Variant
    If Range1.Rows.Count <> Range2.Rows.Count Then
        MsgBox ("Ranges should have equal length")
    Else
        For i = 1 To Range1.Rows.Count
            myRMSE = myRMSE + (Range1.Cells(i, 0) - Range2.Cells(i, 0)) ^ 2
        Next
        RMSE = Math.Sqr(myRMSE / Range1.Rows.Count)
    End If
End Function

Public Function GetShapeByNameFromWorksheet(ByRef MySheet As Worksheet, MyName As String) As Shape
  'finds the shape with a given name on the active worksheet
  Dim myShape As Shape
  For Each myShape In MySheet.Shapes
    If myShape.Name = MyName Then
      Set GetShapeByNameFromWorksheet = myShape
    End If
  Next
End Function


Public Function TrapeziumRangeToYZProfiles(Source As Range, IDCol As Integer, XCol As Integer, YCol As Integer, ProfileHeight As Variant, BedLevelCol As Integer, BedWidthCol As Integer, LeftSlopeCol As Integer, RightSlopeCol As Integer, WetBermLeftBob1Col As Integer, WetBermLeftBob2Col As Integer, WetBermLeftWidthCol As Integer, WetBermLeftSlopeCol As Integer, WetBermRightBob1Col As Integer, WetBermRightBob2Col As Integer, WetBermRightWidthCol As Integer, WetBermRightSlopeCol As Integer, TargetSheet As String) As Boolean
    'this function creates YZ-cross section tables from a given range with trapezium information
    'in this function we assume that nothing is known of the surface level or surface width
    Dim rs As Integer, cs As Integer 'row and column source
    Dim rt As Integer, ct As Integer 'row and column target
    Dim i As Integer, nErrors As Integer
    Dim BedLevel As Variant, BedWidth As Variant
    Dim SlopeLeft As Variant, SlopeRight As Variant
    Dim SurfaceLevel As Variant
    Dim WetBermLeftBob1 As Variant, WetBermLeftBob2 As Variant, WetBermLeftWidth As Variant
    Dim WetBermRightBob1 As Variant, WetBermRightBob2 As Variant, WetBermRightWidth As Variant
    Dim WetBermLeftSlope As Variant, WetBermRightSlope As Variant
    Dim YVals As Collection
    Dim ZVals As Collection
    Dim XCOORD As Variant, YCOORD As Variant, id As String
    
    'initialize the target sheet
    rt = rt + 1
    Worksheets(TargetSheet).Cells(rt, 1) = "ID"
    Worksheets(TargetSheet).Cells(rt, 2) = "XCOORD"
    Worksheets(TargetSheet).Cells(rt, 3) = "YCOORD"
    Worksheets(TargetSheet).Cells(rt, 4) = "Y"
    Worksheets(TargetSheet).Cells(rt, 5) = "X"
           
    'process each cross section
    For rs = 2 To Source.Rows.Count 'skip the header
        nErrors = 0
        Set YVals = New Collection
        Set ZVals = New Collection
        If Source.Cells(rs, IDCol) = "" Then nErrors = nErrors + 1 Else id = Source.Cells(rs, IDCol)
        If Source.Cells(rs, XCol) = "" Then nErrors = nErrors + 1 Else XCOORD = Source.Cells(rs, XCol)
        If Source.Cells(rs, YCol) = "" Then nErrors = nErrors + 1 Else YCOORD = Source.Cells(rs, YCol)
        If Source.Cells(rs, BedLevelCol) = "" Then nErrors = nErrors + 1 Else BedLevel = Source.Cells(rs, BedLevelCol)
        If Source.Cells(rs, BedWidthCol) = "" Then nErrors = nErrors + 1 Else BedWidth = Source.Cells(rs, BedWidthCol)
                
        If nErrors = 0 Then
            If Source.Cells(rs, LeftSlopeCol) <> "" Then SlopeLeft = Source.Cells(rs, LeftSlopeCol) Else SlopeLeft = Source.Cells(rs, RightSlopeCol)
            If Source.Cells(rs, RightSlopeCol) <> "" Then SlopeRight = Source.Cells(rs, RightSlopeCol) Else SlopeRight = SlopeLeft
            
            'add the leftmost coordinate
            SurfaceLevel = BedLevel + ProfileHeight
            Call YVals.Add(0)
            Call ZVals.Add(SurfaceLevel)
            
            'in case we have a left wet berm, add it
            If Source.Cells(rs, WetBermLeftBob1Col) <> "" And Source.Cells(rs, WetBermLeftBob2Col) <> "" And Source.Cells(rs, WetBermLeftWidthCol) > 0 Then
                'calculate the distance of the start of the wet berm
                WetBermLeftBob1 = Source.Cells(rs, WetBermLeftBob1Col)
                WetBermLeftBob2 = Source.Cells(rs, WetBermLeftBob2Col)
                WetBermLeftWidth = Source.Cells(rs, WetBermLeftWidthCol)
                If Source.Cells(rs, WetBermLeftSlopeCol) <> "" Then WetBermLeftSlope = Source.Cells(rs, WetBermLeftSlopeCol) Else WetBermLeftSlope = SlopeLeft
                            
                Call YVals.Add(WetBermLeftSlope * (SurfaceLevel - WetBermLeftBob2))
                Call ZVals.Add(WetBermLeftBob2)
                Call YVals.Add(YVals.Item(YVals.Count) + WetBermLeftWidth)
                Call ZVals.Add(WetBermLeftBob1)
                Call YVals.Add(YVals.Item(YVals.Count) + SlopeLeft * (WetBermLeftBob1 - BedLevel))
                Call ZVals.Add(BedLevel)
            Else
                Call YVals.Add(SlopeLeft * (SurfaceLevel - BedLevel))
                Call ZVals.Add(BedLevel)
            End If
            
            'bed width
            Call YVals.Add(YVals.Item(YVals.Count) + BedWidth)
            Call ZVals.Add(BedLevel)
            
            'in case we have a right wet berm, add it
            If Source.Cells(rs, WetBermRightBob1Col) <> "" And Source.Cells(rs, WetBermRightBob2Col) <> "" And Source.Cells(rs, WetBermRightWidthCol) > 0 Then
                'calculate the distance of the start of the wet berm
                WetBermRightBob1 = Source.Cells(rs, WetBermRightBob1Col)
                WetBermRightBob2 = Source.Cells(rs, WetBermRightBob2Col)
                WetBermRightWidth = Source.Cells(rs, WetBermRightWidthCol)
                If Source.Cells(rs, WetBermRightSlopeCol) <> "" Then WetBermRightSlope = Source.Cells(rs, WetBermRightSlopeCol) Else WetBermRightSlope = SlopeRight
                
                Call YVals.Add(YVals.Item(YVals.Count) + (WetBermRightBob1 - BedLevel) * SlopeRight)
                Call ZVals.Add(WetBermRightBob1)
                Call YVals.Add(YVals.Item(YVals.Count) + WetBermRightWidth)
                Call ZVals.Add(WetBermRightBob2)
                Call YVals.Add(YVals.Item(YVals.Count) + (SurfaceLevel - WetBermRightBob2) * WetBermRightSlope)
                Call ZVals.Add(SurfaceLevel)
            Else
                Call YVals.Add(YVals.Item(YVals.Count) + SlopeRight * (SurfaceLevel - BedLevel))
                Call ZVals.Add(SurfaceLevel)
            End If
            
            For i = 1 To YVals.Count
                rt = rt + 1
                Worksheets(TargetSheet).Cells(rt, 1) = id
                Worksheets(TargetSheet).Cells(rt, 2) = XCOORD
                Worksheets(TargetSheet).Cells(rt, 3) = YCOORD
                Worksheets(TargetSheet).Cells(rt, 4) = YVals(i)
                Worksheets(TargetSheet).Cells(rt, 5) = ZVals(i)
            Next
        End If
    Next


End Function


Public Function GetAngle(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single) As Single
' Return the angle ABC.
' Returns a value between PI and -PI.
' Note that the value is the opposite of what you might expect because Y coordinates increase downward.
    Dim dot_product As Single
    Dim cross_product As Single

    ' Get the dot product and cross product.
    dot_product = DotProduct(Ax, Ay, Bx, By, Cx, Cy)
    cross_product = CrossProductLength(Ax, Ay, Bx, By, Cx, Cy)

    ' Calculate the angle.
    GetAngle = ATan2(cross_product, dot_product)
End Function

Public Function ATan2(ByVal Opp As Single, ByVal adj As Single) As Single
  Dim Angle As Single
  ' Return the angle with tangent opp/hyp. The returned
  ' value is between PI and -PI.

  ' Get the basic angle.
  If Abs(adj) < 0.0001 Then
    Angle = pi / 2
  Else
    Angle = Abs(Atn(Opp / adj))
  End If

  ' See if we are in quadrant 2 or 3.
  If adj < 0 Then
    'angle > PI/2 or angle < -PI/2.
    Angle = pi - Angle
  End If

  'See if we are in quadrant 3 or 4.
  If Opp < 0 Then
    Angle = -Angle
  End If

  'Return the result.
  ATan2 = Angle

End Function


Private Function DotProduct(ByVal Ax As Single, ByVal Ay As Single, ByVal Bx As Single, ByVal By As Single, ByVal Cx As Single, ByVal Cy As Single) As Single
  ' Return the dot product AB · BC.
  ' Note that AB · BC = |AB| * |BC| * Cos(theta).
  Dim BAx As Single
  Dim BAy As Single
  Dim BCx As Single
  Dim BCy As Single
    
  ' Get the vectors' coordinates.
  BAx = Ax - Bx
  BAy = Ay - By
  BCx = Cx - Bx
  BCy = Cy - By
    
  ' Calculate the dot product.
  DotProduct = BAx * BCx + BAy * BCy

End Function

Public Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal By As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single

  ' Return the cross product AB x BC.
  ' The cross product is a vector perpendicular to AB
  ' and BC having length |AB| * |BC| * Sin(theta) and
  ' with direction given by the VBA.Right-hand rule.
  ' For two vectors in the X-Y plane, the result is a
  ' vector with X and Y components 0 so the Z component
  ' gives the vector's length and direction.

  Dim BAx As Single
  Dim BAy As Single
  Dim BCx As Single
  Dim BCy As Single

  ' Get the vectors' coordinates.
  BAx = Ax - Bx
  BAy = Ay - By
  BCx = Cx - Bx
  BCy = Cy - By

  ' Calculate the Z coordinate of the cross product.
  CrossProductLength = BAx * BCy - BAy * BCx

End Function

Public Function NATTEOMTREKAFGEPLATTECIRKEL(r As Variant, Y_center As Variant, Y_snede As Variant) As Variant
  Dim Hoogte As Variant, Breedte As Variant, Hoek As Variant
  Dim Omtrek_cirkel As Variant
  
  Omtrek_cirkel = 2 * pi * r
  Hoogte = Y_snede - Y_center
  
  If Hoogte >= r Then        'volledige cirkel
    NATTEOMTREKAFGEPLATTECIRKEL = 2 * pi * r
  ElseIf Hoogte <= -1 * r Then   'lege cirkel
    NATTEOMTREKAFGEPLATTECIRKEL = 0
  Else                                  'de hoek van de taartpunt die eruit wordt geknipt (radialen)
    Breedte = Sqr(r ^ 2 - Hoogte ^ 2) 'pythagoras
    Hoek = 2 * ArcCos(Hoogte / r)
    NATTEOMTREKAFGEPLATTECIRKEL = (2 * pi - Hoek) * r
  End If
  
End Function

Public Function EllipsBreedte(Breedte As Variant, Hoogte As Variant, H As Variant) As Variant
  'h is gedefinieerd als de hoogte vanaf de bodem van de ellips
  'een ellips voldoet aan de vgl x^2/a^2 + y^2/b^2 = 1
  'waarbij het brandpunt van de ellips als nulpunt moet worden beschouwd, a de halve breedte is en b de halve hoogte
  Dim a As Variant
  Dim b As Variant
  Dim y As Variant 'hoogte y tov brandpunt
  Dim x As Variant
  
  b = Hoogte / 2
  a = Breedte / 2
  
  y = H - b
  
  If H >= 0 And H <= Hoogte Then
    x = Sqr((1 - y ^ 2 / b ^ 2) * a ^ 2)
    EllipsBreedte = x * 2
  Else
    EllipsBreedte = -999
  End If

End Function

' Inverse Sinus
Function ArcSin(x As Variant) As Variant
  ArcSin = Atn(x / Sqr(-x * x + 1))
End Function

'Inverse Cosinus
Function ArcCos(x As Variant) As Variant
  ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

'inverse tangent
Function ArcTan(x As Variant) As Variant
  ArcTan = Atn(x)
End Function

Public Function ArcTan2(ByVal x As Variant, ByVal y As Variant) As Variant
  
  'Code from www.visiblevisual.com
  If x = 0 And y = 0 Then
    ATan2 = 0
  Else
    If x = 0 Then x = 0.00000000001
    ATan2 = Atn(y / x) - pi * (x < 0)
  End If
  End Function

End Function

Public Function DaysInMonth(myDate)

  Dim NextMonth, EndOfMonth
  NextMonth = DateAdd("m", 1, myDate)
  EndOfMonth = NextMonth - DatePart("d", NextMonth)
  DaysInMonth = DatePart("d", EndOfMonth)

End Function

Public Function DaysInMonth2(myMonth As Integer, myYear As Integer, Optional AlwaysInclude29Feb As Boolean = False)

  If myMonth = 1 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 2 Then
    If AlwaysInclude29Feb Then
      DaysInMonth2 = 29
    ElseIf IsLeapYear(myYear) Then
      DaysInMonth2 = 29
    Else
      DaysInMonth2 = 28
    End If
  ElseIf myMonth = 3 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 4 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 5 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 6 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 7 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 8 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 9 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 10 Then
    DaysInMonth2 = 31
  ElseIf myMonth = 11 Then
    DaysInMonth2 = 30
  ElseIf myMonth = 12 Then
    DaysInMonth2 = 31
  End If
End Function

Public Function IsLeapYear(myYear As Integer) As Boolean
  If VBA.Round(myYear / 4, 0) = myYear / 4 Then
    IsLeapYear = True
  Else
    IsLeapYear = False
  End If
End Function

Public Function Kwartaal(myDate)
  Select Case month(myDate)
  Case Is = 1
    Kwartaal = 1
  Case Is = 2
    Kwartaal = 1
  Case Is = 3
    Kwartaal = 1
  Case Is = 4
    Kwartaal = 2
  Case Is = 5
    Kwartaal = 2
  Case Is = 6
    Kwartaal = 2
  Case Is = 7
    Kwartaal = 3
  Case Is = 8
    Kwartaal = 3
  Case Is = 9
    Kwartaal = 3
  Case Is = 10
    Kwartaal = 4
  Case Is = 11
    Kwartaal = 4
  Case Is = 12
    Kwartaal = 4
  End Select
End Function

Public Function Halfjaar(myDate As Date) As String
  Select Case month(myDate)
    
  Case Is = 1
    Halfjaar = year(myDate) - 1 & "-" & VBA.Right(year(myDate), 2) & " winter"
  Case Is = 2
    Halfjaar = year(myDate) - 1 & "-" & VBA.Right(year(myDate), 2) & " winter"
  Case Is = 3
    Halfjaar = year(myDate) - 1 & "-" & VBA.Right(year(myDate), 2) & " winter"
  Case Is = 4
    Halfjaar = year(myDate) & " zomer"
  Case Is = 5
    Halfjaar = year(myDate) & " zomer"
  Case Is = 6
    Halfjaar = year(myDate) & " zomer"
  Case Is = 7
    Halfjaar = year(myDate) & " zomer"
  Case Is = 8
    Halfjaar = year(myDate) & " zomer"
  Case Is = 9
    Halfjaar = year(myDate) & " zomer"
  Case Is = 10
    Halfjaar = year(myDate) & "-" & VBA.Right(year(myDate) + 1, 2) & " winter"
  Case Is = 11
    Halfjaar = year(myDate) & "-" & VBA.Right(year(myDate) + 1, 2) & " winter"
  Case Is = 12
    Halfjaar = year(myDate) & "-" & VBA.Right(year(myDate) + 1, 2) & " winter"
  End Select
  
End Function

Public Function METEOROLOGISCHSEIZOEN(myDate As Date) As String
  'geeft het meteorologische seizoen van een datum terug
  If month(myDate) <= 2 Or month(myDate) = 12 Then
    METEOROLOGISCHSEIZOEN = "winter"
  ElseIf month(myDate) < 6 Then
    METEOROLOGISCHSEIZOEN = "lente"
  ElseIf month(myDate) < 9 Then
    METEOROLOGISCHSEIZOEN = "zomer"
  ElseIf month(myDate) < 12 Then
    METEOROLOGISCHSEIZOEN = "herfst"
  End If
End Function

Public Function METEOROLOGISCHHALFJAAR(myDate As Date) As String
  'geeft het meteorologische halfjaar van een datum terug
  If month(myDate) <= 3 Then
    METEOROLOGISCHHALFJAAR = "winter"
  ElseIf month(myDate) <= 9 Then
    METEOROLOGISCHHALFJAAR = "zomer"
  Else
    METEOROLOGISCHHALFJAAR = "winter"
  End If
End Function

Public Function HYDROLOGISCHSEIZOEN(myDate As Date, WinZomMonth As Long, WinZomDay As Long, ZomWinMonth As Long, ZomWinDay As Long) As String
  'geeft het hydrologisch seizoen van een datum terug
  If month(myDate) < WinZomMonth Then
    HYDROLOGISCHSEIZOEN = "winter"
  ElseIf month(myDate) > ZomWinMonth Then
    HYDROLOGISCHSEIZOEN = "winter"
  ElseIf month(myDate) > WinZomMonth And month(myDate) < ZomWinMonth Then
    HYDROLOGISCHSEIZOEN = "zomer"
  ElseIf month(myDate) = WinZomMonth Then
    If day(myDate) >= WinZomDay Then
      HYDROLOGISCHSEIZOEN = "zomer"
    Else
      HYDROLOGISCHSEIZOEN = "winter"
    End If
  ElseIf month(myDate) = ZomWinMonth Then
    If day(myDate) >= ZomWinDay Then
      HYDROLOGISCHSEIZOEN = "winter"
    Else
      HYDROLOGISCHSEIZOEN = "zomer"
    End If
  End If
  
End Function

Public Function DOUBLE2DATETIMESTRING(myDate As Variant, Optional DateSeparator As String = "/", Optional TimeSeparator As String = ":", Optional DateTimeSeparator As String = "-", Optional YearLen As Long = 4, Optional YearOrder As Integer = 1, Optional MonthOrder As Integer = 2, Optional DayOrder As Integer = 3, Optional HourOrder As Integer = 4, Optional MinuteOrder As Integer = 5, Optional SecondOrder As Integer = 6) As String
Dim YearStr As String
Dim MonthStr As String
Dim DayStr As String
Dim HourStr As String
Dim MinuteStr As String
Dim SecondStr As String

If YearOrder + MonthOrder + DayOrder + HourOrder + MinuteOrder + SecondOrder <> 21 Then
  DOUBLE2DATETIMESTRING = "Error, invalid order specified for datetime-elements"
  Exit Function
Else
  If YearLen = 2 Then
    YearStr = VBA.Format(year(myDate), "00")
  ElseIf YearLen = 4 Then
    YearStr = VBA.Format(year(myDate), "0000")
  Else
    DOUBLE2DATETIMESTRING = "Error, year must be in 2 or 4 digits, e.g. 12 or 2012"
    Exit Function
  End If
  
  MonthStr = VBA.Format(month(myDate), "00")
  DayStr = VBA.Format(day(myDate), "00")
  HourStr = VBA.Format(hour(myDate), "00")
  MinuteStr = VBA.Format(Minute(myDate), "00")
  SecondStr = VBA.Format(Second(myDate), "00")
  
  If YearOrder = 1 And MonthOrder = 2 And DayOrder = 3 And HourOrder = 4 And MinuteOrder = 5 And SecondOrder = 6 Then
    DOUBLE2DATETIMESTRING = YearStr & DateSeparator & MonthStr & DateSeparator & DayStr & DateTimeSeparator & HourStr & TimeSeparator & MinuteStr & TimeSeparator & SecondStr
    Exit Function
  ElseIf YearOrder = 3 And MonthOrder = 2 And DayOrder = 1 And HourOrder = 4 And MinuteOrder = 5 And SecondOrder = 6 Then
    DOUBLE2DATETIMESTRING = DayStr & DateSeparator & MonthStr & DateSeparator & YearStr & DateTimeSeparator & HourStr & TimeSeparator & MinuteStr & TimeSeparator & SecondStr
    Exit Function
  Else
    DOUBLE2DATETIMESTRING = "Error: specified order of date-time elements not (yet) supported."
    Exit Function
  End If
End If

End Function

Public Function DateExists(myYear As Long, myMonth As Long, myDay As Long) As Boolean

DateExists = True
If myDay < 1 Or myDay > 31 Then
  DateExists = False
ElseIf myMonth < 1 Or myMonth > 12 Then
  DateExists = False
ElseIf myMonth = 4 Or myMonth = 6 Or myMonth = 9 Or myMonth = 11 Then
  If myDay > 30 Then
    DateExists = False
  End If
ElseIf myMonth = 2 Then
  If myDay > 29 Then
    DateExists = False
  ElseIf myDay > 28 Then  'alleen geldig bij een schrikkeljaar
    If myYear / 4 <> Round(myYear / 4, 0) Then
      DateExists = False
    End If
  End If
End If

End Function

Public Function DayNumber(myDate As Date, AlwaysInclude29Feb As Boolean) As Integer
  Dim myMonth As Integer
  Dim myNum As Integer
  Dim i As Integer
  
  For i = 1 To 12
    If i = month(myDate) Then
      myNum = myNum + day(myDate)
      DayNumber = myNum
      Exit Function
    Else
      myNum = myNum + DaysInMonth2(i, year(myDate), AlwaysInclude29Feb)
    End If
  Next
  myMonth = month(myDate)
End Function

Public Function DATEHOURWINDOW(myDate As Variant) As Variant
  Dim myHour As Integer
  myHour = hour(myDate)
  
  DATEHOUR = DateSerial(year(myDate), month(myDate), day(myDate))
  DATEHOUR = DATEHOUR + myHour / 24
          
End Function

Public Function DATETWOHOURWINDOW(myDate As Variant) As Variant
  'Author: Siebe Bosch
  'Description: returns the date + the two-hour-window of the day a certain datetime-value falls in
  Dim myHour As Integer
  myHour = hour(myDate)
  DATETWOHOURWINDOW = DateSerial(year(myDate), month(myDate), day(myDate))
  
  If myHour < 2 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 1 / 24
  ElseIf myHour < 4 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 3 / 24
  ElseIf myHour < 6 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 5 / 24
  ElseIf myHour < 8 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 7 / 24
  ElseIf myHour < 10 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 9 / 24
  ElseIf myHour < 12 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 11 / 24
  ElseIf myHour < 14 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 13 / 24
  ElseIf myHour < 16 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 15 / 24
  ElseIf myHour < 18 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 17 / 24
  ElseIf myHour < 20 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 19 / 24
  ElseIf myHour < 22 Then
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 21 / 24
  Else
    DATETWOHOURWINDOW = DATETWOHOURWINDOW + 23 / 24
  End If
  
End Function

Public Function DATEFOURHOURWINDOW(myDate As Variant) As Variant
  'Author: Siebe Bosch
  'Description: returns the date + the quarter of the day a certain datetime-value falls in
  Dim myHour As Integer
  myHour = hour(myDate)
  DATEFOURHOURWINDOW = DateSerial(year(myDate), month(myDate), day(myDate))
  
  If myHour < 6 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 3 / 24  '3 is the middle between 0 and 6
  ElseIf myHour < 12 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 9 / 24  '9 is the middle between 6 and 12
  ElseIf myHour < 18 Then
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 12 / 24 '15 is the middle between 12 and 18
  Else
    DATEFOURHOURWINDOW = DATEFOURHOURWINDOW + 21 / 24 '21 is the middle between 18 and 24
  End If
  
End Function

Public Function DATEFROMSTRING(myDate As String, DateFormat As String) As Variant
   Dim myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMinute As Integer, mySecond As Integer
   
   Select Case DateFormat
     Case Is = "yyyymmddhh"
       myYear = Left(myDate, 4)
       myMonth = Left(Right(myDate, 6), 2)
       myDay = Left(Right(myDate, 4), 2)
       myHour = Right(myDate, 2)
    Case Is = "yyyymmdd"
      myYear = Left(myDate, 4)
      myMonth = Left(Right(myDate, 4), 2)
      myDay = Right(myDate, 2)
   End Select
   
   'corrigeer wanneer 00 uur als 24 wordt weergegeven
   If myHour = 24 Then
     myHour = 0
     myDay = myDay + 1
     If myDay > DaysInMonth2(myMonth, myYear) Then
       myDay = 1
       myMonth = myMonth + 1
       If myMonth > 12 Then
         myMonth = 1
         myYear = myYear + 1
       End If
     End If
   End If
      
   DATEFROMSTRING = DateValue(myYear & "-" & myMonth & "-" & myDay) + TimeValue(myHour & ":" & myMinute & ":" & mySecond)
End Function

Public Function DATE2TEXT(myDate As Date, Formatting As String, MidnightAs24 As Boolean) As String
    Dim myStr As String, hpos As Integer
    
    If MidnightAs24 Then
        If hour(myDate) = 0 And Minute(myDate) = 0 And Second(myDate) = 0 Then
            myStr = Format(myDate - 1, Formatting)
            hpos = InStr(1, Formatting, "hh", vbTextCompare)
            myStr = Left(myStr, hpos - 1) & "24" & Right(myStr, hpos + 2)
        End If
            myStr = Format(myDate, Formatting)
        Else
    End If
    DATE2TEXT = myStr
    
End Function


Public Function TIMEFROMSTRING(myDate As String, timeFormat As String) As Variant
  Dim myHour As Integer, myMinute As Integer, mySecond As Integer
   
  Select Case Trim(LCase(timeFormat))
    Case Is = "hm"
      If VBA.Len(myDate) = 2 Then
        myHour = 0
        myMinute = myDate
      ElseIf VBA.Len(myDate) = 3 Then
        myHour = Left(myDate, 1)
        myMinute = Right(myDate, 2)
      ElseIf VBA.Len(myDate) = 4 Then
        myHour = Left(myDate, 2)
        myMinute = Right(myDate, 2)
      End If
    Case Is = "hhmm"
      myHour = VBA.Left(myDate, 2)
      myMinute = VBA.Right(myDate, 2)
    Case Is = "hhmmss"
      myHour = VBA.Left(myDate, 2)
      myMinute = VBA.Mid(myDate, 3, 2)
      mySecond = VBA.Right(myDate, 2)
   End Select
      
   TIMEFROMSTRING = TimeValue(myHour & ":" & myMinute & ":" & mySecond)
End Function


Public Function DATEANDTIMEFROMSTRINGS(myDateStr As String, myTimeStr As String, DateFormat As String, timeFormat As String) As Variant
  Dim myDay As Integer, myMonth As Integer, myYear As Integer
  Dim myHour As Integer, myMinute As Integer, mySecond As Integer
  Dim myDate As Date
  
  myDate = DATEFROMSTRING(myDateStr, DateFormat)
   
  'timeformat doen we handmatig ivm het mogelijk voorkomen van 24:00
  Select Case timeFormat
    Case Is = "hhmm"
      If VBA.Len(myTimeStr) = 2 Then
        myHour = 0
        myMinute = myTimeStr
      ElseIf VBA.Len(myTimeStr) = 3 Then
        myHour = Left(myTimeStr, 1)
        myMinute = Right(myTimeStr, 2)
      ElseIf VBA.Len(myTimeStr) = 4 Then
        myHour = Left(myTimeStr, 2)
        myMinute = Right(myTimeStr, 2)
      End If
   End Select
      
   If myHour = 24 Then
     myDate = myDate + 1
     myHour = 0
   End If
      
   DATEANDTIMEFROMSTRINGS = myDate + TimeValue(myHour & ":" & myMinute & ":" & mySecond)
   
   
End Function

Public Function HeleUrenUitUren(Uren As Double) As Integer
    HeleUrenUitUren = Application.WorksheetFunction.RoundDown(Uren, 0)
End Function

Public Function MinutenUitUren(Uren As Double) As Integer
    MinutenUitUren = 60 * (Uren - Application.WorksheetFunction.RoundDown(Uren, 0))
End Function


Public Function Decade(myDate As Date) As Integer
    'returns the decade number for a given date.
    'note: 36 is the highest possible value. All values will be limited to 36
    Decade = Maximum(1, Minimum(WorksheetFunction.RoundUp((myDate - DateSerial(year(myDate), 1, 1)) / 10, 0), 36))
End Function

Public Function SOBEKTIMETABLESTRING(WinZomMonthStart As Integer, WinZomDayStart As Integer, WinZomMonthEnd As Integer, WinZomDayEnd As Integer, ZomWinMonthStart As Integer, ZomWinDayStart As Integer, ZomWinMonthEnd As Integer, ZomWinDayEnd As Integer, WinVal As Variant, ZomVal As Variant) As String
    Dim myStr As String
    Dim i As Integer
        myStr = "'1906/01/01;00:00:00' " & WinVal & " <"
    For i = 1906 To 2050
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(WinZomMonthStart, "00") & "/" & Format(WinZomDayStart, "00") & ";00:00:00' " & WinVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(WinZomMonthEnd, "00") & "/" & Format(WinZomDayEnd, "00") & ";00:00:00' " & ZomVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(ZomWinMonthStart, "00") & "/" & Format(ZomWinDayStart, "00") & ";00:00:00' " & ZomVal & " <"
        myStr = myStr & vbCrLf & "'" & i & "/" & Format(ZomWinMonthEnd, "00") & "/" & Format(ZomWinDayEnd, "00") & ";00:00:00' " & WinVal & " <"
    Next
    SOBEKTIMETABLESTRING = myStr
End Function

Function AvgExceedenceValueContiguous(dates As Range, values As Range, continuousTimesteps As Integer) As Double
    Dim dateArr As Variant
    Dim valArr As Variant
    Dim i As Long, j As Long
    Dim minVal As Double, maxVal As Double
    Dim threshold As Double
    Dim eventsPerYear As Double
    Dim yearCount As Double
    Dim years As Object
    
    ' Convert Ranges to Arrays
    dateArr = dates.value
    valArr = values.value
    
    ' Find minimum and maximum values in the range
    minVal = Application.WorksheetFunction.Min(values)
    maxVal = Application.WorksheetFunction.max(values)
    
    ' Initialize threshold
    threshold = (minVal + maxVal) / 2
    
    ' Count the number of different years in the date range
    Set years = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dateArr, 1)
        Dim yr As Integer
        yr = year(dateArr(i, 1))
        If Not years.Exists(yr) Then years.Add yr, 1
    Next i
    yearCount = years.Count
    
    ' Iterative approach to adjust threshold
    Dim iter As Integer
    Dim eventCount As Integer
    For iter = 1 To 1000 ' You can adjust the number of iterations as needed
        eventCount = 0
        For i = 1 To UBound(valArr, 1) - continuousTimesteps + 1
            Dim exceedCount As Integer
            exceedCount = 0
            For j = i To i + continuousTimesteps - 1
                If valArr(j, 1) > threshold Then
                    exceedCount = exceedCount + 1
                Else
                    Exit For
                End If
            Next j
            If exceedCount = continuousTimesteps Then
                eventCount = eventCount + 1
                i = i + continuousTimesteps - 1 ' Skip over the continuous period we've counted
            End If
        Next i
        eventsPerYear = eventCount / yearCount
        
        ' Adjust threshold based on eventsPerYear
        If eventsPerYear > 1 Then
            minVal = threshold
        Else
            maxVal = threshold
        End If
        
        ' Update threshold
        Dim oldThreshold As Double
        oldThreshold = threshold
        threshold = (minVal + maxVal) / 2
        
        ' Convergence check (optional)
        If Abs(threshold - oldThreshold) < 0.001 Then Exit For ' Adjust convergence criterion as needed
    Next iter
    
    ' Set the function result to the found threshold
    AvgExceedenceValueContiguous = threshold
End Function





Public Function DecadeStartDate(year As Integer, Decade As Integer) As Date
    DecadeStartDate = DateSerial(year, 1, 1) - 10 + Decade * 10
End Function

Public Function PAIR_LOOKUP_WORKSHEET(Bereik As Range, IDRow As Integer, LookupValColRelative As Integer, ReturnValColRelative As Integer, LookupID As String, LookupVal As Variant) As Variant
    'this function looks up a value in tables where every item has TWO columns
    'specify the row that contains the ID you're looking a value for
    'specify in which column, relative to the column that contains the ID, the Lookup Value is located
    'specify in which column, relative to the column that contains the ID, the Return Value is located
    Dim c As Integer, r As Integer
    c = 0
    Do While Not Bereik.Cells(IDRow, c + 1) = ""
      c = c + 1
      If Bereik.Cells(IDRow, c) = LookupID Then
        Do While Not Bereik.Cells(r + 1, c) = ""
            r = r + 1
            If Bereik.Cells(r, c + LookupValColRelative) = LookupVal Then
              PAIR_LOOKUP_WORKSHEET = Bereik.Cells(r, c + ReturnValColRelative)
              Exit Function
            End If
        Loop
        Exit Do
      End If
    Loop
    End Function

Public Sub PAIR_LOOKUP_ACTIVEX(InputRange As Range, InputIDRow As Integer, outputRange As Range, ProgressRange As Range)
    'this function looks up a value in tables where every item has TWO columns
    'in the input range we expect the column containing the looked up value to be the same as the one containing the ID
    'in the input range we expect the column containing the return value to be immediately to the right of the lookup column
    Dim ro As Integer, co As Integer, ri As Integer, ci As Integer, n As Long
    Dim LookupID As String, LookupVal As Variant
    Dim MaxNum As Long
    MaxNum = (outputRange.Rows.Count - 1) * (outputRange.Columns.Count - 1)
    For ro = 2 To outputRange.Rows.Count
        For co = 2 To outputRange.Columns.Count
            n = n + 1
            ProgressRange.Cells(1, 1) = n / MaxNum
            LookupID = outputRange.Cells(ro, 1)
            LookupVal = outputRange.Cells(1, co)
            
            For ci = 1 To InputRange.Columns.Count
                If InputRange.Cells(InputIDRow, ci) = LookupID Then
                    For ri = 1 To InputRange.Rows.Count
                        If InputRange.Cells(ri, ci) = LookupVal Then
                            outputRange.Cells(ro, co) = InputRange.Cells(ri, ci + 1)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        Next
        DoEvents
    Next

    End Sub


Public Function VERT_HORIZ_ZOEKEN(Bereik As Range, ZoekVerticaal As String, ZoekHorizontaal As String) As Variant

Dim Kolom As Variant
Dim Rij As Variant
Dim KolomTeller As Integer
Dim RijTeller As Integer
Dim ZoekKolom As Integer
Dim ZoekRij As Integer

    KolomTeller = 0
    RijTeller = 0

    For Each Kolom In Bereik.Columns
        KolomTeller = KolomTeller + 1
        If UCase(Kolom.Columns.Cells(1, 1).value) = UCase(ZoekHorizontaal) Then
            ZoekKolom = KolomTeller
        End If
    Next Kolom
    
    For Each Rij In Bereik.Rows
        RijTeller = RijTeller + 1
        If UCase(Rij.Rows.Cells(1, 1).value) = UCase(ZoekVerticaal) Then
            ZoekRij = RijTeller
        End If
    Next Rij

    If ZoekKolom = 0 Or ZoekRij = 0 Then
        VERT_HORIZ_ZOEKEN = 0
    Else
        VERT_HORIZ_ZOEKEN = Bereik.Cells(ZoekRij, ZoekKolom).value
    End If

End Function

Public Function VERT_ZOEKEN_DOUBLE(SeekValue1 As Variant, SeekValue2 As Variant, MyRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van twee criteria: een waarde in kol1 en een in kol2
  Dim r As Long
  
  VERT_ZOEKEN_DOUBLE = Null
  If MyRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 3 Then Exit Function
  
  For r = 1 To MyRange.Rows.Count
    If Trim(UCase(MyRange.Cells(r, 1))) = Trim(UCase(SeekValue1)) And Trim(UCase(MyRange.Cells(r, 2))) = Trim(UCase(SeekValue2)) Then
      VERT_ZOEKEN_DOUBLE = MyRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function


Public Function HOR_ZOEKEN_DOUBLE(SeekValue1 As Variant, SeekValue2 As Variant, MyRange As Range, ReturnRow As Long) As Variant
  'Deze functie is een uitbreiding op horizontaal zoeken, namelijk dat hij zoekt op basis van twee criteria: een waarde in rij1 en een in rij2
  Dim c As Long
  
  HOR_ZOEKEN_DOUBLE = Null
  If MyRange.Rows.Count < ReturnRow Then Exit Function
  If ReturnRow < 3 Then Exit Function
  
  For c = 1 To MyRange.Columns.Count
    If MyRange.Cells(1, c) = SeekValue1 And MyRange.Cells(2, c) = SeekValue2 Then
      HOR_ZOEKEN_DOUBLE = MyRange.Cells(ReturnRow, c)
      Exit Function
    End If
  Next
  
  HOR_ZOEKEN_DOUBLE = ""
  
End Function

Public Function VERT_ZOEKEN_TRIPLE(SeekValue1 As Variant, SeekValue2 As Variant, SeekValue3 As Variant, MyRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van DRIE criteria: een waarde in kol1 en een in kol2, een in Kol3
  Dim r As Long
  
  VERT_ZOEKEN_TRIPLE = Null
  If MyRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 4 Then Exit Function
  
  For r = 1 To MyRange.Rows.Count
    If MyRange.Cells(r, 1) = SeekValue1 And MyRange.Cells(r, 2) = SeekValue2 And MyRange.Cells(r, 3) = SeekValue3 Then
      VERT_ZOEKEN_TRIPLE = MyRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function


Public Function VERT_ZOEKEN_QUADRUPLE(SeekValue1 As Variant, SeekValue2 As Variant, SeekValue3 As Variant, SeekValue4 As Variant, MyRange As Range, ReturnCol As Long) As Variant
  'Deze functie is een uitbreiding op vertikaal zoeken, namelijk dat hij zoekt op basis van VIER criteria: een waarde in kol1 en een in kol2, een in Kol3 en een in Kol4
  Dim r As Long
  
  VERT_ZOEKEN_QUADRUPLE = Null
  If MyRange.Columns.Count < ReturnCol Then Exit Function
  If ReturnCol < 5 Then Exit Function
  
  For r = 1 To MyRange.Rows.Count
    If MyRange.Cells(r, 1) = SeekValue1 And MyRange.Cells(r, 2) = SeekValue2 And MyRange.Cells(r, 3) = SeekValue3 And MyRange.Cells(r, 4) = SeekValue4 Then
      VERT_ZOEKEN_QUADRUPLE = MyRange.Cells(r, ReturnCol)
      Exit Function
    End If
  Next

End Function

Public Function VERT_ZOEKEN_MIN(id As String, MyRange As Range, ValueColIdx As Long, Optional SkipZero As Boolean = False) As Variant
 Dim r As Long, c As Long, myMin As Variant, n As Long
 
 For r = 1 To MyRange.Rows.Count
   If MyRange.Cells(r, 1) = id And Not (MyRange.Cells(r, ValueColIdx) = 0 And SkipZero = True) Then
     n = n + 1
     If n = 1 Then
       myMin = MyRange.Cells(r, ValueColIdx)
     Else
       If MyRange.Cells(r, ValueColIdx) < myMin Then myMin = MyRange.Cells(r, ValueColIdx)
     End If
   End If
  Next
  
  If n = 0 Then
    VERT_ZOEKEN_MIN = Nothing
  Else
    VERT_ZOEKEN_MIN = myMin
  End If
  
End Function

Public Function VERT_ZOEKEN_MAX(id As String, MyRange As Range, ValueColIdx As Long, Optional SkipZero As Boolean = False) As Variant
 Dim r As Long, c As Long, myMax As Variant, n As Long
 
 For r = 1 To MyRange.Rows.Count
   If MyRange.Cells(r, 1) = id And Not (MyRange.Cells(r, ValueColIdx) = 0 And SkipZero = True) Then
     n = n + 1
     If n = 1 Then
       myMax = MyRange.Cells(r, ValueColIdx)
     Else
       If MyRange.Cells(r, ValueColIdx) > myMax Then myMax = MyRange.Cells(r, ValueColIdx)
     End If
   End If
  Next
  
  If n = 0 Then
    VERT_ZOEKEN_MAX = -999
  Else
    VERT_ZOEKEN_MAX = myMax
  End If

End Function

Public Function VERT_ZOEKEN_GROOTSTEAANDEELHOUDER(id As String, MyRange As Range, AandelhoudersColIdx As Long, ValueColIdx As Long, Optional Absoluut As Boolean = False) As Variant
 Dim r As Long, c As Long, i As Long, GrAand As String, mySum As Variant
 Dim Aandeelhouders() As String
 Dim WaardeSom() As Variant
 
 ReDim Aandeelhouders(1 To MyRange.Rows.Count)  'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 ReDim WaardeSom(1 To MyRange.Rows.Count)       'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 
 For r = 1 To MyRange.Rows.Count
   If MyRange.Cells(r, 1) = id Then
     For i = 1 To UBound(Aandeelhouders())
       If Aandeelhouders(i) = MyRange.Cells(r, AandelhoudersColIdx) Or Aandeelhouders(i) = "" Then
         Aandeelhouders(i) = MyRange.Cells(r, AandelhoudersColIdx)
         If Not Absoluut Then
           WaardeSom(i) = WaardeSom(i) + MyRange.Cells(r, ValueColIdx)
         Else
           WaardeSom(i) = WaardeSom(i) + VBA.Abs(MyRange.Cells(r, ValueColIdx))
         End If
         Exit For
       End If
     Next
   End If
 Next

 For i = 1 To UBound(Aandeelhouders())
   If WaardeSom(i) > mySum Then
     mySum = WaardeSom(i)
     GrAand = Aandeelhouders(i)
   End If
 Next
  
 VERT_ZOEKEN_GROOTSTEAANDEELHOUDER = GrAand
  
End Function

Public Function HEADERBYMAXIMUMVALUE(HeaderRange As Range, ValuesRange As Range) As Variant
 'this function returns the header value that corresponds with the column containing the largest value in a given range
 Dim myVal As Variant
 Dim maxVal As Variant
 Dim Header As Variant
 maxVal = -9.99E+101
 Dim c As Long
 
 If HeaderRange.Columns.Count <> ValuesRange.Columns.Count Then
   HEADERBYMAXIMUMVALUE = "Number of columns for header and values must be equal"
 ElseIf HeaderRange.Rows.Count > 1 Then
   HEADERBYMAXIMUMVALUE = "Header Range can only have one row"
 ElseIf HeaderRange.Rows.Count > 1 Then
   HEADERBYMAXIMUMVALUE = "Values Range can only have one row"
 Else
   For c = 1 To ValuesRange.Columns.Count
     If ValuesRange.Cells(1, c) > maxVal Then
       maxVal = ValuesRange.Cells(1, c)
       Header = HeaderRange.Cells(1, c)
     End If
   Next
 End If
 
 HEADERBYMAXIMUMVALUE = Header
  
End Function

Public Function VERT_ZOEKEN_NEARESTXY(x As Variant, y As Variant, XYVALRANGE As Range, ReturnColIdx As Long, Optional XColIdx As Long = 1, Optional YColIdx As Long = 2) As Variant
  'zoekt voor een gegeven X en Y het meest dichtbijzijnde object uit een range met X,Y en geeft de waarde uit een gespecificeerde kolom terug
  Dim r As Long
  Dim Dist As Variant, minDist As Variant
  Dim dX As Variant, dY As Variant
  Dim minDistVal As Variant  'de waarde die moet worden teruggegeven
  minDist = 9999999999999#
  
  For r = 1 To XYVALRANGE.Rows.Count
    dX = x - XYVALRANGE.Cells(r, XColIdx)
    dY = y - XYVALRANGE.Cells(r, YColIdx)
    Dist = VBA.Math.Sqr(dX ^ 2 + dY ^ 2)
    If Dist < minDist Then
      minDist = Dist
      minDistVal = XYVALRANGE.Cells(r, ReturnColIdx)
    End If
  Next
  VERT_ZOEKEN_NEARESTXY = minDistVal
  
End Function


Public Function VERT_ZOEKEN_MODUS(id As String, MyRange As Range, ValueColIdx As Long) As Variant
 'geeft de meest voorkomende waarde terug behorende bij een vooraf vastgesteld ID
 Dim r As Long, c As Long, i As Long, Found As Boolean, nPoints As Long, myModus As String, MaxNum As Long
 Dim values() As String
 Dim n() As Long
 
 ReDim values(1 To MyRange.Rows.Count)  'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 ReDim n(1 To MyRange.Rows.Count)       'dimensioneer op het meest pessimistische geval: alleen maar unieke waarden
 
 For r = 1 To MyRange.Rows.Count
   If MyRange.Cells(r, 1) = id Then
     For i = 1 To UBound(values())
       If values(i) = MyRange.Cells(r, ValueColIdx) Or values(i) = "" Then
         values(i) = MyRange.Cells(r, ValueColIdx)
         n(i) = n(i) + 1
         Exit For
       End If
     Next
   End If
  Next
    
  MaxNum = 0
  For i = 1 To UBound(values())
    If n(i) > MaxNum Then
      MaxNum = n(i)
      myModus = values(i)
    End If
  Next
  VERT_ZOEKEN_MODUS = myModus

End Function

Public Function VERT_ZOEKEN_SOM(id As String, MyRange As Range, ValueColIdx As Long) As Variant
  'geeft de som terug van alle waarden uit kolomnr ValueColIdx achter een ID in kolom 1 met vooraf vastgestelde waarde
  Dim r As Long, mySum As Variant
 
  mySum = 0
  For r = 1 To MyRange.Rows.Count
   If MyRange.Cells(r, 1) = id Then mySum = mySum + MyRange.Cells(r, ValueColIdx)
  Next
  VERT_ZOEKEN_SOM = mySum

End Function

Public Function FindColumnInRange(MyRange As Range, SeekValue As Variant, assignEmptyColumnIfNotFound As Boolean) As Long
  'deze functie geeft de kolomindex terug, gegeven een gezochte waarde.
  'Let op: de range mag slechts 1 cel hoog zijn.
  Dim FirstEmpty As Long 'de eerst lege kolom die hij tegenkomt
  Dim c As Long
  
  For c = 1 To MyRange.Columns.Count
    If MyRange.Cells(1, c) = SeekValue Then
      FindColumnInRange = c
      Exit Function
    ElseIf MyRange.Cells(1, c) = "" Then
      If FirstEmpty = 0 Then FirstEmpty = c
    End If
  Next
  
  'als hij hier aankomt, heeft hij niets gevonden. Ken dus een nieuwe kolom toe voor de data
  If assignEmptyColumnIfNotFound = True Then
    If FirstEmpty > 0 Then
      FindColumnInRange = FirstEmpty
    Else
      'tel een bij de laatste op
      FindColumnInRange = c + 1
    End If
  Else
    FindColumnInRange = 0
  End If
  
End Function

Public Function FindRowInRange(MyRange As Range, SeekValue As Variant, assignEmptyRowIfNotFound As Boolean) As Long
  'deze functie geeft de rijindex terug, gegeven een gezochte waarde.
  'Let op: de range mag slechts 1 cel breed zijn.
  Dim FirstEmpty As Long 'de eerst lege rij die hij tegenkomt
  Dim r As Long
  
  For r = 1 To MyRange.Rows.Count
    If MyRange.Cells(r, 1) = SeekValue Then
      FindRowInRange = r
      Exit Function
    ElseIf MyRange.Cells(r, 1) = "" Then
      If FirstEmpty = 0 Then FirstEmpty = r
    End If
  Next
  
  'als hij hier aankomt, heeft hij niets gevonden. Ken dus een nieuwe rij toe voor de data
  If assignEmptyRowIfNotFound = True Then
    If FirstEmpty > 0 Then
      FindRowInRange = FirstEmpty
    Else
      'tel een bij de laatste op
      FindRowInRange = r + 1
    End If
  Else
    FindRowInRange = 0
  End If
  
End Function

Public Function AVERAGEFROMRANGE(MyRange As Range, c As Integer, ConditionalColumn As Integer, Condition As String, Optional ByVal UseFirstIfNoFound As Boolean = True) As Variant
  Dim r As Integer, a As Integer, n As Integer
  Dim subRange As Range
  Dim condVal As Variant, myVal As Variant, mySum As Variant
  
  If MyRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function AVERAGEFROMRANGE.")
    End
  End If
  
  If Condition <> "" And ConditionalColumn > 0 Then
    For r = 1 To MyRange.Rows.Count
      myVal = MyRange.Cells(r, c)
      condVal = MyRange.Cells(r, ConditionalColumn)
      If MeetsCondition(condVal, Condition) Then
        n = n + 1
        mySum = mySum + myVal
      End If
    Next
    
    If n > 0 Then
      AVERAGEFROMRANGE = mySum / n
    Else
      If UseFirstIfNoFound Then
        AVERAGEFROMRANGE = MyRange.Cells(1, c)
      Else
        AVERAGEFROMRANGE = -999
      End If
    End If
    
  Else
    AVERAGEFROMRANGE = Application.WorksheetFunction.Average(MyRange.Range(MyRange.Cells(1, c), MyRange.Cells(MyRange.Rows.Count, c)))
  End If

End Function

Public Function MINFROMRANGE(MyRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim minVal As Variant
  minVal = 9999999999999#

  If MyRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function MINFROMRANGE.")
    End
  End If
  
  For r = 1 To MyRange.Rows.Count
    myVal = MyRange.Cells(r, c)
    condVal = MyRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      If myVal < minVal Then minVal = myVal
    End If
  Next

  MINFROMRANGE = minVal
End Function

Public Function MAXFROMRANGE(MyRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim maxVal As Variant
  maxVal = -9999999999999#

  If MyRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function MAXFROMRANGE.")
    End
  End If
  
  For r = 1 To MyRange.Rows.Count
    myVal = MyRange.Cells(r, c)
    condVal = MyRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      If myVal > maxVal Then maxVal = myVal
    End If
  Next

  MAXFROMRANGE = maxVal
End Function

Public Function FIRSTFROMRANGE(MyRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim r As Integer, condVal As Variant, myVal As Variant
  If MyRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function FIRSTFROMRANGE.")
    End
  End If
  
  For r = 1 To MyRange.Rows.Count
    myVal = MyRange.Cells(r, c)
    condVal = MyRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      FIRSTFROMRANGE = myVal
      Exit Function
    End If
  Next
  
End Function

Public Function LASTFROMRANGE(MyRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "") As Variant
  Dim r As Integer, condVal As Variant, myVal As Variant

  If MyRange.areas.Count > 1 Then
    MsgBox ("Error: ranges with multiple areas are not (yet) supported in function LASTFROMRANGE.")
    End
  End If

  For r = MyRange.Rows.Count To 1 Step -1
    myVal = MyRange.Cells(r, c)
    condVal = MyRange.Cells(r, ConditionalColumn)
    If MeetsCondition(condVal, Condition) Then
      LASTFROMRANGE = myVal
      Exit Function
    End If
  Next
  
  End Function

Public Function MOSTCOMMONFROMRANGE(MyRange As Range, c As Integer, Optional ConditionalColumn As Integer = 0, Optional Condition As String = "", Optional ByVal UseFirstIfNoFound As Boolean = True) As Variant
Dim r As Long, i As Long, a As Long, Found As Boolean
Dim myVals() As Variant, myNumbers() As Long
Dim MaxNum As Long, myVal As Variant, condVal As Variant, n As Long

'This function vinds the most common value in a range.

If MyRange.areas.Count > 1 Then
  MsgBox ("Error: ranges with multiple areas are not (yet) supported in function LASTFROMRANGE.")
  End
End If

n = 0

For r = 1 To MyRange.Rows.Count

  myVal = MyRange.Cells(r, c).value
  condVal = MyRange.Cells(r, ConditionalColumn).value

  If MeetsCondition(condVal, Condition) Then
    
    Found = False
    If n > 0 Then
      For i = 1 To UBound(myVals)
        If myVal = myVals(i) Then
          myVals(i) = MyRange.Cells(r, c)
          myNumbers(i) = myNumbers(i) + 1
          Found = True
          Exit For
        End If
      Next
    End If
    
    'if the value was not yet found in the array, add it
    If Not Found Then
      n = n + 1
      ReDim Preserve myVals(1 To n)
      ReDim Preserve myNumbers(1 To n)
      myVals(n) = MyRange.Cells(r, c)
      myNumbers(n) = 1
    End If
  
  End If
Next

MaxNum = 0

If n > 0 Then
  For i = 1 To UBound(myVals)
    If myNumbers(i) > MaxNum Then
      myVal = myVals(i)
      MaxNum = myNumbers(i)
    End If
  Next
ElseIf UseFirstIfNoFound Then
  myVal = MyRange.Cells(1, c)
Else
  myVal = -999
End If

MOSTCOMMONFROMRANGE = myVal

End Function

Public Sub GEWOGEN_GEMIDDELDE(MyRange As Range, IDColIdx As Long, ValColIdx As Long, WeightValColIdx, ResultsRow As Long, ResultsCol As Long, Optional HasHeader As Boolean = True)
  'berekent een gewogen gemiddelde waarde voor ieder ID, gewogen naar bijv. oppervlaktes
  Dim myID As String, checkID As Variant, myResult As Variant, myWeight As Variant, SumOfWeights As Variant
  Dim IDsDone As Collection, IDDone As Boolean
  Dim Vals As Collection, Weights As Collection
  Dim r As Long, r2 As Long, r3 As Long, c As Long, i As Long, startRow As Long
  
  r3 = ResultsRow
  c = ResultsCol
  ActiveSheet.Cells(r3, c) = "ID"
  ActiveSheet.Cells(r3, c + 1) = "Gewogen gemiddelde"
  
  If HasHeader Then
    startRow = 2
  Else
    startRow = 1
  End If
  
  Set IDsDone = New Collection
  
  For r = startRow To MyRange.Rows.Count
    myID = ActiveSheet.Cells(r, IDColIdx)
    IDDone = False
    For Each checkID In IDsDone
      If checkID = MyRange.Cells(r, IDColIdx) Then
        IDDone = True
        Exit For
      End If
    Next
    
    If Not IDDone Then
      Set Vals = New Collection
      Set Weights = New Collection
      Vals.Add (ActiveSheet.Cells(r, ValColIdx))
      Weights.Add (ActiveSheet.Cells(r, WeightValColIdx))
      For r2 = r + 1 To MyRange.Rows.Count
        If ActiveSheet.Cells(r2, IDColIdx) = myID Then
          Vals.Add (ActiveSheet.Cells(r2, ValColIdx))
          Weights.Add (ActiveSheet.Cells(r2, WeightValColIdx))
        End If
      Next
      
      SumOfWeights = 0
      'bereken de som van alle gewichten
      For Each myWeight In Weights
        SumOfWeights = SumOfWeights + myWeight
      Next
      
      myResult = 0
      'bereken de gewogen waarde
      For i = 1 To Vals.Count
        If SumOfWeights <> 0 Then
          myResult = myResult + Vals(i) * Weights(i) / SumOfWeights
        Else
          myResult = 0
        End If
      Next
      
      'schrijf weg
      r3 = r3 + 1
      ActiveSheet.Cells(r3, c) = myID
      ActiveSheet.Cells(r3, c + 1) = myResult
      
     Call IDsDone.Add(myID)
    End If
    
  Next
  
End Sub

Public Sub AGGREGEREN(MyRange As Range, ResultsRow As Long, ResultsCol As Long, ExportEachnRows As Long)
  'Assumes date/time in first column and a header row.
  Dim r As Long, c As Long, r2 As Long, c2 As Long
  
  r2 = ResultsRow
  c2 = ResultsCol
  ActiveSheet.Cells(r2, c2) = "Datum/Tijd"
  For c = 2 To MyRange.Columns.Count
    ActiveSheet.Cells(r2, c2 + c - 1) = MyRange.Cells(1, c)
  Next
    
  For r = 2 To MyRange.Rows.Count Step ExportEachnRows
    DoEvents
    r2 = r2 + 1
    
    ActiveSheet.Cells(r2, c2) = MyRange(r, 1)  'write date/time
    For c = 2 To MyRange.Columns.Count
      ActiveSheet.Cells(r2, c2 + c - 1) = MyRange(r, c)
    Next
  Next
End Sub

Public Sub AGGREGATEFROMRANGE(RangeIncludingHeader As Range, AggregateByColumn As Integer, AggregateColumn As Integer, myMethod As enmAggregateMethod, ResultsRow As Integer, ResultsCol As Integer)
  Dim R1 As Long, r2 As Long, r As Long
  Dim startRow As Long, endRow As Long
  Dim myVal As Variant, Col1Val As Variant
  Dim subRange As Range
  Dim i As Long
  
  Dim curVal As Variant
  curVal = ""
  
  'write the results header
  ActiveSheet.Cells(ResultsRow, ResultsCol) = RangeIncludingHeader.Cells(1, AggregateByColumn)
  ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = RangeIncludingHeader.Cells(1, AggregateColumn)
  
  'walk through the data and find unique blocks based on the AggregateColumn
  For R1 = 2 To RangeIncludingHeader.Rows.Count
    If RangeIncludingHeader.Cells(R1, AggregateByColumn) <> curVal And RangeIncludingHeader.Cells(R1, AggregateByColumn) <> "" Then
       curVal = RangeIncludingHeader.Cells(R1, AggregateByColumn)
       startRow = R1
       For r2 = R1 + 1 To RangeIncludingHeader.Rows.Count
       
         'as soon as the next row in the aggregatebycolumn colummn changes, exit the loop and compute the aggregated value
         If RangeIncludingHeader.Cells(r2, AggregateByColumn) <> RangeIncludingHeader.Cells(R1, AggregateByColumn) Then
           endRow = r2 - 1
           R1 = endRow
           Exit For
         End If
       Next
       
       If endRow > startRow Then
       
       ResultsRow = ResultsRow + 1
       Set subRange = RangeIncludingHeader.Range(RangeIncludingHeader.Cells(startRow, AggregateColumn), RangeIncludingHeader.Cells(endRow, AggregateColumn))
         
         Col1Val = RangeIncludingHeader.Cells(startRow, AggregateByColumn)
         
         If myMethod = Average Then
'           If Application.WorksheetFunction.Sum(SubRange) > 0 Then
'             myVal = Application.Average(SubRange)
'           Else
'             myVal = 0
'           End If
            Dim mySum As Variant
            mySum = 0
            For i = startRow To endRow
              mySum = mySum + RangeIncludingHeader.Cells(i, AggregateColumn)
            Next
            If mySum > 0 Then
              myVal = mySum / (endRow - startRow + 1)
            Else
              myVal = 0
            End If
         ElseIf myMethod = first Then
           myVal = RangeIncludingHeader.Cells(startRow, AggregateColumn).value
         ElseIf myMethod = Last Then
           myVal = RangeIncludingHeader.Cells(endRow, AggregateColumn).value
         ElseIf myMethod = Largest Then
           myVal = Application.max(subRange)
         ElseIf myMethod = Smallest Then
           myVal = Application.Min(subRange)
         ElseIf myMethod = Most Then
           myVal = MOSTCOMMONFROMRANGE(subRange, 1)
         ElseIf myMethod = sum Then
            myVal = Application.sum(subRange)
         End If
         
        ActiveSheet.Cells(ResultsRow, ResultsCol) = Col1Val
        ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = myVal
       
       End If
         
    End If
  Next

End Sub


Public Sub AGGREGATERANGECONDITIONALLY(MyRange As Range, AggregateColumn As Integer, AggregateMethod() As enmAggregateMethod, ConditionalColumn As Integer, Condition() As String, ResultsRow As Integer, ResultsCol As Integer)
  Dim R1 As Integer, r2 As Integer, r As Integer, c As Integer, a As Integer
  Dim startRow As Integer, endRow As Integer
  Dim myMethod As enmAggregateMethod
  Dim myVal As Variant, subRange As Range, myCond As String
    
  If MyRange.Columns.Count <> UBound(AggregateMethod) Then
    MsgBox ("Error: array for aggregation method must have same dimensions as number of columns in range.")
    End
  End If
  
  Dim curVal As Variant
  curVal = ""
  
  'write the results header
  For c = 1 To MyRange.Columns.Count
    ActiveSheet.Cells(ResultsRow, ResultsCol + c - 1) = MyRange.Cells(1, c)
  Next
  
  'walk through the data and find unique blocks based on the AggregateColumn
  For R1 = 2 To MyRange.Rows.Count
    If ActiveSheet.Cells(R1, AggregateColumn) <> curVal And ActiveSheet.Cells(R1, AggregateColumn) <> "" Then
       curVal = ActiveSheet.Cells(R1, AggregateColumn)
       startRow = R1
       For r2 = R1 + 1 To MyRange.Rows.Count
         If ActiveSheet.Cells(r2, AggregateColumn) <> ActiveSheet.Cells(R1, AggregateColumn) Then
           endRow = r2 - 1
           Exit For
         End If
       Next
       
       ResultsRow = ResultsRow + 1
       
       'create a subrange for the block we're in
       Set subRange = MyRange.Range(MyRange.Cells(startRow, 1), MyRange.Cells(endRow, MyRange.Columns.Count))
       
       For c = 1 To subRange.Columns.Count
         myMethod = AggregateMethod(c)
         myCond = Condition(c)
         
         If myMethod = Average Then
           myVal = AVERAGEFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = first Then
           myVal = FIRSTFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Last Then
           myVal = LASTFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Largest Then
           myVal = MAXFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Smallest Then
           myVal = MINFROMRANGE(subRange, c, ConditionalColumn, myCond)
         ElseIf myMethod = Most Then
           myVal = MOSTCOMMONFROMRANGE(subRange, c, ConditionalColumn, myCond)
         End If
         
        ActiveSheet.Cells(ResultsRow, ResultsCol + c - 1) = myVal
         
       Next
    End If
  Next

End Sub

Public Function COLUMNFROMRANGE(MyRange As Range, ColNum As Integer) As Range
  Dim myArea As Range, newRange As Range, subRange As Range
  Dim a As Integer
  
  For a = 1 To MyRange.areas.Count
    Set myArea = MyRange.areas(a)
    Set subRange = myArea.Range(myArea.Cells(-1, ColNum), myArea.Cells(myArea.Rows.Count - 2, ColNum))
    If newRange Is Nothing Then
      Set newRange = subRange
    Else
      Set newRange = Union(newRange, subRange)
    End If
  Next
    
  Set COLUMNFROMRANGE = newRange

End Function

Public Function CONDITIONALSUBRANGE(ByVal MyRange As Range, ByVal ConditionColumn As Integer, ByVal Condition As String) As Range
  'this function applies a given condition to a range and only returns the rows for wich the condtion is met
  'conditions can be: "> x, >= x, < x, <= x, = x, <> x
  Dim newRange As Range, Range2 As Range
  Dim r As Integer
  Dim Operator As String, Operand As Variant
  Dim myVal As Variant, Inuse As Boolean
  
  Condition = VBA.Trim(Condition)
  
  If InStr(1, Condition, " ") <= 0 Then
    MsgBox ("Condition is not valid. Must contain space between operator and operand: " & Condition)
    End
  End If
  
  
  Operator = ParseString(Condition, " ")
  Operand = Condition
  
  For r = 1 To MyRange.Rows.Count
  
    'decide whether the condition is met for this row
    Inuse = False
    myVal = MyRange.Cells(r, ConditionColumn)
    Select Case Operator
      Case Is = ">"
         If myVal > Operand Then Inuse = True
      Case Is = ">="
         If myVal >= Operand Then Inuse = True
      Case Is = "<"
         If myVal < Operand Then Inuse = True
      Case Is = "<="
         If myVal <= Operand Then Inuse = True
      Case Is = "<>"
         If myVal <> Operand Then Inuse = True
      Case Is = "="
         If myVal = Operand Then Inuse = True
      Case Else
        MsgBox ("Error: operator in conditional formatting was not recognized or is not supported: " & Operator)
        End
    End Select
    
    'if the condition is met, add the row to our new range
    Dim n As Integer
    If Inuse = True Then
      n = n + 1
      If newRange Is Nothing Then
        Set newRange = MyRange.Range(MyRange.Cells(r, 1), MyRange.Cells(r, MyRange.Columns.Count))
      Else
        Set Range2 = MyRange.Range(MyRange.Cells(r, 1), MyRange.Cells(r, MyRange.Columns.Count))
        Set newRange = Union(newRange, Range2)
      End If
    End If
  Next
  
  Set CONDITIONALSUBRANGE = newRange

End Function



Public Sub AGGREGERENNAARUREN(DATETIMERANGE As Range, ValRange As Range, ProgressRange As Range, resultrow As Long, resultcol As Long, Optional HeleUren As Boolean = True)
Dim r As Long, r2 As Long, c2 As Long, lastProgress As Variant, Progress As Variant

'voorbeeld aanroep AGGREGERENNAARUREN(Range(Cells(2, 1), Cells(2000, 1)), Range(Cells(2, 2), Cells(2000, 2)), 2, 3)
r2 = resultrow
c2 = resultcol
ActiveSheet.Cells(r2, c2) = "Datum/Tijd"
ActiveSheet.Cells(r2, c2 + 1) = "Waarde"


If HeleUren = True Then
  For r = 1 To DATETIMERANGE.Rows.Count
    Progress = r / DATETIMERANGE.Rows.Count * 100
    If Round(Progress, 0) > Round(lastProgress, 0) Then
      ProgressRange.value = Progress
      DoEvents
      lastProgress = Progress
    End If
    
    If Minute(DATETIMERANGE(r, 1)) = 0 Then
      r2 = r2 + 1
      ActiveSheet.Cells(r2, c2) = DATETIMERANGE(r, 1)
      ActiveSheet.Cells(r2, c2 + 1) = ValRange(r, 1)
    End If
  Next
Else
  MsgBox ("Optie nog niet ondersteund")
End If

End Sub

Public Function CountSequentialExceedances(ValRange As Range, threshold As Variant) As Integer
  Dim r As Integer, n As Integer, nMax As Integer
  
  If ValRange.Columns.Count > 1 Then
    MsgBox ("Error in function CountSequentialExceedances. Range can have no more than one column.")
  Else
    For r = 1 To ValRange.Rows.Count
      If ValRange.Cells(r, 1) > threshold Then
        n = n + 1
        If n > nMax Then nMax = n
      Else
        n = 0
      End If
    Next
  End If
  
  CountSequentialExceedances = nMax

End Function


Public Sub GETASCIIGRIDVALUES(path As String, XYVALRANGE As Range, Optional XColIdx As Long = 1, Optional YColIdx As Long = 2, Optional ValColIdx As Long = 3)
  'haalt voor gegeven XY-coordinaten de bijbehorende waarde uit een ASCII-grid en schrijft deze naar het werkblad
  Dim data() As Variant
  Dim nCols As Long, nRows As Long, xllcorner As Variant, yllcorner As Variant, cellsize As Variant, nodata_value As Variant
  Dim x As Variant, y As Variant, VAL As Variant, rowIdx As Long, colIdx As Long
  Dim yulcorner As Variant, xlrcorner As Variant
  Dim r As Long, c As Long
  
  Call READASCIIGRID(path, nCols, nRows, xllcorner, yllcorner, cellsize, nodata_value, data)
  yulcorner = yllcorner + cellsize * nRows
  xlrcorner = xllcorner + cellsize * nCols
    
  For r = 1 To XYVALRANGE.Rows.Count
   
    x = XYVALRANGE(r, XColIdx)
    y = XYVALRANGE(r, YColIdx)
    
    If x >= xllcorner And x <= xlrcorner And y > yllcorner And y < yulcorner Then
      colIdx = Application.WorksheetFunction.RoundUp((x - xllcorner) / cellsize, 0)
      rowIdx = Application.WorksheetFunction.RoundUp((yulcorner - y) / cellsize, 0)
      XYVALRANGE.Cells(r, ValColIdx) = data(rowIdx, colIdx)
    Else
      XYVALRANGE.Cells(r, ValColIdx) = nodata_value
    End If
  Next
  
End Sub

Public Sub getRowColFromASCIIGRID(xllcenter As Variant, yllcenter As Variant, nCols As Long, nRows As Long, dX As Variant, dY As Variant, x As Variant, y As Variant, ByRef myRow As Long, ByRef myCol As Long)
  Dim xllcorner As Variant, yllcorner As Variant, yurcorner As Variant
  xllcorner = xllcenter - dX / 2
  yllcorner = yllcenter - dY / 2
  yurcorner = yllcorner + dY * nRows
  
  myCol = Application.WorksheetFunction.RoundUp((x - xllcorner) / dX, 0)
  If myCol <= 0 Or myCol > nCols Then myCol = 0
  
  myRow = Application.WorksheetFunction.RoundUp((yurcorner - y) / dY, 0)
  If myRow <= 0 Or myRow > nRows Then myRow = 0
  
End Sub


Public Sub RANGEWITHHEADER2THREECOLRANGE(MyRange As Range, HeaderTitle As String, ResultsRow As Long, ResultsCol As Long)
  'deze routine converteert een reeks waarin X en Y data staan en waarboven telkens een header staat naar een reeks met ID, X en Y in drie kolommen
  'dus van:
         'ID MyID
         'X1 Y1
         'X2 Y2
  'naar:
         'ID X1 Y1
         'ID X2 Y2
  Dim myID As String
  Dim rowIdx As Long
  Dim r As Long
  Dim c As Long
  
  r = ResultsRow - 1
  c = ResultsCol
  
  For rowIdx = 1 To MyRange.Rows.Count
    If MyRange.Cells(rowIdx, 1) = HeaderTitle Then
      myID = MyRange.Cells(rowIdx, 2)
    Else
      r = r + 1
      ActiveSheet.Cells(r, c) = myID
      ActiveSheet.Cells(r, c + 1) = MyRange.Cells(rowIdx, 1)
      ActiveSheet.Cells(r, c + 2) = MyRange.Cells(rowIdx, 2)
    End If
  Next

End Sub

Public Sub WEAVETABLESBLOCKINTERPOLATION(myTable1 As Range, myTable2 As Range, ResultsRow As Long, ResultsCol As Long)
  
  'deze routine weeft twee tabellen (met verspringende x-waarden) ineen
  'gaat standaard uit van blokinterpolatie en als voorgaande waarden ontbreken 0
  Dim Table1() As Variant
  Dim Table2() As Variant
  
  'zorg dat beide ranges geen lege cellen in de eerste kolom bevatten
  'Set myTable1 = TRUNCATERANGEBYEMPTYROWS(myTable1) ROUTINE BEVAT FOUT
  'Set myTable2 = TRUNCATERANGEBYEMPTYROWS(myTable2)
  
  'no 2D-array because the first dimension cannot be resized with redim preserve
  Dim Table3 As Variant
  
  Dim maxRows As Long
  Dim row As Long, col As Long
  Dim i1 As Long, i2 As Long, i3 As Long
  Dim Table1Done As Boolean, Table2Done As Boolean, Done As Boolean
  Dim LastVal1 As Variant, LastVal2 As Variant
  Dim NextVal1 As Variant, NextVal2 As Variant
  
  Table1 = myTable1
  Table2 = myTable2
  LastVal1 = -9999
  LastVal2 = -9999
  NextVal1 = -9999
  NextVal2 = -9999
  
  maxRows = UBound(Table1, 1) + UBound(Table2, 1)
  ReDim Table3(1 To maxRows, 1 To 3)
  
  If Table1(1, 1) <> Table2(1, 1) Then
    MsgBox ("Error: beide tabellen moeten starten met dezelfde x-waarde")
    End
  End If

  i1 = 1
  i2 = 1
  i3 = 1
  
  Table3(i3, 1) = Table1(i1, 1)
  Table3(i3, 2) = Table1(i1, 2)
  Table3(i3, 3) = Table2(i2, 2)
  
  'nu de rest
  While Not (Table1Done And Table2Done)
    
    'If i3 = 159 Then Stop
    
    If i1 >= UBound(Table1, 1) Then Table1Done = True
    If i2 >= UBound(Table2, 1) Then Table2Done = True
    
    If Table1Done And Table2Done Then
      'do nothing
    ElseIf Table1Done And Not Table2Done And i2 < UBound(Table2, 1) Then
      'finish table 2
      i2 = i2 + 1
      i3 = i3 + 1
      
      Table3(i3, 1) = Table2(i2, 1)
      Table3(i3, 2) = Table1(i1, 2)
      Table3(i3, 3) = Table2(i2, 2)
    ElseIf Table2Done And Not Table1Done And i1 < UBound(Table1, 1) Then
      'finish table1
      i1 = i1 + 1
      i3 = i3 + 1
      Table3(i3, 1) = Table1(i1, 1)
      Table3(i3, 2) = Table1(i1, 2)
      Table3(i3, 3) = Table2(i2, 2)
    ElseIf i1 < UBound(Table1, 1) And i2 < UBound(Table2, 1) Then
      NextVal1 = Table1(i1 + 1, 1)
      NextVal2 = Table2(i2 + 1, 1)
      
      If NextVal1 < NextVal2 Then
        'move one up in table 1
        i1 = i1 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table1(i1, 1)
        Table3(i3, 2) = Table1(i1, 2)
        Table3(i3, 3) = Table2(i2, 2) 'de vorige waarde uit tabel 2 is nog altijd van toepassing
      ElseIf NextVal2 < NextVal1 Then
        'move one up in table 2
        i2 = i2 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table2(i2, 1)
        Table3(i3, 2) = Table1(i1, 2) 'de vorige waarde uit tabel 1 is nog altijd van toepassing
        Table3(i3, 3) = Table2(i2, 2)
      ElseIf NextVal1 = NextVal2 Then
        'move one up in both tables
        i1 = i1 + 1
        i2 = i2 + 1
        i3 = i3 + 1
        Table3(i3, 1) = Table1(i1, 1)
        Table3(i3, 2) = Table1(i1, 2)
        Table3(i3, 3) = Table2(i2, 2)
      End If
    End If
        
  Wend
  
  'ReDim Preserve Table3(1 To i3, 1 To 3)
  
  'write the woven table to the worksheet
  row = ResultsRow
  col = ResultsCol
  ActiveSheet.Cells(row, col) = "X"
  ActiveSheet.Cells(row, col + 1) = "YTable1"
  ActiveSheet.Cells(row, col + 2) = "YTable2"
  
  row = row + 1
  
  Call PrintArray(Table3, ActiveSheet.Range(Cells(row, col), Cells(row, col)))
  
  Exit Sub
End Sub

Public Function TRUNCATERANGEBYEMPTYROWS(ByRef MyRange As Range) As Range
  Dim startRow As Long, endRow As Long
  Dim r As Long, i As Long
  
  For i = 1 To MyRange.Rows.Count
    If MyRange.Cells(i, 1) <> "" Then
      startRow = i
      Exit For
    End If
  Next
  
  For i = MyRange.Rows.Count To 1 Step -1
    If MyRange.Cells(i, 1) <> "" Then
      endRow = i
      Exit For
    End If
  Next
  
  Set TRUNCATERANGEBYEMPTYROWS = MyRange.Range(MyRange.Cells(startRow, 1), MyRange.Cells(endRow, MyRange.Columns.Count))
  
End Function

Public Function GETIJDEN_SINUS(amplitude As Variant, Periode As Variant, TijdstipNul As Variant, Evenwichtswaterstand As Variant, DatumTijd As Variant) As Variant
    GETIJDEN_SINUS = amplitude / 2 * Sin(2 * 3.1415 / Periode * (DatumTijd - TijdstipNul)) + Evenwichtswaterstand
End Function

Public Function YZ2TABULATED(YRange As Range, ZRange As Range, ResultsRow As Integer, ResultsCol As Integer, CheckAscending As Boolean) As Boolean
    Dim i As Integer, j As Integer
    Dim r As Integer
    Dim MinZ As Double
    Dim MinIdx As Integer
    
    Dim Level As Double
    Dim SmallestWidth As Double  'used for marking sudden jumps in width
    Dim LargestWidth As Double  'used for marking sudden jumps in width
    Dim PrevZ As Double
    Dim NextZ As Double
    Dim PrevY As Double
    Dim NextY As Double
    Dim ysec As Double
    r = ResultsRow
    Dim lastRow As Integer
    Dim n As Integer
        
    ActiveSheet.Cells(r, ResultsCol) = "Elevation"
    ActiveSheet.Cells(r, ResultsCol + 1) = "Width"
            
    Dim ZValues() As Double
    Dim WValues() As Double
                        
    'check if ascending Y
    If CheckAscending Then
        For i = 1 To YRange.Rows.Count - 1
            If Not YRange.Cells(i + 1, 1) = "" Then
                If YRange.Cells(i + 1, 1) < YRange.Cells(i, 1) Then
                    YZ2TABULATED = False
                    r = r + 1
                    ActiveSheet.Cells(r, ResultsCol) = "Error: non-ascending Y-values"
                    Exit Function
                End If
            End If
        Next
    End If
    
    'find the last row
    For i = 1 To YRange.Rows.Count
        If YRange.Cells(i, 1) = "" Or ZRange.Cells(i, 1) = "" Then
            lastRow = i - 1
            Exit For
        End If
    Next
    If lastRow = 0 Then lastRow = YRange.Rows.Count
                                     
    For i = 1 To lastRow
        
        Level = ZRange.Cells(i, 1)
        SmallestWidth = 0
        LargestWidth = 0
        
                
        'for the current elevation, walk through the entire profile and calculate both the smallest width and the largest width
        For j = 1 To lastRow - 1
            PrevZ = ZRange.Cells(j, 1)
            NextZ = ZRange.Cells(j + 1, 1)
            PrevY = YRange.Cells(j, 1)
            NextY = YRange.Cells(j + 1, 1)
            
            'If Level = 41.27 And NextZ = 41.27 Then Stop
            
            If (PrevZ < Level And NextZ <= Level) Or (PrevZ <= Level And NextZ < Level) Then
                'this entire section is below the current level,so add it entirely
                SmallestWidth = SmallestWidth + (NextY - PrevY)
                LargestWidth = LargestWidth + (NextY - PrevY)
            ElseIf PrevZ = Level And NextZ = Level Then
                'this is a flat section at exactly our current elevation. So only add it to the largestWiddth
                LargestWidth = LargestWidth + (NextY - PrevY)
            ElseIf PrevZ > Level And NextZ > Level Then
                'this entire section is above current level so don't add it at all
            ElseIf PrevZ < Level And NextZ > Level Then
                'only the left part of this section is below current Z, so add that part only. Find the intersection point by interpolation
                ysec = Interpolate(PrevZ, PrevY, NextZ, NextY, Level)
                SmallestWidth = SmallestWidth + (ysec - PrevY)
                LargestWidth = LargestWidth + (ysec - PrevY)
            ElseIf PrevZ > Level And NextZ < Level Then
                'only the right part of this section is below current Z, so add that part only
                ysec = Interpolate(PrevZ, PrevY, NextZ, NextY, Level)
                SmallestWidth = SmallestWidth + (NextY - ysec)
                LargestWidth = LargestWidth + (NextY - ysec)
            End If
        Next
            
        If LargestWidth > SmallestWidth Then n = n + 2 Else n = n + 1
        ReDim Preserve ZValues(1 To n)
        ReDim Preserve WValues(1 To n)
        If LargestWidth > SmallestWidth Then
            ZValues(n - 1) = Level
            WValues(n - 1) = SmallestWidth
            ZValues(n) = Level + 0.00001          'we have to in order to support horizontal floodplains
            WValues(n) = LargestWidth
        Else
            ZValues(n) = Level
            WValues(n) = LargestWidth
        End If
        
    Next
    
    'create a sort index based on the array with Z values
    Dim SortIdx() As Long
    SortIdx = HeapSort(ZValues)
            
    'and write the results to worksheet
    Dim idx As Integer
    
    'here we make sure never to write the same elevation twice. So we'll always write the last instance (= max with) for the given elevation
    'the first record
    r = r + 1
    i = 1
    idx = SortIdx(i)
    ActiveSheet.Cells(r, ResultsCol) = ZValues(idx)
    ActiveSheet.Cells(r, ResultsCol + 1) = WValues(idx)
    
    For i = 2 To UBound(SortIdx)
        idx = SortIdx(i)
        'only if our Z-value exceeds the previous one, write this record to a new row. Otherwise, overwrite the existing record
        If ZValues(idx) > ActiveSheet.Cells(r, ResultsCol) Then r = r + 1
        ActiveSheet.Cells(r, ResultsCol) = ZValues(idx)
        ActiveSheet.Cells(r, ResultsCol + 1) = WValues(idx)
    Next
            
            
End Function


Public Function YZ2TABULATED_ARCH_BRIDGE(YRange As Range, ZRange As Range, ResultsRow As Integer, ResultsCol As Integer) As Boolean
    'nog doen: de hoogte-breedte-tabel sorteren alvorens wegschrijven.
    'let op: deze functie is specifiek geschreven voor boogbruggen waarbij de breedte hogerop groter kan zijn dan lager
    Dim i As Integer, j As Integer, k As Integer
    Dim r As Integer
    Dim MinZ As Double
    Dim MinIdx As Integer
    
    Dim CurZ As Double
    Dim CurY As Double
    Dim CurW As Double
    Dim PrevZ As Double
    Dim NextZ As Double
    Dim PrevY As Double
    Dim NextY As Double
    Dim ysec As Double
    r = ResultsRow
    
    ActiveSheet.Cells(r, ResultsCol) = "Elevation"
    ActiveSheet.Cells(r, ResultsCol + 1) = "Width"
    
    Dim ZValues() As Double
    Dim WValues() As Double
                       
    For i = 1 To ZRange.Rows.Count
        If Not ZRange.Cells(i, 1) = "" And Not YRange.Cells(i, 1) = "" Then
            CurZ = ZRange.Cells(i, 1)
            CurW = 0
            CurY = YRange.Cells(i, 1)
            
            'walk from the left side until we cross the current z-value
            For k = 1 To i - 1
                If Not ZRange.Cells(k, 1) = "" And Not YRange.Cells(k, 1) = "" Then
                    If ZRange.Cells(k, 1) >= CurZ Then
                        'it looks like our starting point is already higher than our current Z. This means we can end here immediately
                        CurW = CurW + (CurY - YRange.Cells(k, 1))
                        Exit For
                    ElseIf ZRange.Cells(k + 1, 1) >= CurZ Then
                        'we found a cross section! calculate the y-value of the crossing
                        ysec = Interpolate(ZRange.Cells(k, 1), YRange.Cells(k, 1), ZRange.Cells(k + 1, 1), YRange.Cells(k + 1, 1), CurZ)
                        CurW = CurW + (CurY - ysec)
                        Exit For
                    End If
                End If
            Next
            
            'now walk from the right side until we cross the current z value
            For j = ZRange.Rows.Count To (i + 1) Step -1
                If Not ZRange.Cells(j, 1) = "" And Not YRange.Cells(j, 1) = "" Then
                    If ZRange.Cells(j, 1) >= CurZ Then
                        'it looks like our endpoint is already higher than our current Z. This means we can end here immediately
                        CurW = CurW + (YRange.Cells(j, 1) - CurY)
                        Exit For
                    ElseIf ZRange.Cells(j - 1, 1) >= CurZ Then
                        'we found a cross section! calculate the y-value of the crossing
                        ysec = Interpolate(ZRange.Cells(j, 1), YRange.Cells(j, 1), ZRange.Cells(j - 1, 1), YRange.Cells(j - 1, 1), CurZ)
                        CurW = CurW + (ysec - CurY)
                        Exit For
                    End If
                End If
            Next
                        
            ReDim Preserve ZValues(1 To i)
            ReDim Preserve WValues(1 To i)
            ZValues(i) = CurZ
            WValues(i) = CurW
            
        End If
    Next
            
    'create a sort index based on the array with Z values
    Dim SortIdx() As Long
    SortIdx = HeapSort(ZValues)
            
    'and write the results to worksheet
    Dim idx As Integer
    For i = 1 To UBound(SortIdx)
        r = r + 1
        idx = SortIdx(i)
        ActiveSheet.Cells(r, ResultsCol) = ZValues(idx)
        ActiveSheet.Cells(r, ResultsCol + 1) = WValues(idx)
    Next
    
    'since we're dealing with an arch bridge we can add 0 width juuuuust above the highest value
    r = r + 1
    ActiveSheet.Cells(r, ResultsCol) = ZValues(SortIdx(UBound(SortIdx))) + 0.01
    ActiveSheet.Cells(r, ResultsCol + 1) = 0
            
        
End Function

Public Function ExitLossCoef(ExitLossUserCoef As Double, Astruc As Double, AchanDn As Double) As Double
    ExitLossCoef = ExitLossUserCoef * (1 - Astruc / AchanDn) ^ 2
End Function

Public Sub ExtraResistance(UpstreamTabulatedProfileRange As Range, TabulatedStructureRange As Range, DownstreamTabulatedProfileRange As Range, Length As Variant, EntranceLossUserCoef As Variant, ExitLossUserCoef As Variant, ChannelFrictionManning As Variant, StructureFrictionManning As Variant, ResultsRow As Integer, ResultsCol As Integer)
    'this function calculates the parameters for an extra resistance node in SOBEK, based on local cross section, bridge cross section and downstream cross section

    'first we need to create a list of unique elevations from all ranges and sort it in ascending order
    Dim Elevations As Collection
    Set Elevations = New Collection
    Dim i As Integer, r As Integer
    
    Dim MinElev As Variant
    MinElev = 9E+99
    Dim MaxElev As Variant
    MaxElev = -9E+99
    
    'find the minimum and maximum elevation
    For r = 1 To UpstreamTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, UpstreamTabulatedProfileRange.Cells(r, 1)) = False Then
            If UpstreamTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = UpstreamTabulatedProfileRange.Cells(r, 1)
            If UpstreamTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = UpstreamTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    For r = 1 To TabulatedStructureRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, TabulatedStructureRange.Cells(r, 1)) = False Then
            If TabulatedStructureRange.Cells(r, 1) < MinElev Then MinElev = TabulatedStructureRange.Cells(r, 1)
            If TabulatedStructureRange.Cells(r, 1) > MaxElev Then MaxElev = TabulatedStructureRange.Cells(r, 1)
        End If
    Next
    
    For r = 1 To DownstreamTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, DownstreamTabulatedProfileRange.Cells(r, 1)) = False Then
            If DownstreamTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = DownstreamTabulatedProfileRange.Cells(r, 1)
            If DownstreamTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = DownstreamTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    'now we create a new collection of elevations, with an increment of 1 cm
    MinElev = Math.Round(MinElev, 2)
    MaxElev = Math.Round(MaxElev, 2)
    Dim Elev As Variant
    For Elev = MinElev - 0.01 To MaxElev + 0.01 Step 0.01
        Call Elevations.Add(Elev)
    Next
        
    'for each elevation we need to interpolate the wetted areas and wetted perimeters from our profile tables
    Dim AchanUp As Collection
    Dim PchanUp As Collection
    Dim RchanUp As Collection
    Dim Astruc As Collection
    Dim Pstruc As Collection
    Dim Rstruc As Collection
    Dim AchanDn As Collection
    Dim PchanDn As Collection
    Dim RchanDn As Collection
    
    'calculate the wetted perimeters (P), wetted areas (A) and hydraulic radiuses (R) for all elevations in our collection
    Set AchanUp = WettedAreasFromTabulatedProfileRange(UpstreamTabulatedProfileRange, Elevations, False)
    Set PchanUp = WettedPerimetersFromTabulatedProfileRange(UpstreamTabulatedProfileRange, Elevations, False)
    Set RchanUp = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanUp, PchanUp)
    Set Astruc = WettedAreasFromTabulatedProfileRange(TabulatedStructureRange, Elevations, False)
    Set Pstruc = WettedPerimetersFromTabulatedProfileRange(TabulatedStructureRange, Elevations, False)
    Set Rstruc = HydraulicRadiusFromWettedAreasAndWettedPerimeters(Astruc, Pstruc)
    Set AchanDn = WettedAreasFromTabulatedProfileRange(DownstreamTabulatedProfileRange, Elevations, False)
    Set PchanDn = WettedPerimetersFromTabulatedProfileRange(DownstreamTabulatedProfileRange, Elevations, False)
    Set RchanDn = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanDn, PchanDn)
        
    'we start with the friction loss coefficient for the structure
    Dim StructureFriction As Collection
    Set StructureFriction = FrictionLossCoefFromHydraulicRadiuses(Rstruc, Length, StructureFrictionManning)
        
    'now also the friction loss coefficient for the hypothetical situation where the structure would be absent
    Dim ChannelFriction As Collection
    Set ChannelFriction = FrictionLossCoefFromHydraulicRadiuses(RchanUp, Length, ChannelFrictionManning)
    
    'compute the added friction loss as a result of the installed bridge.
    'note that this value can be negative
    Dim FrictionLoss As Collection
    Set FrictionLoss = New Collection
    For i = 1 To Elevations.Count
        Call FrictionLoss.Add(StructureFriction(i) - ChannelFriction(i))
    Next
    
    'also write the entrance loss coeficient
    Dim EntranceLoss As Collection
    Set EntranceLoss = New Collection
    For i = 1 To Elevations.Count
        'only if the structure has a smaller wetted area than the upstream channel, the entrance loss applies!
        If Astruc(i) < AchanUp(i) Then
            Call EntranceLoss.Add(EntranceLossUserCoef)
        Else
            Call EntranceLoss.Add(0)
        End If
    Next
    
    'and finally compute the exit loss coefficient
    Dim ExitLoss As Collection
    Set ExitLoss = New Collection
    Dim ExitLossCoefficient As Variant
    For i = 1 To Elevations.Count
        If AchanDn(i) > 0 And AchanDn(i) > AchanUp(i) Then
            Dim ExLs As Double
            Call ExitLoss.Add(ExitLossCoef(ExitLossUserCoef, Astruc(i), AchanDn(i)))
        Else
            Call ExitLoss.Add(0)
        End If
    Next
    
    'now that we have our loss coefficients we can compute the added value due to our structure.
    Dim ksiStructure As Collection
    Set ksiStructure = New Collection
    For i = 1 To Elevations.Count
        If Astruc(i) > 0 Then
            Call ksiStructure.Add(Maximum(0, (EntranceLoss(i) + FrictionLoss(i) + ExitLoss(i)) / (2 * 9.81 * Astruc(i) ^ 2)))
        Else
            Call ksiStructure.Add(0)
        End If
    Next
    
    ActiveSheet.Cells(ResultsRow, ResultsCol) = "Waterhoogte (m + NAP)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "A instroomzijde (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 2) = "P instroomzijde (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 3) = "R instroomzijde (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 4) = "A brug (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 5) = "P brug (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 6) = "R brug (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 7) = "A uitstroomzijde (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 8) = "P uitstroomzijde (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 9) = "R uitstroomzijde (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 10) = "Ruwheidscoef beek (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 11) = "Ruwheidscoef brug (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 12) = "Intreecoef (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 13) = "Uittreecoef (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 14) = "Ksi (-)"
    
    For i = 1 To Elevations.Count
        ActiveSheet.Cells(ResultsRow + i, ResultsCol) = Elevations(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 1) = AchanUp(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 2) = PchanUp(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 3) = RchanUp(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 4) = Astruc(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 5) = Pstruc(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 6) = Rstruc(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 7) = AchanDn(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 8) = PchanDn(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 9) = RchanDn(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 10) = ChannelFriction(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 11) = StructureFriction(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 12) = EntranceLoss(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 13) = ExitLoss(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 14) = ksiStructure(i)
    Next
    
    
    
End Sub

Public Function WettedPerimeterFromYZProfile(YZProfileRange As Range, waterLevel As Double) As Double
    
    'this function calculates the wetted perimeter for a given waterlevel in an YZ type cross section
    Dim myPerimeter As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim Z1 As Double
    Dim Z2 As Double
    Dim y As Double
    Dim i As Integer
    
    For i = 1 To YZProfileRange.Rows.Count - 1
        Y1 = YZProfileRange.Cells(i, 1)
        Y2 = YZProfileRange.Cells(i + 1, 1)
        Z1 = YZProfileRange.Cells(i, 2)
        Z2 = YZProfileRange.Cells(i + 1, 2)
        If waterLevel < Z1 And waterLevel < Z2 Then
            'do nothing
        ElseIf waterLevel >= Z1 And waterLevel >= Z2 Then
            'this entire section is underwater add the diagonal to our perimeter
            myPerimeter = myPerimeter + PYTHAGORAS(Z2 - Z2, Y1 - Y2)
        ElseIf waterLevel >= Z1 And waterLevel < Z2 Then
            'only the left section is under water. Find the Y-value where this section emerges from the water
            y = Interpolate(Z1, Y1, Z2, Y2, waterLevel)
            myPerimeter = myPerimeter + PYTHAGORAS(Z1 - waterLevel, Y1 - y)
        ElseIf waterLevel < Z1 And waterLevel >= Z2 Then
            'only the right section is under water. Find the Y-value where this section emerges from the water
            y = Interpolate(Z1, Y1, Z2, Y2, waterLevel)
            myPerimeter = myPerimeter + PYTHAGORAS(Z2 - waterLevel, Y2 - y)
        End If
    Next
    WettedPerimeterFromYZProfile = myPerimeter
End Function


Public Function WettedAreaFromYZProfile(YZProfileRange As Range, waterLevel As Double) As Double
    
    'this function calculates the wetted area for a given waterlevel in an YZ type cross section
    Dim myArea As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim Z1 As Double
    Dim Z2 As Double
    Dim Zmax As Double
    Dim Zmin As Double
    Dim y As Double
    Dim i As Integer
    
    For i = 1 To YZProfileRange.Rows.Count - 1
        Y1 = YZProfileRange.Cells(i, 1)
        Y2 = YZProfileRange.Cells(i + 1, 1)
        Z1 = YZProfileRange.Cells(i, 2)
        Z2 = YZProfileRange.Cells(i + 1, 2)
        Zmax = Application.WorksheetFunction.max(Z1, Z2)
        Zmin = Application.WorksheetFunction.Min(Z1, Z2)
        If waterLevel < Z1 And waterLevel < Z2 Then
            'do nothing
        ElseIf waterLevel >= Z1 And waterLevel >= Z2 Then
            'this entire section is submerged so add the area
            myArea = myArea + (waterLevel - Zmax) * (Y2 - Y1) 'the rectangular part
            myArea = myArea + (Zmax - Zmin) * (Y2 - Y1) / 2   'the triangular part
        ElseIf waterLevel >= Z1 And waterLevel < Z2 Then
            'only the left section is under water. Find the Y-value where this section emerges from the water
            y = Interpolate(Z1, Y1, Z2, Y2, waterLevel)
            myArea = myArea + (waterLevel - Zmin) * (y - Y1) / 2
        ElseIf waterLevel < Z1 And waterLevel >= Z2 Then
            'only the right section is under water. Find the Y-value where this section emerges from the water
            y = Interpolate(Z1, Y1, Z2, Y2, waterLevel)
            myArea = myArea + (waterLevel - Zmin) * (Y2 - y) / 2
        End If
    Next
    WettedAreaFromYZProfile = myArea
End Function


Public Sub AddedFrictionLoss(OriginalTabulatedProfileRange As Range, NewTabulatedProfileRange As Range, Length As Variant, OriginalFrictionManning As Variant, NewFrictionManning As Variant, ResultsRow As Integer, ResultsCol As Integer)
    'first we need to create a list of unique elevations from all ranges and sort it in ascending order
    Dim Elevations As Collection
    Set Elevations = New Collection
    Dim i As Integer, r As Integer
    
    Dim MinElev As Variant
    MinElev = 9E+99
    Dim MaxElev As Variant
    MaxElev = -9E+99
    
    'find the minimum and maximum elevation
    For r = 1 To OriginalTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, OriginalTabulatedProfileRange.Cells(r, 1)) = False Then
            If OriginalTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = OriginalTabulatedProfileRange.Cells(r, 1)
            If OriginalTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = OriginalTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    For r = 1 To NewTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, NewTabulatedProfileRange.Cells(r, 1)) = False Then
            If NewTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = NewTabulatedProfileRange.Cells(r, 1)
            If NewTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = NewTabulatedProfileRange.Cells(r, 1)
        End If
    Next
        
    'now we create a new collection of elevations, with an increment of 1 cm
    MinElev = Math.Round(MinElev, 2)
    MaxElev = Math.Round(MaxElev, 2)
    Dim Elev As Variant
    For Elev = MinElev - 0.01 To MaxElev + 0.01 Step 0.01
        Call Elevations.Add(Elev)
    Next
        
    'for each elevation we need to interpolate the wetted areas and wetted perimeters from our profile tables
    Dim AchanOrig As Collection
    Dim PchanOrig As Collection
    Dim RchanOrig As Collection
    Dim AchanNew As Collection
    Dim PchanNew As Collection
    Dim RchanNew As Collection
    
    'calculate the wetted perimeters (P), wetted areas (A) and hydraulic radiuses (R) for all elevations in our collection
    Set AchanOrig = WettedAreasFromTabulatedProfileRange(OriginalTabulatedProfileRange, Elevations, False)
    Set PchanOrig = WettedPerimetersFromTabulatedProfileRange(OriginalTabulatedProfileRange, Elevations, False)
    Set RchanOrig = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanOrig, PchanOrig)
    Set AchanNew = WettedAreasFromTabulatedProfileRange(NewTabulatedProfileRange, Elevations, False)
    Set PchanNew = WettedPerimetersFromTabulatedProfileRange(NewTabulatedProfileRange, Elevations, False)
    Set RchanNew = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanNew, PchanNew)
        
    'friction loss for the original situation
    Dim OriginalFriction As Collection
    Set OriginalFriction = FrictionLossCoefFromHydraulicRadiuses(RchanOrig, Length, OriginalFrictionManning)
    
    'we start with the friction loss coefficient for the structure
    Dim NewFriction As Collection
    Set NewFriction = FrictionLossCoefFromHydraulicRadiuses(RchanNew, Length, NewFrictionManning)
        
    'compute the added friction loss as a result of the installed bridge.
    'note that this value can be negative
    Dim FrictionLoss As Collection
    Set FrictionLoss = New Collection
    For i = 1 To Elevations.Count
        Call FrictionLoss.Add(NewFriction(i) - OriginalFriction(i))
    Next
        
    ActiveSheet.Cells(ResultsRow, ResultsCol) = "Waterhoogte (m + NAP)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "R oorspronkelijk (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 2) = "R nieuw (m)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 3) = "Ruwheidscoef oorspronkelijk (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 4) = "Ruwheidscoef nieuw (-)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 5) = "Extra ruwheid (-)"
    
    For i = 1 To Elevations.Count
        ActiveSheet.Cells(ResultsRow + i, ResultsCol) = Elevations(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 1) = RchanOrig(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 2) = RchanNew(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 3) = OriginalFriction(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 4) = NewFriction(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 5) = FrictionLoss(i)
    Next
    
       
End Sub


Public Sub ExitLoss(NarrowTabulatedProfileRange As Range, WideTabulatedProfileRange As Range, ExitLossUserCoef As Variant, ResultsRow As Integer, ResultsCol As Integer)
    'this function calculates the parameters for an extra resistance node in SOBEK, based on local cross section, bridge cross section and downstream cross section

    'first we need to create a list of unique elevations from all ranges and sort it in ascending order
    Dim Elevations As Collection
    Set Elevations = New Collection
    Dim i As Integer, r As Integer
    
    Dim MinElev As Variant
    MinElev = 9E+99
    Dim MaxElev As Variant
    MaxElev = -9E+99
    
    'find the minimum and maximum elevation
    For r = 1 To NarrowTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, NarrowTabulatedProfileRange.Cells(r, 1)) = False Then
            If NarrowTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = NarrowTabulatedProfileRange.Cells(r, 1)
            If NarrowTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = NarrowTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    For r = 1 To WideTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, WideTabulatedProfileRange.Cells(r, 1)) = False Then
            If WideTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = WideTabulatedProfileRange.Cells(r, 1)
            If WideTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = WideTabulatedProfileRange.Cells(r, 1)
        End If
    Next
        
    'now we create a new collection of elevations, with an increment of 1 cm
    MinElev = Math.Round(MinElev, 2)
    MaxElev = Math.Round(MaxElev, 2)
    Dim Elev As Variant
    For Elev = MinElev - 0.01 To MaxElev + 0.01 Step 0.01
        Call Elevations.Add(Elev)
    Next
        
    'for each elevation we need to interpolate the wetted areas and wetted perimeters from our profile tables
    Dim AchanNarrow As Collection
    Dim PchanNarrow As Collection
    Dim RchanNarrow As Collection
    Dim AchanWide As Collection
    Dim PchanWide As Collection
    Dim RchanWide As Collection
    
    'calculate the wetted perimeters (P), wetted areas (A) and hydraulic radiuses (R) for all elevations in our collection
    Set AchanNarrow = WettedAreasFromTabulatedProfileRange(NarrowTabulatedProfileRange, Elevations, False)
    Set PchanNarrow = WettedPerimetersFromTabulatedProfileRange(NarrowTabulatedProfileRange, Elevations, False)
    Set RchanNarrow = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanNarrow, PchanNarrow)
    Set AchanWide = WettedAreasFromTabulatedProfileRange(WideTabulatedProfileRange, Elevations, False)
    Set PchanWide = WettedPerimetersFromTabulatedProfileRange(WideTabulatedProfileRange, Elevations, False)
    Set RchanWide = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanWide, PchanWide)
                
    'and finally compute the exit loss coefficient
    Dim ExitLoss As Collection
    Set ExitLoss = New Collection
    Dim ExitLossCoefficient As Variant
    For i = 1 To Elevations.Count
        If AchanWide(i) > 0 And AchanWide(i) > AchanNarrow(i) Then
            Call ExitLoss.Add(ExitLossUserCoef * (1 - AchanNarrow(i) / AchanWide(i)) ^ 2)
        Else
            Call ExitLoss.Add(0)
        End If
    Next
        
    ActiveSheet.Cells(ResultsRow, ResultsCol) = "Waterhoogte (m + NAP)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "A nauw profiel (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 2) = "A wijd profiel (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 3) = "Uittreecoef (-)"
    
    For i = 1 To Elevations.Count
        ActiveSheet.Cells(ResultsRow + i, ResultsCol) = Elevations(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 1) = AchanWide(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 2) = AchanNarrow(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 3) = ExitLoss(i)
    Next
        
    
End Sub



Public Sub EntranceLoss(WideTabulatedProfileRange As Range, NarrowTabulatedProfileRange As Range, EntranceLossUserCoef As Variant, ResultsRow As Integer, ResultsCol As Integer)
    'this function calculates the entrance loss for a given narrowing of flow paths

    'first we need to create a list of unique elevations from all ranges and sort it in ascending order
    Dim Elevations As Collection
    Set Elevations = New Collection
    Dim i As Integer, r As Integer
    
    Dim MinElev As Variant
    MinElev = 9E+99
    Dim MaxElev As Variant
    MaxElev = -9E+99
    
    'find the minimum and maximum elevation
    For r = 1 To WideTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, WideTabulatedProfileRange.Cells(r, 1)) = False Then
            If WideTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = WideTabulatedProfileRange.Cells(r, 1)
            If WideTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = WideTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    For r = 1 To NarrowTabulatedProfileRange.Rows.Count
        If CollectionContainsNumericalValue(Elevations, NarrowTabulatedProfileRange.Cells(r, 1)) = False Then
            If NarrowTabulatedProfileRange.Cells(r, 1) < MinElev Then MinElev = NarrowTabulatedProfileRange.Cells(r, 1)
            If NarrowTabulatedProfileRange.Cells(r, 1) > MaxElev Then MaxElev = NarrowTabulatedProfileRange.Cells(r, 1)
        End If
    Next
    
    'now we create a new collection of elevations, with an increment of 1 cm
    MinElev = Math.Round(MinElev, 2)
    MaxElev = Math.Round(MaxElev, 2)
    Dim Elev As Variant
    For Elev = MinElev - 0.01 To MaxElev + 0.01 Step 0.01
        Call Elevations.Add(Elev)
    Next
        
    'for each elevation we need to interpolate the wetted areas and wetted perimeters from our profile tables
    Dim AchanWide As Collection
    Dim PchanWide As Collection
    Dim RchanWide As Collection
    Dim AchanNarrow As Collection
    Dim PchanNarrow As Collection
    Dim RchanNarrow As Collection
    
    'calculate the wetted perimeters (P), wetted areas (A) and hydraulic radiuses (R) for all elevations in our collection
    Set AchanWide = WettedAreasFromTabulatedProfileRange(WideTabulatedProfileRange, Elevations, False)
    Set PchanWide = WettedPerimetersFromTabulatedProfileRange(WideTabulatedProfileRange, Elevations, False)
    Set RchanWide = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanWide, PchanWide)
    Set AchanNarrow = WettedAreasFromTabulatedProfileRange(NarrowTabulatedProfileRange, Elevations, False)
    Set PchanNarrow = WettedPerimetersFromTabulatedProfileRange(NarrowTabulatedProfileRange, Elevations, False)
    Set RchanNarrow = HydraulicRadiusFromWettedAreasAndWettedPerimeters(AchanNarrow, PchanNarrow)
            
    'also write the entrance loss coeficient
    Dim EntranceLoss As Collection
    Set EntranceLoss = New Collection
    For i = 1 To Elevations.Count
        'only if the narrow profile has a smaller wetted area than the upstream channel, the entrance loss applies!
        If AchanNarrow(i) < AchanWide(i) Then
            Call EntranceLoss.Add(EntranceLossUserCoef)
        Else
            Call EntranceLoss.Add(0)
        End If
    Next
        
    ActiveSheet.Cells(ResultsRow, ResultsCol) = "Waterhoogte (m + NAP)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "A wijd profiel (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 2) = "A nauw profiel (m2)"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 3) = "Intreecoef (-)"
    
    For i = 1 To Elevations.Count
        ActiveSheet.Cells(ResultsRow + i, ResultsCol) = Elevations(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 1) = AchanWide(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 2) = AchanNarrow(i)
        ActiveSheet.Cells(ResultsRow + i, ResultsCol + 3) = EntranceLoss(i)
    Next
    
End Sub


Public Function HydraulicRadiusFromWettedAreasAndWettedPerimeters(a As Collection, P As Collection) As Collection
    Dim r As Collection
    Dim i As Integer
    Set r = New Collection
    For i = 1 To a.Count
        If P(i) > 0 Then
            Call r.Add(a(i) / P(i))
            If r(i) < 0 Then Stop
        Else
            Call r.Add(0)
        End If
    Next
    Set HydraulicRadiusFromWettedAreasAndWettedPerimeters = r
End Function

Public Function WettedAreasFromTabulatedProfileRange(MyRange As Range, Elevations As Collection, BlockInterpolation As Boolean) As Collection
    'this function calculates the wetted area for a given set of elevations from a tabulated profile (elevation-width)
    'notice that the elevations must be sorted!
    Dim WettedAreas As Collection
    Set WettedAreas = New Collection
        
    Dim i As Integer, CurW As Variant, prevW As Variant, curElev As Variant, prevElev As Variant, curA As Variant, prevA As Variant
    
    curElev = Elevations(1)
    CurW = InterpolateFromRange(curElev, MyRange, 1, 2, False, True, BlockInterpolation, False)
    curA = 0
    Call WettedAreas.Add(curA) 'the lowest elevation has no wetted area by definition
    For i = 2 To Elevations.Count
        prevW = CurW
        prevElev = curElev
        prevA = curA
                
        curElev = Elevations(i)
        CurW = InterpolateFromRange(curElev, MyRange, 1, 2, False, True, BlockInterpolation, False)
        
        If BlockInterpolation Then
            curA = prevA + (curElev - prevElev) * prevW 'since we're assuming block interpolation for this profile, the previous width is leading
        Else
            curA = prevA + (curElev - prevElev) * ((CurW + prevW) / 2)   'area of a trapezium is the average of min/max width * height
        End If
        
        Call WettedAreas.Add(curA)
    Next
    
    'return the WettedAreas collection
    Set WettedAreasFromTabulatedProfileRange = WettedAreas
    
End Function

Public Function WettedPerimetersFromTabulatedProfileRange(MyRange As Range, Elevations As Collection, BlockInterpolation As Boolean) As Collection
    'this function calculates the wetted perimeters for a given set of elevations from a tabulated profile (elevation-width)
    'notice that the elevations must be sorted in ascending order
    Dim WettedPerimeters As Collection
    Set WettedPerimeters = New Collection
        
    Dim i As Integer, CurW As Variant, prevW As Variant, curElev As Variant, prevElev As Variant, prevP As Variant, curP As Variant
    
    curElev = Elevations(1)
    CurW = InterpolateFromRange(curElev, MyRange, 1, 2, False, True, BlockInterpolation, False)
    curP = CurW
    Call WettedPerimeters.Add(curP)
    
    For i = 2 To Elevations.Count
        prevW = CurW
        prevElev = curElev
        prevP = curP
                    
        curElev = Math.Round(Elevations(i), 2)
        CurW = InterpolateFromRange(curElev, MyRange, 1, 2, False, True, BlockInterpolation, False)
        If CurW = 0 And prevW = 0 Then
            curP = prevP
        ElseIf prevW = 0 Then
            curP = curP + Math.Abs(CurW - prevW)
        ElseIf CurW = 0 Then
            curP = curP + Math.Abs(CurW - prevW)
        Else
            If BlockInterpolation Then
                curP = curP + 2 * (curElev - prevElev) + Math.Abs(CurW - prevW)
            Else
                curP = prevP + 2 * PYTHAGORAS((curElev - prevElev), (CurW - prevW) / 2)  'we assume a symmetrical profile so apply pythagoras twice (for both sides half of the added width)
            End If
        End If
        Call WettedPerimeters.Add(curP)
    Next
    
    
    'return the WettedAreas collection
    Set WettedPerimetersFromTabulatedProfileRange = WettedPerimeters
    
End Function

Public Function FrictionLossCoefFromHydraulicRadiuses(HydraulicRadiuses As Collection, Length As Variant, FrictionManning As Variant) As Collection
    'this function calculates the added friction loss coefficient for a collection of hydraulic radiuses of both a structure and the channel cross section
    'IMPORTANT: both collections must have the same size number of records!
    'the formula for friction loss yields: 2gL/(C^2*R)
    
    Dim chezy As Variant
    Dim r As Variant
    Dim F As Variant
    Dim i As Integer
        
    Dim FrictionValues As Collection
    Set FrictionValues = New Collection
    
    For i = 1 To HydraulicRadiuses.Count
    
        'calculate the friction loss coefficient
        chezy = Manning2Chezy(FrictionManning, HydraulicRadiuses(i))
        If HydraulicRadiuses(i) > 0 And chezy > 0 Then
            F = 2 * 9.81 * Length / (chezy ^ 2 * HydraulicRadiuses(i))
        Else
            F = 0
        End If
                
        'calculate the added friction coef for the current record
        Call FrictionValues.Add(F)
    Next
    
    Set FrictionLossCoefFromHydraulicRadiuses = FrictionValues

End Function


Public Function CollectionContainsNumericalValue(ByRef myCollection As Collection, SearchValue As Variant) As Boolean
    Dim i As Integer, result As Boolean
    For i = 1 To myCollection.Count
        If myCollection(i) = SearchValue Then
            result = True
            Exit For
        End If
    Next
    CollectionContainsNumericalValue = result
End Function


Public Function QHEVEL(Diameter As Variant, Lengte As Variant, chezy As Variant, muin As Variant, muUit As Variant, muBuig As Variant, dH As Variant) As Variant

Dim a As Variant
Dim P As Variant
Dim Friction As Variant
Dim mu As Variant

a = pi * (Diameter / 2) ^ 2
P = 2 * pi * (Diameter / 2)

Friction = (2 * 9.81 * Lengte) / (chezy ^ 2 * a / P)
mu = 1 / (Sqr(muin + muUit + Friction + muBuig))

QHEVEL = mu * a * Sqr(2 * 9.81 * dH)

End Function

Public Function QDUIKER(Diameter As Variant, Lengte As Variant, chezy As Variant, muin As Variant, muUit As Variant, dH As Variant) As Variant

Dim a As Variant
Dim P As Variant
Dim Friction As Variant
Dim mu As Variant

a = pi * (Diameter / 2) ^ 2
P = 2 * pi * (Diameter / 2)

Friction = (2 * 9.81 * Lengte) / (chezy ^ 2 * a / P)
mu = 1 / (Sqr(muin + muUit + Friction))

QDUIKER = mu * a * Sqr(2 * 9.81 * dH)

End Function

Public Function QDUIKERRECHTHOEK(BOB As Variant, Breedte As Variant, Hoogte As Variant, Lengte As Variant, chezy As Variant, muin As Variant, muUit As Variant, h1 As Variant, h2 As Variant) As Variant

Dim a As Variant
Dim P As Variant
Dim Friction As Variant
Dim mu As Variant


If h1 >= Hoogte + BOB Then
  'geheel gevuld
  a = Breedte * Hoogte
  P = Breedte * 2 + Hoogte * 2
Else
  'gedeeltelijk gevuld
  a = Breedte * (h1 - BOB)
  P = Breedte + (h1 - BOB) * 2
End If

Friction = (2 * 9.81 * Lengte) / (chezy ^ 2 * a / P)
mu = 1 / (Sqr(muin + muUit + Friction))

QDUIKERRECHTHOEK = mu * a * Sqr(2 * 9.81 * (h1 - h2))

End Function


Public Function QORIFICE(Z As Variant, W As Variant, gh As Variant, mu As Variant, cw As Variant, h1 As Variant, h2 As Variant) As Variant
'Z = crest level
'W = width
'gh = gate height (openningshoogte)
'mu = contraction coef (standaard 0.63)
'cw = lateral contraction coef
'h1 = waterstand bovenstrooms
'h2 = waterstand benedenstrooms
'ce = afvoercoefficient. standaard 1.5

Dim Af As Variant
Dim Ce As Variant
Dim g As Variant
Dim u As Variant 'stroomsnelheid over de kruin. Moet eigenlijk iteratief worden bepaald maar ik zet hem even op 1
u = 1
Ce = 1.5
g = 9.81

'bepaal of hij verdronken of vrij is
If (h1 - Z) >= (3 / 2 * gh) Then   'orifice flow
  If h2 <= (Z + gh) Then 'free orifice flow
    Af = W * mu * gh
    QORIFICE = cw * W * mu * gh * VBA.Sqr(2 * g * (h1 - (Z + mu * gh)))
  ElseIf h2 > (Z + gh) Then 'submerged orifice flow
    Af = W * mu * gh
    QORIFICE = cw * W * mu * gh * VBA.Sqr(2 * g * (h1 - h2))
  End If
ElseIf (h1 - Z) < (3 / 2 * gh) Then 'weir flow
  If (h1 - Z) > (3 / 2 * (h2 - Z)) Then 'free weir flow
    Af = W * 2 / 3 * (h1 - Z)
    QORIFICE = cw * W * 2 / 3 * VBA.Sqr(2 / 3 * g * (h1 - Z) ^ 3 / 2)
  ElseIf (h1 - Z) <= (3 / 2 * (h2 - Z)) Then 'submerged weir flow
    Af = W * (h1 - Z - u ^ 2 / (2 * g))
    QORIFICE = Ce * cw * W * (h1 - Z - (u ^ 2 / (2 * g))) * VBA.Sqr(2 * g * (h1 - h2))
  End If
Else
  MsgBox ("Error: kon niet bepalen of orifice verdronken of vrij was.")
End If


End Function

Public Function QSTUW(Breedte As Variant, DischCoef As Variant, h1 As Variant, h2 As Variant, Z As Variant, Optional LatContrCoef As Variant = 1) As Variant
  Dim Hup As Variant, Hdown As Variant, Multiplier As Variant

'Free flow: als h2 - z < 2/3 * (h1 -z)
If h1 >= h2 Then
  Hup = h1
  Hdown = h2
  Multiplier = 1
Else
  Hup = h2
  Hdown = h1
  Multiplier = -1
End If

If Hup <= Z Then
  QSTUW = 0
ElseIf Hdown < Z Or (Hdown - Z) < 2 / 3 * (Hup - Z) Then
  'Free flow: Q = c * B * 2/3 * SQRT(2/3 * g) * (h1 - z)^1.5
  QSTUW = Multiplier * DischCoef * LatContrCoef * Breedte * 2 / 3 * Sqr(2 / 3 * 9.81) * (Hup - Z) ^ 1.5
Else
  'Drowned flow: Q = c * B * (h2 -z) * SQRT(2 * g *(h1 - h2))
  QSTUW = Multiplier * DischCoef * LatContrCoef * Breedte * (Hdown - Z) * Sqr(2 * 9.81 * (Hup - Hdown))
End If

End Function

Public Function QREHBOCKWEIR(waterLevel As Double, CrestLevel As Double, BedLevel As Double, CrestWidth As Double) As Double
    'see handboek debietmeting https://edepot.wur.nl/216512
    Dim Q As Double
    Dim Ce As Double
    Dim P As Double
    Dim h1 As Double
    Dim he As Double
    
    h1 = waterLevel - CrestLevel    'overstorthoogte
    P = CrestLevel - BedLevel       'apexhoogte (hoogte overstortrand boven bovenstroomse bodem)
    he = h1 + 0.0012                'effectieve overstortende straal
           
    Ce = 0.602 + 0.083 * h1 / P     'NB: is door Boiten zelf ter discussie gesteld. Niet meer gebruiken?
    
    If waterLevel > CrestLevel Then
        QREHBOCKWEIR = Ce * 2 / 3 * Math.Sqr(2 * 9.81) * CrestWidth * he ^ (3 / 2)
    Else
        QREHBOCKWEIR = 0
    End If
            
End Function

Public Function QUNIVERSALWEIRTRIANGULARSECTION(h1 As Variant, h2 As Variant, Y1 As Double, Y2 As Double, Z1 As Double, Z2 As Double, Ce As Double) As Double
    'this function returns the flow over a triangular section of a universal weir
    Dim zCrest As Double 'the lowest of both Z values represents the crest elevation
    Dim dz As Double
    zCrest = Application.WorksheetFunction.Min(Z1, Z2)
    dz = Math.Abs(Z1 - Z2)                              'delta z for this section represents the difference between z-values over this triangle
    Dim u As Double
    Dim W As Double, a As Double
    Dim g As Double
    Dim mlTriangular As Double
    
    g = 9.81
    W = Math.Abs(Y2 - Y1)
    
    'calculate ml
    If (h1 - zCrest) <= 1.25 * dz Then
        mlTriangular = 4 / 5
    Else
        mlTriangular = (2 / 3) + (1 / 6) * (dz / (h1 - zCrest)) '2/3 + een zesde van de verhouding tussen dz en de overstortende straal
    End If
    
    'decide whether it's free flow or submerged flow
    If (h1 <= zCrest) Then
        'no flow over this section
        QUNIVERSALWEIRTRIANGULARSECTION = 0
    ElseIf (h2 - zCrest) / (h1 - zCrest) < mlTriangular Then
        'free flow
        u = Ce * Math.Sqr(2 * g * (1 - mlTriangular) * (h1 - zCrest))
        If (h1 - zCrest) / (1 / mlTriangular) <= dz Then
            a = W * ((mlTriangular * (h1 - zCrest)) ^ 2) / (2 * dz)
        Else
            a = W * ((h1 - zCrest) / (1 / mlTriangular) - dz / 2)
        End If
        QUNIVERSALWEIRTRIANGULARSECTION = u * a
    Else
        'submerged flow
        u = Ce * Math.Sqr(2 * g * (h1 - h2))
        If (h2 - zCrest) <= dz Then
            'the water level is below the wings
            a = W * ((h2 - zCrest) ^ 2) / (2 * dz)
        Else
            'water level is above the wings
            a = W * (h2 - zCrest - dz / 2)
        End If
        QUNIVERSALWEIRTRIANGULARSECTION = u * a
    End If
End Function

Public Function QUNIVERSALWEIRRECTANGULARSECTION(h1 As Variant, h2 As Variant, Y1 As Double, Y2 As Double, Z1 As Double, Z2 As Double, Ce As Double) As Double
    'this function returns the flow over a rectangular section of a universal weir
    Dim zCrest As Double 'the lowest of both Z values represents the crest elevation
    Dim mlRectangular As Double
    zCrest = Application.WorksheetFunction.Min(Z1, Z2)
    Dim u As Double, W As Double, a As Double, g As Double
    
    g = 9.81
    W = Math.Abs(Y2 - Y1)
        
    mlRectangular = 2 / 3       'is a constant for rectangular sections
    
    If (h1 <= zCrest) Then
        'no flow
        QUNIVERSALWEIRRECTANGULARSECTION = 0
    ElseIf (h2 - zCrest) / (h1 - zCrest) < mlRectangular Then
        'free flow
        u = Ce * Math.Sqr(2 / 3 * g) * Math.Sqr(h1 - zCrest)
        a = 2 / 3 * W * (h1 - zCrest)
        QUNIVERSALWEIRRECTANGULARSECTION = u * a
    Else
        'submerged flow
        u = Ce * Math.Sqr(2 * g * (h1 - 2))
        a = W * (h2 - zCrest)
        QUNIVERSALWEIRRECTANGULARSECTION = u * a
    End If
    
End Function


Public Function QUNIVERSALWEIR(h1 As Variant, h2 As Variant, YRange As Range, ZRange As Range, Ce As Double) As Double
    Dim zCrest As Double
    Dim i As Integer
    Dim ui As Double, g As Double
    Dim mlRectangular As Double, mlTriangular As Double
    Dim dz As Double
    Dim Q As Double
    g = 9.81
    Q = 0
    mlRectangular = 2 / 3
    For i = 1 To ZRange.Rows.Count - 1
        zCrest = Application.WorksheetFunction.Min(ZRange.Cells(i, 1), ZRange.Cells(i + 1, 1))
        dz = Math.Abs(ZRange.Cells(i, 1) - ZRange(i + 1, 1))
        If h1 - zCrest < 1.25 * dz Then
            mlTriangular = 4 / 5
        Else
            mlTriangular = (2 / 3) + (1 / 6) * (dz / (h1 - zCrest))
        End If
    
        If ZRange.Cells(i, 1) = ZRange.Cells(i + 1, 1) Then
            'rectangular section
            Q = Q + QUNIVERSALWEIRRECTANGULARSECTION(h1, h2, YRange.Cells(i, 1), YRange.Cells(i + 1, 1), ZRange.Cells(i, 1), ZRange(i + 1, 1), Ce)
        Else
            'triangular section
            Q = Q + QUNIVERSALWEIRTRIANGULARSECTION(h1, h2, YRange.Cells(i, 1), YRange.Cells(i + 1, 1), ZRange.Cells(i, 1), ZRange(i + 1, 1), Ce)
        End If
    Next
    QUNIVERSALWEIR = Q
End Function

Public Function CeThomsonWeir(CrestLevel As Double, waterLevel As Double, BedLevel As Double, ChannelWidth As Double) As Double
    'this function calculates the discharge coefficient for a Thomson weir, based on the chart in https://edepot.wur.nl/411957 (Boiten)
    'this chart was digitized and the results are plotted here
    Dim P As Double, h1 As Double, b As Double
        
    'derive the local variables
    P = CrestLevel - BedLevel
    h1 = waterLevel - BedLevel
    b = ChannelWidth
    
    Dim pB As Double
    pB = Round(P / b, 1) 'since we only have p/B at increments of 0.1
    If pB < 0.1 Then pB = 0.1
    If pB > 1 Then pB = 1
    
    Dim h1p As Double 'h1/p
    h1p = h1 / P
   
    Dim pB01X As Variant, pB01Y As Variant
    Dim pB02X As Variant, pB02Y As Variant
    Dim pB03X As Variant, pB03Y As Variant
    Dim pB04X As Variant, pB04Y As Variant
    Dim pB05X As Variant, pB05Y As Variant
    Dim pB06X As Variant, pB06Y As Variant
    Dim pB07X As Variant, pB07Y As Variant
    Dim pB08X As Variant, pB08Y As Variant
    Dim pB09X As Variant, pB09Y As Variant
    Dim pB10X As Variant, pB10Y As Variant
    Dim XArray As Variant, YArray As Variant

    pB01X = Array(0.317111459968602, 0.376766091051805, 0.445839874411303, 0.511773940345369, 0.568288854003139, 0.649921507064364, 0.737833594976452, 0.822605965463108, 0.938775510204081, 1.05180533751962, 1.13343799058084, 1.24018838304552, 1.34065934065934, 1.41915227629513, 1.51334379905808, 1.6043956043956, 1.70172684458398, 1.78649921507064, 1.87755102040816)
    pB01Y = Array(0.578282208588957, 0.577668711656441, 0.577208588957055, 0.576901840490797, 0.576441717791411, 0.576211656441717, 0.57590490797546, 0.575828220858895, 0.575828220858895, 0.575828220858895, 0.575981595092024, 0.576288343558282, 0.576595092024539, 0.576978527607361, 0.577361963190184, 0.57782208588957, 0.578282208588957, 0.578895705521472, 0.579509202453987)
    pB02X = Array(0.313971742543171, 0.361067503924646, 0.420722135007849, 0.483516483516483, 0.555729984301412, 0.631083202511774, 0.725274725274725, 0.806907378335949, 0.897959183673469, 0.976452119309262, 1.07378335949764, 1.16169544740973, 1.25274725274725, 1.32496075353218, 1.39089481946624, 1.45682888540031, 1.52433281004709, 1.5949764521193, 1.65149136577708, 1.71114599686028, 1.78178963893249, 1.85557299843014)
    pB02Y = Array(0.578282208588957, 0.577668711656441, 0.577361963190184, 0.577055214723926, 0.577055214723926, 0.577285276073619, 0.57782208588957, 0.578435582822085, 0.57920245398773, 0.580199386503067, 0.581426380368098, 0.582653374233128, 0.584187116564417, 0.585644171779141, 0.5870245398773, 0.588558282208589, 0.590092024539877, 0.592009202453987, 0.593312883435582, 0.594923312883435, 0.596763803680981, 0.598987730061349)
    pB03X = Array(0.251177394034536, 0.310832025117739, 0.361067503924646, 0.408163265306122, 0.467817896389325, 0.533751962323391, 0.596546310832024, 0.665620094191522, 0.747252747252747, 0.822605965463108, 0.907378335949764, 0.99215070643642, 1.07378335949764, 1.17111459968602, 1.24332810047095, 1.31868131868131, 1.38775510204081, 1.45368916797488, 1.51020408163265, 1.56357927786499)
    pB03Y = Array(0.578435582822085, 0.578282208588957, 0.578282208588957, 0.578282208588957, 0.578435582822085, 0.578895705521472, 0.579509202453987, 0.58042944785276, 0.581656441717791, 0.583113496932515, 0.585030674846625, 0.587407975460122, 0.589785276073619, 0.592852760736196, 0.595613496932515, 0.598374233128834, 0.601211656441717, 0.603895705521472, 0.606349693251533, 0.608957055214723)
    pB04X = Array(0.197802197802197, 0.235478806907378, 0.298273155416012, 0.392464678178963, 0.470957613814756, 0.555729984301412, 0.627943485086342, 0.700156985871271, 0.769230769230769, 0.839874411302982, 0.897959183673469, 0.951334379905808, 0.993720565149136, 1.04866562009419, 1.09890109890109, 1.14128728414442, 1.1883830455259, 1.23076923076923)
    pB04Y = Array(0.578588957055214, 0.578282208588957, 0.578282208588957, 0.578895705521472, 0.579969325153374, 0.581349693251533, 0.583190184049079, 0.585030674846625, 0.587484662576687, 0.590092024539877, 0.592852760736196, 0.595460122699386, 0.598067484662576, 0.601671779141104, 0.605276073619631, 0.608266871165644, 0.611947852760736, 0.615245398773006)
    pB05X = Array(0.191522762951334, 0.232339089481946, 0.291993720565149, 0.354788069073783, 0.420722135007849, 0.499215070643642, 0.549450549450549, 0.612244897959183, 0.667189952904238, 0.717425431711146, 0.766091051805337, 0.813186813186813, 0.86342229199372, 0.913657770800627, 0.960753532182103, 1.00156985871271)
    pB05Y = Array(0.578588957055214, 0.578282208588957, 0.578282208588957, 0.578895705521472, 0.579969325153374, 0.581656441717791, 0.582883435582822, 0.585184049079754, 0.587791411042944, 0.59032208588957, 0.593159509202454, 0.596073619631901, 0.599754601226993, 0.603588957055214, 0.608190184049079, 0.6120245398773)
    pB06X = Array(0.207221350078492, 0.266875981161695, 0.313971742543171, 0.373626373626373, 0.430141287284144, 0.480376766091051, 0.533751962323391, 0.580847723704866, 0.62480376766091, 0.66248037676609, 0.703296703296703, 0.737833594976452, 0.778649921507064, 0.810047095761381)
    pB06Y = Array(0.578895705521472, 0.579049079754601, 0.579662576687116, 0.580582822085889, 0.581963190184049, 0.58372699386503, 0.585950920245398, 0.588711656441717, 0.591625766871165, 0.594079754601227, 0.597300613496932, 0.600368098159509, 0.603895705521472, 0.607116564417177)
    pB07X = Array(0.200941915227629, 0.248037676609105, 0.29513343799058, 0.345368916797488, 0.389324960753532, 0.423861852433281, 0.467817896389325, 0.500784929356358, 0.549450549450549, 0.577708006279434, 0.609105180533751, 0.631083202511774, 0.659340659340659, 0.675039246467817)
    pB07Y = Array(0.578895705521472, 0.57920245398773, 0.579969325153374, 0.581042944785276, 0.582423312883435, 0.584110429447852, 0.586257668711656, 0.588711656441717, 0.592392638036809, 0.594846625766871, 0.597914110429447, 0.600061349693251, 0.603435582822085, 0.605276073619631)
    pB08X = Array(0.197802197802197, 0.235478806907378, 0.273155416012558, 0.301412872841444, 0.342229199372056, 0.3861852433281, 0.423861852433281, 0.455259026687598, 0.489795918367346, 0.518053375196232, 0.546310832025117, 0.565149136577708, 0.587127158555729)
    pB08Y = Array(0.578895705521472, 0.57920245398773, 0.580122699386503, 0.581349693251533, 0.582730061349693, 0.585184049079754, 0.587484662576687, 0.590092024539877, 0.593159509202454, 0.596380368098159, 0.599754601226993, 0.602668711656441, 0.60542944785276)
    pB09X = Array(0.194662480376766, 0.219780219780219, 0.2574568288854, 0.285714285714285, 0.323390894819466, 0.357927786499215, 0.392464678178963, 0.417582417582417, 0.439560439560439, 0.461538461538461, 0.47723704866562, 0.499215070643642, 0.508634222919937)
    pB09Y = Array(0.578742331288343, 0.578895705521472, 0.579509202453987, 0.580582822085889, 0.582269938650306, 0.58441717791411, 0.587177914110429, 0.589938650306748, 0.592546012269938, 0.595153374233128, 0.597914110429447, 0.601134969325153, 0.60282208588957)
    pB10X = Array(0.122448979591836, 0.156985871271585, 0.200941915227629, 0.248037676609105, 0.27629513343799, 0.310832025117739, 0.335949764521193, 0.357927786499215, 0.376766091051805, 0.395604395604395, 0.414442700156985, 0.433281004709576, 0.445839874411303, 0.452119309262166)
    pB10Y = Array(0.577668711656441, 0.578128834355828, 0.57920245398773, 0.580582822085889, 0.581963190184049, 0.583496932515337, 0.585184049079754, 0.586871165644171, 0.589018404907975, 0.591165644171779, 0.593159509202454, 0.595920245398773, 0.598220858895705, 0.600368098159509)

    If pB = 0.1 Then
        XArray = pB01X
        YArray = pB01Y
    ElseIf pB = 0.2 Then
        XArray = pB02X
        YArray = pB02Y
    ElseIf pB = 0.3 Then
        XArray = pB03X
        YArray = pB03Y
    ElseIf pB = 0.4 Then
        XArray = pB04X
        YArray = pB04Y
    ElseIf pB = 0.5 Then
        XArray = pB05X
        YArray = pB05Y
    ElseIf pB = 0.6 Then
        XArray = pB06X
        YArray = pB06Y
    ElseIf pB = 0.7 Then
        XArray = pB07X
        YArray = pB07Y
    ElseIf pB = 0.8 Then
        XArray = pB08X
        YArray = pB08Y
    ElseIf pB = 0.9 Then
        XArray = pB09X
        YArray = pB09Y
    ElseIf pB = 1 Then
        XArray = pB10X
        YArray = pB10Y
    End If
    
    'now find the item with the nearest h1/p value in our XArray
    Dim i As Integer, MaxDiff As Double, itemnumber As Integer
    MaxDiff = 9E+99
    itemnumber = 0
    For i = 1 To UBound(XArray)
        If Math.Abs(XArray(i) - h1p) < MaxDiff Then
            MaxDiff = Math.Abs(XArray(i) - h1p)
            itemnumber = i
        End If
    Next
    
    'finally return our value from the YArray
    CeThomsonWeir = YArray(itemnumber)
       

End Function

Public Function CeSharpCrestedWeir(waterLevel As Double, CrestLevel As Double, BedLevel As Double) As Double
    'this function computes the discharge coefficient for a sharp crested weir
    'source: http://help.floodmodeller.com/floodmodeller/Technical_Reference/1D_Nodes_Reference/Weirs/Sharp_Crested_Weir.htm
    'equations origin from Kindsvater & Carter, 1957
    'the equation reads: Ce = 0.602 + 0.075 * h1/p1
    'where:
    'h1 = head above crest = waterlevel - crest elevation
    'p1 = apex (crest elevation - bed level)
    CeSharpCrestedWeir = 0.602 + 0.075 * (waterLevel - CrestLevel) / (CrestLevel - BedLevel)
End Function

Public Function QUnsuppressedRectangularWeir(waterLevel As Double, CrestLevel As Double, CrestWidth As Double) As Double
    'this function calculates the discharge over an unsuppressed sharp-crested rectangular weir.
    'this type of weir has no sideway-contraction and its width matches that of the channel
    'source; https://www.brighthubengineering.com/hydraulics-civil-engineering/65880-open-channel-flow-measurement-5-the-rectangular-weir/
    'in S.I. the equation reads: Q = 1.84 * B * H^1.5
    If waterLevel > CrestLevel Then
        QUnsuppressedRectangularWeir = 1.84 * CrestWidth * (waterLevel - CrestLevel) ^ (3 / 2)
    Else
        QUnsuppressedRectangularWeir = 0
    End If
End Function

Public Function QContractedRectangularWeir(waterLevel As Double, CrestLevel As Double, CrestWidth As Double) As Double
    'this function calculates the discharge over a contracted sharp-crested rectangular weir.
    'this type of weir has sideway-contraction since its crest is contained inside a construction of wing walls
    'source: https://www.brighthubengineering.com/hydraulics-civil-engineering/65880-open-channel-flow-measurement-5-the-rectangular-weir/
    'in S.I. the equation reads: Q = 1.84 * (L - 0.2H)*H^1.5
    If waterLevel > CrestLevel Then
        QContractedRectangularWeir = 1.84 * (CrestWidth - 0.2 * (waterLevel - CrestLevel)) * (waterLevel - CrestLevel) ^ 1.5
    Else
        QContractedRectangularWeir = 0
    End If
End Function

Public Function QTHOMSONWEIR(waterLevel As Double, CrestLevel As Double, BedLevel As Double, ChannelWidth As Double) As Double
    'note: the angle is the angle between both legs of the V-shape, in degrees
    'https://edepot.wur.nl/411957 (Boiten et. al.) contains a detailed description for this type of weir
    'it also contains a chart that provides ce as a function of h1/p
    'note: this function is meant for 90-degree V-shapes only. If you need another angle, use the QVNOTCHWEIR function instead!
    
    'discharge coefficient is a function of h1/P where h1 = water depth upstream w.r.t. bedlevel, P = crest elevation w.r.t. bedlevel
    'and also of p/B.
    'ce for a Thomson weir typically lies between 0.57 and 0.62
    Dim P As Double, b As Double, h1 As Double
    Dim Ce As Double
    Dim AngleDegrees As Double
    AngleDegrees = 90
    P = CrestLevel - BedLevel       'apex
    b = ChannelWidth
    h1 = waterLevel - CrestLevel    'waterhoogte min laagste kruin (het puntje)
    Ce = CeThomsonWeir(CrestLevel, waterLevel, BedLevel, ChannelWidth)
       
    Dim he As Double
    Dim kh As Double
    kh = 0.0008     'note: this value is only true when angle = 90 degrees
    
    he = h1 + kh    'effectieve overstortende straal
    QTHOMSONWEIR = 8 / 15 * Math.Sqr(2 * 9.81) * Ce * Math.Tan(DEG2RAD(AngleDegrees / 2)) * he ^ 2.5
    
End Function

Public Function QVNOTCHWEIR(h1 As Double, CrestLevel As Double, ShoulderElevation As Double, AngleDegrees As Double) As Double
    'this function calculates the discharge over a V-notch weir, using the Kinsvater-Shen equation.
    'source: USBR Water Measurement Manual Rev, revision 2001: https://www.usbr.gov/tsc/techreferences/mands/wmm/WMM_3rd_2001.pdf
    'originally, in empiral units, this equation reads: Q = 4.28 * Ce * tan(alpha/2)*h1e^5/2
    'here:
    'Q in ft3
    'h1e = h1 + kh (overstortende straal + correctiehoogte) in ft
    'where kh and Ce are a function of angle alpha
    'converted to SI this yields: Q = 2.362932145 * Ce * tan(alpha/2)*h1e^(5/2)
    
    Dim h1e As Double
    Dim Ce As Double
    Dim Angle As Double
    Dim kh As Double            'height correction
    Angle = DEG2RAD(AngleDegrees)
    
    'we must limit the discharge to the maximum filling of our V-notch
    If h1 > ShoulderElevation Then h1 = ShoulderElevation
        
    'Ce and kh depend on the angle. We have digitized the charts, converted to SI and fitted a trendline
    'kh = 0.0331 * angle(degrees) ^ -0.811
    'Ce = 6E-6* Angledegrees^2 - 0.0008 * Angledegrees + 0.6064
    kh = 0.0331 * AngleDegrees ^ -0.811
    Ce = 0.000006 * AngleDegrees ^ 2 - 0.0008 * AngleDegrees + 0.6064
    h1e = (h1 - CrestLevel) + kh
            
    QVNOTCHWEIR = 2.362932145 * Ce * Math.Tan(Angle / 2) * h1e ^ (5 / 2)
    
End Function

Public Function OutletLoss(outletlosscoef As Variant, Astruc As Variant, Achannel As Variant) As Double
    If Achannel > 0 Then
        OutletLoss = outletlosscoef * (1 - Astruc / Achannel) ^ 2
    Else
        OutletLoss = 0
    End If
End Function

Public Function FrictionLoss(Length As Variant, chezy As Variant, Rstruc As Variant) As Double
    If chezy > 0 And Rstruc > 0 Then
        FrictionLoss = (2 * 9.81 * Length) / (chezy ^ 2 * Rstruc)
    Else
        FrictionLoss = 0
    End If
End Function

Public Function EnergyLoss(InletLoss As Variant, FrictionLoss As Variant, OutletLoss As Variant) As Double
    EnergyLoss = 1 / Math.Sqr(InletLoss + FrictionLoss + OutletLoss)
End Function


Public Function QABUTMENTBRIDGE(h1 As Variant, h2 As Variant, ProfileVerticalShift As Variant, BridgeVerticalShift As Variant, muinlet As Variant, outletlosscoef As Variant, Length As Variant, nManning As Variant, BridgeTableProfileZRange As Range, BridgeTableProfileWRange As Range, ProfileYRange As Range, ProfileZRange As Range, Optional ByVal MaximizeMu As Boolean = False) As Variant
    
    'dit is een complexe functie omdat t.b.v. de outlet loss ook het benedenstroomse profiel bekend moet zijn.
    'let op: om numerieke redenen wordt in SOBEK de uiteindelijke mu gemaximaliseerd op 1. Dit is niet altijd terecht en kan resulteren in een overschatting van het verval over de brug.
    Dim mu As Variant
    Dim muoutlet As Variant
    Dim mufric As Variant
    Dim chezy As Variant
    
    'hydraulic variables
    Dim Abridge As Variant, Pbridge As Variant, Rbridge As Variant
    Dim AProfile As Variant, PProfile As Variant, RProfile As Variant
    
    
    'determine the water depth inside the bridge and inside the profile
    Dim BridgeDepth As Variant
    Dim ProfileDepth As Variant
    BridgeDepth = h1 - BridgeTableProfileZRange.Cells(1, 1) - ProfileVerticalShift
    ProfileDepth = h2 - Application.WorksheetFunction.Min(ProfileZRange) - ProfileVerticalShift
    
    'determine the hydraulic properties
    Call TabulatedHydraulicProperties(BridgeTableProfileZRange, BridgeTableProfileWRange, BridgeDepth, Abridge, Pbridge, Rbridge)
    Call YZHydraulicProperties(ProfileYRange, ProfileZRange, ProfileDepth, AProfile, PProfile, RProfile)
    chezy = Manning2Chezy(nManning, Rbridge)
    
    'calculate the energy loss coefficients
    muoutlet = OutletLoss(outletlosscoef, Abridge, AProfile)
    mufric = FrictionLoss(Length, chezy, Rbridge)
    mu = EnergyLoss(muinlet, muoutlet, mufric)
    
    If MaximizeMu Then mu = Application.WorksheetFunction.Min(mu, 1)
    
    QABUTMENTBRIDGE = mu * Abridge * Math.Sqr(2 * 9.81 * (h1 - h2))
    

End Function

Public Function QABUTMENTBRIDGEBASIC(mu As Double, Abridge As Double, h1 As Double, h2 As Double) As Double
    QABUTMENTBRIDGEBASIC = mu * Abridge * Math.Sqr(2 * 9.81 * (h1 - h2))
End Function


Public Function TabulatedWettedPerimeterByWaterlevel(waterLevel As Variant, ZRange As Range, WRange As Range, ProfileVerticalShift As Variant) As Variant
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Dim Depth As Variant
    Depth = waterLevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, a, P, r)
    TabulatedWettedPerimeterByWaterlevel = P
End Function

Public Function TabulatedHydraulicRadiusByWaterlevel(waterLevel As Variant, ZRange As Range, WRange As Range, ProfileVerticalShift As Variant) As Variant
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Dim Depth As Variant
    Depth = waterLevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, a, P, r)
    TabulatedHydraulicRadiusByWaterlevel = r
End Function


Public Function BREEDTE_STUW(AreaM2 As Variant, Q_MMPD As Variant, DischCoef As Variant, OVStraalM As Variant, Optional LatContrCoef As Variant = 1) As Variant
  
  'Free flow: Q = c * B * 2/3 * SQRT(2/3 * g) * (h1 - z)^1.5
  Dim Q As Variant
  Q = AreaM2 * Q_MMPD / 1000 / 24 / 3600
  BREEDTE_STUW = Q / (DischCoef * LatContrCoef * 2 / 3 * Sqr(2 / 3 * 9.81) * OVStraalM ^ 1.5)
  
End Function
Public Function WeirSubmerged(h1 As Variant, h2 As Variant, Z As Variant) As Boolean

'Free flow: als h2 - z < 2/3 * (h1 -z)
If h2 - Z < 2 / 3 * (h1 - Z) Then
  WeirSubmerged = False
Else
  WeirSubmerged = True
End If

End Function

Public Function DHGEVULDERONDEDUIKER(Q As Variant, d As Variant, L As Variant, n_manning As Variant, zi As Variant, zo As Variant) As Variant
  
  'Q = mu * A * SQR(2 * g * dh)
  'dh = Q^2 / (mu^2 * A^2 * 2g)
  
  Dim mu As Variant
  Dim chezy As Variant
  Dim zf As Variant 'ruwheidsverlies
  Dim a As Variant, P As Variant, r As Variant 'natte doorsnede, natte omtrek en hydraulische straal
  
    'we gaan uit van een volledig gevulde duiker
    a = 3.141592 * (d / 2) ^ 2
    P = 2 * 3.141592 * (d / 2)
    r = a / P
    chezy = Manning2Chezy(n_manning, r)
    zf = (2 * 9.81 * L) / (chezy ^ 2 * r)
    mu = 1 / Sqr(zi + zo + zf)
    DHGEVULDERONDEDUIKER = Q ^ 2 / (mu ^ 2 * a ^ 2 * 2 * 9.81)
  
End Function

Public Function DHRONDEDUIKER(Q As Variant, d As Variant, L As Variant, h1 As Variant, BOB1 As Variant, n_manning As Variant, zi As Variant, zo As Variant) As Variant
  
  'Q = mu * A * SQR(2 * g * dh)
  'dh = Q^2 / (mu^2 * A^2 * 2g)
  
  Dim mu As Variant
  Dim chezy As Variant
  Dim zf As Variant 'ruwheidsverlies
  Dim a As Variant, P As Variant, r As Variant 'natte doorsnede, natte omtrek en hydraulische straal
  
    'we gaan uit van een volledig gevulde duiker
    a = OPPERVLAKAFGEPLATTECIRKEL(d / 2, BOB1 + d / 2, h1)
    P = NATTEOMTREKAFGEPLATTECIRKEL(d / 2, BOB1 + d / 2, h1)
    r = a / P
    chezy = Manning2Chezy(n_manning, r)
    zf = (2 * 9.81 * L) / (chezy ^ 2 * r)
    mu = 1 / Sqr(zi + zo + zf)
    DHRONDEDUIKER = Q ^ 2 / (mu ^ 2 * a ^ 2 * 2 * 9.81)
  
End Function

Public Function WidthOrifice(Q As Variant, h1 As Variant, Drempel As Variant, ContrCoef As Variant, DisCoef As Variant) As Variant
  'berekent de benodigde breedte van een niet-verdronken orifice gegeven een debiet en drempelhoogte
  'Q = mu * c * B * d * SQR(2 * g * (h1 - (z + mu*d)))
  'we gaan uit van een onderkant-schuif die hoger ligt dan het waterpeil, dus d = h1-z
  'B = Q / (mu * c * (h1-z) * sqr(2 * g * (h1-z)))
  
  WidthOrifice = Q / (ContrCoef * DisCoef * (h1 - Drempel) * Sqr(2 * 9.81 * (h1 - Drempel)))
  

End Function

Public Function hydraulicRadius(a As Variant, P As Variant) As Variant
  'calculates the hydraulic radius from wetted area and wetted perimeter
  If P <= 0 Or a <= 0 Then
    hydraulicRadius = 0
  Else
    hydraulicRadius = a / P
  End If
End Function

Public Function Manning2Chezy(n_manning As Variant, r As Variant) As Variant
  'computes the Chezy roughness value from manning's coefficient and the hydraulic radius
  Manning2Chezy = r ^ (1 / 6) / n_manning
End Function
Public Function Chezy2Manning(chezy As Variant, r As Variant) As Variant
  Chezy2Manning = r ^ (1 / 5) / chezy
End Function
Public Function chezy(c As Variant, Depth As Variant, Width As Variant, bedslope As Variant) As Variant
    Dim Q As Variant
    Q = Depth * Width * c * Sqr((Depth * Width) / (Width + 2 * Depth) * bedslope)
    chezy = Q
End Function

Public Function MaatgevendeAfvoer(Oppervlak_ha As Variant, Optional Aanvoer_lpspha As Variant = 1.5) As Variant
  'oppervlak in ha
  'Aanvoer in l/s/ha
  'resultaat in m3/s
  
  MaatgevendeAfvoer = Oppervlak_ha * Aanvoer_lpspha / 1000

End Function

Public Function NeerslagPatroon(ValueRange As Range) As String
  'classificeert een neerslagpatroon (duren 24, 48, 96, 126, 196 uur) als een van de 7 patronen die STOWA onderscheidt
  'in "Nieuwe Neerslagstatistiek voor Waterbeheerders"
  'ga ervan uit dat de gegevens als uurcijfers worden aangeleverd.
  'we delen de patronen op in drie delen. Aan de hand van de onderlinge verhoudingen bepalen we de classificatie
  
  Dim d(1 To 8)
  Dim Gegevens As New Collection
  Dim i As Long
  
  Dim r As Long
  For r = 1 To ValueRange.Rows.Count
    Gegevens.Add ValueRange.Cells(r, 1)
  Next
  
  Dim sum As Variant, Dmax As Variant, PairMax As Variant, QuatMax As Variant
  Dim myPair As Variant, myQuat As Variant
  
  If Gegevens.Count = 24 Then
    Dmax = 0
    PairMax = 0
    QuatMax = 0
    'deel de periode op in drie vakken
    For i = 1 To Gegevens.Count / 8
      d(1) = d(1) + Gegevens(i)
    Next
    If d(1) > Dmax Then Dmax = d(1)
    
    For i = Gegevens.Count / 8 + 1 To Gegevens.Count * 2 / 8
      d(2) = d(2) + Gegevens(i)
    Next
    If d(2) > Dmax Then Dmax = d(2)
    
    For i = Gegevens.Count * 2 / 8 + 1 To Gegevens.Count * 3 / 8
      d(3) = d(3) + Gegevens(i)
    Next
    If d(3) > Dmax Then Dmax = d(3)
    
    For i = Gegevens.Count * 3 / 8 + 1 To Gegevens.Count * 4 / 8
      d(4) = d(4) + Gegevens(i)
    Next
    If d(4) > Dmax Then Dmax = d(4)
    
    For i = Gegevens.Count * 4 / 8 + 1 To Gegevens.Count * 5 / 8
      d(5) = d(5) + Gegevens(i)
    Next
    If d(5) > Dmax Then Dmax = d(5)
    
    For i = Gegevens.Count * 5 / 8 + 1 To Gegevens.Count * 6 / 8
      d(6) = d(6) + Gegevens(i)
    Next
    If d(6) > Dmax Then Dmax = d(6)
    
    For i = Gegevens.Count * 6 / 8 + 1 To Gegevens.Count * 7 / 8
      d(7) = d(7) + Gegevens(i)
    Next
    If d(7) > Dmax Then Dmax = d(7)
    
    For i = Gegevens.Count * 7 / 8 + 1 To Gegevens.Count
      d(8) = d(8) + Gegevens(i)
    Next
    If d(8) > Dmax Then Dmax = d(8)
    
    sum = d(1) + d(2) + d(3) + d(4) + d(5) + d(6) + d(7) + d(8)
    
    'doorloop alle mogelijke paren
    For i = 1 To 7
      myPair = d(i) + d(i + 1)
      If myPair > PairMax Then PairMax = myPair
    Next
    
    'doorloop alle combinaties van 4
    For i = 1 To 5
      myQuat = d(i) + d(i + 1) + d(i + 2) + d(i + 3)
      If myQuat > QuatMax Then QuatMax = myQuat
    Next
    
    'type hoog
    If PairMax > 0.85 * sum Then
      NeerslagPatroon = "hoog"
    ElseIf PairMax > 0.7 * sum Then
      NeerslagPatroon = "middelhoog"
    ElseIf PairMax > 0.55 * sum Then
      NeerslagPatroon = "middellaag"
    ElseIf QuatMax > 0.6 * sum Then
      NeerslagPatroon = "laag"
    ElseIf d(2) > 0.25 * sum And d(7) > 0.25 * sum Then
      NeerslagPatroon = "kort"
    ElseIf d(1) > 0.25 * sum And d(6) > 0.25 * sum Then
      NeerslagPatroon = "kort"
    ElseIf d(3) > 0.25 * sum And d(8) > 0.25 * sum Then
      NeerslagPatroon = "kort"
    ElseIf d(1) > 0.25 * sum And d(8) > 0.25 * sum Then
      NeerslagPatroon = "lang"
    Else
      NeerslagPatroon = "uniform"
    End If
    
  Else
    MsgBox ("Error: alleen neerslagduur 24 uur in uurcijfers wordt geaccepteerd.")
    NeerslagPatroon = ""
  End If
  
  

End Function

Public Function GUMBELFIT(OK As Boolean, xVals As Range, x0Range As Range, betaRange As Range) As Boolean
  Dim maxiter As Integer
  Dim threshold As Variant, mu As Variant, Sigma As Variant, beta_out As Variant
  Dim dif As Variant, olddif As Variant, step As Variant
  Dim i As Integer, dir As Integer, n As Integer
  Dim x0 As Variant, beta As Variant
        
  n = xVals.Rows.Count
  Dim x() As Variant
  ReDim x(1 To xVals.Rows.Count)
  For i = 1 To xVals.Rows.Count
    x(i) = xVals.Cells(i, 1)
  Next
      
  threshold = 0.00001
  maxiter = 50
  mu = Application.WorksheetFunction.Average(xVals)
  Sigma = Application.WorksheetFunction.StDev(xVals)
        
  'first estimate of beta
  beta = Sigma / 1.2825
  
  'failsafe
  If (beta = 0) Then beta = 0.01

  'calc beta
  dif = 999
  dir = 1
  step = 0.5
  i = 0
  OK = True
  While (OK And (Math.Abs(dif) > threshold))
    i = i + 1
    OK = True
    Call calc_pars_Gumbel(OK, beta_out, x0, x, beta, mu)
    If (OK) Then
        olddif = dif
        dif = Abs(beta - beta_out)
        If (olddif < dif) Then
         dir = -dir
        End If
        If (dif >= threshold) Then
          beta = beta - (dif * step * dir)
        End If
        If (i > maxiter) Then
          OK = False
        End If
    End If
 
    If (beta = 0) Then OK = False
  Wend
  
  betaRange.Cells(1, 1) = beta_out
  x0Range.Cells(1, 1) = x0
  GUMBELFIT = True
  
End Function

Public Function GumbelFitMLE(xVals As Range, muCell As Range, sigmaCell As Range, likelihoodCell As Range, AkaikeCell As Range) As Boolean
    'this function fits a gumbel distribution by finding the maximum likelihood
    'it sets the cell value for mu, sigma and the maximum likelihood
    
    Dim muMax As Variant, muMin As Variant
    Dim sMax As Variant, sMin As Variant
    Dim iMu As Integer, iSigma As Integer
    Dim mu As Variant, Sigma As Variant
    Dim Likelihood As Variant, bestLikelihood As Variant
    Dim BestMuIdx As Integer, bestSigmaIdx As Integer
    Dim bestMu As Variant, bestSigma As Variant
    Dim iIter As Integer
    
    'initialize mu and sigma
    muMax = Application.WorksheetFunction.max(xVals)
    muMin = Application.WorksheetFunction.Min(xVals)
    sMax = muMax - muMin
    sMin = 0
    bestLikelihood = -9E+99
    
    For iIter = 1 To 50
        For iMu = 1 To 10
            mu = muMin + (muMax - muMin) / 10 * (iMu - 0.5) 'take the centerpoint from the current section
            For iSigma = 1 To 10
                Sigma = sMin + (sMax - sMin) / 10 * (iSigma - 0.5) 'take the centerpoint from the current selection
                Likelihood = GumbelLogLikelihood(xVals, mu, Sigma)
                'Likelihood = GumbelLikelihood(xVals, mu, sigma)
                If Likelihood > bestLikelihood Then
                    BestMuIdx = iMu
                    bestSigmaIdx = iSigma
                    bestMu = mu
                    bestSigma = Sigma
                    bestLikelihood = Likelihood
                End If
            Next
        Next
        
        'another iteration complete. Narrow down based on the best result
        muMin = muMin + (muMax - muMin) / 10 * (BestMuIdx - 1.5)
        muMax = muMin + (muMax - muMin) / 10 * (BestMuIdx + 1.5)
        sMin = sMin + (sMax - sMin) / 10 * (bestSigmaIdx - 1.5)
        sMax = sMin + (sMax - sMin) / 10 * (bestSigmaIdx + 1.5)
                
    Next
    
    muCell.Cells(1, 1) = bestMu
    sigmaCell.Cells(1, 1) = bestSigma
    likelihoodCell.Cells(1, 1) = bestLikelihood
    AkaikeCell.Cells(1, 1) = Akaike(bestLikelihood, 2)
    GumbelFitMLE = True
    
End Function

Public Function GumbelLogLikelihood(xVals As Range, mu As Variant, Sigma As Variant) As Variant
    'this function computes the log likelihood for a given dataset and Gumbel distribution
    Dim myResult As Variant
    Dim P As Variant
    Dim r As Integer, Z As Variant
    For r = 1 To xVals.Rows.Count
        P = GUMBELPDF(mu, Sigma, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GumbelLogLikelihood = myResult
End Function

Public Function GEVLogLikelihood(xVals As Range, mu As Variant, Sigma As Variant, Zeta As Variant) As Variant
    'this function computes the log likelihood for a given dataset and Gumbel distribution
    Dim myResult As Variant
    Dim P As Variant
    Dim r As Integer, Z As Variant
    For r = 1 To xVals.Rows.Count
        P = GEVPDF(mu, Sigma, Zeta, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GEVLogLikelihood = myResult
End Function

Public Function GenParetoLogLikelihood(xVals As Range, mu As Variant, Sigma As Variant, kappa As Variant) As Variant
    'this function computes the log likelihood for a given dataset and Generalized Pareto distribution
    Dim myResult As Variant
    Dim P As Variant
    Dim r As Integer, Z As Variant
    For r = 1 To xVals.Rows.Count
        P = GENPARETOPDF(mu, Sigma, kappa, xVals.Cells(r, 1))
        If P = 0 Then
          myResult = -9E+99
          Exit For
        Else
          myResult = myResult + Math.Log(P)
        End If
    Next
    GenParetoLogLikelihood = myResult
End Function

Public Function Akaike(LogLikelihood As Variant, nParameters As Integer) As Variant
    Akaike = -2 * LogLikelihood + 2 * nParameters
End Function

Public Function Akaike2(Likelihood As Variant, nParameters As Integer) As Variant
    Akaike = -2 * Math.Log(Likelihood) + 2 * nParameters
End Function


Public Function calc_pars_Gumbel(OK As Boolean, BetaOut As Variant, X0Out As Variant, x() As Variant, BetaIn As Variant, muin As Variant) As Boolean
Dim sum_xi_exp_m_xi_dbeta As Variant
Dim sum_exp_m_xi_dbeta As Variant
Dim i As Integer

OK = True

sum_xi_exp_m_xi_dbeta = 0
sum_exp_m_xi_dbeta = 0
For i = 1 To UBound(x)
    If (BetaIn = 0) Then
        BetaOut = 0.01
        Exit For
    End If
    sum_exp_m_xi_dbeta = sum_exp_m_xi_dbeta + kf_exp(-x(i) / BetaIn)
    sum_xi_exp_m_xi_dbeta = sum_xi_exp_m_xi_dbeta + x(i) * kf_exp(-x(i) / BetaIn)
    If (sum_exp_m_xi_dbeta = 0) Then
      OK = False
      Exit For
    End If
Next
BetaOut = muin - sum_xi_exp_m_xi_dbeta / sum_exp_m_xi_dbeta
X0Out = -BetaOut * Math.Log(sum_exp_m_xi_dbeta / UBound(x))
calc_pars_Gumbel = True

End Function

Public Function kf_exp(V As Variant) As Variant
    'een e-macht met beveiliging tegen crashen
    If (V > 700) Then V = 700
    If (V < -708) Then
        kf_exp = 0
    Else
        kf_exp = Math.Exp(1) ^ V
    End If
End Function

Public Function GUMBELPDF(mu As Variant, Sigma As Variant, x As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de kansdichtheid volgens Gumbel
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  '------------------------------------------------------------------------------------------------
    
  Dim Z As Variant
  Z = (x - mu) / Sigma
  GUMBELPDF = 1 / Sigma * Math.Exp(1) ^ (-(Z + Math.Exp(1) ^ (-Z)))
  
End Function

Public Function GUMBELCDF(mu As Variant, Sigma As Variant, x As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens Gumbel type 1
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  'Kansverdelingsfunctie: F(x) = e^(-e^((mu-x)/sigma))
  '------------------------------------------------------------------------------------------------
  
  Dim e As Variant 'natuurlijke logaritme
  e = Math.Exp(1)
  
  GUMBELCDF = e ^ (-1 * e ^ ((mu - x) / Sigma))

End Function

Public Function CORRELATIONBYWINDOW(Range1 As Range, Range2 As Range, WindowSize As Integer) As Variant
    Dim i As Integer, j As Integer
    Dim Sum1 As Variant, Sum2 As Variant
    Dim Array1 As Collection, Array2 As Collection
    Set Array1 = New Collection
    Set Array2 = New Collection
    If Range1.Rows.Count <> Range2.Rows.Count Then
        MsgBox ("Error: ranges should have equal dimensions.")
    Else
        For i = 1 To Range1.Rows.Count Step WindowSize
            Sum1 = 0
            Sum2 = 0
            For j = i To i + Minimum(Range1.Rows.Count, WindowSize - 1)
                Sum1 = Sum1 + Range1.Cells(j, 1)
                Sum2 = Sum2 + Range2.Cells(j, 1)
            Next
            Array1.Add (Sum1)
            Array2.Add (Sum2)
        Next
    End If
    
    'calculate the correlation coefficient
    CORRELATIONBYWINDOW = CorrelationCoef(Array1, Array2)
        
End Function

Public Function CorrelationCoef(Coll1 As Collection, Coll2 As Collection) As Variant
    
    'start by finding the average value for both collections
    Dim Avg1 As Variant, Avg2 As Variant
    Dim i As Integer
    For i = 1 To Coll1.Count
        Avg1 = Avg1 + Coll1(i)
        Avg2 = Avg2 + Coll2(i)
    Next
    Avg1 = Avg1 / Coll1.Count
    Avg2 = Avg2 / Coll2.Count
    
    'now compute the correlation coefficient
    Dim Teller As Variant, Noemer1 As Variant, Noemer2 As Variant, Noemer As Variant
    For i = 1 To Coll1.Count
        Teller = Teller + (Coll1.Item(i) - Avg1) * (Coll2.Item(i) - Avg2)
        Noemer1 = Noemer1 + (Coll1.Item(i) - Avg1) ^ 2
        Noemer2 = Noemer2 + (Coll2.Item(i) - Avg2) ^ 2
    Next
    Noemer = Math.Sqr(Noemer1 * Noemer2)
    CorrelationCoef = Teller / Noemer
    
End Function

Public Function StdevAvgRatio(MyRange As Range) As Variant
    'computes StDev / Average for a given range.
    StdevAvgRatio = Application.WorksheetFunction.StDev(MyRange) / Application.WorksheetFunction.Average(MyRange)
End Function


Public Function EXPONENTIELEVERDELINGCDF(lambda As Variant, x As Variant, threshold As Variant) As Variant
    If x > threshold Then
        EXPONENTIELEVERDELINGCDF = 1 - Math.Exp(1) ^ (-lambda * (x - threshold))
    Else
        EXPONENTIELEVERDELINGCDF = 0
    End If
End Function

Public Function GUMBELINVERSE(P As Variant, mu As Variant, Sigma As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven de ONDERschrijdingskans p volgens Gumbel type 1
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansdichtheidsfunctie: f(x) = e^-x * e^(-e^(-x))
  'Kansverdelingsfunctie: F(x) = e^(-e^((mu-x)/sigma))
  '------------------------------------------------------------------------------------------------
    
    GUMBELINVERSE = mu - Sigma * (Math.Log(-1 * Math.Log(P)))

End Function

Public Function GENPARETOINVERSE(p_ond As Variant, mu As Variant, Sigma As Variant, kappa As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 25-7-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven de ONDERschrijdingskans p volgens Generalized Pareto
  'Cumulatieve kansdichtheidsfunctie: F(x) = 1-(1+kz)^-1/k waarin:
  'z = (x-mu)/sigma
  'LET OP: het is niet gelukt om deze formule te inverteren, dus we lossen het iteratief op
  '------------------------------------------------------------------------------------------------
  
   'we weten op welke onderschrijdingskans we willen uitkomen en zoeken daarbij de X
   'laten we X iteratief zoeken tussen mu - 10sigma en mu + 10sigma
   Dim iIter As Integer, iSlice As Integer
   Dim Xmin As Variant, Xmax As Variant, Slice As Variant
   Dim Xcur As Variant, Pcur As Variant, Pbest As Variant, BestSlice As Integer, Xbest As Variant
   
   Xmin = mu - 10 * Sigma
   Xmax = mu + 10 * Sigma
   Pbest = -1
   
   For iIter = 1 To 10
     'split the range in ten slices
     Slice = (Xmax - Xmin) / 10
     For iSlice = 1 To 10
        Xcur = Xmin + (iSlice - 0.5) * Slice
        Pcur = GENPARETOCDF(mu, Sigma, kappa, Xcur)
        If Math.Abs(Pcur - p_ond) < Math.Abs(Pbest - p_ond) Then
            Pbest = Pcur
            Xbest = Xcur
            BestSlice = iSlice
        End If
     Next
     
     'narrow down the search window and move on to the next iteration
     'build in some extra security by surrounding slices in the next iteration
     Xmin = Xbest - 2 * Slice
     Xmax = Xbest + 2 * Slice
     
   Next
   
    GENPARETOINVERSE = Xbest
  
End Function

Public Function Log10(myVal As Variant) As Variant
    'omdat VBA met LOG de natuurlijke logaritme bedoelt, voeren we hier een conversie uit naar 10Log
    Log10 = Log(myVal) / Log(10)
End Function



Public Function GLOCDF(mu As Variant, Sigma As Variant, teta As Variant, x As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 5-11-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens de GLO-verdeling (Generalized Logistic)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------
  
  Dim Z As Variant, T As Variant
  Z = (x - mu) / Sigma
    
  If teta = 0 Then
    GLOCDF = (1 + Exp(-Z)) ^ -1
  Else
    GLOCDF = (1 + (1 - teta * Z) ^ (1 / teta)) ^ -1
  End If
    
End Function

Public Function GLOINVERSE(mu As Variant, Sigma As Variant, teta As Variant, value As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 24-12-2018
  'Auteur: Siebe Bosch
  'Deze routine berekent de waarde X gegeven een ONDERschrijdingskans en een GLO-kansverdeling (Generalized Logistic)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------
  
  If teta = 0 Then
    GLOINVERSE = mu - Sigma * Math.Log(1 / value - 1)
  Else
    GLOINVERSE = mu + Sigma * ((1 - (1 / value - 1) ^ teta) / teta)
  End If
  
End Function


Public Function GEVPDF(mu As Variant, Sigma As Variant, Zeta As Variant, x As Variant) As Variant
  Dim e As Variant 'natuurlijke logaritme
  Dim Z As Variant, T As Variant
  
  e = Math.Exp(1)
  Z = (x - mu) / Sigma
    
  If Zeta = 0 Then
    T = e ^ -Z
  Else
    T = (1 + Zeta * Z) ^ (-1 / Zeta)
  End If
  
  GEVPDF = 1 / Sigma * T ^ (Zeta + 1) * e ^ -T
  
End Function


Public Function GEVCDFOLD(mu As Variant, Sigma As Variant, k As Variant, x As Variant) As Variant
  'calculates the cumulative probability density according to the GEV-probability distribution

   Dim Z As Variant
   'Dim arg1 As Variant
   'Dim arg2 As Variant
   
   Z = (x - mu) / Sigma
   'arg1 = (1 + k * z)
   'arg2 = -1 / k
   
   If k <> 0 Then
     GEVCDF = Exp(-1 * (1 + k * Z) ^ (-1 / k)) 'this is the original one
     'GEVCDF = Exp(-1 * arg1 ^ arg2)      'edit: this was necessary to prevent an invalid procedure call due to the numbers inside
   Else
     GEVCDF = Exp(-1 * Math.Exp(-Z))
   End If
   
End Function

Function ReturnPeriodGEV(mu As Double, Sigma As Double, Zeta As Double, x As Double) As Double
    Dim CDF As Double
    
    'Check if Zeta (shape parameter) is zero or not
    If Zeta = 0 Then
        'Calculate CDF for GEV (Gumbel type) distribution when Zeta is zero
        CDF = Exp(-Exp(-(x - mu) / Sigma))
    Else
        'Calculate CDF for GEV distribution when Zeta is not zero
        If ((x - mu) / Sigma) > (-1 / Zeta) Then
            CDF = Exp(-((1 + Zeta * ((x - mu) / Sigma)) ^ (-1 / Zeta)))
        Else
            CDF = 0
        End If
    End If
    
    'Calculate the return period based on the CDF
    If CDF < 1 Then
        ReturnPeriodGEV = 1 / (1 - CDF)
    Else
        ReturnPeriodGEV = CVErr(xlErrValue) ' return Excel #VALUE! error
    End If
    
End Function


Public Function GEVCDF(mu As Variant, Sigma As Variant, Zeta As Variant, x As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Datum: 21-3-2020 het minteken van Zeta omgekeerd. Nu consistent met de STOWA-notatie
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans van een bepaalde parameterwaarde volgens de GEV-verdeling (Gegeneraliseerde Extreme Waarden)
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  'Kansverdelingsfunctie:    F(x;\mu,\sigma,\xi) = \exp\left\{-\left[1+\xi\left(\frac{x-\mu}{\sigma}\right)\right]^{-1/\xi}\right\}
  '------------------------------------------------------------------------------------------------
  
  Dim e As Variant 'natuurlijke logaritme
  Dim Z As Variant, T As Variant
  
  e = Math.Exp(1)
  Z = (x - mu) / Sigma
    
  If Zeta = 0 Then
    T = e ^ -Z
  Else
    T = (1 - Zeta * Z) ^ (1 / Zeta)
  End If
  
  GEVCDF = e ^ -T
  
End Function

Public Function GEVINVERSE(mu As Variant, Sigma As Variant, Zeta As Variant, value As Variant) As Variant
  '------------------------------------------------------------------------------------------------
  'Datum: 9-11-2010
  'Auteur: Siebe Bosch
  'Deze routine berekent de ONDERschrijdingskans p van een bepaalde parameterwaarde volgens GEV-verdeling
  'dit betekent gewoon dat we de verdelingsfunctie gaan berekenen (= de integraal van de kansdichtheidsfunctie)
  '------------------------------------------------------------------------------------------------

  GEVINVERSE = mu + Sigma * (((-1 * Application.WorksheetFunction.Ln(value)) ^ (Zeta) - 1) / -Zeta)

End Function

Public Function EXPPDF(lambda As Variant, y As Variant, value As Variant) As Variant
    EXPPDF = lambda * Math.Exp(-lambda * (value - y))
End Function

Public Function EXPCDF(lambda As Variant, y As Variant, value As Variant) As Variant
    EXPCDF = 1 - Math.Exp(-lambda * (value - y))
End Function

Public Function EXPINVERSE(lambda As Variant, y As Variant, p_ond As Variant) As Variant
    EXPINVERSE = -(Math.Log(1 - p_ond)) / lambda + y
End Function

  Public Sub calcNeerslagStats(ByVal duration As Integer, ByVal Area As Variant, ByRef mu As Variant, ByRef gamma As Variant, ByRef kappa As Variant)
    'deze functie berekent de statistische parameters van de kansdichtheidsfunctie voor neerslagvolume in Nederland:
    'neerslagvolume voldoet namelijk aan de GEV-kansverdeling (Gegeneraliseerde Extremewaardenverdeling)
    'mu = locatieparameter' gamma = schaalparameter, kappa = vormparameter
    'waarden voor a1, a2, b1, b2 en c zijn aangeleverd door HKV-lijn in water
    'document Actuele extreme neerslagstatistiek en neerslag- en verdampingsreeksen, van 7 juli 2011: PR2197.10
    'originele bronvermelding: Overeem, A., T.A. Buishand, I. Holleman en R. Uijlenhoet, Extreme-value modeling of areal rainfall from weather radar, Water Resour. Res., 2010, 46, W09514, doi:10.1029/2009wr008517
    Dim y As Variant
    Dim a1 As Variant, a2 As Variant, b1 As Variant, b2 As Variant, c As Variant

    a1 = 17.92  'was in 2009 17.92
    a2 = 0.225  'was in 2009 0.225
    b1 = -3.57  'was in 2009 -3.57
    b2 = 0.427  'was in 2009 0.43
    c = 0.128   'was in 2009 0.128
    mu = a1 * duration ^ a2 + b1 * Area ^ c + (b2 * Area ^ c) * Math.Log(duration)

    a1 = 0.337  'was in 2009 0.344
    a2 = -0.018 'was in 2009 -0.025
    b1 = -0.014 'was in 2009 -0.016
    b2 = 0      'was in 2009 0.0003
    c = 0       'was in 2009 0
    y = a1 + a2 * Math.Log(duration) + b1 * Math.Log(Area) + b2 * duration * Math.Log(Area)
    gamma = y * mu

    a1 = -0.206 'was in 2009 -0.206
    a2 = 0      'was in 2009 0
    b1 = 0.018  'was in 2009 0.022
    b2 = 0      'was in 2009 -0.004
    c = 0       'was in 2009 0
    kappa = a1 + b1 * Math.Log(Area) + b2 * Math.Log(duration) * Math.Log(Area)

  End Sub


Public Function calcHerhalingstijd(ByVal Volume As Variant, ByVal Duur As Integer, ByVal Area As Variant) As Variant
    'berekent de herhalingstijd van een bui, gegeven Volume, Duur en gebiedsoppervlak)
    
    'Volume in mm
    'Duur in uren
    'Area in km2
    
    Dim mu As Variant, gamma As Variant, kappa As Variant
    Dim F_jaar As Variant 'overschrijdingsfrequentie op jaarbasis
    
    Call calcNeerslagStats(Duur, Area, mu, gamma, kappa)        'bereken de kansdichtheidsparameters
    F_jaar = (1 - kappa / gamma * (Volume - mu)) ^ (1 / kappa)  'frequentie in aantal keren / jaar
    calcHerhalingstijd = 1 / F_jaar                             'bereken de herhalingstijd van de gebeurtenis

    'onderstaand is een test of de terugrekening weer hetzelfde volume genereert
    'Dim myVol = calcNeerslagVolume(Area, Duration, ARI, mu, gamma, kappa)

  End Function
  
  Public Function calcNeerslagVolume(ByVal Area As Variant, ByVal Duur As Integer, ByVal Herhalingstijd As Variant) As Variant
    'Deze functie rekent terug. Gegeven duur, Oppervlak en overschrijdingskans
    'rekent hij het volume over een oppervlak groter dan puntneerslag uit
    Dim F_jaar As Variant
    Dim mu As Variant, gamma As Variant, kappa As Variant
    Call calcNeerslagStats(Duur, Area, mu, gamma, kappa)        'bereken de kansdichtheidsparameters
    
    F_jaar = 1 / Herhalingstijd
    calcNeerslagVolume = mu + gamma / kappa * (1 - F_jaar ^ kappa)

  End Function
  
  Sub ReadKNMIHourlyData(path As String, FromDate As Date, ToDate As Date, sheetName As String)
  
    Dim textFile As Integer
    Dim data As Variant
    Dim Headers As Variant
    Dim row As Long
    Dim resRow As Long
    Dim column As Long
    Dim HeaderCollection As Collection
    Set HeaderCollection = New Collection
    Dim UseRow As Boolean

    Dim targetWorksheet As Worksheet
    Set targetWorksheet = ActiveWorkbook.Worksheets(sheetName)
    
    ' Open the text file for input
    textFile = FreeFile()
    Open path For Input As textFile
        
    ' Read the data from row 34
    Do While Not EOF(textFile)
        Line Input #textFile, data
        data = Split(data, ",")
        row = row + 1
        UseRow = False
        
        If row = 32 Then
            'reads the header
            resRow = resRow + 1
            For column = 0 To UBound(data)
                targetWorksheet.Cells(resRow, column + 1).value = data(column)
                Call HeaderCollection.Add(data(column))
            Next column
        End If
        
        If row >= 34 Then
            For column = 0 To UBound(data)
                If Trim(HeaderCollection(column + 1)) = "YYYYMMDD" Then
                    
                    Dim datestring As String
                    Dim year As Integer
                    Dim month As Integer
                    Dim day As Integer
                    Dim hour As Integer

                    datestring = Trim(data(column))
                    year = Left(datestring, 4)
                    month = Mid(datestring, 5, 2)
                    day = Right(datestring, 2)
                    hour = data(column + 1)         'the hour is one column to the right
                                        
                    Dim myDate As Date
                    myDate = DateSerial(year, month, day)
                
                    ' Create a time value from the hour value
                    Dim myTime As Date
                    myTime = TimeSerial(hour, 0, 0)

                    ' Add the date and time values together to get a datetime value
                    Dim myDateTime As Date
                    myDateTime = myDate + myTime
                    
                    If myDateTime >= FromDate And myDateTime <= ToDate Then
                        resRow = resRow + 1
                        UseRow = True
                        targetWorksheet.Cells(resRow, column + 1).value = myDateTime
                    Else
                        UseRow = False
                    End If
                    
                ElseIf Trim(HeaderCollection(column + 1)) = "RH" And UseRow Then
                    'hourly precipitation is expressed in 0.1 mm so divide by 10. If -1 then < 0.05 mm so assume 0.025
                    If data(column) = -1 Then
                        targetWorksheet.Cells(resRow, column + 1).value = 0.025
                    Else
                        targetWorksheet.Cells(resRow, column + 1).value = data(column) / 10
                   End If
                ElseIf UseRow Then
                    targetWorksheet.Cells(resRow, column + 1).value = data(column)
                End If
            Next column
        End If
    Loop
    
    ' Close the text file
    Close textFile
    
End Sub

  
  Public Sub PrecipitationAreaReduction(ValuesRange As Range, CorrRange As Range, ActivityRange As Range, ProgressRange As Range, ByVal minHerh As Single, Optional ByVal Area As Variant = 6)
    'deze routine identificeert individuele buien uit een tijdreeks met uurlijkse neerslagsommen in Nederland
    'Oppervlak in km2
    'minHerh = minimum Herhalingstijd in jaren
    'voor puntneerslag houden we een standaardoppervlakte van 6 km2 aan
    Dim i As Integer, r As Long, k As Long, Dur As Integer, mySum As Variant, myNextSum As Variant, H As Single
    Dim SkipEvent As Boolean, HERH() As Variant, EventSum() As Variant, duration() As Integer
    Dim myMu As Variant, myGamma As Variant, myKappa As Variant 'probability function parameters
    Dim subRange As Range, CorrSum As Variant
    
    'opschonen bestaand resultaat en herdimensioneren arrays
    Call CorrRange.ClearContents
    ReDim HERH(1 To ValuesRange.Rows.Count)
    ReDim EventSum(1 To ValuesRange.Rows.Count)
    ReDim duration(1 To ValuesRange.Rows.Count)

    'doorloop alle neerslagduren van 1, 2, 4, 8, 12 en 24 uur
    For i = 1 To 6
      Select Case i
        Case Is = 1
          Dur = 1
        Case Is = 2
          Dur = 2
        Case Is = 3
          Dur = 4
        Case Is = 4
          Dur = 8
        Case Is = 5
          Dur = 12
        Case Is = 6
          Dur = 24
      End Select
    
      ActivityRange.Cells(1, 1) = "Analyseren neerslagduur " & Dur & " uur."
      DoEvents

      'doorloop de gecorrigeerde neerslagwaarden en onderscheid buien hierbinnen
      For r = 1 To ValuesRange.Rows.Count - 1
            
        mySum = Application.WorksheetFunction.sum(Range(ValuesRange.Cells(r, 1), ValuesRange.Cells(r + Dur - 1, 1)))
        myNextSum = Application.WorksheetFunction.sum(Range(ValuesRange.Cells(r + 1, 1), ValuesRange.Cells(r + Dur, 1)))

        If myNextSum < mySum Then 'nu weten we dat we een losse bui te pakken hebben
          'bereken de overschrijdingskans van deze puntneerslagsom en haal on the fly ook de bijbehorende Herhalingstijd binnen
          ProgressRange.Cells(1, 1) = r / ValuesRange.Rows.Count
          DoEvents
          
          H = calcHerhalingstijd(mySum, Dur, 6)

          'alleen als de herhalingstijd > minimum is, schrijven we hem weg
          If H >= minHerh Then

            'doorloop eerst de lijst met herhalingstijden om te checken of hij al is toegekend
            SkipEvent = False 'initialiseer SkipEvent
            For k = r To r + Dur - 1
              If HERH(k) > H Then
                'helaas, een gebeurtenis met kortere duur had al een grotere herhalingstijd. We skippen deze bui voor de huidige duur
                SkipEvent = True
                Exit For
              End If
            Next

            'als deze gebeurtenis nog niet is overruled door een zeldzamer herhalingstijd bij kortere duur:
            'leg de herhalingstijd vast!
            If Not SkipEvent Then
              For k = r To r + Dur - 1
                HERH(k) = H                'leg voor deze bui de herhalingstijd vast
                duration(k) = Dur         'leg voor deze bui de neerslagduur vast
                EventSum(k) = mySum       'leg voor deze bui de neerslagsom vast
              Next
            End If
            'Bui is afgehandeld, dus zet r aan het einde van de bui
            r = r + Dur - 1
          End If
        End If
      Next
      
    Next
    
    'update de voortgangsindicatoren
    ProgressRange.Cells(1, 1) = 0
    ActivityRange.Cells(1, 1) = "Berekent gecorrigeerde neerslagvolumes."
    DoEvents
    
    'doorloop nu alle cellen om de gecorrigeerde neerslagvolumes te berekenen en weg te schrijven
    For k = 1 To ValuesRange.Rows.Count
      If HERH(k) > 1 Then
        ProgressRange.Cells(1, 1) = k / ValuesRange.Rows.Count
        CorrSum = calcNeerslagVolume(Area, duration(k), HERH(k))
        CorrRange.Cells(k, 1) = ValuesRange.Cells(k, 1) * CorrSum / EventSum(k)
        DoEvents
      Else
        'geen correctie; neem oorspronkelijke waarde over
        CorrRange.Cells(k, 1) = ValuesRange.Cells(k, 1)
      End If
    Next
    
    'update de voortgangsindicatoren
    ProgressRange.Cells(1, 1) = 100
    ActivityRange.Cells(1, 1) = "Klaar."
    DoEvents
    

  End Sub

Public Sub ANNUALMAXIMUMPRECIPITATIONEVENTS(headerRow As Integer, DateCol As Integer, ValCol As Integer, duration As Integer, ResultsRow As Integer, ResultsCol As Integer, ProgressRange As Range)
    'Deze subroutine loopt door een volledige tijdreeks met neerslagvolumes en extraheert de maxima per jaar en seizoen
    
    Dim ValSubRange As Range
    Dim DateSubRange As Range
    Dim DateValRange As Range
    Dim i As Long, r As Long
    Dim myDate As Date, myYear As Integer, mySeizoen As String
    Dim mySum As Variant
    Dim MergeCells As Range
    
    Dim StartYear As Integer
    Dim EndYear As Integer
    
    'set de range
    r = headerRow
    While Not ActiveSheet.Cells(r + 1, DateCol) = ""
      r = r + 1
    Wend
    Set DateValRange = Range(ActiveSheet.Cells(headerRow + 1, DateCol), ActiveSheet.Cells(r, ValCol))
    
    StartYear = year(DateValRange.Cells(1, DateCol))
    EndYear = year(DateValRange.Cells(DateValRange.Rows.Count, DateCol))

    Dim JaarMaximaZOM() As Variant
    Dim JaarMaximaWin() As Variant
    Dim JaarMaxima() As Variant
    ReDim JaarMaximaZOM(StartYear To EndYear)
    ReDim JaarMaximaWin(StartYear To EndYear)
    ReDim JaarMaxima(StartYear To EndYear)
    
    For i = 1 To DateValRange.Rows.Count - duration + 1
      Set ValSubRange = DateValRange.Range(DateValRange.Cells(i, 2), DateValRange.Cells(i + duration - 1, 2))
      Set DateSubRange = DateValRange.Range(DateValRange.Cells(i, 1), DateValRange.Cells(i + duration - 1, 1))
      myDate = DateValRange.Cells(i, 1)
      myYear = year(myDate)
      mySeizoen = METEOROLOGISCHHALFJAAR(myDate)
      mySum = Application.WorksheetFunction.sum(ValSubRange)
      If mySum > JaarMaxima(myYear) Then
        JaarMaxima(myYear) = mySum
        ProgressRange = i / DateValRange.Rows.Count
        DoEvents
      End If
      If VBA.LCase(mySeizoen) = "zomer" Then
        If mySum > JaarMaximaZOM(myYear) Then JaarMaximaZOM(myYear) = mySum
      ElseIf VBA.LCase(mySeizoen) = "winter" Then
        If mySum > JaarMaximaWin(myYear) Then JaarMaximaWin(myYear) = mySum
      End If
    Next
        
    'create a section header and merge the cells
    r = ResultsRow
    Set MergeCells = Range(Cells(r, ResultsCol), Cells(r, ResultsCol + 3))
    MergeCells.Merge
    ActiveSheet.Cells(r, ResultsCol) = duration & "h"
    
    'write the column headers
    r = r + 1
    ActiveSheet.Cells(r, ResultsCol) = "jaar"
    ActiveSheet.Cells(r, ResultsCol + 1) = "jaarrond"
    ActiveSheet.Cells(r, ResultsCol + 2) = "zomer"
    ActiveSheet.Cells(r, ResultsCol + 3) = "winter"
    
    'write the results
    For i = StartYear To EndYear
      If JaarMaxima(i) > 0 Then
        r = r + 1
        ActiveSheet.Cells(r, ResultsCol) = i
        ActiveSheet.Cells(r, ResultsCol + 1) = JaarMaxima(i)
        ActiveSheet.Cells(r, ResultsCol + 2) = JaarMaximaZOM(i)
        ActiveSheet.Cells(r, ResultsCol + 3) = JaarMaximaWin(i)
      End If
    Next

End Sub

Public Function PLOTTINGPOSITIONFROMANNUALMAXIMA(myVal As Variant, ValuesRange As Range) As Variant
   Dim r As Long, n As Long, i As Long, F As Variant, H As Variant, P As Variant
   n = ValuesRange.Rows.Count
   Dim curVal As Variant
   
   'writes the return period in the second column of the range
  If ValuesRange.Columns.Count <> 1 Then
    MsgBox ("Range must contain only one column, containing the annual maxima.")
  Else
  
    'calculate the index number for the given value
    i = 0
    For r = 1 To ValuesRange.Rows.Count
      curVal = ValuesRange.Cells(r, 1)
      If curVal >= myVal Then i = i + 1
    Next
       
    'calculate the return period based on the index number
    P = (i - 0.3) / (n + 0.4) 'plotting position
    F = -Math.Log(1 - P) 'exceedance frequency in times per year
    PLOTTINGPOSITIONFROMANNUALMAXIMA = 1 / F 'return period
   
   End If
   
End Function

Public Sub IDENTIFYPRECIPITATIONEVENTSPOT(DateTimeCol As Long, ValueCol As Long, startRow As Long, endRow As Long, duration As Integer, POT As Variant, ResultsRow As Integer, ResultsCol As Integer, ProgressRange As Range)
    'Deze subroutine loopt door een volledige tijdreeks met neerslagvolumes en de totaalvolumes die een bepaalde POT-waarde overschrijden
    
    Dim i As Long, j As Long, r As Long, c As Long
    Dim myYear As Integer, mySeizoen As String
    
    Dim PrevRange As Range, CurRange As Range, NextRange As Range
    Dim PrevSum As Variant, CurSum As Variant, NextSum As Variant
    Dim Zomer As Collection, Winter As Collection, Jaarrond As Collection
    Dim myDate As Date
    
    Set Zomer = New Collection
    Set Winter = New Collection
    Set Jaarrond = New Collection
    
    For r = startRow + 1 To endRow - duration - 2
    
      ProgressRange.Cells(1, 1) = (r - startRow) / (endRow - startRow)
      DoEvents
    
      Set PrevRange = ActiveSheet.Range(Cells(r - 1, ValueCol), Cells(r + duration - 2, ValueCol))
      Set CurRange = ActiveSheet.Range(Cells(r, ValueCol), Cells(r + duration - 1, ValueCol))
      Set NextRange = ActiveSheet.Range(Cells(r + 1, ValueCol), Cells(r + duration, ValueCol))
      PrevSum = WorksheetFunction.sum(PrevRange)
      CurSum = WorksheetFunction.sum(CurRange)
      NextSum = WorksheetFunction.sum(NextRange)
      
      If CurSum > PrevSum And CurSum > NextSum And CurSum > POT Then
        myDate = ActiveSheet.Cells(r, DateTimeCol)
        myYear = year(myDate)
        mySeizoen = METEOROLOGISCHHALFJAAR(myDate)
        r = r + duration - 1                            'skip deze bui nu we hem geidentificeerd hebben
        If mySeizoen = "zomer" Then
          Call Zomer.Add(CurSum, str(myDate))
          Call Jaarrond.Add(CurSum, str(myDate))
        ElseIf mySeizoen = "winter" Then
          Call Winter.Add(CurSum, str(myDate))
          Call Jaarrond.Add(CurSum, str(myDate))
        End If
      End If
    Next
    
    r = ResultsRow
    c = ResultsCol
    ActiveSheet.Cells(r, c) = "Zomer"
    For i = 1 To Zomer.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Zomer.Item(i)
    Next
        
    r = ResultsRow
    c = c + 1
    ActiveSheet.Cells(r, c) = "Winter"
    For i = 1 To Winter.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Winter.Item(i)
    Next
    
    r = ResultsRow
    c = c + 1
    ActiveSheet.Cells(r, c) = "Jaarrond"
    For i = 1 To Jaarrond.Count
      r = r + 1
      ActiveSheet.Cells(r, c) = Jaarrond.Item(i)
    Next
    
End Sub

Public Sub CLASSIFYEVENTS(MyRange As Range, duration As Integer, ContainsHeader As Boolean, ClassMin As Variant, ClassMax As Variant)
  'verplicht: 1e kolom = datum, 2e kolom = waarde, 3e kolom = resultaat
  Dim r As Long, i As Long, Done As Boolean
  Dim ValRange As Range, resRange As Range, myCell As Range
  Dim sum As Variant, maxSum As Variant, maxIdx As Long
  Dim RankNum As Long, Ranks() As Long
  ReDim Ranks(1 To duration)
  
  Dim startRow As Integer
  If ContainsHeader Then
    startRow = 2
  Else
    startRow = 1
  End If
  
  'remove old results
  Set resRange = MyRange.Range(MyRange.Cells(startRow, 3), MyRange.Cells(MyRange.Rows.Count, 3))
  Call resRange.ClearContents
  
  While Not Done
    Done = True
    sum = 0
    maxSum = 0
    For r = startRow To MyRange.Rows.Count - duration
      Set ValRange = MyRange.Range(MyRange.Cells(r, 2), MyRange.Cells(r + duration - 1, 2))
      Set resRange = MyRange.Range(MyRange.Cells(r, 3), MyRange.Cells(r + duration - 1, 3))
      
      If Application.WorksheetFunction.sum(resRange) = 0 And Application.WorksheetFunction.sum(ValRange) >= ClassMin And Application.WorksheetFunction.sum(ValRange) <= ClassMax Then
        sum = Application.WorksheetFunction.sum(ValRange)
        If sum > maxSum Then
          maxIdx = r
          maxSum = sum
          Done = False
        End If
      End If
    Next
    
    If Done = False Then
      RankNum = RankNum + 1
      Set resRange = MyRange.Range(MyRange.Cells(maxIdx, 3), MyRange.Cells(maxIdx + duration - 1, 3))
      For i = 1 To duration
        Ranks(i) = RankNum
      Next
      resRange.value = Ranks
    End If
    
  Wend
End Sub

Public Sub MAXFROMMOVINGWINDOW(DATARANGEWITHCOLHEADER As Range, Durations As Collection, ResultsSheet As String, ResultsRow As Integer, ResultsCol As Integer)
    
    Dim curSheet As Worksheet, newSheet As Worksheet
    Dim CurSheetName As String
    CurSheetName = ActiveSheet.Name
    Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
        
    'first create a new worksheet for the results
    If Not WorkSheetExists(ResultsSheet) Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = ResultsSheet
        Set newSheet = ActiveWorkbook.Sheets(ResultsSheet)
    Else
        MsgBox ("Worksheet " & ResultsSheet & " already exists. Please remove the old one first.")
        Exit Sub
    End If
    
    Dim iDur As Integer, myDuration As Integer
    Dim r As Long, c As Long, i As Integer, r2 As Integer, c2 As Integer
    Dim id As String
    Dim myMax As Variant, mySum As Variant
    r2 = ResultsRow
    
    For iDur = 1 To Durations.Count
        r2 = r2 + 1
        c2 = ResultsCol
        myDuration = Durations(iDur)
        newSheet.Cells(r2, ResultsCol) = myDuration                          'write the duration to the results sheet
        For c = 1 To DATARANGEWITHCOLHEADER.Columns.Count
            myMax = 0
            c2 = c2 + 1
            id = DATARANGEWITHCOLHEADER.Cells(1, c)
            newSheet.Cells(ResultsRow, c2) = id                      'write the id to the results sheet
            For r = 2 To DATARANGEWITHCOLHEADER.Rows.Count
                mySum = 0
                For i = 0 To myDuration - 1
                    mySum = mySum + DATARANGEWITHCOLHEADER.Cells(r + i, c)
                Next
                If mySum > myMax Then myMax = mySum
             Next
             newSheet.Cells(r2, c2) = myMax
        Next
    Next
    
    MsgBox ("Operation complete.")
            
End Sub


Public Sub RANKNUMBEROFEXCEEDANCESBYMOVINGWINDOW(StartRowInclHeader As Integer, DataCol As Integer, RankCol As Integer, ResultsCol As Integer, MovingWindowSize As Integer, threshold As Variant, ProgressRange As Range, Optional ByVal OnlyIfSequential As Boolean = False)
  'This routine finds the number of exceedances of a given threshold in a given window size and
  'classifies them, using a moving window.
  Dim r As Long, c As Long, endRow As Long
  Dim ValRange As Range, RankRange As Range, TempRange As Range
  Dim maxFound As Integer, maxRow As Integer
  Dim maxExceedances As Integer, RankNum As Integer
  
  'find the last row
  endRow = StartRowInclHeader
  While Not ActiveSheet.Cells(endRow + MovingWindowSize, DataCol) = ""
    endRow = endRow + 1
  Wend
  
  'cleanup the results range
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, RankCol), Cells(endRow + MovingWindowSize - 1, RankCol))
  Call TempRange.ClearContents
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, ResultsCol), Cells(endRow + MovingWindowSize - 1, ResultsCol))
  Call TempRange.ClearContents
  
  'write the results header
  ActiveSheet.Cells(StartRowInclHeader, RankCol) = "Rangnummer"
  ActiveSheet.Cells(StartRowInclHeader, ResultsCol) = "Aantal aaneengesloten overschrijdingen"
  
  'start searching for threshold exceedances, using the moving window. Start with the largest numer of exceedances, end with the lowest ones
  maxFound = 999 'initialization
  While maxFound > 0
    
    maxFound = 0
    
    'find the window with the highest number of exceedances
    For r = StartRowInclHeader + 1 To endRow
      Set RankRange = Range(ActiveSheet.Cells(r, RankCol), ActiveSheet.Cells(r + MovingWindowSize - 1, RankCol))
      
      If Application.WorksheetFunction.sum(RankRange) = 0 Then
        Set ValRange = Range(ActiveSheet.Cells(r, DataCol), ActiveSheet.Cells(r + MovingWindowSize - 1, DataCol))
        'count the number of exceedances of the threshold
        If OnlyIfSequential Then
          maxExceedances = CountSequentialExceedances(ValRange, threshold)
        Else
          maxExceedances = Application.WorksheetFunction.CountIf(ValRange, "> " & threshold)
        End If
              
        'if the number exceeds the previously found number we'll overwrite the previous value
        If maxExceedances > maxFound Then
          maxFound = maxExceedances
          maxRow = r
        End If
      End If
    Next
    
    'write the ranking number found to the results column
    If maxFound > 0 Then
      RankNum = RankNum + 1
      
      Set TempRange = ActiveSheet.Range(Cells(maxRow, RankCol), Cells(maxRow + MovingWindowSize - 1, RankCol))
      TempRange.value = RankNum
      Set TempRange = ActiveSheet.Range(Cells(maxRow, ResultsCol), Cells(maxRow + MovingWindowSize - 1, ResultsCol))
      TempRange.value = maxFound
    End If
    
    'update the progress indicator
    ProgressRange.value = RankNum & " of maximum " & (endRow - StartRowInclHeader) / MovingWindowSize
    DoEvents
  
  Wend
    
End Sub
    
Public Sub RANKNUMBEROFEXCEEDANCESBYINTERVAL(StartRowInclHeader As Integer, DataCol As Integer, RankCol As Integer, ResultsCol As Integer, IntervalSize As Integer, threshold As Variant, Optional ByVal OnlyIfSequential As Boolean = False)
  'This routine finds the number of exceedances of a given threshold in a given window size and
  'classifies them, using a fixed interval
  Dim r As Long, c As Long, startRow As Long, endRow As Long
  Dim ValRange As Range, RankRange As Range, TempRange As Range
  Dim maxFound As Integer, maxRow As Integer
  Dim nExceedances As Integer, blockNum As Integer
  
  'find the last row
  endRow = StartRowInclHeader
  startRow = StartRowInclHeader + 1
  While Not ActiveSheet.Cells(endRow + IntervalSize, DataCol) = ""
    endRow = endRow + 1
  Wend
  
  'cleanup the results range
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, RankCol), Cells(endRow + IntervalSize - 1, RankCol))
  Call TempRange.ClearContents
  Set TempRange = ActiveSheet.Range(Cells(StartRowInclHeader + 1, ResultsCol), Cells(endRow + IntervalSize - 1, ResultsCol))
  Call TempRange.ClearContents
  
  'write the results header
  ActiveSheet.Cells(StartRowInclHeader, RankCol) = "Rangnummer"
  ActiveSheet.Cells(StartRowInclHeader, ResultsCol) = "Aantal aaneengesloten overschrijdingen"
  
  For r = startRow To endRow Step IntervalSize
    blockNum = blockNum + 1
    Set ValRange = ActiveSheet.Range(Cells(r, DataCol), Cells(r + IntervalSize - 1, DataCol))
    
    If OnlyIfSequential Then
      nExceedances = CountSequentialExceedances(ValRange, threshold)
    Else
      nExceedances = Application.WorksheetFunction.CountIf(ValRange, "> " & threshold)
    End If
    
    Set TempRange = ActiveSheet.Range(Cells(r, RankCol), Cells(r + IntervalSize - 1, RankCol))
    TempRange.value = blockNum
    Set TempRange = ActiveSheet.Range(Cells(r, ResultsCol), Cells(r + IntervalSize - 1, ResultsCol))
    TempRange.value = nExceedances
  Next
        
End Sub

Public Sub POTANALYSISSUM(headerRow As Integer, DateCol As Integer, ValCol As Integer, EventIndexCol As Integer, EventSumCol As Integer, EventExceedanceCol As Integer, duration As Integer, MinimumTimeStepsBetweenEvents As Integer, PotExceedanceFrequencyPerYear As Integer, IncludeSummer As Boolean, IncludeWinter As Boolean, ProgressRange As Range)
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  'Datum: 29-7-2014
  'Auteur: Siebe Bosch
  'Deze routine indexeert de zwaarste neerslaggebeurtenissen uit een opgegeven range en schrijft de indexnummers naar een naastgelegen resultatenkolom
  'Bovendien maakt hij een overzicht van alle bijkomende volumes
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  Dim i As Long, myIdx As Integer, r As Long, lastRow As Long
  Dim subValRange As Range, subIdxRange As Range
  Dim maxSum As Variant, maxIdx As Long, idxSum As Variant
  Dim startDate As Date, EndDate As Date, nDays As Long
  Dim EventSum() As Variant, MaxEvents As Long
  Dim maxCol As Integer
  Dim DateValResultsRange As Range
  
  'zoek het bereik van de gegevens
  r = headerRow
  While Not ActiveSheet.Cells(r + 1, DateCol) = ""
    r = r + 1
  Wend
  maxCol = WorksheetFunction.max(DateCol, ValCol, EventIndexCol, EventSumCol)
  Set DateValResultsRange = ActiveSheet.Range(Cells(headerRow + 1, DateCol), Cells(r, maxCol))
  
  ActiveSheet.Cells(headerRow, EventIndexCol) = "Index for " & duration & "h"
  
  'zoek de start- en einddatum en bereken het gewenste aantal overschrijdingen van de POT-waarde
  startDate = DateValResultsRange.Cells(1, DateCol)
  EndDate = DateValResultsRange.Cells(DateValResultsRange.Rows.Count, DateCol)
  nDays = EndDate - startDate
  MaxEvents = nDays / 365.25 * PotExceedanceFrequencyPerYear
  ReDim EventSum(1 To MaxEvents)
  
  'opschonen oude resultaten
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventIndexCol), Cells(DateValResultsRange.Rows.Count, EventIndexCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventSumCol), Cells(DateValResultsRange.Rows.Count, EventSumCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventExceedanceCol), Cells(DateValResultsRange.Rows.Count, EventExceedanceCol))
  Call subIdxRange.ClearContents
  
  'create a moving window array that contains the sum of each window
  Dim movingWindowSum() As Variant, inUseSum() As Integer
  ReDim movingWindowSum(1 To DateValResultsRange.Rows.Count)
  ReDim inUseSum(1 To DateValResultsRange.Rows.Count)
  For i = 1 To DateValResultsRange.Rows.Count - duration + 1
    movingWindowSum(i) = Application.WorksheetFunction.sum(DateValResultsRange.Range(DateValResultsRange.Cells(i, 2), DateValResultsRange.Cells(i + duration - 1, 2)))
  Next
    
  'next walk through the moving window array to find the highest volumes, starting with the maximum (rank 1) and moving up in rank (=lower volume)
  For myIdx = 1 To MaxEvents
    ProgressRange = myIdx / MaxEvents
    DoEvents
    maxSum = 0
    
    'make a distinction between summer and winter if required
    If IncludeSummer = True And IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          maxSum = movingWindowSum(i)
          maxIdx = i
        End If
      Next
    ElseIf IncludeSummer = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "zomer" Then
            maxSum = movingWindowSum(i)
            maxIdx = i
          End If
        End If
      Next
    ElseIf IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowSum(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowSum(i) > maxSum And inUseSum(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "winter" Then
          maxSum = movingWindowSum(i)
          maxIdx = i
          End If
        End If
      Next
    End If
    
    'write the index number to the worksheet
    Set subIdxRange = DateValResultsRange.Range(DateValResultsRange.Cells(maxIdx, EventIndexCol), DateValResultsRange.Cells(maxIdx + duration - 1, EventIndexCol))
    subIdxRange = myIdx
    
    'zet de relevante velden in de inUse-array plus uitloop voor de minimumruimte tussen twee events op 'bezet'. Let op: ook een stuk terug! de array bevat immers een vooruitblik
    For i = maxIdx To Application.WorksheetFunction.Min(maxIdx + duration - 1 + MinimumTimeStepsBetweenEvents, DateValResultsRange.Rows.Count)
      inUseSum(i) = 1
    Next
    For i = (maxIdx) To Application.WorksheetFunction.max(1, (maxIdx - duration + 1 - MinimumTimeStepsBetweenEvents)) Step -1
      inUseSum(i) = 1
    Next
        
    'store the event sum in the array
    EventSum(myIdx) = maxSum
  Next
  
  'finally write the event sums and threshold exceedance sums to the worksheet
  r = headerRow
  ActiveSheet.Cells(r, EventSumCol) = "Volume." & duration & "h"
  ActiveSheet.Cells(r, EventExceedanceCol) = "Exceedance." & duration & "h"
  For myIdx = 1 To MaxEvents
    r = r + 1
    ActiveSheet.Cells(r, EventSumCol) = EventSum(myIdx)
    ActiveSheet.Cells(r, EventExceedanceCol) = EventSum(myIdx) - EventSum(MaxEvents)
  Next

End Sub

Public Sub POTANALYSISMAX(headerRow As Integer, DateCol As Integer, ValCol As Integer, EventIndexCol As Integer, EventMaxCol As Integer, EventExceedanceCol As Integer, duration As Integer, MinimumTimeStepsBetweenEvents As Integer, PotExceedanceFrequencyPerYear As Integer, IncludeSummer As Boolean, IncludeWinter As Boolean, ProgressRange As Range)
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  'Datum: 29-7-2014
  'Auteur: Siebe Bosch
  'Deze routine indexeert de zwaarste gebeurtenissen (op basis van maximum) uit een opgegeven range en schrijft de indexnummers naar een naastgelegen resultatenkolom
  'Bovendien maakt hij een overzicht van alle bijkomende volumes
  '---------------------------------------------------------------------------------------------------------------------------------------------------
  Dim i As Long, myIdx As Integer, r As Long, lastRow As Long
  Dim subValRange As Range, subIdxRange As Range
  Dim maxVal As Variant, maxIdx As Long, idxMax As Variant
  Dim startDate As Date, EndDate As Date, nDays As Long
  Dim EventMax() As Variant, MaxEvents As Long
  Dim maxCol As Integer
  Dim DateValResultsRange As Range
  
  'zoek het bereik van de gegevens
  r = headerRow
  While Not ActiveSheet.Cells(r + 1, DateCol) = ""
    r = r + 1
  Wend
  maxCol = WorksheetFunction.max(DateCol, ValCol, EventIndexCol, EventMaxCol)
  Set DateValResultsRange = ActiveSheet.Range(Cells(headerRow + 1, DateCol), Cells(r, maxCol))
  
  ActiveSheet.Cells(headerRow, EventIndexCol) = "Index for " & duration & "h"
  
  'zoek de start- en einddatum en bereken het gewenste aantal overschrijdingen van de POT-waarde
  startDate = DateValResultsRange.Cells(1, DateCol)
  EndDate = DateValResultsRange.Cells(DateValResultsRange.Rows.Count, DateCol)
  nDays = EndDate - startDate
  MaxEvents = nDays / 365.25 * PotExceedanceFrequencyPerYear
  ReDim EventMax(1 To MaxEvents)
  
  'opschonen oude resultaten
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventIndexCol), Cells(DateValResultsRange.Rows.Count, EventIndexCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventMaxCol), Cells(DateValResultsRange.Rows.Count, EventMaxCol))
  Call subIdxRange.ClearContents
  Set subIdxRange = DateValResultsRange.Range(Cells(1, EventExceedanceCol), Cells(DateValResultsRange.Rows.Count, EventExceedanceCol))
  Call subIdxRange.ClearContents
  
  'create a moving window array that contains the Max of each window
  Dim movingWindowMax() As Variant, inUseMax() As Integer
  ReDim movingWindowMax(1 To DateValResultsRange.Rows.Count)
  ReDim inUseMax(1 To DateValResultsRange.Rows.Count)
  For i = 1 To DateValResultsRange.Rows.Count - duration + 1
    movingWindowMax(i) = Application.WorksheetFunction.max(DateValResultsRange.Range(DateValResultsRange.Cells(i, 2), DateValResultsRange.Cells(i + duration - 1, 2)))
  Next
    
  'next walk through the moving window array to find the highest volumes, starting with the maximum (rank 1) and moving up in rank (=lower volume)
  For myIdx = 1 To MaxEvents
    ProgressRange = myIdx / MaxEvents
    DoEvents
    maxVal = 0
    
    'make a distinction between Summer and winter if required
    If IncludeSummer = True And IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          maxVal = movingWindowMax(i)
          maxIdx = i
        End If
      Next
    ElseIf IncludeSummer = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "zomer" Then
            maxVal = movingWindowMax(i)
            maxIdx = i
          End If
        End If
      Next
    ElseIf IncludeWinter = True Then
      For i = 1 To DateValResultsRange.Rows.Count - duration + 1
        If movingWindowMax(i) = 0 Then
          i = i + duration - 1
        ElseIf movingWindowMax(i) > maxVal And inUseMax(i) = 0 Then
          If METEOROLOGISCHHALFJAAR(DateValResultsRange.Cells(i, 1)) = "winter" Then
          maxVal = movingWindowMax(i)
          maxIdx = i
          End If
        End If
      Next
    End If
    
    'write the index number to the worksheet
    Set subIdxRange = DateValResultsRange.Range(DateValResultsRange.Cells(maxIdx, EventIndexCol), DateValResultsRange.Cells(maxIdx + duration - 1, EventIndexCol))
    subIdxRange = myIdx
    
    'zet de relevante velden in de inUse-array plus uitloop voor de minimumruimte tussen twee events op 'bezet'. Let op: ook een stuk terug! de array bevat immers een vooruitblik
    For i = maxIdx To Application.WorksheetFunction.Min(maxIdx + duration - 1 + MinimumTimeStepsBetweenEvents, DateValResultsRange.Rows.Count)
      inUseMax(i) = 1
    Next
    For i = (maxIdx) To Application.WorksheetFunction.max(1, (maxIdx - duration + 1 - MinimumTimeStepsBetweenEvents)) Step -1
      inUseMax(i) = 1
    Next
        
    'store the event Max in the array
    EventMax(myIdx) = maxVal
  Next
  
  'finally write the event Maxs and threshold exceedance Maxs to the worksheet
  r = headerRow
  ActiveSheet.Cells(r, EventMaxCol) = "Max." & duration & "h"
  ActiveSheet.Cells(r, EventExceedanceCol) = "Exceedance." & duration & "h"
  For myIdx = 1 To MaxEvents
    r = r + 1
    ActiveSheet.Cells(r, EventMaxCol) = EventMax(myIdx)
    ActiveSheet.Cells(r, EventExceedanceCol) = EventMax(myIdx) - EventMax(MaxEvents)
  Next

End Sub


Public Sub CalculateExtremeEvents(Volumes() As Variant, DuurInUse() As Long, Duur As Long, dIdx As Long, nExtremen As Long, ProgressRange As Range, startRow As Long)
  
  Dim r As Long, c As Long, i As Long, j As Long, k As Long, rMax As Long
  Dim NSLMaxSom As Variant, mySom As Variant 'rMax is het rijnummer van het record van de hoogste dagneerslag, NSLMax de dagsom
  Dim SumRange As Range, Inuse As Long
    
  '----------------------------------------------------------------------------------------------------------------------------------
  'Datum: 1-11-2010
  'Auteur: Siebe Bosch
  'Deze routine zoekt in een array van neerslagvolumes welke neerslaggebeurtenissen met een opgegeven duur de 1000 zwaarste zijn
  'op het werkblad schrijft hij daarna het volgnummer van de zwaarte weg. 1 = zwaarste, 1000 = minst zware
  '----------------------------------------------------------------------------------------------------------------------------------
  For i = 1 To nExtremen                            'bepaal de zwaarste neerslaggebeurtenissen met de opgegeven duur
    ProgressRange.Cells(1, 1) = i / nExtremen
    DoEvents
    
    NSLMaxSom = 0                                   'initialiseer de maximum neerslagsom
    For j = 1 To UBound(Volumes(), 1) - Duur + 1    'doorloop de hele reeks en zoek naar de zwaarste neerslagsom over het opgegeven aantal uren
      
      mySom = Volumes(j, 1)                         'de som begint altijd met het eerste record
      Inuse = DuurInUse(j, dIdx)                    'controleer of dit record niet al aangemerkt is als "inuse" oftewel: al in een maximum verwerkt
      If Inuse = 0 Then                             'alleen als dit record nog niet in gebruik is
        For k = 1 To Duur - 1                       'doorloop de rest van de uurgegevens voor de opgegeven neerslagduur
          mySom = mySom + Volumes(j + k, 1)         'sommeer ze
          Inuse = DuurInUse(j + k, dIdx)            'wederom controle of geen van de opgetelde volumes al in gebruik waren
          If Inuse > 0 Then
            j = j + k + Duur - 1                    'als een record binnen de komende duur al in gebruik is, kunnen we de teller meteen doorzetten tot voorbij de hele neerslaggebeurtenis
            Exit For                                  'als een record al in gebruik is, kunnen we deze loop meteen al verlaten
          End If
        Next
        If mySom > NSLMaxSom And Inuse = 0 Then       'alleen als de som over de duur groter is dan het totnogtoe geregistreerde maximum EN geen van de records is in gebruik, gaan we door
          rMax = j + startRow - 1                     'registreer het rijnummer dat hoort bij de gevonden duur met maximum
          NSLMaxSom = mySom                           'let het gevonden maximum ook als zodanig vast
        End If
      End If
    Next
    
    'nu we de (i-1)-na zwaarste gebeurtenis hebben gevonden, voegen we hem toe aan de collectie
    'door een vlaggetje met het volgnummer naast het record te zetten staat hij ook meteen te boek als reeds verwerkt
    For r = rMax To rMax + Duur - 1
      ActiveSheet.Cells(r, 2 + dIdx) = i
      DuurInUse(r - startRow, dIdx) = i
    Next
    
  Next
End Sub

Public Function NASH_SUTCLIFFE(MyRange As Range, Datumcol As Integer, MeasCol As Integer, ValsCol As Integer, ContainsHeader As Boolean, Optional ByVal Log As Boolean = False) As Variant
  
  On Error GoTo Errorhandling
  
  Dim nObserved As Long
  Dim sum As Variant, sumLog As Variant, AvgObserved As Variant, AvgLogObserved As Variant
  Dim sumTeller As Variant, sumNoemer As Variant
  Dim ErrStr As String, r As Long, startRow As Integer
  
  If ContainsHeader Then
    startRow = 2
  Else
    startRow = 1
  End If
  
  sum = 0
  nObserved = 0
  For r = startRow To MyRange.Rows.Count
    nObserved = nObserved + 1
    sum = sum + MyRange.Cells(r, MeasCol)
    If MyRange.Cells(r, MeasCol) > 0 Then sumLog = sumLog + Math.Log(MyRange.Cells(r, MeasCol)) 'log-NS
  Next
  
  'calculate the average
  If nObserved = 0 Then
    ErrStr = "No measured data found to compare computed data with. Please check from- and to-dates and time series with measured data."
    GoTo Errorhandling
  Else
    AvgObserved = sum / nObserved
    AvgLogObserved = sumLog / nObserved
  End If
  
  For r = startRow To MyRange.Rows.Count
    If Not Log Then
      sumTeller = sumTeller + (MyRange.Cells(r, MeasCol) - MyRange.Cells(r, ValsCol)) ^ 2
      sumNoemer = sumNoemer + (MyRange.Cells(r, MeasCol) - AvgObserved) ^ 2
    Else
      sumTeller = sumTeller + (Math.Log(MyRange.Cells(r, MeasCol)) - Math.Log(MyRange.Cells(r, ValsCol))) ^ 2
      sumNoemer = sumNoemer + (Math.Log(MyRange.Cells(r, MeasCol)) - AvgLogObserved) ^ 2
    End If
  Next
  
  NASH_SUTCLIFFE = 1 - (sumTeller / sumNoemer)
  Exit Function
  
Errorhandling:
  MsgBox ("Error in function calcNashSutcliffe. " & ErrStr)
  End
  
  
End Function


Public Function NASH_SUTCLIFFE_FAST(Observed As Range, Computed As Range, Optional ByVal Log As Boolean = False) As Variant
  
  On Error GoTo Errorhandling
  
  Dim sum As Variant, sumLog As Variant, AvgObserved As Variant, AvgLogObserved As Variant
  Dim sumTeller As Variant, sumNoemer As Variant
  Dim ErrStr As String, r As Long
    
  sum = 0
  For r = 1 To Observed.Rows.Count
    sum = sum + Observed.Cells(r, 1)
    If Observed.Cells(r, 1) > 0 Then sumLog = sumLog + Math.Log(Observed.Cells(r, 1)) 'log-NS
  Next
  
  'calculate the average
  If Observed.Rows.Count = 0 Then
    ErrStr = "No measured data found to compare computed data with. Please check from- and to-dates and time series with measured data."
    GoTo Errorhandling
  Else
    AvgObserved = sum / Observed.Rows.Count
    AvgLogObserved = sumLog / Observed.Rows.Count
  End If
  
  For r = 1 To Computed.Rows.Count
    If Not Log Then
      sumTeller = sumTeller + (Observed.Cells(r, 1) - Computed.Cells(r, 1)) ^ 2
      sumNoemer = sumNoemer + (Observed.Cells(r, 1) - AvgObserved) ^ 2
    Else
      sumTeller = sumTeller + (Math.Log(Observed.Cells(r, 1)) - Math.Log(Computed.Cells(r, 1))) ^ 2
      sumNoemer = sumNoemer + (Math.Log(Observed.Cells(r, 1)) - AvgLogObserved) ^ 2
    End If
  Next
  
  NASH_SUTCLIFFE_FAST = 1 - (sumTeller / sumNoemer)
  Exit Function
  
Errorhandling:
  MsgBox ("Error in function calcNashSutcliffe. " & ErrStr)
  End
  
  
End Function


Public Sub FILTERBASEFLOW(ByRef ValRangeNoHeader As Range, k As Variant, W As Variant, BaseflowCol As Integer, InterFlowCol As Integer)
  'this routine filters the baseflow out of the total discharge
  'it does so by applying the method by prof. Patrick Willems (Leuven University) as implemented in his tool Wetspro
  Dim alpha As Variant, a As Variant, b As Variant, c As Variant, V As Variant
  Dim i As Long, iPar As Long
  
  Dim TotalFlow() As Variant
  Dim InterFlow() As Variant
  Dim BaseFlow() As Variant
  Dim prevTotalFlow As Variant, prevInterFlow As Variant, prevBaseFlow As Variant
  
  ReDim TotalFlow(ValRangeNoHeader.Rows.Count)
  ReDim InterFlow(ValRangeNoHeader.Rows.Count)
  ReDim BaseFlow(ValRangeNoHeader.Rows.Count)
  
  For iPar = 1 To 3
    For i = 1 To ValRangeNoHeader.Count
  
      'retrieve the total, inter and baseflow from the previous timestep
      If i = 1 Then
        prevTotalFlow = 0
        prevInterFlow = 0
        prevBaseFlow = 0
      Else
        prevTotalFlow = TotalFlow(i - 1)
        prevInterFlow = InterFlow(i - 1)
        prevBaseFlow = BaseFlow(i - 1)
      End If
    
      If iPar = 1 Then 'total flow
        TotalFlow(i) = ValRangeNoHeader.Cells(i, 1)
      ElseIf iPar = 2 Then 'interflow
        alpha = Math.Exp(-1 / k)
        V = (1 - W) / W
        a = ((2 + V) * alpha - V) / (2 + V - V * alpha)
        b = 2 / (2 + V - V * alpha)
        c = 0.5 * V
        'curFlow.InterFlow = a * prevFlow.InterFlow + b * (curFlow.Value - alpha * prevFlow.Value)
        InterFlow(i) = a * prevInterFlow + b * (TotalFlow(i) - alpha * prevTotalFlow)

      ElseIf iPar = 3 Then  'baseflow
        alpha = Math.Exp(-1 / k)
        V = (1 - W) / W
        a = ((2 + V) * alpha - V) / (2 + V - V * alpha)
        b = 2 / (2 + V - V * alpha)
        c = 0.5 * V
        BaseFlow(i) = alpha * prevBaseFlow + c * (1 - alpha) * (prevInterFlow + InterFlow(i))
        
      End If
    Next
  Next
  
  ActiveSheet.Cells(ValRangeNoHeader.Cells(1, 1).row - 1, BaseflowCol) = "Baseflow"
  ActiveSheet.Cells(ValRangeNoHeader.Cells(1, 1).row - 1, InterFlowCol) = "Interflow"
  
  For i = 1 To ValRangeNoHeader.Count
    ActiveSheet.Cells(ValRangeNoHeader.Cells(i, 1).row, BaseflowCol) = BaseFlow(i)
    ActiveSheet.Cells(ValRangeNoHeader.Cells(i, 1).row, InterFlowCol) = InterFlow(i)
  Next
  
  
End Sub

Public Function HOOGHOUDT_q(k1 As Variant, k2 As Variant, d As Variant, L As Variant, H As Variant) As Variant
  'k1 = doorlatendheid bovenste laag
  'k2 = doorlatendheid onderste laag
  'Dikte gedraineerde laag
  'L = afstand tussen de drains
  'h = maximale opbolling (m) tussen de drains
  'q = stationaire specifieke afvoer (m/s)
  'let op: K1 en K2 mogen alleen verschillen als de drainagemiddelen exact op de scheidingslaag liggen!
  
  HOOGHOUDT_q = (8 * k2 * d * H + 4 * k1 * H ^ 2) / L ^ 2
  
End Function

Public Function HOOGHOUDT_L(k1 As Variant, k2 As Variant, d As Variant, Q As Variant, H As Variant) As Variant
  'k1 = doorlatendheid bovenste laag
  'k2 = doorlatendheid onderste laag
  'Dikte gedraineerde laag
  'L = afstand tussen de drains
  'h = maximale opbolling (m) tussen de drains
  'q = stationaire specifieke afvoer (m/s)
  'let op: K1 en K2 mogen alleen verschillen als de drainagemiddelen exact op de scheidingslaag liggen!
  
  HOOGHOUDT_L = Sqr((8 * k2 * d * H + 4 * k1 * H ^ 2) / Q)
  
End Function

Public Function YZConveyanceManningSegmentedByDepth(YRange As Range, ZRange As Range, ManningValue As Variant, ResultsRange As Range, DepthColIdx As Integer, KColIdx As Integer)
    Dim row As Integer
    Dim k As Variant, Depth As Variant
    For row = 1 To ResultsRange.Rows.Count
        Depth = ResultsRange.Cells(row, DepthColIdx)
        If Depth > 0 Then
            k = YZConveyanceManningSegmented(YRange, ZRange, ManningValue, Depth)
            ResultsRange.Cells(row, KColIdx) = k
        Else
            ResultsRange.Cells(row, KColIdx) = 0
        End If
    Next
End Function

Function subRange(r As Range, startPos As Integer, endPos As Integer) As Range
    Set subRange = r.Parent.Range(r.Cells(startPos), r.Cells(endPos))
End Function

Public Function YZConveyanceManningSegmented(ByVal YRange As Range, ByVal ZRange As Range, ManningValue As Variant, Depth As Variant) As Variant
    'this function calculates conveyance K for a given YZ cross section, using the 'segmented' method
    'in this (manning) formula, conveyance K is defined as: 1/n*A*R^(2/3)
    'Q is then defined as: K*S^(1/2) where S = slope: (h1-h2)/length
    Dim row As Integer, MinZ As Variant, Level As Variant
    Dim CurY As Variant, NextY As Variant, CurZ As Variant, NextZ As Variant
    Dim yxsect As Variant
    Dim a() As Variant, P() As Variant, r() As Variant, k() As Variant
    Dim newYRange As Range
    Dim newZRange As Range
    MinZ = 9E+99
    YZConveyanceManningSegmented = 0
    
    'truncate the ranges if the table stops before the end
    Dim lastRow As Integer
    Dim i As Integer
    For i = 1 To YRange.Rows.Count
        If YRange.Cells(i + 1, 1) = "" Then
            lastRow = i
            Exit For
        End If
        lastRow = i
    Next
        
    ReDim a(1 To lastRow)
    ReDim P(1 To lastRow)
    ReDim r(1 To lastRow)
    ReDim k(1 To lastRow)
    
    'first find the lowest point and calculate the waterlevel at the current depth
    For row = 1 To lastRow
        If ZRange.Cells(row, 1) < MinZ Then MinZ = ZRange.Cells(row, 1)
    Next
    Level = MinZ + Depth
    
    'now calculate the hydraulic properties
    For row = 1 To lastRow - 1
        a(row) = 0
        P(row) = 0
        r(row) = 0
        k(row) = 0
        CurY = YRange.Cells(row, 1)
        CurZ = ZRange.Cells(row, 1)
        NextY = YRange.Cells(row + 1, 1)
        NextZ = ZRange.Cells(row + 1, 1)
        If Level > CurZ And Level > NextZ Then
            P(row) = P(row) + PYTHAGORAS(NextZ - CurZ, NextY - CurY)
            a(row) = a(row) + (NextY - CurY) * (Level - Maximum(NextZ, CurZ)) + (NextY - CurY) * Math.Abs(NextZ - CurZ) / 2
        ElseIf Level > CurZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            P(row) = P(row) + PYTHAGORAS(Level - CurZ, yxsect - CurY)
            a(row) = a(row) + (Level - CurZ) * Math.Abs(yxsect - CurY)
        ElseIf Level > NextZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            P(row) = P(row) + PYTHAGORAS(Level - NextZ, NextY - yxsect)
            a(row) = a(row) + (Level - NextZ) * Math.Abs(NextY - yxsect)
        End If
         
        If P(row) > 0 Then
            r(row) = a(row) / P(row)
            'K = 1/n*A*R^(2/3)
            'K(row) = A(row) * r(row) ^ (1 / 6) / ManningValue * Sqr(r(row))
            k(row) = a(row) * r(row) ^ (2 / 3) / ManningValue
        Else
            r(row) = 0
            k(row) = 0
        End If
    Next
    
    For row = 1 To lastRow
        YZConveyanceManningSegmented = YZConveyanceManningSegmented + k(row)
    Next
    
    
End Function

Public Function QManning(k As Double, h1 As Double, h2 As Double, distance As Double) As Double
    'this formula calculates discharge Q based on a given conveyance, h1, h2 and distance between both waterlevel points
    'Q = 1/n* R^(2/3)*S^(1/2)
    QManning = k * Math.Sqr(Math.Abs(h1 - h2) / distance)
End Function

Public Function YZHydraulicPropertiesByDepth(YRange As Range, ZRange As Range, ResultsRange As Range, DepthColIdx As Integer, AColIdx As Integer, PColIdx As Integer, RColIdx As Integer) As Boolean
    'this function calculates hydraulic properties for a given YZ cross section, as a function of given depths:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    On Error GoTo errorhandler
    Dim row As Integer
    Dim a As Variant, P As Variant, r As Variant, Depth As Variant
    For row = 1 To ResultsRange.Rows.Count
        Depth = ResultsRange.Cells(row, DepthColIdx)
        If Depth > 0 Then
            Call YZHydraulicProperties(YRange, ZRange, Depth, a, P, r)
            ResultsRange.Cells(row, AColIdx).value = Format(a, "0.00")
            ResultsRange.Cells(row, PColIdx).value = P
            ResultsRange.Cells(row, RColIdx).value = r
        End If
    Next
    YZHydraulicPropertiesByDepth = True
errorhandler:
    MsgBox ("Error")
    
End Function

Public Function XYZTOYZPROFILE(XRange As Range, YRange As Range, ZRange As Range, ResultsRow As Integer, ResultsCol As Integer) As Boolean
    Dim r As Integer, c As Integer
    Dim r2 As Integer, c2 As Integer
    Dim x As Double, y As Double, Z As Double
    Dim myY As Double, myZ As Double
    
    r2 = 0
    c2 = 0
    myY = 0
    myZ = ZRange.Cells(1, 1)
    ActiveSheet.Cells(ResultsRow + r2, ResultsCol) = myY
    ActiveSheet.Cells(ResultsRow + r2, ResultsCol + 1) = myZ
    
    For r = 2 To XRange.Rows.Count
        x = XRange.Cells(r, 1)
        y = YRange.Cells(r, 1)
        Z = ZRange.Cells(r, 1)
                     
        myY = myY + PYTHAGORAS(XRange.Cells(r, 1) - XRange.Cells(r - 1, 1), YRange.Cells(r, 1) - YRange.Cells(r - 1, 1))
        myZ = ZRange.Cells(r, 1)
        
        r2 = r2 + 1
        c2 = c2 + 1
        ActiveSheet.Cells(ResultsRow + r2, ResultsCol) = myY
        ActiveSheet.Cells(ResultsRow + r2, ResultsCol + 1) = myZ
        
    Next
    XYZTOYZPROFILE = True
End Function


Public Function YZWettedArea(YRange As Range, ZRange As Range, Depth As Variant) As Variant
    'this function calculates the wetted area  for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Call YZHydraulicProperties(YRange, ZRange, Depth, a, P, r)
    YZWettedArea = a
End Function

Public Function TabulatedWettedAreaByDepth(YRange As Range, ZRange As Range, Depth As Variant) As Variant
    'this function calculates the wetted area  for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Call TabulatedHydraulicProperties(YRange, ZRange, Depth, a, P, r)
    TabulatedWettedAreaByDepth = a
End Function


Public Function TabulatedWettedAreaByWaterlevel(waterLevel As Variant, ZRange As Range, WRange As Range, ProfileVerticalShift As Variant) As Variant
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Dim Depth As Variant
    Depth = waterLevel - ZRange.Cells(1, 1) - ProfileVerticalShift
    Call TabulatedHydraulicProperties(ZRange, WRange, Depth, a, P, r)
    TabulatedWettedAreaByWaterlevel = a
End Function


Public Function YZHydraulicRadius(YRange As Range, ZRange As Range, Depth As Variant) As Variant
    'this function calculates the hydraulic radius for a given YZ cross section and a given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim a As Variant
    Dim P As Variant
    Dim r As Variant
    Call YZHydraulicProperties(YRange, ZRange, Depth, a, P, r)
    YZHydraulicRadius = r
End Function

Public Function YZWettedAreaByLevel(YRange As Range, ZRange As Range, Level As Variant) As Double
    Dim row As Integer
    Dim CurY As Double
    Dim NextY As Double
    Dim CurZ As Double
    Dim NextZ As Double
    Dim a As Double
    Dim yxsect As Double
    
    'make sure we don't move past the last row that contains data
    Dim lastRow As Integer
    For row = 1 To YRange.Rows.Count - 1
        If YRange.Cells(row + 1, 1) = "" Then
            lastRow = row
            Exit For
        End If
    Next
    If lastRow = 0 Then lastRow = YRange.Rows.Count
    
    'calculate the wetted area for each Z value
    For row = 1 To lastRow - 1
        CurY = YRange.Cells(row, 1)
        CurZ = ZRange.Cells(row, 1)
        NextY = YRange.Cells(row + 1, 1)
        NextZ = ZRange.Cells(row + 1, 1)
        If Level >= CurZ And Level >= NextZ Then
            a = a + (NextY - CurY) * (Level - Maximum(NextZ, CurZ)) + (NextY - CurY) * Math.Abs(NextZ - CurZ) / 2
        ElseIf Level > CurZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            a = a + (Level - CurZ) * Math.Abs(yxsect - CurY)
        ElseIf Level > NextZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            a = a + (Level - NextZ) * Math.Abs(NextY - yxsect)
        End If
     Next
     YZWettedAreaByLevel = a
End Function


Public Function YZHydraulicProperties(YRange As Range, ZRange As Range, Depth As Variant, ByRef a As Variant, ByRef P As Variant, ByRef r As Variant) As Boolean
    'this function calculates hydraulic properties for a given YZ cross section, as a function of given depths:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    
    Dim row As Integer, MinZ As Variant, Level As Variant
    Dim CurY As Variant, NextY As Variant, CurZ As Variant, NextZ As Variant
    Dim yxsect As Variant
    MinZ = 9E+99
    a = 0
    P = 0
    r = 0
    
    'first find the lowest point and calculate the waterlevel at the current depth
    For row = 1 To YRange.Rows.Count
        If ZRange.Cells(row, 1) < MinZ Then MinZ = ZRange.Cells(row, 1)
    Next
    Level = MinZ + Depth
    
    'now calculate the hydraulic properties
    For row = 1 To YRange.Rows.Count - 1
        CurY = YRange.Cells(row, 1)
        CurZ = ZRange.Cells(row, 1)
        NextY = YRange.Cells(row + 1, 1)
        NextZ = ZRange.Cells(row + 1, 1)
        If Level > CurZ And Level > NextZ Then
            P = P + PYTHAGORAS(NextZ - CurZ, NextY - CurY)
            a = a + (NextY - CurY) * (Level - Maximum(NextZ, CurZ)) + (NextY - CurY) * Math.Abs(NextZ - CurZ) / 2
        ElseIf Level > CurZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            P = P + PYTHAGORAS(Level - CurZ, yxsect - CurY)
            a = a + (Level - CurZ) * Math.Abs(yxsect - CurY)
        ElseIf Level > NextZ Then
            yxsect = Interpolate(CurZ, CurY, NextZ, NextY, Level)
            P = P + PYTHAGORAS(Level - NextZ, NextY - yxsect)
            a = a + (Level - NextZ) * Math.Abs(NextY - yxsect)
        End If
     Next
     If P > 0 Then
        r = a / P
     Else
        r = 0
     End If
    YZHydraulicProperties = True
End Function

Public Function TabulatedHydraulicProperties(ZRange As Range, WRange As Range, Depth As Variant, ByRef a As Variant, ByRef P As Variant, ByRef r As Variant) As Boolean
    'this function calculates hydraulic properties for a given tabulated cross section, as a function of given depth:
    'A (wetted area)
    'P (wetted perimeter)
    'R (hydraulic radius)
    Dim BedLevel As Variant
    Dim PrevZ As Variant
    Dim CurZ As Variant
    Dim prevW As Variant
    Dim CurW As Variant
    Dim PrevD As Variant
    Dim CurD As Variant
    Dim WidthAtDepth As Variant
    Dim row As Integer
    
    'makes sure we don't run further than the last row that contains data
    Dim lastRow As Integer
    For row = 1 To ZRange.Rows.Count - 1
        If ZRange.Cells(row + 1, 1) = "" Then
            lastRow = row
            Exit For
        End If
    Next
    If lastRow = 0 Then lastRow = ZRange.Rows.Count
        
    BedLevel = ZRange.Cells(1, 1)
    For row = 2 To lastRow
        PrevZ = ZRange.Cells(row - 1, 1)
        CurZ = ZRange.Cells(row, 1)
        prevW = WRange.Cells(row - 1, 1)
        CurW = WRange.Cells(row, 1)
        PrevD = PrevZ - BedLevel
        CurD = CurZ - BedLevel
                                
        'zolang de waterdiepte groter is dan de 'diepte' van het huidige segment moeten we doorgaan
        If Depth >= CurD Then
            'add the entire segment
            a = a + (CurZ - PrevZ) * (Application.WorksheetFunction.Min(CurW, prevW) + Math.Abs(CurW - prevW) / 2)
            'new in v1.54: make sure the wetted perimeter does not grow when width = 0
            If CurW > 0 Or prevW > 0 Then P = P + 2 * PYTHAGORAS((CurZ - PrevZ), Math.Abs(CurW - prevW) / 2)
            
            'when we're at the highest point, also add the part that exceeds our profile
            If row = lastRow Then
                a = a + (Depth - CurD) * CurW
                'here we don't do anything with the perimeter
            End If
            
        ElseIf Depth >= PrevD Then
            'add part of the segment
            WidthAtDepth = Interpolate(CurD, CurW, PrevD, prevW, Depth)
            a = a + (Depth - PrevD) * (Application.WorksheetFunction.Min(WidthAtDepth, prevW) + Math.Abs(WidthAtDepth - prevW) / 2)
            
            'new in v1.54: make sure the wetted perimeter does not grow when width = 0
            If WidthAtDepth > 0 Or prevW > 0 Then P = P + 2 * PYTHAGORAS((Depth - PrevD), Math.Abs(WidthAtDepth - prevW) / 2)
            Exit For
        End If
    Next
        
    If P > 0 Then r = a / P
    TabulatedHydraulicProperties = True
End Function

Public Function ExtraResistanceKSI(EntranceLoss As Double, FrictionLoss As Double, ExitLoss As Double, Astruc As Double) As Double
    'this formula has been derived from the Abutment Bridge formula, substituted in the extra resistance forula dH = ksi *Q*|Q|
    'abutment: Q = mu * A * sqrt(2g*dH)
    'dH = ksi * Q|Q|
    'substitution:
    'KSI = dH/(mu*A*sqrt(2gdH))^2
    'where mu = 1/sqr(entrance + friction + exit)
    ExtraResistanceKSI = (EntranceLoss + FrictionLoss + ExitLoss) / (Astruc ^ 2 * 2 * 9.81)
End Function

Public Function NELENSCHUURMANSFYSIEKVOORKOMENTONBW(GRIDCODE As Integer) As Integer
  'resultaat: 0= openwater, 1 = akkerbouw, 2 = akkerbouw hoogwaardig, 3 = gras, 4 = natuur, 5 = stedelijk
  
  Select Case GRIDCODE
  Case Is = 1 'Dak
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 2 'Zand (constatering: zand in natuur is al afgedekt (Duin, Heide), dus we nemen aan dat het hier stedelijk zand betreft)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 3 'Half verhard
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 4 'Erf (boerenerf heeft naar ons idee geen recht op T=100 norm. Dus T=10 gehanteerd)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 5 'Gesloten verharding
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 6 'Onverhard
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 7 'Open verharding
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 5
  Case Is = 8 'Groenvoorziening. Aanname: de T=100 geldt alleen boven de dorpelhoogte, dus alle plantsoenen zijn naar ons idee T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 9 'Naaldbos. Dit interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 10 'Gras
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 11 'Grasland. Dit interpreteren wij agrarisch gras, dus T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 12 'Struiken. Wij interpreteren dit als natuur, dus geen norm.
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 13 'Natuurterreinen
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 14 'Boomteelt. Interpreteren wij als hoogwaardige akkerbouw (T=50)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 2
  Case Is = 15 'Duin
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 16 'Heide
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 17 'Bouwland. Wij interpreteren dit als akkerbouw, dus T=25
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 1
  Case Is = 18 'Houtwal. Omringt normaalgesproken bouwland, dus we hanteren T=25
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 1
  Case Is = 19 'Loofbos. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 20 'Rietland. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 21 'Fruitteelt. Interpreteren wij als hoogwaardige land- en tuinbouw
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 2
  Case Is = 22 'Gemengd bos. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 23 'Braakland. Is vrijwel altijd stedelijk, maar niet verhard. Wij interpreteren dit als T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 26 'Moeras. Interpreteren wij als natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 27 'Kwelder. Natuur
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 4
  Case Is = 28 'Waterberm. We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  Case Is = 29 'Water. We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  Case Is = 30 'Overig. Kan vanalles zijn. We hanteren T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 253 'Onbekend. Kan vanalles zijn. We hanteren T=10
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 3
  Case Is = 254 'Water.We hanteren de norm voor water (dus geen norm)
    NELENSCHUURMANSFYSIEKVOORKOMENTONBW = 0
  End Select
  
End Function

Public Function LGN5TONBW(LGN5Code As Integer) As Integer
  'resultaat: 0= openwater, 1 = akkerbouw, 2 = akkerbouw hoogwaardig, 3 = gras, 4 = natuur, 5 = stedelijk
  
  Select Case LGN5Code
  Case Is = 1
    LGN5TONBW = 3
  Case Is = 2
    LGN5TONBW = 1
  Case Is = 3
    LGN5TONBW = 1
  Case Is = 4
    LGN5TONBW = 1
  Case Is = 5
    LGN5TONBW = 1
  Case Is = 6
    LGN5TONBW = 1
  Case Is = 8
    LGN5TONBW = 2
  Case Is = 9
    LGN5TONBW = 2
  Case Is = 10
    LGN5TONBW = 2
  Case Is = 11
    LGN5TONBW = 3
  Case Is = 12
    LGN5TONBW = 3
  Case Is = 16
    LGN5TONBW = 0
  Case Is = 17
    LGN5TONBW = 0
  Case Is = 18
    LGN5TONBW = 5
  Case Is = 19
    LGN5TONBW = 5
  Case Is = 20
    LGN5TONBW = 3
  Case Is = 21
    LGN5TONBW = 3
  Case Is = 22
    LGN5TONBW = 5
  Case Is = 23
    LGN5TONBW = 3
  Case Is = 24
    LGN5TONBW = 3
  Case Is = 25
    LGN5TONBW = 5
  Case Is = 26
    LGN5TONBW = 5
  Case Is = 30
    LGN5TONBW = 0
  Case Is = 35
    LGN5TONBW = 0
  Case Is = 36
    LGN5TONBW = 0
  Case Is = 37
    LGN5TONBW = 0
  Case Is = 38
    LGN5TONBW = 0
  Case Is = 39
    LGN5TONBW = 0
  Case Is = 40
    LGN5TONBW = 0
  Case Is = 41
    LGN5TONBW = 0
  Case Is = 42
    LGN5TONBW = 0
  Case Is = 43
    LGN5TONBW = 0
  Case Is = 45
    LGN5TONBW = 0
  Case Is = 46
    LGN5TONBW = 0
  End Select
  
End Function

Public Function LGN2SOBEK(LGNCODE As Long) As Long

'1 = grass
'2 = corn
'3 = potatoes
'4 = sugarbeet
'5 = grain
'6 = miscellaneous
'7 = non-arable land
'8 = greenhouse area
'9 = orchard
'10 = bulbous plants
'11 = foliage forest
'12 = pine forest
'13 = nature
'14 = fallow
'15 = vegetables
'16 = flowers

'zelf toegevoegd:
'17 = water
'18 = verhard

Select Case LGNCODE
  Case Is = 1 'gras
    LGN2SOBEK = 1
  Case Is = 2 'mais
    LGN2SOBEK = 2
  Case Is = 3 'aardappelen
    LGN2SOBEK = 3
  Case Is = 4 'suikerbiet
    LGN2SOBEK = 4
  Case Is = 5 'graan
    LGN2SOBEK = 5
  Case Is = 6 'overige landbouwgewassen
    LGN2SOBEK = 6
  Case Is = 8 'kassen
    LGN2SOBEK = 8
  Case Is = 9 'boomgaard
    LGN2SOBEK = 9
  Case Is = 10 'bollenteelt
    LGN2SOBEK = 10
  Case Is = 11 'loofbos
    LGN2SOBEK = 11
  Case Is = 12 'naaldbos
    LGN2SOBEK = 12
  Case Is = 16 'zoet water
    LGN2SOBEK = 17
  Case Is = 17 'zout water
    LGN2SOBEK = 17
  Case Is = 18 'stedelijk bebouwd
    LGN2SOBEK = 18
  Case Is = 19 'bebouwd buitengebied
    LGN2SOBEK = 18
  Case Is = 20 'loofbos in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 21 'naaldbos in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 22 'bos met dichte bebouwing
    LGN2SOBEK = 18
  Case Is = 23 'gras in bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 24 'kale grond in bebouwd buitengebied
    LGN2SOBEK = 1
  Case Is = 25 'hoofdwegen en spoorwegen
    LGN2SOBEK = 18
  Case Is = 26 'bebouwing in agrarisch gebied
    LGN2SOBEK = 18
  Case Is = 28 'Gras in secundair bebouwd gebied
    LGN2SOBEK = 1
  Case Is = 30 'kwelders
    LGN2SOBEK = 13
  Case Is = 35 'open stuifzand
    LGN2SOBEK = 13
  Case Is = 36 'heide
    LGN2SOBEK = 13
  Case Is = 37 'matig vergraste heide
    LGN2SOBEK = 13
  Case Is = 38 'sterk vergraste heide
    LGN2SOBEK = 13
  Case Is = 39 'hoogveen
    LGN2SOBEK = 13
  Case Is = 40 'bos in hoogveen
    LGN2SOBEK = 13
  Case Is = 41 'overige moerasvegetatie
    LGN2SOBEK = 13
  Case Is = 42 'rietvegetatie
    LGN2SOBEK = 13
  Case Is = 43 'bos in moerasgebied
    LGN2SOBEK = 13
  Case Is = 45 'overig open begroeid natuurgebied
    LGN2SOBEK = 13
  Case Is = 46 'kale grond in natuurgebied
    LGN2SOBEK = 13
  Case Is = 61 'boomkwekerijen
    LGN2SOBEK = 9
  Case Is = 61 'fruitkwekerijen
    LGN2SOBEK = 9
End Select

End Function

Public Function ERNSTRecord(id As String, a1 As Variant, a2 As Variant, a3 As Variant, a4 As Variant, lv1 As Variant, lv2 As Variant, lv3 As Variant, ainf As Variant) As String
  ERNSTRecord = "ERNS id '" & id & "' nm '" & id & "' cvi " & ainf & " cvo " & a1 & " " & a2 & " " & a3 & " " & a4 & " lv " & lv1 & " " & lv2 & " " & lv3 & " cvs 1 erns"
End Function

Public Function Bod2ZandKleiVeen(bc As String) As String
        Dim CapSimCode As Long
        CapSimCode = BOD2CAPSIM(bc)
        Select Case CapSimCode
                Case Is = 101
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 102
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 103
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 104
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 105
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 106
                        Bod2ZandKleiVeen = "VEEN"
                Case Is = 107
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 108
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 109
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 110
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 111
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 112
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 113
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 114
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 115
                        Bod2ZandKleiVeen = "ZAND"
                Case Is = 116
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 117
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 118
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 119
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 120
                        Bod2ZandKleiVeen = "KLEI"
                Case Is = 121
                        Bod2ZandKleiVeen = "KLEI"
        End Select
End Function

Public Function GHGGLG2GT(GHGmMV As Variant, GLGmMV As Variant) As String
    If GHGmMV < 0.2 Or GLGmMV < 0.5 Then
        GHGGLG2GT = "I"
    
    End If
    
End Function

Public Function BOD2CAPSIM(bc As String) As Long
'converteert bodemtypes uit de Bodemkaart Nederland naar het corresponderende CAPSIM bodemnummer in SOBEK

'knip de grondwatertrap eraf!
bc = ParseString(bc, "-")
If bc = "|a GROEVE" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|b AFGRAV" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|c OPHOOG" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|d EGAL" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|e VERWERK" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|f TERP" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|g MOERAS" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|g WATER" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|h BEBOUW" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|h DIJK" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|i BOVLAND" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "|j MYNSTRT" Then BOD2CAPSIM = 113 'willekeurig bedenksel
If bc = "AAK" Then BOD2CAPSIM = 119 'afgegraven klei
If bc = "AAKp" Then BOD2CAPSIM = 119
If bc = "AAP" Then BOD2CAPSIM = 105
If bc = "ABk" Then BOD2CAPSIM = 119
If bc = "ABkt" Then BOD2CAPSIM = 119
If bc = "ABl" Then BOD2CAPSIM = 121
If bc = "ABv" Then BOD2CAPSIM = 105
If bc = "ABvg" Then BOD2CAPSIM = 105
If bc = "ABvt" Then BOD2CAPSIM = 105
If bc = "ABvx" Then BOD2CAPSIM = 105
If bc = "ABz" Then BOD2CAPSIM = 113
If bc = "ABzt" Then BOD2CAPSIM = 113
If bc = "AD" Then BOD2CAPSIM = 113      'Duin- en kweldergronden
If bc = "AEk9" Then BOD2CAPSIM = 116
If bc = "AEm5" Then BOD2CAPSIM = 115
If bc = "AEm8" Then BOD2CAPSIM = 116
If bc = "AEm9" Then BOD2CAPSIM = 116
If bc = "AEm9A" Then BOD2CAPSIM = 116
If bc = "AEp6A" Then BOD2CAPSIM = 116
If bc = "AEp7A" Then BOD2CAPSIM = 116
If bc = "AFk" Then BOD2CAPSIM = 113
If bc = "AGm9C" Then BOD2CAPSIM = 117   'hollebollige, gemoerde zeekleigronden
If bc = "AFz" Then BOD2CAPSIM = 113
If bc = "Aha" Then BOD2CAPSIM = 121     'glauconiethellingronden
If bc = "AHa" Then BOD2CAPSIM = 121     'glauconiethellingronden
If bc = "AHc" Then BOD2CAPSIM = 121
If bc = "AHk" Then BOD2CAPSIM = 121
If bc = "AHl" Then BOD2CAPSIM = 121
If bc = "Ahs" Then BOD2CAPSIM = 121
If bc = "AHs" Then BOD2CAPSIM = 121
If bc = "AHt" Then BOD2CAPSIM = 121
If bc = "AHv" Then BOD2CAPSIM = 121
If bc = "AHz" Then BOD2CAPSIM = 121
If bc = "AK" Then BOD2CAPSIM = 119
If bc = "AKp" Then BOD2CAPSIM = 119
If bc = "ALu" Then BOD2CAPSIM = 116
If bc = "AM" Then BOD2CAPSIM = 119
If bc = "AMm" Then BOD2CAPSIM = 115
If bc = "AO" Then BOD2CAPSIM = 119
If bc = "AOg" Then BOD2CAPSIM = 119
If bc = "AOp" Then BOD2CAPSIM = 119
If bc = "AOv" Then BOD2CAPSIM = 119
If bc = "AP" Then BOD2CAPSIM = 101
If bc = "App" Then BOD2CAPSIM = 102
If bc = "AQ" Then BOD2CAPSIM = 107
If bc = "AR" Then BOD2CAPSIM = 119
If bc = "AS" Then BOD2CAPSIM = 107
If bc = "aVc" Then BOD2CAPSIM = 101
If bc = "AVk" Then BOD2CAPSIM = 105
If bc = "AVo" Then BOD2CAPSIM = 101
If bc = "aVp" Then BOD2CAPSIM = 102
If bc = "aVpg" Then BOD2CAPSIM = 102
If bc = "aVpx" Then BOD2CAPSIM = 102
If bc = "aVs" Then BOD2CAPSIM = 101
If bc = "aVz" Then BOD2CAPSIM = 102
If bc = "aVzt" Then BOD2CAPSIM = 102
If bc = "aVzx" Then BOD2CAPSIM = 102
If bc = "AWg" Then BOD2CAPSIM = 116
If bc = "AWo" Then BOD2CAPSIM = 106
If bc = "AWv" Then BOD2CAPSIM = 106
If bc = "AZ1" Then BOD2CAPSIM = 114
If bc = "AZW0A" Then BOD2CAPSIM = 107
If bc = "AZW0Al" Then BOD2CAPSIM = 107
If bc = "AZW0Av" Then BOD2CAPSIM = 107
If bc = "AZW1A" Then BOD2CAPSIM = 119
If bc = "AZW1Ar" Then BOD2CAPSIM = 119
If bc = "AZW1Aw" Then BOD2CAPSIM = 119
If bc = "AZW5A" Then BOD2CAPSIM = 119
If bc = "AZW6A" Then BOD2CAPSIM = 119
If bc = "AZW6Al" Then BOD2CAPSIM = 116
If bc = "AZW6Alv" Then BOD2CAPSIM = 118
If bc = "AZW7Al" Then BOD2CAPSIM = 116
If bc = "AZW7A" Then BOD2CAPSIM = 116
If bc = "AZW7Alw" Then BOD2CAPSIM = 116
If bc = "AZW7Alwp" Then BOD2CAPSIM = 119
If bc = "AZW8A" Then BOD2CAPSIM = 116
If bc = "AZW8Al" Then BOD2CAPSIM = 116
If bc = "AZW8Alw" Then BOD2CAPSIM = 116
If bc = "bEZ21" Then BOD2CAPSIM = 112
If bc = "bEZ21g" Then BOD2CAPSIM = 112
If bc = "bEZ21x" Then BOD2CAPSIM = 112
If bc = "bEZ23" Then BOD2CAPSIM = 112
If bc = "bEZ23g" Then BOD2CAPSIM = 112
If bc = "bEZ23t" Then BOD2CAPSIM = 112
If bc = "bEZ23x" Then BOD2CAPSIM = 112
If bc = "bEZ30" Then BOD2CAPSIM = 112
If bc = "bEZ30x" Then BOD2CAPSIM = 112
If bc = "bgMn15C" Then BOD2CAPSIM = 115
If bc = "bgMn25C" Then BOD2CAPSIM = 115
If bc = "bgMn53C" Then BOD2CAPSIM = 117
If bc = "BKd25" Then BOD2CAPSIM = 115
If bc = "BKd25x" Then BOD2CAPSIM = 115
If bc = "BKd26" Then BOD2CAPSIM = 115
If bc = "BKh25" Then BOD2CAPSIM = 115
If bc = "BKh25x" Then BOD2CAPSIM = 115
If bc = "BKh26" Then BOD2CAPSIM = 115
If bc = "BKh26x" Then BOD2CAPSIM = 115
If bc = "BLb6" Then BOD2CAPSIM = 121
If bc = "BLb6g" Then BOD2CAPSIM = 121
If bc = "BLb6k" Then BOD2CAPSIM = 121
If bc = "BLb6s" Then BOD2CAPSIM = 121
If bc = "BLd5" Then BOD2CAPSIM = 121
If bc = "BLd5g" Then BOD2CAPSIM = 121
If bc = "BLd5t" Then BOD2CAPSIM = 121
If bc = "BLd6" Then BOD2CAPSIM = 121
If bc = "BLd6m" Then BOD2CAPSIM = 121
If bc = "BLh5" Then BOD2CAPSIM = 121
If bc = "BLh5m" Then BOD2CAPSIM = 121
If bc = "BLh6" Then BOD2CAPSIM = 121
If bc = "BLh6g" Then BOD2CAPSIM = 121
If bc = "BLh6m" Then BOD2CAPSIM = 121
If bc = "BLh6s" Then BOD2CAPSIM = 121
If bc = "BLn5" Then BOD2CAPSIM = 121
If bc = "BLn5m" Then BOD2CAPSIM = 121
If bc = "BLn5t" Then BOD2CAPSIM = 121
If bc = "BLn6" Then BOD2CAPSIM = 121
If bc = "BLn6g" Then BOD2CAPSIM = 121
If bc = "BLn6m" Then BOD2CAPSIM = 121
If bc = "BLn6s" Then BOD2CAPSIM = 121
If bc = "bMn15A" Then BOD2CAPSIM = 115
If bc = "bMn15C" Then BOD2CAPSIM = 115
If bc = "bMn25A" Then BOD2CAPSIM = 115
If bc = "bMn25C" Then BOD2CAPSIM = 115
If bc = "bMn35A" Then BOD2CAPSIM = 116
If bc = "bMn45A" Then BOD2CAPSIM = 117
If bc = "bMn56Cp" Then BOD2CAPSIM = 119
If bc = "bMn85C" Then BOD2CAPSIM = 116
If bc = "bMn86C" Then BOD2CAPSIM = 117
If bc = "bRn46C" Then BOD2CAPSIM = 117
If bc = "BZd23" Then BOD2CAPSIM = 113
If bc = "BZd24" Then BOD2CAPSIM = 113
If bc = "cHd21" Then BOD2CAPSIM = 108
If bc = "cHd21g" Then BOD2CAPSIM = 110
If bc = "cHd21x" Then BOD2CAPSIM = 111
If bc = "cHd23" Then BOD2CAPSIM = 113
If bc = "cHd23x" Then BOD2CAPSIM = 111
If bc = "cHd30" Then BOD2CAPSIM = 114
If bc = "cHn21" Then BOD2CAPSIM = 109
If bc = "cHn21g" Then BOD2CAPSIM = 110
If bc = "cHn21t" Then BOD2CAPSIM = 111
If bc = "cHn21w" Then BOD2CAPSIM = 111
If bc = "cHn21x" Then BOD2CAPSIM = 111
If bc = "cHn23" Then BOD2CAPSIM = 113
If bc = "cHn23g" Then BOD2CAPSIM = 110
If bc = "cHn23t" Then BOD2CAPSIM = 111
If bc = "cHn23wx" Then BOD2CAPSIM = 111
If bc = "cHn23x" Then BOD2CAPSIM = 111
If bc = "cHn30" Then BOD2CAPSIM = 114
If bc = "cHn30g" Then BOD2CAPSIM = 114
If bc = "cY21" Then BOD2CAPSIM = 109
If bc = "cY21g" Then BOD2CAPSIM = 110
If bc = "cY21x" Then BOD2CAPSIM = 111
If bc = "cY23" Then BOD2CAPSIM = 113
If bc = "cY23g" Then BOD2CAPSIM = 113
If bc = "cY23x" Then BOD2CAPSIM = 111
If bc = "cY30" Then BOD2CAPSIM = 114
If bc = "cY30g" Then BOD2CAPSIM = 114
If bc = "cZd21" Then BOD2CAPSIM = 108
If bc = "cZd21g" Then BOD2CAPSIM = 110
If bc = "cZd23" Then BOD2CAPSIM = 113
If bc = "cZd30" Then BOD2CAPSIM = 114
If bc = "dgMn58Cv" Then BOD2CAPSIM = 117
If bc = "dgMn83C" Then BOD2CAPSIM = 117
If bc = "dgMn88Cv" Then BOD2CAPSIM = 117
If bc = "dhVb" Then BOD2CAPSIM = 101
If bc = "dhVk" Then BOD2CAPSIM = 106
If bc = "dhVr" Then BOD2CAPSIM = 101
If bc = "dkVc" Then BOD2CAPSIM = 103
If bc = "dMn86C" Then BOD2CAPSIM = 117
If bc = "dMv41C" Then BOD2CAPSIM = 118
If bc = "dMv61C" Then BOD2CAPSIM = 118
If bc = "dpVc" Then BOD2CAPSIM = 103
If bc = "dVc" Then BOD2CAPSIM = 101
If bc = "dVd" Then BOD2CAPSIM = 101
If bc = "dVk" Then BOD2CAPSIM = 106
If bc = "dVr" Then BOD2CAPSIM = 101
If bc = "dWo" Then BOD2CAPSIM = 106
If bc = "dWol" Then BOD2CAPSIM = 106
If bc = "EK19" Then BOD2CAPSIM = 115
If bc = "EK19p" Then BOD2CAPSIM = 119
If bc = "EK19x" Then BOD2CAPSIM = 115
If bc = "EK76" Then BOD2CAPSIM = 117
If bc = "EK79" Then BOD2CAPSIM = 116
If bc = "EK79v" Then BOD2CAPSIM = 116
If bc = "EK79w" Then BOD2CAPSIM = 116
If bc = "EL5" Then BOD2CAPSIM = 121
If bc = "eMn12Ap" Then BOD2CAPSIM = 119
If bc = "eMn15A" Then BOD2CAPSIM = 115
If bc = "eMn15Ap" Then BOD2CAPSIM = 119
If bc = "eMn22A" Then BOD2CAPSIM = 119
If bc = "eMn22Ap" Then BOD2CAPSIM = 119
If bc = "eMn25A" Then BOD2CAPSIM = 115
If bc = "eMn25Ap" Then BOD2CAPSIM = 119
If bc = "eMn25Av" Then BOD2CAPSIM = 118
If bc = "eMn35A" Then BOD2CAPSIM = 116
If bc = "eMn35Ap" Then BOD2CAPSIM = 119
If bc = "eMn35Av" Then BOD2CAPSIM = 118
If bc = "eMn35Awp" Then BOD2CAPSIM = 119
If bc = "eMn45A" Then BOD2CAPSIM = 117
If bc = "eMn45Ap" Then BOD2CAPSIM = 117
If bc = "eMn45Av" Then BOD2CAPSIM = 118
If bc = "eMn52Cg" Then BOD2CAPSIM = 119
If bc = "eMn52Cp" Then BOD2CAPSIM = 119
If bc = "eMn52Cwp" Then BOD2CAPSIM = 119
If bc = "eMn56Av" Then BOD2CAPSIM = 118
If bc = "eMn82A" Then BOD2CAPSIM = 119
If bc = "eMn82Ap" Then BOD2CAPSIM = 119
If bc = "eMn82C" Then BOD2CAPSIM = 119
If bc = "eMn82Cp" Then BOD2CAPSIM = 119
If bc = "eMn86A" Then BOD2CAPSIM = 117
If bc = "eMn86Av" Then BOD2CAPSIM = 118
If bc = "eMn86C" Then BOD2CAPSIM = 117
If bc = "eMn86Cv" Then BOD2CAPSIM = 118
If bc = "eMn86Cw" Then BOD2CAPSIM = 117
If bc = "eMo20A" Then BOD2CAPSIM = 119
If bc = "eMo20Ap" Then BOD2CAPSIM = 119
If bc = "eMo80A" Then BOD2CAPSIM = 116
If bc = "eMo80Ap" Then BOD2CAPSIM = 119
If bc = "eMo80C" Then BOD2CAPSIM = 116
If bc = "eMo80Cv" Then BOD2CAPSIM = 118
If bc = "eMOb72" Then BOD2CAPSIM = 119
If bc = "eMOb75" Then BOD2CAPSIM = 116
If bc = "eMOo05" Then BOD2CAPSIM = 115
If bc = "eMv41C" Then BOD2CAPSIM = 118
If bc = "eMv51A" Then BOD2CAPSIM = 118
If bc = "eMv61C" Then BOD2CAPSIM = 118
If bc = "eMv61Cp" Then BOD2CAPSIM = 118
If bc = "eMv81A" Then BOD2CAPSIM = 118
If bc = "eMv81Ap" Then BOD2CAPSIM = 118
If bc = "epMn55A" Then BOD2CAPSIM = 115
If bc = "epMn85A" Then BOD2CAPSIM = 116
If bc = "epMo50" Then BOD2CAPSIM = 115
If bc = "epMo80" Then BOD2CAPSIM = 116
If bc = "epMv81" Then BOD2CAPSIM = 118
If bc = "epRn56" Then BOD2CAPSIM = 117
If bc = "epRn59" Then BOD2CAPSIM = 119
If bc = "epRn86" Then BOD2CAPSIM = 117
If bc = "eRn45A" Then BOD2CAPSIM = 117
If bc = "eRn46A" Then BOD2CAPSIM = 117
If bc = "eRn46Av" Then BOD2CAPSIM = 118
If bc = "eRn47C" Then BOD2CAPSIM = 117
If bc = "eRn52A" Then BOD2CAPSIM = 119
If bc = "eRn66A" Then BOD2CAPSIM = 117
If bc = "eRn66Av" Then BOD2CAPSIM = 118
If bc = "eRn82A" Then BOD2CAPSIM = 119
If bc = "eRn94C" Then BOD2CAPSIM = 117
If bc = "eRn95A" Then BOD2CAPSIM = 116
If bc = "eRn95Av" Then BOD2CAPSIM = 118
If bc = "eRo40A" Then BOD2CAPSIM = 117
If bc = "eRv01A" Then BOD2CAPSIM = 118
If bc = "eRv01C" Then BOD2CAPSIM = 118
If bc = "EZ50A" Then BOD2CAPSIM = 107
If bc = "EZ50Av" Then BOD2CAPSIM = 107
If bc = "EZg21" Then BOD2CAPSIM = 112
If bc = "EZg21g" Then BOD2CAPSIM = 112
If bc = "EZg21v" Then BOD2CAPSIM = 112
If bc = "EZg21w" Then BOD2CAPSIM = 112
If bc = "EZg23" Then BOD2CAPSIM = 112
If bc = "EZg23g" Then BOD2CAPSIM = 112
If bc = "EZg23t" Then BOD2CAPSIM = 112
If bc = "EZg23tw" Then BOD2CAPSIM = 112
If bc = "EZg23w" Then BOD2CAPSIM = 112
If bc = "EZg23wg" Then BOD2CAPSIM = 112
If bc = "EZg23wt" Then BOD2CAPSIM = 112
If bc = "EZg30" Then BOD2CAPSIM = 112
If bc = "EZg30g" Then BOD2CAPSIM = 112
If bc = "EZg30v" Then BOD2CAPSIM = 112
If bc = "fABk" Then BOD2CAPSIM = 119
If bc = "fAFk" Then BOD2CAPSIM = 119
If bc = "fAFz" Then BOD2CAPSIM = 113
If bc = "faVc" Then BOD2CAPSIM = 101
If bc = "faVz" Then BOD2CAPSIM = 102
If bc = "faVzt" Then BOD2CAPSIM = 102
If bc = "FG" Then BOD2CAPSIM = 114
If bc = "fHn21" Then BOD2CAPSIM = 109
If bc = "fhVc" Then BOD2CAPSIM = 101
If bc = "fhVd" Then BOD2CAPSIM = 101
If bc = "fhVz" Then BOD2CAPSIM = 102
If bc = "fiVc" Then BOD2CAPSIM = 105
If bc = "fiVz" Then BOD2CAPSIM = 105
If bc = "fiWp" Then BOD2CAPSIM = 105
If bc = "fiWz" Then BOD2CAPSIM = 105
If bc = "FK" Then BOD2CAPSIM = 121
If bc = "FKk" Then BOD2CAPSIM = 121
If bc = "fkpZg23" Then BOD2CAPSIM = 119
If bc = "fkpZg23g" Then BOD2CAPSIM = 120
If bc = "fkpZg23t" Then BOD2CAPSIM = 119
If bc = "fKRn1" Then BOD2CAPSIM = 119
If bc = "fKRn1g" Then BOD2CAPSIM = 120
If bc = "fKRn2g" Then BOD2CAPSIM = 120
If bc = "fKRn8" Then BOD2CAPSIM = 119
If bc = "fKRn8g" Then BOD2CAPSIM = 120
If bc = "fkVc" Then BOD2CAPSIM = 103
If bc = "fkVs" Then BOD2CAPSIM = 103
If bc = "fkVz" Then BOD2CAPSIM = 104
If bc = "fkWz" Then BOD2CAPSIM = 104
If bc = "fkWzg" Then BOD2CAPSIM = 104
If bc = "fkZn21" Then BOD2CAPSIM = 119
If bc = "fkZn23" Then BOD2CAPSIM = 119
If bc = "fkZn23g" Then BOD2CAPSIM = 120
If bc = "fkZn30" Then BOD2CAPSIM = 120
If bc = "fMn56Cp" Then BOD2CAPSIM = 119
If bc = "fMn56Cv" Then BOD2CAPSIM = 118
If bc = "fpLn5" Then BOD2CAPSIM = 121
If bc = "fpRn59" Then BOD2CAPSIM = 119
If bc = "fpRn86" Then BOD2CAPSIM = 117
If bc = "fpVc" Then BOD2CAPSIM = 103
If bc = "fpVs" Then BOD2CAPSIM = 103
If bc = "fpVz" Then BOD2CAPSIM = 104
If bc = "fpZg21" Then BOD2CAPSIM = 109
If bc = "fpZg21g" Then BOD2CAPSIM = 110
If bc = "fpZg23" Then BOD2CAPSIM = 113
If bc = "fpZg23g" Then BOD2CAPSIM = 113
If bc = "fpZg23t" Then BOD2CAPSIM = 111
If bc = "fpZg23x" Then BOD2CAPSIM = 111
If bc = "fpZn21" Then BOD2CAPSIM = 109
If bc = "fpZn23tg" Then BOD2CAPSIM = 111
If bc = "fRn15C" Then BOD2CAPSIM = 115
If bc = "fRn62C" Then BOD2CAPSIM = 119
If bc = "fRn62Cg" Then BOD2CAPSIM = 120
If bc = "fRn95C" Then BOD2CAPSIM = 116
If bc = "fRo60C" Then BOD2CAPSIM = 116
If bc = "fRv01C" Then BOD2CAPSIM = 118
If bc = "fVc" Then BOD2CAPSIM = 101
If bc = "fvWz" Then BOD2CAPSIM = 102
If bc = "fvWzt" Then BOD2CAPSIM = 102
If bc = "fvWztx" Then BOD2CAPSIM = 102
If bc = "fVz" Then BOD2CAPSIM = 102
If bc = "fZn21" Then BOD2CAPSIM = 107
If bc = "fZn21g" Then BOD2CAPSIM = 107
If bc = "fZn23" Then BOD2CAPSIM = 113
If bc = "fZn23-F" Then BOD2CAPSIM = 113
If bc = "fZn23g" Then BOD2CAPSIM = 113
If bc = "fzVc" Then BOD2CAPSIM = 105
If bc = "fzVz" Then BOD2CAPSIM = 105
If bc = "fzVzt" Then BOD2CAPSIM = 105
If bc = "fzWp" Then BOD2CAPSIM = 105
If bc = "fzWz" Then BOD2CAPSIM = 105
If bc = "fzWzt" Then BOD2CAPSIM = 105
If bc = "gbEZ21" Then BOD2CAPSIM = 112
If bc = "gbEZ30" Then BOD2CAPSIM = 112
If bc = "gcHd30" Then BOD2CAPSIM = 114
If bc = "gcHn21" Then BOD2CAPSIM = 109
If bc = "gcHn30" Then BOD2CAPSIM = 114
If bc = "gcY21" Then BOD2CAPSIM = 109
If bc = "gcY23" Then BOD2CAPSIM = 113
If bc = "gcY30" Then BOD2CAPSIM = 114
If bc = "gcZd30" Then BOD2CAPSIM = 114
If bc = "gHd21" Then BOD2CAPSIM = 108
If bc = "gHd30" Then BOD2CAPSIM = 114
If bc = "gHn21" Then BOD2CAPSIM = 109
If bc = "gHn21t" Then BOD2CAPSIM = 111
If bc = "gHn21x" Then BOD2CAPSIM = 111
If bc = "gHn23" Then BOD2CAPSIM = 113
If bc = "gHn23x" Then BOD2CAPSIM = 111
If bc = "gHn30" Then BOD2CAPSIM = 114
If bc = "gHn30t" Then BOD2CAPSIM = 114
If bc = "gHn30x" Then BOD2CAPSIM = 114
If bc = "gKRd1" Then BOD2CAPSIM = 119
If bc = "gKRd7" Then BOD2CAPSIM = 119
If bc = "gKRn1" Then BOD2CAPSIM = 119
If bc = "gKRn2" Then BOD2CAPSIM = 119
If bc = "gLd6" Then BOD2CAPSIM = 121
If bc = "gLh6" Then BOD2CAPSIM = 121
If bc = "gMK" Then BOD2CAPSIM = 115
If bc = "gMn15C" Then BOD2CAPSIM = 115
If bc = "gMn25C" Then BOD2CAPSIM = 115
If bc = "gMn25Cv" Then BOD2CAPSIM = 115
If bc = "gMn52C" Then BOD2CAPSIM = 119
If bc = "gMn52Cp" Then BOD2CAPSIM = 119
If bc = "gMn52Cw" Then BOD2CAPSIM = 119
If bc = "gMn53C" Then BOD2CAPSIM = 117
If bc = "gMn53Cp" Then BOD2CAPSIM = 119
If bc = "gMn53Cpx" Then BOD2CAPSIM = 119
If bc = "gMn53Cv" Then BOD2CAPSIM = 118
If bc = "gMn53Cw" Then BOD2CAPSIM = 117
If bc = "gMn53Cwp" Then BOD2CAPSIM = 119
If bc = "gMn58C" Then BOD2CAPSIM = 117
If bc = "gMn58Cv" Then BOD2CAPSIM = 117
If bc = "nkZn50A" Then BOD2CAPSIM = 119
If bc = "gMn82C" Then BOD2CAPSIM = 119
If bc = "gMn83C" Then BOD2CAPSIM = 117
If bc = "gMn83Cp" Then BOD2CAPSIM = 117
If bc = "gMn83Cv" Then BOD2CAPSIM = 118
If bc = "gMn83Cw" Then BOD2CAPSIM = 117
If bc = "gMn83Cwp" Then BOD2CAPSIM = 117
If bc = "gMn85C" Then BOD2CAPSIM = 116
If bc = "gMn85Cv" Then BOD2CAPSIM = 118
If bc = "gMn85Cwl" Then BOD2CAPSIM = 116
If bc = "gMn88C" Then BOD2CAPSIM = 117
If bc = "gMn88Cl" Then BOD2CAPSIM = 117
If bc = "gMn88Clv" Then BOD2CAPSIM = 118
If bc = "gMn88Cv" Then BOD2CAPSIM = 118
If bc = "gMn88Cw" Then BOD2CAPSIM = 117
If bc = "gpZg23x" Then BOD2CAPSIM = 111
If bc = "gpZg30" Then BOD2CAPSIM = 114
If bc = "gpZn21" Then BOD2CAPSIM = 109
If bc = "gpZn21x" Then BOD2CAPSIM = 111
If bc = "gpZn23x" Then BOD2CAPSIM = 111
If bc = "gpZn30" Then BOD2CAPSIM = 114
If bc = "gRd10A" Then BOD2CAPSIM = 119
If bc = "gRn15A" Then BOD2CAPSIM = 119
If bc = "gRn94Cv" Then BOD2CAPSIM = 117
If bc = "gtZd30" Then BOD2CAPSIM = 114
If bc = "gvWp" Then BOD2CAPSIM = 102
If bc = "gY21" Then BOD2CAPSIM = 109
If bc = "gY21g" Then BOD2CAPSIM = 109
If bc = "gY23" Then BOD2CAPSIM = 113
If bc = "gY30" Then BOD2CAPSIM = 114
If bc = "gY30-F" Then BOD2CAPSIM = 114
If bc = "gY30-G" Then BOD2CAPSIM = 114
If bc = "gZb30" Then BOD2CAPSIM = 114
If bc = "gZd21" Then BOD2CAPSIM = 107
If bc = "gZd30" Then BOD2CAPSIM = 114
If bc = "gzEZ21" Then BOD2CAPSIM = 112
If bc = "gzEZ23" Then BOD2CAPSIM = 112
If bc = "gzEZ30" Then BOD2CAPSIM = 112
If bc = "gZn30" Then BOD2CAPSIM = 114
If bc = "Hd21" Then BOD2CAPSIM = 108
If bc = "Hd21g" Then BOD2CAPSIM = 108
If bc = "Hd21x" Then BOD2CAPSIM = 108
If bc = "Hd23" Then BOD2CAPSIM = 113
If bc = "Hd23g" Then BOD2CAPSIM = 110
If bc = "Hd23x" Then BOD2CAPSIM = 111
If bc = "Hd30" Then BOD2CAPSIM = 114
If bc = "Hd30g" Then BOD2CAPSIM = 114
If bc = "hEV" Then BOD2CAPSIM = 101
If bc = "Hn21" Then BOD2CAPSIM = 109
If bc = "Hn21-F" Then BOD2CAPSIM = 109
If bc = "Hn21g" Then BOD2CAPSIM = 110
If bc = "Hn21gx" Then BOD2CAPSIM = 110
If bc = "Hn21t" Then BOD2CAPSIM = 111
If bc = "Hn21v" Then BOD2CAPSIM = 109
If bc = "Hn21w" Then BOD2CAPSIM = 109
If bc = "Hn21wg" Then BOD2CAPSIM = 109
If bc = "Hn21x" Then BOD2CAPSIM = 111
If bc = "Hn21x-F" Then BOD2CAPSIM = 111
If bc = "Hn21xg" Then BOD2CAPSIM = 111
If bc = "Hn23" Then BOD2CAPSIM = 113
If bc = "Hn23-F" Then BOD2CAPSIM = 113
If bc = "Hn23g" Then BOD2CAPSIM = 110
If bc = "Hn23t" Then BOD2CAPSIM = 111
If bc = "Hn23x" Then BOD2CAPSIM = 111
If bc = "Hn23x-F" Then BOD2CAPSIM = 111
If bc = "Hn23xg" Then BOD2CAPSIM = 111
If bc = "Hn30" Then BOD2CAPSIM = 114
If bc = "Hn30g" Then BOD2CAPSIM = 114
If bc = "Hn30x" Then BOD2CAPSIM = 114
If bc = "hRd10A" Then BOD2CAPSIM = 119
If bc = "hRd10C" Then BOD2CAPSIM = 119
If bc = "hRd90A" Then BOD2CAPSIM = 116
If bc = "hVb" Then BOD2CAPSIM = 101
If bc = "hVc" Then BOD2CAPSIM = 101
If bc = "hVcc" Then BOD2CAPSIM = 101
If bc = "hVd" Then BOD2CAPSIM = 101
If bc = "hVk" Then BOD2CAPSIM = 106
If bc = "hVkl" Then BOD2CAPSIM = 106
If bc = "hVr" Then BOD2CAPSIM = 101
If bc = "hVs" Then BOD2CAPSIM = 101
If bc = "hVsc" Then BOD2CAPSIM = 101
If bc = "hVz" Then BOD2CAPSIM = 102
If bc = "hVzc" Then BOD2CAPSIM = 102
If bc = "hVzg" Then BOD2CAPSIM = 102
If bc = "hVzx" Then BOD2CAPSIM = 102
If bc = "hZd20A" Then BOD2CAPSIM = 107
If bc = "iVc" Then BOD2CAPSIM = 105
If bc = "iVp" Then BOD2CAPSIM = 105
If bc = "iVpc" Then BOD2CAPSIM = 105
If bc = "iVpg" Then BOD2CAPSIM = 105
If bc = "iVpt" Then BOD2CAPSIM = 105
If bc = "iVpx" Then BOD2CAPSIM = 105
If bc = "iVs" Then BOD2CAPSIM = 105
If bc = "iVz" Then BOD2CAPSIM = 105
If bc = "iVzg" Then BOD2CAPSIM = 105
If bc = "iVzt" Then BOD2CAPSIM = 105
If bc = "iVzx" Then BOD2CAPSIM = 105
If bc = "iWp" Then BOD2CAPSIM = 105
If bc = "iWpc" Then BOD2CAPSIM = 105
If bc = "iWpg" Then BOD2CAPSIM = 105
If bc = "iWpt" Then BOD2CAPSIM = 105
If bc = "iWpx" Then BOD2CAPSIM = 105
If bc = "iWz" Then BOD2CAPSIM = 105
If bc = "iWzt" Then BOD2CAPSIM = 105
If bc = "iWzx" Then BOD2CAPSIM = 105
If bc = "kcHn21" Then BOD2CAPSIM = 119
If bc = "kgpZg30" Then BOD2CAPSIM = 120
If bc = "kHn21" Then BOD2CAPSIM = 119
If bc = "kHn21g" Then BOD2CAPSIM = 120
If bc = "kHn21x" Then BOD2CAPSIM = 119
If bc = "kHn23" Then BOD2CAPSIM = 119
If bc = "kHn23x" Then BOD2CAPSIM = 119
If bc = "kHn30" Then BOD2CAPSIM = 120
If bc = "KK" Then BOD2CAPSIM = 121
If bc = "KM" Then BOD2CAPSIM = 121
If bc = "kMn43C" Then BOD2CAPSIM = 117
If bc = "kMn43Cp" Then BOD2CAPSIM = 117
If bc = "kMn43Cpx" Then BOD2CAPSIM = 117
If bc = "kMn43Cv" Then BOD2CAPSIM = 118
If bc = "kMn43Cwp" Then BOD2CAPSIM = 117
If bc = "kMn48C" Then BOD2CAPSIM = 117
If bc = "kMn48Cl" Then BOD2CAPSIM = 117
If bc = "kMn48Clv" Then BOD2CAPSIM = 118
If bc = "kMn48Cv" Then BOD2CAPSIM = 118
If bc = "kMn48Cvl" Then BOD2CAPSIM = 118
If bc = "kMn48Cw" Then BOD2CAPSIM = 117
If bc = "kMn63C" Then BOD2CAPSIM = 117
If bc = "kMn63Cp" Then BOD2CAPSIM = 119
If bc = "kMn63Cpx" Then BOD2CAPSIM = 119
If bc = "kMn63Cv" Then BOD2CAPSIM = 118
If bc = "kMn63Cwp" Then BOD2CAPSIM = 119
If bc = "kMn68C" Then BOD2CAPSIM = 117
If bc = "kMn68Cl" Then BOD2CAPSIM = 117
If bc = "kMn68Cv" Then BOD2CAPSIM = 118
If bc = "kpZg20A" Then BOD2CAPSIM = 119
If bc = "kpZg21" Then BOD2CAPSIM = 119
If bc = "kpZg21g" Then BOD2CAPSIM = 120
If bc = "kpZg23" Then BOD2CAPSIM = 119
If bc = "kpZg23g" Then BOD2CAPSIM = 120
If bc = "kpZg23t" Then BOD2CAPSIM = 119
If bc = "kpZg23x" Then BOD2CAPSIM = 119
If bc = "kpZn21" Then BOD2CAPSIM = 119
If bc = "kpZn21g" Then BOD2CAPSIM = 120
If bc = "kpZn23" Then BOD2CAPSIM = 119
If bc = "kpZn23x" Then BOD2CAPSIM = 119
If bc = "KRd1" Then BOD2CAPSIM = 119
If bc = "KRd1g" Then BOD2CAPSIM = 120
If bc = "KRd7" Then BOD2CAPSIM = 119
If bc = "KRd7g" Then BOD2CAPSIM = 120
If bc = "KRn1" Then BOD2CAPSIM = 119
If bc = "KRn1g" Then BOD2CAPSIM = 120
If bc = "KRn2" Then BOD2CAPSIM = 119
If bc = "KRn2g" Then BOD2CAPSIM = 120
If bc = "KRn2w" Then BOD2CAPSIM = 119
If bc = "KRn8" Then BOD2CAPSIM = 119
If bc = "KRn8g" Then BOD2CAPSIM = 120
If bc = "KS" Then BOD2CAPSIM = 115
If bc = "kSn13A" Then BOD2CAPSIM = 119
If bc = "kSn13Av" Then BOD2CAPSIM = 119
If bc = "kSn13Aw" Then BOD2CAPSIM = 119
If bc = "kSn14A" Then BOD2CAPSIM = 119
If bc = "kSn14Ap" Then BOD2CAPSIM = 119
If bc = "kSn14Av" Then BOD2CAPSIM = 119
If bc = "kSn14Aw" Then BOD2CAPSIM = 119
If bc = "kSn14Awp" Then BOD2CAPSIM = 119
If bc = "KT" Then BOD2CAPSIM = 115
If bc = "kVb" Then BOD2CAPSIM = 103
If bc = "kVc" Then BOD2CAPSIM = 103
If bc = "kVcc" Then BOD2CAPSIM = 103
If bc = "kVd" Then BOD2CAPSIM = 103
If bc = "kVk" Then BOD2CAPSIM = 106
If bc = "kVr" Then BOD2CAPSIM = 103
If bc = "kVs" Then BOD2CAPSIM = 103
If bc = "kVsc" Then BOD2CAPSIM = 103
If bc = "kVz" Then BOD2CAPSIM = 104
If bc = "kVzc" Then BOD2CAPSIM = 104
If bc = "kVzx" Then BOD2CAPSIM = 104
If bc = "kWp" Then BOD2CAPSIM = 104
If bc = "kWpg" Then BOD2CAPSIM = 104
If bc = "kWpx" Then BOD2CAPSIM = 104
If bc = "kWz" Then BOD2CAPSIM = 104
If bc = "kWzg" Then BOD2CAPSIM = 104
If bc = "kWzx" Then BOD2CAPSIM = 104
If bc = "KX" Then BOD2CAPSIM = 115
If bc = "kZb21" Then BOD2CAPSIM = 119
If bc = "kZb23" Then BOD2CAPSIM = 119
If bc = "kZn10A" Then BOD2CAPSIM = 119
If bc = "kZn10Av" Then BOD2CAPSIM = 119
If bc = "kZn21" Then BOD2CAPSIM = 119
If bc = "kZn21g" Then BOD2CAPSIM = 120
If bc = "kZn21p" Then BOD2CAPSIM = 119
If bc = "kZn21r" Then BOD2CAPSIM = 119
If bc = "kZn21w" Then BOD2CAPSIM = 119
If bc = "kZn21x" Then BOD2CAPSIM = 119
If bc = "kZn23" Then BOD2CAPSIM = 119
If bc = "kZn30" Then BOD2CAPSIM = 120
If bc = "kZn30A" Then BOD2CAPSIM = 120
If bc = "kZn30Ar" Then BOD2CAPSIM = 120
If bc = "kZn30x" Then BOD2CAPSIM = 120
If bc = "kZn40A" Then BOD2CAPSIM = 119
If bc = "kZn40Ap" Then BOD2CAPSIM = 119
If bc = "kZn40Av" Then BOD2CAPSIM = 119
If bc = "kZn50A" Then BOD2CAPSIM = 119
If bc = "kZn50Ap" Then BOD2CAPSIM = 119
If bc = "kZn50Ar" Then BOD2CAPSIM = 119
If bc = "Ld5" Then BOD2CAPSIM = 121
If bc = "Ld5g" Then BOD2CAPSIM = 121
If bc = "Ld5m" Then BOD2CAPSIM = 121
If bc = "Ld5t" Then BOD2CAPSIM = 121
If bc = "Ld6" Then BOD2CAPSIM = 121
If bc = "Ld6a" Then BOD2CAPSIM = 121
If bc = "Ld6g" Then BOD2CAPSIM = 121
If bc = "Ld6k" Then BOD2CAPSIM = 121
If bc = "Ld6m" Then BOD2CAPSIM = 121
If bc = "Ld6s" Then BOD2CAPSIM = 121
If bc = "Ld6t" Then BOD2CAPSIM = 121
If bc = "Ldd5" Then BOD2CAPSIM = 121
If bc = "Ldd5g" Then BOD2CAPSIM = 121
If bc = "Ldd6" Then BOD2CAPSIM = 121
If bc = "Ldh5" Then BOD2CAPSIM = 121
If bc = "Ldh5g" Then BOD2CAPSIM = 121
If bc = "Ldh5t" Then BOD2CAPSIM = 121
If bc = "Ldh6" Then BOD2CAPSIM = 121
If bc = "Ldh6m" Then BOD2CAPSIM = 121
If bc = "lFG" Then BOD2CAPSIM = 114
If bc = "lFK" Then BOD2CAPSIM = 121
If bc = "lFKk" Then BOD2CAPSIM = 121
If bc = "Lh5" Then BOD2CAPSIM = 121
If bc = "Lh5g" Then BOD2CAPSIM = 121
If bc = "Lh6" Then BOD2CAPSIM = 121
If bc = "Lh6g" Then BOD2CAPSIM = 121
If bc = "Lh6s" Then BOD2CAPSIM = 121
If bc = "lKK" Then BOD2CAPSIM = 116
If bc = "lKM" Then BOD2CAPSIM = 116
If bc = "lKRd7" Then BOD2CAPSIM = 119
If bc = "lKS" Then BOD2CAPSIM = 121
If bc = "Ln5" Then BOD2CAPSIM = 121
If bc = "Ln5g" Then BOD2CAPSIM = 121
If bc = "Ln5m" Then BOD2CAPSIM = 121
If bc = "Ln5t" Then BOD2CAPSIM = 121
If bc = "Ln6" Then BOD2CAPSIM = 121
If bc = "Ln6a" Then BOD2CAPSIM = 121
If bc = "Ln6m" Then BOD2CAPSIM = 121
If bc = "Ln6t" Then BOD2CAPSIM = 121
If bc = "Lnd5" Then BOD2CAPSIM = 121
If bc = "Lnd5g" Then BOD2CAPSIM = 121
If bc = "Lnd5m" Then BOD2CAPSIM = 121
If bc = "Lnd5t" Then BOD2CAPSIM = 121
If bc = "Lnd6" Then BOD2CAPSIM = 121
If bc = "Lnd6v" Then BOD2CAPSIM = 121
If bc = "Lnh6" Then BOD2CAPSIM = 121
If bc = "MA" Then BOD2CAPSIM = 116
If bc = "mcY23" Then BOD2CAPSIM = 113
If bc = "mcY23x" Then BOD2CAPSIM = 111
If bc = "mHd23" Then BOD2CAPSIM = 113
If bc = "mHn21x" Then BOD2CAPSIM = 111
If bc = "mHn23x" Then BOD2CAPSIM = 111
If bc = "MK" Then BOD2CAPSIM = 116
If bc = "mKK" Then BOD2CAPSIM = 116
If bc = "mKRd7" Then BOD2CAPSIM = 119
If bc = "mKX" Then BOD2CAPSIM = 115
If bc = "mLd6s" Then BOD2CAPSIM = 121
If bc = "mLh6s" Then BOD2CAPSIM = 121
If bc = "Mn12A" Then BOD2CAPSIM = 119
If bc = "Mn12Ap" Then BOD2CAPSIM = 119
If bc = "Mn12Av" Then BOD2CAPSIM = 119
If bc = "Mn12Awp" Then BOD2CAPSIM = 119
If bc = "Mn15A" Then BOD2CAPSIM = 115
If bc = "Mn15Ap" Then BOD2CAPSIM = 119
If bc = "Mn15Av" Then BOD2CAPSIM = 118
If bc = "Mn15Aw" Then BOD2CAPSIM = 115
If bc = "Mn15Awp" Then BOD2CAPSIM = 119
If bc = "Mn15C" Then BOD2CAPSIM = 115
If bc = "Mn15Clv" Then BOD2CAPSIM = 118
If bc = "Mn15Cv" Then BOD2CAPSIM = 118
If bc = "Mn15Cw" Then BOD2CAPSIM = 115
If bc = "Mn22A" Then BOD2CAPSIM = 119
If bc = "Mn22Alv" Then BOD2CAPSIM = 115
If bc = "Mn22Ap" Then BOD2CAPSIM = 119
If bc = "Mn22Av" Then BOD2CAPSIM = 115
If bc = "Mn22Aw" Then BOD2CAPSIM = 119
If bc = "Mn22Awp" Then BOD2CAPSIM = 119
If bc = "Mn22Ax" Then BOD2CAPSIM = 119
If bc = "Mn25A" Then BOD2CAPSIM = 115
If bc = "Mn25Alv" Then BOD2CAPSIM = 115
If bc = "Mn25Ap" Then BOD2CAPSIM = 119
If bc = "Mn25Av" Then BOD2CAPSIM = 118
If bc = "Mn25Aw" Then BOD2CAPSIM = 115
If bc = "Mn25Awp" Then BOD2CAPSIM = 119
If bc = "Mn25C" Then BOD2CAPSIM = 115
If bc = "Mn25Cp" Then BOD2CAPSIM = 119
If bc = "Mn25Cv" Then BOD2CAPSIM = 118
If bc = "Mn25Cw" Then BOD2CAPSIM = 115
If bc = "Mn35A" Then BOD2CAPSIM = 116
If bc = "Mn35Ap" Then BOD2CAPSIM = 119
If bc = "Mn35Av" Then BOD2CAPSIM = 118
If bc = "Mn35Aw" Then BOD2CAPSIM = 116
If bc = "Mn35Awp" Then BOD2CAPSIM = 119
If bc = "Mn35Ax" Then BOD2CAPSIM = 116
If bc = "Mn45A" Then BOD2CAPSIM = 117
If bc = "Mn45Ap" Then BOD2CAPSIM = 119
If bc = "Mn45Av" Then BOD2CAPSIM = 118
If bc = "Mn52C" Then BOD2CAPSIM = 119
If bc = "Mn52Cp" Then BOD2CAPSIM = 119
If bc = "Mn52Cpx" Then BOD2CAPSIM = 119
If bc = "Mn52Cwp" Then BOD2CAPSIM = 119
If bc = "Mn52Cx" Then BOD2CAPSIM = 119
If bc = "Mn56A" Then BOD2CAPSIM = 117
If bc = "Mn56Ap" Then BOD2CAPSIM = 119
If bc = "Mn56Av" Then BOD2CAPSIM = 118
If bc = "Mn56Aw" Then BOD2CAPSIM = 117
If bc = "Mn56C" Then BOD2CAPSIM = 117
If bc = "Mn56Cp" Then BOD2CAPSIM = 119
If bc = "Mn56Cv" Then BOD2CAPSIM = 118
If bc = "Mn56Cwp" Then BOD2CAPSIM = 119
If bc = "Mn82A" Then BOD2CAPSIM = 119
If bc = "Mn82Ap" Then BOD2CAPSIM = 119
If bc = "Mn82C" Then BOD2CAPSIM = 119
If bc = "Mn82Cp" Then BOD2CAPSIM = 119
If bc = "Mn82Cpx" Then BOD2CAPSIM = 119
If bc = "Mn82Cwp" Then BOD2CAPSIM = 119
If bc = "Mn85C" Then BOD2CAPSIM = 116
If bc = "Mn85Clwp" Then BOD2CAPSIM = 119
If bc = "Mn85Cp" Then BOD2CAPSIM = 119
If bc = "Mn85Cv" Then BOD2CAPSIM = 118
If bc = "Mn85Cw" Then BOD2CAPSIM = 116
If bc = "Mn85Cwp" Then BOD2CAPSIM = 119
If bc = "Mn86A" Then BOD2CAPSIM = 117
If bc = "Mn86Al" Then BOD2CAPSIM = 117
If bc = "Mn86Av" Then BOD2CAPSIM = 118
If bc = "Mn86Aw" Then BOD2CAPSIM = 117
If bc = "Mn86C" Then BOD2CAPSIM = 117
If bc = "Mn86Cl" Then BOD2CAPSIM = 117
If bc = "Mn86Clv" Then BOD2CAPSIM = 117
If bc = "Mn86Clw" Then BOD2CAPSIM = 117
If bc = "Mn86Clwp" Then BOD2CAPSIM = 119
If bc = "Mn86Cp" Then BOD2CAPSIM = 119
If bc = "Mn86Cv" Then BOD2CAPSIM = 118
If bc = "Mn86Cw" Then BOD2CAPSIM = 117
If bc = "Mn86Cwp" Then BOD2CAPSIM = 119
If bc = "Mo10A" Then BOD2CAPSIM = 115
If bc = "Mo10Av" Then BOD2CAPSIM = 115
If bc = "Mo20A" Then BOD2CAPSIM = 115
If bc = "Mo20Av" Then BOD2CAPSIM = 115
If bc = "Mo50C" Then BOD2CAPSIM = 115
If bc = "Mo80A" Then BOD2CAPSIM = 116
If bc = "Mo80Ap" Then BOD2CAPSIM = 119
If bc = "Mo80Av" Then BOD2CAPSIM = 118
If bc = "Mo80C" Then BOD2CAPSIM = 116
If bc = "Mo80Cl" Then BOD2CAPSIM = 116
If bc = "Mo80Cp" Then BOD2CAPSIM = 119
If bc = "Mo80Cv" Then BOD2CAPSIM = 118
If bc = "Mo80Cvl" Then BOD2CAPSIM = 118
If bc = "Mo80Cw" Then BOD2CAPSIM = 116
If bc = "Mo80Cwp" Then BOD2CAPSIM = 119
If bc = "MOb12" Then BOD2CAPSIM = 119
If bc = "MOb15" Then BOD2CAPSIM = 115
If bc = "MOb72" Then BOD2CAPSIM = 119
If bc = "MOb75" Then BOD2CAPSIM = 116
If bc = "MOo02" Then BOD2CAPSIM = 119
If bc = "MOo02v" Then BOD2CAPSIM = 119
If bc = "MOo05" Then BOD2CAPSIM = 115
If bc = "Mv41C" Then BOD2CAPSIM = 118
If bc = "Mv41Cl" Then BOD2CAPSIM = 118
If bc = "Mv41Cp" Then BOD2CAPSIM = 118
If bc = "Mv41Cv" Then BOD2CAPSIM = 118
If bc = "Mv51A" Then BOD2CAPSIM = 118
If bc = "Mv51Al" Then BOD2CAPSIM = 118
If bc = "Mv51Ap" Then BOD2CAPSIM = 118
If bc = "Mv61C" Then BOD2CAPSIM = 118
If bc = "Mv61Cl" Then BOD2CAPSIM = 118
If bc = "Mv61Cp" Then BOD2CAPSIM = 118
If bc = "Mv81A" Then BOD2CAPSIM = 118
If bc = "Mv81Al" Then BOD2CAPSIM = 118
If bc = "Mv81Ap" Then BOD2CAPSIM = 118
If bc = "mY23" Then BOD2CAPSIM = 113
If bc = "mY23x" Then BOD2CAPSIM = 111
If bc = "mZb23x" Then BOD2CAPSIM = 111
If bc = "MZk" Then BOD2CAPSIM = 121
If bc = "MZz" Then BOD2CAPSIM = 107
If bc = "nAO" Then BOD2CAPSIM = 119
If bc = "nkZn21" Then BOD2CAPSIM = 119
If bc = "nkZn50Ab" Then BOD2CAPSIM = 119
If bc = "nMn15A" Then BOD2CAPSIM = 115
If bc = "nMn15Av" Then BOD2CAPSIM = 115
If bc = "nMo10A" Then BOD2CAPSIM = 115
If bc = "nMo10Av" Then BOD2CAPSIM = 118
If bc = "nMo80A" Then BOD2CAPSIM = 116
If bc = "nMo80Aw" Then BOD2CAPSIM = 116
If bc = "nMv61C" Then BOD2CAPSIM = 118
If bc = "npMo50l" Then BOD2CAPSIM = 115
If bc = "npMo80l" Then BOD2CAPSIM = 116
If bc = "nSn13A" Then BOD2CAPSIM = 113
If bc = "nSn13Av" Then BOD2CAPSIM = 113
If bc = "nvWz" Then BOD2CAPSIM = 102
If bc = "nZn21" Then BOD2CAPSIM = 107
If bc = "nZn40A" Then BOD2CAPSIM = 107
If bc = "nZn50A" Then BOD2CAPSIM = 107
If bc = "nZn50Ab" Then BOD2CAPSIM = 107
If bc = "ohVb" Then BOD2CAPSIM = 101
If bc = "ohVc" Then BOD2CAPSIM = 101
If bc = "ohVk" Then BOD2CAPSIM = 106
If bc = "ohVs" Then BOD2CAPSIM = 101
If bc = "opVb" Then BOD2CAPSIM = 103
If bc = "opVc" Then BOD2CAPSIM = 103
If bc = "opVk" Then BOD2CAPSIM = 106
If bc = "opVs" Then BOD2CAPSIM = 103
If bc = "pKRn1" Then BOD2CAPSIM = 119
If bc = "pKRn1g" Then BOD2CAPSIM = 120
If bc = "pKRn2" Then BOD2CAPSIM = 119
If bc = "pKRn2g" Then BOD2CAPSIM = 120
If bc = "pLn5" Then BOD2CAPSIM = 121
If bc = "pLn5g" Then BOD2CAPSIM = 121
If bc = "pMn52A" Then BOD2CAPSIM = 119
If bc = "pMn52C" Then BOD2CAPSIM = 119
If bc = "pMn52Cp" Then BOD2CAPSIM = 119
If bc = "pMn55A" Then BOD2CAPSIM = 115
If bc = "pMn55Av" Then BOD2CAPSIM = 118
If bc = "pMn55Aw" Then BOD2CAPSIM = 115
If bc = "pMn55C" Then BOD2CAPSIM = 115
If bc = "pMn55Cp" Then BOD2CAPSIM = 119
If bc = "pMn56C" Then BOD2CAPSIM = 117
If bc = "pMn56Cl" Then BOD2CAPSIM = 117
If bc = "pMn82A" Then BOD2CAPSIM = 119
If bc = "pMn82C" Then BOD2CAPSIM = 119
If bc = "pMn85A" Then BOD2CAPSIM = 116
If bc = "pMn85Aw" Then BOD2CAPSIM = 116
If bc = "pMn85C" Then BOD2CAPSIM = 116
If bc = "pMn85Cv" Then BOD2CAPSIM = 118
If bc = "pMn86C" Then BOD2CAPSIM = 117
If bc = "pMn86Cl" Then BOD2CAPSIM = 117
If bc = "pMn86Cv" Then BOD2CAPSIM = 118
If bc = "pMn86Cw" Then BOD2CAPSIM = 117
If bc = "pMn86Cwl" Then BOD2CAPSIM = 117
If bc = "pMo50" Then BOD2CAPSIM = 115
If bc = "pMo50l" Then BOD2CAPSIM = 115
If bc = "pMo50w" Then BOD2CAPSIM = 115
If bc = "pMo80" Then BOD2CAPSIM = 116
If bc = "pMo80l" Then BOD2CAPSIM = 116
If bc = "pMo80v" Then BOD2CAPSIM = 118
If bc = "pMv51" Then BOD2CAPSIM = 118
If bc = "pMv81" Then BOD2CAPSIM = 118
If bc = "pMv81l" Then BOD2CAPSIM = 118
If bc = "pMv81p" Then BOD2CAPSIM = 118
If bc = "pRn56" Then BOD2CAPSIM = 119
If bc = "pRn56p" Then BOD2CAPSIM = 119
If bc = "pRn56v" Then BOD2CAPSIM = 118
If bc = "pRn56wp" Then BOD2CAPSIM = 119
If bc = "pRn59" Then BOD2CAPSIM = 119
If bc = "pRn59p" Then BOD2CAPSIM = 119
If bc = "pRn59t" Then BOD2CAPSIM = 119
If bc = "pRn59w" Then BOD2CAPSIM = 119
If bc = "pRn86" Then BOD2CAPSIM = 117
If bc = "pRn86p" Then BOD2CAPSIM = 119
If bc = "pRn86t" Then BOD2CAPSIM = 117
If bc = "pRn86v" Then BOD2CAPSIM = 118
If bc = "pRn86w" Then BOD2CAPSIM = 117
If bc = "pRn86wp" Then BOD2CAPSIM = 119
If bc = "pRn89" Then BOD2CAPSIM = 118
If bc = "pRn89v" Then BOD2CAPSIM = 118
If bc = "pRv81" Then BOD2CAPSIM = 118
If bc = "pVb" Then BOD2CAPSIM = 103
If bc = "pVc" Then BOD2CAPSIM = 103
If bc = "pVcc" Then BOD2CAPSIM = 103
If bc = "pVd" Then BOD2CAPSIM = 103
If bc = "pVk" Then BOD2CAPSIM = 106
If bc = "pVr" Then BOD2CAPSIM = 103
If bc = "pVs" Then BOD2CAPSIM = 103
If bc = "pVsc" Then BOD2CAPSIM = 103
If bc = "pVsl" Then BOD2CAPSIM = 103
If bc = "pVz" Then BOD2CAPSIM = 104
If bc = "pVzx" Then BOD2CAPSIM = 104
If bc = "pZg20A" Then BOD2CAPSIM = 107
If bc = "pZg20Ar" Then BOD2CAPSIM = 107
If bc = "pZg21" Then BOD2CAPSIM = 109
If bc = "pZg21g" Then BOD2CAPSIM = 110
If bc = "pZg21r" Then BOD2CAPSIM = 111
If bc = "pZg21t" Then BOD2CAPSIM = 111
If bc = "pZg21w" Then BOD2CAPSIM = 109
If bc = "pZg21x" Then BOD2CAPSIM = 111
If bc = "pZg23" Then BOD2CAPSIM = 113
If bc = "pZg23g" Then BOD2CAPSIM = 113
If bc = "pZg23r" Then BOD2CAPSIM = 113
If bc = "pZg23t" Then BOD2CAPSIM = 111
If bc = "pZg23w" Then BOD2CAPSIM = 113
If bc = "pZg23x" Then BOD2CAPSIM = 111
If bc = "pZg30" Then BOD2CAPSIM = 114
If bc = "pZg30p" Then BOD2CAPSIM = 114
If bc = "pZg30r" Then BOD2CAPSIM = 114
If bc = "pZg30x" Then BOD2CAPSIM = 114
If bc = "pZn21" Then BOD2CAPSIM = 109
If bc = "pZn21g" Then BOD2CAPSIM = 110
If bc = "pZn21t" Then BOD2CAPSIM = 111
If bc = "pZn21tg" Then BOD2CAPSIM = 109
If bc = "pZn21v" Then BOD2CAPSIM = 109
If bc = "pZn21x" Then BOD2CAPSIM = 111
If bc = "pZn23" Then BOD2CAPSIM = 113
If bc = "pZn23g" Then BOD2CAPSIM = 110
If bc = "pZn23gx" Then BOD2CAPSIM = 110
If bc = "pZn23t" Then BOD2CAPSIM = 111
If bc = "pZn23v" Then BOD2CAPSIM = 113
If bc = "pZn23w" Then BOD2CAPSIM = 113
If bc = "pZn23x" Then BOD2CAPSIM = 111
If bc = "pZn23x-F" Then BOD2CAPSIM = 111
If bc = "pZn30" Then BOD2CAPSIM = 114
If bc = "pZn30g" Then BOD2CAPSIM = 114
If bc = "pZn30r" Then BOD2CAPSIM = 114
If bc = "pZn30w" Then BOD2CAPSIM = 114
If bc = "pZn30x" Then BOD2CAPSIM = 114
If bc = "Rd10A" Then BOD2CAPSIM = 119
If bc = "Rd10Ag" Then BOD2CAPSIM = 119
If bc = "Rd10C" Then BOD2CAPSIM = 119
If bc = "Rd10Cg" Then BOD2CAPSIM = 120
If bc = "Rd10Cm" Then BOD2CAPSIM = 119
If bc = "Rd10Cp" Then BOD2CAPSIM = 119
If bc = "Rd90A" Then BOD2CAPSIM = 116
If bc = "Rd90C" Then BOD2CAPSIM = 116
If bc = "Rd90Cg" Then BOD2CAPSIM = 120
If bc = "Rd90Cm" Then BOD2CAPSIM = 116
If bc = "Rd90Cp" Then BOD2CAPSIM = 119
If bc = "Rn14C" Then BOD2CAPSIM = 117
If bc = "Rn15A" Then BOD2CAPSIM = 115
If bc = "Rn15C" Then BOD2CAPSIM = 115
If bc = "Rn15Cg" Then BOD2CAPSIM = 115
If bc = "Rn15Ct" Then BOD2CAPSIM = 115
If bc = "Rn15Cw" Then BOD2CAPSIM = 115
If bc = "Rn42C" Then BOD2CAPSIM = 119
If bc = "Rn42Cg" Then BOD2CAPSIM = 119
If bc = "Rn42Cp" Then BOD2CAPSIM = 119
If bc = "Rn44C" Then BOD2CAPSIM = 117
If bc = "Rn44Cv" Then BOD2CAPSIM = 118
If bc = "Rn44Cw" Then BOD2CAPSIM = 117
If bc = "Rn45A" Then BOD2CAPSIM = 117
If bc = "Rn45C" Then BOD2CAPSIM = 117
If bc = "Rn46A" Then BOD2CAPSIM = 117
If bc = "Rn46Av" Then BOD2CAPSIM = 118
If bc = "Rn46Aw" Then BOD2CAPSIM = 117
If bc = "Rn47C" Then BOD2CAPSIM = 117
If bc = "Rn47Cg" Then BOD2CAPSIM = 120
If bc = "Rn47Cp" Then BOD2CAPSIM = 119
If bc = "Rn47Cv" Then BOD2CAPSIM = 118
If bc = "Rn47Cw" Then BOD2CAPSIM = 117
If bc = "Rn47Cwp" Then BOD2CAPSIM = 119
If bc = "Rn52A" Then BOD2CAPSIM = 120
If bc = "Rn52Ag" Then BOD2CAPSIM = 120
If bc = "Rn62C" Then BOD2CAPSIM = 119
If bc = "Rn62Cg" Then BOD2CAPSIM = 120
If bc = "Rn62Cp" Then BOD2CAPSIM = 119
If bc = "Rn62Cwp" Then BOD2CAPSIM = 119
If bc = "Rn66A" Then BOD2CAPSIM = 117
If bc = "Rn66Av" Then BOD2CAPSIM = 118
If bc = "Rn67C" Then BOD2CAPSIM = 117
If bc = "Rn67Cg" Then BOD2CAPSIM = 120
If bc = "Rn67Cp" Then BOD2CAPSIM = 119
If bc = "Rn67Cv" Then BOD2CAPSIM = 118
If bc = "Rn67Cwp" Then BOD2CAPSIM = 119
If bc = "Rn82A" Then BOD2CAPSIM = 119
If bc = "Rn82Ag" Then BOD2CAPSIM = 120
If bc = "Rn94C" Then BOD2CAPSIM = 117
If bc = "Rn94Cv" Then BOD2CAPSIM = 118
If bc = "Rn95A" Then BOD2CAPSIM = 116
If bc = "Rn95Av" Then BOD2CAPSIM = 118
If bc = "Rn95C" Then BOD2CAPSIM = 116
If bc = "Rn95Cg" Then BOD2CAPSIM = 120
If bc = "Rn95Cm" Then BOD2CAPSIM = 116
If bc = "Rn95Cp" Then BOD2CAPSIM = 119
If bc = "Ro40A" Then BOD2CAPSIM = 117
If bc = "Ro40Av" Then BOD2CAPSIM = 118
If bc = "Ro40C" Then BOD2CAPSIM = 117
If bc = "Ro40Cv" Then BOD2CAPSIM = 118
If bc = "Ro40Cw" Then BOD2CAPSIM = 117
If bc = "Ro60A" Then BOD2CAPSIM = 116
If bc = "Ro60C" Then BOD2CAPSIM = 116
If bc = "ROb72" Then BOD2CAPSIM = 119
If bc = "ROb75" Then BOD2CAPSIM = 116
If bc = "Rv01A" Then BOD2CAPSIM = 118
If bc = "Rv01C" Then BOD2CAPSIM = 118
If bc = "Rv01Cg" Then BOD2CAPSIM = 118
If bc = "Rv01Cp" Then BOD2CAPSIM = 118
If bc = "saVc" Then BOD2CAPSIM = 101
If bc = "saVz" Then BOD2CAPSIM = 102
If bc = "sHn21" Then BOD2CAPSIM = 109
If bc = "shVz" Then BOD2CAPSIM = 102
If bc = "skVc" Then BOD2CAPSIM = 103
If bc = "skWz" Then BOD2CAPSIM = 104
If bc = "Sn13A" Then BOD2CAPSIM = 113
If bc = "Sn13Ap" Then BOD2CAPSIM = 113
If bc = "Sn13Av" Then BOD2CAPSIM = 113
If bc = "Sn13Aw" Then BOD2CAPSIM = 113
If bc = "Sn13Awp" Then BOD2CAPSIM = 113
If bc = "Sn14A" Then BOD2CAPSIM = 113
If bc = "Sn14Ap" Then BOD2CAPSIM = 113
If bc = "Sn14Av" Then BOD2CAPSIM = 113
If bc = "spVc" Then BOD2CAPSIM = 103
If bc = "spVz" Then BOD2CAPSIM = 104
If bc = "sVc" Then BOD2CAPSIM = 101
If bc = "sVk" Then BOD2CAPSIM = 106
If bc = "sVp" Then BOD2CAPSIM = 102
If bc = "sVs" Then BOD2CAPSIM = 101
If bc = "svWp" Then BOD2CAPSIM = 102
If bc = "svWz" Then BOD2CAPSIM = 102
If bc = "svWzt" Then BOD2CAPSIM = 102
If bc = "sVz" Then BOD2CAPSIM = 102
If bc = "sVzt" Then BOD2CAPSIM = 102
If bc = "sVzx" Then BOD2CAPSIM = 102
If bc = "tZd21" Then BOD2CAPSIM = 107
If bc = "tZd21g" Then BOD2CAPSIM = 110
If bc = "tZd21v" Then BOD2CAPSIM = 107
If bc = "tZd23" Then BOD2CAPSIM = 113
If bc = "tZd30" Then BOD2CAPSIM = 114
If bc = "U4546nr005" Then BOD2CAPSIM = 109 'in omschrijving erachter stond cHn21 (veldpodzol, lemig fijn zand)
If bc = "U4546nr113" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond Hn21 (veldpodzol, zwak lemig fijn zand)
If bc = "U4546nr013" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond Hn21g (veldpodzol, zwak lemig fijn zand)
If bc = "U4546nr127" Then BOD2CAPSIM = 112 'in de omschrijving erachter stond zEZ21 (hoge, zwarte enkeerdgronden, leemarm en zwak lemig fijn zand)
If bc = "U4546nr017" Then BOD2CAPSIM = 112 'in de omschrijving erachter stond pZn21g (gooreerdgronden, lemarm en zwak lemig fijn zand)
If bc = "U4546nr003" Then BOD2CAPSIM = 109 'in de omschrijving erachter stond cHn21 (laarpodzolgronden, leemarm en zwak lemig fijn zand)
If bc = "uWz" Then BOD2CAPSIM = 113 'moerige eerdgronden
If bc = "Vb" Then BOD2CAPSIM = 101
If bc = "Vc" Then BOD2CAPSIM = 101
If bc = "Vd" Then BOD2CAPSIM = 101
If bc = "Vk" Then BOD2CAPSIM = 106
If bc = "Vo" Then BOD2CAPSIM = 101
If bc = "Vp" Then BOD2CAPSIM = 102
If bc = "Vpx" Then BOD2CAPSIM = 102
If bc = "Vr" Then BOD2CAPSIM = 101
If bc = "Vs" Then BOD2CAPSIM = 101
If bc = "Vsc" Then BOD2CAPSIM = 101
If bc = "vWp" Then BOD2CAPSIM = 102
If bc = "vWpg" Then BOD2CAPSIM = 102
If bc = "vWpt" Then BOD2CAPSIM = 102
If bc = "vWpx" Then BOD2CAPSIM = 102
If bc = "vWz" Then BOD2CAPSIM = 102
If bc = "vWzg" Then BOD2CAPSIM = 102
If bc = "vWzr" Then BOD2CAPSIM = 102
If bc = "vWzt" Then BOD2CAPSIM = 102
If bc = "vWzx" Then BOD2CAPSIM = 102
If bc = "Vz" Then BOD2CAPSIM = 102
If bc = "Vzc" Then BOD2CAPSIM = 102
If bc = "Vzg" Then BOD2CAPSIM = 102
If bc = "Vzt" Then BOD2CAPSIM = 102
If bc = "Vzx" Then BOD2CAPSIM = 102
If bc = "Wg" Then BOD2CAPSIM = 106
If bc = "Wgl" Then BOD2CAPSIM = 106
If bc = "Wo" Then BOD2CAPSIM = 106
If bc = "Wol" Then BOD2CAPSIM = 106
If bc = "Wov" Then BOD2CAPSIM = 106
If bc = "Y21" Then BOD2CAPSIM = 109
If bc = "Y21g" Then BOD2CAPSIM = 110
If bc = "Y21x" Then BOD2CAPSIM = 111
If bc = "Y23" Then BOD2CAPSIM = 113
If bc = "Y23b" Then BOD2CAPSIM = 113
If bc = "Y23g" Then BOD2CAPSIM = 110
If bc = "Y23x" Then BOD2CAPSIM = 111
If bc = "Y30" Then BOD2CAPSIM = 114
If bc = "Y30x" Then BOD2CAPSIM = 114
If bc = "Zb20A" Then BOD2CAPSIM = 107
If bc = "Zb21" Then BOD2CAPSIM = 109
If bc = "Zb21g" Then BOD2CAPSIM = 110
If bc = "Zb23" Then BOD2CAPSIM = 113
If bc = "Zb23g" Then BOD2CAPSIM = 113
If bc = "Zb23t" Then BOD2CAPSIM = 111
If bc = "Zb23x" Then BOD2CAPSIM = 111
If bc = "Zb30" Then BOD2CAPSIM = 114
If bc = "Zb30A" Then BOD2CAPSIM = 114
If bc = "Zb30g" Then BOD2CAPSIM = 114
If bc = "Zd20A" Then BOD2CAPSIM = 107
If bc = "Zd20Ab" Then BOD2CAPSIM = 107
If bc = "Zd21" Then BOD2CAPSIM = 107
If bc = "Zd21g" Then BOD2CAPSIM = 107
If bc = "Zd23" Then BOD2CAPSIM = 113
If bc = "Zd30" Then BOD2CAPSIM = 114
If bc = "Zd30A" Then BOD2CAPSIM = 114
If bc = "zEZ21" Then BOD2CAPSIM = 112
If bc = "zEZ21g" Then BOD2CAPSIM = 112
If bc = "zEZ21t" Then BOD2CAPSIM = 112
If bc = "zEZ21w" Then BOD2CAPSIM = 112
If bc = "zEZ21x" Then BOD2CAPSIM = 112
If bc = "zEZ23" Then BOD2CAPSIM = 112
If bc = "zEZ23g" Then BOD2CAPSIM = 112
If bc = "zEZ23t" Then BOD2CAPSIM = 112
If bc = "zEZ23w" Then BOD2CAPSIM = 112
If bc = "zEZ23x" Then BOD2CAPSIM = 112
If bc = "zEZ30" Then BOD2CAPSIM = 112
If bc = "zEZ30g" Then BOD2CAPSIM = 112
If bc = "zEZ30x" Then BOD2CAPSIM = 112
If bc = "zgHd30" Then BOD2CAPSIM = 114
If bc = "zgMn15C" Then BOD2CAPSIM = 115
If bc = "zgMn88C" Then BOD2CAPSIM = 117
If bc = "zgY30" Then BOD2CAPSIM = 114
If bc = "zHd21" Then BOD2CAPSIM = 108
If bc = "zHd21g" Then BOD2CAPSIM = 108
If bc = "zHn21" Then BOD2CAPSIM = 108
If bc = "zHn23" Then BOD2CAPSIM = 109
If bc = "zhVk" Then BOD2CAPSIM = 106
If bc = "zKRn1g" Then BOD2CAPSIM = 120
If bc = "zKRn2" Then BOD2CAPSIM = 119
If bc = "zkVc" Then BOD2CAPSIM = 103
If bc = "zkWp" Then BOD2CAPSIM = 104
If bc = "zMn15A" Then BOD2CAPSIM = 115
If bc = "zMn22Ap" Then BOD2CAPSIM = 119
If bc = "zMn25Ap" Then BOD2CAPSIM = 119
If bc = "zMn56Cp" Then BOD2CAPSIM = 117
If bc = "zMo10A" Then BOD2CAPSIM = 115
If bc = "zMv41C" Then BOD2CAPSIM = 118
If bc = "zMv61C" Then BOD2CAPSIM = 118
If bc = "Zn10A" Then BOD2CAPSIM = 107
If bc = "Zn10Ap" Then BOD2CAPSIM = 107
If bc = "Zn10Av" Then BOD2CAPSIM = 107
If bc = "Zn10Aw" Then BOD2CAPSIM = 107
If bc = "Zn10Awp" Then BOD2CAPSIM = 107
If bc = "Zn21" Then BOD2CAPSIM = 107
If bc = "Zn21-F" Then BOD2CAPSIM = 107
If bc = "Zn21g" Then BOD2CAPSIM = 107
If bc = "Zn21-H" Then BOD2CAPSIM = 107
If bc = "Zn21p" Then BOD2CAPSIM = 107
If bc = "Zn21r" Then BOD2CAPSIM = 107
If bc = "Zn21t" Then BOD2CAPSIM = 107
If bc = "Zn21v" Then BOD2CAPSIM = 107
If bc = "Zn21w" Then BOD2CAPSIM = 107
If bc = "Zn21x" Then BOD2CAPSIM = 107
If bc = "Zn21x-F" Then BOD2CAPSIM = 107
If bc = "Zn23" Then BOD2CAPSIM = 113
If bc = "Zn23-F" Then BOD2CAPSIM = 113
If bc = "Zn23g" Then BOD2CAPSIM = 113
If bc = "Zn23g-F" Then BOD2CAPSIM = 113
If bc = "Zn23-H" Then BOD2CAPSIM = 113
If bc = "Zn23p" Then BOD2CAPSIM = 113
If bc = "Zn23r" Then BOD2CAPSIM = 113
If bc = "Zn23t" Then BOD2CAPSIM = 111
If bc = "Zn23x" Then BOD2CAPSIM = 111
If bc = "Zn30" Then BOD2CAPSIM = 114
If bc = "Zn30A" Then BOD2CAPSIM = 114
If bc = "Zn30Ab" Then BOD2CAPSIM = 114
If bc = "Zn30Ag" Then BOD2CAPSIM = 114
If bc = "Zn30Ar" Then BOD2CAPSIM = 114
If bc = "Zn30g" Then BOD2CAPSIM = 114
If bc = "Zn30r" Then BOD2CAPSIM = 114
If bc = "Zn30v" Then BOD2CAPSIM = 114
If bc = "Zn30x" Then BOD2CAPSIM = 114
If bc = "Zn40A" Then BOD2CAPSIM = 107
If bc = "Zn40Ap" Then BOD2CAPSIM = 107
If bc = "Zn40Ar" Then BOD2CAPSIM = 107
If bc = "Zn40Av" Then BOD2CAPSIM = 107
If bc = "Zn50A" Then BOD2CAPSIM = 107
If bc = "Zn50Ab" Then BOD2CAPSIM = 107
If bc = "Zn50Ap" Then BOD2CAPSIM = 107
If bc = "Zn50Ar" Then BOD2CAPSIM = 107
If bc = "Zn50Aw" Then BOD2CAPSIM = 107
If bc = "zpZn23w" Then BOD2CAPSIM = 113
If bc = "zRd10A" Then BOD2CAPSIM = 119
If bc = "zRn15C" Then BOD2CAPSIM = 115
If bc = "zRn47Cwp" Then BOD2CAPSIM = 117
If bc = "zRn62C" Then BOD2CAPSIM = 119
If bc = "zSn14A" Then BOD2CAPSIM = 113
If bc = "zVc" Then BOD2CAPSIM = 105
If bc = "zVp" Then BOD2CAPSIM = 105
If bc = "zVpg" Then BOD2CAPSIM = 105
If bc = "zVpt" Then BOD2CAPSIM = 105
If bc = "zVpx" Then BOD2CAPSIM = 105
If bc = "zVs" Then BOD2CAPSIM = 105
If bc = "zVz" Then BOD2CAPSIM = 105
If bc = "zVzg" Then BOD2CAPSIM = 105
If bc = "zVzt" Then BOD2CAPSIM = 105
If bc = "zVzx" Then BOD2CAPSIM = 105
If bc = "zWp" Then BOD2CAPSIM = 105
If bc = "zWpg" Then BOD2CAPSIM = 105
If bc = "zWpt" Then BOD2CAPSIM = 105
If bc = "zWpx" Then BOD2CAPSIM = 105
If bc = "zWz" Then BOD2CAPSIM = 105
If bc = "zWzg" Then BOD2CAPSIM = 105
If bc = "zWzt" Then BOD2CAPSIM = 105
If bc = "zWzx" Then BOD2CAPSIM = 105
If bc = "zY21" Then BOD2CAPSIM = 108
If bc = "zY21g" Then BOD2CAPSIM = 108
If bc = "zY23" Then BOD2CAPSIM = 109
If bc = "zY30" Then BOD2CAPSIM = 114

End Function


Public Sub RunDoEvents(PauseTime As Long)
  Dim start, Finish, TotalTime
  start = Timer ' Set start time.
  Do While Timer < start + PauseTime
    DoEvents ' Yield to other processes.
  Loop
End Sub

Public Function ShellandWait(ExeFullPath As String, _
                                Optional TimeOutValue As Long = 0, _
                                Optional CheckReturnCode As Boolean = False, Optional ByRef ReturnCodeFile As String) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    Dim ExeDirectory As String
    
    On Error GoTo errorhandler

    'paths with .'s and or spaces go wrong. So fix it here by surrounding them with double quotes
    lStart = CLng(Timer)
    sExeName = Trim(ExeFullPath)
    
    'set the directory where the executable resides as the current dir
    ExeDirectory = GetDirectory(sExeName)
    Call ChDir(ExeDirectory)
    
    If Left(sExeName, 1) <> Chr(34) Or Right(sExeName, 1) <> Chr(34) Then
      sExeName = Chr(34) & sExeName
      sExeName = sExeName & Chr(34)
    End If

    'Deal with timeout being reset at VBA.Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbMinimizedNoFocus)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    
    If CheckReturnCode Then
      If FileExists(ReturnCodeFile) Then
        Dim hf As Long
        Dim st As String
        hf = FreeFile
        Open ReturnCodeFile For Input As #hf
        Input #hf, st
        Close #hf
        
        ShellandWait = (VAL(st) = 0)
      Else
        ShellandWait = False
      End If
    Else
      ShellandWait = True
    End If
    
    Exit Function
   
errorhandler:
ShellandWait = False
Exit Function
End Function

Public Function HydroZomerWinter(myDate As Variant, Optional SkipFromMonth As Long = 0, Optional SkipFromDay As Long = 0, Optional SkipToMonth As Integer = 0, Optional SkipToDay As Integer = 0) As String
  
  'integriteitscontrole
  If SkipToMonth < SkipFromMonth Then
    HydroZomerWinter = "error in function HydroZomerWinter"
    Exit Function
  End If
  
  'check eerst of hij geskipped moet worden
  If month(myDate) = SkipFromMonth Then
    If day(myDate) >= SkipFromDay Then
      HydroZomerWinter = "overgeslagen"
      Exit Function
    End If
  ElseIf month(myDate) = SkipToMonth Then
    If day(myDate) <= SkipToDay Then
      HydroZomerWinter = "overgeslagen"
      Exit Function
    End If
  ElseIf month(myDate) > SkipFromMonth And month(myDate) < SkipToMonth Then
    HydroZomerWinter = "overgeslagen"
    Exit Function
  End If
  
  Select Case month(myDate)
  Case Is = 1
    HydroZomerWinter = "winter"
  Case Is = 2
    HydroZomerWinter = "winter"
  Case Is = 3
    HydroZomerWinter = "winter"
  Case Is = 4
    If day(myDate) < 15 Then
      HydroZomerWinter = "winter"
    Else
      HydroZomerWinter = "zomer"
    End If
  Case Is = 5
    HydroZomerWinter = "zomer"
  Case Is = 6
    HydroZomerWinter = "zomer"
  Case Is = 7
    HydroZomerWinter = "zomer"
  Case Is = 8
    HydroZomerWinter = "zomer"
  Case Is = 9
    HydroZomerWinter = "zomer"
  Case Is = 10
    If day(myDate) < 15 Then
      HydroZomerWinter = "zomer"
    Else
      HydroZomerWinter = "winter"
    End If
  Case Is = 11
    HydroZomerWinter = "winter"
  Case Is = 12
    HydroZomerWinter = "winter"
  End Select
End Function

Public Function EVAPMAKKINK(Kin As Variant, Tdag As Variant, Tmin As Variant, Tmax As Variant, P As Variant) As Variant
  Dim esat As Variant
  Dim s As Variant
  Dim y As Variant
  Dim lambdaE As Variant

  'lambda = verdampingwarmte van water (2.45E06 J/kg)
  'E = verdampingsflux (kg/m2/s)
  'a_accent = constante (ongeveer 1.1)
  'Rn = nettostraling (W/m2)
  's = afgeleide van ew bij luchttemperatuur T (Pa/K), dus s = dew/dT
  'y = psychrometerconstante in Pa/K
  'G = bodemwarmtestroom
  'Beta = 10 W/m2


  esat = Verzadigingsdampdruk(Tmin, Tmax)
  s = DampDrukGradient(esat, Tdag)
  y = PsychrometerConstante(P, VerdampingswarmteWater(Tdag))

  lambdaE = 0.65 * s / (s + y) * Kin
  
  'converteer naar mm/d
  EVAPMAKKINK = lambdaE / 2450000 / 1000 * 1000 * 3600 * 24 ' =lmbdaE * 0.035

End Function

Public Sub MAKKINK2OPENWATER(startRow As Integer, DateCol As Integer, ValCol As Integer, ResultsRow As Integer, ResultsCol As Integer)
  'This routine converts evaporation according to Makkink (referentiegewasverdamping) to evaporation of openwater bodies
  
  Dim r As Long, c_dat As Long, c_val As Long
  Dim r_res As Long, c_res As Long
  Dim myDate As Date, myVal As Variant
  r = startRow - 1
  c_dat = DateCol
  c_val = ValCol
  r_res = ResultsRow - 1
  c_res = ResultsCol
  
  While Not ActiveSheet.Cells(r + 1, c_dat) = ""
    r = r + 1
    r_res = r_res + 1
    
    myDate = ActiveSheet.Cells(r, c_dat)
    myVal = ActiveSheet.Cells(r, c_val)
    
    ActiveSheet.Cells(r_res, c_res) = myDate
    ActiveSheet.Cells(r_res, c_res + 1) = myVal * OPENWATEREVAPFACTOR(myDate)
  Wend

End Sub

Public Sub EVAPDAY2HOUR(startRow As Integer, DateCol As Integer, ValCol As Integer, ResultsRow As Integer, ResultsCol As Integer)
  'spreads daily evaporation sum out over 24 hours within the day
  Dim r As Long, c_dat As Long, c_val As Long
  Dim r_res As Long, c_res As Long, i As Integer
  Dim myDate As Date, myVal As Variant
  r = startRow - 1
  c_dat = DateCol
  c_val = ValCol
  r_res = ResultsRow - 1
  c_res = ResultsCol
  
  While Not ActiveSheet.Cells(r + 1, c_dat) = ""
    r = r + 1
    myDate = ActiveSheet.Cells(r, c_dat)
    myVal = ActiveSheet.Cells(r, c_val)
    
    For i = 1 To 24
     r_res = r_res + 1
     ActiveSheet.Cells(r_res, c_res) = myDate + (i - 1) / 24
     ActiveSheet.Cells(r_res, c_res + 1) = myVal * HOURLYEVAPORATIONFRACTION(i)
    Next
    
  Wend
End Sub

Public Function HOURLYEVAPORATIONFRACTION(H As Integer) As Variant

  If H <= 6 Then
    HOURLYEVAPORATIONFRACTION = 0
  ElseIf H >= 18 Then
    HOURLYEVAPORATIONFRACTION = 0
  ElseIf H = 7 Or H = 17 Then
    HOURLYEVAPORATIONFRACTION = 0.03
  ElseIf H = 8 Or H = 16 Then
    HOURLYEVAPORATIONFRACTION = 0.07
  ElseIf H = 9 Or H = 15 Then
    HOURLYEVAPORATIONFRACTION = 0.09
  ElseIf H = 10 Or H = 14 Then
    HOURLYEVAPORATIONFRACTION = 0.11
  ElseIf H = 11 Or H = 13 Then
    HOURLYEVAPORATIONFRACTION = 0.13
  ElseIf H = 12 Then
    HOURLYEVAPORATIONFRACTION = 0.14
  Else
  HOURLYEVAPORATIONFRACTION = 0
  End If
  
End Function

Public Function OPENWATEREVAPFACTOR(myDate As Date) As Variant
  'retrieves the openwater evaporation multiplication w.r.t. Makkink evaporation for a given date
  
  Dim MonthDay As String
  MonthDay = VBA.Trim(VBA.str(month(myDate))) & "_" & VBA.Trim(VBA.str(day(myDate)))
  
  
  Select Case MonthDay
Case Is = "1_1"
  OPENWATEREVAPFACTOR = 0.5
Case Is = "1_2"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_3"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_4"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_5"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_6"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_7"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_8"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_9"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_10"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_11"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_12"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_13"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_14"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_15"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_16"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_17"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_18"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_19"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_20"
OPENWATEREVAPFACTOR = 0.5
Case Is = "1_21"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_22"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_23"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_24"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_25"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_26"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_27"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_28"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_29"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_30"
OPENWATEREVAPFACTOR = 0.7
Case Is = "1_31"
OPENWATEREVAPFACTOR = 0.7
Case Is = "2_1"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_2"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_3"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_4"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_5"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_6"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_7"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_8"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_9"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_10"
OPENWATEREVAPFACTOR = 0.8
Case Is = "2_11"
OPENWATEREVAPFACTOR = 1
Case Is = "2_12"
OPENWATEREVAPFACTOR = 1
Case Is = "2_13"
OPENWATEREVAPFACTOR = 1
Case Is = "2_14"
OPENWATEREVAPFACTOR = 1
Case Is = "2_15"
OPENWATEREVAPFACTOR = 1
Case Is = "2_16"
OPENWATEREVAPFACTOR = 1
Case Is = "2_17"
OPENWATEREVAPFACTOR = 1
Case Is = "2_18"
OPENWATEREVAPFACTOR = 1
Case Is = "2_19"
OPENWATEREVAPFACTOR = 1
Case Is = "2_20"
OPENWATEREVAPFACTOR = 1
Case Is = "2_21"
OPENWATEREVAPFACTOR = 1
Case Is = "2_22"
OPENWATEREVAPFACTOR = 1
Case Is = "2_23"
OPENWATEREVAPFACTOR = 1
Case Is = "2_24"
OPENWATEREVAPFACTOR = 1
Case Is = "2_25"
OPENWATEREVAPFACTOR = 1
Case Is = "2_26"
OPENWATEREVAPFACTOR = 1
Case Is = "2_27"
OPENWATEREVAPFACTOR = 1
Case Is = "2_28"
OPENWATEREVAPFACTOR = 1
Case Is = "2_29"
OPENWATEREVAPFACTOR = 1
Case Is = "3_1"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_2"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_3"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_4"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_5"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_6"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_7"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_8"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_9"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_10"
OPENWATEREVAPFACTOR = 1.2
Case Is = "3_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "3_31"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_1"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_2"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_3"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_4"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_5"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_6"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_7"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_8"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_9"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_10"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "4_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_1"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_2"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_3"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_4"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_5"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_6"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_7"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_8"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_9"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_10"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_11"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_12"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_13"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_14"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_15"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_16"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_17"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_18"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_19"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_20"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_21"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_22"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_23"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_24"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_25"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_26"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_27"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_28"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_29"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_30"
OPENWATEREVAPFACTOR = 1.3
Case Is = "5_31"
OPENWATEREVAPFACTOR = 1.3
Case Is = "6_1"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_2"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_3"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_4"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_5"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_6"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_7"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_8"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_9"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_10"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_11"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_12"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_13"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_14"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_15"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_16"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_17"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_18"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_19"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_20"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_21"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_22"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_23"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_24"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_25"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_26"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_27"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_28"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_29"
OPENWATEREVAPFACTOR = 1.31
Case Is = "6_30"
OPENWATEREVAPFACTOR = 1.31
Case Is = "7_1"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_2"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_3"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_4"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_5"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_6"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_7"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_8"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_9"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_10"
OPENWATEREVAPFACTOR = 1.29
Case Is = "7_11"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_12"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_13"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_14"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_15"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_16"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_17"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_18"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_19"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_20"
OPENWATEREVAPFACTOR = 1.27
Case Is = "7_21"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_22"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_23"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_24"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_25"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_26"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_27"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_28"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_29"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_30"
OPENWATEREVAPFACTOR = 1.24
Case Is = "7_31"
OPENWATEREVAPFACTOR = 1.24
Case Is = "8_1"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_2"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_3"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_4"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_5"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_6"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_7"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_8"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_9"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_10"
OPENWATEREVAPFACTOR = 1.21
Case Is = "8_11"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_12"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_13"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_14"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_15"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_16"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_17"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_18"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_19"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_20"
OPENWATEREVAPFACTOR = 1.19
Case Is = "8_21"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_22"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_23"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_24"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_25"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_26"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_27"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_28"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_29"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_30"
OPENWATEREVAPFACTOR = 1.18
Case Is = "8_31"
OPENWATEREVAPFACTOR = 1.18
Case Is = "9_1"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_2"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_3"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_4"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_5"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_6"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_7"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_8"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_9"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_10"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_11"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_12"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_13"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_14"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_15"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_16"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_17"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_18"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_19"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_20"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_21"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_22"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_23"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_24"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_25"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_26"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_27"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_28"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_29"
OPENWATEREVAPFACTOR = 1.17
Case Is = "9_30"
OPENWATEREVAPFACTOR = 1.17
Case Is = "10_1"
OPENWATEREVAPFACTOR = 1
Case Is = "10_2"
OPENWATEREVAPFACTOR = 1
Case Is = "10_3"
OPENWATEREVAPFACTOR = 1
Case Is = "10_4"
OPENWATEREVAPFACTOR = 1
Case Is = "10_5"
OPENWATEREVAPFACTOR = 1
Case Is = "10_6"
OPENWATEREVAPFACTOR = 1
Case Is = "10_7"
OPENWATEREVAPFACTOR = 1
Case Is = "10_8"
OPENWATEREVAPFACTOR = 1
Case Is = "10_9"
OPENWATEREVAPFACTOR = 1
Case Is = "10_10"
OPENWATEREVAPFACTOR = 1
Case Is = "10_11"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_12"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_13"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_14"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_15"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_16"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_17"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_18"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_19"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_20"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_21"
OPENWATEREVAPFACTOR = 0.9
Case Is = "10_22"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_23"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_24"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_25"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_26"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_27"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_28"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_29"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_30"
OPENWATEREVAPFACTOR = 0.8
Case Is = "10_31"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_1"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_2"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_3"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_4"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_5"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_6"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_7"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_8"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_9"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_10"
OPENWATEREVAPFACTOR = 0.8
Case Is = "11_11"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_12"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_13"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_14"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_15"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_16"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_17"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_18"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_19"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_20"
OPENWATEREVAPFACTOR = 0.7
Case Is = "11_21"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_22"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_23"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_24"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_25"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_26"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_27"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_28"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_29"
OPENWATEREVAPFACTOR = 0.6
Case Is = "11_30"
OPENWATEREVAPFACTOR = 0.6
Case Is = "12_1"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_2"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_3"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_4"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_5"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_6"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_7"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_8"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_9"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_10"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_11"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_12"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_13"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_14"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_15"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_16"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_17"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_18"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_19"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_20"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_21"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_22"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_23"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_24"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_25"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_26"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_27"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_28"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_29"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_30"
OPENWATEREVAPFACTOR = 0.5
Case Is = "12_31"
OPENWATEREVAPFACTOR = 0.5
Case Else
OPENWATEREVAPFACTOR = 0
End Select
  
End Function

Public Function EVAPDEBRUINKEIJMAN(Datum As Variant, Kin As Variant, Tdag As Variant, Tmin As Variant, Tmax As Variant, P As Variant, SP As Variant, UG As Variant, d As Variant) As Variant
  
  'Kin = dagsom globale straling (W/m2)
  'Tdag = gemiddelde dagtemperatuur (Celcius en K)
  'Tmin = minimum dagtemperatuur (celcius en K)
  'Tmax = maximum dagtemperatuur (Celcius en K)
  'p = luchtdruk kPa op hoogte z0
  'SP = percentage v d langst mogelijke zonneschijn
  'UG = percentage luchtvochtigheid
  'd = gemiddelde waterdiepte over het gebied
  
  'a_accent = constante (-)
  'beta = 10 W/m2
  'dTdt() verandering watertemperatuur in de tijd K/s.
  'RN = nettostraling in W/m2
  'literatuur:
  'Futurewater, 2006, Berekening openwaterverdamping, in opdracht van wetterskip fryslan
  'STOWA 2009, verbetering bepaling actuele verdamping voor het strategisch waterbeheer, definitiestudie
  
  'OmrekeningSfAcTOren vOOr mm, mJ en WATTS.
  '                       mm d^(-1)             mJ * m^(-2) * d^(-1)  W m^(-2)
  '  1 mm d-1             1.000                 2.451                 28.368
  '  1 MJ m-2 d-1         0.408                 1.000                 11.574
  '  1 W m-2              0.035                 0.086                 1.000
  
  Dim a_accent As Variant    'ongeveer 1.1
  Dim beta As Variant        'beta = 10 W/m2
  Dim dTdt(1 To 12) As Variant 'verandering watertemperatuur in de tijd K/s.
  Dim g As Variant
  Dim lambdaE As Variant     'de verdampingswarmteflux. converteren we m.b.v. de tabel naar mm/d
  Dim Rn As Variant          'nettostraling in W/m2
  Dim esat As Variant        'verzadigingsdampdruk (hPA)
  Dim ez As Variant          'dampdruk (hPa)
  Dim Nrel As Variant        'relatieve zonneschijn
  Dim LNetto As Variant      'netto langgolvige straling
  Dim y As Variant           'psychrometerconstante
  Dim s As Variant           'dampdrukgradiënt
  Dim RH As Variant          'relative humidity (-)
  
  'vergelijking van de bruin-Keijman:
  'lambda * E = a_accent * s/(s + y) *(Rn - G) + Beta
  
  'waarin:
  'lambda = verdampingwarmte van water (2.45E06 J/kg)
  'E = verdampingsflux (kg/m2/s)
  'a_accent = constante (ongeveer 1.1)
  'Rn = nettostraling (W/m2)
  's = afgeleide van ew bij luchttemperatuur T (Pa/K), dus s = dew/dT
  'y = psychrometerconstante in Pa/K
  'G = bodemwarmtestroom
  'Beta = 10 W/m2
  
  '-------------------------------------------------------------------
  'bereken eerst de bodemwarmtestroom G (W/m2)
  g = BodemWarmteStroom(Datum, d)
    
  'bereken verzadigingsdampdruk esat (hPa) en dampdruk e(z) op hoogte z
  esat = Verzadigingsdampdruk(Tmin, Tmax)
  RH = UG / 100     'relatieve luchtvochtigheid
  ez = RH * esat
  
  'bereken de relatieve zonneschijnduur en de netto langgolvige straling (in W/m2)
  Nrel = SP / 100
  LNetto = NettoLanggolvigeStraling(Tmax, Tmin, Nrel, ez)
  
  'bereken Rn (netto straling) (W/m2)
  Rn = NettoStraling(Kin, LNetto)
  
  a_accent = 1.1
  beta = 10
  y = PsychrometerConstante(P, VerdampingswarmteWater(Tdag))
  s = DampDrukGradient(esat, Tdag)

  lambdaE = a_accent * s / (s + y) * (Rn - g) + beta
  
  'converteer naar mm/d
  EVAPDEBRUINKEIJMAN = 0.035 * lambdaE
  
End Function

Public Function Verzadigingsdampdruk(Tmin As Variant, Tmax As Variant) As Variant
  'rekent de verzadigingsdampdruk uit a.d.h.v. minimum en maximum luchttemperatuur
  'eenheid: kPa
  Verzadigingsdampdruk = 0.305 * (Exp(17.27 * Tmin / Celcius2Kelvin(Tmin)) + Exp(17.27 * Tmax / Celcius2Kelvin(Tmax)))
End Function

Public Function BodemWarmteStroom(Datum As Variant, d As Variant) As Variant
  
  'deze functie berekent de bodemwarmtestroom G in W/m2 ten behoeve van de
  'berekening van openwaterverdamping met de De Bruin, Keijman - formule
    
  Dim dTdt(1 To 12) As Variant
  Dim rho_water As Variant   'dichtheid water = 1000 kg/m3
  Dim c_water As Variant     'soortelijke warmte water = 4200 J/kg/K
  'd = gemiddelde waterdiepte in het gebied
  
  rho_water = 1000
  c_water = 4200
  
  'temperatuurveranderingen in de tijd K/s
  'bron: Futurewater, 2006 Tabel A.2
  dTdt(1) = -0.000000746714
  dTdt(2) = 0.000000373357
  dTdt(3) = 0.00000119732
  dTdt(4) = 0.00000112007
  dTdt(5) = 0.00000192901
  dTdt(6) = 0.00000112007
  dTdt(7) = 0.000000385802
  dTdt(8) = 0.000000373357
  dTdt(9) = -0.00000112007
  dTdt(10) = -0.00000115741
  dTdt(11) = -0.00000224014
  dTdt(12) = -0.00000115741

  BodemWarmteStroom = rho_water * c_water * d * dTdt(month(Datum))
   
End Function

Public Function NettoStraling(Kin As Variant, LNetto As Variant) As Variant
'deze functie berekent de nettostraling in W/m2 ten behoeve van verdampingsberekeningen

Dim albedo As Variant
albedo = 0.06 'voor water

NettoStraling = (1 - albedo) * Kin - LNetto

End Function

Public Function NettoLanggolvigeStraling(Tmax As Variant, Tmin As Variant, Nrel As Variant, ez As Variant) As Variant
  'deze functie berekent de netto langgolvige straling t.b.v. verdampingsberekeningen
  'in W/m2
  
  'Tmax = maximale dagtemperatuur
  'Tmin =  minimale dagtemperatuur
  'ez = dampdruk op hoogte z (kPa)
  'Nrel = relatieve zonneschijnduur (-)
  
  Dim sbconst As Variant 'stephan bolzmann constante
  sbconst = 0.000000004903 'MJ/K^4/m2/d
  NettoLanggolvigeStraling = sbconst * ((Celcius2Kelvin(Tmax) ^ 4 + Celcius2Kelvin(Tmin) ^ 4) / 2) * (0.34 - 0.14 * Sqr(ez)) * (0.1 + 0.9 * Nrel) * 11.574

End Function

Public Function DampDrukGradient(esat As Variant, Tdag As Variant) As Variant
  'deze functie berekent de dampdrukgradiënt s bij gemiddelde dagluchttemperatuur T zoals die gebruikt wordt bij verdampingsberekeningen
  'eenheid: kPa/K
  
  DampDrukGradient = 4098 * esat / Celcius2Kelvin(Tdag) ^ 2
End Function

Public Function PsychrometerConstante(P As Variant, lambda As Variant) As Variant
  'deze functie berekent de psychrometerconstante y
  'lambda = verdampingswarmte van water bij de gemiddelde dagtemperatuur (MJ/kg)
  'p = luchtdruk (hPa) op hoogte z0
  
  PsychrometerConstante = 0.00163 * P / lambda

End Function

Public Function VerdampingswarmteWater(Tdag As Variant) As Variant
  'deze functie berekent de verdampingswarmte van water (lambda) in MJ/kg bij daggemiddelde temperatuur T in celcius
  VerdampingswarmteWater = 2.501 - 0.002361 * Tdag
End Function

Public Sub MakeScatterChart(XaxisTitle As String, YAxisTitle As String, MeasTimeRange As Range, MeasDataRange As Range, SobekTimeRange As Range, SobekDataRange As Range, Title As String, Optional minX As Variant = -999, Optional maxX As Variant = -999)
    
  Charts.Add
  With ActiveChart
    
    'maak de eerste sobek case de basis voor deze grafiek
    '.ChartType = xlXYScatterLinesNoMarkers
    .ChartType = xlXYScatter
    .SetSourceData Source:=Union(SobekTimeRange, SobekDataRange), PlotBy:=xlColumns
    
    Call .SetElement(msoElementChartTitleAboveChart)
    '.HasTitle = True
    .ChartTitle.Text = Title
    
    .Axes(xlValue).CrossesAt = -1000  'zorg dat de x-as altijd zo laag mogelijk de y-as snijdt
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Characters.Text = XaxisTitle
    .Axes(xlCategory).TickLabels.NumberFormat = "dd/mm/yy"
    
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YAxisTitle
    .Name = Title
    
    'voeg SOBEK resultaten toe
    .SeriesCollection.NewSeries
    .SeriesCollection(1).Name = "Berekend"
    .SeriesCollection(1).XValues = SobekTimeRange
    .SeriesCollection(1).values = SobekDataRange
    .SeriesCollection(1).MarkerSize = 5
    .SeriesCollection(1).MarkerStyle = xlMarkerStyleDash
    
    'tot slot, voeg de meetgegevens toe als serie aan de grafiek
    .SeriesCollection.NewSeries
    .SeriesCollection(2).ChartType = xlXYScatter
    .SeriesCollection(2).Name = "Gemeten"
    .SeriesCollection(2).XValues = MeasTimeRange
    .SeriesCollection(2).values = MeasDataRange
      
    'opmaak
    .SeriesCollection(2).MarkerBackgroundColorIndex = 40
    .SeriesCollection(2).MarkerForegroundColorIndex = 3
    .SeriesCollection(2).MarkerStyle = xlMarkerStyleDash
    .SeriesCollection(2).Smooth = False
    .SeriesCollection(2).MarkerSize = 5
    .SeriesCollection(2).Shadow = False
    
    If minX <> -999 Then .Axes(xlCategory).MinimumScale = minX
    If maxX <> -999 Then .Axes(xlCategory).MaximumScale = maxX
    
  End With
  

End Sub

Public Sub MakeChart(XaxisTitle As String, YAxisTitle As String, MeasTimeRange As Range, MeasDataRange As Range, SobekTimeRange As Range, SobekDataRange As Range, Title As String, Optional minX As Variant = -999, Optional maxX As Variant = -999)
    
  Charts.Add
  With ActiveChart
    
    'maak de eerste sobek case de basis voor deze grafiek
    .ChartType = xlXYScatterLinesNoMarkers
    .SetSourceData Source:=Union(SobekTimeRange, SobekDataRange), PlotBy:=xlColumns
    
    Call .SetElement(msoElementChartTitleAboveChart)
    '.HasTitle = True
    .ChartTitle.Text = Title
    
    .Axes(xlValue).CrossesAt = -1000  'zorg dat de x-as altijd zo laag mogelijk de y-as snijdt
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Characters.Text = XaxisTitle
    .Axes(xlCategory).TickLabels.NumberFormat = "dd/mm/yy"
    
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YAxisTitle
    .Name = Title
    .Axes(xlCategory, xlPrimary).TickLabels.Orientation = xlUpward
    
    'voeg SOBEK resultaten toe
    .SeriesCollection.NewSeries
    .SeriesCollection(1).Name = "Berekend"
    .SeriesCollection(1).XValues = SobekTimeRange
    .SeriesCollection(1).values = SobekDataRange
    
    'tot slot, voeg de meetgegevens toe als serie aan de grafiek
    .SeriesCollection.NewSeries
    .SeriesCollection(2).ChartType = xlXYScatter
    .SeriesCollection(2).Name = "Gemeten"
    .SeriesCollection(2).XValues = MeasTimeRange
    .SeriesCollection(2).values = MeasDataRange
      
    'opmaak
    .SeriesCollection(2).MarkerBackgroundColorIndex = 40
    .SeriesCollection(2).MarkerForegroundColorIndex = 3
    .SeriesCollection(2).MarkerStyle = xlCircle
    .SeriesCollection(2).Smooth = False
    .SeriesCollection(2).MarkerSize = 2
    .SeriesCollection(2).Shadow = False
    
    If minX <> -999 Then .Axes(xlCategory).MinimumScale = minX
    If maxX <> -999 Then .Axes(xlCategory).MaximumScale = maxX
    
  End With
  
End Sub

Sub ExportChart(ChartIndex As Integer, myFileNameNoExtension As String)
    
    Dim myChart As Chart
    Set myChart = ActiveWorkbook.Charts(ChartIndex)

    'myFileName = "myChart.png"
    

    On Error Resume Next
    Kill ThisWorkbook.path & "\" & myFileNameNoExtension
    On Error GoTo 0

    myChart.Export fileName:=ThisWorkbook.path & "\" & myFileNameNoExtension & ".png", Filtername:="PNG"

    MsgBox "OK"
End Sub

Public Function FileExists(path As String) As Boolean
  'controleert of een bestand bestaat
  If VBA.Trim(path) = "" Then
    FileExists = False
  Else
    FileExists = (VBA.dir(path) > "")
  End If
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Public Function IB(Bruto As Variant) As Variant

Dim SchaalMax(1 To 3) As Variant
Dim SchaalPerc(1 To 4) As Variant

SchaalMax(1) = 18628
SchaalMax(2) = 33436
SchaalMax(3) = 55694

SchaalPerc(1) = 33
SchaalPerc(2) = 41.95
SchaalPerc(3) = 42
SchaalPerc(3) = 52

If Bruto <= SchaalMax(1) Then
  IB = IB + Bruto * SchaalPerc(1) / 100
ElseIf Bruto <= SchaalMax(2) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (Bruto - SchaalMax(1)) * SchaalPerc(2) / 100
ElseIf Bruto <= SchaalMax(3) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (SchaalMax(2) - SchaalMax(1)) * SchaalPerc(2) / 100
  IB = IB + (Bruto - SchaalMax(2)) * SchaalPerc(3) / 100
ElseIf Bruto > SchaalMax(3) Then
  IB = IB + SchaalMax(1) * SchaalPerc(1) / 100
  IB = IB + (SchaalMax(2) - SchaalMax(1)) * SchaalPerc(2) / 100
  IB = IB + (SchaalMax(3) - SchaalMax(2)) * SchaalPerc(3) / 100
  IB = IB + (Bruto - SchaalMax(3)) * SchaalPerc(4) / 100
End If

End Function


Public Function TIDALMINMAXFROMARRAY(myArray() As Variant, RoundHours As Boolean) As Variant()
  'Author: Siebe Bosch
  'Date: 1-9-2013
  'Description: extracts the tidal minima and maxima from a 2D-array with date/time and water levels
  Dim i As Long, j As Long, k As Long, c As Long, n As Long, timeStep As Integer, SearchRadius As Integer
  Dim curVal As Variant, IsMin As Boolean, IsMax As Boolean, Header As String, CurDate As Date
  timeStep = (myArray(3, 1) - myArray(2, 1)) * 24 * 60 'in minutes
  n = UBound(myArray, 1) * (UBound(myArray, 2) - 1)
    
  'diminsioning the arrays
  Dim TidalArray() As Variant
  ReDim TidalArray(1 To n, 1 To 4)
  Dim FinalArray() As Variant
  
  'setting the search radius
  If timeStep <= 10 Then
    SearchRadius = 30 '10 minutes timestep. Detect tidal value by comparing -5 hours and + 5 hours
  ElseIf timeStep <= 15 Then
    SearchRadius = 20 '15 minutes timestep. Detect tidal value by comparing -5 hours +5 hours
  ElseIf timeStep < 60 Then
    SearchRadius = 5 '1 hour timestep. Detect tidal value by comparing -5 and +5 hours
  End If
  
  'walk through the array and search for tides
  For c = 2 To UBound(myArray, 2)
    For i = 2 + SearchRadius To UBound(myArray) - SearchRadius   'we start at row 2 since the first row contains headers
      Header = myArray(1, c)
      CurDate = myArray(i, 1)
      curVal = myArray(i, c)
      
      IsMin = True
      IsMax = True
      
      'search backward
      For j = i - 1 To i - SearchRadius Step -1
        If myArray(j, c) >= curVal Then IsMax = False       'note: the >= here and the > in the next section is importand in case of equal values!
        If myArray(j, c) <= curVal Then IsMin = False       'note: the <= here and the < in the next section is importand in case of equal values!
        If IsMax = False And IsMin = False Then Exit For
      Next
      
      'search forward
      For j = i + 1 To i + SearchRadius Step 1
        If myArray(j, c) > curVal Then IsMax = False
        If myArray(j, c) < curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      
      'identify whether this point is a tidal min or max
      If IsMin Or IsMax Then
        k = k + 1
        TidalArray(k, 1) = Header
        If RoundHours Then
          TidalArray(k, 2) = DATETWOHOURWINDOW(myArray(i, 1))   'since the timing of the computed and observed peak may difference, we'll introduce a certain bandwidth
        Else
          TidalArray(k, 2) = myArray(i, 1)
        End If
        TidalArray(k, 3) = curVal
        If IsMin Then
          TidalArray(k, 4) = "Laag"
        ElseIf IsMax Then
          TidalArray(k, 4) = "Hoog"
        End If
      End If
    Next
  Next
  
  'truncate the tidal array to match the actual number of tides
  ReDim FinalArray(1 To k, 1 To 4)
  For i = 1 To k
    FinalArray(i, 1) = TidalArray(i, 1)
    FinalArray(i, 2) = TidalArray(i, 2)
    FinalArray(i, 3) = TidalArray(i, 3)
    FinalArray(i, 4) = TidalArray(i, 4)
  Next
  
  TIDALMINMAXFROMARRAY = FinalArray

End Function


Public Function TIDALMINMAXSEQUENCE(MyRange As Range, ResultsRow As Integer, ResultsCol As Integer) As Variant

  Dim i As Long, j As Long, k As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Variant, timeStep As Integer, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Variant, lastMaxDate As Variant, lastDate As Variant
  Dim LastMinVal As Variant, lastMaxVal As Variant
  Dim LastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Variant, FirstDone As Boolean, LocationDone As Boolean
  Dim IsMin As Boolean, IsMax As Boolean, curVal As Variant
  
  'Date: 30-8-2013
  'Author: Siebe Bosch
  'Description: subdivides a day into 4 quarters and exports the tidal min or max within that quarter
  
  r = ResultsRow
  c = ResultsCol
  
  ActiveSheet.Cells(r, c) = "Datum/Tijd"
  ActiveSheet.Cells(r, c + 1) = MyRange.Cells(1, 2) 'location name
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  timeStep = (MyRange.Cells(3, 1) - MyRange.Cells(2, 1)) * 24 * 60 'in minutes
  If timeStep <= 10 Then
    SearchRadius = 12 '10 minutes timestep. Detect tidal value by comparing -2 hours and + 2 hours
  ElseIf timeStep <= 15 Then
    SearchRadius = 8 '15 minutes timestep. Detect tidal value by comparing -2 hours +2 hours
  ElseIf timeStep < 60 Then
    SearchRadius = 2 '1 hour timestep. Detect tidal value by comparing -2 and +2 hours
  End If
    
  If MyRange.Columns.Count < 2 Then
    TIDALMINMAXSEQUENCE = "Error: data range must contain at least two columns: date/time and values"
  ElseIf MyRange.Rows.Count < 2 Then
    TIDALMINMAXSEQUENCE = "Error: data range must contain a sufficient number of rows"
  Else
    For i = 2 + SearchRadius To MyRange.Rows.Count - SearchRadius
      curVal = MyRange.Cells(i, 2)
      If MyRange.Cells(i + SearchRadius, 1) = 0 Then Exit For 'reached the end of the timeseries
      
      IsMin = True
      IsMax = True
      For j = i - 1 To i - SearchRadius Step -1
        If MyRange.Cells(j, 2) >= curVal Then IsMax = False
        If MyRange.Cells(j, 2) <= curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      For j = i + 1 To i + SearchRadius Step 1
        If MyRange.Cells(j, 2) >= curVal Then IsMax = False
        If MyRange.Cells(j, 2) <= curVal Then IsMin = False
        If IsMax = False And IsMin = False Then Exit For
      Next
      If IsMin Or IsMax Then
        r = r + 1
        ActiveSheet.Cells(r, c) = DATEHOUR(MyRange.Cells(i, 1))
        ActiveSheet.Cells(r, c + 1) = curVal
      End If
    Next
  End If
  TIDALMINMAXSEQUENCE = True
End Function

Public Function WindRichting(Angle As Variant, ReturnNumeric As Boolean) As Variant
  If (Angle = 0 Or Angle > 360) Then
    WindRichting = "Windstil/Variabel"
  ElseIf (Angle < 22.5 Or Angle >= 337.5) Then
    If ReturnNumeric Then
      WindRichting = 0
    Else
      WindRichting = "N"
    End If
  ElseIf (Angle < 67.5 And Angle >= 22.5) Then
    If ReturnNumeric Then
      WindRichting = 45
    Else
      WindRichting = "NO"
    End If
  ElseIf (Angle < 112.5 And Angle >= 67.5) Then
    If ReturnNumeric Then
      WindRichting = 90
    Else
      WindRichting = "O"
    End If
  ElseIf (Angle < 157.5 And Angle >= 112.5) Then
    If ReturnNumeric Then
      WindRichting = 135
    Else
      WindRichting = "ZO"
    End If
  ElseIf (Angle < 202.5 And Angle >= 157.5) Then
    If ReturnNumeric Then
      WindRichting = 180
    Else
      WindRichting = "Z"
    End If
  ElseIf (Angle < 247.5 And Angle >= 202.5) Then
    If ReturnNumeric Then
      WindRichting = 225
    Else
      WindRichting = "ZW"
    End If
  ElseIf (Angle < 292.5 And Angle >= 247.5) Then
    If ReturnNumeric Then
      WindRichting = 270
    Else
      WindRichting = "W"
    End If
  ElseIf (Angle < 337.5 And Angle >= 292.5) Then
    If ReturnNumeric Then
      WindRichting = 315
    Else
      WindRichting = "NW"
    End If
  Else
    WindRichting = "Windstil/Variabel"
  End If
End Function

Public Function EXTRACTHARMONICFROMRANGE(MyRange As Range, myPeriodDays As Variant, ResultsRow As Integer, ResultsCol As Integer) As Boolean
  'this function extracts a harmonic (sinusoideal function) from a range with date/time (first column) and values (second column). (E.g. tidal movement)
  'for a given period (days) of the harmonic to extract
  'it does so by minimizing the RMS between observed and computed values
  'the remaining timeseries is written to the worksheet as well as the amplitude of the harmonic found
  
  'first calculate the average value inside the range
  Dim avgVal As Variant, minVal As Variant, maxVal As Variant
  avgVal = Application.WorksheetFunction.sum(Range(MyRange.Cells(1, 2), MyRange.Cells(MyRange.Rows.Count, 2))) / MyRange.Rows.Count
  minVal = Application.WorksheetFunction.Min(Range(MyRange.Cells(1, 2), MyRange.Cells(MyRange.Rows.Count, 2)))
  maxVal = Application.WorksheetFunction.max(Range(MyRange.Cells(1, 2), MyRange.Cells(MyRange.Rows.Count, 2)))
  
  ActiveSheet.Cells(ResultsRow, ResultsCol) = "gem:"
  ActiveSheet.Cells(ResultsRow, ResultsCol) = "min:"
  ActiveSheet.Cells(ResultsRow, ResultsCol) = "max:"
  ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = avgVal
  ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = minVal
  ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = maxVal
  
  
End Function

Public Function TIDALMINMAXFROMSERIES(MyRange As Range, ResultsRow As Integer, ResultsCol As Integer) As Boolean

  Dim i As Long, j As Long, k As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Variant, timeStep As Variant, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Variant, lastMaxDate As Variant, lastDate As Variant, startDate As Variant
  Dim LastMinVal As Variant, lastMaxVal As Variant
  Dim LastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Variant, FirstDone As Boolean, LocationDone As Boolean
  
  r = ResultsRow
  c = ResultsCol
  
  ActiveSheet.Cells(r, c) = "Location"
  ActiveSheet.Cells(r, c + 1) = "Date/time low"
  ActiveSheet.Cells(r, c + 2) = "Low"
  ActiveSheet.Cells(r, c + 3) = "Date/time high"
  ActiveSheet.Cells(r, c + 4) = "High"
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  timeStep = MyRange.Cells(3, 1) - MyRange.Cells(2, 1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / timeStep, 0)
  
  If MyRange.Columns.Count < 2 Then
    TIDALMINMAXFROMSERIES = "Error: data range must contain at least two columns: date/time and values"
  ElseIf MyRange.Rows.Count < 2 Then
    TIDALMINMAXFROMSERIES = "Error: data range must contain a sufficient number of rows"
  Else
    For j = 2 To MyRange.Columns.Count
      
      Location = MyRange.Cells(1, j)
      FirstDone = False
      LocationDone = False
      
      i = 2
      startDate = MyRange.Cells(i, 1)
      LastMinVal = MyRange.Cells(i, j)
      lastMaxVal = MyRange.Cells(i, j)
      lastMinDate = MyRange.Cells(i, 1)
      lastMaxDate = MyRange.Cells(i, 1)
      lastDate = MyRange.Cells(i, 1)
      
      'first find the minimum and maximum in the first 13.1 hours which is a little longer than one tidal wave (12.5 h)
      While Not FirstDone = True
        myDate = MyRange.Cells(i, 1)
        myVal = MyRange.Cells(i, j)
        If (myDate - startDate) > (13.1 / 24) Then
          FirstDone = True
        ElseIf i > MyRange.Rows.Count Then
          FirstDone = True
        Else
          If myVal < LastMinVal Then
            LastMinVal = myVal
            lastMinDate = myDate
            LastMinIdx = i
          ElseIf myVal > lastMaxVal Then
            lastMaxVal = myVal
            lastMaxDate = myDate
            lastMaxIdx = i
          End If
        End If
        i = i + 1
      Wend
      
      'write the initial results
      r = r + 1
      ActiveSheet.Cells(r, c) = Location
      ActiveSheet.Cells(r, c + 1) = lastMinDate
      ActiveSheet.Cells(r, c + 2) = LastMinVal
      ActiveSheet.Cells(r, c + 3) = lastMaxDate
      ActiveSheet.Cells(r, c + 4) = lastMaxVal
      
      'now that we have a startlocation we can start looking for the next minima and maxima
      While Not LocationDone
        LastMinVal = 99999999
        lastMaxVal = -99999999
        
        'find the next minimum approximately 12.5 hours later
        nextIdx = Math.Round(LastMinIdx + (12.5 / 24) / timeStep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > MyRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = MyRange.Cells(i, 1)
            myVal = MyRange.Cells(i, j)
            If myVal < LastMinVal Then
              lastMinDate = myDate
              LastMinVal = myVal
              LastMinIdx = i
            End If
          End If
        Next
        
        'find the next maximum
        nextIdx = Math.Round(lastMaxIdx + (12.5 / 24) / timeStep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > MyRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = MyRange.Cells(i, 1)
            myVal = MyRange.Cells(i, j)
            If myVal > lastMaxVal Then
              lastMaxDate = myDate
              lastMaxVal = myVal
              lastMaxIdx = i
            End If
          End If
        Next
        
        'write the results
        r = r + 1
        ActiveSheet.Cells(r, c) = Location
        ActiveSheet.Cells(r, c + 1) = lastMinDate
        ActiveSheet.Cells(r, c + 2) = LastMinVal
        ActiveSheet.Cells(r, c + 3) = lastMaxDate
        ActiveSheet.Cells(r, c + 4) = lastMaxVal
      
      Wend
    Next
  End If
  TIDALMINMAXFROMSERIES = True

End Function

Public Function GetTidalWavePeriodMinutes() As Double
    'https://nl.wikipedia.org/wiki/Getijde_(waterbeweging)#:~:text=De%20periode%20van%20het%20stijgen,minimale%20hoogte%20laagwater%20of%20laagtij.
    GetTidalWavePeriodMinutes = 12 * 60 + 25 + 0.2
End Function


Public Function AddSpuiperiodeToRange(MyRange As Range, DateCol As Integer, TidalCol As Integer, WaterlevelCol As Integer, ResultsCol As Integer, MinimumLevelDifference As Double, MinimumSpuiDurationHours As Double) As Boolean
    'deze functie identificeert de spuiperiode op basis van een reeks met getijden (tidalcol) en een reeks met binnendijkse waterhoogtes
    'de spuiperiode wordt weggeschreven in de vorm ddmmyyyy_x
    'als de laagste waterhoogte na middernacht valt geldt dit als spuiperiode 1; als hij vóór middernacht valt 2
    'een spuiperiode wordt alleen als zodanig aangemerkt als de duur waarover gespuid kan worden groter of gelijk is aan MinimumSpuiDurationHours
    Dim timeStep As Variant, SearchRadius As Integer
    Dim FirstDone As Boolean
    Dim startDate As Variant, myDate As Variant, myVal As Double
    Dim i As Long
    Dim Done As Boolean
    Dim mySpuiperiode As String
    Dim LastMinVal As Double
    Dim lastMinDate As Variant
    Dim LastMinIdx As Integer, nextIdx As Integer
    Dim mySpuiDate As String
    Dim Found As Boolean
    Dim StartIdx As Integer
    Dim EndIdx As Integer
    Dim Spuibaar As Boolean
    Dim SpuiPeriodHours As Double


    'first analyse the timestep in order to specify a window size.
    'then derive a search radius (expressed in n_timesteps) for each following min/max
    timeStep = MyRange.Cells(3, DateCol) - MyRange.Cells(2, DateCol)
    SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / timeStep, 0)
    SearchRadius = Application.WorksheetFunction.max(1, SearchRadius)       'always give our routine at least ONE timestep to check for the lowest surrounding waterlevel
    LastMinVal = 9E+99


    'first find the minimum in the first 13.1 hours which is a little longer than one tidal wave (12.5 h)
    i = 2
    FirstDone = False
    startDate = MyRange.Cells(i, DateCol)
    While Not FirstDone = True
        myDate = MyRange.Cells(i, DateCol)
        myVal = MyRange.Cells(i, TidalCol)
        If (myDate - startDate) > (13.1 / 24) Then
          FirstDone = True
        ElseIf i > MyRange.Rows.Count Then
          FirstDone = True
        Else
          If myVal < LastMinVal Then
            LastMinVal = myVal
            lastMinDate = myDate
            LastMinIdx = i
          End If
        End If
        i = i + 1
    Wend
    
    'now assign our spuiperiod identifier to each timestep where the waterlevel < waterlevel upstream
    mySpuiDate = Format(lastMinDate, "ddMMyyyy")
    If hour(lastMinDate) >= 12 Then mySpuiperiode = mySpuiDate & "_2" Else mySpuiperiode = mySpuiDate & "_1"
    
    Found = False
    i = LastMinIdx
    StartIdx = -1
    While Not Found
        If i < 2 Then
            Found = True
        Else
            i = i - 1
            If MyRange.Cells(i, TidalCol) < MyRange.Cells(i, WaterlevelCol) - MinimumLevelDifference Then
                StartIdx = i
            Else
                Found = True
            End If
        End If
    Wend
    
    Found = False
    i = LastMinIdx
    EndIdx = -1
    While Not Found
        i = i + 1
        If MyRange.Cells(i, TidalCol) < MyRange.Cells(i, WaterlevelCol) - MinimumLevelDifference Then
            EndIdx = i
        Else
            Found = True
        End If
    Wend
        
    '------------------------------------------------------------------------------------------
    'write information about the 'spuibaarheid' for this spuiperiod to the extra column
    '------------------------------------------------------------------------------------------
    If EndIdx > StartIdx And StartIdx > 0 Then
        SpuiPeriodHours = Application.WorksheetFunction.Round((MyRange.Cells(EndIdx, DateCol) - MyRange.Cells(StartIdx, DateCol)) * 24, 2)
        If SpuiPeriodHours >= MinimumSpuiDurationHours Then
            For i = StartIdx To EndIdx
                MyRange.Cells(i, ResultsCol) = mySpuiperiode
            Next
        End If
    End If
    '------------------------------------------------------------------------------------------
        
    
    'now that we have identified our first 'spuiperiode' proceed with the next ones
    Done = False
    While Not Done
        'find the next minimum approximately 12.5 hours later
        nextIdx = Math.Round(LastMinIdx + (12.5 / 24) / timeStep, 0)
        
        'If nextIdx > 657 Then Stop
        
        If MyRange.Cells(nextIdx, TidalCol) = "" Then
            'no next tide found so quit
            Done = True
        Else
            LastMinIdx = nextIdx
            LastMinVal = MyRange.Cells(LastMinIdx, TidalCol)
            lastMinDate = MyRange.Cells(LastMinIdx, DateCol)
            For i = nextIdx - SearchRadius To nextIdx + SearchRadius
                If MyRange.Cells(i, TidalCol) < LastMinVal Then
                    LastMinIdx = i
                    LastMinVal = MyRange.Cells(LastMinIdx, TidalCol)
                    lastMinDate = MyRange.Cells(LastMinIdx, DateCol)
                End If
            Next
            
            
            'now assign our spuiperiod identifier to each timestep where the waterlevel < waterlevel upstream
            mySpuiDate = Format(lastMinDate, "ddMMyyyy")
            If hour(lastMinDate) >= 12 Then mySpuiperiode = mySpuiDate & "_2" Else mySpuiperiode = mySpuiDate & "_1"
            
            'If mySpuiperiode = "04012012_1" Then Stop
            
            'now assign the spuiperiode
            Found = False
            i = LastMinIdx
            StartIdx = -1
            While Not Found
                If i < 2 Then
                    Found = True
                Else
                    i = i - 1
                    If MyRange.Cells(i, TidalCol) < MyRange.Cells(i, WaterlevelCol) - MinimumLevelDifference Then
                        'myRange.Cells(i, ResultsCol) = mySpuiperiode
                        StartIdx = i
                    Else
                        Found = True
                    End If
                End If
            Wend
    
            Found = False
            i = LastMinIdx
            EndIdx = -1
            While Not Found
                i = i + 1
                If MyRange.Cells(i, TidalCol) <> "" And MyRange.Cells(i, TidalCol) < MyRange.Cells(i, WaterlevelCol) - MinimumLevelDifference Then
                    'myRange.Cells(i, ResultsCol) = mySpuiperiode
                    EndIdx = i
                Else
                    Found = True
                End If
            Wend
            
            
            '------------------------------------------------------------------------------------------
            'write information about the 'spuibaarheid' for this spuiperiod to the extra column
            '------------------------------------------------------------------------------------------
            If EndIdx > StartIdx And StartIdx > 0 Then
                SpuiPeriodHours = Application.WorksheetFunction.Round((MyRange.Cells(EndIdx, DateCol) - MyRange.Cells(StartIdx, DateCol)) * 24, 2)
                If SpuiPeriodHours >= MinimumSpuiDurationHours Then
                    For i = StartIdx To EndIdx
                        MyRange.Cells(i, ResultsCol) = mySpuiperiode
                    Next
                End If
            End If
            '------------------------------------------------------------------------------------------

            
        End If
        
    Wend
    
    
    

    
End Function

Public Function TIDALLOWSFROMSERIES(MyRange As Range, ResultsRow As Integer, ResultsCol As Integer) As Boolean

  Dim i As Long, j As Long, k As Long, n As Long, r As Long, c As Long
  Dim Location As String, myVal As Variant, timeStep As Variant, SearchRadius As Integer
  Dim ValRange As Range, DateRange As Range
  Dim lastMinDate As Variant, lastMaxDate As Variant, lastDate As Variant, startDate As Variant
  Dim LastMinVal As Variant, lastMaxVal As Variant
  Dim LastMinIdx As Long, lastMaxIdx As Long, nextIdx As Long
  Dim myDate As Variant, FirstDone As Boolean, LocationDone As Boolean
  
  r = ResultsRow
  c = ResultsCol
  
  ActiveSheet.Cells(r, c) = "Location"
  ActiveSheet.Cells(r, c + 1) = "Date/time"
  ActiveSheet.Cells(r, c + 2) = "Low tide"
  
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  timeStep = MyRange.Cells(3, 1) - MyRange.Cells(2, 1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / timeStep, 0)
  
  If MyRange.Columns.Count < 2 Then
    TIDALLOWSFROMSERIES = "Error: data range must contain at least two columns: date/time and values"
  ElseIf MyRange.Rows.Count < 2 Then
    TIDALLOWSFROMSERIES = "Error: data range must contain a sufficient number of rows"
  Else
    For j = 2 To MyRange.Columns.Count
      
      Location = MyRange.Cells(1, j)
      FirstDone = False
      LocationDone = False
      
      i = 2
      startDate = MyRange.Cells(i, 1)
      LastMinVal = MyRange.Cells(i, j)
      lastMinDate = MyRange.Cells(i, 1)
      lastDate = MyRange.Cells(i, 1)
      
      'first find the minimum in the first 13.1 hours which is a little longer than one tidal wave (12.5 h)
      While Not FirstDone = True
        myDate = MyRange.Cells(i, 1)
        myVal = MyRange.Cells(i, j)
        If (myDate - startDate) > (13.1 / 24) Then
          FirstDone = True
        ElseIf i > MyRange.Rows.Count Then
          FirstDone = True
        Else
          If myVal < LastMinVal Then
            LastMinVal = myVal
            lastMinDate = myDate
            LastMinIdx = i
          End If
        End If
        i = i + 1
      Wend
      
      'write the initial results
      r = r + 1
      ActiveSheet.Cells(r, c) = Location
      ActiveSheet.Cells(r, c + 1) = lastMinDate
      ActiveSheet.Cells(r, c + 2) = LastMinVal
      
      'now that we have a startlocation we can start looking for the next minima
      While Not LocationDone
        LastMinVal = 99999999
        
        'find the next minimum approximately 12.5 hours later
        nextIdx = Math.Round(LastMinIdx + (12.5 / 24) / timeStep, 0)
        For i = nextIdx - SearchRadius To nextIdx + SearchRadius
          If i > MyRange.Rows.Count Then
            LocationDone = True
          Else
            myDate = MyRange.Cells(i, 1)
            myVal = MyRange.Cells(i, j)
            If myVal < LastMinVal Then
              lastMinDate = myDate
              LastMinVal = myVal
              LastMinIdx = i
            End If
          End If
        Next
        
        'write the results
        If Not LocationDone Then
          r = r + 1
          ActiveSheet.Cells(r, c) = Location
          ActiveSheet.Cells(r, c + 1) = lastMinDate
          ActiveSheet.Cells(r, c + 2) = LastMinVal
        End If
      
      Wend
    Next
  End If
  TIDALLOWSFROMSERIES = True

End Function

Public Function TIDALLOWSFROMARRAYS(dates() As Date, Vals() As Single, ByRef DatesLow() As Date, ByRef ValsLow() As Single) As Boolean

  Dim i As Long, j As Long
  Dim myVal As Variant, timeStep As Variant, SearchRadius As Integer
  Dim lastMinDate As Variant, lastDate As Variant, startDate As Variant
  Dim LastMinVal As Variant
  Dim LastMinIdx As Long, nextIdx As Long
  Dim myDate As Variant, FirstDone As Boolean, Done As Boolean
  
  'initialize the size of the output array to be the same as the input array. We'll redim again later
  ReDim DatesLow(1 To UBound(dates))
  ReDim ValsLow(1 To UBound(Vals))
    
  'first analyse the timestep in order to specify a window size.
  'then derive a search radius (expressed in n_timesteps) for each following min/max
  timeStep = dates(2) - dates(1)
  SearchRadius = Application.WorksheetFunction.RoundUp((2 / 24) / timeStep, 0)
  
  FirstDone = False
      
  startDate = dates(1)
  LastMinVal = Vals(1)
  lastMinDate = dates(1)
  lastDate = dates(1)
  i = 1
      
  'first find the minimum in the first 13.1 hours which is a little longer than one tidal wave (12.5 hours)
  While Not FirstDone = True
    myDate = dates(i)
    myVal = Vals(i)
    
    If (myDate - startDate) > (13.1 / 24) Then
      FirstDone = True
    ElseIf i > UBound(dates) Then
      FirstDone = True
    Else
      If myVal < LastMinVal Then
        LastMinVal = myVal
        lastMinDate = myDate
        LastMinIdx = i
      End If
    End If
    i = i + 1
  Wend
      
  'write the initial results
  j = j + 1
  DatesLow(j) = lastMinDate
  ValsLow(j) = LastMinVal
      
  'now that we have a startlocation we can start looking for the next minima
  While Not Done
      
    'initialize the minimum value
    LastMinVal = 99999999
        
    'find the next low tide approximately 12.5 hours later
    nextIdx = Math.Round(LastMinIdx + (12.5 / 24) / timeStep, 0)
    For i = nextIdx - SearchRadius To nextIdx + SearchRadius
      If i > UBound(dates) Then
        Done = True
      Else
        myDate = dates(i)
        myVal = Vals(i)
        
        If myVal < LastMinVal Then
          lastMinDate = myDate
          LastMinVal = myVal
          LastMinIdx = i
        End If
      End If
    Next
        
    'write the results
    If Not Done Then
      j = j + 1
      DatesLow(j) = lastMinDate
      ValsLow(j) = LastMinVal
    End If
  Wend
  
  'define the upper boundary of the output arrays
  ReDim Preserve DatesLow(1 To j)
  ReDim Preserve ValsLow(1 To j)
  
  TIDALLOWSFROMARRAYS = True

End Function

Public Function getAvgMaxFromTide(MyRange As Range) As Variant

Dim Cumulative As Variant, myVal As Variant
Dim i As Long, n As Long

If MyRange.Columns.Count = 1 Then
  For i = 3 To MyRange.Rows.Count - 2
    If IsNumeric(MyRange.Cells(i, 1)) Then
      myVal = MyRange.Cells(i, 1)
      If MyRange.Cells(i - 1, 1) < myVal And myVal >= MyRange.Cells(i + 1, 1) And MyRange.Cells(i - 2, 1) < myVal And myVal >= MyRange.Cells(i + 2, 1) Then
        n = n + 1
        Cumulative = Cumulative + myVal
      End If
    End If
  Next i
End If

getAvgMaxFromTide = Cumulative / n

End Function

Public Function getAvgMinFromTide(MyRange As Range) As Variant

Dim Cumulative As Variant, myVal As Variant
Dim i As Long, n As Long

If MyRange.Columns.Count = 1 Then
  For i = 3 To MyRange.Rows.Count - 2
    If IsNumeric(MyRange.Cells(i, 1)) Then
      myVal = MyRange.Cells(i, 1)
      If MyRange.Cells(i - 1, 1) > myVal And myVal <= MyRange.Cells(i + 1, 1) And MyRange.Cells(i - 2, 1) > myVal And myVal <= MyRange.Cells(i + 2, 1) Then
        n = n + 1
        Cumulative = Cumulative + myVal
      End If
    End If
  Next i
End If

getAvgMinFromTide = Cumulative / n

End Function

Public Sub READHMCZDATA(path As String, TargetSheet As String, startRow As Long, StartCol As Long, IntervalMinutes As Long)
  Dim fn As Long, fileContent As String, FileRecords() As String
  Dim i As Long, r As Long
  Dim spc1 As Long, spc2 As Long, spc3 As Long
  Dim datestr As String, TimeStr As String, valstr As String
  Dim Uur As Long, Minuut As Long
  Dim Tijd As Variant
  
  r = 0
  fn = FreeFile
  
  Open path For Input As #fn
  fileContent = Input(VBA.LOF(fn), #fn)
  FileRecords = Split(fileContent, vbLf)
  Close (fn)
  
  If WorkSheetExists(TargetSheet) Then
    For i = 0 To UBound(FileRecords())
      Dim tmpStr As String
      tmpStr = FileRecords(i)
      
      If VBA.Mid(tmpStr, 3, 1) = "-" Then
        datestr = ParseString(tmpStr, " ")
        TimeStr = ParseString(tmpStr, " ")
        valstr = ParseString(tmpStr, " ")
        Uur = VAL(ParseString(TimeStr, ":"))
        Minuut = VAL(ParseString(TimeStr, ":"))
        
        Dim DateConvert As Date
        DateConvert = VBA.Format(datestr, "dd-mm-yyyy")
        Tijd = TimeSerial(Uur, Minuut, 0)
          
        If IsNumeric(valstr) And Minuut / IntervalMinutes = VBA.Round(Minuut / IntervalMinutes, 0) Then
          r = r + 1
          Worksheets(TargetSheet).Cells(startRow + r, StartCol) = DateConvert + Tijd
          Worksheets(TargetSheet).Cells(startRow + r, StartCol + 1) = valstr / 100
        End If
      End If
      
    Next
  Else
    MsgBox ("Target worksheet does not exist")
  End If

  
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK SOBEK
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ReadSpecificHISResults(HisFile As String, myLoc As String, myPar As String, AllowLeftParMatch As Boolean, AllowLocWildCards As Boolean, ValueSelection As String, Multiplier As Variant) As Collection

' this routine is an example of how you can use ODSSVR20.DLL to read a his file
' give the variable myHis the same structure as the class clsODSServer from the ODSSVR20.DLL library
Dim myHis As ODSSVR20.clsODSServer
Set myHis = New ODSSVR20.clsODSServer

Set ReadSpecificHISResults = New Collection
Dim myResult As clsDateValPair

Dim values() As Single
Dim Loc() As String, par() As String, Tim() As Variant
Dim LocDef() As String, ParDef() As String, TimDef() As Variant
Dim nLoc As Long, nPar As Long, nTim As Long
Dim myVal As Variant

' iRes is just a number that will be returned by the ODSSVR20 library. 0 means: the function was successfully called
' iLoc, iPar and iTim are paramters that will later run from 1 to the number of resp. Locations, Parameters and Timesteps
Dim iRes As Long, iLoc As Long, iPar As Long, iTim As Long
  
myHis.KeepFilesOpen = True
  
myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
  MsgBox "HisFile does not exist: " & HisFile
  Exit Function
End If
  
'hereunder we give you two options:
'1 - to read the whole file
'2 - to first read all locations, parameters & dates/times, and then to choose for which location and parameter to retrieve data
'We've commented out the first option, but feel free to retrieve that one by removing the "'" characters
  
'----------option1-----------------------------------------------
'read the whole file
'The argument "Hisfile" goes into the function, the rest of the parameters are returned to you by the function
'if the hisfile has properly been read, the function returns a value of 0 for itself
'iRes = myHis.GetAllData(Values(), nLoc, nPar, nTim, Hisfile, , Loc(), Par(), Tim())
'If iRes <> 0 Then MsgBox "Function call GetAllData not successful."
'----------/option1-----------------------------------------------
  
'--------option 2------------------------------------------------
'first read the locations
iRes = myHis.GetLoc(HisFile, , nLoc, , Loc())
If iRes <> 0 Then MsgBox "Function call GetLoc not successful."
  
' read the parameters
iRes = myHis.GetPar(HisFile, , nPar, , par())
If iRes <> 0 Then MsgBox "Function call GetPar not successful."
  
' read the dates/times
iRes = myHis.GetTime(HisFile, , nTim, , , Tim())
If iRes <> 0 Then MsgBox "Function call GetTime not successful."

' only read the values for Loc(1) en Par(1). Because of this we'lpl redimension these variables such that they can only contain one value
ReDim LocDef(1 To 1), ParDef(1 To 1) As String, TimDef(1 To nTim) As Variant

'Walk through all parameters and check if the current parameter matches the requested parameter
For iPar = LBound(par()) To UBound(par())
  If (VBA.LCase(Left(par(iPar), VBA.Len(myPar))) = VBA.LCase(myPar) And AllowLeftParMatch = True) Or (LCase(par(iPar)) = LCase(myPar)) Then

    For iLoc = LBound(Loc()) To UBound(Loc())
  
      'Check if the current location matches the requested location
      If (LCase(Loc(iLoc)) = LCase(myLoc)) Or (AllowLocWildCards = True And MATCHWILDCARD(Loc(iLoc), myLoc, False) = True) Then
      
        LocDef(1) = Loc(iLoc)
        ParDef(1) = par(iPar)
        TimDef = Tim
  
        'iRes = myHis.GetData(Values(), nLoc, nPar, nTim, HisFile, strlocdef:=LocDef(), strpardef:=ParDef(), dblTimdef:=TimDef(), strLocLst:=Loc(), strparlst:=Par(), dbltimlst:=Tim())
        iRes = myHis.GetData(values(), nLoc, nPar, nTim, HisFile, strlocdef:=LocDef(), strpardef:=ParDef(), dblTimdef:=TimDef())
        If iRes <> 0 Then MsgBox "Function call GetData not successful."
        
        For iTim = LBound(TimDef()) To UBound(TimDef())
          myVal = values(1, 1, iTim)
          
          'de selectie toepassen
          If VBA.LCase(ValueSelection) = "< 0" Then
            myVal = Minimum(myVal, 0)
          ElseIf VBA.LCase(ValueSelection) = "> 0" Then
            myVal = Maximum(myVal, 0)
          ElseIf VBA.LCase(ValueSelection) = "absolute" Then
            myVal = Math.Abs(myVal)
          ElseIf ValueSelection = "" Then
            'do nothing
          Else
            MsgBox ("Error: value selection was not recognized " & ValueSelection)
            End
          End If
        
          If ReadSpecificHISResults.Count >= iTim Then
            'add to existing value
            Set myResult = ReadSpecificHISResults.Item(iTim)
            myResult.value = myResult.value + myVal * Multiplier
          Else
            'create a new datavalue pair and add it to the collection
            Set myResult = New clsDateValPair
            myResult.Datum = Tim(iTim)
            myResult.value = myVal * Multiplier
            Call ReadSpecificHISResults.Add(myResult)
          End If
        Next
      End If
    Next
    Exit For
  End If
Next

myHis.CloseFiles
myHis.Delete HisFile
Set myHis = Nothing
Erase Loc, par, Tim, LocDef, ParDef, TimDef, values
  
End Function

Public Function ReadHISLocParTim(HisFile As String, ByRef Loc() As String, ByRef par() As String, ByRef Tim() As Variant) As Boolean

' this routine is an example of how you can use ODSSVR20.DLL to read a his file
' give the variable myHis the same structure as the class clsODSServer from the ODSSVR20.DLL library
Dim myHis As ODSSVR20.clsODSServer
Set myHis = New ODSSVR20.clsODSServer

Set ReadSpecificHISResults = New Collection
Dim myResult As clsDateValPair
  
myHis.KeepFilesOpen = True
  
myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
  MsgBox "HisFile does not exist: " & HisFile
  ReadHISLocParTim = False
End If
    
'first read the locations
iRes = myHis.GetLoc(HisFile, , nLoc, , Loc())
If iRes <> 0 Then MsgBox "Function call GetLoc not successful."
  
' read the parameters
iRes = myHis.GetPar(HisFile, , nPar, , par())
If iRes <> 0 Then MsgBox "Function call GetPar not successful."
  
' read the dates/times
iRes = myHis.GetTime(HisFile, , nTim, , , Tim())
If iRes <> 0 Then MsgBox "Function call GetTime not successful."

myHis.CloseFiles
myHis.Delete HisFile
Set myHis = Nothing

ReadHISLocParTim = True
  
End Function

Public Function getNodeStatsFromSobekCase(SbkCaseDir As String, ParIdx As Long) As Collection
  'geeft voor iedere locatie in een hisfile de laatste (in tijd) waarde terug
  Dim HisFile As String, tpFile As String
  HisFile = VBA.Replace(SbkCaseDir & "\calcpnt.his", "\\", "\")
  tpFile = VBA.Replace(SbkCaseDir & "\network.tp", "\\", "\")
  
  Dim myHis As ODSSVR20.clsODSServer
  Dim TpFileContent As clsNetworkTPFileContent
  Set myHis = New ODSSVR20.clsODSServer
  Dim results As Collection
  Set results = New Collection
  
  Set TpFileContent = New clsNetworkTPFileContent
  Call TpFileContent.Read(tpFile)
  
  Dim values() As Single
  Dim Loc() As String, par() As String, Tim() As Variant
  Dim LocDef() As String, ParDef() As String, TimDef() As Variant
  Dim nLoc As Long, nPar As Long, nTim As Long

  ' iRes is just a number that will be returned by the ODSSVR20 library. 0 means: the function was successfully called
  ' iLoc, iPar and iTim are paramters that will later run from 1 to the number of resp. Locations, Parameters and Timesteps
  Dim iRes As Long, iLoc As Long, iPar As Long, iTim As Long
  Dim myLoc As clsSBKNodeStats
  Dim Min As Variant, max As Variant, avg As Variant, mySum As Variant
  
  myHis.KeepFilesOpen = True
  
  myHis.Add HisFile, HisFile, True, True                '.Add is a method that the ODSSVR library supports.
  If Not myHis.Item(HisFile).Exists Then                'give an error message if hisfile does not exist
    MsgBox "HisFile does not exist: " & HisFile
    Exit Function
  End If
    
  'read the whole file
  'The argument "Hisfile" goes into the function, the rest of the parameters are returned to you by the function
  'if the hisfile has properly been read, the function returns a value of 0 for itself
  iRes = myHis.GetAllData(values(), nLoc, nPar, nTim, HisFile, , Loc(), par(), Tim())
  If iRes <> 0 Then
    MsgBox "Function call GetAllData not successful."
  Else
    For iLoc = 1 To UBound(Loc())
      max = -99999999999#
      Min = 99999999999#
      mySum = 0
      Set myLoc = New clsSBKNodeStats
      myLoc.id = Loc(iLoc)
      myLoc.par = par(ParIdx)
      myLoc.first = values(iLoc, ParIdx, 1)
      myLoc.Last = values(iLoc, ParIdx, UBound(Tim))
      For iTim = 1 To UBound(Tim())
        mySum = mySum + values(iLoc, ParIdx, iTim)
        If values(iLoc, ParIdx, iTim) < Min Then Min = values(iLoc, ParIdx, iTim)
        If values(iLoc, ParIdx, iTim) > max Then max = values(iLoc, ParIdx, iTim)
      Next
      If UBound(Tim()) > 0 Then myLoc.avg = mySum / UBound(Tim())
      myLoc.Min = Min
      myLoc.max = max
      
      'zoek nu in de Network.TP file content de X- en Y-coördinaat op
      Dim myNode As clsCFReachNode
      Set myNode = TpFileContent.FindNode(myLoc.id)
      If Not myNode Is Nothing Then
        myLoc.x = myNode.x
        myLoc.y = myNode.y
      End If
      Call results.Add(myLoc)
    Next
  End If
  
  Set getNodeStatsFromSobekCase = results
    
  myHis.CloseFiles
  myHis.Delete HisFile
  Set myHis = Nothing
  Erase Loc, par, Tim, LocDef, ParDef, TimDef, values
End Function

Public Function MergeStorageTables(Table1 As Collection, Table2 As Collection) As Collection
  'voegt twee hoogte/oppervlaktabellen samen
  'beide tabellen moeten een collection zijn van clsLevelAreaPair
  Dim iTable1 As Long
  Dim iTable2 As Long
  Dim Table1Done As Boolean, Table2Done As Boolean
  iTable1 = 0
  iTable2 = 0
  Dim test1 As clsLevelAreaPair
  Dim test2 As clsLevelAreaPair
  Dim newPair As clsLevelAreaPair
  
  Dim newTable As Collection
  Set newTable = New Collection
  
  If Table1.Count = 0 Then
    Set MergeStorageTables = Table2
    Exit Function
  ElseIf Table2.Count = 0 Then
    Set MergeStorageTables = Table1
    Exit Function
  End If
  
  'zet eerst een lijst met alle levels uit tabellen 1 en 2 op
  While Not (Table1Done And Table2Done)
    If Not Table1Done Then Set test1 = Table1(iTable1 + 1)
    If Not Table2Done Then Set test2 = Table2(iTable2 + 1)
    If Table1Done Then 'tabel 1 is al helemaal doorlopen; maak tabel2 in z'n eentje verder af
      iTable2 = iTable2 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table2(iTable2).Level
      Call newTable.Add(newPair)
      If iTable2 = Table2.Count Then Table2Done = True
    ElseIf Table2Done Then 'tabel 2 is al helemaal doorlopen; maak tabel1 in z'n eentje verder af
      iTable1 = iTable1 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table1(iTable1).Level
      Call newTable.Add(newPair)
      If iTable1 = Table1.Count Then Table1Done = True
    ElseIf test1.Level <= test2.Level Then
      iTable1 = iTable1 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table1(iTable1).Level
      Call newTable.Add(newPair)
      If iTable1 = Table1.Count Then Table1Done = True
    Else
      iTable2 = iTable2 + 1
      Set newPair = New clsLevelAreaPair
      newPair.Level = Table2(iTable2).Level
      Call newTable.Add(newPair)
      If iTable2 = Table2.Count Then Table2Done = True
    End If
  Wend
  
  'bereken nu VBA.Middels interpolatie voor elk van de levels het bijbehorende oppervlak
  For Each newPair In newTable
    newPair.Area = InterpolateFromStorageTable(newPair.Level, Table1) + InterpolateFromStorageTable(newPair.Level, Table2)
  Next
  
  Set MergeStorageTables = newTable
  Exit Function
End Function

Public Function InterpolateFromStorageTable(myLevel As Variant, myTable As Collection) As Variant
  'deze functie interpoleert een level binnen een level/area table
  'geeft bijbehorend oppervlak terug
  'het is een specifieke functie omdat je aan de onderkant niet extrapoleert en aan de bovenkant het oppervlak constant houdt
  Dim myPair As clsLevelAreaPair
  Dim minVal As Variant, maxVal As Variant
  Dim lPair As clsLevelAreaPair, uPair As clsLevelAreaPair
  Dim i As Long
  
  Set myPair = myTable(1)
  minVal = myPair.Level
  Set myPair = myTable(myTable.Count)
  maxVal = myPair.Level
  
  If myLevel < minVal Then
    InterpolateFromStorageTable = 0 ' voor alle waarden onder de tabel: geef nul terug
    Exit Function
  ElseIf myLevel >= maxVal Then
    Set myPair = myTable(myTable.Count)
    InterpolateFromStorageTable = myPair.Area ' voor alle waarden boven de tabel: geef maximum waarde terug
    Exit Function
  Else
    'voor alle waarden binnen de tabel: interpoleren
    For i = 1 To myTable.Count - 1
      Set lPair = myTable(i)
      Set uPair = myTable(i + 1)
      If myLevel >= lPair.Level And myLevel < uPair.Level Then
        InterpolateFromStorageTable = Interpolate(lPair.Level, lPair.Area, uPair.Level, uPair.Area, myLevel)
        Exit Function
      End If
    Next
  End If
End Function


Public Function ParseSobekRecords(myPath As String, myToken As String) As Collection
  Dim fn As Long, myStr As String
  Dim fileContent As String, records As Collection
  Set records = New Collection
  
  fileContent = ReadEntireTextFile(myPath)
  records = Split(fileContent, myToken & " ", , vbBinaryCompare)
  
End Function

Public Sub ParseSobekFile(myPath As String, ResultsRow As Long, ResultsCol As Long)
  'leest de inhoud van de Sobek (bijv. Network.CR) file in en schrijft die naar een opgegeven locatie
  Dim fn As Long, myStr As String
  Dim r As Long, c As Long
  fn = FreeFile
  Open myPath For Input As #fn
  
  r = ResultsRow - 1
  While Not EOF(fn)
    Line Input #fn, myStr
    r = r + 1
    c = ResultsCol - 1
    While Not myStr = ""
      c = c + 1
      ActiveSheet.Cells(r, c) = ParseString(myStr, " ")
    Wend
  Wend

  Close (fn)
End Sub

Public Function ParseSobekTable(ByRef myRecord As String) As Variant()
  
  'zoek allereerst naar "TBLE"
  Dim start As Boolean, endsign As Boolean, Done As Boolean
  Dim tmpRecord As String, tmpStr As String
  Dim r, c, nRow, nCol As Long
  myRecord = VBA.Replace(myRecord, vbCrLf, " ")
  myRecord = VBA.Replace(myRecord, "  ", " ")
  tmpRecord = myRecord
  
  Dim myTable() As Variant
  
  c = 0
  r = 1 'ga ervan uit dat de tabel ten minste een rij bezit
  
  'eerst gaan we de dimensies van de tabel vaststellen
  nRow = 0
  nCol = 0
  Done = False
  While Not Done
    tmpStr = ParseString(tmpRecord, " ")
    If tmpStr = "TBLE" Then start = True                'begintoken voor tabel gevonden
    If tmpStr = "<" Then
      endsign = True                 'afsluitend teken voor tabelrij gevonden
      nRow = nRow + 1                'een rij gevonden, dus meteen het tellertje bijhouden
    End If
    If endsign = False And IsNumeric(tmpStr) Then nCol = nCol + 1
    If tmpRecord = "" Or tmpStr = "tble" Then Done = True 'tabel is compleet
  Wend
  
  'nu gaan we de tabel vullen
  ReDim myTable(1 To nRow, 1 To nCol)
  r = 1
  Done = False
  While Not Done
    tmpStr = ParseString(myRecord, " ")
    If tmpStr = "TBLE" Then start = True
    If tmpStr = "<" Then
      r = r + 1
      c = 0
    ElseIf IsNumeric(tmpStr) Then
      c = c + 1
      myTable(r, c) = VAL(tmpStr)
    End If
    If myRecord = "" Or tmpStr = "tble" Then Done = True 'tabel is compleet
  Wend
  ParseSobekTable = myTable

End Function


Public Function ParseBySingleChar(ByRef myString As String) As String
  If VBA.Len(myString) > 0 Then
    ParseBySingleChar = VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  Else
    ParseBySingleChar = ""
  End If
End Function


Public Sub MakeSobekTargetLevelTable(id As String, ZP As Variant, WP As Variant, StartYear As Long, EndYear As Long, ResultsRow As Long, ResultsCol As Long)
  Dim i As Long, r As Long, c As Long
  Dim myDate As Variant
  
  r = ResultsRow
  c = ResultsCol
    
  ActiveSheet.Cells(r, c) = "ID"
  ActiveSheet.Cells(r, c + 1) = "datum"
  ActiveSheet.Cells(r, c + 2) = "tijd"
  ActiveSheet.Cells(r, c + 3) = "waarde"
  For i = StartYear To EndYear
    
    If i = StartYear Then
      r = r + 1
      ActiveSheet.Cells(r, c) = id
      ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 1, 1)
      ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
      ActiveSheet.Cells(r, c + 3) = WP
    End If
    
    r = r + 1
    ActiveSheet.Cells(r, c) = id
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 3, 31)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(23, 59, 0)
    ActiveSheet.Cells(r, c + 3) = WP
    r = r + 1
    ActiveSheet.Cells(r, c) = id
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 4, 1)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
    ActiveSheet.Cells(r, c + 3) = ZP
    r = r + 1
    ActiveSheet.Cells(r, c) = id
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 9, 30)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(23, 59, 0)
    ActiveSheet.Cells(r, c + 3) = ZP
    r = r + 1
    ActiveSheet.Cells(r, c) = id
    ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i, 10, 1)
    ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
    ActiveSheet.Cells(r, c + 3) = WP
    
    If i = EndYear Then
      r = r + 1
      ActiveSheet.Cells(r, c) = id
      ActiveSheet.Cells(r, c + 1) = VBA.DateSerial(i + 1, 1, 1)
      ActiveSheet.Cells(r, c + 2) = VBA.TimeSerial(0, 0, 0)
      ActiveSheet.Cells(r, c + 3) = WP
    End If
    
  Next

End Sub

Public Sub READBUIFILE(myPath As String, ResultsRow As Long, ResultsCol As Long)
  Dim fn As Long, myStr As String
  fn = FreeFile
  Dim RainfallData(1, 1) As Variant
  Dim nStations As Long, nEvents As Long, timeStep As Long
  Dim HeaderRead As Long
  Dim myDate As Variant, myYear As Long, myMonth As Long, myDay As Long, myHour As Long, myMinute As Long, mySecond As Long
  Dim r As Long, c As Long
  Dim i As Long
      
  r = ResultsRow
  c = ResultsCol

  'reads a .bui file (SOBEK rainfall event) and writes it to the worksheet
  Open myPath For Input As #fn
  While Not EOF(fn)
    Line Input #fn, myStr
    If VBA.Trim(VBA.LCase(myStr)) = "*aantal stations" Then
      Line Input #fn, myStr
      nStations = VAL(myStr)
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Trim(VBA.LCase(myStr)) = "*namen van stations" Then
      For i = 1 To nStations
        Line Input #fn, myStr
        ActiveSheet.Cells(r, c + i) = VBA.Replace(myStr, "'", "")
      Next
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Trim(VBA.LCase(myStr)) = "*en het aantal seconden per waarnemingstijdstap" Then
      Line Input #fn, myStr
      nEvents = VBA.VAL(ParseString(myStr, " "))
      timeStep = VBA.VAL(ParseString(myStr, " ")) / 3600 'convert to hours
      HeaderRead = HeaderRead + 1
    ElseIf VBA.Left(myStr, 1) = "*" Then
      'commentaarregel
    ElseIf HeaderRead >= 3 Then  'geen commentaarregels meer
      If HeaderRead = 3 Then
        myYear = VBA.Left(myStr, 4)
        myMonth = VBA.Mid(myStr, 5, 2)
        myDay = VBA.Mid(myStr, 7, 2)
        myHour = VBA.Mid(myStr, 9, 2)
        myMinute = VBA.Mid(myStr, 11, 2)
        mySecond = VBA.Mid(myStr, 13, 2)
        myDate = DateSerial(myYear, myMonth, myDay) + TimeSerial(myHour, myMinute, mySecond)
        'set it back one timestep before reading the first line
        myDate = myDate - timeStep / 24
        HeaderRead = HeaderRead + 1
      Else
        'nieuw record
        r = r + 1
        i = 0
        myDate = myDate + timeStep / 24
        ActiveSheet.Cells(r, c) = myDate
        While Not myStr = ""
          i = i + 1
          ActiveSheet.Cells(r, c + i) = VBA.VAL(ParseString(myStr, " "))
        Wend
      End If
    End If
  Wend
  Close (fn)
End Sub

Public Sub WriteBuiFile(path As String, DataBlock As Range, TSSecs As Integer, ProgressRange As Range)

Dim startDate As Variant, EndDate As Variant
Dim DurDays As Long, DurHours As Long, DurMins As Long, DurSecs As Long, rest As Variant, TotDur As Variant
Dim fn As Long
Dim r As Long, c As Long, i As Long
Dim stations As Collection
Set stations = New Collection
Dim myDate As Variant, myStr As String
Dim TSMins As Long

TSMins = TSSecs / 60

'belangrijk: bovenste rij bevat neerslagstations, linker kolom bevat datum/tijd

'haal gegevens voor de bui op
startDate = DataBlock.Cells(2, 1)
EndDate = DataBlock.Cells(DataBlock.Rows.Count, 1)

TotDur = (DataBlock.Rows.Count - 1) * TSMins 'totale duur in minuten
DurDays = WorksheetFunction.RoundDown(TotDur / 60 / 24, 0)
rest = TotDur - DurDays * 24 * 60
DurHours = WorksheetFunction.RoundDown(rest / 60, 0)
rest = rest - DurHours * 60
DurMins = WorksheetFunction.RoundDown(rest, 0)
rest = rest - DurMins
DurSecs = rest * 60

'enventariseer de neerslagstations
For c = 2 To DataBlock.Columns.Count
  stations.Add DataBlock.Cells(1, c)
Next

fn = FreeFile
Open path For Output As #fn
  Print #fn, "*Name of this file: " & path
  Print #fn, "*Date and time of construction: "
  Print #fn, "1"
  Print #fn, "*Aantal stations"
  Print #fn, stations.Count
  Print #fn, "*Namen van stations"
  Dim myStation As Variant
  For Each myStation In stations
    Print #fn, "'" & myStation & "'"
  Next
  Print #fn, "*Aantal gebeurtenissen (omdat het 1 bui betreft is dit altijd 1)"
  Print #fn, "*en het aantal seconden per waarnemingstijdstap"
  Print #fn, " 1  3600 "
  Print #fn, "*Elke commentaarregel wordt begonnen met een * (asteriks)."
  Print #fn, "*Eerste record bevat startdatum en -tijd, lengte van de gebeurtenis in dd hh mm ss"
  Print #fn, "*Het VBA.Format is: yyyymmdd:hhmmss:ddhhmmss"
  Print #fn, "*Daarna voor elk station de neerslag in mm per tijdstap."
  Print #fn, " " & year(startDate) & " " & month(startDate) & " " & day(startDate) & " " & hour(startDate) & " " & Minute(startDate) & " " & Second(startDate) & " " & DurDays & " " & DurHours & " " & DurMins & " " & DurSecs

  For r = 2 To DataBlock.Rows.Count
  
    If Math.Round(r / 1000, 0) * 1000 = r Then
      ProgressRange.Cells(1, 1) = (r - 1) / DataBlock.Rows.Count
    End If
  
    myDate = DataBlock.Cells(r, 1)
    If myDate >= startDate And myDate <= EndDate Then
      myStr = ""
      For i = 1 To stations.Count
        myStr = myStr & " " & DataBlock.Cells(r, 1 + i)
      Next
      myStr = VBA.Trim(myStr)
      Print #fn, myStr
    End If
  Next


Close (fn)
End Sub

Public Sub WriteRKSFile(path As String, DataBlock As Range, StartEndDatesBlock As Range, ProgressRange As Range, TSSecs As Integer)

Dim startDate As Variant, EndDate As Variant
Dim DurDays As Long, DurHours As Long, DurMins As Long, DurSecs As Long, rest As Variant, TotDur As Variant
Dim fn As Long
Dim r As Long, c As Long, i As Long, j As Long, k As Long
Dim stations As Collection
Set stations = New Collection
Dim myDate As Variant, myStr As String
Dim TSMins As Long

TSMins = TSSecs / 60

'IMPORTANT: bovenste rij DataBlock bevat neerslagstations, linker kolom bevat datum/tijd
'IMPORTANT: StartEndDatesBlock bevat 3 kolommen: links het nummer van de bui, midden de startdatum, rechts einddatum. Geen header

If StartEndDatesBlock.Columns.Count <> 3 Then
  MsgBox ("Fout: StartEndDatesBlock moet drie kolommen bevatten: buinummer, startdatum, einddatum")
  End
End If

If Not IsDate(StartEndDatesBlock.Cells(1, 2)) Then
  MsgBox ("Fout: StartEndDatesBlock mag geen header bevatten. Begin meteen met de eerste bui, met in kolommen 2 en 3 start- en einddatum van de bui.")
End If

'check chronologische volgorde events
For i = 2 To StartEndDatesBlock.Rows.Count
  If StartEndDatesBlock.Cells(i, 2) <= StartEndDatesBlock.Cells(i - 1, 2) Then
    MsgBox ("Fout: het blok met start- en einddatums van buien moet in chronologische volgorde staan.")
    End
  End If
Next

'enventariseer de neerslagstations
For c = 2 To DataBlock.Columns.Count
  stations.Add DataBlock.Cells(1, c)
Next

fn = FreeFile
Open path For Output As #fn
  Print #fn, "*Name of this file: " & path
  Print #fn, "* Gebruik de default dataset voor overige invoer (altijd 1 bij bui, 0 bij reeks)"
  Print #fn, "0"
  Print #fn, "*Aantal stations"
  Print #fn, stations.Count
  Print #fn, "*Namen van de stations"
  Dim myStation As Variant
  For Each myStation In stations
    Print #fn, "'" & myStation & "'"
  Next
  Print #fn, "* Number of events in series and time step size [s]"
  Print #fn, StartEndDatesBlock.Rows.Count & " " & TSSecs
  
  'read each of the start- and enddates
  For i = 1 To StartEndDatesBlock.Rows.Count
    ProgressRange.Cells(1, 1) = i / StartEndDatesBlock.Rows.Count
  
    Print #fn, "* Event " & StartEndDatesBlock.Cells(i, 1) & " duration   " & (StartEndDatesBlock.Cells(i, 3) - StartEndDatesBlock.Cells(i, 2)) * 24 & " [hours]"
    Print #fn, "* Start date and time of the event: yyyy mm dd hh mm ss"
    Print #fn, "* Duration of the event           : dd hh mm ss"
    Print #fn, "* Rainfall value per time step [mm/time step]"
    
    'haal gegevens voor de bui op
    startDate = StartEndDatesBlock.Cells(i, 2)
    EndDate = StartEndDatesBlock.Cells(i, 3)

    'bereken de duur van deze bui in dagen, uren, minuten en seconden
    TotDur = (StartEndDatesBlock.Cells(i, 3) - StartEndDatesBlock.Cells(i, 2)) * 24 * 60 'totale duur in minuten
    DurDays = WorksheetFunction.RoundDown(TotDur / 60 / 24, 0)
    rest = TotDur - DurDays * 24 * 60
    DurHours = WorksheetFunction.RoundDown(rest / 60, 0)
    rest = rest - DurHours * 60
    DurMins = WorksheetFunction.RoundDown(rest, 0)
    rest = rest - DurMins
    DurSecs = rest * 60
    
    Print #fn, " " & year(startDate) & " " & month(startDate) & " " & day(startDate) & " " & hour(startDate) & " " & Minute(startDate) & " " & Second(startDate) & " " & DurDays & " " & DurHours & " " & DurMins & " " & DurSecs
        
    'zoek nu in het datablok de startdatum en schrijf waarden weg
    For j = 2 To DataBlock.Rows.Count
      If DataBlock.Cells(j, 1) >= startDate Then
        If DataBlock.Cells(j, 1) < EndDate Then
          myStr = ""
          For k = 1 To stations.Count
            myStr = myStr & " " & DataBlock.Cells(j, 1 + k)
          Next
          myStr = VBA.Trim(myStr)
          Print #fn, myStr
        Else
          Exit For
        End If
      End If
    Next
  Next
  
Close (fn)
End Sub



Public Function WritePRNFile(path As String, DateValueRange As Range, IncludesHeader As Boolean, DateColIdx As Integer, ValColIdx As Integer) As Boolean

Dim i As Long, fn As Long, myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMin As Integer, mySec As Integer, myVal As Variant
fn = FreeFile
Open path For Output As #fn

Dim startRow As Long
If IncludesHeader Then
  startRow = 2
Else
  startRow = 1
End If

'"1998/01/01;00:00:00" 9.1 <
For i = startRow To DateValueRange.Rows.Count
  myYear = year(DateValueRange.Cells(i, DateColIdx))
  myMonth = month(DateValueRange.Cells(i, DateColIdx))
  myDay = day(DateValueRange.Cells(i, DateColIdx))
  myHour = hour(DateValueRange.Cells(i, DateColIdx))
  myMin = Minute(DateValueRange.Cells(i, DateColIdx))
  mySec = Second(DateValueRange.Cells(i, DateColIdx))
  
  If IsNumeric(DateValueRange.Cells(i, ValColIdx)) Then
    myVal = DateValueRange.Cells(i, ValColIdx)
    Print #fn, Chr(34) & Format(myYear, "0000") & "/" & Format(myMonth, "00") & "/" & Format(myDay, "00") & ";" & Format(myHour, "00") & ":" & Format(myMin, "00") & ":" & Format(mySec, "00") & Chr(34) & " " & myVal & " <"
  End If
  
Next

Close (fn)
WritePRNFile = True

End Function

Public Function WritePRNFiles(OutputDir As String, DateValueRange As Range) As Variant

Dim i As Long, j As Long, fn As Long, myYear As Integer, myMonth As Integer, myDay As Integer, myHour As Integer, myMin As Integer, mySec As Integer, myVal As Variant
Dim path As String

'errorhandling
If DateValueRange.Columns.Count < 2 Then
  WritePRNFiles = "Error: range must contain at least two columns"
  Exit Function
End If


For j = 2 To DateValueRange.Columns.Count
  path = OutputDir & "\" & DateValueRange.Cells(1, j) & ".prn"
  fn = FreeFile
  Open path For Output As #fn
  For i = 2 To DateValueRange.Rows.Count
    myYear = year(DateValueRange.Cells(i, 1))
    myMonth = month(DateValueRange.Cells(i, 1))
    myDay = day(DateValueRange.Cells(i, 1))
    myHour = hour(DateValueRange.Cells(i, 1))
    myMin = Minute(DateValueRange.Cells(i, 1))
    mySec = Second(DateValueRange.Cells(i, 1))
    
    If IsNumeric(DateValueRange.Cells(i, j)) Then
      myVal = DateValueRange.Cells(i, j)
      Print #fn, Chr(34) & Format(myYear, "0000") & "/" & Format(myMonth, "00") & "/" & Format(myDay, "00") & ";" & Format(myHour, "00") & ":" & Format(myMin, "00") & ":" & Format(mySec, "00") & Chr(34) & " " & myVal & " <"
    End If
  Next
  
  Close (fn)
Next

WritePRNFiles = "Complete"


End Function


Public Function WRITERRBOUNDARYDATA(MyRange As Range, File3B As String, FileTBL As String) As String
  'Author: Siebe Bosch
  'Date: 21-6-2013
  'first column must contain ID
  'second column must contain summer target level
  'third column must contain winter target level
  Dim r As Long, c As Long, fn1 As Long, fn2 As Long
  Dim id As String, ZP As Variant, WP As Variant
  
  fn1 = FreeFile
  Open File3B For Output As #fn1
  
  fn2 = FreeFile
  Open FileTBL For Output As #fn2
  
  For r = 1 To MyRange.Rows.Count
    id = MyRange.Cells(r, 1)
    ZP = MyRange.Cells(r, 2)
    WP = MyRange.Cells(r, 3)
    
    'BOUN id 'rrcf121212' bl 1 'rrcf121212' is 0 boun
    Print #fn1, "BOUN id '" & id & "' bl 1 '" & id & "' is 0 boun"
    
    '    bn_t ID 'rrcf121212' nm 'rrcf121212' PDIN 1 1 '365;00:00:00' pdin TBLE
    '    '2000/01/01;00:00:00' -0.5 0 <
    '    '2000/04/15;00:00:00' -0.25 0 <
    '    '2000/10/15;00:00:00' -0.5 0 <
    '    tble bn_t
    
    Print #fn2, "BN_T id '" & id & "' nm '" & id & "' PDIN 1 1 '365;00:00:00' pdin TBLE"
    Print #fn2, "'2000/01/01;00:00:00' " & WP & " 0 <"
    Print #fn2, "'2000/04/15;00:00:00' " & ZP & " 0 <"
    Print #fn2, "'2000/10/15;00:00:00' " & WP & " 0 <"
    Print #fn2, "tble bn_t"
  
  Next
  
  
  
  Close (fn1)
  Close (fn2)
  
  WRITERRBOUNDARYDATA = "COMPLETE"
    
End Function


Public Function getDelwaqID(myNum As Integer) As String
  Dim myStr As String, myNumStr As String
  Dim i As Long
  myStr = "Segment"
  myNumStr = VBA.Trim(VBA.str(myNum))
  For i = VBA.Len(myNumStr) + 1 To 5
    myStr = myStr & " "
  Next
  myStr = myStr & myNumStr
  getDelwaqID = myStr
End Function

Public Function StringRightSectionBySubstring(myString As String, mySubString As String) As String
    Dim StringPos As Integer
    StringPos = InStr(1, myString, mySubString, vbTextCompare)
    If StringPos > 0 Then
        StringRightSectionBySubstring = Right(myString, Len(myString) - StringPos + 1)
    Else
        StringRightSectionBySubstring = ""
    End If
End Function

Public Function StringContainsSubstring(myString As String, mySubString As String) As Boolean
    If InStr(1, myString, mySubString, vbTextCompare) > 0 Then
        StringContainsSubstring = True
    Else
        StringContainsSubstring = False
    End If
End Function

Public Function IDFROMSTRING(myStr As String, Optional prefix As String = "", Optional CutoffString As String = "") As String
  Dim CutOffPos As Long
  Dim PrefixPos As Long
  
  PrefixPos = InStr(1, myStr, prefix)
  CutOffPos = InStr(1, myStr, CutoffString)
  
  If prefix = "" And CutoffString = "" Then                           'geen prefix of afbreekstring opgegeven, dus retourneer de hele string
    IDFROMSTRING = myStr
  ElseIf prefix <> "" And CutoffString = "" Then                      'wel prefix opgegeven maar geen afbreekstring
    If PrefixPos > 0 Then
      IDFROMSTRING = VBA.Right(myStr, VBA.Len(myStr) - PrefixPos + 1)         'prefix aangetroffen
    Else                                                              'prefix niet aangetroffen
      IDFROMSTRING = ""
    End If
  ElseIf prefix = "" And CutoffString <> "" Then                      'geen prefix opgegeven, maar wel een afbreekstring
    If CutOffPos > 0 Then
      IDFROMSTRING = VBA.Left(myStr, CutOffPos - 1)                       'afbreekstring aangetroffen
    Else
      IDFROMSTRING = myStr                                            'afbreekstring niet aangetroffen, dus retourneer de hele string
    End If
  ElseIf prefix <> "" And CutoffString <> "" Then                     'zowel prefix als afbreekstring opgegeven
    If PrefixPos > 0 And CutOffPos > 0 And CutOffPos > PrefixPos Then 'prefix en afbreekstring aangetroffen en afbreekstring ligt achter prefix
      IDFROMSTRING = VBA.Mid(myStr, PrefixPos, (CutOffPos - PrefixPos + 1))
    ElseIf prefixpost > 0 And CutOffPos = 0 Then
      IDFROMSTRING = VBA.Right(myStr, VBA.Len(myStr) - PrefixPos + 1)         'prefix aangetroffen, maar afbreekstring niet
    ElseIf PrefixPos = 0 And CutOffPos > 0 Then
      IDFROMSTRING = VBA.Left(myStr, CutOffPos - 1)                       'afbreekstring aangetroffen
    Else
      IDFROMSTRING = ""
    End If
  End If
  
End Function

Public Function RemovePostFix(myStr As String, Postfix As String) As String
  Dim Pos As Integer
  Pos = InStr(myStr, Postfix)
  If Pos > 0 Then
    RemovePostFix = Left(myStr, Pos - 1)
  Else
    RemovePostFix = myStr
  End If
End Function

Public Sub WRITESTOCHASTXMLFILE(MyRange As Range, myPath As String)
  Dim r As Long, c As Long
  Dim fn As Long
  Dim myID As String, myAlias As String
  Dim myHerh As Variant, myH As Variant
  
  fn = FreeFile
  Open myPath For Output As #fn
  
  Print #fn, "<stochasticAnalysis>"
  
  For r = 2 To MyRange.Rows.Count
    myID = MyRange.Cells(r, 1)
    myAlias = MyRange.Cells(r, 2)
    Print #fn, "  <location>"
    Print #fn, "    <id>" & myID & "</id>"
    Print #fn, "    <alias>" & myAlias & "</alias>"
    For c = 3 To MyRange.Columns.Count
      myHerh = MyRange.Cells(1, c)
      myH = MyRange.Cells(r, c)
      Print #fn, "    <result>"
      Print #fn, "      <frequencyEvent>" & 1 / (MyRange.Columns.Count - 2) & "</frequencyEvent>"
      Print #fn, "      <returnPeriodInYears>" & myHerh & "</returnPeriodInYears>"
      Print #fn, "      <jobname>" & myID & "_" & myAlias & "</jobname>"
      Print #fn, "      <exceedanceWaterLevel>" & myH & "</exceedanceWaterLevel>"
      Print #fn, "    </result>"
    Next
    Print #fn, "  </location>"
  Next
  Print #fn, "</stochasticAnalysis>"
  
  Close (fn)
  
End Sub

Public Sub ReplaceDatesInSettingsDat(templateFile As String, Outfile As String, startDate As Date, EndDate As Date)
  'this routine replaces the start- and end date of a simulation in the settings.dat file
  'NOTE: it might be that the Delft_3B.INI file needs adjustment too!
  
  Dim fn As Integer, i As Integer, tmpStr As String
  fn = FreeFile
  Dim fileContent As String, FileRecords() As String
  
  Open templateFile For Input As #fn
    fileContent = Input(VBA.LOF(fn), #fn)
    FileRecords = Split(fileContent, vbCrLf)
  Close (fn)
    
  fn = FreeFile
  Open Outfile For Output As #fn
  For i = 0 To UBound(FileRecords) - 1
    tmpStr = Replace(FileRecords(i), vbCrLf, "")
    If InStr(1, tmpStr, "BeginYear") > 0 Then
      Print #fn, "BeginYear=" & year(startDate)
    ElseIf InStr(1, tmpStr, "BeginMonth") > 0 Then
      Print #fn, "BeginMonth=" & month(startDate)
    ElseIf InStr(1, tmpStr, "BeginDay") > 0 Then
      Print #fn, "BeginDay=" & day(startDate)
    ElseIf InStr(1, tmpStr, "BeginHour") > 0 Then
      Print #fn, "BeginHour=" & hour(startDate)
    ElseIf InStr(1, tmpStr, "BeginMinute") > 0 Then
      Print #fn, "BeginMinute=" & Minute(startDate)
    ElseIf InStr(1, tmpStr, "BeginSecond") > 0 Then
      Print #fn, "BeginSecond=" & Second(startDate)
    ElseIf InStr(1, tmpStr, "EndYear") > 0 Then
      Print #fn, "EndYear=" & year(EndDate)
    ElseIf InStr(1, tmpStr, "EndMonth") > 0 Then
      Print #fn, "EndMonth=" & month(EndDate)
    ElseIf InStr(1, tmpStr, "EndDay") > 0 Then
      Print #fn, "EndDay=" & day(EndDate)
    ElseIf InStr(1, tmpStr, "EndHour") > 0 Then
      Print #fn, "EndHour=" & hour(EndDate)
    ElseIf InStr(1, tmpStr, "EndMinute") > 0 Then
      Print #fn, "EndMinute=" & Minute(EndDate)
    ElseIf InStr(1, tmpStr, "EndSecond") > 0 Then
      Print #fn, "EndSecond=" & Second(EndDate)
    Else
      Print #fn, tmpStr
    End If
  Next
  Close (fn)
   
End Sub

Public Sub ReplaceDatesInDelft3BINI(templateFile As String, Outfile As String, startDate As Date, EndDate As Date)
  'this routine replaces the start- and end date of a simulation in the delft_3b.ini file
  'NOTE: it might be that the settings.dat file needs adjustment too!
  
  Dim fn As Integer, i As Integer, tmpStr As String
  fn = FreeFile
  Dim fileContent As String, FileRecords() As String
  
  Open templateFile For Input As #fn
    fileContent = Input(VBA.LOF(fn), #fn)
    FileRecords = Split(fileContent, vbCrLf)
  Close (fn)
    
  fn = FreeFile
  Open Outfile For Output As #fn
  For i = 0 To UBound(FileRecords) - 1
    tmpStr = Replace(FileRecords(i), vbCrLf, "")
    If InStr(1, tmpStr, "StartTime") = 1 Then
      Print #fn, "StartTime='" & Format(year(startDate), "0000") & "/" & Format(month(startDate), "00") & "/" & Format(day(startDate), "00") & ";" & Format(hour(startDate), "00") & ":" & Format(Minute(startDate), "00") & ":" & Format(Second(startDate), "00") & "'"
    ElseIf InStr(1, tmpStr, "EndTime") = 1 Then
      Print #fn, "EndTime='" & Format(year(EndDate), "0000") & "/" & Format(month(EndDate), "00") & "/" & Format(day(EndDate), "00") & ";" & Format(hour(EndDate), "00") & ":" & Format(Minute(EndDate), "00") & ":" & Format(Second(EndDate), "00") & "'"
    Else
      Print #fn, tmpStr
    End If
  Next
  Close (fn)
   
End Sub

Public Function WritePaved3BRecord(id As String, ar As Variant, lv As Variant, StorDef As String, MeteoStation As String, DWADef As String, SewageSystem As Integer, POCMixm3ps As Variant, POCDWFm3ps As Variant, MixToObjectType As Integer, DWFToObjectType As Integer) As String
Dim myRecord As String
Dim i As Integer
'SewageSystem: 0=mixed 1=separated 2 = improved separated
'MixToObjectType: 0 = openwater, 1 = boundary, 2 = WWTP
'DWFToObjectType: 0 = openwater, 1 = boundary, 3 = WWTP

'PAVE ID 'J1-11vh' ar 155890 lv -4.41 ss 1 sd 'stor_J1-11' qc 0 0 0 qo 1 1 ms 'J1-11' aaf 1 is 0 np 0 dw 'alg_dwa' ro 1 ru 0.2 qh '' pave
myRecord = "PAVE id '" & id & "' ar " & ar & " lv " & lv & " sd '" & StorDef & "' ss " & SewageSystem & " qc 0 " & POCMixm3ps & " " & POCDWFm3ps & " qo " & MixToObjectType & " " & DWFToObjectType & " ms '" & MeteoStation & "' aaf 1 is 0 np 0 dw '" & DWADef & "' ro 1 ru 0.2 qh '' pave"
WritePaved3BRecord = myRecord
End Function

Public Function WriteUnpaved3BRecord(id As String, gwArea As Variant, areas As Range, lv As Variant, StorDef As String, ErnstDef As String, SeepDef As String, InfDef As String, SoilType As Integer, ig As Variant, MaxGW As Variant, MeteoStation As String) As String
Dim myRecord As String
Dim i As Integer

'J1-11ovl' na 16 ar 2083950 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 ga 2083950 lv -4.41 co 3 rc 0 sd 'sd_J1-11ovl' ad 'P04_2' ed 'ad_J1-11ovl' sp 'sp_J1-11ovl' ic 'ic_J1-11ovl' bt 119 ig 0 1.32 su 0 '' gl 10 mg -4.41 ms 'J1-11' is 0 aaf 1 unpv
If areas.Rows.Count <> 16 Then
    MsgBox ("Landuse table requires exactly 16 entries.")
Else
    myRecord = "UNPV id '" & id & "' na 16 ar"
    For i = 1 To areas.Rows.Count
        myRecord = myRecord & " " & areas(i, 1)
    Next
    myRecord = myRecord & " ga " & gwArea
    myRecord = myRecord & " lv " & lv & " co 3 rc 1 sd '" & StorDef & "' ad '' ed '" & ErnstDef & "' sp '" & SeepDef & "' ic '" & InfDef & "' bt " & SoilType & " ig 0 " & ig & " su 0 '' gl 10 mg " & MaxGW & " ms '" & MeteoStation & "' is 0 aaf 1 unpv"
End If
WriteUnpaved3BRecord = myRecord

End Function

 

Public Sub WriteWagModInput(path As String, startRow As Long, DateCol As Long, PrecCol As Long, EvapCol As Long, Optional MeasCol As Long = 0)
  'schrijft een .dat file voor het Wageningen-model
  Dim fn As Long, i As Long
  Dim r As Long
  Dim myYear As String, myMonth As String, myDay As String, myHour As String
  Dim myPrec As String, myEvap As String, myMeas As String
  
  fn = FreeFile
  r = startRow - 1
  Open path For Output As #fn
    Print #fn, "Deze file is geschreven met de ExcelFuncties van Hydroconsult.nl"
    Print #fn, "op: " & Now
    Print #fn, "                                      <-------P<-----ETp<------Qm"
    Print #fn, "datum#                                <----[mm]<----[mm]<----[mm]"
    While Not ActiveSheet.Cells(r + 1, DateCol) = ""
      r = r + 1
      myYear = VBA.Format(year(ActiveSheet.Cells(r, DateCol)), "0000")
      myMonth = VBA.Format(month(ActiveSheet.Cells(r, DateCol)), "00")
      myDay = VBA.Format(day(ActiveSheet.Cells(r, DateCol)), "00")
      myHour = VBA.str(hour(ActiveSheet.Cells(r, DateCol)))
      myPrec = VBA.Format(ActiveSheet.Cells(r, PrecCol), "0.000")
      myEvap = VBA.Format(ActiveSheet.Cells(r, EvapCol), "0.000")
      If MeasCol > 0 Then
        myMeas = VBA.Format(ActiveSheet.Cells(r, MeasCol), "0.000")
      Else
        myMeas = "0.000"
      End If
      
      While Not VBA.Len(myHour) >= 3
        myHour = " " & myHour
      Wend
      While Not VBA.Len(myPrec) >= 14
        myPrec = " " & myPrec
      Wend
      While Not VBA.Len(myEvap) >= 9
        myEvap = " " & myEvap
      Wend
      While Not VBA.Len(myMeas) >= 9
        myMeas = " " & myMeas
      Wend
      Print #fn, myYear & "/" & myMonth & "/" & myDay & myHour & myPrec & myEvap & myMeas
      
    Wend
    
  Close (fn)
End Sub

Public Sub READWAGMODOUTPUT(Path00P As String, startRow As Integer, StartCol As Integer, ByRef Progress As Range)
    Dim fn As Long, c As Integer, r As Long, AddPos As Integer, AddC As Integer
    Dim myStr As String, tmpStr As String
    Dim nLines As Long, iLine As Long
    r = startRow
    c = StartCol
        
    fn = FreeFile
    Open Path00P For Input As #fn
                
    'count the number of lines
    While Not EOF(fn)
      Line Input #fn, myStr
      nLines = nLines + 1
    Wend
    
    Close #fn
    
    Open Path00P For Input As #fn
    Line Input #fn, myStr
    Line Input #fn, myStr
    Line Input #fn, myStr
    Line Input #fn, myStr
    
    'read the header and write it to the worksheet
    AddC = 0
    While Not myStr = ""
        tmpStr = ParseStringByMultiSpaces(myStr)
        ActiveSheet.Cells(r, c + AddC) = VBA.Trim(tmpStr)
        AddC = AddC + 1
    Wend
    
    'read the file content and write it to the worksheet
    While Not EOF(fn)
        AddC = 0
        iLine = iLine + 1
        
        'update the progress cell
        Progress.Cells(1, 1) = iLine / nLines
        DoEvents
        
        r = r + 1
        Line Input #fn, myStr
        While Not myStr = ""
            tmpStr = ParseStringByMultiSpaces(myStr)
            ActiveSheet.Cells(r, c + AddC) = VBA.Trim(tmpStr)
            AddC = AddC + 1
        Wend
    Wend
        

End Sub

Public Function ParseStringByMultiSpaces(ByRef myStr As String) As String
    Dim CharacterFound As Boolean
    Dim i As Integer
    For i = 1 To VBA.Len(myStr)
        If CharacterFound = False And VBA.Mid$(myStr, i, 1) <> " " Then
            CharacterFound = True
        ElseIf CharacterFound = True And VBA.Mid$(myStr, i, 1) = " " Then
            ParseStringByMultiSpaces = Left(myStr, i - 1)
            myStr = VBA.Right$(myStr, VBA.Len(myStr) - i + 1)
            Exit Function
        End If
    Next
    ParseStringByMultiSpaces = myStr
    myStr = ""
End Function

Public Sub WRITEPCRASTERXYZ(ResultsFile As String, DataRange As Range, XColIdx As Integer, YColIdx As Integer, ValColIdx As Integer)
  Dim fn As Long
  Dim r As Long
   
  fn = FreeFile
  Open ResultsFile For Output As #fn
    Print #fn, "field data"
    Print #fn, "3"
    Print #fn, "xcoord"
    Print #fn, "ycoord"
    Print #fn, "max"
    For r = 1 To DataRange.Rows.Count
      Print #fn, DataRange.Cells(r, XColIdx) & " " & DataRange.Cells(r, YColIdx) & " " & DataRange.Cells(r, ValColIdx)
    Next
  Close (fn)
End Sub

Public Function CropFact(myDate As Date, Crop As String) As Variant
Dim CropIdx As Integer, DayNum As Integer
Dim Fact() As String, DayFacts() As String 'record voor 1 dag, gesplitst
Dim DayVals() As Variant
ReDim Fact(1 To 366)

DayNum = DayNumber(myDate, True)

If LCase(Crop) = "grass" Then
  CropIdx = 1
ElseIf LCase(Crop) = "corn" Then
  CropIdx = 2
ElseIf LCase(Crop) = "potatoes" Then
  CropIdx = 3
ElseIf LCase(Crop) = "sugarbeet" Then
  CropIdx = 4
ElseIf LCase(Crop) = "grain" Then
  CropIdx = 5
ElseIf LCase(Crop) = "miscellaneous" Then
  CropIdx = 6
ElseIf LCase(Crop) = "non-arable land" Then
  CropIdx = 7
ElseIf LCase(Crop) = "greenhouse area" Then
  CropIdx = 8
ElseIf LCase(Crop) = "orchard" Then
  CropIdx = 9
ElseIf LCase(Crop) = "bulbous plants" Then
  CropIdx = 10
ElseIf LCase(Crop) = "foliage forest" Then
  CropIdx = 11
ElseIf LCase(Crop) = "pine forest" Then
  CropIdx = 12
ElseIf LCase(Crop) = "nature" Then
  CropIdx = 13
ElseIf LCase(Crop) = "fallow" Then
  CropIdx = 14
ElseIf LCase(Crop) = "vegetables" Then
  CropIdx = 15
ElseIf LCase(Crop) = "flowers" Then
  CropIdx = 16
End If

Fact(1) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(2) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(3) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(4) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(5) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(6) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(7) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(8) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(9) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(10) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(11) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(12) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(13) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(14) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(15) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(16) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(17) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(18) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(19) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(20) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(21) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(22) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(23) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(24) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(25) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(26) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(27) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(28) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(29) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(30) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(31) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.90,1.20,0.95,0.71,0.71,0.00"
Fact(32) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(33) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(34) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(35) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(36) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(37) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(38) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(39) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(40) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(41) = "0.95,0.63,0.63,0.63,0.63,0.95,0.63,0.00,0.63,0.63,0.90,1.20,0.95,0.63,0.63,0.00"
Fact(42) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(43) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(44) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(45) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(46) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(47) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(48) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(49) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(50) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(51) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.90,1.20,0.95,0.50,0.50,0.00"
Fact(52) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(53) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(54) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(55) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(56) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(57) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(58) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(59) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(60) = "0.95,0.40,0.40,0.40,0.40,0.95,0.40,0.00,0.40,0.40,0.90,1.20,0.95,0.40,0.40,0.00"
Fact(61) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(62) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(63) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(64) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(65) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(66) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(67) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(68) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(69) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(70) = "0.95,0.33,0.33,0.33,0.33,0.95,0.33,0.00,0.33,0.33,1.00,1.20,0.95,0.33,0.33,0.00"
Fact(71) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(72) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(73) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(74) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(75) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(76) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(77) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(78) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(79) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(80) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.23,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(81) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(82) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(83) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(84) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(85) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(86) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(87) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(88) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(89) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(90) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(91) = "0.95,0.23,0.23,0.23,0.23,0.95,0.23,0.00,0.23,0.78,1.00,1.20,0.95,0.23,0.23,0.00"
Fact(92) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(93) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(94) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(95) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(96) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(97) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(98) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(99) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(100) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(101) = "1.00,0.23,0.23,0.23,0.65,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.23,0.00"
Fact(102) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(103) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(104) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(105) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(106) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(107) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(108) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(109) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(110) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(111) = "1.00,0.23,0.23,0.23,0.78,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.52,0.00"
Fact(112) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(113) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(114) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(115) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(116) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(117) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(118) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(119) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(120) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(121) = "1.00,0.23,0.23,0.23,0.91,1.00,0.23,0.00,1.04,0.91,1.05,1.20,1.00,0.23,0.65,0.00"
Fact(122) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(123) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(124) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(125) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(126) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(127) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(128) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(129) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(130) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(131) = "1.00,0.52,0.15,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.78,0.00"
Fact(132) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(133) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(134) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(135) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(136) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(137) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(138) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(139) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(140) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(141) = "1.00,0.52,0.65,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,0.91,0.00"
Fact(142) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(143) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(144) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(145) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(146) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(147) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(148) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(149) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(150) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(151) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(152) = "1.00,0.52,0.91,0.52,1.04,1.00,0.15,0.00,1.43,1.04,1.15,1.20,1.00,0.15,1.04,0.00"
Fact(153) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(154) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(155) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(156) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(157) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(158) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(159) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(160) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(161) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(162) = "1.00,0.79,1.05,0.79,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(163) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(164) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(165) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(166) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(167) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(168) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(169) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(170) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(171) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(172) = "1.00,1.05,1.05,1.05,1.18,1.00,0.15,0.00,1.57,1.05,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(173) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(174) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(175) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(176) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(177) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(178) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(179) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(180) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(181) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(182) = "1.00,1.18,1.18,1.18,1.18,1.00,0.15,0.00,1.57,0.92,1.20,1.20,1.00,0.15,1.18,0.00"
Fact(183) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(184) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(185) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(186) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(187) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(188) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(189) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(190) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(191) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(192) = "1.00,1.29,1.16,1.16,1.03,1.00,0.16,0.00,1.68,0.77,1.25,1.20,1.00,0.16,1.03,0.00"
Fact(193) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(194) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(195) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(196) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(197) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(198) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(199) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(200) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(201) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(202) = "1.00,1.27,1.14,1.14,0.89,1.00,0.16,0.00,1.65,0.64,1.25,1.20,1.00,0.16,0.76,0.00"
Fact(203) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(204) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(205) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(206) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(207) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(208) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(209) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(210) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(211) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(212) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(213) = "1.00,1.24,1.12,1.12,0.74,1.00,0.16,0.00,1.61,0.50,1.25,1.20,1.00,0.16,0.16,0.00"
Fact(214) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(215) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(216) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(217) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(218) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(219) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(220) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(221) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(222) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(223) = "1.00,1.21,1.09,1.09,0.61,1.00,0.17,0.00,1.33,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(224) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(225) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(226) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(227) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(228) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(229) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(230) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(231) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(232) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(233) = "1.00,1.19,0.83,1.07,0.17,1.00,0.17,0.00,1.31,0.17,1.10,1.20,1.00,0.17,0.17,0.00"
Fact(234) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(235) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(236) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(237) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(238) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(239) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(240) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(241) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(242) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(243) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.10,1.20,0.90,0.25,0.25,0.00"
Fact(244) = "0.90,1.18,0.83,1.06,0.25,0.90,0.25,0.00,1.18,0.25,1.05,1.20,0.90,0.25,0.25,0.00"
Fact(245) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(246) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(247) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(248) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(249) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(250) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(251) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(252) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(253) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(254) = "0.90,1.17,0.70,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(255) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(256) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(257) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(258) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(259) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(260) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(261) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(262) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(263) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(264) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(265) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(266) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(267) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(268) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(269) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(270) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(271) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(272) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(273) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(274) = "0.90,1.17,0.26,1.05,0.26,0.90,0.26,0.00,1.17,0.26,1.05,1.20,0.90,0.26,0.26,0.00"
Fact(275) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(276) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(277) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(278) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(279) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(280) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(281) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(282) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(283) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(284) = "0.90,0.40,0.30,0.40,0.30,0.90,0.30,0.00,0.30,0.30,1.00,1.20,0.90,0.30,0.30,0.00"
Fact(285) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(286) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(287) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(288) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(289) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(290) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(291) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(292) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(293) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(294) = "0.95,0.45,0.44,0.44,0.44,0.95,0.44,0.00,0.44,0.44,1.00,1.20,0.95,0.44,0.44,0.00"
Fact(295) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(296) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(297) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(298) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(299) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(300) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(301) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(302) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(303) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(304) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(305) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,1.00,1.20,0.95,0.50,0.50,0.00"
Fact(306) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(307) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(308) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(309) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(310) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(311) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(312) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(313) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(314) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(315) = "0.95,0.50,0.50,0.50,0.50,0.95,0.50,0.00,0.50,0.50,0.95,1.20,0.95,0.50,0.50,0.00"
Fact(316) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(317) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(318) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(319) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(320) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(321) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(322) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(323) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(324) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(325) = "0.95,0.71,0.71,0.71,0.71,0.95,0.71,0.00,0.71,0.71,0.95,1.20,0.95,0.71,0.71,0.00"
Fact(326) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(327) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(328) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(329) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(330) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(331) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(332) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(333) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(334) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(335) = "0.95,0.83,0.83,0.83,0.83,0.95,0.83,0.00,0.83,0.83,0.95,1.20,0.95,0.83,0.83,0.00"
Fact(336) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(337) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(338) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(339) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(340) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(341) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(342) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(343) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(344) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(345) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(346) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(347) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(348) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(349) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(350) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(351) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(352) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(353) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(354) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(355) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(356) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(357) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(358) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(359) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(360) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(361) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(362) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(363) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(364) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(365) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"
Fact(366) = "0.95,1.00,1.00,1.00,1.00,0.95,1.00,0.00,1.00,1.00,0.90,1.20,0.95,1.00,1.00,0.00"

DayFacts = VBA.Split(Fact(DayNum), ",")
CropFact = DayFacts(CropIdx - 1)


End Function

Public Function MAKKINKAVG(myDate As Date) As Variant

Dim MAK() As Variant
ReDim MAK(1 To 12, 1 To 31)
Dim myDay As Integer, myMonth As Integer
myDay = day(myDate)
myMonth = month(myMonth)

MAK(1, 1) = 0.2
MAK(1, 2) = 0.167
MAK(1, 3) = 0.197
MAK(1, 4) = 0.24
MAK(1, 5) = 0.163
MAK(1, 6) = 0.22
MAK(1, 7) = 0.227
MAK(1, 8) = 0.22
MAK(1, 9) = 0.19
MAK(1, 10) = 0.2
MAK(1, 11) = 0.22
MAK(1, 12) = 0.23
MAK(1, 13) = 0.29
MAK(1, 14) = 0.253
MAK(1, 15) = 0.23
MAK(1, 16) = 0.207
MAK(1, 17) = 0.267
MAK(1, 18) = 0.317
MAK(1, 19) = 0.243
MAK(1, 20) = 0.267
MAK(1, 21) = 0.233
MAK(1, 22) = 0.19
MAK(1, 23) = 0.257
MAK(1, 24) = 0.247
MAK(1, 25) = 0.203
MAK(1, 26) = 0.323
MAK(1, 27) = 0.273
MAK(1, 28) = 0.29
MAK(1, 29) = 0.353
MAK(1, 30) = 0.343
MAK(1, 31) = 0.39
MAK(2, 1) = 0.357
MAK(2, 2) = 0.403
MAK(2, 3) = 0.37
MAK(2, 4) = 0.403
MAK(2, 5) = 0.427
MAK(2, 6) = 0.367
MAK(2, 7) = 0.357
MAK(2, 8) = 0.44
MAK(2, 9) = 0.467
MAK(2, 10) = 0.433
MAK(2, 11) = 0.44
MAK(2, 12) = 0.437
MAK(2, 13) = 0.597
MAK(2, 14) = 0.56
MAK(2, 15) = 0.487
MAK(2, 16) = 0.603
MAK(2, 17) = 0.5
MAK(2, 18) = 0.517
MAK(2, 19) = 0.587
MAK(2, 20) = 0.617
MAK(2, 21) = 0.583
MAK(2, 22) = 0.647
MAK(2, 23) = 0.697
MAK(2, 24) = 0.713
MAK(2, 25) = 0.67
MAK(2, 26) = 0.713
MAK(2, 27) = 0.647
MAK(2, 28) = 0.69
MAK(2, 29) = 0.729
MAK(3, 1) = 0.753
MAK(3, 2) = 0.68
MAK(3, 3) = 0.76
MAK(3, 4) = 0.727
MAK(3, 5) = 0.9
MAK(3, 6) = 0.907
MAK(3, 7) = 0.793
MAK(3, 8) = 0.903
MAK(3, 9) = 0.807
MAK(3, 10) = 0.973
MAK(3, 11) = 0.837
MAK(3, 12) = 1
MAK(3, 13) = 0.917
MAK(3, 14) = 0.977
MAK(3, 15) = 0.89
MAK(3, 16) = 0.98
MAK(3, 17) = 0.94
MAK(3, 18) = 0.99
MAK(3, 19) = 0.903
MAK(3, 20) = 1.127
MAK(3, 21) = 1.083
MAK(3, 22) = 1.06
MAK(3, 23) = 1.163
MAK(3, 24) = 1.157
MAK(3, 25) = 1.18
MAK(3, 26) = 1.173
MAK(3, 27) = 1.223
MAK(3, 28) = 1.293
MAK(3, 29) = 1.42
MAK(3, 30) = 1.343
MAK(3, 31) = 1.32
MAK(4, 1) = 1.283
MAK(4, 2) = 1.35
MAK(4, 3) = 1.473
MAK(4, 4) = 1.28
MAK(4, 5) = 1.38
MAK(4, 6) = 1.403
MAK(4, 7) = 1.48
MAK(4, 8) = 1.473
MAK(4, 9) = 1.89
MAK(4, 10) = 1.747
MAK(4, 11) = 1.643
MAK(4, 12) = 1.553
MAK(4, 13) = 1.817
MAK(4, 14) = 1.893
MAK(4, 15) = 1.877
MAK(4, 16) = 1.707
MAK(4, 17) = 1.84
MAK(4, 18) = 1.787
MAK(4, 19) = 1.87
MAK(4, 20) = 1.92
MAK(4, 21) = 1.847
MAK(4, 22) = 2.193
MAK(4, 23) = 1.84
MAK(4, 24) = 2.273
MAK(4, 25) = 2.333
MAK(4, 26) = 2#
MAK(4, 27) = 2.203
MAK(4, 28) = 2.067
MAK(4, 29) = 2.22
MAK(4, 30) = 2.267
MAK(5, 1) = 2.243
MAK(5, 2) = 2.323
MAK(5, 3) = 2.23
MAK(5, 4) = 2.26
MAK(5, 5) = 2.337
MAK(5, 6) = 2.18
MAK(5, 7) = 2.303
MAK(5, 8) = 2.4
MAK(5, 9) = 2.553
MAK(5, 10) = 2.403
MAK(5, 11) = 2.647
MAK(5, 12) = 2.687
MAK(5, 13) = 2.583
MAK(5, 14) = 2.783
MAK(5, 15) = 2.803
MAK(5, 16) = 2.91
MAK(5, 17) = 2.793
MAK(5, 18) = 2.84
MAK(5, 19) = 3.007
MAK(5, 20) = 2.707
MAK(5, 21) = 2.547
MAK(5, 22) = 2.953
MAK(5, 23) = 2.727
MAK(5, 24) = 2.737
MAK(5, 25) = 2.72
MAK(5, 26) = 2.887
MAK(5, 27) = 2.723
MAK(5, 28) = 2.737
MAK(5, 29) = 2.93
MAK(5, 30) = 3.157
MAK(5, 31) = 2.91
MAK(6, 1) = 3.063
MAK(6, 2) = 2.783
MAK(6, 3) = 2.39
MAK(6, 4) = 2.773
MAK(6, 5) = 2.94
MAK(6, 6) = 2.663
MAK(6, 7) = 2.533
MAK(6, 8) = 2.853
MAK(6, 9) = 3.09
MAK(6, 10) = 3.123
MAK(6, 11) = 2.867
MAK(6, 12) = 3.263
MAK(6, 13) = 3.353
MAK(6, 14) = 3.06
MAK(6, 15) = 3.02
MAK(6, 16) = 2.807
MAK(6, 17) = 3.063
MAK(6, 18) = 2.74
MAK(6, 19) = 2.877
MAK(6, 20) = 3.023
MAK(6, 21) = 3.16
MAK(6, 22) = 2.59
MAK(6, 23) = 3.15
MAK(6, 24) = 2.757
MAK(6, 25) = 2.76
MAK(6, 26) = 3.053
MAK(6, 27) = 2.613
MAK(6, 28) = 2.673
MAK(6, 29) = 2.683
MAK(6, 30) = 3.293
MAK(7, 1) = 3.067
MAK(7, 2) = 3.01
MAK(7, 3) = 3.163
MAK(7, 4) = 3.4
MAK(7, 5) = 3.34
MAK(7, 6) = 3.173
MAK(7, 7) = 3.327
MAK(7, 8) = 3.173
MAK(7, 9) = 2.947
MAK(7, 10) = 3.013
MAK(7, 11) = 3.103
MAK(7, 12) = 3.383
MAK(7, 13) = 3.033
MAK(7, 14) = 2.887
MAK(7, 15) = 2.88
MAK(7, 16) = 2.513
MAK(7, 17) = 2.757
MAK(7, 18) = 2.683
MAK(7, 19) = 2.713
MAK(7, 20) = 2.643
MAK(7, 21) = 2.613
MAK(7, 22) = 2.8
MAK(7, 23) = 2.997
MAK(7, 24) = 2.787
MAK(7, 25) = 2.653
MAK(7, 26) = 2.453
MAK(7, 27) = 2.54
MAK(7, 28) = 2.72
MAK(7, 29) = 2.943
MAK(7, 30) = 2.85
MAK(7, 31) = 2.85
MAK(8, 1) = 2.717
MAK(8, 2) = 2.763
MAK(8, 3) = 2.787
MAK(8, 4) = 2.85
MAK(8, 5) = 2.747
MAK(8, 6) = 2.98
MAK(8, 7) = 2.77
MAK(8, 8) = 2.44
MAK(8, 9) = 2.67
MAK(8, 10) = 2.597
MAK(8, 11) = 2.53
MAK(8, 12) = 2.573
MAK(8, 13) = 2.707
MAK(8, 14) = 2.797
MAK(8, 15) = 2.653
MAK(8, 16) = 2.557
MAK(8, 17) = 2.393
MAK(8, 18) = 2.52
MAK(8, 19) = 2.59
MAK(8, 20) = 2.447
MAK(8, 21) = 2.47
MAK(8, 22) = 2.28
MAK(8, 23) = 2.407
MAK(8, 24) = 2.4
MAK(8, 25) = 2.427
MAK(8, 26) = 2.383
MAK(8, 27) = 2.273
MAK(8, 28) = 2.263
MAK(8, 29) = 2.32
MAK(8, 30) = 2.22
MAK(8, 31) = 1.957
MAK(9, 1) = 1.877
MAK(9, 2) = 1.88
MAK(9, 3) = 1.877
MAK(9, 4) = 1.887
MAK(9, 5) = 1.86
MAK(9, 6) = 1.987
MAK(9, 7) = 2#
MAK(9, 8) = 1.977
MAK(9, 9) = 1.787
MAK(9, 10) = 1.673
MAK(9, 11) = 1.657
MAK(9, 12) = 1.71
MAK(9, 13) = 1.577
MAK(9, 14) = 1.547
MAK(9, 15) = 1.49
MAK(9, 16) = 1.48
MAK(9, 17) = 1.487
MAK(9, 18) = 1.523
MAK(9, 19) = 1.68
MAK(9, 20) = 1.57
MAK(9, 21) = 1.547
MAK(9, 22) = 1.483
MAK(9, 23) = 1.497
MAK(9, 24) = 1.437
MAK(9, 25) = 1.177
MAK(9, 26) = 1.263
MAK(9, 27) = 1.333
MAK(9, 28) = 1.403
MAK(9, 29) = 1.343
MAK(9, 30) = 1.093
MAK(10, 1) = 1.327
MAK(10, 2) = 1.257
MAK(10, 3) = 1.163
MAK(10, 4) = 1.213
MAK(10, 5) = 1.09
MAK(10, 6) = 0.98
MAK(10, 7) = 1.017
MAK(10, 8) = 0.937
MAK(10, 9) = 0.917
MAK(10, 10) = 1
MAK(10, 11) = 1.037
MAK(10, 12) = 1
MAK(10, 13) = 0.913
MAK(10, 14) = 0.997
MAK(10, 15) = 0.85
MAK(10, 16) = 0.837
MAK(10, 17) = 0.877
MAK(10, 18) = 0.85
MAK(10, 19) = 0.91
MAK(10, 20) = 0.8
MAK(10, 21) = 0.817
MAK(10, 22) = 0.817
MAK(10, 23) = 0.743
MAK(10, 24) = 0.827
MAK(10, 25) = 0.677
MAK(10, 26) = 0.68
MAK(10, 27) = 0.623
MAK(10, 28) = 0.553
MAK(10, 29) = 0.643
MAK(10, 30) = 0.533
MAK(10, 31) = 0.5
MAK(11, 1) = 0.577
MAK(11, 2) = 0.507
MAK(11, 3) = 0.483
MAK(11, 4) = 0.53
MAK(11, 5) = 0.507
MAK(11, 6) = 0.453
MAK(11, 7) = 0.443
MAK(11, 8) = 0.43
MAK(11, 9) = 0.44
MAK(11, 10) = 0.38
MAK(11, 11) = 0.343
MAK(11, 12) = 0.487
MAK(11, 13) = 0.43
MAK(11, 14) = 0.367
MAK(11, 15) = 0.35
MAK(11, 16) = 0.323
MAK(11, 17) = 0.31
MAK(11, 18) = 0.32
MAK(11, 19) = 0.273
MAK(11, 20) = 0.323
MAK(11, 21) = 0.3
MAK(11, 22) = 0.287
MAK(11, 23) = 0.243
MAK(11, 24) = 0.297
MAK(11, 25) = 0.253
MAK(11, 26) = 0.227
MAK(11, 27) = 0.247
MAK(11, 28) = 0.267
MAK(11, 29) = 0.293
MAK(11, 30) = 0.25
MAK(12, 1) = 0.257
MAK(12, 2) = 0.21
MAK(12, 3) = 0.257
MAK(12, 4) = 0.23
MAK(12, 5) = 0.25
MAK(12, 6) = 0.24
MAK(12, 7) = 0.223
MAK(12, 8) = 0.207
MAK(12, 9) = 0.23
MAK(12, 10) = 0.187
MAK(12, 11) = 0.21
MAK(12, 12) = 0.183
MAK(12, 13) = 0.177
MAK(12, 14) = 0.207
MAK(12, 15) = 0.187
MAK(12, 16) = 0.16
MAK(12, 17) = 0.19
MAK(12, 18) = 0.177
MAK(12, 19) = 0.19
MAK(12, 20) = 0.19
MAK(12, 21) = 0.21
MAK(12, 22) = 0.163
MAK(12, 23) = 0.177
MAK(12, 24) = 0.177
MAK(12, 25) = 0.187
MAK(12, 26) = 0.183
MAK(12, 27) = 0.197
MAK(12, 28) = 0.187
MAK(12, 29) = 0.16
MAK(12, 30) = 0.193
MAK(12, 31) = 0.18

MAKKINKAVG = MAK(myMonth, myDay)

End Function

Public Sub DAYSTOHOURS(MyRange As Range, ResultsRow As Long, resultsol As Long, compOption As String)
    Dim i As Long, j As Long, H As Long
    Dim CurDate As Date, curVal As Variant
    Dim newDate As Date, newVal As Variant
    
    'compOption can have the following values:
    'none
    'divide
    
    
    ActiveSheet.Cells(ResultsRow, ResultsCol) = "Datum/Tijd"
    ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "Waarde"
    If MyRange.Columns.Count <> 2 Then
      MsgBox ("Error: het bereik met gegevens moet twee kolommen bevatten: datum en waarde")
    ElseIf compOption = "none" Or compOption = "divide" Then
    
      For i = 1 To MyRange.Rows.Count
        CurDate = MyRange.Cells(i, 1)
        curVal = MyRange.Cells(i, 2)
        For H = 0 To 23
          newDate = CurDate + H / 24
          If compOption = "divide" Then
            newVal = curVal / 24
          Else
            newVal = curVal
          End If
          ResultsRow = ResultsRow + 1
          ActiveSheet.Cells(ResultsRow, ResultsCol) = newDate
          ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = newVal
        Next
      Next
    Else
      MsgBox ("Error: de variabele compOption moet een van de volgende waarden hebben: none of divide")
    End If
    
End Sub


Public Sub EVAPDAYTOHOUR(DateValuesRange As Range, ResultsRow As Long, ResultsCol As Long)
  'deze routine disaggregeert etmaalverdampingssommen naar uurcijfers
  'en hanteert hiervoor een sinusfunctie
  Dim R1 As Long, r2 As Long, r3 As Long
  Dim myDate As Date, myVal As Variant
  Dim newDate As Date, newVal As Variant
  Dim cyclus As Variant
  
  ActiveSheet.Cells(ResultsRow, ResultsCol) = "Datum/Tijd"
  ActiveSheet.Cells(ResultsRow, ResultsCol + 1) = "Uurwaarde verdamping"
  r3 = ResultsRow
  
  For R1 = 1 To DateValuesRange.Rows.Count
    If IsDate(DateValuesRange.Cells(R1, 1)) Then
      myDate = DateValuesRange.Cells(R1, 1)
      myVal = DateValuesRange.Cells(R1, 2)
      
      For r2 = 0 To 23
        cyclus = (-6 + r2) / 24 * 2 * 3.141592 'de positie in de dagelijkse cyclus
        newVal = myVal / 24 * (Math.Sin(cyclus) + 1)
        r3 = r3 + 1
        ActiveSheet.Cells(r3, ResultsCol) = myDate + r2 / 24
        ActiveSheet.Cells(r3, ResultsCol + 1) = newVal
      Next
    End If
  Next

End Sub

Public Function Neerslagtekort(P As Variant, e As Variant, LastTekort As Variant, GewasFactor As Variant) As Variant
  'berekent het neerslagtekort van een gegeven tijdstip met neerslag en verdamping
  Dim NewTekort As Variant
  NewTekort = LastTekort + e * GewasFactor - P 'neerslagtekort = vorig tekort - neerslag + verdamping
  If NewTekort < 0 Then NewTekort = 0          'aanname: overtollige neerslag wordt meteen afgevoerd, dus een reset naar 0
  Neerslagtekort = NewTekort
End Function

Public Function HIRLAMTRANSLATE(GDALBinDir As String, SourceDir As String, TargetDir As String, SourceProj As String, TargetProj As String, GegevensBandCurrentFiles As Integer, GegevensBandPredictionFiles, myDate As Variant)
  Dim i As Long, j As Long, k As Long, L As Long
  Dim InFile As String, outDir As String, Outfile As String, outFile2 As String
  Dim datestr As String, curDateStr As String, tmpStr As String, CurDate As Variant, predictHour As Integer
  Dim myCollection As Collection
  
  Call ShellandWait("setx PATH " & Chr(34) & "C:\GDAL\bin" & Chr(34))
  
  Set myCollection = New Collection
  Set myCollection = ListFilesInFolder(SourceDir)
  For i = 1 To myCollection.Count
    
    'leid de huidige datum/tijd af
    curDateStr = myCollection(i)
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    tmpStr = ParseString(curDateStr, "_")
    CurDate = DATEFROMSTRING(tmpStr, "yyyymmddhh")
    
    If CurDate = myDate Then
    
      'leid de voorspelhorizon van dit bestand af
      tmpStr = ParseString(curDateStr, "_")
      predictHour = tmpStr
    
      'maak nu een uitvoerdirectory aan voor deze datum/tijd, transformeer het bestand en schrijf het ernaar weg
      datestr = year(CurDate) & VBA.Format(month(CurDate), "00") & VBA.Format(day(CurDate), "00") & VBA.Format(hour(CurDate), "00")
      InFile = SourceDir & "\" & myCollection(i)
      
      outDir = TargetDir & "\" & datestr & "\"
      If Not DirectoryExists(outDir) Then Call VBA.MkDir(outDir)

      Outfile = outDir & VBA.Format(predictHour, "000") & ".tif"
      outFile2 = outDir & VBA.Format(predictHour, "000") & ".asc"
    
      'set the path environment and convert the grids
      Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdalwarp.exe" & Chr(34) & " -s_srs " & Chr(34) & SourceProj & Chr(34) & " -t_srs " & Chr(34) & TargetProj & Chr(34) & " " & Chr(34) & InFile & Chr(34) & " " & Chr(34) & Outfile & Chr(34))
      If predictHour = 0 Then
        Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdal_translate.exe" & Chr(34) & " -b " & GegevensBandCurrentFiles & " -of AAIGrid " & Chr(34) & Outfile & Chr(34) & " " & Chr(34) & outFile2 & Chr(34))
      Else
        Call ShellandWait(Chr(34) & GDALBinDir & "\gdal\apps\gdal_translate.exe" & Chr(34) & " -b " & GegevensBandPredictionFiles & " -of AAIGrid " & Chr(34) & Outfile & Chr(34) & " " & Chr(34) & outFile2 & Chr(34))
      End If
    
    End If
    
  Next

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS GIS
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub READASCIIGRID(path As String, ByRef nCols As Long, ByRef nRows As Long, ByRef xllcorner As Variant, ByRef yllcorner As Variant, ByRef cellsize As Variant, ByRef nodata_value As Variant, ByRef data() As Variant)
  
  Dim fn As Long, myStr As String, tmpStr As String
  Dim r As Long, c As Long
  Dim spcpos As Long
  fn = FreeFile
  
  If FileExists(path) Then
    Open path For Input As #fn
    While Not EOF(fn)
      Line Input #fn, myStr
      myStr = VBA.Trim(myStr)
      If InStr(1, myStr, "/*") > 0 Then
        'commentaarregel
      ElseIf InStr(1, VBA.LCase(myStr), "ncols") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nCols = VBA.VAL(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "nrows") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nRows = VBA.VAL(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "xllcorner") > 0 Then
        tmpStr = ParseString(myStr, " ")
        xllcorner = VBA.VAL(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "yllcorner") > 0 Then
        tmpStr = ParseString(myStr, " ")
        yllcorner = VBA.VAL(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "cellsize") > 0 Then
        tmpStr = ParseString(myStr, " ")
        cellsize = VBA.VAL(myStr)
      ElseIf InStr(1, VBA.LCase(myStr), "nodata_value") > 0 Then
        tmpStr = ParseString(myStr, " ")
        nodata_value = VBA.VAL(myStr)
        ReDim data(1 To nRows, 1 To nCols)
      Else
        r = r + 1
        c = 0
        While Not myStr = ""
          c = c + 1
          data(r, c) = VBA.VAL(ParseString(myStr, " "))
        Wend
      End If
    Wend
    Close (fn)
Else
  MsgBox ("Error: opgegeven bestand bestaat niet.")
End If

End Sub


Public Sub WriteASCIIGridFromEquation(fileName As String, a As Variant, b As Variant, c As Variant, Xmin As Variant, Xmax As Variant, ymin As Variant, ymax As Variant, cellsize As Integer)
    'applies the formula for a 2D plane in order to build the grid: z = ax + by + c
    Dim row As Long, col As Long, x As Variant, y As Variant, Z As Variant
    Dim maxr As Long, maxc As Long, myLine As String
    maxr = (ymax - ymin) / cellsize
    maxc = (Xmax - Xmin) / cellsize
    Dim fn As Long
    fn = FreeFile
    Open fileName For Output As #fn
    
    Print #fn, "ncols        " & maxc
    Print #fn, "nrows        " & maxr
    Print #fn, "xllcorner    " & Xmin
    Print #fn, "yllcorner    " & ymin
    Print #fn, "cellsize     " & cellsize
    Print #fn, "NODATA_value " & -999
    
    For row = maxr To 1 Step -1
        myLine = ""
        y = ymin + row - 1
        For col = 1 To maxc
            x = Xmin + col * cellsize - cellsize / 2   'we gaan uit van cell centered
            myLine = myLine & " " & Format(a * x + b * y + c, "0.0000")
        Next
        Print #fn, myLine
    Next
    Close (fn)

End Sub

Public Sub WriteASCIIGridFromMultipleEquations(fileName As String, aVals As Range, bVals As Range, cVals As Range, xMinVals As Range, xMaxVals As Range, yMinVals As Range, yMaxVals As Range, xMinGrid As Variant, xMaxGrid As Variant, yMinGrid As Variant, yMaxGrid As Variant, cellsize As Variant, nodata_value As Variant, useLowest As Boolean, NullOutEdges As Boolean)
    'applies the formula for a 2D plane in order to build multiple grids: z = ax + by + c
    'for the entire domain it takes the highest z-value from all supplied plains
    'NOTE: the a-values, b-values and c-values are sought within COLUMNS.
    Dim row As Long, col As Long, x As Variant, y As Variant, Z() As Variant, zUse As Variant, i As Integer
    ReDim Z(1 To aVals.Count)
    Dim maxr As Long, maxc As Long, myLine As String
    maxr = (yMaxGrid - yMinGrid) / cellsize
    maxc = (xMaxGrid - xMinGrid) / cellsize
    Dim fn As Long
    
    If aVals.Rows.Count > 1 Or bVals.Rows.Count > 1 Or cVals.Rows.Count > 1 Then
        MsgBox ("Error: het bereik met a-, b-, en c-waarden mag alleen uit kolommen bestaan.")
        End
    ElseIf aVals.Columns.Count <> bVals.Columns.Count Or aVals.Columns.Count <> cVals.Columns.Count Then
        MsgBox ("Error: het bereik voor a-, b-, en c-waarden moet gelijk zijn.")
        End
    End If
    
    fn = FreeFile
    Open fileName For Output As #fn
    
    Print #fn, "ncols        " & maxc
    Print #fn, "nrows        " & maxr
    Print #fn, "xllcorner    " & xMinGrid
    Print #fn, "yllcorner    " & yMinGrid
    Print #fn, "cellsize     " & cellsize
    Print #fn, "NODATA_value " & nodata_value
    
    For row = maxr To 1 Step -1
        myLine = ""
        y = yMinGrid + (row - 0.5) * cellsize 'we gaan uit van cell centered
        For col = 1 To maxc
            x = xMinGrid + (col - 0.5) * cellsize   'we gaan uit van cell centered
            If useLowest Then zUse = 999 Else zUse = -999 'initialize the z-value that will be used.
            For i = 1 To aVals.Count
                If x >= xMinVals(, i) And x <= xMaxVals(, i) And y >= yMinVals(, i) And y <= yMaxVals(, i) Then
                    Z(i) = aVals(, i) * x + bVals(, i) * y + cVals(, i)
                    If useLowest = True And Z(i) < zUse Then zUse = Z(i)
                    If useLowest = False And Z(i) > zUse Then zUse = Z(i)
                End If
                If NullOutEdges Then
                    If row = 1 Then
                        zUse = nodata_value
                    ElseIf row = maxr Then
                        zUse = nodata_value
                    ElseIf col = 1 Then
                        zUse = nodata_value
                    ElseIf col = maxc Then
                        zUse = nodata_value
                    End If
                End If
                
            Next
            myLine = myLine & " " & Format(zUse, "0.0000")
        Next
        Print #fn, myLine
    Next
    Close (fn)

End Sub

Public Sub GRIDINTEGERS(path As String, ByRef nCols As Long, ByRef nRows As Long, ByRef xllcorner As Variant, ByRef yllcorner As Variant, ByRef cellsize As Variant, ByRef nodata_value As Variant, ByRef data() As Integer)
  
  Dim fn As Long, myStr As String
  Dim i As Long, j As Long
  fn = FreeFile
  
  Open path For Output As #fn
  Print #fn, "ncols         " & nCols
  Print #fn, "nrows         " & nRows
  Print #fn, "xllcorner     " & xllcorner
  Print #fn, "yllcorner     " & yllcorner
  Print #fn, "cellsize      " & cellsize
  Print #fn, "NODATA_value  " & nodata_value
        
  For i = 1 To nRows
    myStr = ""
    For j = 1 To nCols - 1
      myStr = myStr & data(i, j) & " "
    Next
    Print #fn, myStr & data(i, j)
  Next
  Close (fn)

End Sub
Public Sub ASCII2XYZ(ASCPath As String, XYZPath As String)
  Dim fn As Long, r As Long, c As Long, x As Variant, y As Variant, Z As Variant
  Dim nCols As Long, nRows As Long, xllcorner As Variant, yllcorner As Variant, cellsize As Variant, NodataValue As Variant, data() As Variant
  'converteert een ASCII grid in een bestand met X Y Z
     
  Call READASCIIGRID(ASCPath, nCols, nRows, xllcorner, yllcorner, cellsize, NodataValue, data)
  
  fn = FreeFile
  Open XYZPath For Output As #fn
  
  For r = 1 To nRows
    y = yllcorner + (nRows - r + 0.5) * cellsize
    For c = 1 To nCols
      x = xllcorner + cellsize * (c - 0.5)
      Z = data(r, c)
      If Not Z = NodataValue Then Print #fn, x & " " & y & " " & Z
    Next
  Next
  Close (fn)
  
End Sub

Public Sub READMT940(path As String, Tenaamstelling As String, ByRef row As Integer, StartCol As Integer)
  'MT940 is een bestandsformaat voor rekeningafschriften, o.a. gebruikt door ABN AMRO
  'uitleg: https://nl.wikipedia.org/wiki/MT940
  'specificaties:
  'Elk record dient te worden voorafgegaan door een TAG (:XX:)
  'Een bestand kan één of meer berichten omvatten
  'Een bericht omvat één rekening
  'Een bericht kan meerdere :61: en :86: records omvatten
  ':20: = transactiereferentie
  ':21: = omschrijving
  ':25: = rekeningnummer
  ':28: of :28C: = afschriftinformatie
  ':60F:, :60M: = beginsaldo
  ':61: = transactie/mutatiegegevens
  ':86: omschrijving
  ':62F:, :62M: = eindsaldo
  
  Dim Cpos As Integer
  Dim Dpos As Integer
  Dim Npos As Integer
  Dim Commapos As Integer
  Dim Colonpos As Integer
  Dim Spacepos As Integer
  
  Dim fn As Long, i As Long, r As Long, c As Long, myStr As String, tmpStr As String, CD As String 'credit debet
  Dim mult As Integer
  Dim OmschrijvingActief As Boolean
  
  OmschrijvingActief = False
  Dim Omschrijving As String
  Dim TransactieType As String
  Dim TransactieNummer As String
  Dim TransactieDatum As String
  Dim Tegenrekening As String
  Dim Naam As String
  Dim Bedrag As Double
  Dim machtiging As String
  Dim TransactieOmschrijving As String
  Dim Resttekst As String
  Dim Groep As String
  Dim Categorie As String
  Dim SubCategorie As String
    
  fn = FreeFile
  r = row
  c = StartCol
  
  
  
  ActiveSheet.Range(Cells(r, StartCol).Address, Cells(r + 1000000, StartCol + 11).Address).ClearContents
      
      
  ActiveSheet.Cells(r, c) = "tenaamstelling"
  ActiveSheet.Cells(r, c + 1) = "bedrag"
  ActiveSheet.Cells(r, c + 2) = "transactietype"
  ActiveSheet.Cells(r, c + 3) = "transactienummer"
  ActiveSheet.Cells(r, c + 4) = "transactiedatum"
  ActiveSheet.Cells(r, c + 5) = "tegenrekening"
  ActiveSheet.Cells(r, c + 6) = "naam"
  ActiveSheet.Cells(r, c + 7) = "groep"
  ActiveSheet.Cells(r, c + 8) = "categorie"
  ActiveSheet.Cells(r, c + 9) = "subcategorie"
  ActiveSheet.Cells(r, c + 10) = "machtiging"
  ActiveSheet.Cells(r, c + 11) = "TransactieOmschrijving"
  ActiveSheet.Cells(r, c + 12) = "Resttekst"
        
  Open path For Input As #fn
  
  Close #fn
    
  If FileExists(path) Then
    Open path For Input As #fn
    While Not EOF(fn)
      Line Input #fn, myStr
                        
      If VBA.Left(myStr, 4) = ":20:" Then
        OmschrijvingActief = False
        'nieuw afschrift. Let op: kan meerdere transacties bevatten!
      ElseIf VBA.Left(myStr, 4) = ":61:" Then
        'een nieuwe transactie
        
        'schrijf eerst de voorgaande transactie naar het werkblad
        If Omschrijving <> "" And Bedrag <> 0 Then
            r = r + 1
            ActiveSheet.Cells(r, c) = Tenaamstelling
            ActiveSheet.Cells(r, c + 1) = Bedrag
            
            'we gaan nu de omschrijving parsen en wegschrijven naar het werkblad
            TransactieType = ""
            TransactieNummer = ""
            TransactieDatum = ""
            Tegenrekening = ""
            Naam = ""
            Groep = ""
            Categorie = ""
            SubCategorie = ""
            machtiging = ""
            TransactieOmschrijving = ""
            Resttekst = ""
                        
            While Not Omschrijving = ""
                tmpStr = UCase(Trim(ParseString(Omschrijving, " ")))
                If tmpStr = "SEPA" Then
                    TransactieType = tmpStr
                ElseIf tmpStr = "BEA" Then
                    TransactieType = "BETAALAUTOMAAT"
                    'in het geval van een betaalautomaat volgt nummer, datum, naam en pasnummer
                    TransactieNummer = ParseString(Omschrijving, " ")
                    TransactieDatum = ParseString(Omschrijving, " ")
                    
                    Commapos = VBA.InStr(Omschrijving, ",")
                    If Commapos > 0 Then
                        Naam = VBA.Left(Omschrijving, Commapos - 1)
                    Else
                        Naam = ParseString(Omschrijving, " ")
                    End If
                ElseIf tmpStr = "OVERBOEKING" Then
                    TransactieType = TransactieType & " " & tmpStr
                ElseIf tmpStr = "INCASSO" Then
                    TransactieType = TransactieType & " " & tmpStr
                ElseIf tmpStr = "IBAN:" Then
                    Tegenrekening = ParseString(Omschrijving, " ")
                ElseIf tmpStr = "NAAM:" Then
                    'de naam wordt gezocht door de eerstvolgende dubbleepunt te vinden
                    Colonpos = VBA.InStr(1, Omschrijving, ":")
                    If Colonpos > 0 Then
                        'zoek terugwaarts naar de eerste spatie
                        For Spacepos = Colonpos - 1 To 1 Step -1
                            If VBA.Mid(Omschrijving, Spacepos, 1) = " " Then
                                Exit For
                            End If
                        Next
                        Naam = Trim(VBA.Left(Omschrijving, Spacepos))
                        Omschrijving = VBA.Right(Omschrijving, VBA.Len(Omschrijving) - Spacepos)
                    Else
                        Naam = Omschrijving
                    End If
                ElseIf tmpStr = "MACHTIGING:" Then
                    machtiging = ParseString(Omschrijving, " ")
                ElseIf tmpStr = "OMSCHRIJVING:" Then
                    TransactieOmschrijving = ParseString(Omschrijving, " ")
                Else
                    Resttekst = Resttekst & " " & tmpStr
                End If
            Wend
            
            Call ClassifyExpenseByName(Naam & " " & TransactieOmschrijving & " " & Resttekst, Groep, Categorie, SubCategorie)
            
            ActiveSheet.Cells(r, c + 2) = TransactieType
            ActiveSheet.Cells(r, c + 3) = TransactieNummer
            ActiveSheet.Cells(r, c + 4) = TransactieDatum
            ActiveSheet.Cells(r, c + 5) = Tegenrekening
            ActiveSheet.Cells(r, c + 6) = Naam
            ActiveSheet.Cells(r, c + 7) = Groep
            ActiveSheet.Cells(r, c + 8) = Categorie
            ActiveSheet.Cells(r, c + 9) = SubCategorie
            ActiveSheet.Cells(r, c + 10) = machtiging
            ActiveSheet.Cells(r, c + 11) = TransactieOmschrijving
            ActiveSheet.Cells(r, c + 12) = Resttekst
        End If
        
        'het bedrag staat direct achter de C/D en voor de N
        OmschrijvingActief = False
        Omschrijving = ""
        Cpos = VBA.InStr(1, myStr, "C") 'credit (bijschrijving)
        Dpos = VBA.InStr(1, myStr, "D") 'debet (afschrijving)
        Npos = VBA.InStr(1, myStr, "N")
        
        If Cpos > Dpos Then
        Bedrag = VBA.Mid(myStr, Cpos + 1, Npos - Cpos - 1)
        ElseIf Dpos > Cpos Then
        Bedrag = Replace(VBA.Mid(myStr, Dpos + 1, Npos - Dpos - 1), ",", ".") * -1
        
             
        
               End If
    ElseIf VBA.Left(myStr, 4) = ":86:" Then
      OmschrijvingActief = True
      Omschrijving = VBA.Right(myStr, VBA.Len(myStr) - 4)
      'dit record bevat de omschrijving. Let op: kan vervolgregels bevatten!
    ElseIf VBA.Left(myStr, 1) = ":" Then
        OmschrijvingActief = False
    Else
        'omschrijving continues
        Omschrijving = Omschrijving & " " & myStr
    End If
      
      

    Wend
  Close (fn)
  row = r + 1
Else
  MsgBox ("Error: opgegeven bestand bestaat niet.")
End If
  
End Sub

Public Function ClassifyExpenseByName(Naam As String, ByRef Groep As String, ByRef Categorie As String, ByRef SubCategorie As String)
    
    Groep = ""
    Categorie = ""
    SubCategorie = ""
    
    If InStr(1, Naam, "HEIJN", vbTextCompare) > 0 Or InStr(1, Naam, "OTTEN", vbTextCompare) > 0 Or InStr(1, Naam, "SUEZKANAAL", vbTextCompare) > 0 Or InStr(1, Naam, "FRESH", vbTextCompare) > 0 Or InStr(1, Naam, "PICNIC", vbTextCompare) > 0 Or InStr(1, Naam, "LIDL", vbTextCompare) > 0 Or InStr(1, Naam, "FIFI", vbTextCompare) > 0 Or InStr(1, Naam, "PLUS", vbTextCompare) > 0 Or InStr(1, Naam, "JUMBO", vbTextCompare) > 0 Or InStr(1, Naam, "SPAR", vbTextCompare) > 0 Or InStr(1, Naam, "WEERNEKERS", vbTextCompare) > 0 Or InStr(1, Naam, "ALDI", vbTextCompare) > 0 Or InStr(1, Naam, "ROODENRIJS", vbTextCompare) > 0 Or InStr(1, Naam, "MERKEZ", vbTextCompare) > 0 Or InStr(1, Naam, "FIRAT", vbTextCompare) > 0 Or InStr(1, Naam, "ZILTE ZEE", vbTextCompare) > 0 Or InStr(1, Naam, "VISHANDEL", vbTextCompare) > 0 Then
        Groep = "Boodschappen"
        Categorie = "Voeding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "NOTENBAR", vbTextCompare) > 0 Or InStr(1, Naam, "EKOPLAZA", vbTextCompare) > 0 Or InStr(1, Naam, "VLINDER", vbTextCompare) > 0 Then
        Groep = "Boodschappen"
        Categorie = "Luxe Voeding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "KRUIDVAT", vbTextCompare) > 0 Or InStr(1, Naam, "ETOS", vbTextCompare) > 0 Or InStr(1, Naam, "HAIRSTYLING", vbTextCompare) > 0 Then
        Groep = "Boodschappen"
        Categorie = "Verzorging"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "PLACE", vbTextCompare) > 0 Or InStr(1, Naam, "VOGELKELDER", vbTextCompare) > 0 Then
        Groep = "Huisdieren"
        Categorie = "Voer en benodigdheden"
    ElseIf InStr(1, Naam, "DIERENARTS", vbTextCompare) > 0 Then
        Groep = "Huisdieren"
        Categorie = "Medische kosten"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "PEET", vbTextCompare) > 0 Or InStr(1, Naam, "MCDONALD", vbTextCompare) > 0 Then
        Groep = "Voeding"
        Categorie = "Snacken"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "EXPRESS-SO", vbTextCompare) > 0 Or InStr(1, Naam, "PLSTK", vbTextCompare) > 0 Or InStr(1, Naam, "MYTIKAS", vbTextCompare) > 0 Or InStr(1, Naam, "CLAIRE", vbTextCompare) > 0 Or InStr(1, Naam, "ZEBEDEUS", vbTextCompare) > 0 Or InStr(1, Naam, "LUNCHROOM", vbTextCompare) > 0 Or InStr(1, Naam, "ALBRON", vbTextCompare) > 0 Or InStr(1, Naam, "STRANDCLUB", vbTextCompare) > 0 Or InStr(1, Naam, "KIJKDUINSE", vbTextCompare) > 0 Or InStr(1, Naam, "STRANDPAVILJOEN", vbTextCompare) > 0 Or InStr(1, Naam, "BOSMAN", vbTextCompare) > 0 Or InStr(1, Naam, "VENEZIA", vbTextCompare) > 0 Or InStr(1, Naam, "TSM", vbTextCompare) > 0 Or InStr(1, Naam, "IJSZEE", vbTextCompare) > 0 Or InStr(1, Naam, "SPIJSSALON", vbTextCompare) > 0 Then
        Groep = "Consumpties"
        Categorie = "Buiten de deur"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "DE HAAN", vbTextCompare) > 0 Or InStr(1, Naam, "BP", vbTextCompare) > 0 Or InStr(1, Naam, "ESSO", vbTextCompare) > 0 Or InStr(1, Naam, "SHELL", vbTextCompare) > 0 Or InStr(1, Naam, "TOTAL", vbTextCompare) > 0 Or InStr(1, Naam, "FIETEN", vbTextCompare) > 0 Then
        Groep = "Transport"
        Categorie = "Brandstof"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "PARKING", vbTextCompare) > 0 Or InStr(1, Naam, "PARKEREN", vbTextCompare) > 0 Or InStr(1, Naam, "ACADEMISCH", vbTextCompare) > 0 Then
        Groep = "Transport"
        Categorie = "Parkeren"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "HTM", vbTextCompare) > 0 Or InStr(1, Naam, "ARRIVA", vbTextCompare) > 0 Then
        Groep = "Transport"
        Categorie = "OV"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "GARAGE", vbTextCompare) > 0 Then
        Groep = "Transport"
        Categorie = "AUTO"
        SubCategorie = "Onderhoud"
    ElseIf InStr(1, Naam, "FLETCHER", vbTextCompare) > 0 Then
        Groep = "Vakantie"
        Categorie = "Hotels"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "JOHANNA", vbTextCompare) > 0 Or InStr(1, Naam, "SCHOUSTRA", vbTextCompare) > 0 Or InStr(1, Naam, "OSTERIA", vbTextCompare) > 0 Then
        Groep = "Vakantie"
        Categorie = "Consumpties"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "TUI", vbTextCompare) > 0 Then
        Groep = "Vakantie"
        Categorie = "Reisvakanties"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "BLOEMEN", vbTextCompare) > 0 Or InStr(1, Naam, "BLOEMBOETIEK", vbTextCompare) > 0 Or InStr(1, Naam, "XENOS", vbTextCompare) > 0 Or InStr(1, Naam, "GOUDENREGEN", vbTextCompare) > 0 Or InStr(1, Naam, "HOUTSPEL", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Aankleding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "APOT", vbTextCompare) > 0 Then
        Groep = "Zorgkosten"
        Categorie = "Medicijnen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "BASALT", vbTextCompare) > 0 Then
        Groep = "Zorgkosten"
        Categorie = "Revalidatie"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "JAGRO", vbTextCompare) > 0 Or InStr(1, Naam, "FA-MED", vbTextCompare) > 0 Or InStr(1, Naam, "CMIB", vbTextCompare) > 0 Then
        Groep = "Zorgkosten"
        Categorie = "Tandarts"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "TACKLE", vbTextCompare) > 0 Or InStr(1, Naam, "HENGEL", vbTextCompare) > 0 Or InStr(1, Naam, "ACTION", vbTextCompare) > 0 Or InStr(1, Naam, "BIGBAZAR", vbTextCompare) > 0 Or InStr(1, Naam, "PPRO", vbTextCompare) > 0 Or InStr(1, Naam, "ALIPAY", vbTextCompare) Or InStr(1, Naam, "HOBBYHOEK", vbTextCompare) > 0 Then
        Groep = "Boodschappen"
        Categorie = "Hobby"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "GAMMA", vbTextCompare) > 0 Or InStr(1, Naam, "PRAXIS", vbTextCompare) > 0 Or InStr(1, Naam, "VERPLOEGEN", vbTextCompare) > 0 Or InStr(1, Naam, "KARWEI", vbTextCompare) > 0 Or InStr(1, Naam, "HORNBACH", vbTextCompare) > 0 Or InStr(1, Naam, "TRIPLE S PROJECTS", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Klussen en verbouwen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "VERVAT", vbTextCompare) > 0 Or InStr(1, Naam, "VRIJBUITER", vbTextCompare) > 0 Or InStr(1, Naam, "CAMPINGGAZ", vbTextCompare) > 0 Then
        Groep = "Vakantie"
        Categorie = "Kampeerspullen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "SCHADEVERZEKERING", vbTextCompare) > 0 Or InStr(1, Naam, "SAA VERZEKERING", vbTextCompare) > 0 Then
        Groep = "Verzekeringen"
        Categorie = "Schadeverzekeringen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "LEVENSVE", vbTextCompare) > 0 Then
        Groep = "Verzekeringen"
        Categorie = "Levensverzekeringen"
        SubCategorie = "ALLIANZ"
    ElseIf InStr(1, Naam, "NATIONALE-NEDERLANDEN", vbTextCompare) > 0 Or InStr(1, Naam, "NN VERZEKEREN", vbTextCompare) > 0 Or InStr(1, Naam, "NATIONALE NEDERLANDEN", vbTextCompare) > 0 Then
        Groep = "Verzekeringen"
        Categorie = "Levensverzekeringen"
        SubCategorie = "NN"
    ElseIf InStr(1, Naam, "BLOKWEG", vbTextCompare) > 0 Then
        Groep = "Verzekeringen"
        Categorie = "Schadeverzekeringen"
        SubCategorie = "BLOKWEG"
    ElseIf InStr(1, Naam, "ACHMEA", vbTextCompare) > 0 Or InStr(1, Naam, "AFREKENING INCASSO", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Hypotheek"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "BOSCH", vbTextCompare) > 0 Then
        Groep = "Bijschrijvingen"
        Categorie = "Siebe"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "RUISENDAAL", vbTextCompare) > 0 Then
        Groep = "Bijschrijvingen"
        Categorie = "Lynn"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "UNITED", vbTextCompare) > 0 Then
        Groep = "Bijschrijvingen"
        Categorie = "United Consumers"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "ZWEM", vbTextCompare) > 0 Then
        Groep = "School"
        Categorie = "Zwemles"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "SCHOOLFOTO", vbTextCompare) > 0 Then
        Groep = "School"
        Categorie = "Schoolfoto's"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "CHRISTELIJK", vbTextCompare) > 0 Then
         Groep = "School"
         Categorie = "Elout"
    ElseIf InStr(1, Naam, "GLASHOUWER", vbTextCompare) > 0 Or InStr(1, Naam, "ZEEMAN", vbTextCompare) > 0 Or InStr(1, Naam, "TOM", vbTextCompare) > 0 Or InStr(1, Naam, "TUMBLE", vbTextCompare) > 0 Then
        Groep = "Kleding"
        Categorie = "Kinderkleding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "HEMA", vbTextCompare) > 0 Then
        Groep = "Boodschappen"
        Categorie = "Non-food"
        SubCategorie = "HEMA"
    ElseIf InStr(1, Naam, "BELLE", vbTextCompare) > 0 Or InStr(1, Naam, "2SAMEN", vbTextCompare) > 0 Then
        Groep = "Kinderopvang"
        Categorie = "Crèche"
    ElseIf InStr(1, Naam, "CHARLY", vbTextCompare) > 0 Then
        Groep = "Kinderopvang"
        Categorie = "Oppas"
    ElseIf InStr(1, Naam, "SPEELZA", vbTextCompare) > 0 Then
        Groep = "Kinderopvang"
        Categorie = "Peuterspeelzaal"
    ElseIf InStr(1, Naam, "CONSUMENTENBOND", vbTextCompare) > 0 Or InStr(1, Naam, "NRC", vbTextCompare) > 0 Or InStr(1, Naam, "DONALD", vbTextCompare) > 0 Then
        Groep = "Kranten en bladen"
        Categorie = "Lidmaatschap"
    ElseIf InStr(1, Naam, "AKO", vbTextCompare) > 0 Then
        Groep = "Kranten en bladen"
        Categorie = "Los"
    ElseIf InStr(1, Naam, "BATAVIA", vbTextCompare) > 0 Then
        Groep = "Lidmaatschap"
        Categorie = "Kortingspassen"
    ElseIf InStr(1, Naam, "NRC", vbTextCompare) > 0 Then
        Groep = "Kranten en bladen"
        Categorie = "Krant"
        SubCategorie = "NRC"
    ElseIf InStr(1, Naam, "WELKOM ENERGIE", vbTextCompare) > 0 Or InStr(1, Naam, "ENECO", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Energie"
    ElseIf InStr(1, Naam, "VGZ", vbTextCompare) > 0 Then
        Groep = "Verzekeringen"
        Categorie = "Zorgverzekering"
    ElseIf InStr(1, Naam, "T-MOBILE", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Telefoon, internet en televisie"
    ElseIf InStr(1, Naam, "DUNEA", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Drinkwater"
    ElseIf InStr(1, Naam, "SHOPHOEF", vbTextCompare) > 0 Or InStr(1, Naam, "H&M", vbTextCompare) > 0 Or InStr(1, Naam, "TOPVINTAGE", vbTextCompare) > 0 Or InStr(1, Naam, "BOMMEL", vbTextCompare) > 0 Or InStr(1, Naam, "WOODSTOCK", vbTextCompare) > 0 Or InStr(1, Naam, "FRUUGO", vbTextCompare) > 0 Or InStr(1, Naam, "BONPRIX", vbTextCompare) > 0 Or InStr(1, Naam, "FUNKY", vbTextCompare) > 0 Then
        Groep = "Kleding"
        Categorie = "Kleding"
    ElseIf InStr(1, Naam, "IELM", vbTextCompare) > 0 Then
        Groep = "Kleding"
        Categorie = "Kleding"
        Categorie = "IELM"
    ElseIf InStr(1, Naam, "LIMANGO", vbTextCompare) > 0 Or InStr(1, Naam, "C & A", vbTextCompare) > 0 Then
        Groep = "Kleding"
        Categorie = "Kleding"
        Categorie = "LIMANGO"
    ElseIf InStr(1, Naam, "BOL.COM", vbTextCompare) > 0 Or InStr(1, Naam, "AMAZON", vbTextCompare) > 0 Then
        Groep = "Diversen"
        Categorie = "Mooie spullen"
    ElseIf InStr(1, Naam, "DHL", vbTextCompare) > 0 Or InStr(1, Naam, "POSTNL", vbTextCompare) > 0 Then
        Groep = "Diversen"
        Categorie = "Pakketverzendingen"
    ElseIf InStr(1, Naam, "BETAALVERZOEK", vbTextCompare) > 0 Or InStr(1, Naam, "TIKKIE", vbTextCompare) > 0 Or InStr(1, Naam, "ABN AMRO", vbTextCompare) > 0 Then
        Groep = "Diversen"
        Categorie = "Tikkies"
        SubCategorie = "TIKKIE"
    ElseIf InStr(1, Naam, "WIERTZ", vbTextCompare) > 0 Or InStr(1, Naam, "SMEENK", vbTextCompare) > 0 Or InStr(1, Naam, "HERTOGHS", vbTextCompare) > 0 Or InStr(1, Naam, "HAASTERT", vbTextCompare) > 0 Or InStr(1, Naam, "RIKKEN", vbTextCompare) > 0 Or InStr(1, Naam, "WESTERHOF", vbTextCompare) > 0 Or InStr(1, Naam, "NEEFJES", vbTextCompare) > 0 Or InStr(1, Naam, "HASSELMAN", vbTextCompare) > 0 Or InStr(1, Naam, "VAN DEN HEUVEL", vbTextCompare) > 0 Or InStr(1, Naam, "MARKTPLAATS", vbTextCompare) > 0 Or InStr(1, Naam, "KNETEMAN", vbTextCompare) > 0 Or InStr(1, Naam, "PATTERSON", vbTextCompare) > 0 Then
        Groep = "Tweedehands"
        Categorie = "Marktplaats"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "TIFFANY", vbTextCompare) > 0 Or InStr(1, Naam, "TABAK", vbTextCompare) > 0 Or InStr(1, Naam, "LIL", vbTextCompare) > 0 Or InStr(1, Naam, "GREETZ", vbTextCompare) > 0 Or InStr(1, Naam, "SURPRISE", vbTextCompare) > 0 Or InStr(1, Naam, "CHOCOLADEBEZORGD", vbTextCompare) > 0 Or InStr(1, Naam, "FIEP", vbTextCompare) > 0 Then
        Groep = "Cadeaus"
        Categorie = "Verjaardagen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "EARLYBIRDS", vbTextCompare) > 0 Then
        Groep = "Goede doelen"
        Categorie = "Kinderen"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "WWF.NL", vbTextCompare) > 0 Then
        Groep = "Goede doelen"
        Categorie = "Natuur"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "EVENT", vbTextCompare) > 0 Then
        Groep = "Goede doelen"
        Categorie = "Stichting The Event Foundation"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "IKEA", vbTextCompare) > 0 Or InStr(1, Naam, "LIL NEDERLAND", vbTextCompare) > 0 Or InStr(1, Naam, "WELKOOP", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Woninginrichting"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "GOFUNDME", vbTextCompare) > 0 Or InStr(1, Naam, "LIL NEDERLAND", vbTextCompare) > 0 Then
        Groep = "Onbekend"
        Categorie = "Onbekend"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "KAROO", vbTextCompare) > 0 Or InStr(1, Naam, "LIL NEDERLAND", vbTextCompare) > 0 Then
        Groep = "Kleding"
        Categorie = "Kleding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "BELASTINGDIENST", vbTextCompare) > 0 Or InStr(1, Naam, "LIL NEDERLAND", vbTextCompare) > 0 Then
        Groep = "Transport"
        Categorie = "Wegenbelasting"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "TOYCHAMP", vbTextCompare) > 0 Or InStr(1, Naam, "STICHTING PAY", vbTextCompare) > 0 Or InStr(1, Naam, "LIL NEDERLAND", vbTextCompare) > 0 Then
        Groep = "Kinderen"
        Categorie = "Opvoeding"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "TOYS", vbTextCompare) > 0 Then
        Groep = "Kinderen"
        Categorie = "Speelgoed"
        SubCategorie = ""
    ElseIf InStr(1, Naam, "NETFLIX", vbTextCompare) > 0 Then
        Groep = "Abonnementen"
        Categorie = "Entertainment"
        SubCategorie = "STREAMING"
    ElseIf InStr(1, Naam, "OCW", vbTextCompare) > 0 Or InStr(1, Naam, "PARKLINE", vbTextCompare) > 0 Then
        Groep = "AUTO"
        Categorie = "Parkeren"
    ElseIf InStr(1, Naam, "THUISBEZORGD", vbTextCompare) > 0 Or InStr(1, Naam, "TACO MUNDO", vbTextCompare) > 0 Or InStr(1, Naam, "NYPD", vbTextCompare) > 0 Then
        Groep = "Voeding"
        Categorie = "Thuisbezorgd"
    ElseIf InStr(1, Naam, "REGIONALE BELASTINGGROEP", vbTextCompare) > 0 Or InStr(1, Naam, "REGIONALE BELASTING GROEP", vbTextCompare) > 0 Then
        Groep = "Belastingen"
        Categorie = "Waterschap"
    ElseIf InStr(1, Naam, "HAAG-BELASTINGEN", vbTextCompare) > 0 Or InStr(1, Naam, "GEMEENTE DEN HAAG", vbTextCompare) > 0 Or InStr(1, Naam, "GEMEENTEBELA", vbTextCompare) > 0 Then
        Groep = "Belastingen"
        Categorie = "Gemeente"
    ElseIf InStr(1, Naam, "ONLINE PAYMENTS FOUNDATION", vbTextCompare) > 0 Then
        Groep = "Sport"
        Categorie = "Kleding en assesoires"
    ElseIf InStr(1, Naam, "COACHING", vbTextCompare) > 0 Then
        Groep = "Sport"
        Categorie = "Lessen"
    ElseIf InStr(1, Naam, "UITHOF", vbTextCompare) > 0 Then
        Groep = "Vrijetijd"
        Categorie = "Sportuitjes"
    ElseIf InStr(1, Naam, "PANASJ", vbTextCompare) > 0 Or InStr(1, Naam, "TENNISVERENIGING", vbTextCompare) > 0 Then
        Groep = "Sport"
        Categorie = "Lidmaatschap"
    ElseIf InStr(1, Naam, "TIQETS", vbTextCompare) > 0 Or InStr(1, Naam, "COLIJNSPLAAT", vbTextCompare) > 0 Or InStr(1, Naam, "NATURALIS", vbTextCompare) > 0 Or InStr(1, Naam, "WORLD FORUM", vbTextCompare) > 0 Or InStr(1, Naam, "HOLIDAY", vbTextCompare) > 0 Or InStr(1, Naam, "MADURODAM", vbTextCompare) > 0 Or InStr(1, Naam, "DOEKSEN", vbTextCompare) > 0 Or InStr(1, Naam, "EUROSPAREN", vbTextCompare) > 0 Or InStr(1, Naam, "BALLORIG", vbTextCompare) > 0 Or InStr(1, Naam, "TICKETMASTER", vbTextCompare) > 0 Then
        Groep = "Vrijetijd"
        Categorie = "Uitjes"
    ElseIf InStr(1, Naam, "APPLE", vbTextCompare) > 0 Then
        Groep = "Software"
        Categorie = "Apps"
    ElseIf InStr(1, Naam, "SQULA", vbTextCompare) > 0 Then
        Groep = "School"
        Categorie = "Bijles"
    ElseIf InStr(1, Naam, "SVB", vbTextCompare) > 0 Then
        Groep = "Toeslagen"
        Categorie = "Kinderbijslag"
    ElseIf InStr(1, Naam, "BLOKKER", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Huishoudartikelen"
    ElseIf InStr(1, Naam, "IBOOD", vbTextCompare) > 0 Then
        Groep = "Wonen"
        Categorie = "Gereedschappen"
    ElseIf InStr(1, Naam, "PIANOTHUIS", vbTextCompare) > 0 Or InStr(1, Naam, "WELOVE2DANCE", vbTextCompare) > 0 Then
        Groep = "Kinderen"
        Categorie = "Cultuurlessen"
    ElseIf InStr(1, Naam, "BANKGIRO", vbTextCompare) > 0 Then
        Groep = "Loterijen"
        Categorie = "Bankgiro"
    ElseIf InStr(1, Naam, "PARTY", vbTextCompare) > 0 Then
        Groep = "Feesten en partijen"
        Categorie = "Feestartikelen"
    ElseIf InStr(1, Naam, "BIBL", vbTextCompare) > 0 Then
        Groep = "Kinderen"
        Categorie = "Bibliotheek"
    ElseIf InStr(1, Naam, "LEBKOV", vbTextCompare) > 0 Or InStr(1, Naam, "SIRTAKI", vbTextCompare) > 0 Or InStr(1, Naam, "TIJSTERMA", vbTextCompare) > 0 Or InStr(1, Naam, "SHEREEN", vbTextCompare) > 0 Or InStr(1, Naam, "LOFT", vbTextCompare) > 0 Or InStr(1, Naam, "NOH", vbTextCompare) > 0 Then
        Groep = "Vrijetijd"
        Categorie = "Uit eten"
    ElseIf InStr(1, Naam, "IN3", vbTextCompare) > 0 Or InStr(1, Naam, "VENDINGWORK", vbTextCompare) > 0 Or InStr(1, Naam, "PANNENKOEKEN", vbTextCompare) > 0 Or InStr(1, Naam, "WITZIER", vbTextCompare) > 0 Or InStr(1, Naam, "PROMOTIONAL", vbTextCompare) > 0 Or InStr(1, Naam, "PP RO PAYMENT SERVICES", vbTextCompare) > 0 Then
        Groep = "Onbekend"
        Categorie = "Onbekend"
    ElseIf InStr(1, Naam, "SCIENCE", vbTextCompare) > 0 Then
        Groep = "Inkomsten"
        Categorie = "Loon"
    ElseIf InStr(1, Naam, "LENING", vbTextCompare) > 0 Then
        Groep = "Leningen"
        Categorie = ""
    Else
    End If
    
    
End Function

Public Function MATCHWILDCARD(myStr As String, myMask As String, CaseSensitive As Boolean) As Boolean
  'Date: 8-12-2013
  'Author: Siebe Bosch
  'Description: matches a given string with a string with wildcards
  'Note: only tested for SOMETHING* so far.
  Dim tmpMask As String, tmpStr As String, checkStr As String, i As Integer, startPos As Integer
  Dim maskPart As String, partPos As Integer
  
  'if case insensitive, convert both strings to uppercase
  If CaseSensitive = False Then
    myStr = VBA.UCase(myStr)
    myMask = VBA.UCase(myMask)
  End If
  
  'create a new string that consists of asteriskses only and that has the length of myStr
  For i = 1 To VBA.Len(myStr)
    checkStr = checkStr & "*"
  Next
  
  'now start parsing the mask in order to find its components (disregarding the wildcards for now)
  startPos = 1
  tmpMask = myMask
  While Not tmpMask = ""
    maskPart = ParseString(tmpMask, "*")
    partPos = InStr(startPos, myStr, maskPart, vbBinaryCompare)
    If partPos > 0 Then
      'embed the string we found in checkStr, at the exact same location
      checkStr = Left(checkStr, partPos - 1) & maskPart & VBA.Right(checkStr, VBA.Len(checkStr) - (partPos - 1) - VBA.Len(maskPart))
    End If
  Wend
  
  'now that we have a checkStr that only consists of * and parts from the mask, we can reduce it to its minimum
  'and check whether it matches our original mask
  While InStr(1, checkStr, "**") > 0
    checkStr = VBA.Replace(checkStr, "**", "*")
  Wend
  
  If checkStr = myMask Then
    MATCHWILDCARD = True
  Else
    MATCHWILDCARD = False
  End If

End Function

Public Sub CONCATENATECOMBINATIONS(MyRange As Range, startRow As Integer, ResultsCol As Integer)
    'this function combines all unique values from all columns in the given range (except for empty cells)
    'and concatenates them
    Dim r As Long, n As Long
    Dim Found As Boolean
    r = startRow
    
    Dim i As Long, j As Long
    Dim myStr As String
    Dim newArray() As Integer
    ReDim newArray(1 To MyRange.Columns.Count)
    While MileageOneUp(1, MyRange.Rows.Count, newArray)
    
        If newArray(1) = 21 And newArray(2) = 21 And newArray(3) = 20 Then Stop
        
        n = n + 1
        Found = True
        myStr = ""
        For j = 1 To MyRange.Columns.Count
            If MyRange.Cells(newArray(j), j) = "" Then Found = False
            myStr = myStr & MyRange.Cells(newArray(j), j)
        Next
        If Found Then
          r = r + 1
          ActiveSheet.Cells(r, ResultsCol) = myStr
        End If
        If n > MyRange.Rows.Count ^ MyRange.Columns.Count Then End
    Wend

End Sub

Public Function ReadEntireTextFile(myPath) As String
  Dim fn As Long, myStr As String
  Dim fileContent As String

  'reads the entire file to memory
  Open myPath For Input As #fn
  
  If FileExists(myPath) Then
    Open myPath For Input As #fn
    fileContent = VBA.Input(LOF(ifn), ifn)
    Close #fn
  Else
    MsgBox ("Error: file does not exist: " & myPath)
    End
  End If
    
  'return the result
  ReadEntireTextFile = fileContent

End Function


Public Function SetFilenameLengthByInsertingZeroes(HostFolder As String, TargetLength As Integer, InsertZeroesAfterCharacterNo As Integer) As Boolean
    Dim FileSystem As Object
    Dim HostFolder As String
    Dim NewFileName As String

    ' Create an instance of the File System Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")

    ' Loop through all files in the directory
    For Each file In FileSystem.GetFolder(HostFolder).Files
            If Len(file.Name) < RequiredLength Then
                'add leading zeroes
                Dim nZeroes As Integer
                nZeroes = TargetLength - Len(file.Name)
                NewFileName = Left(file.Name, InsertZeroesAfterCharacterNo)
                For i = 1 To nZeroes
                  NewFileName = NewFileName & "0"
                Next
                NewFileName = NewFileName & Right(file.Name, Len(file.Name) - InsertZeroesAfterCharacterNo)
                Debug.Print (file.Name)
                Debug.Print (NewFileName)
                
                FileSystem.MoveFile file, HostFolder & "\" & NewFileName
            End If
    Next file
    SetFilenameLengthByInsertingZeroes = True
End Function



Public Sub JoinNodes(MyRange As Range, IDCol As Long, XCol As Long, YCol As Long, rIDcol As Long, rXcol As Long, rYcol As Long, Mergedistance As Variant, Optional ResultsNodePrefix As String = "", Optional BNACol As Long = 0)
  'maakt nieuwe knopen aan door knopen die dicht bijeen liggen samen te voegen. Handig als lozingspunten van meerdere afwateringseenheden dicht bijeen liggen.
  Dim JoinedNodes As New Collection
  Dim JoinedNode As clsMultiNodeObject, Node As clsNode
  Dim i As Long, j As Long, k As Long, n As Long, myDist As Variant
  Dim Found As Boolean
    
  For i = 1 To MyRange.Rows.Count
    If i = 1 Then
      n = 1
      Set JoinedNode = New clsMultiNodeObject
      JoinedNode.id = ResultsNodePrefix & n
      Set Node = New clsNode
      Node.id = MyRange.Cells(i, IDCol).value
      Node.x = MyRange.Cells(i, XCol)
      Node.y = MyRange.Cells(i, YCol)
      Call JoinedNode.AddNode(Node)
      Call JoinedNodes.Add(JoinedNode)
    Else
      Found = False
      Set Node = New clsNode
      Node.id = MyRange.Cells(i, IDCol)
      Node.x = MyRange.Cells(i, XCol)
      Node.y = MyRange.Cells(i, YCol)
      
      For j = 1 To JoinedNodes.Count
        Set JoinedNode = JoinedNodes(j)
        myDist = PointDistance(JoinedNode.XAvg, JoinedNode.YAvg, Node.x, Node.y)
        If myDist <= Mergedistance Then
          Found = True
          Call JoinedNode.AddNode(Node)
        End If
      Next
      If Not Found Then
        n = n + 1
        Set JoinedNode = New clsMultiNodeObject
        JoinedNode.id = ResultsNodePrefix & n
        Call JoinedNode.AddNode(Node)
        Call JoinedNodes.Add(JoinedNode)
      End If
    End If
  Next

  'schrijf de resultaten weg
  For i = 1 To MyRange.Rows.Count
    For j = 1 To JoinedNodes.Count
      Set JoinedNode = JoinedNodes(j)
      For k = 1 To JoinedNode.Nodes.Count
        Set Node = JoinedNode.Nodes(k)
        If MyRange.Cells(i, IDCol) = Node.id Then
          MyRange.Cells(i, rIDcol) = JoinedNode.id
          MyRange.Cells(i, rXcol) = JoinedNode.XAvg
          MyRange.Cells(i, rYcol) = JoinedNode.YAvg
          Exit For
        End If
      Next
    Next
  Next
  
  'optie BNA-string wegschrijven
  If BNACol > 0 Then
    For j = 1 To JoinedNodes.Count
      Set JoinedNode = JoinedNodes(j)
      MyRange.Cells(j, BNACol) = BNAString(JoinedNode.id, JoinedNode.XAvg, JoinedNode.YAvg)
    Next
  End If
End Sub


'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'-----------------------------------------STRINGBEWERKINGEN--------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Function VERWIJDERDAGNAAMUITDATUM(myString As String) As String
  myString = VBA.LCase(myString)
  myString = VBA.Replace(myString, "maandag", "")
  myString = VBA.Replace(myString, "dinsdag", "")
  myString = VBA.Replace(myString, "woensdag", "")
  myString = VBA.Replace(myString, "donderdag", "")
  myString = VBA.Replace(myString, "vrijdag", "")
  myString = VBA.Replace(myString, "zaterdag", "")
  myString = VBA.Replace(myString, "zondag", "")
  myString = VBA.Trim(myString)
  VERWIJDERDAGNAAMUITDATUM = myString

End Function

Public Function MAKEXMLTOKEN(myToken As String, myValue As String) As String
  MAKEXMLTOKEN = myToken & "=" & VBA.str(34) & myValue & VBA.str(34)
End Function

Public Function getDoubleFromXMLRecord(xmlStr As String, TokenID As String) As Variant
  Dim result As String
  result = VBA.LCase(xmlStr)
  result = VBA.Replace(result, "<" & VBA.LCase(TokenID) & ">", "")
  result = VBA.Replace(result, "</" & VBA.LCase(TokenID) & ">", "")
  result = VBA.Trim(result)
  getDoubleFromXMLRecord = result
End Function


Public Function STRINGPOSITIE(SearchString As String, SeekString As String, Optional startPos As Long = 1) As Long
  Dim myPos As Long
  myPos = InStr(startPos, SearchString, SeekString)
  STRINGPOSITIE = myPos
End Function

Public Function ReplaceString(SearchStr As String, FindStr As String, ReplaceStr As String) As String
  ReplaceString = VBA.Replace(SearchStr, FindStr, ReplaceStr, , , vbTextCompare)
End Function

Public Sub REPLACESTRINGINALLFILES(SearchDir As String, FindStr As String, ReplaceStr As String)
  Dim myCollection As Collection, myFile As String, myContent As String, Found As Boolean
  Dim fn As Long, of As Long, i As Long
  
  Set myCollection = New Collection
  Set myCollection = ListFilesInFolder(SearchDir)
  For i = 1 To myCollection.Count
    myFile = SearchDir & "\" & myCollection.Item(i)
    myFile = ReplaceString(myFile, "\\", "\")       'make sure we only have one backslash at a time in the path
    Found = False
    fn = FreeFile
    Open myFile For Input As #fn
      If LOF(fn) > 0 Then
        myContent = Input(LOF(fn), fn)
        If InStr(1, myContent, FindStr, vbTextCompare) > 0 Then
          myContent = ReplaceString(myContent, FindStr, ReplaceStr)
          Found = True
        End If
      End If
    Close
    
    If Found Then
      of = FreeFile
      Open myFile For Output As #of
      Print #of, myContent
      Close #of
    End If
  Next
  
End Sub

Public Function DOUBLEIDSINSTRINGCOLLECTION(myCollection As Collection, ByRef doubleStr As String) As Boolean
  'checkt of een collectie van strings dubbele waarden bevat
  Dim i As Long, j As Long
  
  DOUBLEIDSINSTRINGCOLLECTION = False
  For i = 1 To myCollection.Count
    For j = i + 1 To myCollection.Count
      If myCollection(i) = myCollection(j) Then
        doubleStr = myCollection(i)
        DOUBLEIDSINSTRINGCOLLECTION = True
        Exit Function
      End If
    Next
  Next

End Function

Public Function TRIMUSINGCUSTOMSTRING(myStr As String, myTrimStr As String, Optional CaseSensitive As Boolean = False) As String

If Not CaseSensitive Then
  While VBA.Left(VBA.LCase(myStr), VBA.Len(myTrimStr)) = VBA.LCase(myTrimStr)
    myStr = VBA.Right(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
  While VBA.Right(VBA.LCase(myStr), VBA.Len(myTrimStr)) = VBA.LCase(myTrimStr)
    myStr = VBA.Left(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
Else
  While VBA.Left(myStr, VBA.Len(myTrimStr)) = myTrimStr
    myStr = VBA.Right(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
  While VBA.Right(myStr, VBA.Len(myTrimStr)) = myTrimStr
    myStr = VBA.Left(myStr, VBA.Len(myStr) - VBA.Len(myTrimStr))
  Wend
End If

TRIMUSINGCUSTOMSTRING = myStr

End Function

Public Function UnifyString(myStr As String) As String
  'deze functie uniformeert eeen string door de uppercase te nemen en hem te VBA.Trimmen.
  'handig om te gebruiken als key in collections
  UnifyString = VBA.UCase(VBA.Trim(myStr))
End Function

Public Function IsBankNumber(myStr As String) As Boolean
  myStr = VBA.Trim(myStr)
  If Mid(myStr, 3, 1) = "." And VBA.Mid(myStr, 6, 1) = "." And VBA.Mid(myStr, 9, 1) = "." Then
    IsBankNumber = True
  Else
    IsBankNumber = False
  End If
End Function

Public Function FindNearestObjectInRange(x As Variant, y As Variant, SearchListRange As Range, IDColIdx As Long, XColIdx As Long, YColIdx As Long) As String

Dim Dist As Variant, tmpDist As Variant, id As String, tmpID As String, r As Long, c As Long

'initialiseren
Dist = Sqr((x - SearchListRange.Cells(1, XColIdx)) ^ 2 + (y - SearchListRange.Cells(1, YColIdx)) ^ 2)
id = SearchListRange.Cells(1, IDColIdx)

For r = 2 To SearchListRange.Rows.Count
  tmpDist = Sqr((x - SearchListRange.Cells(r, XColIdx)) ^ 2 + (y - SearchListRange.Cells(r, YColIdx)) ^ 2)
  tmpID = SearchListRange.Cells(r, IDColIdx)
  If tmpDist < Dist Then
    Dist = tmpDist
    id = tmpID
  End If
Next

FindNearestObjectInRange = id

End Function

'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'-----------------------------------------BESTANDEN----------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------

Public Function OpenSingleFile() As String
  Dim Filter As String, Title As String
  Dim FilterIndex As Integer
  Dim fileName As Variant
  
  ' File filters
  Filter = "MT940 Files (*.sta),*.sta, All Files (*.*),*.*"
  FilterIndex = 3

  ' Set Dialog Caption
  Title = "Selecteer een bestand."
  'ChDrive ("C")
  'ChDir ("E:\Chapters\chap14")
  With Application
    ' Set File Name to selected File
    fileName = .GetOpenFilename(Filter, FilterIndex, Title)
    ' Reset Start Drive/Path
    ChDrive (VBA.Left(.DefaultFilePath, 1))
    ChDir (.DefaultFilePath)
  End With

 ' Exit on Cancel
 If fileName = False Then
    MsgBox "No file was selected."
    Exit Function
  End If
  OpenSingleFile = fileName
End Function

Public Function ListFilesInFolder(SourceFolderName As String, Optional EXT As String = "*") As Collection
  'The macro example below assumes that your VBA project has added a reference to the Microsoft Scripting Runtime library.
  'You can do this from within the VBE by selecting the menu Etra, References and selecting Microsoft Scripting Runtime.
  
  ' lists information about the files in SourceFolder
  ' example: ListFilesInFolder "C:\FolderName\", True
  Dim myFile As String
  Dim myCollection As Collection
  Set myCollection = New Collection
  
  myFile = dir$(SourceFolderName & "\*." & EXT)
  Do While myFile <> ""
    myCollection.Add myFile
    myFile = dir$
  Loop
  Set ListFilesInFolder = myCollection

End Function

Public Function DirectoryExists(DName As String) As Boolean

Dim sDummy As String
On Error Resume Next

If VBA.Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = dir$(DName & "*.*", vbDirectory)
DirectoryExists = Not (sDummy = "")

End Function

Public Function CONTAINSKEY(ByRef col As Collection, ByVal key As Variant) As Boolean

Dim obj As Variant
On Error GoTo err
  CONTAINSKEY = True
  obj = col(key)
  Exit Function
err:
  CONTAINSKEY = False

End Function

Public Function CONTAINSKEY_BYOBJECTID(ByRef col As Collection, ByVal id As String) As Boolean

'uses the .ID element of the objects in a collection as a key
'this is because VBA has no way of retrieving objects from a collection by Key
'note: this only works if the elements of the collection actually HAVE an element named ID

Dim i As Long
For i = 1 To col.Count
  If VBA.Trim(VBA.UCase(col.Item(i).id)) = VBA.Trim(VBA.UCase(id)) Then
    CONTAINSKEY_BYOBJECTID = True
    Exit Function
  End If
Next

'not found
CONTAINSKEY_BYOBJECTID = False


End Function

Public Sub DELETESHAPEFILE(path As String)
  Dim myPath As String
  myPath = path
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = Replace(path, ".shp", ".dbf")
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = Replace(path, ".shp", ".shx")
  If FileExists(myPath) Then Call DeleteFile(myPath)
  myPath = Replace(path, ".shp", ".prj")
  If FileExists(myPath) Then Call DeleteFile(myPath)
End Sub

Public Sub MoveFile(FromDir As String, ToDir As String, fileName As String)
  Dim FromFile As String, ToFile As String
  FromFile = FromDir & "\" & fileName
  ToFile = ToDir & "\" & fileName

  If FileExists(FromFile) Then
    If DirectoryExists(ToDir) Then
      Call FileCopy(FromFile, ToFile)
      Call Kill(FromFile)
    Else
      MsgBox ("Error: target directory does not exist:" & ToDir)
    End If
  Else
      MsgBox ("Error: file does not exist:" & FromFile)
  End If

End Sub

Public Sub DIRECTORYCOPY(FromDir As String, ToDir As String)
  'This example copy all files and subfolders from FromPath to ToPath.
  'Note: If ToPath already exist it will overwrite existing files in this folder
  'if ToPath not exist it will be made for you.
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(FromDir) = False Then
        MsgBox FromDir & " doesn't exist"
        Exit Sub
    End If
    FSO.CopyFolder Source:=FromDir, Destination:=ToDir
End Sub

Public Function FOLDERBROWSER(strPath As String) As String
  Dim fldr As FileDialog
  Dim sItem As String
  Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
  With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
  End With
NextCode:
  FOLDERBROWSER = sItem
  Set fldr = Nothing
End Function

Public Sub ReplaceInFile(InFile As String, Outfile As String, ReplaceString As String, ReplaceByString As String)
  
  Dim fn As Long, fn2 As Long, myStr As String
  fn = FreeFile
  Open InFile For Input As #fn
  fn2 = FreeFile
  Open Outfile For Output As #fn2
  
  While Not EOF(fn)
    Line Input #fn, myStr
    myStr = Replace(ReplaceString, myStr, ReplaceByString)
    Print #fn2, myStr
  Wend
  
  Close (fn)
  Close (fn2)

End Sub

Public Function ReplaceInString(SourceStr As String, ReplaceStr As String, ReplaceBy As String) As String
    ReplaceInString = Replace(SourceStr, ReplaceStr, ReplaceBy)
End Function

Function FileNameFromPath(path) As String
    FileNameFromPath = Right(path, InStrRev(path, "\" - 1))
End Function
Function DirFromPath(path) As String
   DirFromPath = Left(path, InStrRev(path, "\"))
End Function

Function GetDirectory(path) As String
   GetDirectory = Left(path, InStrRev(path, "\"))
End Function

Function WorkSheetExists(wksName As String) As Boolean
  'checks of een worksheet al bestaat
  On Error Resume Next
  WorkSheetExists = CBool(Len(Worksheets(wksName).Name) > 0)
End Function

Public Function SumRange(MyRange As Range) As Variant
    Dim CurCell As Object
    Dim mySum As Variant
    For Each CurCell In MyRange
      mySum = mySum + CurCell.value
    Next
    SumRange = mySum
    Exit Function
End Function

Public Function FRACTIONOFDAYSUM(myDateTimeCell As Range, DateTimeCol As Long, valuesCol As Long) As Variant
  'Deze functie rekent uit welk aandeel van de dagsom in een bepaalde cel staat
  'Dit betekent dat je moet opgeven: de kolom waarin datum/tijd staat, de kolom waarin de bijbehorende waarden staan
  'én natuurlijk de cel met de datum/tijd waarvoor je de fractie wilt weten en de cel waarin de waarde staat.
  'de functie deelt de waarde uit de gezochte cel door de som van de waarden van alle cellen die op dezelfde datum vallen
  
  Dim myDay As Variant
  Dim myYear As Variant
  Dim mySum As Variant
  Dim myCell As Object
  Dim myValue As Variant
  Dim nCells As Long
  Dim r As Long
  Dim Done As Boolean
  
  myDay = day(myDateTimeCell.value)
  myYear = year(myDateTimeCell.value)
  myValue = ActiveSheet.Cells(myDateTimeCell.row, valuesCol).value
  mySum = myValue
  nCells = 1
  
  If myDateTimeCell.Count <> 1 Then
    MsgBox ("Error: één cel selecteren voor huidige datum/tijd")
  End If
  
  'we lopen vanaf de gevraagde cel omhoog tot de datum verschilt
  r = myDateTimeCell.row
  Done = False
  While Not Done
    r = r - 1
    If r > 0 And IsDate(ActiveSheet.Cells(r, DateTimeCol)) Then
      If day(ActiveSheet.Cells(r, DateTimeCol)) = myDay And year(ActiveSheet.Cells(r, DateTimeCol)) = myYear Then
        nCells = nCells + 1
        mySum = mySum + ActiveSheet.Cells(r, valuesCol)
      Else
        Done = True
      End If
    Else
      Done = True
    End If
  Wend
  
  'en nu omlaag
  r = myDateTimeCell.row
  Done = False
  While Not Done
    r = r + 1
    If IsDate(ActiveSheet.Cells(r, DateTimeCol)) Then
      If day(ActiveSheet.Cells(r, DateTimeCol)) = myDay And year(ActiveSheet.Cells(r, DateTimeCol)) = myYear Then
        nCells = nCells + 1
        mySum = mySum + ActiveSheet.Cells(r, valuesCol)
      Else
        Done = True
      End If
    Else
      Done = True
    End If
  Wend
  
  If mySum = 0 Then
    FRACTIONOFDAYSUM = 1 / nCells
  Else
    FRACTIONOFDAYSUM = myValue / mySum
  End If

End Function

Public Function IsRangeAscending(MyRange As Range) As Boolean
'checkt of een range (1e kolom) een oplopende volgorde heeft
Dim r As Long
IsRangeAscending = True
  If MyRange.Rows.Count > 1 Then
    For r = 2 To MyRange.Rows.Count
      If MyRange.Cells(r, 1).value < MyRange.Cells(r - 1, 1).value Then
        IsRangeAscending = False
      End If
    Next
  Else
    IsRangeAscending = True
  End If
End Function


Public Function MinYFromXYRange(myWorksheet As String, myXRange As Range, myYRange As Range, Optional fromX As Variant = -10000000000000#, Optional toX As Variant = 10000000000000#) As Variant
Dim row As Long, curSheet As String
'retrieves te lowest Y value from a Range with X and Y values
'XcolIdx is the index number of the column within the range in which the X values can be found
'YColIdx is the index number of the column within the range in which the Y values can be found
'fromX and toX are optional and can be used to restrict the search to the part of the range where X falls between these values
curSheet = ActiveWorkbook.ActiveSheet.Name

If myXRange.Rows.Count <> myYRange.Rows.Count Then
  MsgBox ("Error in function MinYFromXYRange. Ranges must be of equal length.")
  Exit Function
ElseIf myXRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MinYFromXYRange. Range containing X values must consist of only one column.")
  Exit Function
ElseIf myYRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MinYFromXYRange. Range containing Y values must consist of only one column.")
  Exit Function
End If

MinYFromXYRange = 10000000000000#
Worksheets(myWorksheet).Activate
For row = 1 To myXRange.Rows.Count
  If IsNumeric(myXRange.Cells(row, 1)) And IsNumeric(myYRange.Cells(row, 1)) Then
    If myYRange.Cells(row, 1) < MinYFromXYRange And myXRange.Cells(row, 1) >= fromX And myXRange.Cells(row, 1) <= toX Then
      MinYFromXYRange = myYRange.Cells(row, 1)
    End If
  Else
    'MsgBox ("Error in function MinYFromXYRange: non numeric value encountered in row index " & row & " of the data range.")
    'Exit Function
  End If
Next row

End Function

Public Function MaxYFromXYRange(myWorksheet As String, myXRange As Range, myYRange As Range, Optional fromX As Variant = -10000000000000#, Optional toX As Variant = 10000000000000#) As Variant
Dim row As Long, curSheet As String
'retrieves te highest Y value from a Range with X and Y values
'XcolIdx is the index number of the column within the range in which the X values can be found
'YColIdx is the index number of the column within the range in which the Y values can be found
'fromX and toX are optional and can be used to restrict the search to the part of the range where X falls between these values
curSheet = ActiveWorkbook.ActiveSheet.Name

If myXRange.Rows.Count <> myYRange.Rows.Count Then
  MsgBox ("Error in function MaxYFromXYRange. Ranges must be of equal length.")
  Exit Function
ElseIf myXRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MaxYFromXYRange. Range containing X values must consist of only one column.")
  Exit Function
ElseIf myYRange.Columns.Count <> 1 Then
  MsgBox ("Error in function MaxYFromXYRange. Range containing Y values must consist of only one column.")
  Exit Function
End If

MaxYFromXYRange = -10000000000000#
Worksheets(myWorksheet).Activate
For row = 1 To myXRange.Rows.Count
  If IsNumeric(myXRange.Cells(row, 1)) And IsNumeric(myYRange.Cells(row, 1)) Then
    If myYRange.Cells(row, 1) > MaxYFromXYRange And myXRange.Cells(row, 1) >= fromX And myXRange.Cells(row, 1) <= toX Then
      MaxYFromXYRange = myYRange.Cells(row, 1)
    End If
  Else
    'MsgBox ("Error in function MaxYFromXYRange: non numeric value encountered in row index " & row & " of the data range.")
    'Exit Function
  End If
Next row

End Function

Public Function CONCATENATEALGEBRAIC(MyRange As Range, AlgebraString As String) As String
  Dim i As Long, result As String
  If MyRange.Columns.Count <> 1 Then
    MsgBox ("Error in function CONCATENATEALGEBRAIC. Range must consist of one column.")
  Else
   result = MyRange.Rows(1)
   For i = 2 To MyRange.Rows.Count
     result = result & " " & AlgebraString & " " & MyRange.Rows(i)
   Next
   CONCATENATEALGEBRAIC = result
  End If
End Function

Public Function CONCATENATEWITHDELIMITER(MyRange As Range, delimiter As String, Optional surroundingCharacter As String = "") As String
  Dim result As String
  Dim cellValue As String
  Dim r As Long, c As Long
  
  If delimiter = "\t" Then delimiter = vbTab
  
  For r = 1 To MyRange.Rows.Count
    For c = 1 To MyRange.Columns.Count
      ' Apply the surrounding character to the current cell's value
      cellValue = surroundingCharacter & MyRange.Cells(r, c).Text & surroundingCharacter
      
      ' Check if it's the first cell to avoid adding the delimiter before it
      If r = 1 And c = 1 Then
        result = cellValue
      Else
        result = result & delimiter & cellValue
      End If
    Next c
  Next r
  
  CONCATENATEWITHDELIMITER = result
End Function

Function ConcatenateFortranStyle(rng As Range, maxDecimals As Integer, totalPositions As Integer) As String
    Dim result As String
    Dim cell As Range
    Dim formattedNumber As String
    
    For Each cell In rng
        If IsNumeric(cell.value) Then
            ' Convert to scientific notation
            formattedNumber = Format(cell.value, "0." & String(maxDecimals, "#") & "E+00")
            
            ' Adjust the length
            If Len(formattedNumber) > totalPositions Then
                ' Truncate if longer
                formattedNumber = Left(formattedNumber, totalPositions)
            ElseIf Len(formattedNumber) < totalPositions Then
                ' Add trailing spaces if shorter
                formattedNumber = formattedNumber & Space(totalPositions - Len(formattedNumber))
            End If
        Else
            ' Handle non-numeric values
            formattedNumber = Left(cell.value & Space(totalPositions), totalPositions)
        End If
        
        ' Concatenate to the result
        result = result & formattedNumber
    Next cell
    
    ConcatenateFortranStyle = result
End Function

Public Sub AddWorkSheet(sheetName As String)

  If WorkSheetExists(sheetName) Then
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    Worksheets.Add
    ActiveSheet.Name = sheetName
  Else
    Worksheets.Add
    ActiveSheet.Name = sheetName
  End If

End Sub

Public Function FindColumnOnWorkSheet(sheetName As String, Header As String, row As Long, Optional GiveWarning As Boolean) As Long
Dim col As Long

FindColumnOnWorkSheet = 0
For col = 1 To 100
  If VBA.LCase(Worksheets(sheetName).Cells(row, col)) = VBA.LCase(Header) Then
    FindColumnOnWorkSheet = col
    Exit For
  End If
Next col

If FindColumnOnWorkSheet = 0 And GiveWarning Then
  MsgBox ("Column " & Header & " not found.")
End If

End Function

Public Function UnpivotMultiHeader(ByRef MyRange As Range, nHeaderRows As Integer, nHeaderCols As Integer, ResultsStartRow As Integer, ResultsStartCol As Integer) As Boolean
    'this routine converts a given table including multiple headers into a pivot-ready table
    Dim r As Long, c As Long
    Dim row As Integer, col As Integer
    Dim RowHeaderIdx As Integer
    Dim ColHeaderIdx As Integer
    
    row = ResultsStartRow
    col = ResultsStartCol
    For r = (nHeaderRows + 1) To MyRange.Rows.Count
        For c = nHeaderCols + 1 To MyRange.Columns.Count
            col = ResultsStartCol
            For RowHeaderIdx = 1 To nHeaderRows
                ActiveSheet.Cells(row, col) = MyRange.Cells(RowHeaderIdx, c)
                col = col + 1
            Next
            For ColHeaderIdx = 1 To nHeaderCols
                ActiveSheet.Cells(row, col) = MyRange.Cells(r, ColHeaderIdx)
                col = col + 1
            Next
            ActiveSheet.Cells(row, col) = MyRange.Cells(r, c)
            row = row + 1
        Next
    Next

End Function

Public Function UnpivotTable(ByRef MyRange As Range, ColumnHeaderVariableName As String, RowHeaderVariableName As String, ValueHeaderName As String) As Boolean
    'a very simple unpivot
    Dim curSheet As Worksheet, newSheet As Worksheet
    Dim CurSheetName As String
    Dim NewSheetName As String
    CurSheetName = ActiveSheet.Name
    NewSheetName = CurSheetName & ".UnPivot"
    Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
    
    'create a new worksheet for the results
    If Not WorkSheetExists(NewSheetName) Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = NewSheetName
        Set newSheet = ActiveWorkbook.Sheets(NewSheetName)
        UnpivotTable = True
    Else
        MsgBox ("Worksheet " & NewSheetName & " already exists. Please remove the old one first.")
        UnpivotTable = False
    End If
    
    Dim r2 As Integer
    r2 = 1
    newSheet.Cells(r2, 1) = ColumnHeaderVariableName
    newSheet.Cells(r2, 2) = RowHeaderVariableName
    newSheet.Cells(r2, 3) = ValueHeaderName
    

     
    Dim r As Integer, c As Integer
    For r = 2 To MyRange.Rows.Count
        For c = 2 To MyRange.Columns.Count
            r2 = r2 + 1
            newSheet.Cells(r2, 1) = MyRange.Cells(1, c)
            newSheet.Cells(r2, 2) = MyRange.Cells(r, 1)
            newSheet.Cells(r2, 3) = MyRange.Cells(r, c)
        Next
    Next
    
    
End Function

Public Function UnPivot2(ByRef MyRange As Range, ValuesStartCol As Integer, ValuesEndCol As Integer, ResultsStartRow As Integer, ResultsStartCol As Integer)
    Dim r As Integer, c As Integer
    Dim Header1 As String, Header2 As String, value As Variant
    Dim resRow As Integer, rescol As Integer
    
    resRow = ResultsStartRow
    rescol = ResultsStartCol
    ActiveSheet.Cells(resRow, rescol) = "Header1"
    ActiveSheet.Cells(resRow, rescol + 1) = "Header2"
    ActiveSheet.Cells(resRow, rescol + 2) = "Value"
            
    For r = 2 To MyRange.Rows.Count
        Header1 = MyRange.Cells(r, 1)
        For c = ValuesStartCol To ValuesEndCol
            Header2 = MyRange.Cells(1, c)
            value = MyRange.Cells(r, c)
            resRow = resRow + 1
            ActiveSheet.Cells(resRow, rescol) = Header1
            ActiveSheet.Cells(resRow, rescol + 1) = Header2
            ActiveSheet.Cells(resRow, rescol + 2) = value
        Next
    Next

End Function


Public Function UnPivot(ByRef MyRange As Range, HeaderColNum As Integer, UnpivotColNum As Integer, ValueColNum As Integer) As Boolean
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way
  'Note: as input a range with exactly three (3) columns is required! One for the row headers in the result, one for the new columns and one containing the values
  Dim r As Long, c As Long, r2 As Long, c2 As Long
  Dim nRow As Integer, nCol As Integer
  nRow = MyRange.Rows.Count
  nCol = MyRange.Columns.Count
  Dim myArray() As Variant
  Dim CurSheetName As String
  Dim NewSheetName As String
  Dim PivotList As Collection
  Set PivotList = New Collection
  Dim HeaderList As Collection
  Set HeaderList = New Collection
      
  Dim curSheet As Worksheet, newSheet As Worksheet
  CurSheetName = ActiveSheet.Name
  NewSheetName = CurSheetName & ".UnPivot"
  Set curSheet = ActiveWorkbook.Sheets(CurSheetName)
    
  'since we've been given a column that needs unpivoting, we'll first create a collection of all unique elements inside that column
  For r = 2 To MyRange.Rows.Count
    If CollectionGetIndex(PivotList, MyRange.Cells(r, UnpivotColNum)) = 0 Then Call PivotList.Add(MyRange.Cells(r, UnpivotColNum))
  Next
  
  'figure out the range at which the header value is unique. After that it starts repeating itself, so quit
  For r = 2 To MyRange.Rows.Count
    If CollectionGetIndex(HeaderList, MyRange.Cells(r, HeaderColNum)) = 0 Then
        Call HeaderList.Add(MyRange.Cells(r, HeaderColNum))
    Else
        Exit For
    End If
  Next
  
  'redim the results array
  ReDim myArray(HeaderList.Count + 1, PivotList.Count + 1)
  
  'write the column headers
  For c = 1 To PivotList.Count
    myArray(1, c + 1) = PivotList(c)
  Next

  'write the row headers
  For r = 1 To HeaderList.Count
    myArray(r + 1, 1) = HeaderList(r)
  Next

  'write the block
  Dim Header As Variant
  Dim PivotHeader As Variant
  Dim myValue As Variant
  Dim myR As Integer, myC As Integer
  
  For r = 2 To MyRange.Rows.Count
    Header = MyRange.Cells(r, HeaderColNum)
    PivotHeader = MyRange.Cells(r, UnpivotColNum)
    myValue = MyRange.Cells(r, ValueColNum)
      
    'now find the row and column number
    myR = CollectionGetIndex(HeaderList, Header) + 1
    myC = CollectionGetIndex(PivotList, PivotHeader) + 1
    myArray(myR, myC) = myValue
  Next
    
  If Not WorkSheetExists(NewSheetName) Then
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = NewSheetName
    Set newSheet = ActiveWorkbook.Sheets(NewSheetName)
    newSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray
    UnPivot = True
  Else
    MsgBox ("Worksheet " & NewSheetName & " already exists. Please remove the old one first.")
    UnPivot = False
  End If

  
End Function

Public Function CollectionContains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    CollectionContains = True
    obj = col(key)
    Exit Function
err:
    CollectionContains = False
End Function

Public Function CollectionGetIndex(ByRef col As Collection, key As Variant) As Integer
    Dim idx As Integer
    For idx = 1 To col.Count
        If col(idx) = key Then
            CollectionGetIndex = idx
            Exit Function
        End If
    Next
    CollectionGetIndex = 0
End Function


Public Sub UnPivot2CSV(ByRef MyRange As Range, StartDataCol As Integer, ResultsFile As String, delimiter As String, DataColName As String)
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way to csv
  Dim r As Long, c As Long, fn As Long
  Dim BaseStr As String, DataStr As String, myStr As String
  
  fn = FreeFile
  Open ResultsFile For Output As #fn
  
  'WRITE THE HEADER
  myStr = MyRange.Cells(1, 1)
  If StartDataCol > 2 Then
    For c = 2 To StartDataCol - 1
      myStr = myStr & "," & MyRange.Cells(1, c)
    Next
  End If
  myStr = myStr & "," & DataColName
  Print #fn, myStr
  
  'WRITE THE DATA
  For r = 2 To MyRange.Rows.Count
    myStr = MyRange.Cells(r, 1)
    If StartDataCol > 2 Then
      For c = 2 To StartDataCol - 1
        BaseStr = BaseStr & delimiter & MyRange.Cells(r, c)
      Next
    End If
    
    For c = StartDataCol To MyRange.Columns.Count
      If MyRange.Cells(r, c) <> "" Then
        DataStr = MyRange.Cells(1, c)
        Print #fn, BaseStr & delimiter & DataStr
      End If
    Next
  Next
  
  Close (fn)
  
End Sub

Public Sub Timeseries2CSV(ByRef DataRange As Range, ResultsFile As String, DateCol As Integer, ValCol As Integer, delimiter As String, Append As Boolean, DateHeader As String, ValueHeader As String, DateFormat As String, ExcludeValue As Double)

  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way to csv
  Dim r As Long, c As Long, fn As Long, tmpStr As String
  Dim myDateStr As String
  Dim myValue As Double
  
  fn = FreeFile
  If Append Then
      Open ResultsFile For Append As #fn
  Else
      Open ResultsFile For Output As #fn
  End If
  
  'first write the header
    Print #fn, DateHeader & delimiter & ValueHeader
    
  'next write the data
  For r = 1 To DataRange.Rows.Count
    DoEvents
    myDateStr = DataRange.Cells(r, DateCol)
    myValue = DataRange.Cells(r, ValCol)
    If myValue <> ExcludeValue Then
        tmpStr = Format(myDateStr, DateFormat) & delimiter & myValue
        Print #fn, tmpStr
    End If
  Next
  Close (fn)
  
End Sub

Public Sub Range2CSV(ByRef MyRange As Range, ResultsFile As String, delimiter As String, Append As Boolean, WriteHeader As Boolean)
  'This routine creates a new worksheet and writes data from the original sheet in an unpivoted way to csv
  Dim r As Long, c As Long, fn As Long, tmpStr As String
  
  fn = FreeFile
  If Append Then
      Open ResultsFile For Append As #fn
  Else
      Open ResultsFile For Output As #fn
  End If
  
  'first write the header
  If WriteHeader Then
    tmpStr = MyRange.Cells(1, 1)
    For c = 2 To MyRange.Columns.Count
      tmpStr = tmpStr & delimiter & MyRange.Cells(1, c)
    Next
    Print #fn, tmpStr
  End If
  
  'next write the data
  For r = 2 To MyRange.Rows.Count
    DoEvents
    tmpStr = MyRange.Cells(r, 1)
    For c = 2 To MyRange.Columns.Count
      tmpStr = tmpStr & delimiter & MyRange.Cells(r, c)
    Next
    Print #fn, tmpStr
  Next
  Close (fn)
  
End Sub

Function TimeseriesFromCSV(filePath As String, delimiter As String, headerRowIdx As Long, dateColName As String, valueColName As String, worksheetName As String, startRow As Long) As Long
    'returns the last row in the results worksheet data has been written to
    Dim i As Long

    Dim objFSO As Object
    Dim objFile As Object
    Dim objTS As Object
    
    Dim csvData As String
    Dim csvRows() As String
    Dim csvCols() As String
    
    Dim dateColIndex As Long
    Dim valueColIndex As Long
    
    Dim dateArray() As Variant
    Dim valueArray() As Variant
    
    ' Open the CSV file and read its contents
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(filePath)
    Set objTS = objFile.OpenAsTextStream(1, -2)
    csvData = objTS.ReadAll
    objTS.Close
    
    ' Split the CSV data into rows and columns
    csvRows = Split(csvData, vbCrLf)
    csvCols = Split(csvRows(headerRowIdx), delimiter)
    
    ' Find the column indexes for the date and value columns
    For i = 0 To UBound(csvCols)
        If csvCols(i) = dateColName Then
            dateColIndex = i
        ElseIf csvCols(i) = valueColName Then
            valueColIndex = i
        End If
    Next i
    
    ' Populate the date and value arrays
    ReDim dateArray(UBound(csvRows) - headerRowIdx)
    ReDim valueArray(UBound(csvRows) - headerRowIdx)
    For i = headerRowIdx + 1 To UBound(csvRows)
        csvCols = Split(csvRows(i), delimiter)
        
        If UBound(csvCols) >= valueColIndex And UBound(csvCols) >= dateColIndex Then
            ' Check if the cell value in the date column can be converted to a date
            If IsDate(csvCols(dateColIndex)) Then
                dateArray(i - headerRowIdx - 1) = CDate(csvCols(dateColIndex))
            Else
                dateArray(i - headerRowIdx - 1) = ""
            End If
    
            valueArray(i - headerRowIdx - 1) = CDbl(csvCols(valueColIndex))
        End If
    
    Next i
    

    ' Write the date and value arrays to the specified worksheet row by row
    With Worksheets(worksheetName)
            
        If startRow = 1 Then startRow = 2   'make sure we won't write values to the header row
        .Cells(1, 1).value = "Date"         'write the header
        .Cells(1, 2).value = "Value"        'write the header
        
        'loop through our data arrays and write them to the results sheet
        For i = 0 To UBound(dateArray)
            .Cells(startRow + i, 1).value = dateArray(i)
            .Cells(startRow + i, 2).value = valueArray(i)
        Next i
        
    End With
    
    'return the last row number we have written data to
    TimeseriesFromCSV = startRow + UBound(dateArray) + 1
    
End Function


Public Sub GoalSeekMultiple(ByRef GoalCell As Range, myGoal As Variant, ByRef MyRange As Range)
  'this function attempts to optimize a cell by adjusting values in multiple cells
  'it is a fairly simple approach, so it won't always work!!!
  'the routine optimizes by adjusting only one cell at a time
  Dim r As Long, c As Long
  For r = 1 To MyRange.Rows.Count
    For c = 1 To MyRange.Columns.Count
      GoalSeekMultiple = GoalCell.GoalSeek(myGoal, MyRange.Cells(r, c))
    Next
  Next

End Sub


Public Sub GoalSeekTriple(ByRef GoalCell As Range, myGoal As Variant, Adjust As Range, l1 As Variant, u1 As Variant, l2 As Variant, u2 As Variant, l3 As Variant, u3 As Variant, nIterations As Integer)
  Dim r As Long, c As Long, i As Integer, j As Long, k As Long, nIter As Integer
  Dim minI As Integer, minJ As Integer, minK As Integer
  Dim Range1 As Variant, Range2 As Variant, range3 As Variant
  Dim rowIdx As Integer, colIdx As Integer
  Dim myErr As Variant, minErr As Variant
  
  If Adjust.Count <> 3 Then
    MsgBox ("Error: range must contain 3 cells.")
    End
  End If
  
  Range1 = u1 - l1
  Range2 = u2 - l2
  range3 = u3 - l3
  
  Dim results(10, 10, 10) As Variant
  
  For nIter = 1 To nIterations
  
    For i = 1 To 10
      Adjust.Cells(1, 1) = l1 + (i - 0.5) * (u1 - l1) / 10
      For j = 1 To 10
      
        If Adjust.Rows.Count > 1 Then
          Adjust.Cells(2, 1) = l2 + (j - 0.5) * (u2 - l2) / 10
        ElseIf Adjust.Columns.Count > 1 Then
          Adjust.Cells(1, 2) = l2 + (j - 0.5) * (u2 - l2) / 10
        End If
      
        For k = 1 To 10
          If Adjust.Rows.Count > 1 Then
            Adjust.Cells(3, 1) = l3 + (k - 0.5) * (u3 - l3) / 10
          ElseIf Adjust.Columns.Count > 1 Then
            Adjust.Cells(1, 3) = l3 + (k - 0.5) * (u3 - l3) / 10
          End If
        
          'set the values for the 10x10x10 matrix
          If IsNumeric(GoalCell.value) Then
            results(i, j, k) = GoalCell.value
          Else
            results(i, j, k) = 99999999
          End If
        Next
      Next
    Next
    
    'find the value that's closest to the target
    minErr = 99999999
    For i = 1 To 10
      For j = 1 To 10
        For k = 1 To 10
           myErr = Math.Abs(results(i, j, k) - myGoal)
           If myErr < minErr Then
             minI = i
             minJ = j
             minK = k
             minErr = myErr
           End If
        Next
      Next
    Next
    
    'set the final value
    If Adjust.Rows.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(2, 1) = l2 + (minJ - 0.5) * (u2 - l2) / 10
      Adjust.Cells(3, 1) = l3 + (minK - 0.5) * (u3 - l3) / 10
    ElseIf Adjust.Columns.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(1, 2) = l2 + (minJ - 0.5) * (u2 - l2) / 10
      Adjust.Cells(1, 3) = l3 + (minK - 0.5) * (u3 - l3) / 10
    End If
    
    'adjust the boundaries to initiate the next iteration
    l1 = l1 + (minI - 1) * Range1 / 10
    u1 = l1 + Range1 / 10
    l2 = l2 + (minJ - 1) * Range2 / 10
    u2 = l2 + Range2 / 10
    l3 = l3 + (minK - 1) * range3 / 10
    u3 = l3 + range3 / 10
    Range1 = u1 - l1
    Range2 = u2 - l2
    range3 = u3 - l3
  
  Next
  

  

End Sub

Public Sub GoalSeekDouble(ByRef GoalCell As Range, myGoal As Variant, Adjust As Range, l1 As Variant, u1 As Variant, l2 As Variant, u2 As Variant, nIterations As Integer)
  Dim r As Long, c As Long, i As Integer, j As Long, nIter As Integer
  Dim minI As Integer, minJ As Integer
  Dim Range1 As Variant, Range2 As Variant
  Dim rowIdx As Integer, colIdx As Integer
  Dim myErr As Variant, minErr As Variant
  
  Range1 = u1 - l1
  Range2 = u2 - l2
  
  If Adjust.Count <> 2 Then
    MsgBox ("Error: range must contain 2 cells.")
    End
  End If

  
  Dim results(10, 10) As Variant
  
  For nIter = 1 To nIterations
  
    For i = 1 To 10
      Adjust.Cells(1, 1) = l1 + (i - 0.5) * (u1 - l1) / 10
      For j = 1 To 10
        If Adjust.Rows.Count > 1 Then
          Adjust.Cells(2, 1) = l2 + (j - 0.5) * (u2 - l2) / 10
        ElseIf Adjust.Columns.Count > 1 Then
          Adjust.Cells(1, 2) = l2 + (j - 0.5) * (u2 - l2) / 10
        End If
      
        'set the values for the 10x10 matrix
        If IsNumeric(GoalCell.value) Then
          results(i, j) = GoalCell.value
        Else
          results(i, j) = 99999999
        End If
      Next
    Next
    
    'find the value that's closest to the target
    minErr = 99999999
    For i = 1 To 10
      For j = 1 To 10
        myErr = Math.Abs(results(i, j) - myGoal)
        If myErr < minErr Then
          minI = i
          minJ = j
          minErr = myErr
        End If
      Next
    Next
    
    'set the final value
    If Adjust.Rows.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(2, 1) = l2 + (minJ - 0.5) * (u2 - l2) / 10
    ElseIf Adjust.Columns.Count > 1 Then
      Adjust.Cells(1, 1) = l1 + (minI - 0.5) * (u1 - l1) / 10
      Adjust.Cells(1, 2) = l2 + (minJ - 0.5) * (u2 - l2) / 10
    End If
    
    'adjust the boundaries to initiate the next iteration
    l1 = l1 + (minI - 1) * Range1 / 10
    u1 = l1 + Range1 / 10
    l2 = l2 + (minJ - 1) * Range2 / 10
    u2 = l2 + Range2 / 10
    Range1 = u1 - l1
    Range2 = u2 - l2
  Next

End Sub

Public Function COLUMN_NUMBER(ByVal myVal As Variant, ByVal MyRange As Range) As Long
  Dim c As Long
  For c = 1 To MyRange.Columns.Count
    If MyRange.Cells(1, c) = myVal Then
      COLUMN_NUMBER = c
      Exit Function
    End If
  Next
  COLUMN_NUMBER = 0
End Function

Public Sub PrintArray(ByRef data As Variant, ByRef Cl As Range)
    Cl.Resize(UBound(data, 1), UBound(data, 2)) = data
End Sub

Public Function RANGEVERTASCENDING(ByRef MyRange As Range, Optional AllowEqualValues As Boolean = True) As Boolean
  Dim i As Long
  
  If AllowEqualValues Then
    For i = 1 To MyRange.Rows.Count - 1
      If MyRange(i, 1) > MyRange(i + 1, 1) Then
        RANGEVERTASCENDING = False
        Exit Function
      End If
    Next
  Else
    For i = 1 To MyRange.Rows.Count - 1
      If MyRange(i, 1) >= MyRange(i + 1, 1) Then
        RANGEVERTASCENDING = False
        Exit Function
      End If
    Next
  End If
  
  RANGEVERTASCENDING = True
  
End Function

Public Function VALUEFROMCELLADDRESS(row As Integer, col As Integer, sheetName As String) As Variant
    'deze functie geeft een celwaarde terug op basis van het adres (rij, kolom)
    'merk op dat het ook direct als werkbladfunctie kan: INDIRECT(ADDRESS(RIJ, KOLOM,,,WERKBLADNAAM))
    VALUEFROMCELLADDRESS = ActiveWorkbook.Sheets(sheetName).Cells(row, col)
End Function

Public Function FormatRoman(ByVal n As Integer) As String
   ' Author: Christian d'Heureuse (www.source-code.biz)
   If n = 0 Then VBA.FormatRoman = "0": Exit Function
      ' There is no roman symbol for 0, but we don't want to return an empty string.
   Const r = "IVXLCDM"              ' roman symbols
   Dim i As Integer: i = Abs(n)
   Dim s As String, P As Integer
   For P = 1 To 5 Step 2
      Dim d As Integer: d = i Mod 10: i = i \ 10
      Select Case d                 ' VBA.Format a decimal digit
         Case 0 To 3: s = String(d, VBA.Mid(r, P, 1)) & s
         Case 4:      s = VBA.Mid(r, P, 2) & s
         Case 5 To 8: s = VBA.Mid(r, P + 1, 1) & String(d - 5, VBA.Mid(r, P, 1)) & s
         Case 9:      s = VBA.Mid(r, P, 1) & VBA.Mid(r, P + 2, 1) & s
         End Select
      Next
   s = String(i, "M") & s           ' VBA.Format thousands
   If n < 0 Then s = "-" & s        ' insert sign if negative (non-standard)
   VBA.FormatRoman = s

End Function

Public Function LSHA2MMPD(myVal As Variant) As Variant
  Dim newVal As Variant
  newVal = myVal * 3600 * 24 / 10000
  LSHA2MMPD = newVal
End Function

Public Function MMPD2LSHA(myVal As Variant) As Variant
  Dim newVal As Variant
  newVal = myVal / 3600 / 24 * 10000
  MMPD2LSHA = newVal
End Function

Public Function M3PS2MMPD(CAP As Variant, Opp As Variant) As Variant
  'cap in m3/s
  'opp in m2
  M3PS2MMPD = CAP / Opp * 1000 * 3600 * 24
End Function

Public Function M3PS2MMPU(CAP As Variant, Opp As Variant) As Variant
  'cap in m3/s
  'opp in m2
  If Opp > 0 Then
    M3PS2MMPU = CAP / Opp * 1000 * 3600
  Else
    M3PS2MMPU = 0
  End If
End Function

Public Function MMPU2M3PS(CAP As Variant, Opp As Variant) As Variant
  'cap in mm/u
  'opp in m2
  MMPU2M3PS = CAP / 3600 / 1000 * Opp
End Function

Public Function MMPD2M3PS(CAP As Variant, Opp As Variant) As Variant
  'cap in mm/d
  'opp in m2
  MMPD2M3PS = CAP / 1000 / 24 / 3600 * Opp
End Function


Public Function Celcius2Kelvin(Celcius As Variant)
  Celcius2Kelvin = Celcius + 273.15
End Function
Public Function Kelvin2Celcius(Kelvin As Variant)
  Kelvin2Celcius = Kelvin - 273.15
End Function

Public Function RD2LATLONG(x As Variant, y As Variant, Optional ByRef Latitude As Variant = 0, Optional ByRef Longitude As Variant = 0) As String
  Dim dX As Variant, dY As Variant
  Dim SomN As Variant, SomE As Variant

  dX = (x - 155000) * 10 ^ (-5)
  dY = (y - 463000) * 10 ^ (-5)
  SomN = (3235.65389 * dY) + (-32.58297 * dX ^ 2) + (-0.2475 * dY ^ 2) + (-0.84978 * dX ^ 2 * dY) + (-0.0655 * dY ^ 3) + (-0.01709 * dX ^ 2 * dY ^ 2) + (-0.00738 * dX) + (0.0053 * dX ^ 4) + (-0.00039 * dX ^ 2 * dY ^ 3) + (0.00033 * dX ^ 4 * dY) + (-0.00012 * dX * dY)
  SomE = (5260.52916 * dX) + (105.94684 * dX * dY) + (2.45656 * dX * dY ^ 2) + (-0.81885 * dX ^ 3) + (0.05594 * dX * dY ^ 3) + (-0.05607 * dX ^ 3 * dY) + (0.01199 * dY) + (-0.00256 * dX ^ 3 * dY ^ 2) + (0.00128 * dX * dY ^ 4) + (0.00022 * dY ^ 2) + (-0.00022 * dX ^ 2) + (0.00026 * dX ^ 5)
  Latitude = 52.15517 + (SomN / 3600)
  Longitude = 5.387206 + (SomE / 3600)
 
  RD2LATLONG = Latitude & ";" & Longitude

End Function

Public Function RD2LAT(x As Variant, y As Variant) As Variant
  Dim Latitude As Variant, Longitude As Variant
  Call RD2LATLONG(x, y, Latitude, Longitude)
  RD2LAT = Latitude
End Function
Public Function RD2LON(x As Variant, y As Variant) As Variant
  Dim Latitude As Variant, Longitude As Variant
  Call RD2LATLONG(x, y, Latitude, Longitude)
  RD2LON = Longitude
End Function

Public Function RD2WGS84(x As Variant, y As Variant, Optional ByRef lat As Variant = 0, Optional ByRef lon As Variant = 0) As String
  'converteert RD-coordinaten naar Lat/Long (WGS84)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim phi As Variant
  Dim lambda As Variant
  Call RD2BESSEL(x, y, phi, lambda)
  Call BESSEL2WGS84(phi, lambda, lat, lon)
  RD2WGS84 = lat & "," & lon
End Function

Public Function WGS842RD(lat As Variant, lon As Variant, Optional ByRef x As Variant = 0, Optional ByRef y As Variant = 0) As String
  'converteert WGS84-coordinaten (Lat/Long) naar RD
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim phiBes As Variant
  Dim LambdaBes As Variant
  Call WGS842BESSEL(lat, lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842RD = x & "," & y
  
End Function

Public Function WGS842X(lat As Variant, lon As Variant) As Variant
  'converteert WGS84-coordinaten (Lat/Long) naar RD (alleen de X-coordinaat)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim x As Variant, y As Variant
  Dim phiBes As Variant
  Dim LambdaBes As Variant
  Call WGS842BESSEL(lat, lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842X = x
  
End Function

Public Function WGS842Y(lat As Variant, lon As Variant) As Variant
  'converteert WGS84-coordinaten (Lat/Long) naar RD (alleen de Y-coordinaat)
  'maakt gebruik van de routines van Ejo Schrama: schrama @geo.tudelft.nl
  Dim x As Variant, y As Variant
  Dim phiBes As Variant
  Dim LambdaBes As Variant
  Call WGS842BESSEL(lat, lon, phiBes, LambdaBes)
  Call BESSEL2RD(phiBes, LambdaBes, x, y)
  WGS842Y = y
  
End Function

Public Function WGS84DEG2DECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to decimal
  Dim tmpStr As String, Pos As Integer, startPos As Integer
  Dim DegNB As Variant, MinNB As Variant, SecNB As Variant, Northing As Variant
  Dim DegOL As Variant, MinOL As Variant, SecOL As Variant, Easting As Variant
  Deg = VBA.Trim(Deg)
  
  'find out where the actual value for northing begins and clean the string up
  For Pos = 1 To Len(Deg)
   If IsNumeric(VBA.Mid(Deg, Pos, 1)) Then
     startPos = Pos
     Exit For
   End If
  Next
  Deg = VBA.Right(Deg, Len(Deg) - startPos + 1)
  
  'determine the coordinate for northing
  DegNB = VAL(ParseString(Deg, "°", 0))
  MinNB = VAL(ParseString(Deg, "'", 0)) / 60
  SecNB = VAL(ParseString(Deg, Chr(34), 0)) / 3600
  Northing = DegNB + MinNB + SecNB
  
  'find out where the actual value for easting begins
  For Pos = 1 To Len(Deg)
   If IsNumeric(VBA.Mid(Deg, Pos, 1)) Then
     startPos = Pos
     Exit For
   End If
  Next
  Deg = VBA.Right(Deg, Len(Deg) - startPos + 1)
  
  'retrieve the coordinates for Easting
  DegOL = VAL(ParseString(Deg, "°", 0))
  MinOL = VAL(ParseString(Deg, "'", 0)) / 60
  SecOL = VAL(ParseString(Deg, Chr(34), 0)) / 3600
  Easting = DegOL + MinOL + SecOL
    
  WGS84DEG2DECIMAL = Northing & "," & Easting

End Function

Public Function WGS84DEG2LATDECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to Latitude in Decimals
  Dim Decimals As String
  Decimals = WGS84DEG2DECIMAL(Deg)
  WGS84DEG2LATDECIMAL = ParseString(Decimals, ",")
End Function

Public Function WGS84DEG2LONDECIMAL(Deg As String) As String
  'converts WGS84 coordinates from degrees to Latitude in Decimals
  Dim Decimals As String, tmpStr As String
  Decimals = WGS84DEG2DECIMAL(Deg)
  tmpStr = ParseString(Decimals, ",")
  WGS84DEG2LONDECIMAL = Decimals
End Function

Public Sub RD2BESSEL(x As Variant, y As Variant, ByRef phi As Variant, ByRef lambda As Variant)

'converteert RD-coordinaten naar phi en lambda voor een Bessel-functie
'code is geheel gebaseerd op de routines van Ejo Schrama's software:
'schrama@geo.tudelft.nl

Dim x0 As Variant
Dim y0 As Variant
Dim k As Variant
Dim bigr As Variant
Dim m As Variant
Dim n As Variant
Dim lambda0 As Variant
Dim phi0 As Variant
Dim l0 As Variant
Dim b0 As Variant
Dim e As Variant
Dim a As Variant

Dim d_1 As Variant, d_2 As Variant, r As Variant, sa As Variant, ca As Variant, psi As Variant, cpsi As Variant, spsi As Variant
Dim sb As Variant, cb As Variant, b As Variant, sdl As Variant, dl As Variant, W As Variant, Q As Variant, phiprime As Variant
Dim dq As Variant, i As Long, pi As Variant

x0 = 155000
y0 = 463000
k = 0.9999079
bigr = 6382644.571
m = 0.003773953832
n = 1.00047585668

pi = Application.WorksheetFunction.pi
'pi = 3.14159265358979
lambda0 = pi * 2.99313271611111E-02
phi0 = pi * 0.289756447533333
l0 = pi * 2.99313271611111E-02
b0 = pi * 0.289561651383333

e = 0.08169683122
a = 6377397.155

d_1 = x - x0
d_2 = y - y0
r = Sqr(d_1 ^ 2 + d_2 ^ 2)

If r <> 0 Then
  sa = d_1 / r
  ca = d_2 / r
Else
  sa = 0
  ca = 0
End If

psi = Application.WorksheetFunction.ATan2(k * 2 * bigr, r) * 2
cpsi = Cos(psi)
spsi = Sin(psi)

sb = ca * Cos(b0) * spsi + Sin(b0) * cpsi
d_1 = sb
cb = Sqr(1 - d_1 ^ 2)
b = Application.WorksheetFunction.Acos(cb)
sdl = sa * spsi / cb
dl = Application.WorksheetFunction.Asin(sdl)
lambda = dl / n + lambda0
W = Application.WorksheetFunction.Ln(Tan(b / 2 + pi / 4))
Q = (W - m) / n

phi = Atn(Exp(1) ^ Q) * 2 - pi / 2 'phi prime
For i = 1 To 4
  dq = e / 2 * Application.WorksheetFunction.Ln((e * Sin(phi) + 1) / (1 - e * Sin(phi)))
  phi = Atn(Exp(1) ^ (Q + dq)) * 2 - pi / 2
Next

lambda = lambda / pi * 180
phi = phi / pi * 180

End Sub

Public Sub BESSEL2WGS84(phi As Variant, lambda As Variant, ByRef PhiWGS As Variant, ByRef LamWGS As Variant)
  Dim dphi As Variant, dlam As Variant, phicor As Variant, lamcor As Variant

  dphi = phi - 52
  dlam = lambda - 5
  phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
  lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
  PhiWGS = phi + phicor
  LamWGS = lambda + lamcor


End Sub

Public Sub WGS842BESSEL(PhiWGS As Variant, LamWGS As Variant, ByRef phi As Variant, ByRef lambda As Variant)
  Dim dphi As Variant, dlam As Variant, phicor As Variant, lamcor As Variant

  dphi = PhiWGS - 52
  dlam = LamWGS - 5
  phicor = (-96.862 - dphi * 11.714 - dlam * 0.125) * 0.00001
  lamcor = (dphi * 0.329 - 37.902 - dlam * 14.667) * 0.00001
  phi = PhiWGS - phicor
  lambda = LamWGS - lamcor
  
End Sub

Public Sub BESSEL2RD(phiBes As Variant, lamBes As Variant, ByRef x As Variant, ByRef y As Variant)

'converteert Lat/Long van een Bessel-functie naar X en Y in RD
'code is geheel gebaseerd op de routines van Ejo Schrama's software:
'schrama@geo.tudelft.nl

Dim x0 As Variant
Dim y0 As Variant
Dim k As Variant
Dim bigr As Variant
Dim m As Variant
Dim n As Variant
Dim lambda0 As Variant
Dim phi0 As Variant
Dim l0 As Variant
Dim b0 As Variant
Dim e As Variant
Dim a As Variant

Dim d_1 As Variant, d_2 As Variant, r As Variant, sa As Variant, ca As Variant, psi As Variant, cpsi As Variant, spsi As Variant
Dim sb As Variant, cb As Variant, b As Variant, sdl As Variant, dl As Variant, W As Variant, Q As Variant, phiprime As Variant
Dim dq As Variant, i As Long, pi As Variant, phi As Variant, lambda As Variant, s2psihalf As Variant, cpsihalf As Variant, spsihalf As Variant
Dim tpsihalf As Variant

x0 = 155000
y0 = 463000
k = 0.9999079
bigr = 6382644.571
m = 0.003773953832
n = 1.00047585668

pi = Application.WorksheetFunction.pi
'pi = 3.14159265358979
lambda0 = pi * 2.99313271611111E-02
phi0 = pi * 0.289756447533333
l0 = pi * 2.99313271611111E-02
b0 = pi * 0.289561651383333

e = 0.08169683122
a = 6377397.155

phi = phiBes / 180 * pi
lambda = lamBes / 180 * pi

Q = Application.WorksheetFunction.Ln(Tan(phi / 2 + pi / 4))
dq = e / 2 * Application.WorksheetFunction.Ln((e * Sin(phi) + 1) / (1 - e * Sin(phi)))
Q = Q - dq
W = n * Q + m
b = Atn(Exp(1) ^ W) * 2 - pi / 2
dl = n * (lambda - lambda0)
d_1 = Sin((b - b0) / 2)
d_2 = Sin(dl / 2)
s2psihalf = d_1 * d_1 + d_2 * d_2 * Cos(b) * Cos(b0)
cpsihalf = Sqr(1 - s2psihalf)
spsihalf = Sqr(s2psihalf)
tpsihalf = spsihalf / cpsihalf
spsi = spsihalf * 2 * cpsihalf
cpsi = 1 - s2psihalf * 2
sa = Sin(dl) * Cos(b) / spsi
ca = (Sin(b) - Sin(b0) * cpsi) / (Cos(b0) * spsi)
r = k * 2 * bigr * tpsihalf
x = Round(r * sa + x0, 0)
y = Round(r * ca + y0, 0)

End Sub

Sub ExtractCoordinatesFromWKT(wkt As String, ResultsSheet As String, startRow As Integer, StartCol As Integer)

    Dim coordinates As String
    coordinates = ""
    
    Dim r As Long
    r = startRow
    
    'Remove unnecessary characters from WKT string
    wkt = Replace(wkt, "MULTIPOLYGON", "")
    wkt = Replace(wkt, "(", "")
    wkt = Replace(wkt, ")", "")
    wkt = Replace(wkt, ",", " ")
    
    'Split the WKT string into an array of coordinates
    Dim coords() As String
    coords = Split(Trim(wkt))
    
    'Loop through the array of coordinates and extract x and y values
    Dim i As Integer
    For i = 0 To UBound(coords) Step 2
        Dim x As String
        Dim y As String
        x = coords(i)
        y = coords(i + 1)
        
        r = r + 1
        Worksheets(ResultsSheet).Cells(r, StartCol) = x
        Worksheets(ResultsSheet).Cells(r, StartCol + 1) = y
        
        
        'coordinates = coordinates & x & "," & y & vbCrLf
    Next i
    
    'ExtractCoordinatesFromWKT = coordinates

End Sub


Public Function MultiParse(ByRef myString As String, returnInstanceNumber As Integer, Optional delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String
  Dim tmpString As String, i As Long
  For i = 1 To returnInstanceNumber
    tmpString = ParseString(myString, delimiter, QuoteHandlingFlag)
  Next
  MultiParse = tmpString
End Function

Public Function ParseNumeric(ByRef myString As String) As String
  Dim i As Integer, myChar As String, Done As Boolean
  'knabbelt net zo lang een karakter van de linker kant van een string af tot het resultaat niet langer numeriek is
  
  While Not Done
    myChar = VBA.Left(myString, 1)
    If Not (IsNumeric(myChar) Or myChar = ".") Then
      Exit Function
    Else
      ParseNumeric = ParseNumeric & myChar
      myString = Right(myString, VBA.Len(myString) - 1)
    End If
  Wend
  
End Function

Public Function ParseString(ByRef myString As String, Optional delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String

Dim Pos As Long, quoteEven As Boolean
quoteEven = True

'Quotehandlingflag: default = 1
'0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
'1 = items between single quotes are being treated as separate items (no parsing between single quotes)
'2 = items between double quotes are being treated as separate items (no parsing between double quotes)

Dim i As Long
For i = 1 To VBA.Len(myString)
  
  'als we een quote tegenkomen, houden we bij of het even of oneven is. zo weten we of we een omsloten object hebben
  If VBA.Left(myString, 1) = "'" And QuoteHandlingFlag = 1 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then myString = VBA.Right(myString, VBA.Len(myString) - 1)
      Exit Function
    End If
  
  ElseIf VBA.Left(myString, 1) = VBA.Chr(34) And QuoteHandlingFlag = 2 Then 'double quote encountered
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then myString = VBA.Right(myString, VBA.Len(myString) - 1)
      Exit Function
    End If
  'als het teken gelijk is aan de delimiter, kijken we of we al geldige tekens hadden gevonden
  'zo ja, wegschrijven
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 1 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 2 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 0 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseString) > 0 Then
      Exit Function
    End If
  Else
    'hier gebeurt het werkelijke parsen
    ParseString = ParseString & VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  End If
Next

End Function

Public Function ParseStringPlus(ByRef myString As String, ByRef QuotesFound As Boolean, Optional delimiter As String = " ", Optional QuoteHandlingFlag As Long = 1) As String

Dim Pos As Long, quoteEven As Boolean
quoteEven = True
QuotesFound = False

'Differences with ParseString:
'- Uses a byref boolean to return whether an item surrounded by quoted was found

'Quotehandlingflag: default = 1
'0 = items between quotes are NOT being treated as separate items (parsing also between quotes)
'1 = items between single quotes are being treated as separate items (no parsing between single quotes)
'2 = items between double quotes are being treated as separate items (no parsing between double quotes)

Dim i As Long
For i = 1 To VBA.Len(myString)
  
  'als we een quote tegenkomen, houden we bij of het even of oneven is. zo weten we of we een omsloten object hebben
  If VBA.Left(myString, 1) = "'" And QuoteHandlingFlag = 1 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then
        myString = VBA.Right(myString, VBA.Len(myString) - 1)
        QuotesFound = True
        Exit Function
      End If
    End If
  
  ElseIf VBA.Left(myString, 1) = VBA.Chr(34) And QuoteHandlingFlag = 2 Then 'double quote encountered
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If quoteEven Then
      quoteEven = False
      'we voegen niets toe aan de geparste string want quotes zelf doen niet mee
    Else
      quoteEven = True
      'we hebben een omsloten object gevonden dus is weer even
      If VBA.Len(myString) > 0 Then
        myString = VBA.Right(myString, VBA.Len(myString) - 1)
        QuotesFound = True
        Exit Function
      End If
    End If
  'als het teken gelijk is aan de delimiter, kijken we of we al geldige tekens hadden gevonden
  'zo ja, wegschrijven
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 1 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 2 And quoteEven = True Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = delimiter And QuoteHandlingFlag = 0 Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  ElseIf VBA.Left(myString, 1) = vbCrLf Then
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
    If VBA.Len(ParseStringPlus) > 0 Then
      Exit Function
    End If
  Else
    'hier gebeurt het werkelijke parsen
    ParseStringPlus = ParseStringPlus & VBA.Left(myString, 1)
    myString = VBA.Right(myString, VBA.Len(myString) - 1)
  End If
Next

End Function

Public Sub TextSnippet(startPos As Long, endPos As Long, myString As String, ByRef LeftSnippet As String, ByRef Snippet As String, ByRef RightSnippet As String)
  'cuts a string in three parts based on two string positions
  LeftSnippet = Left(myString, startPos - 1)
  Snippet = Mid(myString, startPos, endPos - startPos + 1)
  RightSnippet = VBA.Right(myString, Len(myString) - endPos)
End Sub

Public Function BNAString(id As String, Name As String, x As Variant, y As Variant) As String
  BNAString = VBA.Chr(34) & id & VBA.Chr(34) & "," & VBA.Chr(34) & Name & VBA.Chr(34) & ",1," & x & "," & y
End Function

Public Function WagModSTAString(myDate As Variant, Prec As Variant, EVAP As Variant, Qmeas As Variant) As String
  Dim TimeStr As String
  Dim myPrec As String, myEvap As String, myQm As String
  
  If hour(myDate) = 0 Then
    TimeStr = "0"
  Else
    TimeStr = VBA.Trim(hour(myDate) & "00")
  End If
  
  While VBA.Len(TimeStr) < 4
    TimeStr = " " & TimeStr
  Wend
  
  myPrec = VBA.Format(Prec, "0.000")
  While VBA.Len(myPrec) < 13
    myPrec = " " & myPrec
  Wend

  myEvap = VBA.Format(EVAP, "0.000")
  While VBA.Len(myEvap) < 8
    myEvap = " " & myEvap
  Wend
  
  myQm = VBA.Format(Qmeas, "0.000")
  While VBA.Len(myQm) < 8
    myQm = " " & myQm
  Wend

  WagModSTAString = year(myDate) & "/" & VBA.Format(month(myDate), "00") & "/" & VBA.Format(day(myDate), "00") & " " & TimeStr & " " & myPrec & " " & myEvap & " " & myQm
  
  
End Function

Public Function WalrusDATString(myDate As Variant, Prec As Variant, EVAP As Variant, Qmeas As Variant) As String
  Dim TimeStr As String
  Dim myPrec As String, myEvap As String, myQm As String

  WalrusDATString = year(myDate) & VBA.Format(month(myDate), "00") & VBA.Format(day(myDate), "00") & VBA.Format(hour(myDate), "00") & " " & VBA.Format(Prec, "0.0000") & " " & VBA.Format(EVAP, "0.0000") & " " & VBA.Format(Qmeas, "0.0000") & " 0 0 0 0"
  
End Function

'Binary Conversions
'The Functions in this module are designed to aid in working with BINARY
'numbers. Visual Basic does not include nor allow any representation of a
'number in binary VBA.Format.  Therefore, all of these functions work strictly on
'strings.  All of the parameters passed into them and returned from them are
'strings.
'
'              CONVERSION NEEDED                 FUNCTION
'            ------------------------------------------------------
'              Binary to Hex            BinToHex(BinNum As String)
'              Binary to Octal          BinToOct(BinNum As String)
'              Binary to Decimal        BinToDec(BinNum As String)
'              Hex to Binary            HexToBin(HexNum As String)
'              Octal to Binary          OctToBin(OctNum As String)
'              Decimal to Binary        DecToBin(DecNum As String)
'
'
Function BinToHex(BinNum As String) As String
   Dim BinLen As Integer, i As Integer
   Dim HexNum As Variant
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
   BinLen = VBA.Len(BinNum)
   For i = BinLen To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         HexNum = ""
         err.Raise 1002, "BinToHex", "Invalid Input"
      End If
'     Calculate HEX value of BinNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         HexNum = HexNum + 2 ^ Abs(i - BinLen)
      End If
   Next i
'  Return HexNum as String
   BinToHex = Hex(HexNum)
errorhandler:
End Function

Function BinToOct(BinNum As String) As String
   Dim BinLen As Integer, i As Integer
   Dim OctNum As Variant
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
   BinLen = VBA.Len(BinNum)
   For i = BinLen To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         OctNum = ""
         err.Raise 1002, "BinToOct", "Invalid Input"
      End If
'     Calculate Octal value of BinNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         OctNum = OctNum + 2 ^ Abs(i - BinLen)
      End If
   Next i
'  Return OctNum as String
   BinToOct = Oct(OctNum)
errorhandler:
End Function

Public Function BinToDec(BinNum As String) As String
   Dim i As Integer
   Dim DecNum As Long
   
   On Error GoTo errorhandler
   BinNum = VBA.Trim(BinNum)
'  Loop thru BinString
   For i = VBA.Len(BinNum) To 1 Step -1
'     Check the string for invalid characters
      If Asc(VBA.Mid(BinNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(BinNum, i, 1)) > 49 Then
         DecNum = ""
         err.Raise 1002, "BinToDec", "Invalid Input"
      End If
'     If bit is 1 then raise 2^LoopCount and add it to DecNum
      If VBA.Mid(BinNum, i, 1) And 1 Then
         DecNum = DecNum + 2 ^ (Len(BinNum) - i)
      End If
   Next i
'  Return DecNum as a String
   BinToDec = DecNum
errorhandler:
End Function

Public Function OctToBin(OctNum As String) As String
   Dim BinNum As String
   Dim lOctNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   OctNum = VBA.Trim(OctNum)
'  Check the string for invalid characters
   For i = 1 To VBA.Len(OctNum)
      If (Asc(VBA.Mid(OctNum, i, 1)) < 48 Or Asc(VBA.Mid(OctNum, i, 1)) > 55) Then
         BinNum = ""
         err.Raise 1008, "OctToBin", "Invalid Input"
      End If
   Next i

   i = 0
   lOctNum = VAL("&O" & OctNum)
   
   Do
      If lOctNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lOctNum
'  Return BinNum as a String
   OctToBin = BinNum
errorhandler:
End Function

Public Function DecToBin(DecNum As String) As String
   Dim BinNum As String
   Dim lDecNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   DecNum = VBA.Trim(DecNum)
   
'  Check the string for invalid characters
   For i = 1 To VBA.Len(DecNum)
      If Asc(VBA.Mid(DecNum, i, 1)) < 48 Or _
         Asc(VBA.Mid(DecNum, i, 1)) > 57 Then
         BinNum = ""
         err.Raise 1010, "DecToBin", "Invalid Input"
      End If
   Next i
   
   i = 0
   lDecNum = VAL(DecNum)
   
   Do
      If lDecNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lDecNum
'  Return BinNum as a String
   DecToBin = BinNum
errorhandler:
End Function

Public Function HexToBin(HexNum As String) As String
   Dim BinNum As String
   Dim lHexNum As Long
   Dim i As Integer
   
   On Error GoTo errorhandler
   HexNum = VBA.str(HexNum)
'  Check the string for invalid characters
   For i = 1 To VBA.Len(HexNum)
      If ((Asc(VBA.Mid(HexNum, i, 1)) < 48) Or _
          (Asc(VBA.Mid(HexNum, i, 1)) > 57 And _
           Asc(UCase(VBA.Mid(HexNum, i, 1))) < 65) Or _
          (Asc(UCase(VBA.Mid(HexNum, i, 1))) > 70)) Then
         BinNum = ""
         err.Raise 1016, "HexToBin", "Invalid Input"
      End If
   Next i
   
   i = 0
   lHexNum = VAL("&h" & HexNum)
   Do
      If lHexNum And 2 ^ i Then
         BinNum = "1" & BinNum
      Else
         BinNum = "0" & BinNum
      End If
      i = i + 1
   Loop Until 2 ^ i > lHexNum
'  Return BinNum as a String
   HexToBin = BinNum
errorhandler:
End Function

Public Function Ohm(Volt As Double, Ampere As Double) As Double
    'returns the required resistance (ohm) for a given combination of Voltage (V) and Current (Amperes)
    'Resistance = Volts / Amperes
    Ohm = Volt / Ampere
End Function

Public Function LedResistance(Vsource As Double, Vrequired As Double, i As Double) As Double
    'calculates the required resistance (ohm) for a given input voltage, requested voltage and requested current (Ampere)
    'R = (Vsource - Vrequired)/I
    'source: https://www.digikey.nl/nl/resources/conversion-calculators/conversion-calculator-led-series-resistor
    LedResistance = (Vsource - Vrequired) / i
End Function

Function processViForisProfiles()
    'this function processes cross sections as provided by ViForis in the sense that it changes values with postfix 'm' by the numerical value.
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim numericValue As Variant
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Set the range to columns B and C, starting from row 2
        Set rng = ws.Range("B2:C" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row)
        
        ' Loop through each cell in the range
        For Each cell In rng.Cells
            
            ' Check if the cell value ends with "m"
            If Right(cell.value, 1) = "m" Then
                
                ' Remove the "m" postfix and convert the cell value to a numeric value
                numericValue = Left(cell.value, Len(cell.value) - 1)
                
                ' Check if the remaining value is numeric, and if so, update the cell value
                If IsNumeric(numericValue) Then
                    cell.value = CDbl(numericValue)
                End If
                
            End If
            
        Next cell
        
    Next ws

End Function

Sub WriteYZProfileFiles(profileDatPath As String, profileDefPath As String, prefix As String, frictionDatPath As String, frictionValue As Double)

    'this function writes profile.dat and profile.def and friction.dat, based on YZ data in this workbook
    'assumptions:
    'ONE CROSS SECTION PER WORKSHEET
    'ID of the cross section = name of the worksheet
    'Y-values in column B
    'Z-values in column C

    Dim ws As Worksheet
    Dim id As String
    Dim profileDatFile As Integer
    Dim profileDefFile As Integer
    Dim frictionDatFile As Integer
    Dim i As Long
    Dim yzTable As String

    ' Open the files for writing
    profileDatFile = FreeFile()
    Open profileDatPath For Output As profileDatFile

    profileDefFile = FreeFile()
    Open profileDefPath For Output As profileDefFile

    frictionDatFile = FreeFile()
    Open frictionDatPath For Output As frictionDatFile

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        id = prefix & ws.Name
        
        ' Write profile.dat record
        Print #profileDatFile, "CRSN id '" & id & "' di '" & id & "' rl 0 rs 50.284 crsn"

        ' Prepare the YZ table data
        yzTable = "TBLE" & vbCrLf
        For i = 2 To ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            yzTable = yzTable & ws.Cells(i, "B").value & " " & ws.Cells(i, "C").value & " <" & vbCrLf
        Next i
        yzTable = yzTable & "tble"

        ' Write profile.def record
        Print #profileDefFile, "CRDS id '" & id & "' nm '" & id & "' ty 10 st 0 lt sw 0 0 lt yz"
        Print #profileDefFile, yzTable
        Print #profileDefFile, "gl 0"
        Print #profileDefFile, "gu 0"
        Print #profileDefFile, "crds"

        ' Write friction.dat record
        Print #frictionDatFile, "CRFR id '" & id & "' nm 'Friction' cs '" & id & "'"
        Print #frictionDatFile, "lt ys"
        Print #frictionDatFile, "TBLE"
        Print #frictionDatFile, "-12.22 13.16 <"
        Print #frictionDatFile, "tble"
        Print #frictionDatFile, "ft ys"
        Print #frictionDatFile, "TBLE"
        Print #frictionDatFile, "1 " & frictionValue & " <"
        Print #frictionDatFile, "tble"
        Print #frictionDatFile, "fr ys"
        Print #frictionDatFile, "TBLE"
        Print #frictionDatFile, "1 " & frictionValue & " <"
        Print #frictionDatFile, "tble crfr"

    Next ws

    ' Close the files
    Close profileDatFile
    Close profileDefFile
    Close frictionDatFile

End Sub

Function RoundDownDateToHour(InputDateTime As Date) As Date
    RoundDownDateToHour = DateAdd("h", hour(InputDateTime), DateValue(InputDateTime))
End Function


Option Explicit

Sub ReadMidasPrecipitation(FolderPath As String, sheetName As String)

    Dim fileName As String
    Dim ws As Worksheet
    Dim csvFile As String
    Dim textline As String
    Dim dataFlag As Boolean
    Dim record() As String
    Dim outputRow As Long
    Dim outputFile As String
    Dim fs As Object, txtFile As Object
    
    ' Ensure FolderPath ends with a backslash
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
        
    ' Specify the worksheet where you want to store data
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ws.Cells(1, 1).value = "Date"
    ws.Cells(1, 2).value = "Precipitation (mm)"
    
    ' Initialize the row where the first record will be written
    outputRow = 2

    ' Create a FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Get the first file
    fileName = dir(FolderPath & "*.csv")
    
    ' Loop through each CSV file in the folder
    While fileName <> ""
    
        ' Open the file
        Set txtFile = fs.OpenTextFile(FolderPath & fileName, 1)
        
        dataFlag = False
        
        ' Loop through each line in file
        Do Until txtFile.AtEndOfStream
            
            textline = txtFile.ReadLine
            
            ' Check if line is "data"
            If textline = "data" Then
                dataFlag = True
                ' Skip the header
                textline = txtFile.ReadLine
            ElseIf textline = "end data" Then
                dataFlag = False
            ElseIf dataFlag Then
                ' Split line into fields
                record = Split(textline, ",")
                
                ' Write fields to worksheet
                ws.Cells(outputRow, 1).value = CDate(record(0)) ' ob_end_time
                If IsNumeric(record(8)) Then ws.Cells(outputRow, 2).value = CDbl(record(8)) ' prcp_amt
                
                ' Move to next row
                outputRow = outputRow + 1
            End If
        Loop
        
        ' Close the file
        txtFile.Close
        
        ' Get the next file
        fileName = dir
    Wend
        
    Set fs = Nothing

End Sub


Sub DisaggregateTimeSeriesToEquidistantByImplementingSmallestTimestep(FirstRowContainsHeader As Boolean, DateCol As String, ValCol As String, DateColResult As String, ValColResult As String)

    Dim lastRow As Long, i As Long, j As Long
    Dim smallestTimeStep As Double, timeStep As Double
    Dim startTime As Date, endTime As Date
    Dim value As Double
    Dim outputRow As Long
    
    Dim PrevDate As Date, CurDate As Date
    Dim CurValue As Double

    Dim startRow As Integer
    If FirstRowContainsHeader Then startRow = 2 Else startRow = 1
    

    ' Find the last row in the original data
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, DateCol).End(xlUp).row

    ' Initialize the smallest timestep as a large number
    smallestTimeStep = 1000000

    ' Find the smallest timestep in the original data
    For i = startRow To lastRow - 1
        timeStep = ActiveSheet.Cells(i + 1, DateCol).value - ActiveSheet.Cells(i, DateCol).value
        If timeStep < smallestTimeStep Then smallestTimeStep = timeStep
    Next i

    ' Initialize the row where the first record will be written
    outputRow = startRow
    
    'simply copy the first record since we don't know the timestep size of the first one
    startTime = ActiveSheet.Cells(startRow, DateCol).value
    value = ActiveSheet.Cells(startRow, ValCol).value
        
    ActiveSheet.Cells(outputRow, DateColResult).value = startTime
    ActiveSheet.Cells(outputRow, ValColResult).value = value
    
    'Disaggregate the rest of the series
    For i = startRow + 1 To lastRow
        PrevDate = ActiveSheet.Cells(i - 1, DateCol).value
        CurDate = ActiveSheet.Cells(i, DateCol).value
        CurValue = ActiveSheet.Cells(i, ValCol).value
    
        'decide how many timesteps
        Dim nSteps As Integer
        nSteps = (CurDate - PrevDate) / smallestTimeStep
        
        For j = 1 To nSteps
            outputRow = outputRow + 1
            ActiveSheet.Cells(outputRow, DateColResult).value = PrevDate + j * smallestTimeStep
            ActiveSheet.Cells(outputRow, ValColResult).value = CurValue / nSteps
        Next
    Next i

End Sub

Function DateRoundUpToNearestHour(inputDate As Date, hoursString As String) As Date
    Dim roundedDate As Date
    Dim hourValue As Integer
    Dim i As Integer
    Dim j As Integer
    Dim nextHour As Integer
    Dim hoursArray() As String
    Dim roundedHour As Integer
    
    ' Convert the string of hours into an array
    hoursArray = Split(hoursString, ",")
    
    roundedDate = inputDate
    
    For i = 0 To 24
        roundedHour = hour(roundedDate) ' get the hour
        
        'iterate through the array of hours and see if there is a match
        For j = LBound(hoursArray) To UBound(hoursArray)
            If roundedHour = CInt(hoursArray(j)) Then
                'this is it! We found the first matching date
                DateRoundUpToNearestHour = roundedDate
                Exit Function
            End If
        Next j
        roundedDate = DateAdd("h", 1, roundedDate) 'add an hour to the date
    Next i
    
End Function

Option Explicit

Sub WriteSumaquaTimeseriesToCSV(sheetName As String, Title As String, strColumns As String, startDate As Date, intStartRow As Integer, strFilePath As String)
    'this macro writes the results of long simulations with Sumaqua as presented in Excel in multiple columns to a single CSV file
    'it also includes a date since no date is given in the
    Dim ws As Worksheet
    Dim rng As Range
    Dim rngCell As Range
    Dim arrColumns() As String
    Dim fileNumber As Integer
    Dim dataRow As Long
    Dim i As Long, n As Long
    
    ' Get an array of columns
    arrColumns = Split(strColumns, ",")
    
    ' Set the worksheet based on the sheetName argument
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Open a file for writing
    fileNumber = FreeFile
    Open strFilePath For Output As #fileNumber
    
    'write the header
    Print #fileNumber, "Date" & ";" & Title
    
    ' Iterate over each column
    Dim lastRow As Long

    For i = LBound(arrColumns) To UBound(arrColumns)
        ' Calculate last row for each column individually
        lastRow = ws.Cells(ws.Rows.Count, arrColumns(i)).End(xlUp).row
        ' Get the range of the column
        Set rng = ws.Range(ws.Cells(intStartRow, arrColumns(i)), ws.Cells(lastRow, arrColumns(i)))
    
        ' Iterate over each cell in the range until empty cell
        dataRow = 0
        For Each rngCell In rng
            If Not IsEmpty(rngCell.value) Then
                ' Write the formatted date/time and cell value to the file
                Dim days As Long
                Dim hours As Long

                days = n \ 24           'calculate the number of days. notice the \ sign!!!
                hours = n Mod 24
                Print #fileNumber, Format(startDate + days + TimeSerial(hours, 0, 0), "yyyy/MM/dd HH:mm:ss") & ";" & rngCell.value
                dataRow = dataRow + 1
                n = n + 1
            Else
                ' Exit loop on empty cell
                Exit For
            End If
        Next rngCell
    Next i
    
    ' Close the file
    Close #fileNumber
End Sub

Sub FitSineWave(DateRange As Range, ValuesRange As Range, AmplitudeCell As Range, PhaseCell As Range)
    Dim i As Long, numPoints As Long
    Dim a As Double, b As Double
    Dim amplitude As Double, phase As Double

    ' Count the number of data points
    numPoints = DateRange.Rows.Count

    ' Check if the date and values ranges have the same dimensions
    If numPoints <> ValuesRange.Rows.Count Then
        Exit Sub
    End If

    ' Temporary arrays for calculations
    Dim cosValues() As Double, sinValues() As Double, yValues() As Double
    ReDim cosValues(1 To numPoints), sinValues(1 To numPoints), yValues(1 To numPoints)

    ' Fill temporary arrays
    For i = 1 To numPoints
        cosValues(i) = Cos(DateRange.Cells(i, 1).value)
        sinValues(i) = Sin(DateRange.Cells(i, 1).value)
        yValues(i) = ValuesRange.Cells(i, 1).value
    Next i

    ' Calculate a and b coefficients using LINEST function
    Dim linestResults As Variant
    linestResults = Application.Run("LINEST", yValues, Application.WorksheetFunction.Transpose(Array(cosValues, sinValues)), True, False)
    a = linestResults(1, 1)
    b = linestResults(2, 1)

    ' Calculate amplitude and phase
    amplitude = Sqr(a ^ 2 + b ^ 2)
    phase = WorksheetFunction.ATan2(b, a)

    ' Write amplitude and phase to the specified cells
    AmplitudeCell.value = amplitude
    PhaseCell.value = phase
End Sub

Public Function CalculateQHRelationship(YZProfileRange As Range, slopeCell As Range, roughnessCell As Range, Optional maxDepth As Double = 10, Optional depthStep As Double = 0.1) As Variant
    Dim results() As Variant
    Dim waterLevel As Double
    Dim wetArea As Double, wetPerimeter As Double, hydraulicRadius As Double
    Dim chezy As Double, discharge As Double
    Dim slope As Double, roughness As Double
    Dim i As Long, numSteps As Long
    Dim areaCollection As New Collection
    Dim perimeterCollection As New Collection
    Dim radiusCollection As Collection
    Dim lowestPoint As Double
    
    ' Read slope and roughness
    slope = slopeCell.value
    roughness = roughnessCell.value
    
    ' Find the lowest point in the YZ-profile
    lowestPoint = Application.WorksheetFunction.Min(YZProfileRange.Columns(2))
    
    ' Calculate number of steps
    numSteps = Application.RoundUp(maxDepth / depthStep, 0)
    
    ' Initialize results array
    ReDim results(1 To numSteps, 1 To 2)
    
    ' Calculate QH relationship
    For i = 1 To numSteps
        
        waterLevel = lowestPoint + (i - 1) * depthStep
                
        ' Calculate wetted area and perimeter
        wetArea = WettedAreaFromYZProfile(YZProfileRange, waterLevel)
        wetPerimeter = WettedPerimeterFromYZProfile(YZProfileRange, waterLevel)
        
        ' Add to collections
        areaCollection.Add wetArea
        perimeterCollection.Add wetPerimeter
        
        ' Calculate hydraulic radius
        Set radiusCollection = HydraulicRadiusFromWettedAreasAndWettedPerimeters(areaCollection, perimeterCollection)
        If wetPerimeter <> 0 Then
            hydraulicRadius = wetArea / wetPerimeter
        Else
            hydraulicRadius = 0
        End If
        
        
        ' Convert Manning's n to Chezy C
        chezy = Manning2Chezy(roughness, hydraulicRadius)
        
        ' Calculate discharge using Chezy formula
        discharge = wetArea * chezy * Sqr(hydraulicRadius * slope)
        
        ' Store results
        results(i, 1) = waterLevel
        results(i, 2) = discharge
        
        ' Clear collections for next iteration
        Set areaCollection = New Collection
        Set perimeterCollection = New Collection
    Next i
    
    CalculateQHRelationship = results
End Function


Function GetAttributeValueFromString(attributeString As String, attributeName As String, Optional attributeSeparator As String = " ", Optional nameValueSeparator As String = "=") As String
    Dim attributes() As String
    Dim myattribute As Variant
    Dim i As Long
    
    'first clean up our string by removing any soft returns
    attributeString = Replace(attributeString, vbCrLf, "")
    attributeString = Replace(attributeString, vbCr, "")
    attributeString = Replace(attributeString, vbLf, "")
        
    ' Split the string into individual attributes
    attributes = Split(attributeString, attributeSeparator)
    
    ' Loop through each attribute
    For i = LBound(attributes) To UBound(attributes)
        ' Split each attribute into name and value
        myattribute = Split(attributes(i), nameValueSeparator)
        
        ' Check if the attribute has both name and value
        If UBound(myattribute) >= 1 Then
            ' Check if the attribute name matches the one we're looking for
            If LCase(Trim(myattribute(0))) = LCase(Trim(attributeName)) Then
                ' Return the value if found
                GetAttributeValueFromString = Trim(myattribute(1))
                Exit Function
            End If
        End If
    Next i
    
    ' If attribute is not found, return an empty string
    GetAttributeValueFromString = ""
End Function

