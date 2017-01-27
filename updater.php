<?php
/**
 * Copyright (c) 2017 Arvis Lācis
 * Automatizēts skripts Latvijas pasta indeksu iegūšanai un strukturēšanai uzskatāmā formā
 */

// Skripta darbībai nepieciešama PHPExcel bibliotēka izklājlapu satura nolasīšanai
// https://github.com/PHPOffice/PHPExcel
require_once "./PHPExcel.php";

$outputFilename = "data.csv";

// RegEx izteiksme pasta indeksu izklājlapu saišu iegūšanai no analizējamās Latvijas Pasta URL
$fileURLRegex = "/\\/files\\/pakalpojumi_pp\\/.*\\.xls/";
// Analizējāmās Latvijas Pasta vietnes adrese (URL)
$prefixURL = "http://www.pasts.lv/";
$sourceURL = $prefixURL . "lv/uznemumiem/noderigi/pasta_indeksu_gramata/";
$sourceContent = file_get_contents($sourceURL);

// Pagaidu datnes nosaukums, kurā īslaicīgi tiek uzglabāti lejupielādētie pasta indeksu dati
$dataFilename = "index_data.xls";
$excelReader = PHPExcel_IOFactory::createReader("Excel5");
// Masīvs visu iegūto pasta indeksu uzglabāšanai
$indexDatabase = array();

// Iegūst saites uz pasta indeksu izklājlapām to tālākai apstrādei
preg_match_all($fileURLRegex, $sourceContent, $urls);

// Iterē caur iegūto saišu sarakstu, apstrādājot katru pasta indeksu datni atsevišķi
foreach ($urls[0] as $url) {
    // Apstrādā tikai pilsētu un novadu indeksus, ignorējot speciālos indeksus
    if (!strpos($url, "Specialie")) {
        // Lejupielādē apstrādājamo izklājlapu datni un saglabā tās saturu
        file_put_contents($dataFilename, file_get_contents($prefixURL . $url));

        // Veic lejupielādētās datnes atvēršanu/ielādi ar PHPExcel bibliotēkas palīdzību
        $excel = $excelReader->load($dataFilename);
        $sheets = $excel->getSheetCount();

        // Izvēlas tālāko apstrādes algoritmu - pilsētas vai novadu datu apstrādei
        if (!strpos($url, "Novadi")) {
            // Pilsētas pasta indeksu apstrāde
            // Pārbauda vai datne satur tikai vienu izklājlapu (viena izklājlapa = viena pilsēta)
            if ($sheets === 1) {
                // Aktīvās izklājlapas iestatīšana un parametru nolasīšana (pilsētas nosaukums, izklājlapas rindu un kolonnu skaits)
                $activeSheet = $excel->setActiveSheetIndex(0);
                $city = $activeSheet->getTitle();
                $columns = PHPExcel_Cell::columnIndexFromString($activeSheet->getHighestDataColumn());
                $rows = $activeSheet->getHighestRow();

                // Pārbauda vai izklājlapa satur tieši 4 datu kolonnas
                if ($columns == 4) {
                    // Iterē caur izklājlapas rindām, veic to šūnu vērtību nolasīšanu un apstrādi
                    for ($i = 2; $i <= $rows; $i++) {
                        // Nolasa ielas nosaukumu un pasta indeksu
                        $street = $activeSheet->getCellByColumnAndRow(1, $i)->getValue();
                        $index = $activeSheet->getCellByColumnAndRow(3, $i)->getValue();

                        // Pārbauda vai ir iestatīta gan ielas nosaukuma, gan pasta indeksa vērtība,
                        // pretējā gadījumā apstrādājamā rinda ir tukša (satur tikai alfabēta burta galveni) un to var izlaist
                        if ($street and $index) {
                            // Atsevišķi izdala mājas numuru vērtības
                            $numbers = explode(",", $activeSheet->getCellByColumnAndRow(2, $i)->getValue());

                            // Iterē caur iegūtajām mājas numuru vērtībām un formē pilno pasta indeksa ierakstu,
                            // kas tiek pievienots kopējā pasta indeksu masīvā
                            foreach ($numbers as $number) {
                                $indexDatabase[] = array(
                                    "index" => $index,
                                    "street" => $street,
                                    "street_number" => $number,
                                    "region" => "",
                                    "city" => $city,
                                    "parish" => "",
                                    "village" => ""
                                );
                            }
                        }
                    }
                } else {
                    throw new Exception("Pilsētas pasta indeksu izklājlapa nesatur tieši 4 datu kolonnas!");
                }
            } else {
                throw new Exception("Pilsētas pasta indeksu datne var saturēt tikai vienu izklājlapu!");
            }
        } else {
            // Novadu pasta indeksu apstrāde
            //
            for ($i = 0; $i < $sheets; $i++) {
                // Aktīvās izklājlapas iestatīšana un parametru nolasīšana (novada nosaukums, izklājlapas rindu un kolonnu skaits)
                $activeSheet = $excel->setActiveSheetIndex($i);
                $region = $activeSheet->getTitle();
                $columns = PHPExcel_Cell::columnIndexFromString($activeSheet->getHighestDataColumn());
                $rows = $activeSheet->getHighestRow();

                // Pārbauda vai izklājlapa satur tieši 5 datu kolonnas
                if ($columns === 5) {
                    // Iterē caur izklājlapas rindām, veic to šūnu vērtību nolasīšanu un apstrādi
                    for ($j = 2; $j < $rows; $j++) {
                        // Nolasa novada nosaukumu un pasta indeksu
                        $regionTest = $activeSheet->getCellByColumnAndRow(0, $j)->getValue();
                        $index = $activeSheet->getCellByColumnAndRow(4, $j)->getValue();

                        // Papildus pārbaude, lai pārliecinātos, ka izklājlapas nosaukumā norādītais novada nosaukums
                        // sakrīt ar datu kolonnas šūnā sniegto
                        if ($region === $regionTest) {
                            // Iegūst pilsētas, pagasta un ciema nosaukumu (ja tādi ir norādīti)
                            $city = $activeSheet->getCellByColumnAndRow(1, $j)->getValue();
                            $parish = $activeSheet->getCellByColumnAndRow(2, $j)->getValue();
                            $village = $activeSheet->getCellByColumnAndRow(3, $j)->getValue();

                            // Formē pilno pasta indeksu un pievieno to kopējā indeksu masīvā
                            $indexDatabase[] = array(
                                "index" => $index,
                                "street" => "",
                                "street_number" => "",
                                "region" => $region,
                                "city" => $city,
                                "parish" => $parish,
                                "village" => $village
                            );
                        } else {
                            throw new Exception("Izklājlapas nosaukumā sniegtais novada nosaukums nesakrīt ar datu kolonnā norādīto novadu!");
                        }
                    }
                } else {
                    throw new Exception("Novadu pasta indeksu izklājlapa nesatur tieši 5 datu kolonnas!");
                }
            }
        }
    }
}

/**
 * Funkcija divu teksta virkņu salīdzināšanai.
 * Tiek ņemti vērā latviešu burti (UTF-8) un dabiskās salīdzināšanas/kārtošanas principi.
 *
 * @param string $string1 Pirmā teksta virkne
 * @param string $string2 Otrā teksta virkne
 * @return int Salīdzināšanas rezultāts: 1 - pirmā virkne ir lielāka (kārtojama pēc); -1 - pirmā virkne ir mazāka (kārtojama pirms); 0 - abas virknes ir vienādas
 */
function compareStrings($string1, $string2)
{
    // Pirmās teksta virknes garums
    $length = strlen($string1);

    // Veic cikla izpildi tik ilgi, kamēr analizēti visi pirmās virknes simboli vai arī
    // kamēr ir zināms salīdzināšanas rezultāts un tiek atgriezta funkcijas vērtība
    for ($i = 0; $i < $length; $i++) {
        // Ja otrajā virknē vairs nav simbolu, tad neapšaubāmi pirmā virkne ir lielāka
        if (!isset($string2[$i])) {
            return 1;
        }

        // Pārbauda vai kāds no aktīvajiem abu virkņu simboliem nav cipars (ir burts vai cita rakstzīme)
        if (!is_numeric($string1[$i]) or !is_numeric($string2[$i])) {
            // Veic parastu burtu, simbolu vērtību salīdzināšanu
            if ($string1[$i] > $string2[$i]) {
                return 1;
            } elseif ($string1[$i] < $string2[$i]) {
                return -1;
            } else {
                $next = $i + 1;

                // Ja abas virknes tiktāl ir vienādas un sasniegts pirmās virknes garums, tad
                // pārbauda vai otrajā teksta virknē vēl atlicis kāds simbols (līdzīga pārbaude kā cikla sākumā)
                if ($next === $length and isset($string2[$next])) {
                    return -1;
                }
            }
        } else {
            // Specifiska apstrādes, kārtošanas procedūra galvenokārt ielas numuru vērtībām, tas ir, lai
            // realizētu dabisko kārtošanu, kur pēc 1 seko 2 un 3, nevis 11 un 111 utml.

            // Ja sākotnējie virkņu simboli nav bijuši cipari, tad šo daļu nogriež un vairs neizmanto
            // Ar RegEx izteiksmes palīdzību veic pirmās sastopamās un lielākās iespējamās skaitliskās vērtības izgūšanu un
            // tās pārveidošanu vesela skaitļa (int) vērtībā
            $stringPart1 = substr($string1, $i);
            preg_match("/\\d+/", $stringPart1, $matches);
            $intNumber1 = (int)$matches[0];

            $stringPart2 = substr($string2, $i);
            preg_match("/\\d+/", $stringPart2, $matches);
            $intNumber2 = (int)$matches[0];

            // Veic parastu skaitlisko vērtību salīdzināšanu
            if ($intNumber1 > $intNumber2) {
                return 1;
            } elseif ($intNumber1 < $intNumber2) {
                return -1;
            } else {
                // Ja abas skaitliskās vērtības ir vienādas, tad pārbauda specifiskus papildu gadījumus
                // Piemēram, ielas numurs ir "5 k-1", "5 k-2", "5 k-11" utml.
                $part1KPosition = strpos($stringPart1, "k-");
                $part2KPosition = strpos($stringPart2, "k-");

                // Ja abas virknes papildus satur "k-", tad rekursīvi izsauc šo virkņu salīdzināšanas funkciju
                // atlikušajām teksta virkņu daļām
                if ($part1KPosition > 0 and $part2KPosition > 0) {
                    return compareStrings(
                        substr($stringPart1, $part1KPosition + 2),
                        substr($stringPart2, $part2KPosition + 2)
                    );
                } else {
                    // Pretējā gadījumā veic abu virkņu garuma pārbaudi, piemēram, lai apstrādātu
                    // situācijas, kad ielas numurs ir "90" un "90A" utml.
                    $part1Length = strlen($stringPart1);
                    $part2Length = strlen($stringPart2);

                    // Veic salīdzināšanu pēc atlikušās daļas garuma, lai "3A" vienmēr sekotu "3" utml.
                    if ($part1Length > $part2Length) {
                        return 1;
                    } elseif ($part1Length < $part2Length) {
                        return -1;
                    } else {
                        // Ja abas virknes aizvien ir vienāda garuma, piemēram, "5B" un "5C", tad
                        // veic skaitliskās daļas atmešanu un atlikušajai abu virkņu daļai
                        // rekursīvi izsauc šo virkņu salīdzināšanas funkciju
                        $intNumber1Length = strlen((string)$intNumber1);
                        $intNumber2Length = strlen((string)$intNumber2);

                        return compareStrings(
                            substr($stringPart1, strpos($stringPart1, (string)$intNumber1) + $intNumber1Length),
                            substr($stringPart2, strpos($stringPart2, (string)$intNumber2) + $intNumber2Length)
                        );
                    }
                }
            }
        }
    }

    return 0;
}

// Pielāgota kārtošanas funkcija izveidotā pasta indeksu masīva sakārtošanai pārskatāmā (alfabēta) veidā
// Kārtošanas secība: pasta indekss, ielas nosaukums, ielas numurs, novads, pilsēta, pagasta un ciems
usort($indexDatabase, function($a, $b) {
    $compare = compareStrings($a["index"], $b["index"]);
    if ($compare) return $compare;

    $compare = compareStrings($a["street"], $b["street"]);
    if ($compare) return $compare;

    $compare = compareStrings($a["street_number"], $b["street_number"]);
    if ($compare) return $compare;

    $compare = compareStrings($a["region"], $b["region"]);
    if ($compare) return $compare;

    $compare = compareStrings($b["city"], $a["city"]);
    if ($compare) return $compare;

    $compare = compareStrings($a["parish"], $b["parish"]);
    if ($compare) return $compare;

    return compareStrings($a["village"], $b["village"]);
});

// Pievieno galvenes rindiņu izvadāmajiem datiem
$outputContent = "indekss,iela,ielas numurs,novads,pilsēta,pagasts,ciems\n";

// Iterē caur formēto, sakārtoto pasta indeksu masīvu un pievieno pasta indeksu rindiņas
// gala rezultātā saglabājamajā izvadē
foreach ($indexDatabase as $record) {
    $outputContent .= implode(",", $record) . "\n";
}

// Saglabā izvades teksta virkni pasta indeksu CSV datnē
file_put_contents($outputFilename, $outputContent);
// Dzēš lejupielādēto pagaidu pasta indeksu datni
unlink("index_data.xls");
