<#
.SYNOPSIS
   Lädt alle Statistikdateien der PV-Anlage eines Jahres herunter und speichert sie in einer Sammeldatei für das Jahr
   
.DESCRIPTION
   <A detailed description of the script>
.PARAMETER <paramName>
   -jahr: Jahr für das die Dateien heruntergeladen werden sollen
   -modus: "vollstaendig": alle Dateien eines Jahres 
           "update":       alle ab der laufenden Woche bis zur zuletzt heruntergeladenen Wochendatei
.EXAMPLE
   <An example of using the script>
#>
param([string]$jahr, [validateSet("vollstaendig", "update")][string]$modus)

class Datensatz {
	[DateTime] $Zeitstempel
	[Double[]]$werte = @(0.0) * 9
}	

function datensatzSplitten {
	param([string]$textDatensatz)
	$datensatz = [Datensatz]::new()
	[String[]]$daten = $textDatensatz.Split(";")

	# Daten auf korrektes Format testen und in Hashtable zuweisen
	# erster Wert muss Datum sein
	if (($datensatz.zeitstempel = $daten[0] -as [datetime]) -eq $null) {
		"falsches Datumsformat: $daten[0]"
		$fehler = $true
	}
	else {
		#$zahlenformat = (Get-Culture).numberformat
		# alle folgende Werte müssen Zahlen sein
		for ($i = 1; $i -lt $daten.Length; $i++) {
			try {
				$datensatz.werte[$i-1] = $daten[$i].ToDouble($zahlenformat)
			}
			catch [FormatException] {
				"falsches Zahlenformat: $daten[0] - $i. Wert: $"
				$fehler = $true
			}
		}
	}
	if ($fehler) {$datensatz = $null}
	return $datensatz
}

function ausgabeZusammensetzen {
	param ([datetime]$zeit, [Double[]]$werte, [String]$format)
	$zf = (Get-Culture).NumberFormat
    $zeitformat = (Get-Culture).DateTimeFormat
	[String]$zeitString = ""
	switch ($format) {
		"h" {$zeitString = ($zeit.ToString((get-culture).datetimeformat) -replace "(.* \d\d:)\d\d:\d\d", '$1') + "00"}
		"d" {$zeitString = ($zeit.ToString((get-culture).datetimeformat) -replace "(.* )\d\d:\d\d:\d\d", '$1') + "00:00" }
		"m" {$zeitString = ($zeit.ToString((get-culture).datetimeformat) -replace "(.* )\d\d:\d\d:\d\d", '$1') + "00:00"}
	}
	
	$ausgabeZeile = [String]::Join(";",($zeitString.ToString($zeitformat), $werte[2].ToString($zf), $werte[5].ToString($zf),`
                                                                     $werte[1].toString($zf), $werte[0].ToString($zf),` 
                                                                     $werte[3].toString($zf), $werte[4].ToString($zf)))
	return $ausgabeZeile
}


$datenPfad = "C:\Daten\Anwendungen\PhotovoltaikDaten\Jahr$jahr\"
$dateiMitGewichtung = $datenPfad + "jahr-" + $Jahr + "mitGewichtung.csv"
$stundenDatei = $datenPfad + "Jahr-" + $Jahr + "-Stunde.csv"
$tagesDatei = $datenPfad + "Jahr-" + $Jahr + "-Tag.csv"
$monatsDatei = $datenPfad + "Jahr-" + $Jahr + "-Monat.csv"

 [System.Collections.ArrayList]$stundenDaten = @()
 [System.Collections.ArrayList]$tagesDaten = @()
 [System.Collections.ArrayList]$monatsDaten = @()

# Die Daten über Stunden, Tage und Monate kummulieren 
# und in separate Dateien schreiben
$jahresDaten = Get-Content -Path $dateiMitGewichtung
# erster Datensatz muss die Spaltenüberschriften beinhalten
if ($jahresDaten[0] -notlike "Uhrzeit*") {
	exit
}


# Aus den Spaltenüberschriften die Gewichtung rausnehmen, wird nicht mehr benötigt
$ue = $jahresDaten[0].Split(";")
$spaltenUeberschriften = [String]::Join(";",($ue[0], $ue[3], $ue[6], $ue[2], $ue[1], $ue[4], $ue[5])); 
#$spaltenUeberschriften | Out-File $stundenDatei
#$spaltenUeberschriften | Out-File $tagesDatei
#$spaltenUeberschriften | Out-File $monatsDatei

$stundenDaten.Add($spaltenUeberschriften) | Out-Null
$tagesDaten.Add($spaltenUeberschriften) | Out-Null
$monatsDaten.Add($spaltenUeberschriften) | Out-Null

[String]$ausgabeZeile = ""

$zahlenformat = (get-culture).numberformat
$zeitformat = (Get-Culture).DateTimeFormat

$datensatz = datensatzSplitten -textDatensatz $jahresDaten[1]
if ($datensatz -eq $null) {
	exit
}

$anzahlWerte = $datensatz.werte.length - 2
[Double[]]$summeStunden = ($Datensatz.werte[0..($anzahlWerte-1)]).clone()
[Double[]]$summeTage = $summeStunden.Clone()
[Double[]]$summeMonate = $summeStunden.Clone()

$aktuelleStunde = $datensatz.zeitstempel
$aktuellerTag = $datensatz.zeitstempel
$aktuellerMonat = $datensatz.zeitstempel

#foreach ($zeile in $jahresDaten[2..($jahresDaten.Length-1)]) {
for ($j = 2; $j -lt $jahresDaten.Length; $j++) {
#$zeile in $jahresDaten[2..($jahresDaten.Length-1)]) {
	#$datensatz = datensatzSplitten -textDatensatz $zeile
    $datensatz = datensatzSplitten -textDatensatz $jahresDaten[$j]

	if ($datensatz -ne $null) {
		for ($i = 0;$i -lt $anzahlWerte;$i++) {
			$wert = $datensatz.werte[$i] * $datensatz.werte[-1]
			$summeStunden[$i] += $wert
			$summeTage[$i] += $wert
			$summeMonate[$i] += $wert
		}
		
		if ($datensatz.zeitstempel.hour -ne $aktuelleStunde.Hour) {
			$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuellestunde -werte $summeStunden -format "h"
			$stundenDaten.Add($ausgabeZeile) | Out-Null
			$aktuelleStunde = $datensatz.zeitstempel
			for ($i = 0; $i -lt $summeStunden.length; $i++) {
				$summeStunden[$i] = 0.0
			}

		}
		if ($datensatz.zeitstempel.Day -ne $aktuellerTag.Day) {
			$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuellerTag -werte $summeTage -format "d"
			$tagesDaten.Add($ausgabeZeile) | Out-Null
			$aktuellerTag = $datensatz.zeitstempel
			for ($i = 0; $i -lt $summeTage.length; $i++) {
				$summeTage[$i] = 0.0
			}
		}
		if ($datensatz.zeitstempel.Month -ne $aktuellerMonat.Month) {
			$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuellerMonat -werte $summeMonate -format "m"
			$monatsDaten.Add($ausgabeZeile) | Out-Null
			$aktuellerMonat = $datensatz.zeitstempel
			for ($i = 0; $i -lt $summeMonate.length; $i++) {
				$summeMonate[$i] = 0.0
			}
		}
	}
}

$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuelleStunde -werte $summeStunden -format "h"
$stundenDaten.Add($ausgabeZeile) | Out-Null
$stundenDaten | Set-Content $stundenDatei

$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuellerTag -werte $summeTage -format "d"
$tagesDaten.Add($ausgabeZeile) | Out-Null
$tagesDaten | Set-Content $tagesDatei

$ausgabeZeile = ausgabeZusammensetzen -zeit $aktuellerMonat -werte $summeMonate -format "m"
$monatsDaten.Add($ausgabeZeile) | Out-Null
$monatsDaten | Set-Content $monatsDatei






