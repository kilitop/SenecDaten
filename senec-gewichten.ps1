<#
.SYNOPSIS
   - �berpr�ft die Jahresdatei 
   	 gespeichert im Ordner = "C:\Daten\Anwendungen\PhotovoltaikDaten\", Unterordner "Jahr<Jjahr>\"
	 Dateiname "jahr-<Jahr>.csv" mit <Jahr> = als Paramter �bergebenes Jahr
	 auf korrektes Datenformat
     Der 1. Datensatz muss eine Spalten�berschrift sein. Alle anderen noch enthaltenen �berschriftszeilen
	 werden �berlesen, alle anderen Datens�tze m�sssen in der ersten Spalte ein Datum/Uhrzeit enthalten,
	 die anderen Spalten Gleitkommazahlen, die Spalten sind durch ein ";" getrennt
   - f�gt eine Gewichtung gem�� der Dauer (Zeitabstand zwischen Vorg�nger- und aktuellem datensatz) hinzu
   - Ein Datensatz der �ber eine Stundengrenze hinausgeht (Stunde des Vorg�ngers kleiner als aktuelle Stunde)
     wird in zwei Datens�tze gesplittet, wobei die Grenze die volle Stunde ist. Die Gewichtung der beiden 
	 Datens�tze wird entsprechend gesetzt
   - wird in der Datei "Jahr-<Jahr der auswertung>Gewichtet.csv" gespeichert
   
.DESCRIPTION
   <A detailed description of the script>
.PARAMETER <paramName>
   -paramDatei: Datei mit URLs der SENEC-Seiten, Dateipfade f�r Daten und die Anmeldeinformationen
   -jahr:       Jahr f�r das die Dateien gewichtet werden sollen
.EXAMPLE
   <An example of using the script>
#>
param([String]$paramDatei, [string]$jahr)

class Datensatz {
	[DateTime] $zeitstempel 
	$werte = [System.Collections.Generic.List[Double]]@()
} 

# Datensatz auf Konsistenz pr�fen:
# 1. Spalte hat Datumsformat, alle anderen sind Gleitkommazahlen, Trennungszeichen ist ";"
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
		# alle folgende Werte m�ssen Zahlen sein
		for ($i = 1; $i -lt $daten.Length; $i++) {
			try {
				$datensatz.werte.Add($daten[$i].ToDouble($zahlenformat))
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

# Parameter �berpr�fen ---------------------------------
$fehler = 0
if ($paramDatei -eq "" ) {
	"XML-Datei mit URLs und Datenpfaden nicht angegeben."
	$fehler++
}	
elseif (-not (Test-Path $paramDatei)) {
	"XML-Datei mit URLs und Datenpfaden existiert nicht: " + $paramDatei
	$fehler++
}

if ($jahr -eq $null-or $jahr -eq "") {
	'Parameter "Jahr" nicht angegeben' 
	$fehler++
}

if ($fehler -gt 0) {
	exit
}
# Ende Parameter �berpr�fen -----------------------------

# Parameterdatei einlesen und Datenpfade URLs und Login-Daten f�r SENEC-Homepage setzen.
[xml]$datenpfade = Get-Content $ParamDatei
$datenPfad = $datenpfade.dataPathes.folders.data

$datensatz = [Datensatz]::new()
[DateTime]$vorZeitstempel = Get-Date
[DateTime]$Zeitstempel = Get-Date
[DateTime]$zwischenZeitstempel = Get-Date

$zahlenformat = (get-culture).numberformat
$zeitformat = (Get-Culture).DateTimeFormat

# Dateien definieren: <Datenpfad aus xml-Datei>\jahr<Jahr aus Parameter>\jahr-<Jahr aus Parameter>.csv bzw
#                     <Datenpfad aus xml-Datei>\jahr<Jahr aus Parameter>\jahr-<Jahr aus Parameter>MitGweichtung.csv
$jahresdatei = $datenPfad + "Jahr" + $jahr + "\" + "jahr-" + $jahr + ".csv"
$jahresDateiMitGewichtung = $datenPfad + "jahr" + $Jahr  + "\" + "jahr-" + $jahr + "MitGewichtung.csv"

# Jahresdatei einlesen
$jahresDaten = Get-Content -Path $jahresdatei

# Den Daten eine Gewichtung �ber die Zeitdauer geben 
# Datens�tze, die �ber eine Stundengrenze hinweggehen, auf zwei Datens�tze aufsplitten

# erster Datensatz muss die Spalten�berschriften beinhalten
if ($jahresDaten[0] -notlike "Uhrzeit*") {
	exit
}

# leeres Ausgabe-Array anlegen
 [System.Collections.ArrayList]$gewichtet = @()
 
# Aus den Spalten�berschriften die Ma�einheiten, Leerzeichen und Sonderzeichen/Umlaute entfernen
# Gewichtung als Spalte am Ende hinzuf�gen
$spaltenUeberschriften = ($jahresDaten[0] -replace "\[.*?\]|\-" , "") -replace " ", ""
$spaltenUeberschriften = $spaltenUeberschriften -replace "�" , "ae"
$spaltenUeberschriften += ";Gewichtung"
#$spaltenUeberschriften | Out-File $jahresDateiMitGewichtung
$index = $gewichtet.Add($spaltenUeberschriften)

# erster Datensatz wird mit 5 Minuten veranschlagt
$datensatz = datensatzSplitten -textDatensatz $jahresDaten[1]
if ($datensatz -eq $null) {
	exit
}

$jahresDaten[1] + ";" + (300.0 / 3600.0).ToString($zahlenformat) | Out-File $jahresDateiMitGewichtung -Append
$vorZeitstempel = $datensatz.zeitstempel

"Anzahl Datens�tze = $($jahresDaten.Length)"
foreach ($zeile in $jahresDaten[2..($jahresDaten.Length-1)]) {
	if ($zeile -notlike "Uhrzeit*"){
		$datensatz = datensatzSplitten -textDatensatz $zeile

		if ($datensatz -ne $null) {
			$verhaeltnis = 1.0
			$gewichtung = [Double](([timespan]($datensatz.zeitstempel - $vorZeitstempel)).TotalSeconds) / 3600.0 
			# Datensatz �berschreitet Stundengrenze. Es wird ein zus�tzlicher Datensatz bis zur vollen Stunde erstellt
			# Der zus�tzliche und der aktuelle Datensatz werden entsprechend ihrer Dauer gewichtet
			if ($datensatz.zeitstempel.hour -ne $vorZeitstempel.Hour) {
				$zeile1 = $zeile.split(";")
				$zwischenZeitstempel = $datensatz.zeitstempel
				$zwischenZeitstempel = $zwischenZeitstempel.AddMinutes(-$zwischenZeitstempel.Minute)
				$zwischenZeitstempel = $zwischenZeitstempel.AddSeconds(-$zwischenZeitstempel.Second)
				$zwischenZeitstempel = $zwischenZeitstempel.AddMilliseconds(-($zwischenZeitstempel.MilliSecond+1))
				$zeile1[0] = $zwischenZeitstempel.ToString($zeitformat)
				$verhaeltnis = [Double](([timespan]($zwischenZeitstempel - $vorZeitstempel)).TotalSeconds)/ `
							   [Double](([timespan]($datensatz.zeitstempel - $vorZeitstempel)).TotalSeconds)
				[String]$zwischenzeile = ([String]::Join(";",$zeile1)) + ";" + ($gewichtung * $verhaeltnis).ToString($zahlenformat)
				$index = $gewichtet.Add($zwischenzeile)
				$verhaeltnis = 1.0 - $verhaeltnis
			}
			
			$zeile = $zeile + ";" + ($gewichtung * $verhaeltnis).ToString($zahlenformat)
			#$zeile | Out-File $JahresdateiMitGewichtung -Append
			$index = $gewichtet.Add($zeile)
			$vorZeitstempel = $datensatz.zeitStempel;
		}
	}
}

$gewichtet | Set-Content $jahresDateiMitGewichtung





