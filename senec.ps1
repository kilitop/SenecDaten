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
	[Double[]] $werte
}	

function download_statistikdaten {
	param($woche, $jahr, $zielDatei, $modus)

	$downloadParams["woche"] = $woche
	try {
		$statistikdaten = Invoke-WebRequest -Uri $uridownload -WebSession $senec -Body $downloadParams `
		-OutFile $zielDatei -ErrorAction silentlycontinue `
		-ErrorVariable downloadError
	}
	catch {
		"Category: "
		$downloadError[0].ErrorRecord.CategoryInfo | fl *
		"Exception: "
		$downloadError[0].errorRecord.exception
		$downloadError[0].errorRecord.exception.Response | fl *
	}
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
		# alle folgende Werte müssen Zahlen sein
		for ($i = 1; $i -lt $daten.Length; $i++) {
			if (($datensatz.werte[$i - 1] = $daten[$i] -as [double]) -eq $null) {
				"falsches Zahlenformat: $daten[0] - $i. Wert: $"
				$fehler = $true
			}
		}
	}

	if ($fehler) {$datensatz = $null}

	return $datensatz
}



$datensatz = [Datensatz]::new()
[DateTime]$vorZeitstempel = $null
[DateTime]$Zeitstempel = $null
[DateTime]$zwischenZeitstempel = $null

$uri = "https://mein-senec.de/auth/login"
$uriStatistik = "https://mein-senec.de/endkunde/#/0/statistischeDaten"
$uriDownload = "https://mein-senec.de/endkunde/api/statistischeDaten/download"

$datenPfad = "C:\Daten\Anwendungen\PhotovoltaikDaten\"
$dateiMitGewichtung = $datenPfad + "jahr" + $Jahr + "-mitGewichtung"
$stundenDatei = $datenPfad + "jahr" + $Jahr + "-Stunde.csv"
$tagesDatei = $datenPfad + "jahr" + $Jahr + "-Tag.csv"
$monatsDatei = $datenPfad + "jahr" + $Jahr + "-Monat.csv"
$credentialDatei = "C:\Daten\allgemein\senecLogin.xml"
$passwordDatei = "C:\Daten\allgemein\senecPW.txt"

#Paramter für Download der Dateien
$downloadParams = @{anlageNummer = '0'
	woche = '1'
	jahr = '2020'}

#lese Anmeldeinformation aus Datei
$credentials = Import-Clixml -Path $credentialDatei
#$password = Get-Content $passwordDatei | ConvertTo-SecureString
$password = $credentials.GetNetworkCredential().password

#Anmeldeseite von Senec aufrufen
$loginpage = Invoke-WebRequest -Uri $uri -SessionVariable 'senec' 

# Das darin enthaltene Anmeldefomular füllen
$loginpage.Forms[0].Fields["username"] = $credentials.getNetworkCredential().UserName
$loginpage.Forms[0].Fields["password"] = $credentials.GetNetworkCredential().Password

# Anmelderequest abschicken
$loginresponse = Invoke-WebRequest -Uri $uri -WebSession $senec -Method Post -Body $loginpage.Forms[0]

# Nachprüfen, ob Login erfolgreich
if ($loginresponse.ParsedHtml.body.outerText -like "*invalid login data*") {
	"ungültige Anmeldedaten"
	break
}
$laufendeWoche = Get-Date -UFormat %V

$datenPfadDownloads = $datenPfad + "Jahr" + $jahr + "\"
$letzteDateiNochmals = $true 

"downloads beginnen"
for ($i = [int]$laufendeWoche; $i -ge 1; i--) {
	$textWoche = ("{0:d2}" -f $i)
	$downloadDatei = $datenpfadDownloads + "woche-" + $jahr + "-" + $textWoche + ".csv"
	if (!(Test-Path $downloadDatei) -or $modus -eq "vollstaendig") {
		download_statistikdaten -woche $textWoche -jahr $jahr -zielDatei $downloadDatei -modus $modus
	}
	else{
		if ($letzteDateiNochmals) {
			download_statistikdaten -woche $textWoche -jahr $jahr -zielDatei $downloadDatei -modus $modus
			$letzteDateiNochmals = $false
			break
		}
	}
}

if ($jahr -lt (Get-Date).Year) {
	$downloadDatei = $datenpfadDownloads + "woche-" + ($jahr + 1) + "-" + "55.csv"
	download_statistikdaten -woche "01" -jahr $(jahr +1) -zielDatei $downloadDatei -modus $modus
}

"downloads beendet"
# alle Wochendateien in eine Jahresdatei überführen
$jahresDaten = Get-ChildItem -Path $datenPfadDownloads -Filter "woche-$jahr-*.csv" | Get-Content | Set-Content -Path ("$datenPfadDownloads" + "jahr-$jahr" + ".csv")

# Den Daten eine Gewichtung über die Zeitdauer geben 
# Dateien mit Datensätze für ganze Stunden, ganze Tage und ganze Monate erstellen

# erster Datensatz muss die Spaltenüberschriften beinhalten
if ($jahresDaten[0] -notlike "Uhrzeit*") {
	exit
}

# Aus den Spaltenüberschriften die Maßeinheiten, Leerzeichen und Sonderzeichen/Umlaute entfernen
$spaltenUeberschriften = ($jahresDaten[0] -replace "\[.*?\]|\-" , "") -replace " ", ""
$spaltenUeberschriften = $jahresDaten[0] -replace "ä" , "ae"
$spaltenUeberschriften += ";Gewichtung"
$spaltenUeberschriften | Out-File $dateiMitGewichtung
$spaltenUeberschriften | Out-File $stundenDatei
$spaltenUeberschriften | Out-File $tagesDatei
$spaltenUeberschriften | Out-File $monatsDatei


$vorZeitstempel = $null
foreach ($zeile in $jahresDaten) {
	if ($zeile -notlike "Uhrzeit*"){
		$datensatz = datensatzSplitten -datensatz $zeile -textDatensatz $textDatensatz -vorherigerTextDatensatz $vorherigerTextDatensatz

		if ($datensatz -ne $null) {
			if ($vorZeitstempel -ne $null) {
				if ($datensatz.zeitstempel.hour -ne $vorZeitstempel.Hour) {
					$zwischenZeitstempel = $datensatz.zeitstempel
					$zwischenZeitstempel.Minute = 0
					$zwischenZeitstempel.Second = 0
					$zwischenZeitstempel.Millisecond = 0
					$verhaeltnis1 = [Double]([timediff]($zwischenZeitstempel - $vorZeitstempel).TotalSeconds)/[Double]([timediff]($datensatz.zeitstempel - $vorZeitstempel).TotalSeconds)
					$datensatz2.zeitStempel = $zwischenZeitstempel
					for ($i = 0;$i -lt $datensatz.werte.length-1;$i++) {
						$datensatz2.werte[$i] = $datensatz.werte[$i]
						
				
				$datensatz.werte[9] = [timediff]($datensatz.zeitstempel - $vorZeitstempel).TotalSeconds / 3600.0
			}
			else {
				# für den ersten Datensatz wird angenommen, dass er für 5 Minuten (300 Sekunden) gilt
				$datensatz.werte[9] = 300.0 / 3600.0
			}
			
			$vorZeitstempel = $datensatz.zeitStempel;
			$ausgabeZeile = $datensatz.zeitStempel 
			Out-File $dateiMitGewichtung -Append
		}
	}

}





