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

$uri = "https://mein-senec.de/auth/login"
$uriStatistik = "https://mein-senec.de/endkunde/#/0/statistischeDaten"
$uriDownload = "https://mein-senec.de/endkunde/api/statistischeDaten/download"

$datenPfad = "C:\Daten\Anwendungen\PhotovoltaikDaten\"
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
$laufendeWoche = [System.Globalization.DateTimeFormatInfo]::CurrentInfo.Calendar.GetWeekOfYear((get-date),2,1)
"laufende Woche = $laufendeWoche"

$datenPfadDownloads = $datenPfad + "Jahr" + $jahr + "\"
$letzteDateiNochmals = $true 

"downloads beginnen"
for ($i = [int]$laufendeWoche; $i -ge 1; $i--) {
	"Download Daten Woche $i" 
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





