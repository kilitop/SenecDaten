<#
.SYNOPSIS
   <A brief description of the script>
.DESCRIPTION
   <A detailed description of the script>
.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
#>

$credentials = Get-Credential
$credentials.Password | ConvertFrom-SecureString | Set-Content "senecPW.txt"
$credentials | Export-Clixml -Path "senecLogin.xml"


