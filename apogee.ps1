
while($true)
{

cls
###############
# Definizione database
Import-Module PSSQLite

###############
# Funzioni globali
function Output-Heading {
 param( [int]$direction)
If ($direction -lt 22)  {
Write-Output "N" } 
ElseIf ($direction -lt 67)  {
Write-Output "NE" } 
ElseIf ($direction -lt 112)  {
Write-Output "E" } 
ElseIf ($direction -lt 157)  {
Write-Output "SE" } 
ElseIf ($direction -lt 212)  {
Write-Output "S" } 
ElseIf ($direction -lt 247)  {
Write-Output "SW" } 
ElseIf ($direction -lt 292)  {
Write-Output "W" } 
ElseIf ($direction -lt 337)  {
Write-Output "NW" } 
Else  { Write-Output "N" } 
}


###############
# Variabili calcolo temperatura
$A = 1.1292241*[math]::Pow(10,-3)
$B = 2.341077*[math]::Pow(10,-4)
$C = 8.775468*[math]::Pow(10,-8)
#2018-11-14T13:16:45
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"
$data_status=Get-Date -UFormat "%d-%m-%Y %H:%M:%S"
$unixtime=[Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime()-uformat "%s"))

###############
# Directory di lavoro Adamcmd
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\

###############
# Controllo dimensione file per reset
$file = 'C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee_windspeed.txt'
if ((Get-Item $file).length -gt 40kb) {
    "Data, Temperatura" | Out-File 'apogee_temp.txt'
    "Data,Radiazione solare" | Out-File 'apogee_rad.txt'
    "Data, RPM" | Out-File 'apogee_rpm.txt'
    "Data, Direzione vento" | Out-File 'apogee_winddir.txt'
    "Data, Velocità vento" | Out-File 'apogee_windspeed.txt'
    Write-Host "RESET FILES - RESET FILES - RESET FILES"
}

Write-Host "-------------------------------------"
Write-Host "Dati del $data_status"
"Dati del: <b>$data_status</b>" | Out-File 'apogee_status.txt'
Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# TEMPERATURA
#
$excVin0 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 0 -p COM5)
Write-Host "Lettura Excitation: $excVin0"
"Lettura Excitation: $excVin0 volt"  | Out-File 'apogee_status.txt' -Append

$sensVin1 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 1 -p COM5)
Write-Host "Lettura Sensore: $sensVin1"
"Lettura Sensore: $sensVin1 volt" | Out-File 'apogee_status.txt' -Append

$ResistanceSens1 = 24900*((($excVin0/10000)/($sensVin1/10000))-1)
Write-Host "Resistenza Sens1: $ResistanceSens1"
"Resistenza Sens1: $ResistanceSens1 ohm" | Out-File 'apogee_status.txt' -Append

$TempKelvinSens1 = 1/($A+($B*[math]::Log($ResistanceSens1))+$C*([math]::Pow([math]::Log($ResistanceSens1),3)))
Write-Host "TempKelvin Sens1: $TempKelvinSens1"
"TempKelvin Sens1: $TempKelvinSens1 °K" | Out-File 'apogee_status.txt' -Append

$TempCelsiusSens1 = [math]::Round(($TempKelvinSens1 - 273.15),1)
Write-Host "TempCelsius Sens1: $TempCelsiusSens1 °C"
"<b>Temperatura: $TempCelsiusSens1 °C</b>" | Out-File 'apogee_status.txt' -Append


"$data,$TempCelsiusSens1" | Out-File 'apogee_temp.txt' -Append
# Data per confronto su 192.168.2.205
"$TempCelsiusSens1" | Out-File 'apogee_temp_single.txt'
# Data per gauge su kwos.org
"var temperatura_single = `'$TempCelsiusSens1'`;"  | Out-File 'temperatura_single.txt'

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# RADIAZIONE SOLARE
#
$pyrVin5 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 5 -p COM5)
$pyrVin5_int = [double]::Parse($pyrVin5)
$rad = ($pyrVin5_int*1000)*5
Write-Host "Lettura pyranometer: $pyrVin5"
"Lettura pyranometer: $pyrVin5 mV"  | Out-File 'apogee_status.txt' -Append

Write-Host "Radiazione solare: $rad W/m2"
"<b>Radiazione solare: $rad W/m2</b>" | Out-File 'apogee_status.txt' -Append
"$data,$rad" | Out-File 'apogee_rad.txt' -Append
# Data per confronto su 192.168.2.205
"$rad" | Out-File 'apogee_radiation_single.txt'
# Data per gauge su kwos.org
"var radiazione_single = `'$rad'`;"  | Out-File 'radiazione_single.txt'

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# VELOCITA' VENTOLA 
#
$PORT='COM5'
$BAUDRATE='9600'
$Parity=[System.IO.Ports.Parity]::None
$dataBits=8
$stopBits=[System.IO.Ports.StopBits]::one

$period= [timespan]::FromSeconds($DelaySeconds)
$port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
$port.Open()
$StartTime = Get-Date

$port.WriteLine("#020`r")
$iCountOnPowerShell = 0
while($port.BytesToRead -lt 9) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
    #$port.BytesToRead
}
$countCN0_full = $port.ReadExisting()
$port.Close()
$countCN0 = $countCN0_full -replace '[>]',""
$countCN0_dec = [int]"0x$countCN0"
$rpm = $countCN0_dec*33.8
Write-Host "Lettura counter: $countCN0"
"Lettura counter: $countCN0" | Out-File 'apogee_status.txt' -Append

Write-Host "Velocità ventola: $rpm RPM"
"Velocità ventola: $rpm RPM" | Out-File 'apogee_status.txt' -Append
"$data,$rpm" | Out-File 'apogee_rpm.txt' -Append

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# VELOCITA' - DIREZIONE VENTO 
#

$PORT='COM6'
$BAUDRATE='9600'
$Parity=[System.IO.Ports.Parity]::None
$dataBits=8
$stopBits=[System.IO.Ports.StopBits]::one


$port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
$port.Open()
$StartTime = Get-Date
# :1,238,0.00,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,78

while($port.BytesToRead -lt 69) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
}
$risposta1 = $port.ReadExisting()

#Write-Host "Stringa anemometro: $risposta1"
$risposta1 | foreach-object {
  $dati = $_ -split ","
  "{0} {1} {2}" -f $dati[0],$dati[1],$dati[2]
}

#"Stringa anemometro: $dati[0],$dati[1],$dati[2]" | Out-File 'apogee_status.txt' -Append
$direction = [int]$dati[1]
$Vmps = [double]::Parse($dati[2])
$port.Close()

$Vkmh = $Vmps*3.60
$Vkmh_round =  [math]::Round($Vkmh,2)

# Controllo ERRORI nella lettura
#
If ( $Vkmh_round -gt 100 ) {
    $Vkmh_round = 0.0
    $Vkmh = 0.0
}

Write-Host "Velocità mph: $Vmps mps"
"Velocità mps: $Vmps mps" | Out-File 'apogee_status.txt' -Append
Write-Host "Velocità kmh: $Vkmh_round km/h"
"<b>Velocità kmh:  $Vkmh_round km/h</b>" | Out-File 'apogee_status.txt' -Append

"$data,$Vkmh_round" | Out-File 'apogee_windspeed.txt' -Append

# Data per gauge su kwos.org
"var windspeed_single = `'$Vkmh_round'`;"  | Out-File 'windspeed_single.txt'

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# DIREZIONE VENTO 
#

$heading=Output-Heading -direction $direction

Write-Host "Direzione vento: $direction ° - $heading"
"<b>Direzione vento: $direction ° - $heading</b>" | Out-File 'apogee_status.txt' -Append
"$data,$direction" | Out-File 'apogee_winddir.txt' -Append

# Data per gauge su kwos.org
"var winddir_single = `'$direction'`;"  | Out-File 'winddir_single.txt'


Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
# 
# Setting velocità ventola dipendente da radiazione e vento 
#

Write-Host "Setting velocità ventola:"
"Setting velocità ventola:"  | Out-File 'apogee_status.txt' -Append

If ($rad -lt 100 -And $Vkmh -gt 0) {
    Write-Host "R=$rad < 100 & W=$Vkmh_round > 0"
    Write-Host "Rad. < 100 & Wind > 0"
    "Rad. < 100 & Wind > 0" | Out-File 'apogee_status.txt' -Append
    $min = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 0)
    Write-Host $min  
    $min | Out-File 'apogee_status.txt' -Append
    Start-Sleep -Seconds 2
} ElseIf ($rad -lt 300 -Or ($rad -lt 100 -And $Vkmh -eq 0) )  {
    Write-Host "R=$rad < 300 Or. R=$rad < 100 & W=$Vkmh_round = 0"
    Write-Host "Rad. < 300 Or Rad. < 100 & Wind = 0"
    "Rad. < 300 Or Rad. < 100 & Wind = 0" | Out-File 'apogee_status.txt' -Append
    $med = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 50)
    Write-Host $med
    $med | Out-File 'apogee_status.txt' -Append
    Start-Sleep -Seconds 2
} Else {
    Write-Host "R=$rad > 300"
    Write-Host "Rad. > 300"
    "Rad. > 300" | Out-File 'apogee_status.txt' -Append    
    $max = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 100)
    Write-Host $max
    $max | Out-File 'apogee_status.txt' -Append
    Start-Sleep -Seconds 2
} 

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status.txt' -Append

###############
#
# Scrittura sul DataBase
#
# $query_crea_DB = "CREATE TABLE archive (dateTime INTEGER NOT NULL UNIQUE PRIMARY KEY, 
# outTemp REAL, windSpeed REAL, windDir REAL, rainRate REAL, rain REAL, rpm REAL, radiation REAL)"
# 
# 


$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$query = "INSERT INTO archive (dateTime, outTemp, windSpeed, windDir, radiation, rpm) 
          VALUES ('$unixtime', '$TempCelsiusSens1', '$Vkmh', '$direction', '$rad', '$rpm')"

Invoke-SqliteQuery -DataSource $Database -Query $query

$query_test = "SELECT * FROM archive ORDER BY dateTime DESC LIMIT 1"
Invoke-SqliteQuery -DataSource $Database -Query $query_test

Write-Host "-------------------------------------"

###############
#
# Lettura dal DataBase
# Creazione Massimi, minimi, altro
#
#
$beginoftoday=[Math]::Floor([decimal](Get-Date(Get-Date -Hour 0 -Minute 00).ToUniversalTime()-uformat "%s"))

$temp_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(outTemp) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$temp_min=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(outTemp) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
#
$rad_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(radiation) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
#
$windSpeed_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(windSpeed) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$windSpeed_max_round=[math]::Round($windSpeed_max,2)
#
$WindDir_avg=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(windDir) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$heading_avg=Output-Heading -direction $WindDir_avg
$WindDir_avg_round =  [math]::Round($WindDir_avg,2)

Write-Host "Temperatura MAX oggi: $temp_max °C"
Write-Host "Temperatura MIN oggi: $temp_min °C"
Write-Host "Radiazione MAX oggi: $rad_max W/m2"
Write-Host "WindSpeed MAX oggi: $windSpeed_max_round Km/h"
Write-Host "WindDir AVG oggi: $WindDir_avg_round - $heading_avg"

"<b>Temperatura MAX oggi: $temp_max °C</b>" | Out-File 'apogee_status.txt' -Append
"<b>Temperatura MIN oggi: $temp_min °C</b>" | Out-File 'apogee_status.txt' -Append
"<b>Radiazione MAX oggi: $rad_max W/m2</b>" | Out-File 'apogee_status.txt' -Append
"<b>WindSpeed MAX oggi: $windSpeed_max_round Km/h</b>" | Out-File 'apogee_status.txt' -Append
"<b>WindDir AVG oggi: $WindDir_avg_round - $heading_avg</b>" | Out-File 'apogee_status.txt' -Append

"------------------------------"  | Out-File 'apogee_status.txt' -Append
Write-Host "-------------------------------------"

###############
#
# Invio FTP su KWOS
#

Write-Host "Invio su FTP KWOS"
$server = "ftp.kwos.it"
$filelist = "windspeed_single.txt radiazione_single.txt winddir_single.txt temperatura_single.txt apogee_temp.txt apogee_rad.txt apogee_rpm.txt apogee_status.txt apogee_windspeed.txt apogee_winddir.txt"   
$user = "**********"
$password = "**********"
$dir = "/www.kwos.org/apogee/"

"open $server
user $user $password
binary  
cd $dir     
" +
($filelist.split(' ') | %{ "put ""$_""`n" }) | ftp -i -in

Write-Host "Invio su FTP 192.168.2.205"
#
# Invio FTP su 192.168.2.205 per confronto temperature
#
$server = "192.168.2.205"
$filelist = "apogee_temp_single.txt apogee_radiation_single.txt"   
$user = "**********"
$password = "**********"
$dir = "/root/"

"open $server
user $user $password
binary  
cd $dir     
" +
($filelist.split(' ') | %{ "put ""$_""`n" }) | ftp -i -in



###############
#
Write-Host "-------------------------------------"
Write-Host "Fine loop"
Write-Host "-------------------------------------"
# Fine loop
Start-Sleep -Seconds 300
}