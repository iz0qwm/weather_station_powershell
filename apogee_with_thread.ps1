
# Ferma tutti i Job
Remove-Job *

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

# Setta i cronometri per gli eventi a minuti (FTP e scrittura su DataBase
$timeoutDB = new-timespan -Minutes 5
$timeoutFTP = new-timespan -Minutes 3
$cronometroDB.Stop()
$cronometroFTP.Stop()
$cronometroDB = [diagnostics.stopwatch]::StartNew()
$cronometroFTP = [diagnostics.stopwatch]::StartNew()

# Variabili per accumulo istantaneo pluviometri.
# Setto la variabile lastmem partendo dal valore nel DB
$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$last_acc_pluvio=Invoke-SqliteQuery -DataSource $Database -Query "SELECT rain FROM tampone ORDER BY dateTime DESC limit 1" -As SingleValue
$last_acc_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT raintb FROM tampone ORDER BY dateTime DESC limit 1" -As SingleValue
$last_acc_pluvio_archive=Invoke-SqliteQuery -DataSource $Database -Query "SELECT rain FROM archive ORDER BY dateTime DESC limit 1" -As SingleValue
$last_acc_pluviotb_archive=Invoke-SqliteQuery -DataSource $Database -Query "SELECT raintb FROM archive ORDER BY dateTime DESC limit 1" -As SingleValue

$last_acc_pluvio = [double]::Parse($last_acc_pluvio)
$last_acc_pluviotb = [double]::Parse($last_acc_pluviotb)
$last_acc_pluvio_archive = [double]::Parse($last_acc_pluvio_archive)
$last_acc_pluviotb_archive = [double]::Parse($last_acc_pluviotb_archive)

# Se l'accumulo nella memoria tampone è 0 cerco in archive
If ( $last_acc_pluvio -eq $null -OR $last_acc_pluvio -eq 0.0 )  {
    $lastmem_acc_pluvio = $last_acc_pluvio_archive
} Else {
    $lastmem_acc_pluvio = $last_acc_pluvio
}
If ( $last_acc_pluviotb -eq $null -OR $last_acc_pluviotb -eq 0.0 )  {
    $lastmem_acc_pluviotb = $last_acc_pluviotb_archive
} Else {
    $lastmem_acc_pluviotb = $last_acc_pluviotb
}

$acc_pluvio = 0.0
$acc_pluviotb = 0.0


# INIZIO LOOP
while($true)
{

# Ferma tutti i Job e pulisce la memoria
Remove-Job *
#[System.GC]::Collect()

cls

###############
# Definizione Date e Tempi
#
#2018-11-14T13:16:45
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"
$data_status=Get-Date -UFormat "%d-%m-%Y %H:%M:%S"
$unixtime=[Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime()-uformat "%s"))

###############
# Directory di lavoro Adamcmd
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\

###############
# Controllo orario per cancellazione accumulo pioggia
#$min = Get-Date '01:00'
#$max = Get-Date '01:03'
$min = Get-Date '23:57'
$max = Get-Date '23:59'

$now = Get-Date

if ($min.TimeOfDay -le $now.TimeOfDay -and $max.TimeOfDay -ge $now.TimeOfDay) {
       
    Write-Host "-------------------------------------"
    Write-Host "E' ora di resettare l'accumulo della pioggia"
    Write-Host "-------------------------------------"
    Write-Host " DISDROMETRO "
    $PORT='COM7'
    $BAUDRATE='9600'
    $Parity=[System.IO.Ports.Parity]::None
    $dataBits=8
    $stopBits=[System.IO.Ports.StopBits]::one

    $port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$Parity,$dataBits,$stopBits
    $port.Open()
    $port.WriteLine(">RESET`r`n")
    Start-Sleep -Milliseconds 1000
    $port.Close()

    Write-Host " PLUVIOMETRO TB "
    $PORT='COM5'
    $BAUDRATE='9600'
    $Parity=[System.IO.Ports.Parity]::None
    $dataBits=8
    $stopBits=[System.IO.Ports.StopBits]::one

    $period= [timespan]::FromSeconds($DelaySeconds)
    $port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
    $port.ReadTimeout = 5000
    $port.Open()
    $StartTime = Get-Date

#    $port.WriteLine("`$02510`r")
#    while($port.BytesToRead -lt 3) {
#        $iCountOnPowerShell++
#        Start-Sleep -Milliseconds 1000
        #$port.BytesToRead
#    }
    $port.WriteLine("`$0360`r")
    while($port.BytesToRead -lt 3) {
        $iCountOnPowerShell++
        Start-Sleep -Milliseconds 1000
        #$port.BytesToRead
    }
    $risposta1 = $port.ReadExisting()
#    $port.WriteLine("`$02511`r")
#    while($port.BytesToRead -lt 3) {
#        $iCountOnPowerShell++
#        Start-Sleep -Milliseconds 1000
        #$port.BytesToRead
#    }
    Start-Sleep -Milliseconds 1000
    $port.Close()
}

###############
# Controllo dimensione file per reset
$file = 'C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee_temp.txt'
if ((Get-Item $file).length -gt 80kb) {
    "Data, ST-110, HC2-S3" | Out-File 'apogee_temp.txt'
    "Data, Umidita" | Out-File 'apogee_umidita.txt'
    "Data,Radiazione solare" | Out-File 'apogee_rad.txt'
    "Data, RPM" | Out-File 'apogee_rpm.txt'
    "Data, Direzione vento" | Out-File 'apogee_winddir.txt'
    "Data, Velocità vento" | Out-File 'apogee_windspeed.txt'
    "Data, Pressione" | Out-File 'apogee_pressione.txt'
    "Data, Disdrometro acc., Disdrometro intens., Pluviom. acc., Pluviom. intens." | Out-File 'apogee_pluviometro.txt'
    Write-Host "RESET FILES - RESET FILES - RESET FILES"
}

Write-Host "-------------------------------------"
Write-Host "Dati del $data_status"
"Dati del: <b>$data_status</b>" | Out-File 'apogee_status_PRE.txt'
Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_PRE.txt' -Append


######################################################
# THREAD PER COM5

$ricevi_da_COM5 = {

Write-Host "=========================="
Write-Host "JOB: RICEVI DA COM5"
Write-Host "=========================="

#2018-11-14T13:16:45
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"
$data_status=Get-Date -UFormat "%d-%m-%Y %H:%M:%S"
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\

###############
#
# TEMPERATURA
#
# Variabili calcolo temperatura
$A = 1.1292241*[math]::Pow(10,-3)
$B = 2.341077*[math]::Pow(10,-4)
$C = 8.775468*[math]::Pow(10,-8)

Write-Host " "
Write-Host "-------------------------------------"
$excVin0 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 0 -p COM5)
Write-Host "Lettura Excitation: $excVin0"
"Lettura Excitation: $excVin0 volt"  | Out-File 'apogee_status_COM5.txt'

$sensVin1 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 1 -p COM5)
Write-Host "Lettura Sensore: $sensVin1"
"Lettura Sensore: $sensVin1 volt" | Out-File 'apogee_status_COM5.txt' -Append

$ResistanceSens1 = 24900*((($excVin0/10000)/($sensVin1/10000))-1)
Write-Host "Resistenza Sens1: $ResistanceSens1"
"Resistenza Sens1: $ResistanceSens1 ohm" | Out-File 'apogee_status_COM5.txt' -Append

$TempKelvinSens1 = 1/($A+($B*[math]::Log($ResistanceSens1))+$C*([math]::Pow([math]::Log($ResistanceSens1),3)))
Write-Host "TempKelvin Sens1: $TempKelvinSens1"
"TempKelvin Sens1: $TempKelvinSens1 &#176;K" | Out-File 'apogee_status_COM5.txt' -Append

$TempCelsiusSens1 = [math]::Round(($TempKelvinSens1 - 273.15),1)
Write-Host "TempCelsius Sens1: $TempCelsiusSens1 °C"
"<b>Temperatura ST-110: $TempCelsiusSens1 &#176;C</b>" | Out-File 'apogee_status_COM5.txt' -Append


Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append
###############
#
# PRESSIONE ATMOSFERICA
#
# Costanti
$hpamv=0.218
$hpaoffset=118
$altitude=284
$constaltitude=0.0065
$basekelvin=273.15

# Lettura valore precedente del QNH
Import-Module PSSQLite
$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$QNH_round_last=Invoke-SqliteQuery -DataSource $Database -Query "SELECT qnh FROM tampone WHERE dateTime = (SELECT MAX(dateTime)  FROM tampone)" -As SingleValue
$QNH_round_last_archive=Invoke-SqliteQuery -DataSource $Database -Query "SELECT qnh FROM archive WHERE dateTime = (SELECT MAX(dateTime)  FROM archive)" -As SingleValue

#

$preVin2 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 2 -p COM5)
Write-Host "Lettura barometro: $preVin2"
"Lettura barometro: $preVin2 V"  | Out-File 'apogee_status_COM5.txt' -Append
$preVin2_int = [double]::Parse($preVin2)*1000
$pressione = ($preVin2_int*$hpamv)+$hpaoffset
Write-Host "Pressione: $pressione hPa"
"Pressione: $pressione hPa" | Out-File 'apogee_status_COM5.txt' -Append

$sotto=[double]::Parse($TempCelsiusSens1)+([double]::Parse($constaltitude)*$altitude)+[double]::Parse($basekelvin)
$sopra=([double]::Parse($constaltitude)*$altitude)
$sottosopra=[double]::Parse($sopra)/[double]::Parse($sotto)
$parentesi=[math]::Pow((1-[double]::Parse($sottosopra)),-5.257)
$QNH=$pressione*[double]::Parse($parentesi)
$QNH_round=[math]::Round($QNH,2)

Write-Host "QNH_letto: $QNH_round"
Write-Host "QNH_last: $QNH_round_last"

If ( $QNH_round_last -eq $null -OR $QNH_round_last -eq 0.0 )  {
    $QNH_round_last=$QNH_round_last_archive
}
If ( $QNH_round_last_archive -eq $null -OR $QNH_round_last_archive -eq 0.0 )  {
    $QNH_round_last=$QNH_round
}
$diff=[math]::abs($QNH_round-$QNH_round_last)
If ( $diff -gt 5 ) {
    Write-Host "QNH_letto-QNH_last = $diff"
    Write-Host "Errore nella lettura QNH uso QNH_last: $QNH_round_last"
    $QNH_round=$QNH_round_last
} Else {
    Write-Host "QNH_letto-QNH_last = $diff"
    Write-Host "Lettura QNH OK uso QNH_letto: $QNH_round"
}


Write-Host "QNH: $QNH_round hPa"
"<b>QNH: $QNH_round hPa</b>" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

# pressione usata ora setto pressione_last
$global:pressione_last=$pressione
Write-Host "Pressione_last_settata=$pressione_last"
###############
#
# RADIAZIONE SOLARE
#
$pyrVin5 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 5 -p COM5)
$pyrVin5_int = [double]::Parse($pyrVin5)
$rad = ($pyrVin5_int*1000)*5
Write-Host "Lettura pyranometer: $pyrVin5"
"Lettura pyranometer: $pyrVin5 V"  | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "Radiazione solare: $rad W/m2"
"<b>Radiazione solare: $rad W/m2</b>" | Out-File 'apogee_status_COM5.txt' -Append
"$data,$rad" | Out-File 'apogee_rad.txt' -Append
# Data per confronto su 192.168.2.205
"$rad" | Out-File 'apogee_radiation_single.txt'
# Data per gauge su kwos.org
"var radiazione_single = `'$rad'`;"  | Out-File 'radiazione_single.txt'

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

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
$port.ReadTimeout = 5000
$port.Open()
$StartTime = Get-Date

$port.WriteLine("`$0260`r")
while($port.BytesToRead -lt 3) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
    #$port.BytesToRead
}
$risposta1 = $port.ReadExisting()
$port.WriteLine("`$02501`r")
while($port.BytesToRead -lt 3) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
    #$port.BytesToRead
}
$risposta2 = $port.ReadExisting()
Start-Sleep -Seconds 5
$port.WriteLine("#020`r")
$iCountOnPowerShell = 0
while($port.BytesToRead -lt 9) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
    #$port.BytesToRead
}
$port.WriteLine("`$02500`r")
$countCN0_full = $port.ReadExisting()
$port.Close()
$countCN0 = $countCN0_full -replace '[>]',""
$countCN0_dec = [int]"0x$countCN0"
$rpm = $countCN0_dec*5.5
Write-Host "Lettura counter: $countCN0 - $countCN0_dec"
"Lettura counter: $countCN0_dec" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "Velocità ventola: $rpm RPM"
"Velocit&#225; ventola: $rpm RPM" | Out-File 'apogee_status_COM5.txt' -Append


Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

###############
#
# PLUVIOMETRO TB 
#
$PORT='COM5'
$BAUDRATE='9600'
$Parity=[System.IO.Ports.Parity]::None
$dataBits=8
$stopBits=[System.IO.Ports.StopBits]::one

$period= [timespan]::FromSeconds($DelaySeconds)
$port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
$port.ReadTimeout = 5000
$port.Open()
$StartTime = Get-Date

$port.WriteLine("#030`r")
$iCountOnPowerShell = 0
while($port.BytesToRead -lt 9) {
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 1000
    #$port.BytesToRead
}
$countCN1_full = $port.ReadExisting()
$port.Close()
$countCN1 = $countCN1_full -replace '[>]',""
$countCN1_dec = [int]"0x$countCN1"

$acc_pluviotb = $countCN1_dec*0.1

Write-Host "Lettura counter: $countCN1 - $countCN1_dec"
"Lettura counter: $countCN1_dec" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "Accumulo TB: $acc_pluviotb mm"
"<b>Accumulo TB: $acc_pluviotb mm</b>" | Out-File 'apogee_status_COM5.txt' -Append


###############
# Calcolo intensità pluvio TB in mm/min
# Leggo l'ultimo valore di accumulo scritto un minuto prima nel tampone o in archive
Import-Module PSSQLite
$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$last_acc_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT raintb FROM tampone ORDER BY dateTime DESC LIMIT 1" -As SingleValue
If ( !$last_acc_pluviotb ) {
    $last_acc_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT raintb FROM archive ORDER BY dateTime DESC LIMIT 1" -As SingleValue
}
$last_intens_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT rainRatetb FROM archive ORDER BY dateTime DESC LIMIT 1" -As SingleValue


$acc_pluviotb = [double]::Parse($acc_pluviotb)
$last_acc_pluviotb = [double]::Parse($last_acc_pluviotb)

#rain_rate = delta_rain + rain_rate_prec * 59/60
#delta_rain = acc_pluviotb dello scorso minuto
#rain_rate_prec = last rain_rate

$intens_pluviotb_min = $acc_pluviotb-$last_acc_pluviotb
$intens_pluviotb = ($intens_pluviotb_min+$last_intens_pluviotb)*0.99

If ( $intens_pluviotb -lt 0.0 ) {
    $intens_pluviotb = 0.0
}


Write-Host "Intensità: $intens_pluviotb mm/h"
"<b>Intensit&#225;:  $intens_pluviotb mm/h</b>" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

#
# TEMPERATURA ROTRONIC 
#

$sensVin3 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 3 -p COM5)
$TempCelsiusSens2 = [math]::Round(([double]::Parse($sensVin3)*100)-40,1)

$difftemp=[math]::abs($TempCelsiusSens2-$TempCelsiusSens1)
If ( $difftemp -gt 4 ) {
    Write-Host "TempSens2-TempSens1 = $difftemp"
    Write-Host "Lettura TempCelsiusSens2 ERRORE: $sensVin3 V - $TempCelsiusSens2 °C"  
    $TempCelsiusSens2=$TempCelsiusSens1
} Else {
    Write-Host "TempSens2-TempSens1 = $difftemp"
    Write-Host "Lettura TempCelsiusSens2 OK: $sensVin3 V - $TempCelsiusSens2"
}

Write-Host "TempCelsiusSens2: $TempCelsiusSens2 °C"
"Lettura Sensore: $sensVin3 volt" | Out-File 'apogee_status_COM5.txt' -Append
"<b>Temperatura HC2-S3: $TempCelsiusSens2 &#176;C</b>" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

#
# UMIDITA' ROTRONIC 
#

$sensVin4 = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\AdamCmd.exe read-ai -u 1 -c 4 -p COM5)


$HumSens = [double]::Parse($sensVin4)*100

# Taratura igrometro
#########################
$HumSens = $HumSens + 1.5
#

$Hum=[math]::Round($HumSens,1)

Write-Host "Lettura umidità: $sensVin4"
"Lettura umidit&#225;: $sensVin4 volt"  | Out-File 'apogee_status_COM5.txt' -Append
Write-Host "Umidità: $Hum %"
"<b>Umidit&#225;:  $Hum %</b>" | Out-File 'apogee_status_COM5.txt' -Append

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM5.txt' -Append

#OUTPUT VARIABILI
#
Write-Output $QNH_round,$TempCelsiusSens1,$rad,$rpm,$acc_pluviotb,$intens_pluviotb,$Hum,$TempCelsiusSens2

}
################################
# FINE THREAD COM5



######################################################
# THREAD PER COM6

$ricevi_da_COM6 = {

Write-Host "=========================="
Write-Host "JOB: RICEVI DA COM6"
Write-Host "=========================="

#2018-11-14T13:16:45
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"
$data_status=Get-Date -UFormat "%d-%m-%Y %H:%M:%S"
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\


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
#
# VELOCITA' - DIREZIONE VENTO 
#

$PORT='COM6'
$BAUDRATE='9600'
$Parity=[System.IO.Ports.Parity]::None
$dataBits=8
$stopBits=[System.IO.Ports.StopBits]::one
$DelaySeconds=10
$iCountOnPowerShell=0


$port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
$port.ReadTimeout = 5000
$port.Open()
$StartTime = Get-Date
# :1,238,0.00,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,78


while($port.BytesToRead -lt 69) {
    if ( $iCountOnPowerShell -gt 5 ) {
        break
    }
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 500
}
$risposta = $port.ReadExisting()

If ( $risposta -eq "" ) {
    $risposta = ":1,0,0.00,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,ER"
} 

Write-Host "Stringa anemometro: $risposta"
$risposta1 = $risposta -creplace '^[^:1]*', ''
Write-Host "Stringa pulita: $risposta1"
$risposta1 | foreach-object {
  $dati = $_ -split ","
  "{0} {1} {2}" -f $dati[0],$dati[1],$dati[2]
}

#"Stringa anemometro: $dati[0],$dati[1],$dati[2]" | Out-File 'apogee_status.txt' -Append
$direction = [int]$dati[1]
$Vmps = [double]::Parse($dati[2])
$port.Close()

# Controllo ERRORI nella lettura
#
If ( $Vmps -match "/" ) {
    $Vmps = 0.0
}
If ( $direction -match "/" ) {
    $direction = 0
}

$Vkmh = $Vmps*3.60
$Vkmh_round =  [math]::Round($Vkmh,2)

# Controllo ERRORI nella lettura
#
If ( $Vkmh_round -gt 100 ) {
    $Vkmh_round = 0.0
    $Vkmh = 0.0
}


Write-Host "Velocità mph: $Vmps mps"
"Velocit&#225; mps: $Vmps mps" | Out-File 'apogee_status_COM6.txt'
Write-Host "Velocità kmh: $Vkmh_round km/h"
"<b>Velocit&#225; kmh:  $Vkmh_round km/h</b>" | Out-File 'apogee_status_COM6.txt' -Append



Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM6.txt' -Append

###############
#
# DIREZIONE VENTO 
#

$heading=Output-Heading -direction $direction

Write-Host "Direzione vento: $direction ° - $heading"
"<b>Direzione vento: $direction &#176; - $heading</b>" | Out-File 'apogee_status_COM6.txt' -Append



Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM6.txt' -Append

#OUTPUT VARIABILI
#
Write-Output $Vkmh_round,$direction

}
################################
# FINE THREAD COM6


######################################################
# THREAD PER COM7

$ricevi_da_COM7 = {

Write-Host "=========================="
Write-Host "JOB: RICEVI DA COM7"
Write-Host "=========================="

#2018-11-14T13:16:45
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"
$data_status=Get-Date -UFormat "%d-%m-%Y %H:%M:%S"
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\

###############
#
# DISDROMETRO 
#

$PORT='COM7'
$BAUDRATE='9600'
$Parity=[System.IO.Ports.Parity]::None
$dataBits=8
$stopBits=[System.IO.Ports.StopBits]::one
$iCountOnPowerShell=0

$port = new-Object System.IO.Ports.SerialPort $PORT,$BAUDRATE,$PArity,$dataBits,$stopBits
$port.ReadTimeout = 5000
$port.Open()
$StartTime = Get-Date
# :1,/,/,/,/,/,/,0,0.0,0.0,0,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,B5

while($port.BytesToRead -lt 138) {
    if ( $iCountOnPowerShell -gt 5 ) {
        break
    }
    $iCountOnPowerShell++
    Start-Sleep -Milliseconds 500
}
$risposta = $port.ReadExisting()

If ( $risposta -eq "" ) {
    $risposta = ":1,/,/,/,/,/,/,0,0.0,0.0,0,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,/,ER"
} 

Write-Host "Stringa pluviometro: $risposta"
$risposta1 = $risposta -creplace '^[^:1]*', ''
Write-Host "Stringa pulita: $risposta1"
$risposta1 | foreach-object {
  $dati = $_ -split ","
  "{0} {1} {2} {2} {3} {4} {5} {6} {7} {8} {9} {10}" -f $dati[0],$dati[1],$dati[2],$dati[3],$dati[4],$dati[5],$dati[6],$dati[7],$dati[8],$dati[9],$dati[10]
}


$stato_pluvio = [int]$dati[7] # 4=Solid / 2=snow / 1=rain
$intens_pluvio = ([double]::Parse($dati[8]))/10 # mm/h
$acc_pluvio = ([double]::Parse($dati[9]))/10   # To clear >CLR RFS\n\r
$unit_pluvio = [int]$dati[10]

If ( $acc_pluvio -match "\?" ) {
    $acc_pluvio = 0.0
    $intens_pluvio = 0.0
}

If ( $intens_pluvio -match "/" ) {
    $acc_pluvio = 0.0
    $intens_pluvio = 0.0
}

$port.Close()



If ( $stato_pluvio -eq 4 ) {
    $tipo_pluvio = "grandine" } 
ElseIf ( $stato_pluvio -eq 2 ) {
    $tipo_pluvio = "neve" }
ElseIf ( $stato_pluvio -eq 1 ) {
    $tipo_pluvio = "pioggia" }
Else { $tipo_pluvio = "nessuna" }

Write-Host "Tipo precipitazione: $tipo_pluvio"
"<b>Tipo precipitazione: $tipo_pluvio</b>" | Out-File 'apogee_status_COM7.txt'
Write-Host "Intensità: $intens_pluvio mm/h"
"<b>Intensit&#225;: $intens_pluvio mm/h</b>" | Out-File 'apogee_status_COM7.txt' -Append
Write-Host "Accumulo: $acc_pluvio mm"
"<b>Accumulo: $acc_pluvio mm</b>" | Out-File 'apogee_status_COM7.txt' -Append



Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_COM7.txt' -Append



#OUTPUT VARIABILI
#
Write-Output $tipo_pluvio

}
################################
# FINE THREAD COM7
################################





# PARTENZA MULTI-THREAD

Start-Job -scriptblock $ricevi_da_COM5 -Name "ricevi_da_COM5"
Start-Job -scriptblock $ricevi_da_COM6 -Name "ricevi_da_COM6"
Start-Job -scriptblock $ricevi_da_COM7 -Name "ricevi_da_COM7"

Get-Job | Wait-Job -Timeout 60

$ricevi_dai_Jobs = Get-Job | Receive-Job 
#$ricevi_da_COM6 = Get-Job | Receive-Job

#$output_COM5 = Get-Job -Name 'ricevi_da_COM5' | Receive-Job
Write-Host "========================================================="
Write-Host "OUTPUT-OUTPUT JOBS: $ricevi_dai_Jobs"
Write-Host "========================================================="

$QNH_round, $TempCelsiusSens1, $rad, $rpm, $acc_pluviotb, $intens_pluviotb, $Hum, $TempCelsiusSens2, $null, $null, $null, $Vkmh_round, $direction, $null, $null, $null, $null, $null, $null, $null, $null, $stato_pluvio, $intens_pluvio, $acc_pluvio, $null, $tipo_pluvio = $ricevi_dai_Jobs -split ' '

# Ripristino variabili
$QNH_round = [double]::Parse($QNH_round)
$TempCelsiusSens1 = [double]::Parse($TempCelsiusSens1)
$rad = [double]::Parse($rad)
$rpm = [double]::Parse($rpm)
$acc_pluviotb = [double]::Parse($acc_pluviotb)
$intens_pluviotb = [double]::Parse($intens_pluviotb)
$Hum = [math]::Round([double]::Parse($Hum),1)
$TempCelsiusSens2 = [double]::Parse($TempCelsiusSens2)
$direction = [int]$direction
$Vkmh_round = [double]::Parse($Vkmh_round)
$stato_pluvio = [int]$stato_pluvio # 4=Solid / 2=snow / 1=rain
$intens_pluvio = ([double]::Parse($intens_pluvio))/10 # mm/h
$acc_pluvio = ([double]::Parse($acc_pluvio))/10   # To clear >CLR RFS\n\r
$tipo_pluvio = $tipo_pluvio

# Controllo dati per esclusione scrittura su DB - Possibile errore nella lettura della stringa via seriale
If ( $intens_pluvio -match "/" -Or $intens_pluvio -match ":" ) {
    $QNH_round = "N"
    $TempCelsiusSens1 = "N"
    $TempCelsiusSens2 = "N"
    $rad = "N"
    $rpm = "N"
    $acc_pluviotb = "N"
    $intens_pluviotb = "N"
    $direction = "N"
    $Vkmh_round = "N"
    $stato_pluvio = "N"
    $intens_pluvio = "N"
    $acc_pluvio = "N"
    $Hum = "N"
}
If ( $acc_pluvio -match "/" -Or $acc_pluvio -match ":" ) {
    $QNH_round = "N"
    $TempCelsiusSens1 = "N"
    $TempCelsiusSens2 = "N"
    $rad = "N"
    $rpm = "N"
    $acc_pluviotb = "N"
    $intens_pluviotb = "N"
    $direction = "N"
    $Vkmh_round = "N"
    $stato_pluvio = "N"
    $intens_pluvio = "N"
    $acc_pluvio = "N"
    $Hum = "N"
}
If ( $Vkmh_round -match "/" -Or $Vkmh_round -match ":" ) {
    $QNH_round = "N"
    $TempCelsiusSens1 = "N"
    $TempCelsiusSens2 = "N"
    $rad = "N"
    $rpm = "N"
    $acc_pluviotb = "N"
    $intens_pluviotb = "N"
    $direction = "N"
    $Vkmh_round = "N"
    $stato_pluvio = "N"
    $intens_pluvio = "N"
    $acc_pluvio = "N"
    $Hum = "N"
}



Remove-Job ricevi_da_COM5
Remove-Job ricevi_da_COM6
Remove-Job ricevi_da_COM7

# FINE MULTI-THREAD
################################

# SCRITTURA FILES PER GRAFICI E GAUGES
#########################################
$data=Get-Date -UFormat "%Y-%m-%dT%H:%M:%S"

# TEMPERATURA
"$data,$TempCelsiusSens1,$TempCelsiusSens2" | Out-File 'apogee_temp.txt' -Append
# Data per confronto su 192.168.2.205
"$TempCelsiusSens1" | Out-File 'apogee_temp_single.txt'
# Data per gauge su kwos.org
"var temperatura_single = `'$TempCelsiusSens1'`;"  | Out-File 'temperatura_single.txt'

# TEMPERATURA ROTRONIC
# Data per confronto su 192.168.2.205
"$TempCelsiusSens2" | Out-File 'apogee_temp2_single.txt'
# Data per gauge su kwos.org
"var temperatura2_single = `'$TempCelsiusSens2'`;"  | Out-File 'temperatura2_single.txt'

# DISDROMETRO
"$data,$acc_pluvio,$intens_pluvio,$acc_pluviotb,$intens_pluviotb" | Out-File 'apogee_pluviometro.txt' -Append
# Data per gauge su kwos.org
"var pluviometro_single = `'$acc_pluvio'`;"  | Out-File 'pluviometro_single.txt'

# PLUVIOMETRO TB
#"$data,$acc_pluviotb,$intens_pluviotb" | Out-File 'apogee_pluviometrotb.txt' -Append
# Data per gauge su kwos.org
"var pluviometrotb_single = `'$acc_pluviotb'`;"  | Out-File 'pluviometrotb_single.txt'

# VENTO
"$data,$direction" | Out-File 'apogee_winddir.txt' -Append
# Data per gauge su kwos.org
"var winddir_single = `'$direction'`;"  | Out-File 'winddir_single.txt'
"$data,$Vkmh_round" | Out-File 'apogee_windspeed.txt' -Append
# Data per gauge su kwos.org
"var windspeed_single = `'$Vkmh_round'`;"  | Out-File 'windspeed_single.txt'

# VENTOLA
"$data,$rpm" | Out-File 'apogee_rpm.txt' -Append

# PRESSIONE
# Data per grafico su kwos.org
"$data,$QNH_round" | Out-File 'apogee_pressione.txt' -Append
# Data per gauge su kwos.org
"var pressione_single = `'$QNH_round'`;"  | Out-File 'pressione_single.txt'

# UMIDITA'
# Data per grafico su kwos.org
"$data,$Hum" | Out-File 'apogee_umidita.txt' -Append
# Data per gauge su kwos.org
"var umidita_single = `'$Hum'`;"  | Out-File 'umidita_single.txt'

# SCRITTURA FILE PER TCP SERVER
#########################################
# Calcolo pioggia istantanea da cumulata per TCP server


$acc_pluvio = [double]::Parse($acc_pluvio)
$acc_pluviotb = [double]::Parse($acc_pluviotb)
$lastmem_acc_pluvio = [double]::Parse($lastmem_acc_pluvio)
$lastmem_acc_pluviotb = [double]::Parse($lastmem_acc_pluviotb)

# Ora calcolo istantaneo sottraendo attuale da DB tampone se non è 0

$inst_acc_pluvio = $acc_pluvio - $lastmem_acc_pluvio
$inst_acc_pluviotb = $acc_pluviotb - $lastmem_acc_pluviotb

if ( $inst_acc_pluvio -le 0 ) {
    $inst_acc_pluvio = 0.0
}
if ( $inst_acc_pluviotb -le 0 ) {
    $inst_acc_pluviotb = 0.0
}

$inst_acc_pluvio = [double]::Parse($inst_acc_pluvio)
$inst_acc_pluviotb = [double]::Parse($inst_acc_pluviotb)

#
Write-Host "========================================================="
Write-Host "                 SCRITTURA FILE PER TCP SERVER"
Write-Host " "
Write-Host "DIS: rainacc=$acc_pluvio - lastmem_acc=$lastmem_acc_pluvio - inst_acc_pluvio=$inst_acc_pluvio"
Write-Host "-TB: rainacc=$acc_pluviotb - lastmem_acc=$lastmem_acc_pluviotb - inst_acc_pluvio=$inst_acc_pluviotb"
Write-Host "outTemp=$TempCelsiusSens1,outHumidity=$Hum,inTemp=$TempCelsiusSens2,barometer=$QNH_round,rain=$inst_acc_pluvio,rainRate=$intens_pluvio,windDir=$direction,windSpeed=$Vkmh_round,radiation=$rad,hail=$inst_acc_pluviotb,hailRate=$intens_pluviotb,consBatteryVoltage=$rpm"
Write-Host "========================================================="
#outTemp=14.8,outHumidity=78,inTemp=14.8,barometer=1010.4,rain=0.2,rainRate=1.0,windDir=140,windSpeed=6,radiation=734,hail=0.3,hailRate=0.1,consBatteryVoltage=300
"outTemp=$TempCelsiusSens1,outHumidity=$Hum,inTemp=$TempCelsiusSens2,barometer=$QNH_round,rain=$inst_acc_pluvio,rainRate=$intens_pluvio,windDir=$direction,windSpeed=$Vkmh_round,radiation=$rad,hail=$inst_acc_pluviotb,hailRate=$intens_pluviotb,consBatteryVoltage=$rpm"  | Out-File 'misure_server.txt'

# lastmem diventa il nuovo valore
$lastmem_acc_pluvio = $acc_pluvio 
$lastmem_acc_pluviotb = $acc_pluviotb
###############
# Scrittura su DataBase tampone per MAX e MIN durante i 5 min

If ($intens_pluvio -eq "N" -OR $Vkmh_round -eq "N" -OR $acc_pluvio -eq "N") {
  # Non scrivo nulla   
} Else {
$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$query_tampone = "INSERT INTO tampone (dateTime, outTemp, windSpeed, windDir, rain, rainRate, rainType, radiation, rpm, qnh, raintb, rainRatetb, hum, outTemp2) 
          VALUES ('$unixtime', '$TempCelsiusSens1', '$Vkmh_round', '$direction', '$acc_pluvio', '$intens_pluvio', '$tipo_pluvio', '$rad', '$rpm', '$QNH_round', '$acc_pluviotb', '$intens_pluviotb', '$Hum', '$TempCelsiusSens2')"
Invoke-SqliteQuery -DataSource $Database -Query $query_tampone
}

###############
# 
# Setting velocità ventola dipendente da radiazione e vento 
#
Write-Host "Setting velocità ventola:"
"Setting velocit&#225; ventola:"  | Out-File 'apogee_status_END.txt'

# Prendo il valore massimo del vento nel tampone per non fare attacca e stacca alla ventola
$Vkmh_round=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(windSpeed) FROM tampone" -As SingleValue
$last_rad=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(radiation) FROM tampone" -As SingleValue

If ($rad -le 500 -And $Vkmh_round -gt 0) {
    If ($last_rad -gt 500) {
        Write-Host "R=$rad < 500 & LastR=$last_rad"
        Write-Host "Rad. < 500 & Wind > 0 but LastRad > 500"
        "Rad. < 500 & Wind > 0 but LastRad > 500" | Out-File 'apogee_status_END.txt' -Append
        $min = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 50)
        Write-Host $min  
        $min | Out-File 'apogee_status_END.txt' -Append
        Start-Sleep -Seconds 2
    } ElseIf ($last_rad -le 500 -And $last_rad -gt 400) {
        Write-Host "R=$rad < 500 & LastR=$last_rad"
        Write-Host "Rad. < 500 & Wind > 0 but 500 < LastRad > 400"
        "Rad. < 500 & Wind > 0 but 500 < LastRad > 400" | Out-File 'apogee_status_END.txt' -Append
        $min = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 40)
        Write-Host $min  
        $min | Out-File 'apogee_status_END.txt' -Append
        Start-Sleep -Seconds 2
    } ElseIf ($last_rad -le 400 -And $last_rad -gt 300) {
        Write-Host "R=$rad < 400 & LastR=$last_rad"
        Write-Host "Rad. < 500 & Wind > 0 but 400 < LastRad > 300"
        "Rad. < 500 & Wind > 0 but 400 < LastRad > 300" | Out-File 'apogee_status_END.txt' -Append
        $min = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 30)
        Write-Host $min  
        $min | Out-File 'apogee_status_END.txt' -Append
        Start-Sleep -Seconds 2
    } Else {
        Write-Host "R=$rad < 500 & W=$Vkmh_round > 0"
        Write-Host "Rad. < 500 & Wind > 0"
        "Rad. < 500 & Wind > 0" | Out-File 'apogee_status_END.txt' -Append
        $min = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 0)
        Write-Host $min  
        $min | Out-File 'apogee_status_END.txt' -Append
        Start-Sleep -Seconds 2
    }
#} ElseIf ($rad -lt 500 -Or ($rad -lt 100 -And $Vkmh_round -eq 0) )  {
} ElseIf ( ($rad -le 500 -And $Vkmh_round -eq 0) -Or ($rad -le 500 -And $Vkmh_round -eq 0 -And $direction -eq 0) )  {
    Write-Host "Rad. < 500 & Wind = 0"
    Write-Host "R=$rad < 500 & W=$Vkmh_round = 0"
    "Rad. < 500 & Wind = 0" | Out-File 'apogee_status_END.txt' -Append
    $med = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 30)
    Write-Host $med
    $med | Out-File 'apogee_status_END.txt' -Append
    Start-Sleep -Seconds 2
} ElseIf ($rad -gt 500 -And $rad -le 800) {
    Write-Host "500 < Rad. > 800"
    Write-Host "500 < R=$rad > 800"
    "500 < Rad. > 800" | Out-File 'apogee_status_END.txt' -Append
    $med = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 50)
    Write-Host $med
    $med | Out-File 'apogee_status_END.txt' -Append
    Start-Sleep -Seconds 2
} Else {
    Write-Host "Rad. > 800"
    Write-Host "R=$rad > 800"
    "Rad. > 800" | Out-File 'apogee_status_END.txt' -Append    
    $max = $(C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\YPwmOutput.exe APOGEE-RPM set_dutyCycle 100)
    Write-Host $max
    $max | Out-File 'apogee_status_END.txt' -Append
    Start-Sleep -Seconds 2
} 

Write-Host "-------------------------------------"
"------------------------------"  | Out-File 'apogee_status_END.txt' -Append

###############
#
# Lettura dal DataBase
# Creazione Massimi, minimi, altro
#
#
$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"
$beginoftoday=[Math]::Floor([decimal](Get-Date(Get-Date -Hour 0 -Minute 00).ToUniversalTime()-uformat "%s"))

$temp_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(outTemp) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$temp_min=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(outTemp) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$temp2_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(outTemp2) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$temp2_min=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(outTemp2) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
#
$hum_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(hum) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$hum_min=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(hum) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
#
$rad_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(radiation) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
#
$windSpeed_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(windSpeed) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$windSpeed_max_round=[math]::Round($windSpeed_max,2)
#
$WindDir_avg=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(windDir) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$heading_avg=Output-Heading -direction $WindDir_avg
$WindDir_avg_round =  [math]::Round($WindDir_avg,2)
#
$intens_pluvio_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(rainRate) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$acc_pluvio_today=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(rain) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$tipo_pluvio_today=Invoke-SqliteQuery -DataSource $Database -Query "SELECT DISTINCT rainType FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$intens_pluviotb_max=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(rainRatetb) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue
$acc_pluviotb_today=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(raintb) FROM archive WHERE dateTime>$beginoftoday AND dateTime<=$unixtime" -As SingleValue

Write-Host "Temp. ST-110 MAX oggi: $temp_max °C"
Write-Host "Temp. ST-110 MIN oggi: $temp_min °C"
Write-Host "Temp. HC2-S3 MAX oggi: $temp2_max °C"
Write-Host "Temp. HC2-S3 MIN oggi: $temp2_min °C"
Write-Host "Umidita MAX oggi: $hum_max %"
Write-Host "Umidita MIN oggi: $hum_min %"
Write-Host "Radiazione MAX oggi: $rad_max W/m2"
Write-Host "WindSpeed MAX oggi: $windSpeed_max_round Km/h"
Write-Host "WindDir AVG oggi: $WindDir_avg_round - $heading_avg"
Write-Host "Rain Rate MAX oggi: $intens_pluvio_max mm/h"
Write-Host "Rain accum. oggi: $acc_pluvio_today mm"
Write-Host "Tipo precip. oggi: $tipo_pluvio_today"
Write-Host "Rain Rate TB MAX oggi: $intens_pluviotb_max mm/h"
Write-Host "Rain accum. TB oggi: $acc_pluviotb_today mm"

"<b>Temp. ST-110 MAX oggi: $temp_max &#176;C</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Temp. ST-110 MIN oggi: $temp_min &#176;C</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Temp. HC2-S3 MAX oggi: $temp2_max &#176;C</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Temp. HC2-S3 MIN oggi: $temp2_min &#176;C</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Umidit&#225; MAX oggi: $hum_max &#37;</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Umidit&#225; MIN oggi: $hum_min &#37;</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Radiazione MAX oggi: $rad_max W/m2</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>WindSpeed MAX oggi: $windSpeed_max_round Km/h</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>WindDir AVG oggi: $WindDir_avg_round - $heading_avg</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Rain Rate MAX oggi: $intens_pluvio_max mm/h</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Rain accum. oggi: $acc_pluvio_today mm</b>" | Out-File 'apogee_status_END.txt' -Append 
"<b>Tipo precip. oggi: $tipo_pluvio_today</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Rain Rate TB MAX oggi: $intens_pluviotb_max mm/h</b>" | Out-File 'apogee_status_END.txt' -Append
"<b>Rain accum. TB oggi: $acc_pluviotb_today mm</b>" | Out-File 'apogee_status_END.txt' -Append 

"------------------------------"  | Out-File 'apogee_status_END.txt' -Append
Write-Host "-------------------------------------"



###############
# Unione file status

Remove-Item apogee_status.txt
add-content apogee_status.txt -value (get-content apogee_status_PRE.txt, apogee_status_COM5.txt, apogee_status_COM6.txt, apogee_status_COM7.txt, apogee_status_END.txt )
#Get-Content apogee_status_PRE.txt, apogee_status_COM5.txt, apogee_status_COM6.txt, apogee_status_END.txt| Set-Content apogee_status.txt




################
# Evento ogni 5 minuti - SCRITTURA SUL DATABASE
#
If ($cronometroDB.elapsed -gt $timeoutDB){

###############
#
# Scrittura sul DataBase
#
# $query_crea_DB = "CREATE TABLE archive (dateTime INTEGER NOT NULL UNIQUE PRIMARY KEY, 
# outTemp REAL, windSpeed REAL, windDir REAL, rainRate REAL, rain REAL, rpm REAL, radiation REAL)"
# 
# 
Write-Host "Scrittura sul database ogni 5 minuti"

$Database = "C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\apogee.sdb"

#$envlib = $env:Lib
#Remove-Item env:Lib

# Sul database vengono inseriti i valori presi dalla tabella tampone
#
If ($rad -lt 10) {
    # Radiazione solare <10 vuol dire che è notte. Prendo la minima del tampone
    $TempCelsiusSens1=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(outTemp) FROM tampone" -As SingleValue
    $TempCelsiusSens2=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MIN(outTemp2) FROM tampone" -As SingleValue
} else {
    # Radiazione solare >10 vuol dire che è giorno. Prendo la massima del tampone
    $TempCelsiusSens1=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(outTemp) FROM tampone" -As SingleValue
    $TempCelsiusSens2=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(outTemp2) FROM tampone" -As SingleValue
}
$rad=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(radiation) FROM tampone" -As SingleValue
$Vkmh_round=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(windSpeed) FROM tampone" -As SingleValue
$direction=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(windDir) FROM tampone" -As SingleValue
$rpm=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(rpm) FROM tampone" -As SingleValue
$Hum=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(hum) FROM tampone" -As SingleValue
$QNH_round=Invoke-SqliteQuery -DataSource $Database -Query "SELECT AVG(qnh) FROM tampone" -As SingleValue
$intens_pluvio=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(rainRate) FROM tampone" -As SingleValue
$acc_pluvio=Invoke-SqliteQuery -DataSource $Database -Query "SELECT rain FROM tampone ORDER BY dateTime DESC LIMIT 1" -As SingleValue
$tipo_pluvio=Invoke-SqliteQuery -DataSource $Database -Query "SELECT rainType FROM tampone ORDER BY dateTime DESC LIMIT 1" -As SingleValue
$intens_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT MAX(rainRatetb) FROM tampone" -As SingleValue
$acc_pluviotb=Invoke-SqliteQuery -DataSource $Database -Query "SELECT raintb FROM tampone ORDER BY dateTime DESC LIMIT 1" -As SingleValue
#


$query = "INSERT INTO archive (dateTime, outTemp, windSpeed, windDir, rain, rainRate, rainType, radiation, rpm, qnh, raintb, rainRatetb, hum, outTemp2) 
          VALUES ('$unixtime', '$TempCelsiusSens1', '$Vkmh_round', '$direction', '$acc_pluvio', '$intens_pluvio', '$tipo_pluvio', '$rad', '$rpm', '$QNH_round', '$acc_pluviotb', '$intens_pluviotb', '$Hum', '$TempCelsiusSens2')"

Invoke-SqliteQuery -DataSource $Database -Query $query

$query_test = "SELECT * FROM archive ORDER BY dateTime DESC LIMIT 1"
Invoke-SqliteQuery -DataSource $Database -Query $query_test

# Cancella tampone
$query_cancella_tampone = "DELETE FROM tampone"
Invoke-SqliteQuery -DataSource $Database -Query $query_cancella_tampone

$query_vacuum = "VACUUM"
Invoke-SqliteQuery -DataSource $Database -Query $query_vacuum

Write-Host "-------------------------------------"
$env:Lib = $envlib

# Reset cronometro
$cronometroDB.Stop()
$cronometroDB = [diagnostics.stopwatch]::StartNew()

}
# FINE - Evento ogni 5 minuti - SCRITTURA SUL DATABASE
################
#



################
# Evento ogni 5 minuti - INVIO FTP
#

If ($cronometroFTP.elapsed -gt $timeoutFTP){


######################################################
# THREAD PER FTP

$invia_FTP = {

###############
#
# Invio FTP su KWOS
#
###############
# Directory di lavoro Adamcmd
cd C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\

Write-Host "Invio su FTP KWOS ogni 3 minuti"
$server = "ftp.***********.it"
$filelist = "pluviometrotb_single.txt pluviometro_single.txt windspeed_single.txt umidita_single.txt pressione_single.txt radiazione_single.txt winddir_single.txt temperatura_single.txt temperatura2_single.txt apogee_pressione.txt apogee_umidita.txt apogee_temp.txt apogee_rad.txt apogee_rpm.txt apogee_status.txt apogee_windspeed.txt apogee_winddir.txt apogee_pluviometro.txt"   
$user = "***********"
$password = "***********"
$dir = "/www.***********.org/apogee/"

"open $server
user $user $password
binary  
cd $dir     
" +
($filelist.split(' ') | %{ "put ""$_""`n" }) | ftp -i -in

Write-Host "Invio su FTP 192.168.2.205 ogni 3 minuti"
#
# Invio FTP su 192.168.2.205 per confronto temperature
#
$server = "192.168.2.205"
$filelist = "apogee_temp_single.txt apogee_temp2_single.txt apogee_radiation_single.txt"   
$user = "***********"
$password = "***********"
$dir = "/root/"

"open $server
user $user $password
binary  
cd $dir     
" +
($filelist.split(' ') | %{ "put ""$_""`n" }) | ftp -i -in


}
# Fine Thread FTP
#################

################################
# PARTENZA MULTI-THREAD

Start-Job -scriptblock $invia_FTP -Name "invia_FTP"

Get-Job | Wait-Job -Timeout 180

Remove-Job invia_FTP -force


# Reset cronometro
$cronometroFTP.Stop()
$cronometroFTP = [diagnostics.stopwatch]::StartNew()

# FINE MULTI-THREAD
################################

}
# FINE - Evento ogni 3 minuti - INVIO FTP
################
#

###############
#
Write-Host "====================================="
Write-Host "Fine loop"
Write-Host "====================================="
# Fine loop
Start-Sleep -Seconds 60
}

