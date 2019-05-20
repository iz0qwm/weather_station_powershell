$endpoint = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 8987)
$listener = New-Object System.Net.Sockets.TcpListener $endpoint 
if($listener -eq $null) { return; } 


$listener.Start()
[console]::WriteLine("Listening on :8987")
while ($true) {
    $client = $listener.AcceptTcpClient() # will block here until connection
    [console]::WriteLine("{0} >> Accepted Client " -f (Get-Date).ToString())
    $stream = $client.GetStream();
    $reader = New-Object System.IO.StreamReader $stream
    [console]::WriteLine("Inside Processing")
    #try
    #{
        while ($true) { 
               try{
               
                    $stringa = Get-Content -Path C:\Users\Ruggero\AdamLib-master\AdamCmd\bin\Debug\netcoreapp2.1\win7-x64\publish\misure_server.txt
                    #$stringa = $stringa + "`r`n"
                    $bytes = [System.Text.Encoding]::ASCII.GetByteCount($stringa)
                    $stream.Write([text.Encoding]::Ascii.GetBytes($stringa), 0, $bytes)
                    [console]::WriteLine($stringa)
                    Start-Sleep -Seconds 30
                } catch {
                    $reader.Dispose()
                    $stream.Dispose()
                    $client.Dispose()
                    $reader.Close() 
                    $stream.Close()
                    $client.Close()
                    #$listener.Stop()
                    [console]::WriteLine("Client disconnected") 
                    break
                } 
        } 
    #}
    #finally
    ##{
        #[console]::WriteLine("Connection exit")

    #}
} 
