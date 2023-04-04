##Clean Teams Cache V2 - this one may be outdated - need to review.
##This version deletes all of the folders under Teams except the Backgrounds and meeting-addin. 
##It also deletes preference and settings files, so please inform the user preferred settings will need to be set again

$challenge = Read-Host "Are you sure you want to delete Teams Cache (Y/N)?"
$challenge = $challenge.ToUpper()

if ($challenge -eq "N"){

    Stop-Process -Id $PID

    }
    elseif ($challenge -eq "Y"){
        Write-Host "Stopping Teams Process" -ForegroundColor Yellow
            try{
                 
                    Get-Process -ProcessName Teams | Stop-Process -Force
                    Start-Sleep -Seconds 3
                    Write-Host "Teams Process Sucessfully Stopped" -ForegroundColor Green

                }
                catch{

                    echo $_

                     }
                
                    Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
            try{
                   
                    $TeamsPath = $env:APPDATA+"\Microsoft\Teams\"
                        Get-ChildItem -Path $TeamsPath -Exclude "Backgrounds", "meeting-addin" | foreach ($_) {
                        Write-Host "CLEANING :" + $_.fullname
                        Remove-Item $_.fullname -Force -Recurse
                        Write-Host "CLEANED... :" + $_.fullname
                    }   

                    #Get-ChildItem -Path $env:APPDATA\"Microsoft\teams" | Remove-Item -Confirm:$true
                    #Write-Host "Teams Disk Cache Cleaned" -ForegroundColor Green

                }

                catch{
                        echo $_
                      }

            try{

                    Write-Host "Cleanup Complete... Launching Teams" -ForegroundColor Green
                    Start-Process -File $env:LOCALAPPDATA\Microsoft\Teams\Update.exe -ArgumentList '--processStart "Teams.exe"'
                    Stop-Process -Id $PID

                }
                catch{
                        echo $_

                      }
    }

<#Write-Host "Stopping Chrome Process" -ForegroundColor Yellow
try{
    
    Get-Process -ProcessName Chrome| Stop-Process -Force
    Start-Sleep -Seconds 3
    Write-Host "Chrome Process Sucessfully Stopped" -ForegroundColor Green

    }

    catch{
    echo $_

    }

    Write-Host "Clearing Chrome Cache" -ForegroundColor Yellow

    try{

        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cache" | Remove-Item -Confirm:$false
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cookies" -File | Remove-Item -Confirm:$false
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Web Data" -File | Remove-Item -Confirm:$false
        Write-Host "Chrome Cleaned" -ForegroundColor Green

        }
        
        catch{

        echo $_

        }

Write-Host "Stopping IE Process" -ForegroundColor Yellow

try{

    Get-Process -ProcessName MicrosoftEdge | Stop-Process -Force
    Get-Process -ProcessName IExplore | Stop-Process -Force
    Write-Host "Internet Explorer and Edge Processes Sucessfully Stopped" -ForegroundColor Green

    }
    
    catch{

        echo $_

    }

Write-Host "Clearing IE Cache" -ForegroundColor Yellow

try{

    RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
    RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
    Write-Host "IE and Edge Cleaned" -ForegroundColor Green

    }
    
    catch{

        echo $_
    }

Write-Host "Cleanup Complete... Launching Teams" -ForegroundColor Green
Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe
Stop-Process -Id $PID

}

else{

    Stop-Process -Id $PID
}
#>