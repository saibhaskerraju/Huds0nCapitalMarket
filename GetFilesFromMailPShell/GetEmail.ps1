.“$(get-location)\Logging_Functions.ps1” #load the logger library here

try
{

$xdoc = [xml] (get-content “.\GetEmailConfig.xml”) #Load the Config File

$filepath=$xdoc.Data.Destination.value # where to download the file

$subject=$xdoc.Data.subject.value

$searchFilter=$xdoc.Data.FileFilter.value

$logfilePath=("{0}\{1}" -f $(get-location),"LogFile.txt") # Give name of log file here


            $ObjOutlook = New-Object -comobject outlook.application
            $namespace = $ObjOutlook.GetNamespace("MAPI")
            $inBox = $namespace.PickFolder() 
        
            $inBox.Items | where-object {$_.ReceivedTime.ToShortDateString() -eq (get-date).ToShortDateString() -and $_.Subject -eq $subject} | Sort-Object $_.ReceivedTime –Descending | select-object -first 1 | foreach-object {$_.attachments|foreach {
            Write-Host $_.filename
             $mailFileName = $_.filename
            If ($a.Contains($searchFilter)) {
             $_.saveasfile((Join-Path $filepath "$mailFileName"))
             }
            }}
}
catch
{
    $ErrorMessage = $_.Exception.Message
    
     Log-Error -LogPath $logfilePath -ErrorDesc $ErrorMessage -ExitGracefully $True

}



