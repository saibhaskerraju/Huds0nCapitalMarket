.“$(get-location)\Logging_Functions.ps1” #load the logger library here

try
{

$xdoc = [xml] (get-content “.\GetEmailConfig.xml”) #Load the Config File

$filepath=$xdoc.Data.Destination # where to download the file

$subjectList=$xdoc.Data.subject   #separate mails to search by comma

$searchFolder=$xdoc.Data.searchfolder

$searchFilter=$xdoc.Data.FileFilter

$emailID=$xdoc.Data.emailID

$logfilePath=("{0}\{1}" -f $(get-location),"LogFile.txt") # Give name of log file here


            $ObjOutlook = New-Object -comobject outlook.application
            $namespace = $ObjOutlook.GetNamespace("MAPI")
            
            $Account = $namespace.Folders | ? { $_.Name -eq $emailID };
            $inBox = $Account.Folders | ? { $_.Name -match $searchFolder };
            
            #$inBox = $namespace.PickFolder() 
            
            Foreach ($subject in $subjectList)
            {
                            $inBox.Items | where-object {$_.ReceivedTime.ToShortDateString() -eq (get-date).ToShortDateString() -and $_.Subject -eq $subject} | Sort-Object $_.ReceivedTime –Descending | select-object -first 1 | foreach-object {$_.attachments|foreach {
                            Write-Host $_.filename
                            $mailFileName = $_.filename
                            
                            If ($mailFileName.Contains($searchFilter)) {
                             $_.saveasfile((Join-Path $filepath "$mailFileName"))
                             $canDelete=$true
                             }
                             
                            }#end of foreach attachment iterator
                            
                            if($canDelete){
                                $_.Delete()    #delete mail after downloading file
                            }
                            else
                            {
                                $errStr="The Mail is either not deleted or Attachment not downloaded"
                                write-host $errStr
                                Log-Error -LogPath $logfilePath -ErrorDesc $errStr -ExitGracefully $True
                            }
                            
                            }#end of foreach object iterator
            
            }#end of subject foreach iterator
        
           
}
catch
{
    $ErrorMessage = $_.Exception.Message
    write-host "In Catch Block"
     Log-Error -LogPath $logfilePath -ErrorDesc $ErrorMessage -ExitGracefully $True

}



