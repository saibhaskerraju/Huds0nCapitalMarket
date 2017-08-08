try
{
 # ************************************** Variables Section *****************************************************
 
 
 # NOTE : To Filter for Multiple Objects , always separate them with "," when initializing the variables Ex: $searchFilter="xlsx , pdf"
    
    $downloadFolder="C:\Temp\" # where to download the file
    $subjectList="Lone Star Pipeline - Lone Star"   # To Search for Multiple Mails , separate the Subject Names by a ','
    $searchFilter="xlsx"
    $sender="business.objects@freedommortgage.com"
    $ewsUrlPath = "https://ews.hudson-advisors.com/EWS/Exchange.asmx" 
    $ewsDLLPath = "\Microsoft.Exchange.WebServices.dll" 
    $logfilePath=("{0}\{1}" -f $(get-location),"LogFile.txt") # Give name of log file here
    $userName = "capitalmarketsitinbo"
    $password = "ma9CaKef#Fen"
    $domain = "advisors.hal"
    
        
    
    

    # Email Variables
    $smtp = "dallassmtpt01.advisors.hal"
    <#prepare string array as below if email has to be sent to more than one recipient#>
    #$to = "Ashfaq Syed <asyed@hudson-advisors.com>"
    $to = "hguddanti@hudson-advisors.com"
    <#$to = "PM_IT_Support <PM_IT_Support@Hudson-Advisors.com>"#>
    #$from = "PM-IT Reporting <hadba_noreply@hudson-advisors.com>"
    $from = "hguddanti@hudson-advisors.com"
    $subject = "Freedom Mortage File Downloaded Successfully"
    
    
    
    $hasExcelAttach=$false
    $excelFileName=$null
    $errorMsg=$null
        
        
 
 # ****************************************************************************************************************   
       
    #check for Download Directory available or not
    if(test-path $downloadFolder)
    {
        New-Item -Path $downloadFolder -ItemType directory    
    }
    
    
    
    [Reflection.Assembly]::LoadFile($ewsDLLPath)
    $exchange = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $exchange.Credentials = New-Object Net.NetworkCredential($userName, $password, $domain)
    $exchange.Url = New-Object System.Uri($ewsUrlPath)


    $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchange,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, $sender)
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
    
    $results = $inbox.FindItems($searchFilter, $view)
    
     Foreach ($subject in $subjectList)
    {
             # Filtering the Inbox/Folder to get the Latest Mail for a Particular Day and for Particular Subject
             $filteredResult=$results.Items | where-object {$_.DateTimeReceived.ToShortDateString() -eq (get-date).ToShortDateString() -and $_.Subject -eq $subject} | Sort-Object $_.DateTimeReceived –Descending | select-object -first 1 
             
             if($filteredResult -ne $null)
             {
                     $filteredResult.Load()
                    
                    if($filteredResult.HasAttachments)
                    {
                            foreach($attach in $filteredResult.Attachments)
                             {
                                    
                                    if($attach.Name.ToString().Split('.')[1].ToLower() -eq $searchFilter)
                                    {
                                        
                                        $attach.Load()
                                		$attachFile = new-object System.IO.FileStream(($downloadFolder + “\” + $attach.Name.ToString()), [System.IO.FileMode]::Create)
                                		$attachFile.Write($attach.Content, 0, $attach.Content.Length)
                                		$attachFile.Close()
                                		#write-host "Downloaded Attachment : " + (($downloadFolder + “\” + $attach.Name.ToString()))
                                        
                                        $hasExcelAttach=$true
                                        $excelFileName=$attach.Name.ToString()
                                    }
                            		
                             }
                    
                    }
                    else
                    {
                        # send no attachments mail
                         $subject = "Freedom Mortage File Download Failed"
                         $errorMsg="Freedom Mortage File Download Failed with Error Message: No Attachments were found"
                          send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body ($errorMsg) -BodyAsHtml
                          throw $errorMsg
                    }
                     
             
             }
             else
             {
             
                #Send Mail saying Email Not Arrived
                $subject = "Freedom Mortage File Download Failed"
                $errorMsg="Freedom Mortage File Download Failed with Error Message: Email for "+(get-date).ToShortDateString()+" Not Found"
                send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body ($errorMsg) -BodyAsHtml
                throw $errorMsg
             }
            
            
            #Send mail if we didnt find the searched File
            if(-not $hasExcelAttach)
            {
                  $subject = "Freedom Mortage File Download Failed"
                  $errorMsg="Freedom Mortage File Download Failed with Error Message: Excel File Not Found"
                  send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body ($errorMsg) -BodyAsHtml
                  throw $errorMsg
  
            }

                   
    }# End of Foreach of Subject List
    
    
    # send success mail   
    send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body ("Freedom Mortage File "+$excelFileName+" has been downloaded to "+$downloadFolder+"\"+$excelFileName+" successfully") -BodyAsHtml
    
}
catch
{
    
     <#Sending Email#>
    $subject = "Freedom Mortage File Download Failed"
    $errorMsg="Freedom Mortage File Download Failed with Error Message: " + $_.exception.message
    send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body ($errorMsg) -BodyAsHtml
    throw $errorMsg
}

