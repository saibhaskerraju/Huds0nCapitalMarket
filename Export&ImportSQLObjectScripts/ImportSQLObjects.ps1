

."$(get-location)\Logging_Functions.ps1" #load the logger library here
$logfilePath=("{0}\{1}" -f $(get-location),"LogFile_Import.txt") # Give name of log file here
$rootLoc="$(get-location)\";

$bulkInsertQuery="bulk insert <tablename>
from <path>
with
(
	fieldterminator='|'
)"

try
{

        $allData= Get-Content -Path "$(get-location)\AllObjectData.txt"
        
        if($allData -ne $null)
        {
    
                   foreach($data in $allData)
                    {
        
                        $dbName=$data.split('-')[0]
                        $fullTbName=$data.split('-')[1]
                        $objType=$data.split('-')[2]
                        $serverName=$data.split('-')[3]
                        
                        $fnlScrPath=$rootLoc+"SQLScripts\"+$dbName
               
                        
                        
                        if(test-path $fnlScrPath)
                        {
                            
                           $srchFileLoc= Get-ChildItem -Path $fnlScrPath -Filter "$($fullTbName).sql" -Recurse -ErrorAction SilentlyContinue -Force | select -first 1
                           if($srchFileLoc -ne $null)
                           {
                           
                                switch($objType.Trim())
                                {
                                    "U"
                                    {
                                    
                                              if(Test-Path "$($srchFileLoc.Directory)\$($fullTbName)_data.txt")
                                                {
                                                    #create object script
                                                    invoke-sqlcmd -inputfile $srchFileLoc.fullname -serverinstance $serverName -database $dbName
                                                    
                                                    #load data into table
                                                    $bulkInsertQuery=$bulkInsertQuery.replace("<path>","'$($srchFileLoc.Directory)\$($fullTbName)_data.txt'").replace("<tablename>","$($fullTbName)")
                                                    Invoke-Sqlcmd -Query $bulkInsertQuery -ServerInstance $serverName -database $dbName
                                                }
                                                else
                                                {
                                                    #log
                                                    write-host "in fourth if cond"
                                                }
                                    }
                                    default
                                    {
                                    
                                             #create object script
                                                invoke-sqlcmd -inputfile $srchFileLoc.fullname -serverinstance $serverName -database $dbName
                                    }
                                
                                }
                           
          
                              
                           }
                           else
                           {
                            #log mail
                            write-host "in third if cond"
                           }
                        }
                        else
                        {
                            #write into log file  or mail 
                            write-host "in 2nd if cond"
                        }

                        
                    }
        }
        else
        {

            write-host "in first if cond"
        }
     

}
catch
{
            $ErrorMessage = $_.Exception.Message
    write-host "In Catch Block"
Log-Error -LogPath $logfilePath -ErrorDesc $ErrorMessage -ExitGracefully $True

}