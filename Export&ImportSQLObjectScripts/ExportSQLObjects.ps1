
."$(get-location)\Logging_Functions.ps1" #load the logger library here
$logfilePath=("{0}\{1}" -f $(get-location),"LogFile_Export.txt") # Give name of log file here
try
{
            $xdoc = [xml] (get-content “.\GetExportConfig.xml”) #Load the Config File
            $serverName=$xdoc.Data.Server
            $dbNames=$xdoc.Data.Database
            $query=$xdoc.Data.Query


            # Object Initialization

                  [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
                  [System.Reflection.Assembly]::LoadWithPartialName("System.Data") | Out-Null
                  $srv = new-object "Microsoft.SqlServer.Management.SMO.Server" $serverName
                  $srv.SetDefaultInitFields([Microsoft.SqlServer.Management.SMO.View], "IsSystemObject")
                  $db = New-Object "Microsoft.SqlServer.Management.SMO.Database"
                 
                  $scr = New-Object "Microsoft.SqlServer.Management.Smo.Scripter"
                  $deptype = New-Object "Microsoft.SqlServer.Management.Smo.DependencyType"
              
                  $options = New-Object "Microsoft.SqlServer.Management.SMO.ScriptingOptions"
                  $options.AllowSystemObjects = $false
                  $options.IncludeDatabaseContext = $true
                  $options.IncludeIfNotExists = $false
                  $options.ClusteredIndexes = $true
                  $options.Default = $true
                  $options.DriAll = $true
                  $options.Indexes = $true
                  $options.NonClusteredIndexes = $true
                  $options.IncludeHeaders = $false
                  $options.ToFileOnly = $true
                  $options.AppendToFile = $true
                  $options.ScriptDrops = $false
                   

        

            # end of obejct initialization


foreach($dbName in $dbNames)
{
                   $db = $srv.Databases[$dbName]
                   $scr.Server = $srv
                   # Set options for SMO.Scripter
                   $scr.Options = $options
                        
                   $scriptpath=$(Get-Location).ToString()+"\"+$xdoc.Data.ScriptLocation+"\"+$dbName+"\"

                   If(!(test-path $scriptpath)) # create storing directory for scripts if doesnt exists
                    {
                    New-Item -ItemType Directory -Force -Path $scriptpath
                    }

                $QueryResult = Invoke-Sqlcmd -ServerInstance $serverName -Database $dbName -Query $query

                $tableNames = New-Object System.Collections.ArrayList
                $procNames = New-Object System.Collections.ArrayList
                $allObjNames = New-Object System.Collections.ArrayList
                
                

                foreach($row in $QueryResult)
                {   
                
                    switch($($row.Type).Trim())
                    {
                        "P"{$procNames.Add($($row.Schema_Name)+"."+$($row.Object_Name));break;}
                        "U"{$tableNames.Add($($row.Schema_Name)+"."+$($row.Object_Name));break;}                    
                    }
                    
                    $allObjNames.Add($dbName.Trim()+"-"+$($row.Schema_Name)+"."+$($row.Object_Name)+"-"+$($row.Type).Trim()+"-"+$serverName);
                }

                $allObjNames | out-file -Append "$(get-location)\AllObjectData.txt" # this file is used for importing and creating objects again
                # We here have all the SQLObjects that were created in a time frame ,  We have to generate scripts of these Objects

                  #=============
                  # Tables
                  #=============
                  
                  Foreach ($tb in $db.Tables)
                  {
                         $tablePath=($scriptpath+"Tables\")
                                  If(!(test-path $tablePath)) # create storing directory for scripts if doesnt exists
                                    {
                                    New-Item -ItemType Directory -Force -Path $tablePath
                                    }
                       
                       if($tableNames.Contains(($tb.Owner+"."+$tb.Name)))# checking for required table names only
                       {
                     
                            $options.FileName = $tablePath + "$($tb.Owner+"."+$tb.Name).sql"  # create new .sql file with table name
                            New-Item $options.FileName -type file -force | Out-Null

                            If ($tb.IsSystemObject -eq $FALSE)
                               {
                                $smoObjects = New-Object Microsoft.SqlServer.Management.Smo.UrnCollection
                                $smoObjects.Add($tb.Urn)
                                $scr.Script($smoObjects)
                               }

                            # table script created
                            #generate its data
                            $queryStr="Select * from "+$($tb.Owner+"."+$tb.Name)
                            $outExcel=$tablePath + "$($tb.Owner+"."+$tb.Name)_data.txt"
                            bcp $queryStr queryout $outExcel -S $serverName /d $dbName /c /t "|" /T
                            #end of generate data
                       }

                  
                  }

                  #=============
                  # End of Tables
                  #=============

                   #=============
                  # Stored Procedures
                  #=============
                  
                  Foreach ($sp in $db.StoredProcedures)
                  {
                      $procPath=($scriptpath+"Procedures\")
                     If(!(test-path $procPath)) # create storing directory for scripts if doesnt exists
                        {
                        New-Item -ItemType Directory -Force -Path $procPath
                        }
                       if($procNames.Contains(($sp.Owner+"."+$sp.Name)))# checking for required table names only
                       {
                     
                            $options.FileName = $procPath + "$($sp.Owner+"."+$sp.Name).sql"  # create new .sql file with table name
                            New-Item $options.FileName -type file -force | Out-Null

                            If ($sp.IsSystemObject -eq $FALSE)
                               {
                                $smoObjects = New-Object Microsoft.SqlServer.Management.Smo.UrnCollection
                                $smoObjects.Add($sp.Urn)
                                $scr.Script($smoObjects)
                               }

                            # SP script created

                       }

                  
                  }

                  #=============
                  # End of Stored Procedures
                  #=============

}# end of foreach of databases



}
catch
{
            $ErrorMessage = $_.Exception.Message
    write-host "In Catch Block"
Log-Error -LogPath $logfilePath -ErrorDesc $ErrorMessage -ExitGracefully $True

}



