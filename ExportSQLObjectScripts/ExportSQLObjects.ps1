
."$(get-location)\Logging_Functions.ps1" #load the logger library here
$logfilePath=("{0}\{1}" -f $(get-location),"LogFile.txt") # Give name of log file here
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
                  $db = $srv.Databases[$dbName]
                  $scr = New-Object "Microsoft.SqlServer.Management.Smo.Scripter"
                  $deptype = New-Object "Microsoft.SqlServer.Management.Smo.DependencyType"
                  $scr.Server = $srv
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
                   

                  # Set options for SMO.Scripter
                  $scr.Options = $options

            # end of obejct initialization


foreach($dbName in $dbNames)
{

                        
                   $scriptpath=$xdoc.Data.ScriptLocation+"\"+$dbName+"\"

                   If(!(test-path $scriptpath)) # create storing directory for scripts if doesnt exists
                    {
                    New-Item -ItemType Directory -Force -Path $scriptpath
                    }

                $QueryResult = Invoke-Sqlcmd -ServerInstance $serverName -Database $dbName -Query $query

                $tableNames = New-Object System.Collections.ArrayList
                $procNames = New-Object System.Collections.ArrayList

                foreach($row in $QueryResult)
                {   
                
                write-host $($row.Type)

                if($($row.Type).Trim() -eq "P"){
                
                    $procNames.Add($($row.Schema_Name)+"."+$($row.Object_Name))
                }

                if($row.Type.Trim() -eq "U"){
                
                   $tableNames.Add($($row.Schema_Name)+"."+$($row.Object_Name))
                }
                    
                }

                # We here have all the tables that were createded in a time frame 

                # We have to generate scripts of these tables

                 

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
                            bcp $queryStr queryout $outExcel -S $serverName /d $dbName /c /T
                            #end of generate data
                       }

                  
                  }

                  #=============
                  # End of Tables
                  #=============

                   #=============
                  # Stored Procedures
                  #=============
                  
                  Foreach ($tb in $db.StoredProcedures)
                  {
                  $procPath=($scriptpath+"Procedures\")
                     If(!(test-path $procPath)) # create storing directory for scripts if doesnt exists
                    {
                    New-Item -ItemType Directory -Force -Path $procPath
                    }
                       if($procNames.Contains(($tb.Owner+"."+$tb.Name)))# checking for required table names only
                       {
                     
                            $options.FileName = $procPath + "$($tb.Owner+"."+$tb.Name).sql"  # create new .sql file with table name
                            New-Item $options.FileName -type file -force | Out-Null

                            If ($tb.IsSystemObject -eq $FALSE)
                               {
                                $smoObjects = New-Object Microsoft.SqlServer.Management.Smo.UrnCollection
                                $smoObjects.Add($tb.Urn)
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



