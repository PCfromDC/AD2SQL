# Array of Domain Controller Server Names
$DCs = @("DC01","DC02","DC03")

# Database Server
$dbServer = "sql2012-03"

# Database Name
$databaseName = "pcDemo_Personnel"

# Production System User Table Name
$activeTableName = "pcDemo_SystemUsers"

# create out file location
$saveLocation = "c:\psOutputs\"

# Days to Keep Synchronization File History
$daysSaved = 6

# Saved File Name
$date = Get-Date -Format s
$fileName = "synchedUsers_" + $date + ".txt"
$fileName = $fileName.Replace(":","_")

# Saved File
$file = $saveLocation + $fileName

# Verify Folder Exists else Create It
if ((Test-Path $saveLocation) -eq $false)
{
    [IO.Directory]::CreateDirectory($saveLocation)
}

<# 
#  **************************
#  * Create Functions Below *
#  ************************** 
#>
#region *** Function Definitions ***

####################### 
function Get-Type 
{ 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
} #Get-Type 
 
####################### 
<# 
.SYNOPSIS 
Creates a DataTable for an object 
.DESCRIPTION 
Creates a DataTable based on an objects properties. 
.INPUTS 
Object 
    Any object can be piped to Out-DataTable 
.OUTPUTS 
   System.Data.DataTable 
.EXAMPLE 
$dt = Get-psdrive| Out-DataTable 
This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable 
.NOTES 
Adapted from script by Marc van Orsouw see link 
Version History 
v1.0  - Chad Miller - Initial Release 
v1.1  - Chad Miller - Fixed Issue with Properties 
v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
v1.4  - Chad Miller - Corrected issue with DBNull 
v1.5  - Chad Miller - Updated example 
v1.6  - Chad Miller - Added column datatype logic with default to string 
v1.7 - Chad Miller - Fixed issue with IsArray 
.LINK 
http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx 
#> 
function Out-DataTable 
{ 
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = new-object Data.datatable   
        $First = $true  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
} #Out-DataTable

try {add-type -AssemblyName "Microsoft.SqlServer.ConnectionInfo, Version=10.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" -EA Stop} 
catch {add-type -AssemblyName "Microsoft.SqlServer.ConnectionInfo"} 
 
try {add-type -AssemblyName "Microsoft.SqlServer.Smo, Version=10.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" -EA Stop}  
catch {add-type -AssemblyName "Microsoft.SqlServer.Smo"}  
 
#######################  
function Get-SqlType  
{  
    param([string]$TypeName)  
  
    switch ($TypeName)   
    {  
        'Boolean' {[Data.SqlDbType]::Bit}  
        'Byte[]' {[Data.SqlDbType]::VarBinary}  
        'Byte'  {[Data.SQLDbType]::VarBinary}  
        'Datetime'  {[Data.SQLDbType]::DateTime}  
        'Decimal' {[Data.SqlDbType]::Decimal}  
        'Double' {[Data.SqlDbType]::Float}  
        'Guid' {[Data.SqlDbType]::UniqueIdentifier}  
        'Int16'  {[Data.SQLDbType]::SmallInt}  
        'Int32'  {[Data.SQLDbType]::Int}  
        'Int64' {[Data.SqlDbType]::BigInt}  
        'UInt16'  {[Data.SQLDbType]::SmallInt}  
        'UInt32'  {[Data.SQLDbType]::Int}  
        'UInt64' {[Data.SqlDbType]::BigInt}  
        'Single' {[Data.SqlDbType]::Decimal} 
        default {[Data.SqlDbType]::VarChar}  
    }  
      
} #Get-SqlType 
 
#######################  
<#  
.SYNOPSIS  
Creates a SQL Server table from a DataTable  
.DESCRIPTION  
Creates a SQL Server table from a DataTable using SMO.  
.EXAMPLE  
$dt = Invoke-Sqlcmd2 -ServerInstance "Z003\R2" -Database pubs "select *  from authors"; Add-SqlTable -ServerInstance "Z003\R2" -Database pubscopy -TableName authors -DataTable $dt  
This example loads a variable dt of type DataTable from a query and creates an empty SQL Server table  
.EXAMPLE  
$dt = Get-Alias | Out-DataTable; Add-SqlTable -ServerInstance "Z003\R2" -Database pubscopy -TableName alias -DataTable $dt  
This example creates a DataTable from the properties of Get-Alias and creates an empty SQL Server table.  
.NOTES  
Add-SqlTable uses SQL Server Management Objects (SMO). SMO is installed with SQL Server Management Studio and is available  
as a separate download: http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=ceb4346f-657f-4d28-83f5-aae0c5c83d52  
Version History  
v1.0   - Chad Miller - Initial Release  
v1.1   - Chad Miller - Updated documentation 
v1.2   - Chad Miller - Add loading Microsoft.SqlServer.ConnectionInfo 
v1.3   - Chad Miller - Added error handling 
v1.4   - Chad Miller - Add VarCharMax and VarBinaryMax handling 
v1.5   - Chad Miller - Added AsScript switch to output script instead of creating table 
v1.6   - Chad Miller - Updated Get-SqlType types 
#>  
function Add-SqlTable  
{  
  
    [CmdletBinding()]  
    param(  
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance,  
    [Parameter(Position=1, Mandatory=$true)] [string]$Database,  
    [Parameter(Position=2, Mandatory=$true)] [String]$TableName,  
    [Parameter(Position=3, Mandatory=$true)] [System.Data.DataTable]$DataTable,  
    [Parameter(Position=4, Mandatory=$false)] [string]$Username,  
    [Parameter(Position=5, Mandatory=$false)] [string]$Password,  
    [ValidateRange(0,8000)]  
    [Parameter(Position=6, Mandatory=$false)] [Int32]$MaxLength=1000, 
    [Parameter(Position=7, Mandatory=$false)] [switch]$AsScript 
    )  
  
 try { 
    if($Username)  
    { $con = new-object ("Microsoft.SqlServer.Management.Common.ServerConnection") $ServerInstance,$Username,$Password }  
    else  
    { $con = new-object ("Microsoft.SqlServer.Management.Common.ServerConnection") $ServerInstance }  
      
    $con.Connect()  
  
    $server = new-object ("Microsoft.SqlServer.Management.Smo.Server") $con  
    $db = $server.Databases[$Database]  
    $table = new-object ("Microsoft.SqlServer.Management.Smo.Table") $db, $TableName  
  
    foreach ($column in $DataTable.Columns)  
    {  
        $sqlDbType = [Microsoft.SqlServer.Management.Smo.SqlDataType]"$(Get-SqlType $column.DataType.Name)"  
        if ($sqlDbType -eq 'VarBinary' -or $sqlDbType -eq 'VarChar')  
        {  
            if ($MaxLength -gt 0)  
            {$dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType, $MaxLength} 
            else 
            { $sqlDbType  = [Microsoft.SqlServer.Management.Smo.SqlDataType]"$(Get-SqlType $column.DataType.Name)Max" 
              $dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType 
            } 
        }  
        else  
        { $dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType }  
        $col = new-object ("Microsoft.SqlServer.Management.Smo.Column") $table, $column.ColumnName, $dataType  
        $col.Nullable = $column.AllowDBNull  
        $table.Columns.Add($col)  
    }  
  
    if ($AsScript) { 
        $table.Script() 
    } 
    else { 
        $table.Create() 
    } 
} 
catch { 
    $message = $_.Exception.GetBaseException().Message 
    Write-Error $message 
} 
   
} #Add-SqlTable

####################### 
<# 
.SYNOPSIS 
Writes data only to SQL Server tables. 
.DESCRIPTION 
Writes data only to SQL Server tables. However, the data source is not limited to SQL Server; any data source can be used, as long as the data can be loaded to a DataTable instance or read with a IDataReader instance. 
.INPUTS 
None 
    You cannot pipe objects to Write-DataTable 
.OUTPUTS 
None 
    Produces no output 
.EXAMPLE 
$dt = Invoke-Sqlcmd2 -ServerInstance "Z003\R2" -Database pubs "select *  from authors" 
Write-DataTable -ServerInstance "Z003\R2" -Database pubscopy -TableName authors -Data $dt 
This example loads a variable dt of type DataTable from query and write the datatable to another database 
.NOTES 
Write-DataTable uses the SqlBulkCopy class see links for additional information on this class. 
Version History 
v1.0   - Chad Miller - Initial release 
v1.1   - Chad Miller - Fixed error message 
.LINK 
http://msdn.microsoft.com/en-us/library/30c3y597%28v=VS.90%29.aspx 
#> 
function Write-DataTable 
{ 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$true)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$true)] [string]$TableName, 
    [Parameter(Position=3, Mandatory=$true)] $Data, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=5, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$BatchSize=50000, 
    [Parameter(Position=7, Mandatory=$false)] [Int32]$QueryTimeout=0, 
    [Parameter(Position=8, Mandatory=$false)] [Int32]$ConnectionTimeout=15 
    ) 
     
    $conn=new-object System.Data.SqlClient.SQLConnection 
 
    if ($Username) 
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout } 
    else 
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout } 
 
    $conn.ConnectionString=$ConnectionString 
 
    try 
    { 
        $conn.Open() 
        $bulkCopy = new-object ("Data.SqlClient.SqlBulkCopy") $connectionString 
        $bulkCopy.DestinationTableName = $tableName 
        $bulkCopy.BatchSize = $BatchSize 
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut 
        $bulkCopy.WriteToServer($Data) 
        $conn.Close() 
    } 
    catch 
    { 
        $ex = $_.Exception 
        Write-Error "$ex.Message" 
        continue 
    } 
 
} #Write-DataTable

#######################
function writeStartTime($string)
{
    # add start time/date to outfile
    $startTime = Get-Date
    $string + " started: " + $startTime | Out-File $file -Append
} #writeQueryStartTime

#######################
function writeFinishTime($string)
{
    # add finish time/date to outfile
    $endTime = Get-Date
    $string + " finished :" + $endTime | Out-File $file -Append

    # add duration time to outfile
    $queryDuration = ($endTime - $startTime).duration()
    $string + " duration: " + $queryDuration | Out-File $file -Append
} #writeQueryFinishTime


####################
# Validate Servers #
####################

Function validateServer ($s)
{
    $alive = $true

    if(!(Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet))

    {    

    "Problem connecting to $s" | Out-File $file -Append
    ipconfig /flushdns | out-null
    ipconfig /registerdns | out-null
    nslookup $s | out-null
    if(!(Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet))
        {
            $alive = $false
        }

    ELSE 
        {
            "Resolved problem connecting to $s" | Out-File $file -Append
            $alive = $true
        } #end if

   } # end if

   return $alive # always a good sign!

} # Validate Server Alive

#endregion
<# 
#  **************************
#  * Create Functions Above *
#  ************************** 
#> 

<# 
#  *************************
#  * Synchronize AD to SQL *
#  *************************
#> 

# Create Out-File and add start time/date
$PoSH_startTime = Get-Date
"Synchronize AD to SQL PowerShell started: " + $PoSH_startTime | Out-File $file

# Validate Domain Controllers
$OUs = @() 
foreach ($DC in $DCs)
{
    $a = validateServer($DC)
    
    if ($a)
    {
        "$DC is alive: " + $a | Out-File $file -Append
        $OUs += $DC
    }
}

$counter = 0
foreach ($OU in $OUs)
{
    # Get current OU Server Name
    $ouServer = $OUs[$counter]

    # Create Table Name
    $tableName = "temp_" + $ouServer + "_Table"

    # Drop table if it exists
    $query1 = "IF OBJECT_ID('dbo.$tableName', 'U') IS NOT NULL DROP TABLE dbo.$tableName"
    Invoke-Sqlcmd -Query $query1 -Database $databaseName -ServerInstance $dbServer

    # add AD query start time/date to outfile
    $startTime = Get-Date
    "Query AD " + $ouServer + " started: " + $startTime | Out-File $file -Append

    # Set AD Properties to return 
    if ($counter -lt 1)
    {
        $properties = ("sAMAccountName","displayName","mail","telephoneNumber","physicalDeliveryOfficeName","department","userAccountControl","company","title","lastLogon","manager","givenName","Surname")
    }
    else
    {
        $properties = ("sAMAccountName","lastLogon")
    }

    # Get Users and their properties out of AD where the displayName is not blank
    $users = Get-ADUser -Filter * -Server $ouServer -Properties (foreach{$properties}) | Select (foreach{$properties})

  # $users = Get-ADUser -Filter {displayName -like "*"}  -Server $ouServer -Properties (foreach{$properties}) | Select (foreach{$properties})

    # add AD query finish time/date to outfile
    $endTime = Get-Date
    "Query AD " + $ouServer + " finished :" + $endTime | Out-File $file -Append

    # add duration time to outfile
    $queryDuration = ($endTime - $startTime).duration()
    "Query AD " + $ouServer + " duration: " + $queryDuration | Out-File $file -Append
    
    
    # Clean up lastLogon values
    foreach ($user in $users)
    {
        if (!$user.lastLogon)
            {
                $user.lastLogon = 0
            }
        else
            {
                $user.lastLogon = [datetime]::FromFileTime($user.lastLogon).ToString('yyyy-MM-dd HH:mm:ss.fff') 
            }
    }
    
    # SQL Write start time/date to outfile
    $sqlStartTime = Get-Date
    "SQL Creation started: " + $sqlStartTime | Out-File $file -Append

    # Turn $users into DataTable
    $dt1 = $users | Out-DataTable 
  
    # Create SQL Table
    Add-SqlTable -ServerInstance $dbServer -Database $databaseName -TableName $tableName -DataTable $dt1

    # Write DataTable into SQL
    Write-DataTable -ServerInstance $dbServer -Database $databaseName -TableName $tableName -Data $dt1

    # Clean up new table from NULL error work around
    $query2 = "UPDATE [dbo].$tableName SET lastLogon = NULL WHERE lastLogon = '0'"
    Invoke-Sqlcmd -Query $query2 -Database $databaseName -ServerInstance $dbServer

    # SQL Write finish time/date to outfile
    $sqlEndTime = Get-Date
    "SQL Creation finished :" + $sqlEndTime | Out-File $file -Append

    # add duration time to outfile
    $sqlQueryDuration = ($sqlEndTime - $sqlStartTime).duration()
    "SQL Creation duration: " + $sqlQueryDuration | Out-File $file -Append

    $counter ++   
} #Synchronize AD to SQL

<#
# **********************************************************************
# * Move Last Logon Times to Temp Table If Multiple Domain Controllers *
# **********************************************************************
#>

if ($OUs.Count -gt 1)
{
    
    # Drop table if it exists
    $query3 = "IF OBJECT_ID('dbo.temp_lastLogonTimes', 'U') IS NOT NULL DROP TABLE dbo.temp_lastLogonTimes"
    Invoke-Sqlcmd -Query $query3 -Database $databaseName -ServerInstance $dbServer

    # Create temp_lastLogonTimes Table
    $query4 = "CREATE TABLE temp_lastLogonTimes (sAMAccountName varchar(1000))"
    Invoke-Sqlcmd -Query $query4 -Database $databaseName -ServerInstance $dbServer

    # Add a column for each OU

    foreach ($OU in $OUs)
    {
        # Create OU Columns
        $columnName = $OU + "_lastLogon"
        $query5 = "ALTER TABLE temp_lastLogonTimes ADD " + $columnName + " varchar(1000)"
        Invoke-Sqlcmd -Query $query5 -Database $databaseName -ServerInstance $dbServer
    }

    # Insert and Update Times Into Temp Table
    $counter = 0
    foreach ($OU in $OUs)
    {
        if ($counter -lt 1)
        {
            # Insert Names and Times
            $query6 = "INSERT INTO [dbo].[temp_lastLogonTimes] 
                            ([sAMAccountName]
                            ,[" + $OU + "_lastLogon])
                       Select
                            sAMAccountName 
                           ,lastLogon
                       FROM
                           temp_" + $OU + "_Table"
            Invoke-Sqlcmd -Query $query6 -Database $databaseName -ServerInstance $dbServer
        }

        # Update OU lastLogon Times *** Adjust Query Timeout Accordingly ***
        $query7 = "UPDATE [dbo].[temp_lastLogonTimes] 
                   SET " + $OU + "_lastLogon = lastLogon
                   FROM temp_" + $OU + "_Table
                   WHERE temp_lastLogonTimes.sAMAccountName = temp_" + $OU + "_Table.sAMAccountName"
        Invoke-Sqlcmd -Query $query7 -Database $databaseName -ServerInstance $dbServer # -QueryTimeout 600
        $counter ++
    }

    <#
     # ***************************
     # * Get Max lastLogon Times *
     # ***************************
     #>

    # Get Table and Update Last Logon Value
    $str_OUs = @()
    foreach ($OU in $OUs)
    {
        $str_OUs += "ISNULL(" + $OU + "_lastLogon, 0) as " + $OU + "_lastLogon"
    }
    $str_OUs = $str_OUs -join ", "
    
    $query8 = "SELECT sAMAccountName, " + $str_OUs + " from temp_lastLogonTimes"
    $arrayLLT = @()
    $arrayLLT = Invoke-Sqlcmd -Query $query8 -Database $databaseName -ServerInstance $dbServer
    $arrayLLT | Add-Member -MemberType NoteProperty -Name "lastLogon" -Value ""
    $arrayLength = $arrayLLT[0].Table.Columns.Count - 1

    $counter = 0
    foreach ($sAM in $arrayLLT.sAMAccountName)
    {
        $max = $arrayLLT[$counter][1..$arrayLength] | Measure -Maximum
        $arrayLLT[$counter].lastLogon = $max.Maximum
        # $arrayLLT[$counter].lastLogon = [datetime]::FromFileTime($max.Maximum).ToString('yyyy-MM-dd HH:mm:ss.fff') 
        $counter ++
    }

    # Drop table if it exists
    $tableNameLLT = "temp_lastLogons"
    $query9 = "IF OBJECT_ID('dbo.$tableNameLLT', 'U') IS NOT NULL DROP TABLE dbo.$tableNameLLT"
    Invoke-Sqlcmd -Query $query9 -Database $databaseName -ServerInstance $dbServer

    # Turn $users into DataTable
    $arrayLLT = $arrayLLT | Select sAMAccountName, lastLogon 
    $dt2 = $arrayLLT | Out-DataTable 
  
    # Create SQL Table
    Add-SqlTable -ServerInstance $dbServer -Database $databaseName -TableName $tableNameLLT -DataTable $dt2

    # Write DataTable into SQL
    Write-DataTable -ServerInstance $dbServer -Database $databaseName -TableName $tableNameLLT -Data $dt2

    # Clean up new table from NULL error work around
    $query10 = "UPDATE [dbo].$tableNameLLT SET lastLogon = NULL WHERE lastLogon = '0'"
    Invoke-Sqlcmd -Query $query10 -Database $databaseName -ServerInstance $dbServer
}

<#
#     ********************************************
#     * Update Current Users In $activeTableName *
#     ********************************************
#>

$tempTableName = "temp_" + $OUs[0] + "_Table"
$query11 = "UPDATE active
		    SET
	            active.UserLogin = LOWER(temp.sAMAccountName),
                active.UserFullName = temp.displayName,
                active.UserLastName = temp.Surname,
                active.UserFirstName = temp.givenName,
                active.UserCompany = temp.company,
                active.UserOfficeLocation = temp.physicalDeliveryOfficeName,
                active.UserTitle = temp.title,
                active.Manager = temp.manager,
                active.UserPhone = temp.telephoneNumber,
                active.UserEmail = temp.mail,
                active.lastLogon = CONVERT(DATETIME, temp.lastLogon),
                active.userAccountControl = temp.userAccountControl,
	            active.Department = temp.department	  		 
           FROM " + $activeTableName + " active
		            inner join " + $tempTableName + " temp
		                on active.UserLogin = temp.sAMAccountName
    	   WHERE LOWER(active.UserLogin) = LOWER(temp.sAMAccountName)"
Invoke-Sqlcmd -Query $query11 -Database $databaseName -ServerInstance $dbServer

<#
#     *********************************************
#     * Insert New Accounts Into $activeTableName *
#     *********************************************
#>

$query12 = "INSERT INTO [" + $databaseName + "].[dbo].[" + $activeTableName + "]
(
      [UserLogin],
      [UserFullName],
      [UserLastName],
      [UserFirstName],
      [UserCompany],
      [UserOfficeLocation],
      [Department],
      [UserTitle],
      [Manager],
      [UserPhone],
      [UserEmail],
      [System_Role],
      [ReadOnly],
      [lastLogon],
      [userAccountControl]
)
SELECT 
	LOWER(sAMAccountName),
	[displayName],
	[givenName],
	[Surname],
	[company],
	[physicalDeliveryOfficeName],
	[department],
	[title],
	[manager],
	[telephoneNumber],
	[mail],
    [System_Role] = 'User',
	[ReadOnly] = 'Y',
	CONVERT(DATETIME, [lastLogon]),
	[userAccountControl]
FROM " + $tempTableName + " AS temp
WHERE sAMAccountName <> '' and not exists
(
	SELECT LOWER(UserLogin)
	FROM " + $activeTableName + " AS active
	WHERE LOWER(active.UserLogin) = LOWER(temp.sAMAccountName)
)"
Invoke-Sqlcmd -Query $query12 -Database $databaseName -ServerInstance $dbServer

<#
#     ***************************************************************
#     * Update lastLogon Time In $activeTableName IF more than 1 DC *
#     ***************************************************************
#>

if ($OUs.Count -gt 1)
{
        $query13 = "UPDATE [dbo].[" + $activeTableName + "] 
                   SET " + $activeTableName + ".lastLogon = temp_lastLogons.lastLogon
                   FROM temp_lastLogons
                   WHERE LOWER(temp_lastLogons.sAMAccountName) = LOWER(" + $activeTableName + ".UserLogin)"
        Invoke-Sqlcmd -Query $query13 -Database $databaseName -ServerInstance $dbServer 
}

# Write Number of People Found in AD
"Total number of users imported from AD: " + $users.count | Out-File $file -Append

# Clean Up Old Files
Get-ChildItem $saveLocation -Recurse | Where {$_.LastWriteTime -lt (Get-Date).AddDays(-$daysSaved)} | Remove-Item -Force

# add PoSH finish time/date to outfile
$PoSH_endTime = Get-Date
"Synchronize AD to SQL PowerShell finished: " + $PoSH_endTime | Out-File $file -Append

# add PoSH duration time to outfile
$queryDuration = ($PoSH_endTime - $PoSH_startTime).duration()
"Synchronize AD to SQL PowerShell duration: " + $queryDuration | Out-File $file -Append
