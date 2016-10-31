# ************************************************************************************
#  Author:   Brad Hurst
#  Date:     10/31/2016
#  Name:     SSIS_Project_Deploy.ps1
#  Purpose:  Show simple example of deploying ssis project model to SSIS Catalog (SSISDB)
#
#  Description: Deploys a simple ssis project (ispac) that contains a single empty package
#               along with 2 project-level connection managers, SourceDB and TargetDB
#               The Script will deploy the ispac to the SSISDB to a folder called ETL,
#               Create an environment called "DEV" and add 2 variables to it that 
#               Represent connection strings to Source and Target.
#  
#  Required Setup: The machine on which this script runs must have SSIS installed.
# ************************************************************************************
   
$environment = "DEV"
$BaseDir = ".\Project1\bin\Development" #replace with your base path if different.
$ProjectName = "Project1"
$SourceConnectionString = "Data Source=(local);Initial Catalog=Source;Provider=SQLNCLI11.1;Integrated Security=SSPI;"
$TargetConnectionString = "Data Source=(local);Initial Catalog=Target;Provider=SQLNCLI11.1;Integrated Security=SSPI;"

$IsPacFilePath = "${BaseDir}\${ProjectName}.ispac"
$IntegrationServicesCatalog = "SSISDB"
$FolderName = "ETL"

if (!(Test-Path $IsPacFilePath)) {	
	Write-Host "Error: $IsPacFilePath does not exist."
	throw
}

# Load the IntegrationServices Assembly
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.IntegrationServices") | Out-Null;
 
# Store the IntegrationServices Assembly namespace to avoid typing it every time
$ISNamespace = "Microsoft.SqlServer.Management.IntegrationServices"
 
# Create a connection to the server
$sqlConnectionString = "Data Source=(local);Initial Catalog=master;Integrated Security=SSPI;"

$sqlConnection = New-Object System.Data.SqlClient.SqlConnection $sqlConnectionString
try {		
	Write-Host "Attempting to connect to SSIS catalog on localhost..."
	$integrationServices = New-Object $ISNamespace".IntegrationServices" $sqlConnection
	Write-Host "Connection to SSIS Catalog successful."
}
catch {
	Write-Host "Error: failed to establish connection to the ssis catalog server, localhost, failed."
	throw
}

try {
    Write-Host "getting catalog object..."
	$catalog = $integrationServices.Catalogs["SSISDB"]
}
catch {
	Write-Host "Error: Integration Services Catalog does not exist on localhost."
	throw
}

# $existingFolders = $catalog.Folders | Foreach {"$($_.Name)" }
Write-Host "Getting $FolderName folder..."

$folder = $null
if ($catalog.Folders -ne $null)  {
    $folder = $catalog.Folders["$FolderName"]
}
Write-Host "Catalog: $catalog"
Write-Host "FolderName: $FolderName"
if ($folder -eq $null) {
	try {
		Write-Host "Creating Folder " $FolderName " ..."	 
		# Create a new folder
		$folder = New-Object $ISNamespace".CatalogFolder" ($catalog, $FolderName, "Folder description")
		$folder.Create()
	}
	catch {
		Write-Host "Error: failed creating folder $FolderName."
		throw
	}
}
else {
	$folder = $catalog.Folders["$FolderName"]
}

Write-Host "Deploying " $ProjectName " project to the SSIS Catalog on localhost ..."
 
# Read the project file, and deploy it to the folder
[byte[]] $projectFile = [System.IO.File]::ReadAllBytes($IsPacFilePath)
$folder.DeployProject($ProjectName, $projectFile)
 
Write-Host "Creating environment, $environment, in the SSIS Catalog on localhost ..."

# get existing environment if it is there
$environmentInfo = $folder.Environments["$environment"]
 
if ($environmentInfo -eq $null) {
	try {
		$environmentInfo = New-Object $ISNamespace".EnvironmentInfo" ($folder, $environment, "Description")
		$environmentInfo.Create()            
	}
	catch {
		Write-Host "Error: failed to create environment $environment on localhost."
		throw
	}
}
else {	
	# Remove all existing Environment variables before re-adding: there is no update feature. Must delete then re-add. 
    # Probably need to code this to only remove the variables we are replacing, not all existing variables... -blh
	try {
		Write-Host "Removing all existing environment variables..."
		foreach ($variable in $environmentInfo.Variables) {
			$environmentInfo.Variables.Remove($variable)			
		}			
		$environmentInfo.Alter()
	}
	catch {
		Write-Host "Error: Failed to remove existing environment variables."
		throw
	}
}
 
Write-Host "Adding server variables ..."

# Re-add variables to Environment	
try {
	$environmentInfo.Variables.Add("SourceDBConnectionString", [System.TypeCode]::String, "${SourceDBConnectionString}", $false, "SourceDBConnectionString")
	$environmentInfo.Variables.Add("TargetDBConnectionString", [System.TypeCode]::String, "${TargetDBConnectionString}", $false, "TargetDBConnectionString")
	$environmentInfo.Alter()
}
catch {
	Write-Host "Error: Failed adding updated environment variables."
	throw	
}

Write-Host "Adding environment reference to $ProjectName project ..."
 
# add project reference to this environment
$project = $folder.Projects[$ProjectName]
$envRef = $project.References["$environment", "$FolderName"]
if ($envRef -eq $null) {
	try {
		$project.References.Add($environment, $folder.Name)
		$project.Alter() 
	}
	catch {
		Write-Host "Error: failed adding environment referent to $ProjectName project."
		throw
	}
}	

# Reference Environment Variables in the project connection managers
Write-Host "Setting up project connection managers to use Environment Variables..."

try {
	
	$parConnectionString = "CM.SourceDB.ConnectionString"
	$param = $project.Parameters[$parConnectionString]
	if ($param -ne $null) {
		$param.Set("Referenced","SourceDBConnectionString")
	}
	 
	$parConnectionString = "CM.TargetDB.ConnectionString"
	$param = $project.Parameters[$parConnectionString]
	if ($param -ne $null) {
		$param.Set("Referenced","TargetDBConnectionString")
	}
	 
	$project.Alter() 
}
catch {
	Write-Host "Error: failed setting project connection managers to use Environment Variables."
	throw
}

Write-Host "All done."