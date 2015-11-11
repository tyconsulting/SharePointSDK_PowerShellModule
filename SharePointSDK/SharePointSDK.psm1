Function Import-SPClientSDK
{
[OutputType('System.Boolean')]
<# 
 .Synopsis
  Load SharePoint Client SDK DLLs

 .Description
   Load SharePoint Client SDK DLLs from either the Global Assembly Cache or from the DLLs located in SharePointSDK PS module directory. It will use GAC if the DLLs are already loaded in GAC.

 .Example
  # Load the OpsMgr SDK DLLs
  Import-SPClientSDK
#>
    #OpsMgr 2012 R2 SDK DLLs
    $DLLPath = (Get-Module SharePointSDK).ModuleBase
    $arrDLLs = @()
    $arrDLLs += 'Microsoft.SharePoint.Client.dll'
    $arrDLLs += 'Microsoft.SharePoint.Client.Runtime.dll'
	$AssemblyVersion = "15.0.0.0"
	$AssemblyPublicKey = "71e9bce111e9429c"
    #Load SharePoint Client SDKs
    $bSDKLoaded = $true

    Foreach ($DLL in $arrDLLs)
    {
        $AssemblyName = $DLL.TrimEnd('.dll')
        If (!([AppDomain]::CurrentDomain.GetAssemblies() |Where-Object { $_.FullName -eq "$AssemblyName, Version=$AssemblyVersion, Culture=neutral, PublicKeyToken=$AssemblyPublicKey"}))
		{
			Write-verbose 'Loading Assembly $AssemblyName...'
			Try {
                $DLLFilePath = Join-Path $DLLPath $DLL
                [Void][System.Reflection.Assembly]::LoadFrom($DLLFilePath)
            } Catch {
                Write-Verbose "Unable to load $DLLFilePath. Please verify if the SDK DLLs exist in this location!"
                $bSDKLoaded = $false
            }
		}
    }
    $bSDKLoaded
}

Function New-SPCredential
{
[OutputType('Microsoft.SharePoint.Client.SharePointOnlineCredentials')]
[OutputType('System.Net.NetworkCredential')]
<# 
 .Synopsis
  Create a SharePoint credential that can be used authenticating to a SharePoint (online or On-Premise) site.

 .Description
  Create a Network Credential object(System.Net.NetworkCredential) or SharePoint Online credential object(Microsoft.SharePoint.Client.SharePointOnlineCredentials) that can be used authenticating to a SharePoint site. This function will return a Microsoft.SharePoint.Client.SharePointOnlineCredentials object if its going to be used on a SharePoint Online site, or a System.Net.NetworkCredential object if it is to be used on a On-Premise SharePoint site.
  
 .Parameter -SPConnection
  SharePoint SDK Connection object (SMA / Azure Automation connection or hash table).

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Example
  # Create a Credential object using a SMA connection object named "MySPOSite":
  $SPCred = New-SPCredential -SPConnection "MySPOSite"

 .Example
  # Create a Credential for a SharePoint Online site by specifying the user name and password:
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCreds = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $SPCred = New-SPCredential -Credential $MyPSCred -IsSharePointOnlineSite $true

 .Example
  # Create a Credential for a On-Premise SharePoint site by specifying the user credential:
  $Username = "YourDomain\YourUserName"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $SPCred = New-SPCredential -Credential $MyPSCred -IsSPO $false

  .Example
  # Create a Credential for a On-Premise SharePoint site by specifying the user credential and prompt user to enter the password:
  $Username = "YourDomain\YourUserName"
  $SecurePassword = Read-Host -Prompt 'Input password here' -AsSecureString
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username", $SecurePassword)
  $SPCred = New-SPCredential -Credential $MyPSCred -IsSPO $false
#>
    [CmdletBinding()]
    PARAM (
        [Parameter(ParameterSetName='SMAConnection',Mandatory=$true,HelpMessage='Please specify the SMA Connection object')][Alias('Connection','c')][Object]$SPConnection,
        [Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
		[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
    )

    #Firstly, make sure the Microsoft.SharePoint.Client.Dll and Microsoft.SharePoint.Client.Runtime.dll are loaded
	$ImportSDK = Import-SPClientSDK
	If ($ImportSDK -eq $false)
	{
		Write-Error "Unable to load SharePoint Client DLLs. Aborting."
		Return
	}
	
	If ($SPConnection)
	{
		$Username = $SPConnection.Username
		$Password= $SPConnection.Password
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
		$IsSharePointOnlineSite = $SPConnection.IsSharePointOnlineSite
        $Credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)
	} else {
		$Username = $Credential.UserName
        $SecurePassword = $Credential.Password
	}
	If ($IsSharePointOnlineSite)
	{
		Write-Verbose "Creating a SharePointOnlineCredentials object for user $Username - to be used on a SharePoint Online site."
		$SPcredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $SecurePassword)
	} else {
		Write-Verbose "Creating a NetworkCredential object for user $Username - to be used on an On-Premise SharePoint site."
		$SPcredential = $Credential.GetNetworkCredential()
	}
	$SPcredential
}

Function Get-SPServerVersion
{
[OutputType('System.Version')]
<# 
 .Synopsis
  Get SharePoint server version.

 .Description
  Get SharePoint server version using SharePoint CSOM (Client-Side Object Model)
  
 .Parameter -SPConnection
  SharePoint SDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Example
  $ServerVersion = Get-ServerVersion -SPConnection $SPConnection

 .Example
  $Username = "YourDomain\YourUserName"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ServerVersion = Get-SPServerVersion -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSharePointOnlineSite $true

.Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ServerVersion = Get-SPServerVersion -SiteUrl "http://Shrepoint.YourCompany.com" -Credential $MyPSCred -IsSharePointOnlineSite $false
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential
	$Context.executequery()

	#Retrieve Server version
	$ServerVersion = $Context.ServerVersion
	$ServerVersion
}

Function Get-SPListFields
{
[OutputType('System.Array')]
<# 
 .Synopsis
  Get all fields from a list on a SharePoint site.

 .Description
  Get all fields from a list on a SharePoint site.
  
 .Parameter -SPConnection
  SharePoint SDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -List Name
  Name of the list

 .Example
  $ListFields = Get-SPListFields -SPConnection $SPConnection -ListName "Test List"

 .Example
  $Username = "YourDomain\YourUserName"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ListFields = Get-SPListFields -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -ListName "Test List" -IsSharePointOnlineSite $true

.Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ListFields = Get-SPListFields -SiteUrl "http://Shrepoint.YourCompany.com" -Credential $MyPSCred -ListName "Test List" -IsSharePointOnlineSite $false
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Get List fields
	$colListFields = $List.Fields                                                                                    
	$Context.Load($colListFields)                                                                                    
	$Context.ExecuteQuery()

	#add fields to an array
	$ListFields = @()
	Foreach ($item in $colListFields)
	{
		$ListFields +=$item
	}
	,$ListFields
}

Function Add-SPListItem
{
[OutputType('System.Int32')]
<# 
 .Synopsis
  Add a list item to the SharePoint site.

 .Description
  Add a list item to the SharePoint site. When the item has been successfully added, the List Item ID is returned. a NULL value is returned if the item is not added to the list.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListFieldsValues
  A hash table containing List item that to be added.

 .Example
  $HashTableListFieldsValues= @{ "Title", "List Item Title"
	"Description", "List Item Description Text"
  }
  $AddListItem = Add-SPListItem -SPConnection $SPConnection -ListName "Test List" -ListFieldsValues $HashTableListFieldsValues

 .Example
   $HashTableListFieldsValues= @{
    "Title" = "List Item Title"
	"Description" = "List Item Description Text"
  }
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $AddListItem = Add-SPListItem -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSharePointOnlineSite $true -ListName "Test List" -ListFieldsValues $HashTableListFieldsValues

#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the value of each list field in a hash table')][Object]$ListFieldsValues
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Adds an item to the list
	$ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	$Item = $List.AddItem($ListItemInfo)
	Try {
		Foreach ($Field in $ListFieldsValues.Keys)
		{
			$item["$Field"] = $ListFieldsValues.$Field
		}
		$Item.Update()
		$Context.ExecuteQuery()
		$bCreated= $True
	} Catch {
		throw "Unable to add list item to the list $ListName`: " + "`n" + $ListFieldsValues
		$bCreated = $false
	}
	If ($bCreated)
	{
		#Return list item ID when it's successfully created
		$Context.Load($item)
		$Context.ExecuteQuery()
		$ListItemId = $Item.FieldValues.ID
	} else {
		$ListItemId = $NULL
	}
	$ListItemId
}

Function Get-SPListItem
{
[OutputType('System.Array')]
[OutputType('System.Collections.Hashtable')]
<# 
 .Synopsis
  Get all items from a list on a SharePoint site or a specific item by specifying the List Item ID.

 .Description
  Get all items from a list on a SharePoint site or a specific item by specifying the List Item ID. A Hash Table is used to store the property and value of each list item returned.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Example
  $ListItems = Get-SPListItem -SPConnection $SPConnection -ListName "Test List"

 .Example
  $ListItem = Get-SPListItem -SPConnection $SPConnection -ListName "Test List" -ListItemID 1

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ListItems = Get-SPListItem -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List"

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $ListItem = Get-SPListItem -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA Connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(Mandatory=$false,HelpMessage='Please specify the list item ID if retrieving an individual list item')][int]$ListItemId=$null,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	If ($ListItemId -ne $null)
	{
		#Get the list item with the particular ID
		Write-Verbose "List Item ID specified. Retrieving item with ID $ListItemID..."
		Try {
			$ListItem = $List.GetItemById($ListItemID)
			$Context.Load($ListItem)
			$Context.ExecuteQuery()
			$bItemFound = $True
		} Catch {
			Write-Error "Unable to find list item with ID $ListItemID from the list $ListName."
			$bItemFound = $false
		}
		if ($bItemFound)
		{
			$htListItem = @{}
			Foreach ($property in $ListItem.FieldValues.keys)
			{
				$Value = $ListItem.FieldValues.$property
				$htListItem.Add($Property, $Value)
			}
			$htListItem
		}
	} else {
		#Get all List items
		Write-Verbose "List Item ID not specified. Retrieving all items from the list."
		#$camlQuery=  New-Object Microsoft.SharePoint.Client.CamlQuery
		#$camlQuery.ViewXML = "<View />"
		$camlQuery = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(10000)
		$colListItems = $list.GetItems($camlQuery)
		$context.Load($colListItems)
		$Context.ExecuteQuery()

		#add items to an array
		$ListItems = @()
		Foreach ($item in $colListItems)
		{
			$ListItem = @{}
			Foreach ($property in $item.FieldValues.keys)
			{
				$Value = $item.FieldValues.$property
				$ListItem.Add($Property, $Value)
			}
			$ListItems +=$ListItem
		}
		,$ListItems
	}
}

Function Remove-SPListItem
{
[OutputType('System.Boolean')]
<# 
 .Synopsis
  Delete a list item to the SharePoint site.

 .Description
  Delete a list item to the SharePoint site.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListItemID
  The ID of the item to be deleted

 .Example
  $DeleteListItem = Remove-SPListItem -SPConnection $SPConnection -ListName "Test List" -ListItemID 1

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $DeleteListItem = Remove-SPListItem -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA Connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the ID of the list item to be deleted')][int]$ListItemID
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Retrieve the list item
	Try {
		$ListItem = $List.GetItemById($ListItemID)
		$Context.Load($ListItem)
		$Context.ExecuteQuery()
		$bItemFound = $True
	} Catch {
		Write-Error "Unable to find list item with ID $ListItemID from the list $ListName."
		$bItemFound = $false
	}


	#Delete the item
	if ($bItemFound)
	{
		Try {
			$ListItem.DeleteObject()
			$Context.ExecuteQuery()
			$bRemoved= $True
		} Catch {
			Write-Error "Unable to delete the list item (ID: $ListItemID) from the list $ListName."
			$bRemoved = $false
		}
	} else {
		$bRemoved = $false
	}
	$bRemoved
}

Function Update-SPListItem
{
[OutputType('System.Boolean')]
<# 
 .Synopsis
  Update a list item to the SharePoint site.

 .Description
  UPdate a list item to the SharePoint site.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListItemID
  The ID of the item to be deleted

 .Parameter -ListFieldsValues
  A hash table containing List item that to be added.

 .Example
  $HashTableListFieldsValues= @{
    "Title" = "List Item Title"
	"Description" = "List Item Description Text"
  }
  $UpdateListItem = Update-SPListItem -SPConnection $SPConnection -ListName "Test List" -ListItemID 1 -ListFieldsValues $HashTableListFieldsValues

 .Example
  $HashTableListFieldsValues= @{
    "Title" = "List Item Title"
	"Description" = "List Item Description Text"
  }
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $UpdateListItem = Update-SPListItem -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSharePointOnlineSite $true -ListName "Test List" -ListItemID 1 -ListFieldsValues $HashTableListFieldsValues

#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA Connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the ID of the list item to be updated')][int]$ListItemID,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the value of each list field in a hash table')][Object]$ListFieldsValues
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Retrieve the list item
	Try {
		$ListItem = $List.GetItemById($ListItemID)
		$Context.Load($ListItem)
		$Context.ExecuteQuery()
		$bItemFound = $True
	} Catch {
		throw "Unable to find list item with ID $ListItemID from the list $ListName."
		$bItemFound = $false
	}

	#Update the list item
	if ($bItemFound)
	{
		Try {
			Foreach ($Field in $ListFieldsValues.Keys)
			{
				$ListItem["$Field"] = $ListFieldsValues.$Field
			}
			$ListItem.Update()
			$Context.ExecuteQuery()
			$bUpdated= $True
		} Catch {
			throw "Unable to add list item to the list $ListName`: " + "`n" + $ListFieldsValues
			$bUpdated = $false
		}
	} else {
		$bUpdated = $false
	}
	
	$bUpdated
}

Function Get-SPListItemAttachments
{
[OutputType('System.Int32')]
<# 
 .Synopsis
  Download all attachments from a SharePoint list item.

 .Description
  Download all attachments from a SharePoint list item.  Please note this function DOES NOT work on SharePoint 2010 sites.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListItemID
  The ID of the item to be deleted

 .Parameter -DestinationFolder
  The destination folder of where attachments will be saved to.

 .Example
  $DownloadAttachments = Get-SPListItemAttachments -SPConnection $SPConnection -ListName "Test List" -ListItemID 1 -DestinationFolder "\\Server01\ShareFolder"

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $DownloadAttachments = Get-SPListItemAttachments -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1 -Destination "\\Server01\ShareFolder"
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA Connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the ID of the list item')][int]$ListItemID,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the destination folder of where attachments will be saved to')][Alias('Destination')][string]$DestinationFolder
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
		$ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
		$ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}
	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Check SharePoint server version
	If ($ServerVersion.Major -lt 15)
	{
		#Version below SharePoint 2013 (Major version 15)
		Write-Error "Get-SPListItemAttachments DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
		Return
	}

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Retrieve the list item
	Try {
		$ListItem = $List.GetItemById($ListItemID)
		$Context.Load($ListItem)
		$Context.ExecuteQuery()
		$bItemFound = $True
	} Catch {
		Write-Error "Unable to find list item with ID $ListItemID from the list $ListName."
		$bItemFound = $false
	}

	#Download attachments
	$iDownloadCount = 0
	if ($bItemFound)
	{
		#Get Attached Files
		Write-Verbose "Getting attachments for the list item"
		$AttachmentFiles = $ListItem.AttachmentFiles
		$Context.Load($AttachmentFiles)
		$Context.ExecuteQuery()
		$FileCount = $AttachmentFiles.count
		Write-Verbose "Number of attachments found: $FileCount"
		Foreach ($attachment in $AttachmentFiles)
		{
			Write-Verbose "Saving $($attachment.FileName) to $DestinationFolder"
			Try
			{
				$file = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $attachment.ServerRelativeUrl)
				$path = Join-Path $DestinationFolder $attachment.FileName
				$fs = new-object System.IO.FileStream($path, "OpenOrCreate") 
				$file.Stream.CopyTo($fs)
				$fs.Close()
				$iDownloadCount ++
			} catch {
				Write-Error "Unable to save $($attachment.FileName) to $DestinationFolder"
			}
		}

	}
	Write-Verbose "Total number of attachments downloaded: $iDownloadCount"
	$iDownloadCount
}

Function Add-SPListItemAttachment
{
[OutputType('System.Boolean')]
<# 
 .Synopsis
  Upload a file as a SharePoint list item attachment.

 .Description
  Upload a file as a SharePoint list item attachment. Please note this function DOES NOT work on SharePoint 2010 sites.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListItemID
  The ID of the item to be deleted

 .Parameter -FilePath
  The file path of the file to be attached to the list item.

 .Example
  $AddAttachment = Add-SPListItemAttachment -SPConnection $SPConnection -ListName "Test List" -ListItemID 1 -FilePath "\\Server01\ShareFolder\File.txt"

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $AddAttachment = Add-SPListItemAttachment -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1 -FilePath "\\Server01\ShareFolder\File.txt"

 .Example
  $text = "hello world"
  [Byte[]]$Bytes=[System.Text.Encoding]::Default.GetBytes($text)
  $AddAttachment = Add-SPListItemAttachment -SPConnection $SPConnection -ListName "Test List" -ListItemID 1 -ContentByteArray $Bytes -FileName "HelloWord.txt"

 .Example
  [Byte[]]$bytes = [System.IO.File]::ReadAllBytes("C:\Temp\Original.zip")
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username", $SecurePassword)
  $AddAttachment = Add-SPListItemAttachment -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1 -ContentByteArray $Bytes -FileName "renamed.zip"
#>
    [CmdletBinding()]
	Param(
	[Parameter(ParameterSetName="SMAUploadFile",Mandatory=$True,HelpMessage='Please specify the SMA Connection object')]
    [Parameter(ParameterSetName="SMACreateFile",Mandatory=$True,HelpMessage='Please specify the SMA Connection object')]
	[Object]$SPConnection,

	[Parameter(ParameterSetName="IndividualUploadFile",Mandatory=$True,HelpMessage='Please specify the request URL')]
    [Parameter(ParameterSetName="IndividualCreateFile",Mandatory=$True,HelpMessage='Please specify the request URL')]
	[String]$SiteUrl,

	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,

    [Parameter(ParameterSetName="IndividualUploadFile",Mandatory=$True,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName="IndividualCreateFile",Mandatory=$True,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')]	
    [Alias('cred')]
    [PSCredential]$Credential,

	[Parameter(ParameterSetName="IndividualUploadFile",Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName="IndividualCreateFile",Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')]
	[Alias('IsSPO')]
	[boolean]$IsSharePointOnlineSite,

	[Parameter(Mandatory=$true,HelpMessage='Please specify the ID of the list item')][int]$ListItemID,

	[Parameter(ParameterSetName='SMAUploadFile',Mandatory=$true,HelpMessage='Please specify the file path of the file to be attached to the list item.')]
	[Parameter(ParameterSetName="IndividualUploadFile",Mandatory=$true,HelpMessage='Please specify the file path of the file to be attached to the list item.')]
	[Alias('Path')]
	[string]$FilePath,

	[Parameter(ParameterSetName='SMACreateFile',Mandatory=$true,HelpMessage='Please specify the file content byte array.')]
    [Parameter(ParameterSetName="IndividualCreateFile",Mandatory=$true,HelpMessage='Please specify the file content byte array.')]
	[Alias('ByteArray')]
	[Byte[]]$ContentByteArray,

	[Parameter(ParameterSetName='SMACreateFile',Mandatory=$true,HelpMessage='Please specify the file name of the content byte array to be attached to the list item.')]
	[Parameter(ParameterSetName="IndividualCreateFile",Mandatory=$true,HelpMessage='Please specify the file name of the content byte array to be attached to the list item.')]
	[Alias('Name')]
	[string]$FileName
    )
	#firstly, make sure the file exists if uploading an existing file
	if ($FilePath)
	{
		If (Test-Path $FilePath)
		{
			$FileName = Split-Path -leaf $FilePath
			Write-Verbose "File name: $FileName"
		} else {
			throw "$FilePath does not exists!"
			Return
		}
	}

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
		$ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
		$ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}

	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Check SharePoint server version
	If ($ServerVersion.Major -lt 15)
	{
		#Version below SharePoint 2013 (Major version 15)
		Write-Error "Add-SPListItemAttachment DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
		Return
	}

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Retrieve the list item
	Try {
		$ListItem = $List.GetItemById($ListItemID)
		$Context.Load($ListItem)
		$Context.ExecuteQuery()
		$bItemFound = $True
	} Catch {
		Write-Error "Unable to find list item with ID $ListItemID from the list $ListName."
		$bItemFound = $false
	}

	#Upload file
	if ($bItemFound)
	{
		Write-Verbose "Attaching $FilePath to List item (ID: $ListItemID)."
		Try
		{
			#Get existing attachments
			Write-Verbose "Getting existing attachments for the list item"
			$AttachmentFiles = $ListItem.AttachmentFiles
			$Context.Load($AttachmentFiles)
			$Context.ExecuteQuery()
			Write-verbose "Creating Microsoft.SharePoint.Client.AttachmentCreationInformation object"
			$AttachFile = New-Object Microsoft.SharePoint.Client.AttachmentCreationInformation
			If ($FilePath)
			{
				Write-Verbose "Converting the content of `"$FilePath`" to byte array."
				[Byte[]]$bytes = [System.IO.File]::ReadAllBytes($FilePath)
			} else {
				$bytes = $ContentByteArray
			}
			$mStream = New-Object System.IO.MemoryStream @(,$bytes)
			$AttachFile.ContentStream = $mStream
			$AttachFile.FileName = $FileName
			[Void]$AttachmentFiles.Add($AttachFile)
			$Context.Load($AttachmentFiles)
			$Context.ExecuteQuery()
			$bUploaded = $true
			Write-Verbose "$FileName uploaded."
		} Catch {
			Write-Error "Failed to uploade $FilePath to the list Item (ID: $ListItemID)."
			$bUploaded = $false
		}
	}
	Write-Verbose "Done"
	$bUploaded
}

Function Remove-SPListItemAttachment
{
[OutputType('System.Int32')]
<# 
 .Synopsis
  Remove a SharePoint list item attachment.

 .Description
  Remove a SharePoint list item attachment.  Please note this function DOES NOT work on SharePoint 2010 sites.
  
 .Parameter -SPConnection
  SharePointSDK Connection object (SMA connection or hash table).

 .Parameter -SiteUrl
  SharePoint Site Url

 .Parameter -Credential
  The PSCredential object that contains the user name and password required to connect to the SharePoint site.

 .Parameter -IsSharePointOnlineSite
  Specify if the site is a SharePoint Online site

 .Parameter -ListName
  Name of the list

 .Parameter -ListItemID
  The ID of the item to be deleted

 .Parameter -FileName
  The file name of the attachment to be removed from the list item.

 .Example
  $DeleteAttachment = Remove-SPListItemAttachment -SPConnection $SPConnection -ListName "Test List" -ListItemID 1 -FileName "File.txt"

 .Example
  $Username = "you@yourcompany.com"
  $SecurePassword = ConvertTo-SecureString "password1234" -AsPlainText -Force
  $MyPSCred = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
  $DeleteAttachment = Remove-SPListItemAttachment -SiteUrl "https://yourcompany.sharepoint.com" -Credential $MyPSCred -IsSPO $true -ListName "Test List" -ListItemID 1 -FileName "File.txt"
#>
    [CmdletBinding()]
	Param(
    [Parameter(ParameterSetName='SMAConnection',Mandatory=$True,HelpMessage='Please specify the SMA Connection object')][Object]$SPConnection,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify the request URL')][String]$SiteUrl,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the name of the list')][String]$ListName,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$true,HelpMessage='Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')][PSCredential]$Credential,
	[Parameter(ParameterSetName='IndividualParameter',Mandatory=$True,HelpMessage='Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the ID of the list item')][int]$ListItemID,
	[Parameter(Mandatory=$true,HelpMessage='Please specify the file name of the attachment to be removed from the list item')][Alias('Path')][string]$FileName
    )

	If($SPConnection)
	{
		$SPCredential = New-SPCredential -SPConnection $SPConnection
		$SiteUrl = $SPConnection.SharePointSiteURL
		$ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
	} else {
		$SPCredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
		$ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
	}

	#Make sure the SharePoint Client SDK Runtime DLLs are loaded
	$ImportDLL = Import-SPClientSDK

	#Check SharePoint server version
	If ($ServerVersion.Major -lt 15)
	{
		#Version below SharePoint 2013 (Major version 15)
		Write-Error "Remove-SPListItemAttachment DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
		Return
	}

	#Bind to site collection
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
	$Context.Credentials = $SPCredential

	#Retrieve list
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List)
	$Context.ExecuteQuery()

	#Retrieve the list item
	Try {
		$ListItem = $List.GetItemById($ListItemID)
		$Context.Load($ListItem)
		$Context.ExecuteQuery()
		$bItemFound = $True
	} Catch {
		Write-Error "Unable to find list item with ID $ListItemID from the list $ListName."
		$bItemFound = $false
	}

	#delete file
	$iDeleted = 0
	if ($bItemFound)
	{
			#Get attachments
			Write-Verbose "Getting attachments for the list item"
			$AttachmentFiles = $ListItem.AttachmentFiles
			$Context.Load($AttachmentFiles)
			$Context.ExecuteQuery()
			Foreach ($attachment in $AttachmentFiles)
			{
				if ($attachment.FileName -ieq $FileName)
				{
					$ToBedeleted = $attachment
				}
			}
            If ($ToBedeleted)
            {
                #Delete this file
				try
				{
					$ToBedeleted.DeleteObject()
					$iDeleted ++
				} catch {
					Write-Error "Unable to delete file (ServerRelativeUrl: $($attachment.ServerRelativeUrl))"
				}
            }

			#Commit changes
			if ($iDeleted -gt 0)
			{
				$Context.Load($AttachmentFiles)
				$Context.ExecuteQuery()
			}
	}
	Write-Verbose "Number of files deleted: $iDeleted"
	$iDeleted
}