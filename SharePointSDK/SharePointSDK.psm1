# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPCredential
{
  [OutputType('Microsoft.SharePoint.Client.SharePointOnlineCredentials')]
  [OutputType('System.Net.NetworkCredential')]
  [CmdletBinding()]
  PARAM (
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Alias('Connection','c')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )
	
  If ($SPConnection)
  {
    $Username = $SPConnection.Username
    $Password = $SPConnection.Password
    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    $IsSharePointOnlineSite = $SPConnection.IsSharePointOnlineSite
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($Username, $SecurePassword)
  }
  else 
  {
    $Username = $Credential.UserName
    $SecurePassword = $Credential.Password
  }
  If ($IsSharePointOnlineSite)
  {
    Write-Verbose -Message "Creating a SharePointOnlineCredentials object for user $Username - to be used on a SharePoint Online site."
    $SPcredential = New-Object -TypeName Microsoft.SharePoint.Client.SharePointOnlineCredentials -ArgumentList ($Username, $SecurePassword)
  }
  else 
  {
    Write-Verbose -Message "Creating a NetworkCredential object for user $Username - to be used on an On-Premise SharePoint site."
    $SPcredential = $Credential.GetNetworkCredential()
  }
  $SPcredential
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPServerVersion
{
  [OutputType('System.Version')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $Context.executequery()

  #Retrieve Server version
  $ServerVersion = $Context.ServerVersion
  $ServerVersion
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPListFields
{
  [OutputType('System.Array')]
  [OutputType('System.Array')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][Alias('ListTitle')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

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
    $ListFields += $item
  }
  ,$ListFields
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Add-SPListItem
{
  [OutputType('System.Int32')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the value of each list field in a hash table')][Object]$ListFieldsValues
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Adds an item to the list
  $ListItemInfo = New-Object -TypeName Microsoft.SharePoint.Client.ListItemCreationInformation
  $item = $List.AddItem($ListItemInfo)
  Try 
  {
    Foreach ($Field in $ListFieldsValues.Keys)
    {
      $item["$Field"] = $ListFieldsValues.$Field
    }
    $item.Update()
    $Context.ExecuteQuery()
    $bCreated = $true
  }
  Catch 
  {
    throw "Unable to add list item to the list $ListName`: " + "`n" + $ListFieldsValues
    $bCreated = $false
  }
  If ($bCreated)
  {
    #Return list item ID when it's successfully created
    $Context.Load($item)
    $Context.ExecuteQuery()
    $ListItemId = $item.FieldValues.ID
  }
  else 
  {
    $ListItemId = $NULL
  }
  $ListItemId
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPListItem
{
  [OutputType('System.Array')]
  [OutputType('System.Collections.Hashtable')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the list item ID if retrieving an individual list item')][int]$ListItemId = $NULL,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Work out if $ListItemId actually contains value
  If ($PSBoundParameters.ContainsKey('ListItemId'))
  {
    $ListItemId = $PSBoundParameters.ListItemId
  }
  else 
  {
    Remove-Variable -Name ListItemId
  }

  If ($ListItemId -ne $NULL)
  {
    #Get the list item with the particular ID
    Write-Verbose -Message "List Item ID specified. Retrieving item with ID $ListItemId..."
    Try 
    {
      $ListItem = $List.GetItemById($ListItemId)
      $Context.Load($ListItem)
      $Context.ExecuteQuery()
      $bItemFound = $true
    }
    Catch 
    {
      Write-Error -Message "Unable to find list item with ID $ListItemId from the list $ListName."
      $bItemFound = $false
    }
    if ($bItemFound)
    {
      $htListItem = @{}
      Foreach ($property in $ListItem.FieldValues.keys)
      {
        $Value = $ListItem.FieldValues.$property
        $htListItem.Add($property, $Value)
      }
      $htListItem
    }
  }
  else 
  {
    #Get all List items
    Write-Verbose -Message 'List Item ID not specified. Retrieving all items from the list.'
    #$camlQuery=  New-Object Microsoft.SharePoint.Client.CamlQuery
    #$camlQuery.ViewXML = "<View />"
    $camlQuery = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(10000)
    $colListItems = $List.GetItems($camlQuery)
    $Context.Load($colListItems)
    $Context.ExecuteQuery()

    #add items to an array
    $ListItems = @()
    Foreach ($item in $colListItems)
    {
      $ListItem = @{}
      Foreach ($property in $item.FieldValues.keys)
      {
        $Value = $item.FieldValues.$property
        $ListItem.Add($property, $Value)
      }
      $ListItems += $ListItem
    }
    ,$ListItems
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPListItem
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the ID of the list item to be deleted')][int]$ListItemId
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Retrieve the list item
  Try 
  {
    $ListItem = $List.GetItemById($ListItemId)
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
    $bItemFound = $true
  }
  Catch 
  {
    Write-Error -Message "Unable to find list item with ID $ListItemId from the list $ListName."
    $bItemFound = $false
  }


  #Delete the item
  if ($bItemFound)
  {
    Try 
    {
      $ListItem.DeleteObject()
      $Context.ExecuteQuery()
      $bRemoved = $true
    }
    Catch 
    {
      Write-Error -Message "Unable to delete the list item (ID: $ListItemId) from the list $ListName."
      $bRemoved = $false
    }
  }
  else 
  {
    $bRemoved = $false
  }
  $bRemoved
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Update-SPListItem
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the ID of the list item to be updated')][int]$ListItemId,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the value of each list field in a hash table')][Object]$ListFieldsValues
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Retrieve the list item
  Try 
  {
    $ListItem = $List.GetItemById($ListItemId)
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
    $bItemFound = $true
  }
  Catch 
  {
    throw "Unable to find list item with ID $ListItemId from the list $ListName."
    $bItemFound = $false
  }

  #Update the list item
  if ($bItemFound)
  {
    Try 
    {
      Foreach ($Field in $ListFieldsValues.Keys)
      {
        $ListItem["$Field"] = $ListFieldsValues.$Field
      }
      $ListItem.Update()
      $Context.ExecuteQuery()
      $bUpdated = $true
    }
    Catch 
    {
      throw "Unable to add list item to the list $ListName`: " + "`n" + $ListFieldsValues
      $bUpdated = $false
    }
  }
  else 
  {
    $bUpdated = $false
  }
	
  $bUpdated
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPListItemAttachments
{
  [OutputType('System.Int32')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the ID of the list item')][int]$ListItemId,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the destination folder of where attachments will be saved to')][Alias('Destination')][string]$DestinationFolder
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
    $ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
  }
  else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
    $ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Check SharePoint server version
  If ($ServerVersion.Major -lt 15)
  {
    #Version below SharePoint 2013 (Major version 15)
    Write-Error -Message "Get-SPListItemAttachments DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
    Return
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Retrieve the list item
  Try 
  {
    $ListItem = $List.GetItemById($ListItemId)
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
    $bItemFound = $true
  }
  Catch 
  {
    Write-Error -Message "Unable to find list item with ID $ListItemId from the list $ListName."
    $bItemFound = $false
  }

  #Download attachments
  $iDownloadCount = 0
  if ($bItemFound)
  {
    #Get Attached Files
    Write-Verbose -Message 'Getting attachments for the list item'
    $AttachmentFiles = $ListItem.AttachmentFiles
    $Context.Load($AttachmentFiles)
    $Context.ExecuteQuery()
    $FileCount = $AttachmentFiles.count
    Write-Verbose -Message "Number of attachments found: $FileCount"
    Foreach ($attachment in $AttachmentFiles)
    {
      Write-Verbose -Message "Saving $($attachment.FileName) to $DestinationFolder"
      Try
      {
        $file = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $attachment.ServerRelativeUrl)
        $path = Join-Path -Path $DestinationFolder -ChildPath $attachment.FileName
        $fs = New-Object -TypeName System.IO.FileStream -ArgumentList ($path, 'OpenOrCreate') 
        $file.Stream.CopyTo($fs)
        $fs.Close()
        $iDownloadCount ++
      }
      catch 
      {
        Write-Error -Message "Unable to save $($attachment.FileName) to $DestinationFolder"
      }
    }
  }
  Write-Verbose -Message "Total number of attachments downloaded: $iDownloadCount"
  $iDownloadCount
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Add-SPListItemAttachment
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAUploadFile',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')]
    [Parameter(ParameterSetName = 'SMACreateFile',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualUploadFile',Mandatory = $true,HelpMessage = 'Please specify the request URL')]
    [Parameter(ParameterSetName = 'IndividualCreateFile',Mandatory = $true,HelpMessage = 'Please specify the request URL')]
    [String]$SiteUrl,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,

    [Parameter(ParameterSetName = 'IndividualUploadFile',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualCreateFile',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]	
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualUploadFile',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualCreateFile',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')]
    [boolean]$IsSharePointOnlineSite,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the ID of the list item')][int]$ListItemId,

    [Parameter(ParameterSetName = 'SMAUploadFile',Mandatory = $true,HelpMessage = 'Please specify the file path of the file to be attached to the list item.')]
    [Parameter(ParameterSetName = 'IndividualUploadFile',Mandatory = $true,HelpMessage = 'Please specify the file path of the file to be attached to the list item.')]
    [Alias('Path')]
    [string]$FilePath,

    [Parameter(ParameterSetName = 'SMACreateFile',Mandatory = $true,HelpMessage = 'Please specify the file content byte array.')]
    [Parameter(ParameterSetName = 'IndividualCreateFile',Mandatory = $true,HelpMessage = 'Please specify the file content byte array.')]
    [Alias('ByteArray')]
    [Byte[]]$ContentByteArray,

    [Parameter(ParameterSetName = 'SMACreateFile',Mandatory = $true,HelpMessage = 'Please specify the file name of the content byte array to be attached to the list item.')]
    [Parameter(ParameterSetName = 'IndividualCreateFile',Mandatory = $true,HelpMessage = 'Please specify the file name of the content byte array to be attached to the list item.')]
    [Alias('Name')]
    [string]$FileName
  )
  #firstly, make sure the file exists if uploading an existing file
  if ($FilePath)
  {
    If (Test-Path $FilePath)
    {
      $FileName = Split-Path -Leaf -Path $FilePath
      Write-Verbose -Message "File name: $FileName"
    }
    else 
    {
      throw "$FilePath does not exists!"
      Return
    }
  }

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
    $ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
  }
  else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
    $ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Check SharePoint server version
  If ($ServerVersion.Major -lt 15)
  {
    #Version below SharePoint 2013 (Major version 15)
    Write-Error -Message "Add-SPListItemAttachment DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
    Return
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Retrieve the list item
  Try 
  {
    $ListItem = $List.GetItemById($ListItemId)
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
    $bItemFound = $true
  }
  Catch 
  {
    Write-Error -Message "Unable to find list item with ID $ListItemId from the list $ListName."
    $bItemFound = $false
  }

  #Upload file
  if ($bItemFound)
  {
    Write-Verbose -Message "Attaching $FilePath to List item (ID: $ListItemId)."
    Try
    {
      #Get existing attachments
      Write-Verbose -Message 'Getting existing attachments for the list item'
      $AttachmentFiles = $ListItem.AttachmentFiles
      $Context.Load($AttachmentFiles)
      $Context.ExecuteQuery()
      Write-Verbose -Message 'Creating Microsoft.SharePoint.Client.AttachmentCreationInformation object'
      $AttachFile = New-Object -TypeName Microsoft.SharePoint.Client.AttachmentCreationInformation
      If ($FilePath)
      {
        Write-Verbose -Message "Converting the content of `"$FilePath`" to byte array."
        [Byte[]]$bytes = [System.IO.File]::ReadAllBytes($FilePath)
      }
      else 
      {
        $bytes = $ContentByteArray
      }
      $mStream = New-Object -TypeName System.IO.MemoryStream -ArgumentList @(,$bytes)
      $AttachFile.ContentStream = $mStream
      $AttachFile.FileName = $FileName
      [Void]$AttachmentFiles.Add($AttachFile)
      $Context.Load($AttachmentFiles)
      $Context.ExecuteQuery()
      $bUploaded = $true
      Write-Verbose -Message "$FileName uploaded."
    }
    Catch 
    {
      Write-Error -Message "Failed to uploade $FilePath to the list Item (ID: $ListItemId)."
      $bUploaded = $false
    }
  }
  Write-Verbose -Message 'Done'
  $bUploaded
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPListItemAttachment
{
  [OutputType('System.Int32')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion Connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the request URL')][String]$SiteUrl,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the list')][String]$ListName,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the ID of the list item')][int]$ListItemId,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the file name of the attachment to be removed from the list item')][Alias('Path')][string]$FileName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
    $ServerVersion = Get-SPServerVersion -SPConnection  $SPConnection
  }
  else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
    $ServerVersion = Get-SPServerVersion -SiteUrl $SiteUrl -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Check SharePoint server version
  If ($ServerVersion.Major -lt 15)
  {
    #Version below SharePoint 2013 (Major version 15)
    Write-Error -Message "Remove-SPListItemAttachment DOES NOT work on pre SharePoint 2013 versions. Current Version: $ServerVersion"
    Return
  }

  #Bind to site collection
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential

  #Retrieve list
  $List = $Context.Web.Lists.GetByTitle($ListName)
  $Context.Load($List)
  $Context.ExecuteQuery()

  #Retrieve the list item
  Try 
  {
    $ListItem = $List.GetItemById($ListItemId)
    $Context.Load($ListItem)
    $Context.ExecuteQuery()
    $bItemFound = $true
  }
  Catch 
  {
    Write-Error -Message "Unable to find list item with ID $ListItemId from the list $ListName."
    $bItemFound = $false
  }

  #delete file
  $iDeleted = 0
  if ($bItemFound)
  {
    #Get attachments
    Write-Verbose -Message 'Getting attachments for the list item'
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
      }
      catch 
      {
        Write-Error -Message "Unable to delete file (ServerRelativeUrl: $($attachment.ServerRelativeUrl))"
      }
    }

    #Commit changes
    if ($iDeleted -gt 0)
    {
      $Context.Load($AttachmentFiles)
      $Context.ExecuteQuery()
    }
  }
  Write-Verbose -Message "Number of files deleted: $iDeleted"
  $iDeleted
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPList
{
  [OutputType('Microsoft.SharePoint.Client.List')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the title of the list')][ValidateNotNullOrEmpty()][String]$ListTitle,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the description of the list')][String]$ListDescription,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if the list should be displayed on Quick Launch')][Boolean]$QuickLaunch = $false,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the list template type')][String]
    [ValidateScript(
        {
          if ([Microsoft.SharePoint.Client.ListTemplateType]::$_)
          {
            $true
          }
          else 
          {
            $false
          }
        }
    )]$ListTemplateType = 'GenericList'
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  #Create list
  Write-Verbose -Message "Creating 'Microsoft.SharePoint.Client.ListCreationInformation' object for list `'$ListTitle`'."
  $CreationInfo = New-Object -TypeName Microsoft.SharePoint.Client.ListCreationInformation
  $CreationInfo.Title = $ListTitle
  if ($ListDescription.Length -gt 0)
  {
    Write-Verbose -Message "List Description: '$ListDescription'."
    $CreationInfo.Description = $ListDescription
  }
  $CreationInfo.TemplateType = [Microsoft.SharePoint.Client.ListTemplateType]::$ListTemplateType
  try
  {
    #Create the list
    Write-Verbose -Message "Creating list '$ListTitle'."
    $List = $web.lists.Add($CreationInfo)
    $List.Update()
    $Context.ExecuteQuery()
    #Quick Launch
    Write-Verbose -Message "List on Quick Launch: '$QuickLaunch'."
    If ($QuickLaunch -eq $true)
    {
      $List.OnQuickLaunch = $QuickLaunch
      $List.Update()
      $Context.ExecuteQuery()
    }
    #Retrieve the list
    Write-Verbose -Message 'Retrieving the list.'
    $Context.Load($List)
    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    Write-Verbose -Message 'Returning the list.'
    $List
  }
  catch 
  {
    throw $_.Exception.InnerException
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPList
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Delete the list
  Write-Verbose -Message 'Deleting the list'
  try 
  {
    $List.DeleteObject()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPList
{
  [OutputType('Microsoft.SharePoint.Client.List')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  $Context.Load($List)
  $Context.Load($List.Fields)
  $Context.ExecuteQuery()
  Return $List
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListLookupField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Id of the source list of where the lookup field is getting information from')]
    [ValidateNotNullOrEmpty()][Guid]$SourceListId,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the Field Relationship Delete Behavior')]
    [ValidateSet('None', 'Restrict', 'Cascade')][String]$RelationshipDeleteBehavior = 'None',

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the field to be displayed')]
    [ValidateNotNullOrEmpty()][String]$ShowField,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the additional fields from the source list to be added')]
    [ValidateNotNullOrEmpty()][String[]]$AdditionalSourceFields,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
  $Context.Load($web)
  $Context.ExecuteQuery()
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #set the source list to itself if the SourceListId is not specified
  if (!$SourceListId)
  {
    $SourceListId = $strListId
  }
  else 
  {
    $SourceListId = $SourceListId.ToString()
  }
    
  #Index
  if ($RelationshipDeleteBehavior -ieq 'cascade' -or $RelationshipDeleteBehavior -ieq 'restrict' -or $EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
  Write-Verbose -Message "List Id: '$strListId'."
  $LookupFieldGUID = [System.Guid]::NewGuid().ToString()
  $fieldXML = "<Field Type='Lookup' ID='{$LookupFieldGUID}' DisplayName='$FieldDisplayName' List='{$SourceListId}' ShowField='$ShowField' RelationshipDeleteBehavior='$RelationshipDeleteBehavior' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())'/>"
    
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $Context.Load($ThisField)
    $Context.ExecuteQuery()
    If ($AdditionalSourceFields.count -gt 0)
    {
      Write-Verbose -Message 'Adding additional lookup fields from the source list.'
      Write-Verbose -Message 'Retrieving the source list first.'
      $SourceList = $web.Lists.GetById($SourceListId)
      $Context.Load($SourceList)
      $Context.Load($SourceList.Fields)
      $Context.ExecuteQuery()
      Write-Verbose -Message "Source List Title: '$($SourceList.Title)'."
      Foreach ($AdditionalFieldDisplayName in $AdditionalSourceFields)
      {
        $WebId = $web.Id
        Write-Verbose -Message "Additional field name: '$AdditionalFieldName'."
        $SourceField = $SourceList.Fields.GetByInternalNameOrTitle($AdditionalFieldDisplayName)
        $Context.Load($SourceField)
        $Context.ExecuteQuery()
        #$SourceFieldVersion= $([xml]$SourceField.SchemaXml).Field.Version
        $SourceFieldStaticName = $SourceField.StaticName
        $SourceFieldInternalName = $SourceField.InternalName
        Write-Verbose -Message "Source Field Static Name: '$SourceFieldStaticName'."
        $AdditionalFieldGUID = [System.Guid]::NewGuid().ToString()
        $AdditionalFieldName = "$FieldDisplayName`_$SourceFieldStaticName"
        $AdditionalFieldXML = "<Field Type='Lookup' DisplayName='$FieldDisplayName`:$AdditionalFieldDisplayName' List='{$SourceListId}' WebId='$WebId' ShowField='$SourceFieldInternalName' FieldRef='$LookupFieldGUID' ReadOnly='TRUE' UnlimitedLengthInDocumentLibrary='FALSE' ID='{$AdditionalFieldGUID}' SourceID='{$($List.Id.ToString())}' StaticName='$AdditionalFieldName' Name='$AdditionalFieldName'/>"
        #$AdditionalFieldXML = "<Field Type='Lookup' DisplayName='$FieldDisplayName`:$AdditionalFieldDisplayName' List='{$SourceListId}' WebId='$WebId' ShowField='$SourceFieldInternalName' FieldRef='$LookupFieldGUID' ReadOnly='TRUE' UnlimitedLengthInDocumentLibrary='FALSE' ID='{$AdditionalFieldGUID}' SourceID='{$($List.Id.ToString())}' StaticName='$AdditionalFieldName' Name='$AdditionalFieldName' Version='$SourceFieldVersion'/>"
                
        Write-Verbose -Message "Additional field XML: '$AdditionalFieldXML'"
        if ($AddToDefaultView)
        {
          Write-Verbose -Message 'The additional field will be added to the default view.'
          $List.Fields.AddFieldAsXml($AdditionalFieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
        }
        else 
        {
          Write-Verbose -Message 'The additional field will NOT be added to the default view.'
          $List.Fields.AddFieldAsXml($AdditionalFieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
        }
        $List.Update()
        $Context.ExecuteQuery()
      }
    }

    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListCheckboxField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is checked by default')]
    [ValidateNotNullOrEmpty()][Boolean]$CheckedByDefault = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()
    
  #Determine the integer for the default value
  if ($CheckedByDefault)
  {
    $iDefault = 1
  }
  else 
  {
    $iDefault = 0
  }

  Write-Verbose -Message "List Id: '$strListId'."
  $fieldXML = "<Field Type='Boolean' DisplayName='$FieldDisplayName' Name='$FieldName' Description='$FieldDescription'><Default>$iDefault</Default></Field>"
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListSingleLineTextField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the maximum number of characters for this Field')]
    [ValidateRange(1,255)][Int]$MaxLength = 255,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][String]$DefaultValue,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #Index
  if ($EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
        
  Write-Verbose -Message "List Id: '$strListId'."
  If ($DefaultValue.Length -eq 0)
  {
    $fieldXML = "<Field Type='Text' DisplayName='$FieldDisplayName' Name='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())' MaxLength='$MaxLength' />"
  }
  else 
  {
    $fieldXML = "<Field Type='Text' DisplayName='$FieldDisplayName' Name='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())' MaxLength='$MaxLength' ><Default>$DefaultValue</Default></Field>"
  } 
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListMultiLineTextField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the number of lines for this Field')]
    [ValidateScript({
          $_ -ge 1
    })][Int]$NumberOfLines = 6,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field can contain Rich Text')]
    [ValidateNotNullOrEmpty()][Boolean]$RichText = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the rich text type for the Field')]
    [ValidateScript({
          if ($RichText -eq $true)
          {
            $_ -ieq 'compatiple' -or $_ -ieq 'fullhtml'
          }
    })]
    [ValidateSet('Compatiple','FullHTML')]
    [String]$RichTextMode,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()
        
  Write-Verbose -Message "List Id: '$strListId'."
  Write-Verbose -Message "Rich Text: $RichText"
  If ($RichText -eq $false)
  {
    $fieldXML = "<Field Type='Note' DisplayName='$FieldDisplayName' Name='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' Sortable='FALSE' StaticName='$FieldName' NumLines='$NumberOfLines' RichText='FALSE'/>"
  }
  else 
  {
    switch ($RichTextMode.ToLower())
    {
      'compatiple' 
      {
        $strRichTextMode = 'Compatiple'
      }
      'fullhtml' 
      {
        $strRichTextMode = 'FullHtml'
      }
    }
    $fieldXML = "<Field Type='Note' DisplayName='$FieldDisplayName' Name='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' Sortable='FALSE' StaticName='$FieldName' NumLines='$NumberOfLines' RichText='TRUE' RichTextMode='$strRichTextMode'/>"
  } 
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListNumberField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the minimum value of this Field')]
    [ValidateNotNullOrEmpty()][Alias('min')][Double]$Minimum,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the maximum value of this Field')]
    [ValidateNotNullOrEmpty()][Alias('max')][Double]$Maximum,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the number of decimal places')]
    [ValidateNotNullOrEmpty()][Int][Alias('decimal')]$DecimalPlaces,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be shown as percentage')]
    [ValidateNotNullOrEmpty()][Boolean]$Percentage = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][String]$DefaultValue,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #Index
  if ($EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
        
  Write-Verbose -Message "List Id: '$strListId'."
  #construct the field XML
  $fieldXML = "<Field Type='Number' DisplayName='$FieldDisplayName' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())'"
  if ($PSBoundParameters.ContainsKey('Minimum'))
  #if ($Minimum)
  {
    $fieldXML = "$fieldXML Min='$Minimum'"
  }
  if ($PSBoundParameters.ContainsKey('Maximum'))
  #if ($Maximum)
  {
    $fieldXML = "$fieldXML Max='$Maximum'"
  }
  if ($PSBoundParameters.ContainsKey('DecimalPlaces'))
  #if ($DecimalPlaces)
  {
    $fieldXML = "$fieldXML Decimals='$DecimalPlaces'"
  }
  if ($Percentage)
  {
    $fieldXML = "$fieldXML Percentage='TRUE'"
  }
  if ($PSBoundParameters.ContainsKey('DefaultValue'))
  #if ($DefaultValue)
  {
    $fieldXML = "$fieldXML><Default>$DefaultValue</Default></Field>"
  }
  else 
  {
    $fieldXML = "$fieldXML />"
  }
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListChoiceField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the minimum value of this Field')]
    [ValidateSet('DropDown', 'RadioButtons', 'CheckBoxes')][String]$Style,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify an array containing a list of choices')]
    [ValidateNotNullOrEmpty()][String[]]$Choices,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateScript({
          $Choices.Contains($_)
    })][String]$DefaultValue,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if Fill-in choices are allowed')]
    [ValidateNotNullOrEmpty()][Boolean]$FillInChoice = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #Index
  if ($EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
        
  Write-Verbose -Message "List Id: '$strListId'."
  #construct the field XML
  $fieldXML = "<Field DisplayName='$FieldDisplayName' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())' FillInChoice='$($FillInChoice.ToString().ToUpper())'"
  Switch ($Style.ToLower())
  {
    'dropdown'
    {
      $fieldXML = "$fieldXML Type='Choice' Format='Dropdown'"
    }
    'radiobuttons'
    {
      $fieldXML = "$fieldXML Type='Choice' Format='RadioButtons'"
    }
    'checkboxes'
    {
      $fieldXML = "$fieldXML Type='MultiChoice'"
    }
  }
  $fieldXML = "$fieldXML>
  <CHOICES>"
  Foreach ($item in $Choices)
  {
    $fieldXML = "$fieldXML
    <CHOICE>$item</CHOICE>"
  }
  $fieldXML = "$fieldXML
    </CHOICES>
  "
  If ($DefaultValue)
  {
    $fieldXML = "$fieldXML
      <Default>$DefaultValue</Default>
      </Field>
    "
  }
  else 
  {
    $fieldXML = "$fieldXML</Field>
    "
  }
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  Try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field '$FieldDisplayName' will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field '$FieldDisplayName' will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field '$FieldDisplayName' will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $Context.ExecuteQuery()
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $Context.ExecuteQuery()
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListDateTimeField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Field Style ("DateTime" or "DateOnly")')]
    [ValidateSet('DateTime', 'DateOnly')][String]$Style,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if the Friendly display format should be used')]
    [ValidateNotNullOrEmpty()][Boolean]$FriendlyDisplay = $false,

    [Parameter(Mandatory = $false,HelpMessage = "Please specify if use today's date as the default value.")]
    [ValidateNotNullOrEmpty()][Boolean]$UseTodayAsDefaultValue = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($ListId) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($PSBoundParameters.ContainsKey('ListId'))
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #Index
  if ($EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
        
  Write-Verbose -Message "List Id: '$strListId'."
  #construct the field XML
  $fieldXML = "<Field Type='DateTime' DisplayName='$FieldDisplayName' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())'"
  Switch ($Style.ToLower())
  {
    'dateonly'
    {
      $fieldXML = "$fieldXML Format='DateOnly'"
    }
    'datetime'
    {
      $fieldXML = "$fieldXML Format='DateTime'"
    }
  }
  #Friendly display format
  If ($FriendlyDisplay -eq $true)
  {
    $fieldXML = "$fieldXML FriendlyDisplayFormat='Relative'>"
  }

  #Default value
  If ($UseTodayAsDefaultValue -eq $true)
  {
    $fieldXML = "$fieldXML><Default>[today]</Default></Field>"
  }
  else 
  {
    $fieldXML = "$fieldXML />"
  }
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListHyperLinkField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the minimum value of this Field')]
    [ValidateSet('Hyperlink', 'Picture')][String]$Style = 'Hyperlink',

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()
    
  Write-Verbose -Message "List Id: '$strListId'."
  #construct the field XML
  $fieldXML = "<Field Type='URL' DisplayName='$FieldDisplayName' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())'"
  Switch ($Style.ToLower())
  {
    'hyperlink'
    {
      $fieldXML = "$fieldXML Format='Hyperlink'/>"
    }
    'picture'
    {
      $fieldXML = "$fieldXML Format='Image'/>"
    }
  }
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPListPersonField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the display name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDisplayName')][String]$FieldDisplayName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the Field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnDescription')][String]$FieldDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field is a required Field')]
    [ValidateNotNullOrEmpty()][Boolean]$Required = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should enforce unique values')]
    [ValidateNotNullOrEmpty()][Alias('unique')][Boolean]$EnforceUniqueValues = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the User Selection Mode')]
    [ValidateSet('PeopleAndGroups', 'PeopleOnly')][String]$UserSelectionMode = 'PeopleAndGroups',

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the User Selection Mode')]
    [Alias('UserSelectionScope','From')][Int]$FromGroupId = 0,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if selecting multiple users or groups is allowed')]
    [ValidateScript({
          If ($EnforceUniqueValues -eq $true)
          {
            $_ -ne $true
          }
          else 
          {
            $true
          }
    })][Boolean]$AllowMultiple = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if this Field should be added to the default view')]
    [ValidateNotNullOrEmpty()][Boolean]$AddToDefaultView = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInEditForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInNewForm = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form')]
    [ValidateNotNullOrEmpty()][Boolean]$ShowInDisplayForm = $true
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }
  #Load the list
  $Context.Load($List)
  $Context.ExecuteQuery()
  $strListId = $List.Id.ToString()

  #Index
  if ($EnforceUniqueValues -eq $true)
  {
    $bIndexed = $true
  }
  else 
  {
    $bIndexed = $false
  }
        
  Write-Verbose -Message "List Id: '$strListId'."
  #construct the field XML
  $fieldXML = "<Field DisplayName='$FieldDisplayName' Name='$FieldName' StaticName='$FieldName' Description='$FieldDescription' Required='$($Required.ToString().ToUpper())' EnforceUniqueValues='$($EnforceUniqueValues.ToString().ToUpper())' Indexed='$($bIndexed.ToString().ToUpper())' UserSelectionScope='$FromGroupId' Mult='$($AllowMultiple.ToString().ToUpper())'"
  If ($AllowMultiple -eq $true)
  {
    $fieldXML = "$fieldXML Type='UserMulti'"
  }
  else 
  {
    $fieldXML = "$fieldXML Type='User'"
  }

  Switch ($UserSelectionMode.ToLower())
  {
    'peopleandgroups'
    {
      $fieldXML = "$fieldXML UserSelectionMode='PeopleAndGroups'/>"
    }
    'peopleonly'
    {
      $fieldXML = "$fieldXML UserSelectionMode='PeopleOnly'/>"
    }
  }
  Write-Verbose -Message "Field XML: `"$fieldXML`""
  Write-Verbose -Message 'Adding the Field to the list'
    
  try
  {
    if ($AddToDefaultView)
    {
      Write-Verbose -Message 'The Field will be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    }
    else 
    {
      Write-Verbose -Message 'The Field will NOT be added to the default view.'
      $List.Fields.AddFieldAsXml($fieldXML,$false, [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
    }
    $List.Update()
    $Context.ExecuteQuery()

    #Update the field
    Write-Verbose -Message "Field will show in display form: $ShowInDisplayForm"
    Write-Verbose -Message "Field will show in edit form: $ShowInEditForm"
    Write-Verbose -Message "Field will show in new form: $ShowInNewForm"

    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
    $ThisField = $List.Fields.GetByInternalNameOrTitle($FieldDisplayName)
    $ThisField.SetShowInDisplayForm($ShowInDisplayForm)
    $ThisField.SetShowInEditForm($ShowInEditForm)
    $ThisField.SetShowInNewForm($ShowInNewForm)
    $ThisField.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    Throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPListField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Field name')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName','ColumnTitle','FieldTitle')][String]$FieldName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return
  }

  #Get the field
  Write-Verbose -Message 'Deleting the Field'
  try
  {
    $Field = $List.Fields.GetByInternalNameOrTitle($FieldName)
    $Field.DeleteObject()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Update-SPListField
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Field name')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName','ColumnTitle','FieldTitle')][String]$FieldName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the updated field schema XML')]
    [ValidateNotNullOrEmpty()][Alias('schema')][String]$SchemaXML
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return $false
  }

  try
  {
    $Context.Load($List)
    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
  }
  Catch 
  {
    throw $_.Exception.InnerException
    return $false
  }

  #Get the field
  Write-Verbose -Message "Updating the Field '$FieldName'"
  try
  {
    Write-Verbose -Message "New Schema XML: '$SchemaXML'"
    $Field = $List.Fields.GetByInternalNameOrTitle($FieldName)
    $Context.Load($Field)
    $Context.ExecuteQuery()
    $Field.SchemaXml = $SchemaXML
    $Field.Update()
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Set-SPListFieldVisibility
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Field name')]
    [ValidateNotNullOrEmpty()][Alias('ColumnName','ColumnTitle','FieldTitle')][String]$FieldName,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the edit form.')]
    [Boolean]$ShowInEditForm,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the new form.')]
    [Boolean]$ShowInNewForm,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the if this field should be shown in the display form.')]
    [Boolean]$ShowInDisplayForm
  )

  #Retrieve field visibility parameters from $PSBoundParameters
  Write-Verbose -Message 'Determining field visibility configurations.'
  If (!($PSBoundParameters.ContainsKey('ShowInEditForm')) -and !($PSBoundParameters.ContainsKey('ShowInNewForm')) -and !($PSBoundParameters.ContainsKey('ShowInDisplayForm')))
  {
    Throw "Please specify at least one of the following parameters: 'ShowInEditForm', 'ShowInNewForm', 'ShowInDisplayForm'"
    Return $false
  }

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return $false
  }

  try
  {
    $Context.Load($List)
    $Context.Load($List.Fields)
    $Context.ExecuteQuery()
  }
  Catch 
  {
    throw $_.Exception.InnerException
    return $false
  }

  #Get the field
  Write-Verbose -Message "Updating the Field '$FieldName'"
  try
  {
    $Field = $List.Fields.GetByInternalNameOrTitle($FieldName)
    #ShowInEditForm
    If ($PSBoundParameters.ContainsKey('ShowInEditForm'))
    {
      Write-Verbose -Message "Setting ShowInEditForm to '$ShowInEditForm'"
      $Field.SetShowInEditForm($ShowInEditForm)
      $Context.ExecuteQuery()
    }

    #ShowInNewForm
    If ($PSBoundParameters.ContainsKey('ShowInNewForm'))
    {
      Write-Verbose -Message "Setting ShowInNewForm to '$ShowInNewForm'"
      $Field.SetShowInNewForm($ShowInNewForm)
      $Context.ExecuteQuery()
    }

    #ShowInDisplayForm
    If ($PSBoundParameters.ContainsKey('ShowInDisplayForm'))
    {
      Write-Verbose -Message "Setting ShowInDisplayForm to '$ShowInDisplayForm'"
      $Field.SetShowInDisplayForm($ShowInDisplayForm)
      $Context.ExecuteQuery()
    }
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPUser
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the user email')]
    [ValidateNotNullOrEmpty()][String]$Email,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the user login name')]
    [ValidateNotNullOrEmpty()][String]$LoginName,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the user title')]
    [ValidateNotNullOrEmpty()][String]$Title,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if the user should be granted Site Visitors rights.')]
    [ValidateNotNullOrEmpty()][Boolean]$IsSiteVisitor = $true,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if the user should be granted Site Members rights.')]
    [ValidateNotNullOrEmpty()][Boolean]$IsSiteMember = $false,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify if the user should be granted Site Owners rights.')]
    [ValidateNotNullOrEmpty()][Boolean]$IsSiteOwner = $false
  )
  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  #Determine the default group membership
  If ($IsSiteOwner)
  {
    $GroupId = 6
  }
  elseif ($IsSiteMember)
  {
    $GroupId = 8
  }
  elseif ($IsSiteVisitor)
  {
    $GroupId = 7
  }
  else 
  {
    $GroupId = $NULL
  }
  #start creating the user
  Write-Verbose -Message 'Start creating the SharePoint user.'
  try
  {
    $UserCreationInfo = New-Object -TypeName Microsoft.SharePoint.Client.UserCreationInformation
    $UserCreationInfo.Title = $Title
    $UserCreationInfo.Email = $Email
    $UserCreationInfo.LoginName = $LoginName
    IF ($GroupId -ne $NULL)
    {
      $Group = $Context.web.SiteGroups.GetById($GroupId)
      $Context.Load($Group)
      $Context.ExecuteQuery()
      $User = $Group.Users.Add($UserCreationInfo)
    }
    else 
    {
      $User = $web.SiteUsers.Add($UserCreationInfo)
    }
    $Context.ExecuteQuery()
    return $true
  }
  catch 
  {
    Throw $_.Exception.InnerException
    return $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPGroup
{
  [OutputType('Microsoft.SharePoint.Client.GroupCollection')]
  [OutputType('Microsoft.SharePoint.Client.Group')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the title of the group')]
    [ValidateNotNullOrEmpty()][String]$GroupTitle = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the Id of the group')]
    [ValidateNotNullOrEmpty()][Int]$GroupId = $NULL
  )

  #Work out if $GroupTitle and $GroupId contains values
  #GroupTitle
  If ($PSBoundParameters.ContainsKey('GroupTitle'))
  {
    $GroupTitle = $PSBoundParameters.GroupTitle
  }
  else 
  {
    Remove-Variable -Name GroupTitle
  }
  #GroupId
  If ($PSBoundParameters.ContainsKey('GroupId'))
  {
    $GroupId = $PSBoundParameters.GroupId
  }
  else 
  {
    Remove-Variable -Name GroupId
  }

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint Group.'
  Try
  {
    $Groups = $web.SiteGroups
    $Context.Load($Groups)
    $Context.ExecuteQuery()
    If ($PSBoundParameters.ContainsKey('GroupTitle'))
    {
      Write-Verbose -Message "Group Title specified: '$GroupTitle'."
      $Group = $Groups.GetByName($GroupTitle)
      $Context.Load($Group)
      $Context.ExecuteQuery()
      $Context.Load($Group.Users)
      $Context.ExecuteQuery()
      $Result = $Group
    }
    elseif ($PSBoundParameters.ContainsKey('GroupId')) 
    {
      Write-Verbose -Message "Group Id specified: '$GroupId'."
      $Group = $Groups.GetById($GroupId)
      $Context.Load($Group)
      $Context.ExecuteQuery()
      $Context.Load($Group.Users)
      $Context.ExecuteQuery()
      $Result = $Group
    }
    else 
    {
      Write-Verbose -Message 'Get all groups'
      Foreach ($Group in $Groups)
      {
        $Context.Load($Group)
        $Context.ExecuteQuery()
        $Context.Load($Group.Users)
        $Context.ExecuteQuery()
      }
      $Result = $Groups
    }
  }
  Catch 
  {
    throw $_.exception.innerexception
    return
  }
  Return $Result
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPGroup
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the title of the group')]
    [ValidateNotNullOrEmpty()][String]$GroupTitle,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the description of the group')]
    [ValidateNotNullOrEmpty()][String]$GroupDescription = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the title of the group')]
    [ValidateNotNullOrEmpty()][String[]]$GroupRoles
  )
  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  #Validate each group roles before continuing
  if ($GroupRoles.count -gt 0)
  {
    Write-Verbose -Message 'Validating the group roles specified.'
    $bValidRoles = $true
    Foreach ($role in $GroupRoles)
    {
      If ([Microsoft.SharePoint.Client.RoleType]::$role)
      {
        if ([Microsoft.SharePoint.Client.RoleType]::$role -eq 'None')
        {
          Write-Error -Message "There is no need to specify the role 'None'. This role will be automatically assigned if no other roles have been specified."
          $bValidRoles = $false
        }
        else 
        {
          Write-Verbose -Message "The group role '$role' is a valid role."
        }
      }
      else 
      {
        Write-Error -Message "The group role specified: '$role' is not a valid role."
        $bValidRoles = $false
      }
    }
    If ($bValidRoles -eq $false)
    {
      Throw 'Group roles validation failed. unable to continue.'
      Return $false
    }
  }

  #start creating the group
  Write-Verbose -Message 'Start creating the SharePoint group.'
  try
  {
    $GroupCreationInfo = New-Object -TypeName Microsoft.SharePoint.Client.GroupCreationInformation
    $GroupCreationInfo.Title = $GroupTitle
    $GroupCreationInfo.Description = $GroupDescription
    $Group = $web.SiteGroups.Add($GroupCreationInfo)
    $colRoleDefBinding = New-Object -TypeName Microsoft.SharePoint.Client.RoleDefinitionBindingCollection -ArgumentList ($Context)
    $Context.Load($Group)
    $Context.ExecuteQuery()
    If ($GroupRoles.count -gt 0)
    {
      Foreach ($role in $GroupRoles)
      {
        Write-Verbose -Message "Assigning role '$role' to group '$GroupTitle'."
        $RoleDefinition = $web.RoleDefinitions.GetByType([Microsoft.SharePoint.Client.RoleType]::$role)
        $colRoleDefBinding.add($RoleDefinition)
        $Context.Load($RoleDefinition)
      }
      $AssignRole = $web.RoleAssignments.Add($Group, $colRoleDefBinding)
      $Context.ExecuteQuery()
    }
    else 
    {
      Write-Verbose -Message "No roles are assigned to the group '$GroupTitle'."
    }
    $Context.ExecuteQuery()
    return $true
  }
  catch 
  {
    Throw $_.Exception.InnerException
    return $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPGroupMember
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the Group')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the Group')]
    [ValidateNotNullOrEmpty()][String]$GroupTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the Group')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the Group')]
    [ValidateNotNullOrEmpty()][Int]$GroupId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify user name of the user to be added to the group')]
    [ValidateNotNullOrEmpty()][Alias('New')][String]$NewMemberUserName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the group
  Write-Verbose -Message 'Getting the SharePoint group.'
  Try
  {
    $Groups = $web.SiteGroups
    $Context.Load($Groups)
    $Context.ExecuteQuery()
    If ($PSBoundParameters.ContainsKey('GroupTitle'))
    {
      Write-Verbose -Message "Group Title specified: '$GroupTitle'."
      $Group = $Groups.GetByName($GroupTitle)
    }
    elseif ($PSBoundParameters.ContainsKey('GroupId')) 
    {
      Write-Verbose -Message "Group Id specified: '$GroupId'."
      $Group = $Groups.GetById($GroupId)
    }
    else 
    {
      Throw 'Unable to retrieve the group.'
      Return $false
    }
    $Context.Load($Group)
    $Context.ExecuteQuery()
    $Context.Load($Group.Users)
    $Context.ExecuteQuery()
  }
  Catch 
  {
    throw $_.exception.innerexception
    return $false
  }
  Write-Verbose -Message "adding the new user '$NewMemberUserName' to the group"
  try
  {
    $NewMember = $web.EnsureUser($NewMemberUserName)
    $Context.Load($NewMember)
    $AddUser = $Group.Users.AddUser($NewMember)
    $Context.Load($AddUser)
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPGroupMember
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the Group')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the Group')]
    [ValidateNotNullOrEmpty()][String]$GroupTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the Group')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the Group')]
    [ValidateNotNullOrEmpty()][Int]$GroupId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify user name of the user to be removed to the group')]
    [ValidateNotNullOrEmpty()][Alias('Remove')][String]$RemoveUserName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the group
  Write-Verbose -Message 'Getting the SharePoint group.'
  Try
  {
    $Groups = $web.SiteGroups
    $Context.Load($Groups)
    $Context.ExecuteQuery()
    If ($PSBoundParameters.ContainsKey('GroupTitle'))
    {
      Write-Verbose -Message "Group Title specified: '$GroupTitle'."
      $Group = $Groups.GetByName($GroupTitle)
    }
    elseif ($PSBoundParameters.ContainsKey('GroupId')) 
    {
      Write-Verbose -Message "Group Id specified: '$GroupId'."
      $Group = $Groups.GetById($GroupId)
    }
    else 
    {
      Throw 'Unable to retrieve the group.'
      Return $false
    }
    $Context.Load($Group)
    $Context.ExecuteQuery()
    $Context.Load($Group.Users)
    $Context.ExecuteQuery()
  }
  Catch 
  {
    throw $_.exception.innerexception
    return $false
  }
  Write-Verbose -Message "removing the user '$RemoveUserName' from the group"
  try
  {
    $ToBeRemovedMember = $web.EnsureUser($RemoveUserName)
    $Context.Load($ToBeRemovedMember)
    $Group.Users.Remove($ToBeRemovedMember)
    $Context.ExecuteQuery()
    $true
  }
  catch 
  {
    throw $_.Exception.InnerException
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Clear-SPSiteRecycleBin
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
  $site = $Context.Site
    
  #Get the group
  Write-Verbose -Message 'Clearing recycle bin.'
  Try
  {
    $web.RecycleBin.DeleteAll()
    $site.RecycleBin.DeleteAll()
    $Context.ExecuteQuery()
    $true
  }
  Catch 
  {
    throw $_.exception.innerexception
    $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPSiteTemplate
{
  [OutputType('Microsoft.SharePoint.Client.WebTemplateCollection')]
  [OutputType('Microsoft.SharePoint.Client.WebTemplate')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualParameters',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the title of the site template')][String]$TemplateTitle = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the name of the site template')][String]$TemplateName = $NULL,

    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the locale for the site template to be retrieved')]
    [ValidateScript({
          if ([System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag($_))
          {
            $true
          }
          else
          {
            $false
          }
    })]
    [Alias('Locale')][String]$TemplateLocale = 'en-us'
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the all avaialble templates
  Write-Verbose -Message "Getting all avaialbe SharePoint site template for locale '$TemplateLocale'."
  Try
  {
    $LCID = [System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag($TemplateLocale).LCID
    $AvailableTemplates = $web.GetAvailableWebTemplates($LCID,$true)
    $Context.Load($AvailableTemplates)
    $Context.ExecuteQuery()
  }
  catch 
  {
    Throw $_.Exception.InnerException
    Return
  }
  If ($AvailableTemplates.count -eq 0)
  {
    Throw "No site templates found for locale '$TemplateLocale'."
    Return
  }
  #Get the specific template if the template title or name is specified
  If ($PSBoundParameters.ContainsKey('TemplateTitle'))
  {
    Write-Verbose -Message "Getting site template with title '$TemplateTitle'."
    $SpecifiedTemplate = $AvailableTemplates | Where-Object -FilterScript {
      $_.Title -ieq $TemplateTitle
    }
    If (!$SpecifiedTemplate)
    {
      Throw "The specified site template with title '$TemplateTitle' for locale '$TemplateLocale' does not exist!"
      Return
    }
    else 
    {
      Return $SpecifiedTemplate
    }
  }
  elseif ($PSBoundParameters.ContainsKey('TemplateName')) 
  {
    Write-Verbose -Message "Getting site template with name '$TemplateName'."
    $SpecifiedTemplate = $AvailableTemplates | Where-Object -FilterScript {
      $_.Name -ieq $TemplateName
    }
    If (!$SpecifiedTemplate)
    {
      Throw "The specified site template with name '$TemplateName' for locale '$TemplateLocale' does not exist!"
      Return
    }
    else 
    {
      Return $SpecifiedTemplate
    }
  }
  else 
  {
    return $AvailableTemplates
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function New-SPSubSite
{
  [OutputType('Microsoft.SharePoint.Client.Site')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Root Site URL')][Alias('RootSiteURL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the title of the new site')][ValidateNotNullOrEmpty()][String]$NewSiteTitle,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Url leaf name of the new site')][ValidateNotNullOrEmpty()][String]$NewSiteUrlLeaf,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the description of the new site')][String]$NewSiteDescription = $NULL,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the name of the site template to be used for the new site')][ValidateNotNullOrEmpty()][Alias('TemplateName')][String]$NewSiteTemplateName,
    [Parameter(Mandatory = $false,HelpMessage = 'Please specify the locale for the site template to be used for the new site')]
    [ValidateScript({
          if ([System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag($_))
          {
            $true
          }
          else
          {
            $false
          }
    })]
    [Alias('TemplateLocale','Locale')][String]$NewSiteTemplateLocale = 'en-us'
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  #Validate specified site template
  Try
  {
    $LCID = [System.Globalization.CultureInfo]::GetCultureInfoByIetfLanguageTag($NewSiteTemplateLocale).LCID
    $AvailableTemplates = $web.GetAvailableWebTemplates($LCID,$true)
    $Context.Load($AvailableTemplates)
    $Context.ExecuteQuery()
    $SpecifiedTemplate = $AvailableTemplates | Where-Object -FilterScript {
      $_.Name -ieq $NewSiteTemplateName
    }
    If (!$SpecifiedTemplate)
    {
      Throw "The specified site template '$NewSiteTemplateName' does not exist!"
    }
  }
  catch 
  {
    Throw $_.Exception.InnerException
    Return
  }

  #Create new site
  Write-Verbose -Message "Creating 'Microsoft.SharePoint.Client.WebCreationInformation' object for new site `'$NewSiteTitle`'."
  $CreationInfo = New-Object -TypeName Microsoft.SharePoint.Client.WebCreationInformation
  $CreationInfo.Url = $NewSiteUrlLeaf
  $CreationInfo.Title = $NewSiteTitle
  if ($NewSiteDescription.length -gt 0)
  {
    Write-Verbose -Message "New site Description: '$NewSiteDescription'."
    $CreationInfo.Description = $NewSiteDescription
  }
  $CreationInfo.WebTemplate = $NewSiteTemplateName
  try
  {
    #Create the site
    Write-Verbose -Message "Creating list '$ListTitle'."
    $NewSite = $web.webs.Add($CreationInfo)
    $Context.Load($NewSite)
    $Context.ExecuteQuery()
        
    #Retrieve the site
    Write-Verbose -Message 'Retrieving the new site.'
    $Context.Load($NewSite)
    $Context.ExecuteQuery()
    Write-Verbose -Message 'Returning the new site.'
    $NewSite
  }
  catch 
  {
    throw $_.Exception.InnerException
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Get-SPSubSite
{
  [OutputType('System.Array')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Root Site URL')][Alias('RootSiteURL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  $Sites = $web.Webs
  $Context.Load($web)
  $Context.Load($Sites)
  $Context.ExecuteQuery()
  $arrSites = @()
  Foreach ($site in $Sites)
  {
    $Context.Load($site)
    $Context.ExecuteQuery()
    $arrSites += $site
  }
  Write-Verbose -Message "Total number of subsites: $($arrSites.count)."
  ,$arrSites
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPSubSite
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAConnection',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')][Object]$SPConnection,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Root Site URL')][Alias('RootSiteURL')][String]$SiteUrl,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')][Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,
    [Parameter(ParameterSetName = 'IndividualParameter',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')][Alias('IsSPO')][boolean]$IsSharePointOnlineSite,
    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the Url of the sub site that is going to be deleted')][ValidateNotNullOrEmpty()][String]$SubSiteURL
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web

  #Get all subsites
  Write-Verbose -Message 'Get all subsites'
  $Subsites = $web.Webs
  $Subsites = $web.Webs
  $Context.Load($web)
  $Context.Load($Subsites)
  $Context.ExecuteQuery()
  Foreach ($site in $Subsites)
  {
    $Context.Load($site)
    $Context.ExecuteQuery()
    If ($site.url -ieq $SubSiteURL)
    {
      Write-Verbose -Message "Subsite found. Site Title: '$($site.Title)'."
      $SiteToBeDeleted = $site
    }
  }
  If ($SiteToBeDeleted)
  {
    Write-Verbose -Message "Deleting site '$($SiteToBeDeleted.Url)'."
    Try
    {
      $SiteToBeDeleted.DeleteObject()
      $Context.ExecuteQuery()
      $Result = $true
    }
    Catch 
    {
      Throw $_.Exception.InnerException
      $Result = $false
    }
  }
  else 
  {
    Throw "Unable to find sub site with URL '$SubSiteURL'."
    $Result = $false
  }
  $Result
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Add-SPListFieldToDefaultView
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the internal name of the list field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnInternalName')][String]$FieldInternalName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return $false
  }
  $Context.Load($List)
  $Context.Load($List.Fields)
  $Context.ExecuteQuery()
    
  #Add the field to the default view
  Write-Verbose -Message 'Adding the list field to the default view'
  Try 
  {
    $Context.Load($List.DefaultView)
    $List.DefaultView.ViewFields.Add($FieldInternalName)
    $List.DefaultView.Update()
    $Context.ExecuteQuery()
    Return $true
  }
  Catch 
  {
    Throw $_.Exception.InnerException
    Return $false
  }
}

# .EXTERNALHELP SharePointSDK.psm1-Help.xml
Function Remove-SPListFieldFromDefaultView
{
  [OutputType('System.Boolean')]
  [CmdletBinding()]
  Param(
    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the SMA / Azure Autoamtion connection object')]
    [Object]$SPConnection,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the SharePoint Site URL')]
    [String]$SiteUrl,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the user credential to connect to the SharePoint or SharePoint Online site')]
    [Alias('cred')]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify if the site is a SharePoint Online site')]
    [Alias('IsSPO')][boolean]$IsSharePointOnlineSite,

    [Parameter(ParameterSetName = 'SMAByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [Parameter(ParameterSetName = 'IndividualByTitle',Mandatory = $true,HelpMessage = 'Please specify the title of the list')]
    [ValidateNotNullOrEmpty()][String]$ListTitle,

    [Parameter(ParameterSetName = 'SMAById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [Parameter(ParameterSetName = 'IndividualById',Mandatory = $true,HelpMessage = 'Please specify the Id of the list')]
    [ValidateNotNullOrEmpty()][Guid]$ListId,

    [Parameter(Mandatory = $true,HelpMessage = 'Please specify the internal name of the list field')]
    [ValidateNotNullOrEmpty()][Alias('ColumnInternalName')][String]$FieldInternalName
  )

  If($SPConnection)
  {
    $SPcredential = New-SPCredential -SPConnection $SPConnection
    $SiteUrl = $SPConnection.SharePointSiteURL
  } else 
  {
    $SPcredential = New-SPCredential -Credential $Credential -IsSharePointOnlineSite $IsSharePointOnlineSite
  }

  #Bind to site collection
  Write-Verbose -Message 'Bind to site collection'
  $Context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext -ArgumentList ($SiteUrl)
  $Context.Credentials = $SPcredential
  $web = $Context.Web
    
  #Get the list
  Write-Verbose -Message 'Getting the SharePoint list.'
  If ($PSBoundParameters.ContainsKey('ListTitle'))
  {
    Write-Verbose -Message "List Title specified: '$ListTitle'."
    $List = $web.Lists.GetByTitle($ListTitle)
  }
  elseif ($PSBoundParameters.ContainsKey('ListId')) 
  {
    Write-Verbose -Message "List Id specified: '$($ListId.ToString())'."
    $List = $web.Lists.GetById($ListId)
  }
  else 
  {
    Throw 'Unable to retrieve the list.'
    Return $false
  }
  $Context.Load($List)
  $Context.Load($List.Fields)
  $Context.ExecuteQuery()
    
  #Remove the field from the default view
  Write-Verbose -Message 'Removing the list field From the default view'
  Try 
  {
    $Context.Load($List.DefaultView)
    $List.DefaultView.ViewFields.Remove($FieldInternalName)
    $List.DefaultView.Update()
    $Context.ExecuteQuery()
    Return $true
  }
  Catch 
  {
    Throw $_.Exception.InnerException
    Return $false
  }
}
