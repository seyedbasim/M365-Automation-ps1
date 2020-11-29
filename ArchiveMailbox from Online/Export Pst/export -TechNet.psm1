function Enter-ExchangeOnlineSession {
  <#
      .Synopsis
      Sets up a user session with ExchangeOnline, typically to fetch mail.
      .DESCRIPTION
      Sets up a user session with ExchangeOnline, typically to fetch mail.
      .EXAMPLE
      $CredMail = Get-Credential -UserName user@domain.com
      $Service = Enter-ExchangeOnlineSession -Credential $CredMail -MailDomain domain.com
      .INPUTS
      Inputs to this cmdlet (if any)
      .OUTPUTS
      Output from this cmdlet (if any)
      .NOTES
      General notes
      .COMPONENT
      The component this cmdlet belongs to
      .ROLE
      The role this cmdlet belongs to
      .FUNCTIONALITY
      The functionality that best describes this cmdlet
  #>
  Param(
    [Parameter(ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$EWSAssemblyPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    # Parameter Credential
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=1)]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    # Parameter Maildomain
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=2)]
    [string]$MailDomain,

    # Parameter Maildomain
    [Parameter(Mandatory=$false, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=3)]
    [string]$AutoDiscoverCallbackUrl = 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml'

  )

  # Set the path to your copy of EWS Managed API and load the assembly.
  [void][Reflection.Assembly]::LoadFile($EWSAssemblyPath) 

  # Establish the session and return the service connection object.
  $service =  New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
  $Service.Credentials = New-Object System.Net.NetworkCredential($Credential.UserName, $Credential.GetNetworkCredential().Password, $MailDomain)

  $TestUrlCallback = {
    param ([string] $url)
    if ($url -eq $AutoDiscoverCallbackUrl) {$true} else {$false}
  }
  $service.AutodiscoverUrl($Credential.UserName, $TestUrlCallback)

  return $Service
}

function Get-ExchangeOnlineMailContent {
  <#
      .Synopsis
      Fetches content from an Exchange Online user session, typically mail.
      .DESCRIPTION
      Fetches content from an Exchange Online user session, typically mail.
      .EXAMPLE
      $Service = Enter-ExchangeOnlineSession -Credential $CredMail -MailDomain example.com
      Get-ExchangeOnlineMailContent -ServiceObject $Service -PageSize 5 -Offset 0 -PageIndexLimit 15 -WellKnownFolderName Inbox
      .INPUTS
      Inputs to this cmdlet (if any)
      .OUTPUTS
      Output from this cmdlet (if any)
      .NOTES
      General notes
      .COMPONENT
      The component this cmdlet belongs to
      .ROLE
      The role this cmdlet belongs to
      .FUNCTIONALITY
      The functionality that best describes this cmdlet
  #>

  Param(
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=0)]
    $ServiceObject,

    # Define paging as described in https://msdn.microsoft.com/en-us/library/office/dn592093(v=exchg.150).aspx
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=1)]
    [int]$PageSize,

    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=2)]
    [int]$Offset,

    #Translates into multiples of $PageSize
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=3)]
    [int]$PageIndexLimit,

    # WellKnownFolderNames doc https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx
    [Parameter(Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=4)]
    [ValidateSet('Calendar',
                'Contacts',
                'DeletedItems',
                'Drafts',
                'Inbox',
                'Journal',
                'Notes',
                'Outbox',
                'SentItems',
                'Tasks',
                'MsgFolderRoot',
                'PublicFoldersRoot',
                'Root',
                'JunkEmail',
                'SearchFolders',
                'VoiceMail',
                'RecoverableItemsRoot',
                'RecoverableItemsDeletions',
                'RecoverableItemsVersions',
                'RecoverableItemsPurges',
                'ArchiveRoot',
                'ArchiveMsgFolderRoot',
                'ArchiveDeletedItems',
                'ArchiveRecoverableItemsRoot',
                'ArchiveRecoverableItemsDeletions',
                'ArchiveRecoverableItemsVersions',
                'ArchiveRecoverableItemsPurges',
                'SyncIssues',
                'Conflicts',
                'LocalFailures',
                'ServerFailures',
                'RecipientCache',
                'QuickContacts',
                'ConversationHistory',  
                'ToDoSearch')]
    [string]$WellKnownFolderName,

    [Parameter(Mandatory=$false, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=5)]
    [switch]$ParseOriginalRecipient,

    [Parameter(Mandatory=$false, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=6)]
    [switch]$OriginalRecipientAddressOnly,

    [Parameter(Mandatory=$false, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=7)]
    [datetime]$MailFromDate,

    [Parameter(Mandatory=$false, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true, 
        ValueFromRemainingArguments=$false, 
        Position=8)]
    [switch]$ConsoleOutput
  )

  # Create Property Set to include body and header, set body to text.
  $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
  $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

  if ($ParseOriginalRecipient)
  {
    $PR_TRANSPORT_MESSAGE_HEADERS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x007D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
    $PropertySet.add($PR_TRANSPORT_MESSAGE_HEADERS)
  }

  $PageIndex = 0

  # Page through folder.
  do 
  { 
    # Limit the view to $pagesize number of email starting at $Offset.
    $PageView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize,$PageIndex,$Offset)

    # Get folder data.
    $FindResults = $ServiceObject.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName,$PageView) 
    foreach ($MailItem in $FindResults.Items)
    {
      # Load extended properties.
      $MailItem.Load($propertySet)

      if ($ParseOriginalRecipient)
      {
        # Extended properties are one string, split by linebreak then find the line beginning with 'To:', containing original recipient address before exchange aliasing replaces it.
        $OriginalRecipientStringRaw = ($MailItem.ExtendedProperties.value.split([Environment]::NewLine) | Where-Object {$_ -match '^To:'}).trimstart('To:').trim()

        $MailItem | Add-Member -NotePropertyName ExtendedHeader -NotePropertyValue $OriginalRecipientStringRaw

        if ($OriginalRecipientAddressOnly)
        {
          # Removes everything but the address 'my.name@example.com' when string has form of e.g. '"My Name" <my.name@example.com>'
          $MailItem | Add-Member -NotePropertyName OriginalRecipient -NotePropertyValue ($OriginalRecipientStringRaw | ForEach-Object {($_.ToString() -creplace '^[^<]*<', '').trimend('>')})
        }
	else
        {
          $MailItem | Add-Member -NotePropertyName OriginalRecipient -NotePropertyValue $OriginalRecipientStringRaw
        }
        if ($ConsoleOutput) {write-host -ForegroundColor Cyan "$($MailItem.DateTimeReceived) | $($MailItem.OriginalRecipient) | $($MailItem.Sender)"}
      }

      # Output result.
      $MailItem | Select-Object -Property *
    } 

    # Increment $index to next page.
    $PageIndex += $PageSize
  } while (($FindResults.MoreAvailable) `
            -and ($PageIndex -lt $PageIndexLimit) `
            -and ($MailItem.DateTimeReceived -gt $MailFromDate)) # Do/While there are more emails to retrieve and pagelimit is not exceeded and datetimereceived is later than date.
}