<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#>

#requires -Version 2

#Import Localized Data
Import-LocalizedData -BindingVariable Messages
#Load .NET Assembly for Windows PowerShell V2
Add-Type -AssemblyName System.Core

$webSvcInstallDirRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -PSProperty "Install Directory" -ErrorAction:SilentlyContinue
if ($webSvcInstallDirRegKey -ne $null) {
	$moduleFilePath = $webSvcInstallDirRegKey.'Install Directory' + 'Microsoft.Exchange.WebServices.dll'
	Import-Module $moduleFilePath
} else {
	$errorMsg = $Messages.InstallExWebSvcModule
	throw $errorMsg
}

Function New-OSCPSCustomErrorRecord
{
	#This function is used to create a PowerShell ErrorRecord
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true,Position=1)][String]$ExceptionString,
		[Parameter(Mandatory=$true,Position=2)][String]$ErrorID,
		[Parameter(Mandatory=$true,Position=3)][System.Management.Automation.ErrorCategory]$ErrorCategory,
		[Parameter(Mandatory=$true,Position=4)][PSObject]$TargetObject
	)
	Process
	{
		$exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
		$customError = New-Object System.Management.Automation.ErrorRecord($exception,$ErrorID,$ErrorCategory,$TargetObject)
		return $customError
	}
}

Function Connect-OSCEXOWebService
{
	#.EXTERNALHELP Connect-OSCEXOWebService-Help.xml

	[cmdletbinding()]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$true,Position=1)]
		[System.Management.Automation.PSCredential]$Credential,
		[Parameter(Mandatory=$false,Position=2)]
		[Microsoft.Exchange.WebServices.Data.ExchangeVersion]$ExchangeVersion="Exchange2010_SP2",
		[Parameter(Mandatory=$false,Position=3)]
		[string]$TimeZoneStandardName,
		[Parameter(Mandatory=$false)]
		[switch]$Force
	)
	Process
	{
		#Get specific time zone info
		if (-not [System.String]::IsNullOrEmpty($TimeZoneStandardName)) {
			Try
			{
				$tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneStandardName)
			}
			Catch
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
		} else {
			$tzInfo = $null
		}

		#Create the callback to validate the redirection URL.
		$validateRedirectionUrlCallback = {
			param ([string]$Url)
			if ($Url -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml") {
				return $true
			} else {
				return $false
			}
		}	

		#Try to get exchange service object from global scope
		$existingExSvcVar = (Get-Variable -Name exService -Scope Global -ErrorAction:SilentlyContinue) -ne $null

		#Establish the connection to Exchange Web Service
		if ((-not $existingExSvcVar) -or $Force) {
			$verboseMsg = $Messages.EstablishConnection
			$PSCmdlet.WriteVerbose($verboseMsg)
			if ($tzInfo -ne $null) {
				$exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
							[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion,$tzInfo)			
			} else {
				$exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
							[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
			}

			#Set network credential
			$userName = $Credential.UserName
			$exService.Credentials = $Credential.GetNetworkCredential()
			Try
			{
				#Set the URL by using Autodiscover
				$exService.AutodiscoverUrl($userName,$validateRedirectionUrlCallback)
				$verboseMsg = $Messages.SaveExWebSvcVariable
				$PSCmdlet.WriteVerbose($verboseMsg)
				Set-Variable -Name exService -Value $exService -Scope Global -Force
			}
			Catch [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException]
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
			Catch
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
		} else {
			$verboseMsg = $Messages.FindExWebSvcVariable
			$verboseMsg = $verboseMsg -f $exService.Credentials.Credentials.UserName
			$PSCmdlet.WriteVerbose($verboseMsg)            
		}
	}
}

Function New-OSCEXOAppointment
{
    #.EXTERNALHELP New-OSCEXOAppointment-Help.xml

    [cmdletbinding(DefaultParameterSetName="NoRecurrence")]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true)]
		[string]$Identity,

		[Parameter(Mandatory=$true,Position=2)]
		[string]$Subject,

		[Parameter(Mandatory=$false,Position=3)]
		[string]$Location,

		[Parameter(Mandatory=$false,Position=4)]
		[string]$Body,

		[Parameter(Mandatory=$false,Position=5)]
		[datetime]$StartDate,

		[Parameter(Mandatory=$false,Position=6)]
        [ValidateScript({$_ -gt $StartDate})]
		[datetime]$EndDate,

		[Parameter(Mandatory=$false)]
		[switch]$AllDayEvent,

		[Parameter(Mandatory=$false)]
		[switch]$UseImpersonation,

		[Parameter(Mandatory=$false, ParameterSetName="DailyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="WeeklyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="MonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="YearlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeMonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeYearlyRecurrence")]
		[datetime]$RecurrenceRangeStart,

		[Parameter(Mandatory=$false, ParameterSetName="DailyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="WeeklyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="MonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="YearlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeMonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeYearlyRecurrence")]
		[int]$RecurrenceRangeEndAfter,

		[Parameter(Mandatory=$false, ParameterSetName="DailyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="WeeklyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="MonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="YearlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeMonthlyRecurrence")]
        [Parameter(Mandatory=$false, ParameterSetName="RelativeYearlyRecurrence")]
		[datetime]$RecurrenceRangeEndBy,

		[Parameter(Mandatory=$true, ParameterSetName="DailyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="WeeklyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="MonthlyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="RelativeMonthlyRecurrence")]
		[int]$Interval,
        
        [Parameter(Mandatory=$true, ParameterSetName="WeeklyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="RelativeMonthlyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="RelativeYearlyRecurrence")]
		[Microsoft.Exchange.WebServices.Data.DayOfTheWeek]$DayOfTheWeek,

        [Parameter(Mandatory=$true, ParameterSetName="YearlyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="RelativeYearlyRecurrence")]
		[Microsoft.Exchange.WebServices.Data.Month]$Month,

        [Parameter(Mandatory=$true, ParameterSetName="MonthlyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="YearlyRecurrence")]
		[int]$DayOfMonth,

        [Parameter(Mandatory=$true, ParameterSetName="RelativeMonthlyRecurrence")]
        [Parameter(Mandatory=$true, ParameterSetName="RelativeYearlyRecurrence")]
		[Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex]$DayOfTheWeekIndex
	)
    Begin
    {        
        if ($exService -eq $null) {
			$errorMsg = $Messages.RequireConnection
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString $errorMsg `
			-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
			$PSCmdlet.ThrowTerminatingError($customError)
        }

        #Initilize StartDate and EndDate
        $now = Get-Date

        if ($StartDate -eq $null) {
            if (-not $AllDayEvent) {
                $nextHour = $now.AddMinutes(30).Hour
                if ($now.Hour -eq $nextHour) {
	                $StartDate = New-Object System.DateTime($now.Year,$now.Month,$now.Day,$nextHour,30,0)
                } else {
	                $StartDate = New-Object System.DateTime($now.Year,$now.Month,$now.Day,$nextHour,0,0)
                }
            } else {
                $StartDate = $now.Date
            }
        }

        if ($EndDate -eq $null) {
            $EndDate = $StartDate.AddMinutes(30)
        }

        #You cannot use RecurrenceRangeEndAfter and RecurrenceRangeEndBy at the same time
        if (($RecurrenceRangeEndAfter -ne 0) -and ($RecurrenceRangeEndBy -ne $null)) {
			$errorMsg = $Messages.InvalidRecurrenceRangeEnd
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString $errorMsg `
			-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
			$PSCmdlet.ThrowTerminatingError($customError)
        }
    }
    Process
    {
        $verboseMsg = $Messages.ResolveIdentity
        $verboseMsg = $verboseMsg -f $Identity
        $PSCmdlet.WriteVerbose($verboseMsg)

        #Resolve Identity parameter to get user's SMTP address
        $nameResolutionCollection = $exService.ResolveName($Identity,`
                                    [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$false)
        if ($nameResolutionCollection.Count -ne 1) {
			$errorMsg = $Messages.InvalidIdentity
			$errorMsg = $errorMsg -f $Identity
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString $errorMsg `
			-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
            $PSCmdlet.ThrowTerminatingError($customError)
        } else {
            $userSMTPAddress = $nameResolutionCollection[0].Mailbox.Address
        }

        $verboseMsg = $Messages.CreateAppointment
        $verboseMsg = $verboseMsg -f $Subject
        $PSCmdlet.WriteVerbose($verboseMsg)

        #Create a new appiontment
        $newAppointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($exService)
        $newAppointment.Subject = $Subject
        $newAppointment.Location = $Location
        $newAppointment.Body = $Body
        $newAppointment.Start = $StartDate
        $newAppointment.End = $EndDate

        #Set default values for reminder
        if (-not $AllDayEvent) {
            $newAppointment.ReminderMinutesBeforeStart = 15
            $newAppointment.ReminderDueBy = $StartDate.AddMinutes(15)
            $newAppointment.LegacyFreeBusyStatus = "Busy"
        } else {
            $newAppointment.ReminderMinutesBeforeStart = 1080
            $newAppointment.ReminderDueBy = $StartDate
            $newAppointment.LegacyFreeBusyStatus = "Free"
            $newAppointment.IsAllDayEvent = $true
        }

        #Set default date and time for recurrence range start
        if ($RecurrenceRangeStart -eq $null) {
            $RecurrenceRangeStart = $StartDate
        }

        #Enable recurrence setting if applicable
        Try
        {
            Switch ($PSCmdlet.ParameterSetName) {
                "DailyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+DailyPattern(`
                                                 $RecurrenceRangeStart,$Interval)
                }
                "WeeklyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern(`
                                                 $RecurrenceRangeStart,$Interval,$DayOfTheWeek)
                }
                "MonthlyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+MonthlyPattern(`
                                                 $RecurrenceRangeStart,$Interval,$DayOfMonth)
                }
                "YearlyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+YearlyPattern(`
                                                 $RecurrenceRangeStart,$Month,$DayOfMonth)
                }
                "RelativeMonthlyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+RelativeMonthlyPattern(`
                                                 $RecurrenceRangeStart,$Interval,$DayOfTheWeek,$DayOfTheWeekIndex)
                }
                "RelativeYearlyRecurrence" {
                    $newAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+RelativeYearlyPattern(`
                                                 $RecurrenceRangeStart,$Month,$DayOfTheWeek,$DayOfTheWeekIndex)
                }
            }
        }
        Catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        #Set default date and time for recurrence range start
        Try
        {
            if ($PSCmdlet.ParameterSetName -ne "NoRecurrence") {
                if ($RecurrenceRangeEndAfter -ne 0) {
                    $newAppointment.Recurrence.NumberOfOccurrences = $RecurrenceRangeEndAfter
                }

                if ($RecurrenceRangeEndBy -ne $null) {
                    $newAppointment.Recurrence.EndDate = $RecurrenceRangeEndBy
                }

                if (($RecurrenceRangeEndAfter -eq 0) -and ($RecurrenceRangeEndBy -eq $null)) {
                    $newAppointment.Recurrence.NeverEnds()
                }
            }
        }
        Catch
        {
            $PSCmdlet.ThrowTerminatingError($_)
        }

        $verboseMsg = $Messages.SaveAppointment
        $verboseMsg = $verboseMsg -f $Identity
        $PSCmdlet.WriteVerbose($verboseMsg)

        #Save appointment to the default Calendar folder
        Try
        {
            if ($UseImpersonation) {
                $impersonationUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(`
                                       [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,`
                                        ,$userSMTPAddress)
                $exService.ImpersonatedUserId = $impersonationUserId
                $newAppointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,`
                                     [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)

                #Clear ImpersonationUserId
                $exService.ImpersonatedUserId = $null
            } else {
                $mailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($userSMTPAddress)
                $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(`
                           [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,`
                           $mailbox)
                $newAppointment.Save($folderId, [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
            }
        }
        Catch
        {
            $PSCmdlet.WriteError($_)
        }
    }
    End {}
}

Export-ModuleMember -Function "Connect-OSCEXOWebService","New-OSCEXOAppointment"