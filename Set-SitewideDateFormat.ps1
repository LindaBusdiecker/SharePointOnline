<#
.SYNOPSIS
    For a particular site, this will set the date format of all date fields to the 
    particular format specified.
.DESCRIPTION
    Given the URL to a site, this traverses all lists and libraries and sets 
    the format of all date fields to the format specified, using the enumeration
    specified in https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/jj174848(v=office.15).

.PARAMETER SiteURL
    Example: "https://bogus.sharepoint.com/teams/sitename"

.PARAMETER FriendlyDate

    For field with TypeAsString of DateTime, the property FriendlyDisplayFormat is set.

    Name: "Unspecified"
    Description: Undefined. The default rendering will be used. Value = 0. 

    Name: "Disabled"
    Description: The standard absolute representation will be used. Value = 1. 

    Name: "Relative"
    Description: The standard friendly relative representation will be used (for example, "today at 3:00 PM"). Value = 2. 

.EXAMPLE 

    *** Disabling friendly date for existing date fields, but with a federated sign-in realm issue

    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus.sharepoint.com/teams/sitename" -FriendlyDate "Disabled" -verbose

    cmdlet Get-Credential at command pipeline position 1
    Supply values for the following parameters:
    User: linda.busdiecker@hennnepin.us
    Password for user linda.busdiecker@hennnepin.us: ***********
    Connect-PnPOnline : Identity Client Runtime Library (IDCRL) could not look up the realm information for a federated sign-in.
    At H:\PS\SP\Set-SitewideDateFormat.ps1:63 char:13
    +     $Site = Connect-PnPOnline -Url $SiteUrl -Credential $Credentials  ...
    +             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : NotSpecified: (:) [Connect-PnPOnline], IdcrlException
        + FullyQualifiedErrorId : Microsoft.SharePoint.Client.IdcrlException,SharePointPnP.PowerShell.Commands.Base.ConnectOnline

    VERBOSE: List [appdata], Field [Modified] initial value: [Relative]
    VERBOSE: List [appdata], Field [Modified] updated value: [Disabled]

    VERBOSE: List [appdata], Field [Created] initial value: [Relative]
    VERBOSE: List [appdata], Field [Created] updated value: [Disabled]

    ...
.EXAMPLE

    *** Successful disabling friendly date for existing date fields

    PS H:\PS\SP> Set-SitewideDateFormat -SiteURL "https://bogus.sharepoint.com/teams/sitename" -FriendlyDate "Disabled" -verbose
    cmdlet Get-Credential at command pipeline position 1
    Supply values for the following parameters:
    User: linda.busdiecker@hennepin.us
    Password for user linda.busdiecker@hennepin.us: ***********
    VERBOSE: PnP PowerShell Cmdlets (3.12.1908.1): Connected to https://hennepin.sharepoint.com/teams/fs-sumnerlibrary
    VERBOSE: List [appdata], Field [Modified] initial value: [Disabled]
    VERBOSE: List [appdata], Field [Modified] updated value: [Disabled]

    VERBOSE: List [appdata], Field [Created] initial value: [Disabled]
    VERBOSE: List [appdata], Field [Created] updated value: [Disabled]

    ...
.EXAMPLE
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus.sharepoint.com/teams/sitename" -FriendlyDate "Relative" -verbose
.EXAMPLE
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus.sharepoint.com/teams/sitename" -FriendlyDate "Unspecified" -verbose
.NOTES
    Author: Linda Busdiecker
    Date: 08/21/2019
#>
#Import-Module SharePointPnPPowerShellOnline -Scope Local -ErrorAction Stop
#Requires -Module SharePointPnPPowerShellOnline 
function Set-SitewideDateFormat {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String] $SiteURL,
        [Parameter(Mandatory = $true)] [ValidateSet("Undefined", "Disabled", "Relative")] [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType] $FriendlyDate
    )

    # $Modules = Get-InstalledModule
    # ForEach ($Module in $Modules) {
    #     If ($Module.Name -eq "SharePointPnPPowerShellOnline") {
    #         $Continue = $true
    #     }
    #     ElseIf ($Module.Name.StartsWith("SharePointPnPPowerShell")) {
    #         Try {
    #             Remove-Module $Module.Name -Scope Local -Verbose -Force -ErrorAction Stop
    #             $Continue = $true 
    #         }
    #         Catch {
    #             Write-Error "[$($Module.Name)] is installed and could cause issues"
    #         }
    #     }
    # }

    $Credentials = Get-Credential
    
    $Site = Connect-PnPOnline -Url $SiteUrl -Credential $Credentials -ReturnConnection -ErrorAction Stop

    $Context = $Site.Context
    $Context.ExecuteQuery()

    $Lists = $Context.Web.Lists
    $Context.Load($Lists)
    $Context.ExecuteQuery()

    ForEach ($List in $Lists) {

        $Fields = $List.Fields
        $Site.Context.Load($Fields)
        $Context.ExecuteQuery()

        ForEach ($Field in $Fields) {
            If ($Field.TypeAsString -eq "DateTime") {
                Write-Verbose "List [$($List.Title)], Field [$($Field.Title)] initial value: [$($Field.FriendlyDisplayFormat)]"
                $Field.FriendlyDisplayFormat = $FriendlyDate
                $Field.Update()
                Write-Verbose "List [$($List.Title)], Field [$($Field.Title)] updated value: [$($Field.FriendlyDisplayFormat)]`n"
            }
        }
    }
}