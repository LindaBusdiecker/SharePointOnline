<#
.SYNOPSIS
    For a particular site, this will set all existing date columns to the particular format specified.

.DESCRIPTION
    Given the URL to a SharePoint site, this traverses all lists and libraries and sets the
    format of all existing date columns (aka fields) to the format specified. The web interface refers 
    to the formats as Standard and Friendly, which correspond to Disabled and Relative.

    The cmdlet will not run if SharePointPnPPowerShellOnline has not been imported.

    This requires -the module SharePointPnPPowerShellOnline. If you're not certain 
    it is available in your session, use Get-Module to find out or if it is installed,
    just import it, e.g.
        Import-Module SharePointPnPPowerShellOnline -Scope Local -ErrorAction Stop

.PARAMETER SiteURL
    SiteURL is the URL for the site for which the format of existing Date and Time
    columns should be standardized.

    Example: "https://bogus.sharepoint.com/teams/sitename"

.PARAMETER FriendlyDate

    FriendlyDate is the enumeration you want to standardize current Date and Time column formats.
    The enumeration is specified in
    https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/jj174848(v=office.15).

    This does not impact Date and Time columns entered after this is run.

    In the PowerShell script, you get the Date and Time by looking for columns whose
    TypeAsString is DateTime, the format gets changed by modifying the property FriendlyDisplayFormat.

    Name: "Unspecified"
    Description: Undefined. The default rendering will be used. Value = 0. 

    Name: "Disabled"
    Description: The standard absolute representation will be used. Value = 1. 

    Name: "Relative"
    Description: The standard friendly relative representation will be used (for example, "today at 3:00 PM"). Value = 2. 

.EXAMPLE 
    Credentials are passed in so they are not requested each time this is run, so before running 
    this function, you might run
        C:\PS> $Credential = Get-Credential

    This example causes existing date columns to have the Standard (rather than Friendly) format.
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus.sharepoint.com/teams/sitename" `
           -FriendlyDate "Disabled" -Credential $Credential

.EXAMPLE
If you are setting the format across a number of sites, consider splatting the 
common parameters, e.g.
    C:\PS> $Params = @{'FriendlyDate'='Disabled';'Credential'=$Credential}
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus1.sharepoint.com/teams/sitename" @Params
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus2.sharepoint.com/teams/sitename" @Params
    C:\PS> Set-SitewideDateFormat -SiteURL "https://bogus3.sharepoint.com/teams/sitename" @Params

.NOTES
    Author: Linda Busdiecker
    Date: 08/21/2019

    Many thanks to @SharePointDiary - this work was based off of
    https://www.sharepointdiary.com/2017/01/how-to-change-friendly-date-format-in-sharepoint.html
#>
#Requires -Module SharePointPnPPowerShellOnline 
function Set-SitewideDateFormat {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] 
        [String] $SiteURL,

        [Parameter(Mandatory = $true)] 
        [ValidateSet("Undefined", "Disabled", "Relative")] 
        [Microsoft.SharePoint.Client.DateTimeFieldFriendlyFormatType] $FriendlyDate,

        
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    Try {
        $Site = Connect-PnPOnline -Url $SiteUrl -Credential $Credential -ReturnConnection -ErrorAction Stop

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
                    Write-Verbose "Before: List [$($List.Title)], Field [$($Field.Title)] value: [$($Field.FriendlyDisplayFormat)]"
                    $Field.FriendlyDisplayFormat = $FriendlyDate
                    $Field.Update()
                    Write-Verbose "After: List [$($List.Title)], Field [$($Field.Title)] value: [$($Field.FriendlyDisplayFormat)]`n"
                }
            }
        }
    }
    Catch {
        Write-Error "Error - $_"
    }
}
