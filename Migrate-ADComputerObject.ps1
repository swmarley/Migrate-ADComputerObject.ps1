Function Migrate-ADComputerObject 
{
<#
.SYNOPSIS
    Moves one or more Active Directory objects into a different Organizational Unit and provides logging options in CSV and HTML formats.

.PARAMETER ComputerName
    Optional parameter that can be used to target an AD object by its computer name. 
    (Multiple values are accepted if separated by a comma)

.PARAMETER InputFile
    Optional parameter that can be used to target multiple objects by pointing to a text file containing a list of computer names.
    (Only one value for a text file is accepted)

.PARAMETER SourceOU
    Optional parameter to target all objects currently in a particular Organizational Unit.
    (Value must be entered in the format "OU=test,DC=contoso,DC=com")

.PARAMETER DestinationOU
    Mandatory parameter to specify the destination Organizational Unit for the targeted object(s).
    (Value must be entered in the format "OU=test,DC=contoso,DC=com")

.PARAMETER Report
    Optional parameter to generate a report showing the objects to be moved, along with their current Organizational Unit and the destination.
    (Report can be generated either in CSV or HTML and is saved by default to the desktop of the user running the script.

.EXAMPLE
    Migrate-ADComputerObject -ComputerName TESTMACHINE1 -DestinationOU "OU=test,DC=contoso,DC=com"

.EXAMPLE   
    Migrate-ADComputerObject -InputFile "C:\Users\Test\Desktop\servers.txt" -DestinationOU "OU=test,DC=contoso,DC=com" -Report CSV

.EXAMPLE
    Migrate-ADComputerObject -SourceOU "OU=old,DC=contoso,DC=com" -DestinationOU "OU=new,DC=contoso,DC=com" -Report HTML

.NOTES
    1.0 | 5/14/2018 | Sutton Marley
        Initial Version
#>
    [CmdletBinding()]

    PARAM (
        [Parameter(Mandatory=$false)]
        [ValidateScript({Get-ADComputer -Identity $_})]
        [String[]] $ComputerName,

        [Parameter(Mandatory=$false)]
        [ValidateScript({Test-Path -Path $_})]
        [String] $InputFile,

        [Parameter(Mandatory=$false)]
        [ValidateScript({Get-ADObject -Identity $_})]
        [String] $SourceOU,

        [Parameter(Mandatory=$true, HelpMessage="A destination OU is needed before AD object can be moved.")]
        [ValidateScript({Get-ADObject -Identity $_})]
        [String] $DestinationOU,

        [Parameter(Mandatory=$false)]
        [ValidateSet("CSV","HTML")]
        [String] $Report
    )


    BEGIN {
        TRY {
            #Checks for script pre-reqs (ActiveDirectory module and at least one target parameter)
            If (-not (Get-Module -Name ActiveDirectory)) {Import-Module -Name ActiveDirectory -ErrorAction Stop}
            If (!$ComputerName -and !$InputFile -and !$SourceOU) {
                throw "At least one parameter (ComputerName,InputFile, or SourceOU) must be specified..."
                }
        }
        CATCH{
            Write-Warning -Message "[BEGIN] An error occurred."
        }
    }

    PROCESS {
        
        Function Get-SourceOUMembers {
            [CmdletBinding()]

            PARAM (
                [String] $Source
            )

            $groupmembers = Get-ADComputer -Filter * -SearchBase $Source

            return $groupmembers
        }

    If ($Report) {

        Function Get-SourceOU {
            [CmdletBinding()]

            PARAM (
                [String] $Object
            )
            #Gets source OU information and formats to same syntax as DestinationOU parameter in main function
            $distinguishedName = Get-ADComputer $Object | select -ExpandProperty DistinguishedName
            $retrievedSourceOU = $distinguishedName.substring($distinguishedName.IndexOf(",") + 1)

            return $retrievedSourceOU
        }
        
        Function New-CSVReport {
            [CmdletBinding()]
            PARAM (
               [object] $Object
            )
            $CSVTimeStamp = Get-Date -Format M-d-yyyy-hh-mm
            $CSVPath = "C:\Users\$($env:USERNAME)\Desktop\ObjectMigrationLog-$($CSVTimeStamp).csv"
            $Object | Export-Csv -Path $CSVPath -Append
        }

        Function New-HTMLReport {
            [CmdletBinding()]
            PARAM (
               [object] $Object
            )
            #CSS table customizations for HTML report
            $Header = @"
<style>
table {font-family: "Arial"; font-weight: lighter;}
th {text-align: center; background-color: #4682B4; padding: .75em; color: white;}
td {text-align: center;}
table tr:nth-child(even) {background-color: #FFFFFF;}
table tr:nth-child(odd) {background-color: #DCDCDC;}
</style>
"@
            $HTMLTimeStamp = Get-Date -Format M-d-yyyy-hh-mm
            $HTMLPath = "C:\Users\$($env:USERNAME)\Desktop\ObjectMigrationLog-$($HTMLTimeStamp).html"
            $Object | ConvertTo-Html -Head $Header | Out-File -FilePath $HTMLPath -Append
        }

        Function New-ReportObject {
            [CmdletBinding()]
            PARAM (
                [string] $Name
            )

            $ReportObject = [PSCustomObject] @{
                ComputerName = $Name
                SourceOU = Get-SourceOU -Object $Name
                DestinationOU = $DestinationOU
            }

            return $ReportObject
        }

        Function Get-ReportPreference {
            [CmdletBinding()]
            PARAM (
                $Object 
            )

            If ($Report -eq "CSV") {
                New-CSVReport -Object $Object
            }

            If ($Report -eq "HTML") {
                New-HTMLReport -Object $Object
            }
        }
        #Builds report entries in specified format (CSV/HTML) when ComputerName parameter is used in main function
        If ($ComputerName) {
            #The parentObject variable contains report information for all objects and is directly used when generating reports
            $parentObject = @()

            foreach($computer in $ComputerName) {
                $childObject = New-ReportObject -Name $computer
                $parentObject += $childObject
            }

            Get-ReportPreference -Object $parentObject  
        }
        #Builds report entries in specified format (CSV/HTML) when InputFile parameter is used in main function
        If ($InputFile) {

            $servers = Get-Content $InputFile
            $parentObject = @()

           
            foreach($server in $servers) {
                $childObject = New-ReportObject -Name $server
                $parentObject += $childObject
            }

            Get-ReportPreference -Object $parentObject  
        }
        #Builds report entries in specified format (CSV/HTML) when SourceOU parameter is used in main function
        If ($SourceOU) {
            $servers = Get-SourceOUMembers -source $SourceOU | select -ExpandProperty Name
            $parentObject =@()

            foreach($server in $servers) {
                $childObject = New-ReportObject -Name $server
                $parentObject += $childObject
            }

            Get-ReportPreference -Object $parentObject
        }
    }

        TRY{
            #This is where the actual moving of objects in AD is defined for each paramter type of the main function
            If ($ComputerName){
                foreach ($computer in $ComputerName) {
                    $name = Get-ADComputer $computer | Select -ExpandProperty Name
                    Get-ADComputer $name | Move-ADObject -TargetPath $DestinationOU
                }
            }

            If ($InputFile){
                $targetlist = Get-Content $InputFile
                foreach($computer in $targetlist){
                    $name = Get-ADComputer $computer | Select -ExpandProperty Name
                    Get-ADComputer $name | Move-ADObject -TargetPath $DestinationOU
                }
            }

            If ($SourceOU){
                $targetlist = Get-SourceOUMembers -source $SourceOU | select -ExpandProperty Name
                foreach ($computer in $targetlist){
                    Get-ADComputer $computer | Move-ADObject -TargetPath $DestinationOU
                }
            }

        }
        CATCH{
            Write-Warning -Message "[PROCESS] An error occurred."
        }
    }

    END{}
} 
