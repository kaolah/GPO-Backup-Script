#SCRIPT TO BACKUP GPO, GPO REPORTS, WMI Filters

# Variables that needs to be changed

$Server = "dc1"
$Domain = "ko.local"
$ServerFQDN = "$Server.$Domain"
$RetentionPeriod = "10"
$RootBackupFolder = "C:\Backup\GPO\GPOBackups\"
$RootReportFolder = "C:\Backup\GPO\GPOReports\"
$RootWMIFolder = "C:\Backup\GPO\WMI\"

Import-Module GroupPolicy
Import-Module ActiveDirectory

# BACKUP CURRENT GPOs and GENERATE REPORTS 

If (!(Test-path C:\Backup\GPO\GPOBackups)) {
	new-item -path C:\Backup\GPO\GPOBackups -itemType Directory -force
	}
If (!(Test-path C:\Backup\GPO\GPOReports)) {
	new-item -path C:\Backup\GPO\GPOReports -itemType Directory -force
	}
If (!(Test-path C:\Backup\GPO\WMI)) {
	new-item -path C:\Backup\GPO\WMI -itemType Directory -force
	}

$Prefix = (get-date).ToString("yyyyMMdd_HHmmss")

$BackupFolder = join-path $RootBackupFolder $Prefix

$ReportFolder = join-path $RootReportFolder $Prefix

$WMIFolder = join-path $RootWMIFolder $Prefix

new-item -path $BackupFolder -itemType Directory -force

new-item -path $ReportFolder -itemType Directory -force

new-item -path $WMIFolder -itemType Directory -force

#Backup-GPO –All –Path $BackupFolder –server $Server
#Get-GPO –All –Domain $Domain –server $Server | % { Get-GPOReport –Name $_.DisplayName –ReportType HTML –Path “$ReportFolder\$($_.DisplayName).html” –server $Server }

$allGPOs = Get-GPO –All –Domain $Domain –server $ServerFQDN | sort displayname

foreach ($gpo in $allGPOs) {
    Backup-GPO -Name $gpo.DisplayName -Path $BackupFolder -server $ServerFQDN -Domain $Domain
    Get-GPOReport –Name $gpo.DisplayName –ReportType HTML –Path “$ReportFolder\$($gpo.DisplayName).html” –server $ServerFQDN -Domain $Domain
    }



# BACKUP WMI FILTERS -- XML file
#Connect to the Active Directory to get details of the WMI filters
    $WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' `
                               -Properties msWMI-Author, msWMI-ID, msWMI-Name, msWMI-Parm1, msWMI-Parm2 `
                               -Server $ServerFQDN `
                               -ErrorAction SilentlyContinue


#If $WMIFilters contains objects write these to an XML file
    If ($WmiFilters) {

        #Create a variable for the XML file representing the WMI filters
        $WmiXML = "$WMIFolder\WmiFilters.xml"

        #Export our array of WMI filters to XML so they can be easily re-imported as objects
        $WmiFilters | Export-Clixml -Path $WmiXML

    }   #End of If ($WmiFilters)

# BACKUP WMI FILTERS -- TXT file
get-adobject –LDAPFilter "(ObjectClass=msWMI-som)" -properties * | Sort-Object msWMI-Name | Select @{Name="Name";Expression= {$_.'msWMI-Name'}}, @{Name="GUID";Expression={$_.Name}}, @{Name="Description";Expression={$_.'msWMI-Parm1'}}, @{Name="Namespace";Expression={$_.'msWMI-Parm2'.Split(";")[-3] }}, @{Name="WMIQuery";Expression={$_.'msWMI-Parm2'.Split(";")[-2] }}, Created,Modified | Out-File $WMIFolder\WmiFilters.txt


# REPORT to CSV

# Grab a list of all GPOs
$GPOs = Get-GPO -All –Domain $Domain –server $ServerFQDN | Select-Object ID, Path, DisplayName, GPOStatus, WMIFilter

# Create a hash table for fast GPO lookups later in the report.
# Hash table key is the policy path which will match the gPLink attribute later.
# Hash table value is the GPO object with properties for reporting.
$GPOsHash = @{}
ForEach ($GPO in $GPOs) {
    $GPOsHash.Add($GPO.Path,$GPO)
}

# Empty array to hold all possible GPO link SOMs
$gPLinks = @()

# GPOs linked to the root of the domain
#  !!! Get-ADDomain does not return the gPLink attribute
$gPLinks += `
 Get-ADObject -Identity (Get-ADDomain).distinguishedName -Properties name, distinguishedName, gPLink, gPOptions |
 Select-Object name, distinguishedName, gPLink, gPOptions, @{name='Depth';expression={0}}

# GPOs linked to OUs
#  !!! Get-GPO does not return the gPLink attribute
# Calculate OU depth for graphical representation in final report
$gPLinks += `
 Get-ADOrganizationalUnit -Filter * -Properties name, distinguishedName, gPLink, gPOptions |
 Select-Object name, distinguishedName, gPLink, gPOptions, @{name='Depth';expression={($_.distinguishedName -split 'OU=').count - 1}}

# GPOs linked to sites
$gPLinks += `
 Get-ADObject -LDAPFilter '(objectClass=site)' -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions |
 Select-Object name, distinguishedName, gPLink, gPOptions, @{name='Depth';expression={0}}

# Empty report array
$report = @()

# Loop through all possible GPO link SOMs collected
ForEach ($SOM in $gPLinks) {
    # Filter out policy SOMs that have a policy linked
    If ($SOM.gPLink) {
        # If an OU has 'Block Inheritance' set (gPOptions=1) and no GPOs linked,
        # then the gPLink attribute is no longer null but a single space.
        # There will be no gPLinks to parse, but we need to list it with BlockInheritance.
        If ($SOM.gPLink.length -gt 1) {
            # Use @() for force an array in case only one object is returned (limitation in PS v2)
            # Example gPLink value:
            #   [LDAP://cn={7BE35F55-E3DF-4D1C-8C3A-38F81F451D86},cn=policies,cn=system,DC=wingtiptoys,DC=local;2][LDAP://cn={046584E4-F1CD-457E-8366-F48B7492FBA2},cn=policies,cn=system,DC=wingtiptoys,DC=local;0][LDAP://cn={12845926-AE1B-49C4-A33A-756FF72DCC6B},cn=policies,cn=system,DC=wingtiptoys,DC=local;1]
            # Split out the links enclosed in square brackets, then filter out
            # the null result between the closing and opening brackets ][
            $links = @($SOM.gPLink -split {$_ -eq '[' -or $_ -eq ']'} | Where-Object {$_})
            # Use a for loop with a counter so that we can calculate the precedence value
            For ( $i = $links.count - 1 ; $i -ge 0 ; $i-- ) {
                # Example gPLink individual value (note the end of the string):
                #   LDAP://cn={7BE35F55-E3DF-4D1C-8C3A-38F81F451D86},cn=policies,cn=system,DC=wingtiptoys,DC=local;2
                # Splitting on '/' and ';' gives us an array every time like this:
                #   0: LDAP:
                #   1: (null value between the two //)
                #   2: distinguishedName of policy
                #   3: numeric value representing gPLinkOptions (LinkEnabled and Enforced)
                $GPOData = $links[$i] -split {$_ -eq '/' -or $_ -eq ';'}
                # Add a new report row for each GPO link
                $report += New-Object -TypeName PSCustomObject -Property @{
                    Depth             = $SOM.Depth;
                    Name              = $SOM.Name;
                    DistinguishedName = $SOM.distinguishedName;
                    PolicyDN          = $GPOData[2];
                    Precedence        = $links.count - $i
                    GUID              = "{$($GPOsHash[$($GPOData[2])].ID)}";
                    DisplayName       = $GPOsHash[$GPOData[2]].DisplayName;
                    GPOStatus         = $GPOsHash[$GPOData[2]].GPOStatus;
                    WMIFilter         = $GPOsHash[$GPOData[2]].WMIFilter.Name;
                    Config            = $GPOData[3];
                    LinkEnabled       = [bool](!([int]$GPOData[3] -band 1));
                    Enforced          = [bool]([int]$GPOData[3] -band 2);
                    BlockInheritance  = [bool]($SOM.gPOptions -band 1)
                } # End Property hash table
            } # End For
        } Else {
            # BlockInheritance but no gPLink
            $report += New-Object -TypeName PSCustomObject -Property @{
                Depth             = $SOM.Depth;
                Name              = $SOM.Name;
                DistinguishedName = $SOM.distinguishedName;
                BlockInheritance  = [bool]($SOM.gPOptions -band 1)
            }
        } # End If
    } Else {
        # No gPLink at this SOM
        $report += New-Object -TypeName PSCustomObject -Property @{
            Depth             = $SOM.Depth;
            Name              = $SOM.Name;
            DistinguishedName = $SOM.distinguishedName;
            BlockInheritance  = [bool]($SOM.gPOptions -band 1)
        }
    } # End If
} # End ForEach

# Output the results to CSV file for viewing in Excel
$report |
 Select-Object @{name='SOM';expression={$_.name.PadLeft($_.name.length + ($_.depth * 5),'_')}}, `
  DistinguishedName, BlockInheritance, LinkEnabled, Enforced, Precedence, `
  DisplayName, GPOStatus, WMIFilter, GUID, PolicyDN |
 Export-CSV $ReportFolder\___gPLink_Report.csv -NoTypeInformation -Encoding UTF8


# REMOVE BACKUPS OLDER THAN $RetentionPeriod

Get-ChildItem -path $RootBackupFolder -Directory | Sort-Object -Property Name -Descending | Select-Object -Skip $RetentionPeriod |
	ForEach-Object { remove-item $_.fullname -recurse -force }

# REMOVE REPORTS OLDER THAN $RetentionPeriod

Get-ChildItem -path $RootReportFolder -Directory | Sort-Object -Property Name -Descending | Select-Object -Skip $RetentionPeriod |  
	ForEach-Object { remove-item $_.fullname -recurse -force }

# REMOVE WMI FILTERS OLDER THAN $RetentionPeriod

Get-ChildItem -path $RootWMIFolder -Directory | Sort-Object -Property Name -Descending | Select-Object -Skip $RetentionPeriod | 
	ForEach-Object { remove-item $_.fullname -recurse -force }

#END OF BACKUP SCRIPT
