<#
.SYNOPSIS
    Script to retrieve WhoIs information from a list of domains
.DESCRIPTION
    This script will, by default, create a report of WhoIs information on 1 or 
    more Internet domains. 
    
    Report options are CSV, XML, HTML and object (default) output.  Dates in the CSV and HTML options
    are formatted for the culture settings on the PC.  Columns in HTML report are also sortable,
    just click on the headers.
    
.PARAMETER Domain
    One or more domain names to check.  Accepts pipeline.
    
.PARAMETER Path
    Path where the resulting HTML or CSV report will be saved.  If not specified
    report will be saved in the same folder where the script is located.
    
.PARAMETER RedThresold
    If the number of days left before the domain expires falls below this number 
    the entire row will be highlighted in Red (HTML reports only).
    
.PARAMETER YellowThresold
    If the number of days left before the domain expires falls below this number 
    the entire row will be highlighted in Yellow (HTML reports only).
    
.PARAMETER GreyThresold
    If the number of days left before the domain expires falls below this number 
    the entire row will be highlighted in Grey (HTML reports only).
    
.PARAMETER OutputType
    Specify what kind of report you want.  Valid types are HTML, CSV or Object.
    
.INPUTS
    String that only contains domain name
.OUTPUTS
    HTML report
    CSV report
    XML report
    PSCustomObject
        DomainName
        Registrar
        WhoIsServer
        NameServers
        DomainLock
        LastUpdated
        Created
        Expiration
        DaysLeft

.EXAMPLE
    Get-Content c:\domains.txt | .\Get-Whois.ps1 -RedThreshold 160 -YellowThresold 365 -GreyThreshold 720 -OutputType HTML -Verbose
    
    Take every domain in domains.txt and produce an HTML report.  Any domain expiring 
    in 160 days or less will be highlighted Red.  Expiring in 365 days will be Yellow
    and 720 days will be Grey.  Script will communicate about what it's doing as it goes
    though the different domains (Verbose output) but this information will not be
    included in the report.
    
.EXAMPLE
    .\Get-Whois.ps1 -Domain "thesurlyadmin.com","google.com"
    
    Will create object output of the domain registration data.
    
.EXAMPLE
    .\Get-Whois.ps1 -Domain (Get-Content c:\domains.txt) -OutputType CSV
    
    Take every domain in domains.txt and produce a CSV report.
    
.NOTES
    Author:             Martin Pugh
    Twitter:            @thesurlyadm1n
    Spiceworks:         Martin9700
    Blog:               www.thesurlyadmin.com
      
    Changelog:
        2.0             Complete rewrite, no longer uses JSON feed but a web-app instead.  Added XML output support.
        1.0             Initial Release
.LINK
    http://community.spiceworks.com/scripts/show/2809-whois-report-get-whois-ps1
#>
#requires -Version 3.0
[CmdletBinding()]
Param(
    [Parameter(Mandatory,ValueFromPipeline,Position=1)]
    [String[]]$Domain,
    
    [Parameter(Position=2)]
    [string]$Path,
    
    [Parameter(Position=0)]
    [int]$RedThresold = 30,
    
    [Parameter(Position=0)]
    [int]$YellowThresold = 90,
    
    [Parameter(Position=0)]
    [int]$GreyThresold = 365,
    
    [Parameter(Position=0)]
    [ValidateSet("object","html","csv","xml")]
    [string]$OutputType = "object"
)

Begin {
    Write-Verbose "$(Get-Date): Get-WhoIs.ps1 script beginning." -Verbose
    
    #Validate the path
    If ($Path)
    {   If (Test-Path $Path)
        {   If (-not (Get-Item $Path).PSisContainer)
            {   Throw "You cannot specify a file in the Path parameter, must be a folder: $Path"
            }
        }
        Else
        {   Throw "Unable to locate: $Path"
        }
    }
    Else
    {   $Path = Split-Path $MyInvocation.MyCommand.Path
    }
    
    #Create the Web Proxy instance
    $WC = New-WebServiceProxy 'http://www.webservicex.net/whois.asmx?WSDL'
    $Data = @()
}

Process {
    $Data += ForEach ($Dom in $Domain)
    {   Write-Verbose "$(Get-Date): Querying for $Dom"
        $DNError = ""
        Try {
            $Raw = $WC.GetWhoIs($Dom)
        }
        Catch {
            #Some domains throw an error, I assume because the WhoIs server isn't returning standard output
            $DNError = "$($Dom.ToUpper()): Unknown Error retrieving WhoIs information"
        }

        #Test if the domain name is good or if the data coming back is ok--Google.Com just returns a list of domain names so no good
        If ($Raw -match "No match for")
        {   $DNError = "$($Dom.ToUpper()): Unable to find registration for domain"
        }
        ElseIf ($Raw -notmatch "Domain Name: (.*)")
        {   $DNError = "$($Dom.ToUpper()): WhoIs data not in correct format"
        }
        
        If ($DNError)
        {   #Use 999899 to tell the script later that this is a bad domain and color it properly in HTML (if HTML output requested)
            [PSCustomObject]@{
                DomainName = $DNError
                Registrar = ""
                WhoIsServer = ""
                NameServers = ""
                DomainLock = ""
                LastUpdated = ""
                Created = ""
                Expiration = ""
                DaysLeft = 999899
            }
            Write-Warning $DNError
        }
        Else
        {   #Parse out the DNS servers
            $NS = ForEach ($Match in ($Raw | Select-String -Pattern "Name Server: (.*)" -AllMatches).Matches)
            {   $Match.Groups[1].Value
            }

            #Parse out the rest of the data
            [PSCustomObject]@{
                DomainName = ($Raw | Select-String -Pattern "Domain Name: (.*)").Matches.Groups[1].Value
                Registrar = ($Raw | Select-String -Pattern "Registrar: (.*)").Matches.Groups[1].Value
                WhoIsServer = ($Raw | Select-String -Pattern "WhoIs Server: (.*)").Matches.Groups[1].Value
                NameServers = $NS -join ", "
                DomainLock = ($Raw | Select-String -Pattern "Status: (.*)").Matches.Groups[1].Value
                LastUpdated = [datetime]($Raw | Select-String -Pattern "Updated Date: (.*)").Matches.Groups[1].Value
                Created = [datetime]($Raw | Select-String -Pattern "Creation Date: (.*)").Matches.Groups[1].Value
                Expiration = [datetime]($Raw | Select-String -Pattern "Expiration Date: (.*)").Matches.Groups[1].Value
                DaysLeft = (New-TimeSpan -Start (Get-Date) -End ([datetime]($Raw | Select-String -Pattern "Expiration Date: (.*)").Matches.Groups[1].Value)).Days
            }
        }
    }
}

End {
    Write-Verbose "$(Get-Date): Producing $OutputType report"
    $WC.Dispose()
    $Data = $Data | Sort DomainName

    Switch ($OutputType)
    {   "object"
        {   Write-Output $Data | Select DomainName,Registrar,WhoIsServer,NameServers,DomainLock,LastUpdated,Created,Expiration,@{Name="DaysLeft";Expression={If ($_.DaysLeft -eq 999899) { 0 } Else { $_.DaysLeft }}}
        }
        "csv"
        {   $ReportPath = Join-Path -Path $Path -ChildPath "WhoIs.CSV"
            $Data | Select DomainName,Registrar,WhoIsServer,NameServers,DomainLock,@{Name="LastUpdated";Expression={Get-Date $_.LastUpdated -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},@{Name="Created";Expression={Get-Date $_.Created -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},@{Name="Expiration";Expression={Get-Date $_.Expiration -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},DaysLeft | Export-Csv $ReportPath -NoTypeInformation
        }
        "xml"
        {   $ReportPath = Join-Path -Path $Path -ChildPath "WhoIs.XML"
            $Data | Select DomainName,Registrar,WhoIsServer,NameServers,DomainLock,LastUpdated,Created,Expiration,@{Name="DaysLeft";Expression={If ($_.DaysLeft -eq 999899) { 0 } Else { $_.DaysLeft }}} | Export-Clixml $ReportPath
        }
        "html"
        {   
            $Header = @"
<script src="http://kryogenix.org/code/browser/sorttable/sorttable.js"></script>
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TR:Hover TD {Background-Color: #C1D5F8;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;cursor: pointer;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
<title>
WhoIS Report
</title>
"@

            $PreContent = @"
<p><h1>WhoIs Report</h1></p>
"@

            $PostContent = @"
<p><br/><h3>Legend</h3>
<pre><span style="background-color:red">    </span>  Expires in under $RedThreshold days
<span style="background-color:yellow">    </span>  Expires in under $YellowThreshold days
<span style="background-color:#B0C4DE">    </span>  Expires in under $GreyThreshold days
<span style="background-color:#DEB887">    </span>  Problem retrieving information about domain/Domain not found</pre></p>
<h6><br/>Run on: $(Get-Date)</h6>
"@

            $RawHTML = $Data | Select DomainName,Registrar,WhoIsServer,NameServers,DomainLock,@{Name="LastUpdated";Expression={Get-Date $_.LastUpdated -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},@{Name="Created";Expression={Get-Date $_.Created -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},@{Name="Expiration";Expression={Get-Date $_.Expiration -Format (Get-Culture).DateTimeFormat.ShortDatePattern}},DaysLeft | ConvertTo-Html -Head $Header -PreContent $PreContent -PostContent $PostContent 

            $HTML = ForEach ($Line in $RawHTML)
            {   If ($Line -like "*<tr><td>*")
                {   $Value = [float](([xml]$Line).SelectNodes("//td").'#text'[-1])
                    If ($Value)
                    {   If ($Value -eq 999899)
                        {   $Line.Replace("<tr><td>","<tr style=""background-color: #DEB887;""><td>").Replace("<td>999899</td>","<td>0</td>")
                        }
                        ElseIf ($Value -lt $RedThreshold)
                        {   $Line.Replace("<tr><td>","<tr style=""background-color: red;""><td>")
                        }
                        ElseIf ($Value -lt $YellowThreshold)
                        {   $Line.Replace("<tr><td>","<tr style=""background-color: yellow;""><td>")
                        }
                        ElseIf ($Value -lt $GreyThreshold)
                        {   $Line.Replace("<tr><td>","<tr style=""background-color: #B0C4DE;""><td>")
                        }
                        Else
                        {   $Line
                        }
                    }
                }
                ElseIf ($Line -like "*<table>*")
                {   $Line.Replace("<table>","<table class=""sortable"">")
                }
                Else
                {   $Line
                }
            }
            
            $ReportPath = Join-Path -Path $Path -ChildPath "WhoIs.html"
            $HTML | Out-File $ReportPath -Encoding ASCII
            
            #Immediately display the html if in debug mode
            If ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent)
            {   & $ReportPath
            }
        }
    }

    Write-Verbose "$(Get-Date): Get-WhoIs.ps1 script completed!" -Verbose
}
