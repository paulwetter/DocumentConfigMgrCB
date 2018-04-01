#Requires -Version 4.0

#region help text

<#
.SYNOPSIS
	Script attempts to fully document a Microsoft Configuration Manager environment.
.DESCRIPTION
	This script will fully document a Configuration Manager environment.  The original 
    script developed several years ago by David O'Brien required Microsoft Word to create 
    the documentation.  This updated script is more detailed and outputs the documentation
    in pure HTML.  If you so desire, you can import this HTML report into Word for easier
    editing.
.PARAMETER Title
	The title you would like to use for this documentation.  default is "Configuration Manager Site Documentation".
.PARAMETER FilePath
	This is the path of the documentation file.  By default, the file will be created in the same directory as the
    where the script is currently located. And named CMDocumentation.html
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be ReportName_2014-06-01_1800.html.
.PARAMETER CompanyName
	This is the name of the company or organization that the documentation will be created for.
.PARAMETER CompanyLogo
    This is a UNC or URL path to a image file jpg, or png to embed into the document on the title page.  By default,
    the Cyber Advisors logo will display.
.PARAMETER Author
    This is the report author.  Their name appears in the lower right corner of the title page.
.PARAMETER Vendor
    This displays a company name in the lower right corner of the title page.
.PARAMETER Software
    Specifies whether the script should run an inventory of Applications, Packages and OSD related objects.
.PARAMETER ListAllInformation
    Specifies whether the script should only output an overview of what is configured (like count of collections) or 
    a full output with verbose information.
.PARAMETER SMSProvider
    Some information rely on WMI queries that need to be executed against the SMS Provider directly. 
    Please specify as FQDN.
    If not specified, it assumes localhost.
.PARAMETER UnknownClientSettings
    With new releases of CM come new client settings.  If this parameter is added, it will display raw 
    information for these client settings.
.EXAMPLE
	DocumentCMCB.ps1 -ListAllInformation
.EXAMPLE
	DocumentCMCB.ps1 -CompanyLogo 'http://www.contoso.com/logo.jpg' -ListAllInformation
.EXAMPLE
	DocumentCMCB.ps1 -CompanyLogo 'http://www.contoso.com/logo.jpg' -Author "Bugs Bunny" -Vendor "Acme" -ListAllInformation
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a HTML document.
.NOTES
	NAME: DocumentCMCB.ps1
	VERSION: 1.00
	AUTHOR: Paul Wetter, David O'Brien
	LASTEDIT: March 07, 2018
#>

#endregion

#region script parameters
#[CmdletBinding(DefaultParameterSetName = "Word") ]
[CmdletBinding()]

Param(
	[parameter(Mandatory=$True)] 
	[string]$CompanyName,
    
	[parameter(Mandatory=$False)] 
    [string]$CompanyLogo = "https://blog.cyberadvisors.com/hubfs/CAI_logo.jpg",

	[parameter(Mandatory=$False)] 
	[Switch]$ListAllInformation,

	[parameter(Mandatory=$False)] 
	[string]$Author="Paul Wetter",

	[parameter(Mandatory=$False)] 
    [string]$Vendor = "Cyber Advisors",

	[parameter(Mandatory=$False)] 
	[String]$Title = "Configuration Manager Site Documentation",

	[parameter(Mandatory=$False)] 
	[String]$FilePath = "$([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))\CMDocumentation.html",

	[parameter(Mandatory=$False)] 
	[Switch]$Software,

	[parameter(Mandatory=$False)] 
	[string]$SMSProvider='localhost',

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
    [switch]$UnknownClientSettings

	)
#endregion

<#
$FilePath = "C:\Users\Administrator\Desktop\TestDocumentation.html"
$Title = "Configuration Manager Site Documentation"
$Author = "Snow Lizard"
$Vendor = "FYLCP"
$Org = "Wetter Servicing"
$SMSProvider = 'localhost'
$UnknownClientSettings = $true
$ListAllInformation = $true
#>

$CMPSSuppressFastNotUsedCheck = $true
$Global:DocTOC = @()

Function Write-HtmlTable{
<#
.SYNOPSIS
    This will take an input array of objects and turn it into an HTML table.  Optionally, you can set a border for the table as well.
.PARAMETER InputObject
    This is an array of objects that will be built into a HTML table.
.PARAMETER Padding
    This is the amount of space in each field between the border and the text.
.PARAMETER Spacing
    This is the amount of space between the borders of each field.  Rarely anything other than zero (0).
.PARAMETER Level
    This is the amount of space that the table will indented by.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HtmlTable -InputObject $folders -Border 1
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is an array of objects that will be built into a HTML table.")]
        $InputObject,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="If this table has a border, select the thickness here. Default:1")]
        [int]$Border=1,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space in each field between the border and the text. Default:3")]
        [int]$Padding=3,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space between the borders of each field.  Rarely anything other than zero (0).")]
        [int]$Spacing=0,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the table will indent by")]
        [ValidateRange(0,9)]
        [int]$Level=0,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the file that the HTML will be written to")]
        [string]$File
    )

    switch ($Level) 
    { 
        0 {$Indent=0} 
        1 {$Indent=5} 
        2 {$Indent=15} 
        3 {$Indent=25} 
        4 {$Indent=35} 
        5 {$Indent=45}
        6 {$Indent=55} 
        7 {$Indent=65} 
        8 {$Indent=75} 
        9 {$Indent=85} 
        default {$Indent=5}
    }
    $table = $InputObject|ConvertTo-Html -Fragment
    $table[0] = "<table cellpadding=$Padding cellspacing=$Spacing border=$Border style=`"margin-left:$($Indent)px;`">"
    $table = $table -replace "--CRLF--","<BR />"
    If ($File) {$table | Out-File -filepath $File -Append}
    Else {Return $table}
}

Function Write-HtmlList{
<#
.SYNOPSIS
    This will take an input array of strings and turn them into an HTML list.  This can be an ordered or unordered list (Numbered or bulleted).
.PARAMETER InputObject
    This is an array of strings that will be made into the list.
.PARAMETER Title
    This is the title text for the list.
.PARAMETER Description
    This is html formatted test that will appear as a description or paragraph between the Title and actual list.
.PARAMETER Level
    This is the amount of space that the list will indented by.
.PARAMETER Type
    Choose ordered (OL) or unordered (UL) list.  Unordered or bulleted is the default
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HtmlList -InputObject @('Red','Blue','Green','Yellow') -Title "Colors of the Rainbow" -Description "This is a <i>list</i> of colors in the rainbow." -level 1
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is an array of strings that will be made into the list.")]
        $InputObject,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the title text for the list")]
        $Title,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is html formatted test that will appear as a description or paragraph between the Title and actual list")]
        $Description,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the list will indent by")]
        [ValidateRange(0,6)]
        [int]$Level=0,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="Choose ordered (OL) or unordered (UL) list.  Unordered or bulleted is the default")]
        [ValidateSet("OL","UL")]
        [string]$Type="UL",
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the file that the HTML will be written to")]
        [string]$File
    )

    switch ($Level) 
    { 
        0 {$Indent=0} 
        1 {$Indent=5} 
        2 {$Indent=15} 
        3 {$Indent=25} 
        4 {$Indent=35} 
        5 {$Indent=45}
        6 {$Indent=55} 
        default {$Indent=5}
    }
    if ($Title)
        {
          $ListHTML = "<P style=`"margin-left:$($indent)px;`"><B>$($Title)</B>"
        }Else{
          $ListHTML = "<P style=`"margin-left:$($indent)px;`">"
        }
    if ($Description)
        {
          $ListHTML = $ListHTML + "<Div style=`"margin-left:$($indent + 5)px;`">$Description</div>"
        }
    $GroupList = "<$Type style=`"margin-left:$($indent)px;margin-top:0px;`">"
    If ($InputObject){
        foreach ($Item in $InputObject)
            {
                $GroupList = $GroupList + "<LI>$Item</LI>"
            }
    }
    $GroupList = $GroupList + "</$Type>"
    $ListHTML = $ListHTML + $GroupList + "</P>"
    If ($File) {$ListHTML | Out-File -filepath $File -Append}
    Else {Return $ListHTML}
}



Function Write-HTMLHeading{
<#
.SYNOPSIS
    This will format text as a heading in HTML.  Optionally, it will add a page break to the heading so that when printing, it will appear on a new page.
.PARAMETER Text
    This is the text that will appear inside the heading <H#`> tag
.PARAMETER Level
    This is the level of the heading. 1,2,3,4,5,6 are valid options.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HTMLHeading -Text "Test Heading 1"
.EXAMPLE
    Write-HTMLHeading -Text "Test Heading 2" -Level 2
.EXAMPLE
    Write-HTMLHeading -Text "Test Heading 3" -Level 3 -PageBreak
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear inside the heading <H#`> tag")]
        [string]$Text,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the level of the heading. 1,2,3,4,5,6 are valid options")]
        [ValidateRange(1,6)]
        [int]$Level=1,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This will style the HTML so it will print a page break")]
        [switch]$PageBreak,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This will exclude the heading from the Table of Contents")]
        [switch]$ExcludeTOC,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the file that the HTML will be written to")]
        [string]$File
    )
    $PropertyID = $Text.Replace(' ','')
    If(-not $ExcludeTOC){
        $Global:DocTOC += New-Object -TypeName PSObject -Property @{'Level'=$level; 'Title'="$Text"; 'Id'=$PropertyID}
    }
    If ($PageBreak) {$HtmlClass = " Class=`"pagebreak`""}
    $HeadLine = "<H$Level$HtmlClass id=`"$PropertyID`">$Text</H$Level>"
    If ($File) {$HeadLine | Out-File -filepath $File -Append}
    Else {Return $HeadLine}
}


Function Write-HTMLParagraph{
<#
.SYNOPSIS
    This will format text as a paragraph in HTML.  Optionally, it will allow you to indent the text to match the headings.
.PARAMETER Text
    This is the text that will appear inside the heading <P> tag.
.PARAMETER Level
    This is the amount of space that the paragraph will indent by.  This is equivelent to the heading level indent +5.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HTMLParagraph -Text "This is a bunch of text. It is a lot to go into the paragraph."
.EXAMPLE
    Write-HTMLParagraph -Text "This is also a bunch of text. It is a lot to go into the paragraph as well." -Indent
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear inside the heading <P> tag")]
        [string]$Text,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the indent level of the paragraph")]
        [ValidateRange(0,6)]
        [int]$Level=0,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the table will indent by")]
        [string]$File
    )
    
    switch ($Level) 
    { 
        0 {$Indent=0} 
        1 {$Indent=5} 
        2 {$Indent=15} 
        3 {$Indent=25} 
        4 {$Indent=35} 
        5 {$Indent=45}
        6 {$Indent=55} 
        default {$Indent=5}
    }
    $Paragraph = "<P style=`"margin-left:$($Indent)px;`">$Text</p>"
    If ($File) {$Paragraph | Out-File -filepath $File -Append}
    Else {Return $Paragraph}
}



Function Write-HTMLHeader{
<#
.SYNOPSIS
    This will write the header for the document/HTML.  This also resets the document to no text (does not append to the document).
.PARAMETER Title
    This is the title for the document.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HTMLHeader -Title "This is a bunch of text for the title"
.EXAMPLE
    Write-HTMLHeader -Title "This is also a bunch of text for the title" -file "C:\test.html"
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear in title tag of the header")]
        [string]$Title,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the table will indent by")]
        [string]$File
    )
    $Header = @()
    $Header += "<html>"
    $Header += "<Head>"
    $Header += "<Title>$Title</Title>"
    $Header += "<Style>"
    $Header += 'H1  {background-color:royalblue; border-top: 1px solid black;}'
    $Header += 'H2	{margin-left:10px;background-color:steelblue; border-top: 1px solid black;}'
    $Header += 'H3	{margin-left:20px;background-color:lightblue; border-top: 1px solid black;}'
    $Header += 'H4	{margin-left:30px;background-color:lightsteelblue; border-top: 1px solid black;}'
    $Header += 'H5	{margin-left:40px;background-color:lightcyan; border-top: 1px solid black;}'
    $Header += 'H6	{margin-left:50px;background-color:lavender; border-top: 1px solid black;}'
    $Header += ".pagebreak { page-break-before: always; }"
    $Header += "TH  {background-color:LightBlue;padding: 3px; border: 2px solid black;}"
    $Header += "TD  {padding: 3px; border: 1px solid black;}"
    $Header += "TABLE	{border-collapse: collapse;}"
    $Header += "</Style>"
    $Header += "</Head>"
    $Header += "<Body>"
    If ($File) {IF (Test-Path -Path $File) {Remove-Item -Path $File -Force}}
    If ($File) {$header | Out-File -filepath $File -Append}
    Else {Return $Header}
}

Function Write-HTMLCoverPage{
<#
.SYNOPSIS
    This will write the title/cover page for the document.
.PARAMETER Title
    This is the title for the document.
.PARAMETER Author
    This is the name of the person that is creating the document.
.PARAMETER Vendor
    This is the name of the vendor that is creating the document.
.PARAMETER Org
    This is the organization that the documentation was created for.  Typically, they are the owner of the CM environment.
.PARAMETER ImagePath
    This is the path to an optional image to put on the cover page.  It will appear in the lower left of the body of the page.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HTMLCoverPage -Text "This is a bunch of text. It is a lot to go into the paragraph."
.EXAMPLE
    Write-HTMLCoverPage -Text "This is also a bunch of text. It is a lot to go into the paragraph as well." -Indent
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear in title tag of the header")]
        [string]$Title,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear in the lower right by line")]
        [string]$Author,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This text will also appear in to lower right by line, below")]
        [string]$Vendor,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="Will apprear in the top left, below the title.")]
        [string]$Org,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is an image logo that will be embedded in the title page")]
        [string]$ImagePath,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the file that the HTML text will be written to")]
        [string]$File
    )
    $Cover = @()
    $Cover += "<Table border=0 cellspacing=0 cellpadding=0 style=`"width:100%;border: 0px`">"
    $Cover += "<TR><TD Height=50 VAlign=`"top`" align=`"left`" style=`"border: 0px;font-size:48pt`">$Title</TD></TR>"
    $Cover += "<TR><TD Height=20 VAlign=`"top`" align=`"left`" style=`"border: 0px;font-size:24pt;padding-left:10px`">Report Prepared for: $Org</TD></TR>"
    If ($ImagePath){
        $ImageData=Convert-Image2Base64 -Path $ImagePath
    }
    If ($ImageData){
        $Cover += "<TR><TD Height=700 VAlign=`"bottom`" align=`"right`" style=`"border: 0px`"><img src=`"$ImageData`"></TD></TR>"
    }Else{
        $Cover += "<TR><TD Height=700 VAlign=`"top`" style=`"border: 0px`">&nbsp;</TD></TR>"
    }
    $Cover += "<TR><TD Height=30 VAlign=`"top`" Align=`"right`" style=`"border: 0px;font-size:18pt`">Report Prepared By: $Author</TD></TR>"
    If ($Vendor) {$Cover += "<TR><TD Height=30 VAlign=`"top`" Align=`"right`" style=`"border: 0px;font-size:24pt`">$Vendor</TD></TR>"}
    $Cover += "</Table>"
    If ($File) {$Cover | Out-File -filepath $File -Append}
    Else {Return $Cover}
}

Function Convert-Image2Base64{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName=$false,Mandatory=$true,ValueFromPipeline=$True,
        HelpMessage="this is a path to either a file on the web or locally on the network to convert")]
        [string]$Path
    )
    If (($Path -match '^[A-z]:\\.*(\.png|\.jpg)$') -or ($Path -match '^\\\\*\\.*(\.png|\.jpg)$')){
        If (Test-Path -Path "filesystem::$Path"){
            $EncodedImage = [convert]::ToBase64String((get-content $Path -encoding byte))
        }else{
            Write-Error "Path not found: $path"
            Return $false
        }
    }
    ElseIf ($Path -match '^http[s]://.*(\.png|\.jpg)$'){
        $ext=$Path.Substring($Path.Length-4)
        $tempfile = "${env:TEMP}\logo31337$ext"
        if (Test-Path $tempfile) {Remove-Item -Path $tempfile -Force}
        Try{Invoke-WebRequest -Uri $Path -OutFile $tempfile}
        Catch{
            Write-Host -ForegroundColor Yellow "Image for title page not found. Building title page without image."
            Return $false
        }
        $EncodedImage = [convert]::ToBase64String((get-content $tempfile -encoding byte))
    }else{
        Write-Error "Path does not match pattern: $path"
        Return $false
    }
    if($path.EndsWith(".jpg")){$imgtype = "jpg"}
    elseif($path.EndsWith(".png")){$imgtype = "png"}
    "data:image/$imgtype;base64,$EncodedImage"
}

function Write-HTMLTOC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the array of texts that will make up the table of contents")]
        $InputObject,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the file that the table of contents text will be written to")]
        [string]$File,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the key text that the table of contents will be inserted after. Each line of the HTML file is searched to find this text.  On each find, it will begin the insert.")]
        [string]$InsertPoint = "TOC_Insert_Point"
    )
    $TOC = @()
    foreach ($heading in $InputObject){
        If ($heading.level -le 4){
            Switch ($heading.level){
                1{$Style = "Margin-left:10;Font-Size:16pt"}
                2{$Style = "Margin-left:30"}
                3{$Style = "Margin-left:50"}
                4{$Style = "Margin-left:70"}
            }
            $TOC += "<DIV style=`"$Style`"><a href=`"`#$($heading.Id)`" style=`"color:blue`">$($heading.Title)</a></DIV>"
        }
    }
    (Get-Content $File) | 
        Foreach-Object {
            $_ # send the current line to output
            if ($_ -match $InsertPoint) 
            {
                #Add Lines after the selected pattern 
                $TOC
            }
        } | Set-Content $File
}

function Write-HtmliLink{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,ParameterSetName='Standard',
        HelpMessage="This is the text that will appear in title tag of the header")]
        [string]$LinkID,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,ParameterSetName='Standard',
        HelpMessage="This is the amount of space that the table will indent by")]
        [string]$Text,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='TOC',
        HelpMessage="This is the text that will appear in the lower left by line")]
        [Switch]$ReturnTOC,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='TOC',
        HelpMessage="This is the text that will appear in the lower left by line")]
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,ParameterSetName='Standard',
        HelpMessage="This is the text that will appear in the lower left by line")]
        [String]$File
    )
    if($ReturnTOC){
        $iLink = "<div style=`"text-align:right`"><a href=`"#TableofContents`" style=`"color:DarkRed`">Return to Table of Contents</a></div>"
    }Else{
        $iLink = "<div><a href=`"#$LinkID`" style=`"color:blue`">$Text</a></div>"
    }
    If ($File) {$iLink | Out-File -filepath $File -Append}
    Else {$iLink}
}


function Write-HtmlComment{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear in the HTML comment")]
        [string]$Text,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear in the lower left by line")]
        [String]$File
    )
    "<!--$Text-->" | Out-File -filepath $File -Append
}



#$folders=Get-ChildItem -Path 'C:\AMD\WU-CCC2\ccc2_install' | Where {$_.PSIsContainer -eq $false} | select Name,FullName,Mode,Length
#Write-HTMLHeadLine -Text "Applications" -Level 2 -PageBreak
#Write-HTMLParagraph -Text "Please read this text. It is good to share with all." -Indent 20
#Write-HtmlTable -InputObject $folders -Border 1 -Indent 25
#Write-HTMLHeadLine -Text "Programs" -Level 2
#Write-HTMLParagraph -Text "Please read this text. It is good to share with all. Don't indent." -Indent 20
#Write-HtmlTable -InputObject $folders -Border 1 -Indent 25

##########################################################################################################################
###############################################HTML Format functions above################################################
##########################################################################################################################

function Ping-Host { 
  Param([string]$computername=$(Throw "You must specify a computername.")) 
  Write-Debug "In Ping-Host function" 
  $query="Select * from Win32_PingStatus where address='$computername'" 
  $wmi=Get-WmiObject -query $query 
  if([string]::IsNullOrEmpty($wmi.ResponseTime)){$false}Else{$true}
}


function Invoke-SqlDataReader {
 
<#
.SYNOPSIS
    Runs a select statement query against a SQL Server database.
 
.DESCRIPTION
    Invoke-SqlDataReader is a PowerShell function that is designed to query
    a SQL Server database using a select statement without the need for the SQL
    PowerShell module or snap-in being installed.
 
.PARAMETER ServerInstance
    The name of an instance of the SQL Server database engine. For default instances,
    only specify the server name: 'ServerName'. For named instances, use the format
    'ServerName\InstanceName'.
 
.PARAMETER Database
    The name of the database to query on the specified SQL Server instance.
 
.PARAMETER Query
    Specifies one Transact-SQL select statement query to be run.
 
.PARAMETER Credential
    SQL Authentication userid and password in the form of a credential object.
 
.EXAMPLE
     Invoke-SqlDataReader -ServerInstance Server01 -Database Master -Query '
     select name, database_id, compatibility_level, recovery_model_desc from sys.databases'
 
.EXAMPLE
     'select name, database_id, compatibility_level, recovery_model_desc from sys.databases' |
     Invoke-SqlDataReader -ServerInstance Server01 -Database Master
 
.EXAMPLE
     'select name, database_id, compatibility_level, recovery_model_desc from sys.databases' |
     Invoke-SqlDataReader -ServerInstance Server01 -Database Master -Credential (Get-Credential)
 
.INPUTS
    String
 
.OUTPUTS
    DataRow
 
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>
 
    [CmdletBinding()]
    param (        
        [Parameter(Mandatory)]
        [string]$ServerInstance,
 
        [Parameter(Mandatory)]
        [string]$Database,
        
        [Parameter(Mandatory,
                   ValueFromPipeline)]
        [string]$Query,
        
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    BEGIN {
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
 
        if (-not($PSBoundParameters.Credential)) {
            $connectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=True;"
        }
        else {
            $connectionString = "Server=$ServerInstance;Database=$Database;Integrated Security=False;"
            $userid= $Credential.UserName -replace '^.*\\|@.*$'
            ($password = $credential.Password).MakeReadOnly()
            $sqlCred = New-Object -TypeName System.Data.SqlClient.SqlCredential($userid, $password)
            $connection.Credential = $sqlCred
        }
 
        $connection.ConnectionString = $connectionString
        $ErrorActionPreference = 'Stop'
        
        try {
            $connection.Open()
            Write-Verbose -Message "Connection to the $($connection.Database) database on $($connection.DataSource) has been successfully opened."
        }
        catch {
            Write-Error -Message "An error has occurred. Error details: $($_.Exception.Message)"
        }
        
        $ErrorActionPreference = 'Continue'
        $command = $connection.CreateCommand()
    }
 
    PROCESS {
        $command.CommandText = $Query
        $ErrorActionPreference = 'Stop'
 
        try {
            $result = $command.ExecuteReader()
        }
        catch {
            Write-Error -Message "An error has occured. Error Details: $($_.Exception.Message)"
        }
 
        $ErrorActionPreference = 'Continue'
 
        if ($result) {
            $dataTable = New-Object -TypeName System.Data.DataTable
            $dataTable.Load($result)
            $dataTable
        }
    }
 
    END {
        $connection.Close()
    }
 
}


Function Read-ScheduleToken {
  
  $SMS_ScheduleMethods = 'SMS_ScheduleMethods'
  $class_SMS_ScheduleMethods = [wmiclass]''
  $class_SMS_ScheduleMethods.psbase.Path ="ROOT\SMS\Site_$($SiteCode):$($SMS_ScheduleMethods)"
  
  $script:ScheduleString = $class_SMS_ScheduleMethods.ReadFromString($ServiceWindow.ServiceWindowSchedules)
  return $ScheduleString
}

Function Convert-WeekDay {
  [CmdletBinding()]
  param (
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [string]$Day
  )
  ### day of week
  switch ($Day)
  {
    1 {$weekday = 'Sunday'}
    2 {$weekday = 'Monday'}
    3 {$weekday = 'Tuesday'}
    4 {$weekday = 'Wednesday'}
    5 {$weekday = 'Thursday'}
    6 {$weekday = 'Friday'}
    7 {$weekday = 'Saturday'}
  }
  return $weekday
}

Function Convert-Time {
  param (
    [int]$time
  )
  $min = $time % 60
  if ($min -le 9) {$min = "0$($min)" }
  $hrs = [Math]::Truncate($time/60)
  
  $NewTime = "$($hrs):$($min)"
  return $NewTime
}

Function Get-SiteCode
{
  $wqlQuery = 'SELECT * FROM SMS_ProviderLocation'
  $a = Get-WmiObject -Query $wqlQuery -Namespace 'root\sms' -ComputerName $SMSProvider
  $a | ForEach-Object {
    if($_.ProviderForLocalSite)
    {
      $script:SiteCode = $_.SiteCode
    }
  }
  return $SiteCode
}

##Recursively processes through all the steps in a task sequence
Function Process-TSSteps{
    param ($Sequence,$GroupName)
    foreach ($node in $Sequence.ChildNodes){
        switch($node.localname) {
            'step'{
                if (-not [string]::IsNullOrEmpty($node.Description)){
                    $StepDescription = "$($node.Description)"
                }
                try {
                        if (-not [string]::IsNullOrEmpty($node.disable)){
                            $StepStatus = 'Disabled'
                        }else{
                            $StepStatus = 'Enabled'
                        }
                    }   
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $StepStatus = 'Enabled'
                }
                if($GroupName){
                    #"$($GroupName):  $($node.name) $($node.action)"
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="$GroupName";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="$($node.Action)";'Status'="$StepStatus"}
                }else{
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="N/A";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="$($node.Action)";'Status'="$StepStatus"}
                    #"$($node.name) $($node.action)"
                }
                $TSStep
            }
            'group'{
                $TSStepNumber++
                if (-not [string]::IsNullOrEmpty($node.Description)){
                    $StepDescription = "$($node.Description)"
                }
                try {
                        if (-not [string]::IsNullOrEmpty($node.disable)){
                            $StepStatus = 'Disabled'
                        }else{
                            $StepStatus = 'Enabled'
                        }
                    }   
                catch [System.Management.Automation.PropertyNotFoundException] {
                    $StepStatus = 'Enabled'
                }
                #"Group: $($node.Name)"
                $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="$($node.Name)";'Step Name'="N/A";'Description'="$StepDescription";'Action'="N/A";'Status'="$StepStatus"}
                $TSStep
                Process-TSSteps -Sequence $node -GroupName "$($node.Name)" -TSSteps $TSSteps -StepCounter $TSStepNumber
            }
            default{}
        }
    }
}

####################################################################################################################################################################
####################################################################################################################################################################
#####################################################################Starting#######################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
$StartingPath = (get-location).Path

$SiteCode = Get-SiteCode

Write-Verbose "$(Get-Date): Start writing report data"

$LocationBeforeExecution = Get-Location

Write-HTMLHeader -Title $Title -File $FilePath
Write-HTMLCoverPage -Title $Title -Author $Author -Vendor $Vendor -Org $CompanyName -ImagePath $CompanyLogo -File $FilePath



#Import the CM Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
{
  Write-Verbose "$(Get-Date):   CM PowerShell module has not been imported yet, will import it now."
  Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1') | Out-Null
}
#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
{
  Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
  exit 1
}

#### Administration
#### Site Configuration

Write-HTMLHeading -Text 'Table of Contents' -Level 1 -PageBreak -ExcludeTOC -File $FilePath
Write-HtmlComment -Text "TOC_Insert_Point" -File $FilePath
Write-HTMLHeading -Text 'Summary of all Sites in this Hierarchy' -Level 1 -PageBreak -File $FilePath
Write-Verbose "$(Get-Date):   Getting Site Information"
$CMSites = Get-CMSite

$CAS                    = $CMSites | Where-Object {$_.Type -eq 4}
$ChildPrimarySites      = $CMSites | Where-Object {$_.Type -eq 3}
$StandAlonePrimarySite  = $CMSites | Where-Object {$_.Type -eq 2}
$SecondarySites         = $CMSites | Where-Object {$_.Type -eq 1}

#region CAS
if (-not [string]::IsNullOrEmpty($CAS))
{
  Write-HTMLParagraph -Text 'The following Central Administration Site is installed:' -level 1 -File $FilePath
  $CAS = New-Object -TypeName psobject -Property @{'Site Name' = $CAS.SiteName; 'Site Code' = $CAS.SiteCode; Version = $CAS.Version };
  
  Write-HtmlTable -InputObject $CAS -Border 1 -Level 1 -File $FilePath
}
else {
  Write-HTMLParagraph -Text 'No <b>CAS</b> detected. continue with Primary Sites.' -level 1 -File $FilePath
}
#endregion CAS

#region Child Primary Sites
if (-not [string]::IsNullOrEmpty($ChildPrimarySites))
{
  Write-Verbose "$(Get-Date):   Enumerating all child Primary Site."
  Write-HTMLParagraph -Text 'The following child Primary Sites are installed:' -level 1 -File $FilePath
  $ChildSite = New-Object -TypeName psobject -Property @{'Site Name' = $ChildPrimarySites.SiteName; 'Site Code' = $ChildPrimarySites.SiteCode; Version = $ChildPrimarySites.Version };
  
  Write-HtmlTable -InputObject $ChildSite -Border 1 -Level 1 -File $FilePath
}
#endregion Child Primary Sites


#region Standalone Primary
if (-not [string]::IsNullOrEmpty($StandAlonePrimarySite))
{
  Write-Verbose "$(Get-Date):   Enumerating a standalone Primary Site."
  Write-HTMLParagraph -Text 'The following Primary Site is installed:' -level 1 -File $FilePath
  $CMSiteID = Get-WmiObject -Class SMS_Identification -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider| where {$_.ThisSiteCode -eq "$SiteCode"}
  $ReleaseVersion = $CMSiteID.MonthlyReleaseVersion
  $StandAlonePrimarySite = New-Object -TypeName psobject -Property @{'Site Name' = $StandAlonePrimarySite.SiteName; 'Site Code' = $StandAlonePrimarySite.SiteCode; Version = $StandAlonePrimarySite.Version; 'Build' = $StandAlonePrimarySite.BuildNumber; 'Release Version' = $ReleaseVersion};
  
  $StandAlonePrimarySite = $StandAlonePrimarySite |select 'Site Name','Site Code','Release Version',Version,Build
  Write-HtmlTable -InputObject $StandAlonePrimarySite -Border 1 -Level 1 -File $FilePath
}
#endregion Standalone Primary

#region Secondary Sites
if (-not [string]::IsNullOrEmpty($SecondarySites))
{
  Write-Verbose "$(Get-Date):   Enumerating all secondary sites."
  Write-HTMLParagraph -Text 'The following Secondary Sites are installed:' -level 1 -File $FilePath
  $SecondarySites = New-Object -TypeName psobject -Property @{'Site Name' = $SecondarySites.SiteName; 'Site Code' = $SecondarySites.SiteCode; Version = $SecondarySites.Version };
  
  Write-HtmlTable -InputObject $SecondarySites -Border 1 -Level 1 -File $FilePath
}
#endregion Secondary Sites


#region Site Configuration report
foreach ($CMSite in $CMSites)
{  
  Write-Verbose "$(Get-Date):   Checking each site's configuration."
  Write-HTMLHeading -Text "Configuration Summary for Site $($CMSite.SiteCode)" -level 1 -File $FilePath
  Write-HTMLHeading -Text "Updates and Servicing" -Level 2 -File $FilePath

  #region Site Servicing Updates
  Write-Verbose "$(Get-Date):   Enumerating Configuration Manager Update Status and History"
  Write-HTMLHeading -Text "Update Status and History" -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "Below is a history of updates that have been made available to this Site.  It includes information for if, or when, they were installed.  Some older updates may be listed as ready to install, however, they were never installed nor will they be avialable to install as they are superseded by the newer updates." -Level 3 -File $FilePath
  $SiteUpdateHistory = Get-CMSiteUpdateHistory| select Name,FullVersion,Impact,State,UpdateType,LastUpdateTime|sort LastUpdateTime
  if(-not [string]::IsNullOrEmpty($SiteUpdateHistory)){
    $SiteUpdates = @()
    foreach ($SiteUpdate in $SiteUpdateHistory){
        Switch($SiteUpdate.State){
            196612{$UpdateState = "Installed"}
            262146{$UpdateState = "Ready to Install"}
            default{$UpdateState = "Other ($($SiteUpdate.State))"}
        }
        If($UpdateState -eq "Installed"){
            $InstalledDate = $SiteUpdate.LastUpdateTime
        }else{
            $InstalledDate = "N/A"
        }
        $SiteUpdates += New-Object -TypeName PSObject -Property @{'Name'="$($SiteUpdate.Name)";'Version'="$($SiteUpdate.FullVersion)";'Status'="$UpdateState";'Installed Date'="$InstalledDate"}
    }
    $SiteUpdates = $SiteUpdates | select 'Name','Version','Status','Installed Date'
    Write-HtmlTable -InputObject $SiteUpdates -Border 1 -Level 3 -File $FilePath
  }
  Write-Verbose "$(Get-Date):   Completed Configuration Manager Update Status and History"
  #endregion Site Servicing Updates

  #region Site Features
  Write-Verbose "$(Get-Date):   Enumerating Configuration Manager Site Features"
  Write-HTMLHeading -Text "Site Features" -Level 3 -File $FilePath
  $features=Get-CMSiteFeature
  #region release features
  $ReleaseFeatures = $features | Where{$_.FeatureType -eq 1}|Sort-Object Name
  $FeatureTable = @()
  Foreach ($feature in $ReleaseFeatures){
    $FeatureName = $feature.Name
    Switch($feature.Status){
        1{$FeatureStatus = "On"}
        0{$FeatureStatus = "Off"}
        default{$FeatureStatus = "Unknown"}
    }
    $FeatureTable += New-Object -TypeName PSObject -Property @{'Feature Name'="$FeatureName";'Status'="$FeatureStatus"}
  }
  Write-HTMLHeading -Text "Release Features" -Level 4 -File $FilePath
  Write-HTMLParagraph -Text "Below is a list of all released features in this Configuration Manager site and which ones are enabled and which are not.  Once a feature is turned on, it cannot be turned off." -Level 4 -File $FilePath
  Write-HtmlTable -InputObject $FeatureTable -Border 1 -Level 4 -File $FilePath
  #endregion release features

  #region PreRelease features
  $PreReleaseFeatures = $features | Where{$_.FeatureType -eq 0}|Sort-Object Name
  $PreFeatureTable = @()
  Foreach ($feature in $PreReleaseFeatures){
    $FeatureName = $feature.Name
    Switch($feature.Status){
        1{$FeatureStatus = "On"}
        0{$FeatureStatus = "Off"}
        default{$FeatureStatus = "Unknown"}
    }
    $PreFeatureTable += New-Object -TypeName PSObject -Property @{'Feature Name'="$FeatureName";'Status'="$FeatureStatus"}
  }
  Write-HTMLHeading -Text "Pre-Release Features" -Level 4 -File $FilePath
  Write-HTMLParagraph -Text "Below is a list of all pre-release features in this Configuration Manager site and which ones are enabled and which are not.  Once a feature is turned on, it cannot be turned off." -Level 4 -File $FilePath
  Write-HtmlTable -InputObject $PreFeatureTable -Border 1 -Level 4 -File $FilePath
  #region PreRelease features
  Write-HtmliLink -ReturnTOC -File $FilePath
  Write-Verbose "$(Get-Date):   Completed Configuration Manager Site Features"
  #endregion Site Features

  $SiteRolesTable = @()  
  $SiteRoles = Get-CMSiteRole -SiteCode $CMSite.SiteCode | Select-Object -Property NALPath, rolename

  Write-HTMLHeading -Text "Site Roles" -Level 2 -File $FilePath
  Write-HTMLParagraph  -Text "The following Site Roles are installed in this site:" -Level 2 -File $FilePath
  foreach ($SiteRole in $SiteRoles) {
    if (-not (($SiteRole.rolename -eq 'SMS Component Server') -or ($SiteRole.rolename -eq 'SMS Site System'))) {
        $SiteRoleobject = New-Object -TypeName PSObject -Property @{'Server Name' = ($SiteRole.NALPath).ToString().Split('\\')[2]; 'Role' = $SiteRole.RoleName}
        $SiteRolesTable += $SiteRoleobject
    }
  }
  Write-HtmlTable -InputObject $SiteRolesTable -Border 1 -Level 2 -File $FilePath
  Write-HtmliLink -ReturnTOC -File $FilePath

  $SiteMaintenanceTaskTable = @()
  $SiteMaintenanceTasks = Get-CMSiteMaintenanceTask -SiteCode $CMSite.SiteCode
  Write-HTMLHeading -Text "Site Maintenance Tasks for Site $($CMSite.SiteCode)" -Level 2 -File $FilePath
  
  foreach ($SiteMaintenanceTask in $SiteMaintenanceTasks) {
    $SiteMaintenanceTaskRowHash = New-Object -TypeName PSObject -Property @{'Task Name' = $SiteMaintenanceTask.TaskName; Enabled = $SiteMaintenanceTask.Enabled};
    $SiteMaintenanceTaskTable += $SiteMaintenanceTaskRowHash;
  }

  $SiteMaintenanceTaskTable = $SiteMaintenanceTaskTable|Select 'Task Name',Enabled
  Write-HtmlTable -InputObject $SiteMaintenanceTaskTable -Border 1 -Level 2 -File $FilePath
  Write-HtmliLink -ReturnTOC -File $FilePath
  
  #region Site SQL Info
  Write-HTMLHeading -Text "Summary of SQL database info for Site $($CMSite.SiteCode)" -PageBreak -Level 2 -File $FilePath
  $SiteDef = Get-CMSiteDefinition -SiteCode $($CMSite.SiteCode)
  $SQLServer = $SiteDef.SQLServerName
  $CMDatabase = $SiteDef.SQLDatabaseName
  $SQLInfo = @("Site SQL Server: <b>$SQLServer</b>","Site Database Name: <b>$CMDatabase</b>")
  Write-HtmlList -InputObject $SQLInfo -Level 2 -File $FilePath
  #Write-HTMLParagraph -Text "$($SQLInfo)" -Level 2 -File $FilePath
  #Query SQL Server WMI for basic hardware Information: CPU,RAM,Drives
  $SQLHWDesc = "$SQLServer Hardware Info:"
  $SQLHWInfo = @()
  try {
    $Capacity = 0
    Get-WmiObject -Class win32_physicalmemory -ComputerName $SQLServer | ForEach-Object {[int64]$Capacity = $Capacity + [int64]$_.Capacity}
    $TotalMemory = $Capacity / 1024 / 1024 / 1024
    $CPUs = Get-WmiObject -Class win32_processor -ComputerName $SQLServer 
    [int]$Cores=0
    foreach ($CPU in $CPUs) {
        $Cores = $Cores + $CPU.NumberOfCores
        $CPUModel = $CPU.Name
    }
    [int]$Threads=0
    foreach ($CPU in $CPUs) {$Threads = $Threads + $CPU.NumberOfLogicalProcessors}
    $SQLHWInfo += "$CPUModel"
    $SQLHWInfo += "$Cores Cores ($Threads logical)"
    $SQLHWInfo += "$($TotalMemory) GB RAM"
    $Drives=Get-WmiObject -Class win32_LogicalDisk -ComputerName $SQLServer | Where {$_.DriveType -eq 3}
    Foreach($Drive in $Drives){
        $SQLHWInfo += "Drive $($Drive.DeviceID) size: $([math]::Round($Drive.size/1024/1024/1024,1)) GB ($([math]::Round($Drive.FreeSpace/1024/1024/1024,1)) GB Free)"
    }
  }
  catch {
    $SQLHWInfo += "Failed to access server: $SQLServer" 
  }
  Write-HtmlList -InputObject $SQLHWInfo -Description $SQLHWDesc -Level 2 -File $FilePath
  
  Write-Verbose "$(Get-Date):   Getting SQL Database detailed info."
  $SQLVersion = Invoke-SqlDataReader -ServerInstance $SQLServer -Database Master -Query "SELECT SERVERPROPERTY (`'edition`') Edition, SERVERPROPERTY(`'productversion`') Version, SERVERPROPERTY (`'productlevel`') SP, SERVERPROPERTY (`'ProductUpdateLevel`') CU"
  $SQLConfig = Invoke-SqlDataReader -ServerInstance $SQLServer -Database Master -Query "SELECT name ServerSetting,value_in_use Value FROM sys.configurations where configuration_id = 1543 OR configuration_id = 1544 OR configuration_id = 1539"
  $DatabaseFiles = Invoke-SqlDataReader -ServerInstance $SQLServer -Database Master -Query "SELECT db.name `'DatabaseName`',type_desc `'FileType`',physical_name `'FilePath`',mf.state_desc `'Status`',size*8/1024 `'FileSizeMB`',max_size `'MaximumSize`',growth `'GrowthRate`',(CASE WHEN is_percent_growth = 1 THEN `'Percent`' ELSE `'MB`' END) `'GrowthUnit`',create_date `'DateCreated`',compatibility_level `'DBLevel`',user_access_desc `'AccessMode`',recovery_model_desc `'RecoveryModel`' FROM sys.master_files mf INNER JOIN sys.databases db ON db.database_id = mf.database_id where db.name = `'$CMDatabase`'"
  $SQLVersion = $SQLVersion | Select Edition,Version,SP,CU
  $SQLConfig = $SQLConfig | Select @{Name='Server Setting';Expression={$_.ServerSetting}},Value
  $DatabaseFiles = $DatabaseFiles | Select @{Name='File Type';Expression={$_.FileType}},@{Name='File Path';Expression={$_.FilePath}},Status,@{Name='File Size MB';Expression={'{0:N0}' -f $_.FileSizeMB}},@{Name='Maximum Size';Expression={$(IF($_.MaximumSize -eq -1){"Unlimited"}else{'{0:N0}' -f ($_.MaximumSize/128)})}},@{Name='Growth Rate';Expression={"$(IF($_.GrowthUnit -eq "Percent"){"$($_.GrowthRate)%"}Else{"$($_.GrowthRate/128)MB"})"}},@{Name='Recovery Model';Expression={$_.RecoveryModel}}
  $IndexFragmentation = Invoke-SqlDataReader -ServerInstance $SQLServer -Database $CMDatabase -Query "SELECT SUM(CASE WHEN indexstats.avg_fragmentation_in_percent > 75 THEN  1 ELSE 0 END) [Over 75],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 50 AND indexstats.avg_fragmentation_in_percent <= 75) THEN  1 ELSE 0 END) [Over 50],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 25 AND indexstats.avg_fragmentation_in_percent <= 50) THEN  1 ELSE 0 END) [Over 25],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 1 AND indexstats.avg_fragmentation_in_percent <= 25) THEN  1 ELSE 0 END) [Under 25],SUM(CASE WHEN indexstats.avg_fragmentation_in_percent < 1 THEN  1 ELSE 0 END) [Not Fragmented] FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, NULL) AS indexstats JOIN sys.tables dbtables on dbtables.[object_id] = indexstats.[object_id] WHERE indexstats.database_id = DB_ID()"
  $IndexFragmentation = $IndexFragmentation | Select 'Over 75','Over 50','Over 25','Under 25','Not Fragmented'
  Write-HTMLParagraph -Text "SQL instance version information:" -Level 2 -File $FilePath
  Write-HtmlTable -InputObject $SQLVersion -Border 1 -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "The following are important global settings on the SQL server.  Typically, this SQL server should be dedicated to Configuration Manager." -Level 2 -File $FilePath
  Write-HtmlTable -InputObject $SQLConfig -Border 1 -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "Below are the database files for the site database ($CMDatabase):" -Level 2 -File $FilePath
  Write-HtmlTable -InputObject $DatabaseFiles -Border 1 -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "Below is a fragmentation summary (%) for indexes on the site database ($CMDatabase):" -Level 2 -File $FilePath
  Write-HtmlTable -InputObject $IndexFragmentation -Border 1 -Level 3 -File $FilePath
  Write-Verbose "$(Get-Date):   SQL detailed info complete."
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion Getting Site SQL Info

  #region Management Points
  $CMManagementPoints = Get-CMManagementPoint -SiteCode $CMSite.SiteCode
  Write-HTMLHeading -Text "Summary of Management Points for Site $($CMSite.SiteCode)" -PageBreak -Level 2 -File $FilePath
  foreach ($CMManagementPoint in $CMManagementPoints)
  {
    $MPText = @()
    #Write-Verbose "$(Get-Date):   Management Point: $($CMManagementPoint)"
    $MPName = $CMManagementPoint.NetworkOSPath.Split('\\')[2]
    Write-Verbose "$(Get-Date):   Management Point Name: $MPName"
    [bool]$SSLENabled = if($CMManagementPoint.SslState -eq 0){$false}else{$true}
    $MPText += "SSL Enabled: $SSLENabled"
    $UseSiteDB = ($CMManagementPoint.props|Where{$_.PropertyName -like "UseSiteDatabase"}).value
    [bool]$UseSiteDB = if($UseSiteDB -eq 1) {$true}else{$false}
    $MPText += "Using Site Database: $UseSiteDB"
    $MPIntranet = ($CMManagementPoint.props|Where{$_.PropertyName -like "MPIntranetFacing"}).value
    $MPInternet = ($CMManagementPoint.props|Where{$_.PropertyName -like "MPInternetFacing"}).value
    Write-Verbose "$(Get-Date): Internet: $MPInternet Intranet: $MPIntranet"
    If (!($MPIntranet) -and !($MPInternet)) {[bool]$MPIntranet = $true; [bool]$MPInternet = $false}
    Else {
        [bool]$MPIntranet = If($MPIntranet -eq 1){$true}else{$false}
        [bool]$MPInternet = If($MPInternet -eq 1){$true}else{$false}
    }
    $MPText += "Intranet Clients: $MPIntranet"
    $MPText += "Internet Clients: $MPInternet"
    Write-HtmlList -InputObject $MPText -Title "Management Point Name: <B>$MPName</B>" -Level 2 -File $FilePath
    Remove-Variable MPIntranet
    Remove-Variable MPInternet
    Write-Verbose "$(Get-Date):   Test-Path -Path `"filesystem::\\$MPName\C$`""
    $local1 = (Get-Location).path
    Set-Location C:
    $PathTest = Test-Path -Path "filesystem::\\$MPName\C$"
    Write-Verbose "$(Get-Date):   Testing Access to Management Point: $MPName -- $PathTest"
    If (Test-Path -Path "filesystem::\\$MPName\C$") {$CMMPServerName=$MPName}
    Set-Location $local1
  }
  Write-HtmliLink -ReturnTOC -File $FilePath
  Write-Verbose "$(Get-Date):   Default Management Point: $CMMPServerName"
  #endregion Management Points
  
  #region Distribution Point details
  Write-HTMLHeading -Text "Summary of Distribution Points for Site $($CMSite.SiteCode)" -Level 2 -PageBreak -File $FilePath
  $CMDistributionPoints = Get-CMDistributionPoint -SiteCode $CMSite.SiteCode
  
  foreach ($CMDistributionPoint in $CMDistributionPoints)
  {
    $CMDPServerName = $CMDistributionPoint.NetworkOSPath.Split('\\')[2]
    Write-Verbose "$(Get-Date):   Found DP: $($CMDPServerName)"
    Write-HTMLHeading -Text "$CMDPServerName" -Level 3 -File $FilePath
    Write-Verbose "Trying to ping $($CMDPServerName)"
    $PingResult = Ping-Host $CMDPServerName
    if (-not ($PingResult))
    {
      Write-Verbose "Ping Failed: $($CMDPServerName)"
      Write-HTMLParagraph -Text "The Distribution Point $($CMDPServerName) is not reachable. Check connectivity." -Level 3 -File $FilePath
    }
    else
    {
      Write-Verbose "Ping Succeeded: $($CMDPServerName)"
      Write-HTMLParagraph -Text "Disk Information:" -Level 4 -File $FilePath
      $CMDPDrives = (Get-WmiObject -Class SMS_DistributionPointDriveInfo -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider).Where({$PSItem.NALPath -like "*$CMDPServerName*"})
      foreach ($CMDPDrive in $CMDPDrives)
      {
        $DPText = "<b>Partition $($CMDPDrive.Drive):</b>"
        $Size = ''
        $Size = $CMDPDrive.BytesTotal / 1024 / 1024
        $Freesize = ''
        $Freesize = $CMDPDrive.BytesFree / 1024 / 1024
        
        $HTMLList = "<UL><LI>$([MATH]::Round($Size,2)) GB Size in total</LI><LI>$([MATH]::Round($Freesize,2)) GB Free Space</LI><LI>Still $($CMDPDrive.PercentFree) percent free.</LI></UL>"
        $DPText = $DPText + $HTMLList
        Write-HTMLParagraph -Text $DPText -Level 4 -File $FilePath
      }
      
      $DPText = "<B>Hardware Info:</B>"
      try {
          $Capacity = 0
          Get-WmiObject -Class win32_physicalmemory -ComputerName $CMDPServerName | ForEach-Object {[int64]$Capacity = $Capacity + [int64]$_.Capacity}
          $TotalMemory = $Capacity / 1024 / 1024 / 1024
          $CPUs = Get-WmiObject -Class win32_processor -ComputerName $CMDPServerName 
          [int]$Cores=0
          foreach ($CPU in $CPUs) {$Cores = $Cores + $CPU.NumberOfCores}
          $CPUModel = $CPU.Name
          $DPText = $DPText + "<BR /><UL><LI>$CPUModel</LI><LI>$Cores Cores</LI><LI>$($TotalMemory) GB RAM</LI></UL>"
      }
      catch {
        $DPText = $DPText + "<BR />Failed to access server $CMDPServerName.<BR /><BR />" 
        }
    }
    Write-HTMLParagraph -Text "$DPText" -Level 4 -File $FilePath
    $DPText = "<B>Additional Configuration:</B><ul>"
    $DPInfo = $CMDistributionPoint.Props
    $IsPXE = ($DPInfo.Where({$_.PropertyName -eq 'IsPXE'})).Value
    $UnknownMachines = ($DPInfo.Where({$_.PropertyName -eq 'SupportUnknownMachines'})).Value
    switch ($IsPXE)
    {
      1 
      {
        $DPText = $DPText + "<li>PXE Enabled</li>"
        switch ($UnknownMachines)
        {
          1 { $DPText = $DPText + "<li>Supports unknown machines: true</li>" }
          0 { $DPText = $DPText + "<li>Supports unknown machines: false</li>" }
        }
      }
      0
      {
        $DPText = $DPText + "<li>PXE Disabled</li>"
      }
    }
    $DPText = $DPText + "</ul>"
    Write-HTMLParagraph -Text $DPText -Level 4 -File $FilePath
    $DPGroupMembers = $Null
    $DPGroupIDs = $Null
    $DPGroupMembers = (Get-WmiObject -class SMS_DPGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider) | Where {$_.DPNALPath -ilike "*$($CMDPServerName)*"}
    if (-not [string]::IsNullOrEmpty($DPGroupMembers))
    {
      $DPGroupIDs = $DPGroupMembers.GroupID
    }
    
    #enumerating DP Group Membership
    $DPText = "<B>Distribution Point Group Membership:</B>"
    if (-not [string]::IsNullOrEmpty($DPGroupIDs))
    {
      $GroupList = "<UL>"
      foreach ($DPGroupID in $DPGroupIDs)
      {
        $DPGroupName = (Get-CMDistributionPointGroup -Id "$($DPGroupID)").Name
        $GroupList = $GroupList + "<LI>$DPGroupName</LI>"
      }
      $DPText = $DPText + $GroupList + "</UL>"
    }
    else
    {
      $DPText = $DPText + "<ul><li>This Distribution Point is not a member of any DP Group.</li></ul>"
    }
    Write-HTMLParagraph -Text $DPText -Level 4 -File $FilePath
  }
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion Distribution Point details

  #region enumerating Software Update Points and Configuration
  Write-HTMLHeading -Text "Software Update configuration for Site $($CMSite.SiteCode)" -Level 2 -PageBreak -File $FilePath
  
  Write-HTMLHeading -Text "Software Update Point Component Settings for Site $($CMSite.SiteCode)" -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "This is a list of all of the software update classifications and products that are syncronized into the site as well as some of the general site configuration settings." -Level 3 -File $FilePath

  $cats=Get-CMSoftwareUpdateCategory
  $UpdatesClassifications = $cats|where {$_.CategoryTypeName -eq "UpdateClassification" -and $_.AllowSubscription -eq $true}
  $SubscribedUpdatesClassifications = $UpdatesClassifications|where {$_.IsSubscribed -eq $true}
  $Products = $cats|where {$_.CategoryTypeName -eq "Product" -and $_.AllowSubscription -eq $true}
  $SubscribedProducts = $Products|where {$_.IsSubscribed -eq $true}
  $SupProperties = (Get-CMSoftwareUpdatePointComponent).props

  $SUPPropertyList = @()
  $SUPPropertyList += "Synchronizing $($SubscribedUpdatesClassifications.Count) of $($UpdatesClassifications.Count) update classifications."
  $SUPPropertyList += "Synchronizing $($SubscribedProducts.count) of $($Products.count) products."
  Foreach ($SupProp in $SupProperties){
    Switch ($SupProp.PropertyName){
        'Call WSUS Cleanup'{
            if ($SupProp.value -eq 1){
                $SUPPropertyList += "Run WSUS cleanup wizard: Enabled"
            }Elseif ($SupProp.value -eq 0){
                $SUPPropertyList += "Run WSUS cleanup wizard: Disabled"
            }
        }
        'Sync Supersedence Age'{
            $SUPPropertyList += "months to wait before a superseded software update is expired: $($SupProp.value)"
        }
        'Sync Supersedence Mode'{
            switch ($SupProp.value){
                1{$SUPPropertyList += "Do not expire a superseded software update until the software update is superseded for a specified period"}
                0{$SUPPropertyList += "Immediately expire a superseded software update (ignore `'months to wait before a superseded software update is expired`')"}
            }
        }
        'SupportedUpdateLanguages'{
            $SUPPropertyList += "Software Update File languages: $($SupProp.Value2)"
        }
        'SupportedTitleLanguages'{
            $SUPPropertyList += "Update Summary Details languages: $($SupProp.Value2)"
        }
    }
  }

  
  Write-HTMLHeading -Text "Software Update Point Base Settings" -Level 4 -File $FilePath
  Write-HtmlList -InputObject $SUPPropertyList -Level 4 -File $FilePath
  Write-HTMLHeading -Text "Selected Software Update Classifications" -Level 4 -File $FilePath
  Write-HtmlList -InputObject ($SubscribedUpdatesClassifications.LocalizedCategoryInstanceName) -Level 4 -File $FilePath
  Write-HTMLHeading -Text "Selected Software Update Point Software Products" -Level 4 -File $FilePath
  Write-HtmlList -InputObject ($SubscribedProducts.LocalizedCategoryInstanceName) -Level 4 -File $FilePath


  Write-Verbose "$(Get-Date):   Enumerating all Software Update Points"
  Write-HTMLHeading -Text "Software Update Point Servers for Site $($CMSite.SiteCode)" -Level 3 -File $FilePath
  Write-Verbose "Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $CMMPServerName | Where-Object {$_.rolename -eq `'SMS Software Update Point`'}"
  $CMSUPs = Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $CMMPServerName | Where-Object {$_.rolename -eq 'SMS Software Update Point'}
  #$CMSUPs = (Get-CMSoftwareUpdatePoint).Where({$_.SiteCode -eq "$($CMSite.SiteCode)"})
  if (-not [string]::IsNullOrEmpty($CMSUPs))
  {
    foreach ($CMSUP in $CMSUPs) {
      $SUPPropertyTable = @();
      $CMSUPServerName = $CMSUP.NetworkOSPath.split('\\')[2]
      Write-Verbose "$(Get-Date):   Found SUP: $($CMSUPServerName)"
      Write-HTMLHeading -Text "$($CMSUPServerName)" -Level 4 -File $FilePath
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'WSUS IIS Port'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'WSUSIISPORT'}).value)}
      #8530
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Database'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'DBServerName'}).value2)}
      #soup-cm1.soup.steamedsoup.com\MICROSOFT##WID
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Access Account'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'WSUSAccessAccount'}).value2)}
      #soup.steamedsoup.com\SVC-SCCM-RAA
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SSL Enabled'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'SSLWSUS'}).value)}
      #0
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SSL Port'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'WSUSIISSSLPORT'}).value)}
      #8531
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SUP Enabled'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'Enabled'}).value)}
      #1
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Proxy Enabled'; Value = (($CMSUP.props|select Propertyname,Value,Value1,Value2| where {$_.PropertyName -like 'UseProxy'}).value)}
      #0
      $SUPPropertyTable = $SUPPropertyTable|Select Name,Value
      Write-HtmlTable -InputObject $SUPPropertyTable -Border 1 -Level 4 -File $FilePath
      }
  }
  else
  {
    Write-HTMLParagraph -Text "This site has no Software Update Points installed." -Level 3 -File $FilePath
  }
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion enumerating Software Update Points and Configuration
}


##### Hierarchy wide configuration
Write-HTMLHeading -Level 1 -PageBreak -Text "Summary of Hierarchy Wide Configuration" -File $FilePath

#region enumerating Boundaries
Write-Verbose "$(Get-Date): Enumerating all Site Boundaries"
Write-HTMLHeading -Level 2 -Text "Summary of Site Boundaries" -File $FilePath

$Boundaries = Get-CMBoundary
    if (-not [string]::IsNullOrEmpty($Boundaries))
{
  $SubnetBoundaryTable = @();
  $ADBoundaryTable = @();
  $IPv6BoundaryTable = @();
  $IPRangeTable = @();
  
  ##Boundary Site Types: 0=IP Subnet; 1=AD Site; 2=IPv6 Prefix; 3=IP Address Range
  foreach ($Boundary in $Boundaries) {       
    if ($Boundary.BoundaryType -eq 0) {
      $BoundaryType = 'IP Subnet';
      $NamesOfBoundarySiteSystems = $Null
      if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
      {
        ForEach-Object -Begin {$BoundarySiteSystems= $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
      }
      else 
      {
        $NamesOfBoundarySiteSystems = 'n/a'
      } 
      $Subnet = New-Object -TypeName psobject -Property @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)";
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems"
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $SubnetBoundaryTable += $Subnet;
    }
    elseif ($Boundary.BoundaryType -eq 1) { 
      $BoundaryType = 'Active Directory Site';
      $NamesOfBoundarySiteSystems = $Null
      if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
      {
        ForEach-Object -Begin {$BoundarySiteSystems= $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
      }
      else 
      {
        $NamesOfBoundarySiteSystems = 'n/a'
      } 
      $ADBoundary = New-Object -TypeName psobject -Property @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)";
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems";
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $ADBoundaryTable += $ADBoundary;
    }
    elseif ($Boundary.BoundaryType -eq 2) { 
      $BoundaryType = 'IPv6 Prefix';
      $NamesOfBoundarySiteSystems = $Null
      if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
      {
        ForEach-Object -Begin {$BoundarySiteSystems= $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
      }
      else 
      {
        $NamesOfBoundarySiteSystems = 'n/a'
      } 
      $IPv6Boundary = New-Object -TypeName psobject -Property @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)";
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems";
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $IPv6BoundaryTable += $IPv6Boundary;
    }
    elseif ($Boundary.BoundaryType -eq 3) 
    { 
      $BoundaryType = 'IP Range';
      $NamesOfBoundarySiteSystems = $Null
      if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
      {
        ForEach-Object -Begin {$BoundarySiteSystems= $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
      }
      else 
      {
        $NamesOfBoundarySiteSystems = 'n/a'
      } 
      $IPRangeBoundary = New-Object -TypeName psobject -Property @{'Boundary Type' = $BoundaryType;
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)";
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems";
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $IPRangeTable += $IPRangeBoundary
    }
  }
}

#region IPv6 Boundaries Table
      Write-HTMLHeading -Level 3 -Text "IPv6 Boundaries" -File $FilePath
      If ($IPv6BoundaryTable){
          $IPv6BoundaryTable = $IPv6BoundaryTable|select @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $IPv6BoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IPv6 boundaries defined." -Level 3 -File $FilePath
      }
#endregion IPv6 Boundaries Table
      Write-HTMLParagraph -Text '&nbsp;' -File $FilePath

#region IP Subnet Boundaries Table
      Write-HTMLHeading -Level 3 -Text "IP Subnet Boundaries" -File $FilePath
      If ($SubnetBoundaryTable){
          $SubnetBoundaryTable = $SubnetBoundaryTable|select @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $SubnetBoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IP subnet boundaries defined." -Level 3 -File $FilePath
      }
#endregion IP Subnet Boundaries Table
      Write-HTMLParagraph -Text '&nbsp;' -File $FilePath

#region IP Range Boundaries Table
      Write-HTMLHeading -Level 3 -Text "IP Range Boundaries" -File $FilePath
      If ($IPRangeTable){
          $IPRangeTable = $IPRangeTable|select @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $IPRangeTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IP Range boundaries defined." -Level 3 -File $FilePath
      }
#endregion IP Range Boundaries Table

#region AD Site Boundaries Table
      Write-HTMLHeading -Level 3 -Text "AD Site Boundaries" -File $FilePath
      If ($ADBoundaryTable){
          $ADBoundaryTable = $ADBoundaryTable|select @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $ADBoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No AD Site boundaries defined." -Level 3 -File $FilePath
      }
#endregion AD Site Boundaries Table
    Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating Boundaries


#region enumerating all Boundary Groups and their members

Write-HTMLHeading -Level 2 -Text "Site Boundary Groups" -PageBreak -File $FilePath

#User Defined Boundary Groups
Write-Verbose "$(Get-Date):   Enumerating all Boundary Groups and their members"

$BoundaryGroups = Get-CMBoundaryGroup
Write-HTMLHeading -Level 3 -Text "User Defined Boundary Groups" -File $FilePath

$BoundaryGroupTable = @();
if (-not [string]::IsNullOrEmpty($BoundaryGroups))
{
  foreach ($BoundaryGroup in $BoundaryGroups) {
    $BGSystems = @()
    $MemberNames = @();
    if ($BoundaryGroup.SiteSystemCount -gt 0)
    {
        $CMSiteSystems = Get-WmiObject -Class SMS_BoundaryGroupSiteSystems -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where {$_.GroupID -eq "$($BoundaryGroup.GroupID)"}
        foreach($SS in $CMSiteSystems){
            $BGSystems +=[regex]::Match($SS.ServerNALPath,'\[\"Display=\\\\(.*)\\\"\]MSWNET').Groups[1].value
        }
        $BoundaryGroupSiteSystems = $BGSystems -join '--CRLF--'
    }
    Else
    {
        $BoundaryGroupSiteSystems = "None"
    }
    $MemberIDs = (Get-WmiObject -Class SMS_BoundaryGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object -FilterScript {$_.GroupID -eq "$($BoundaryGroup.GroupID)"}).BoundaryID
    if ($MemberIDs)
    {
      foreach ($MemberID in $MemberIDs)
      {
        $MemberName = (Get-CMBoundary -Id $MemberID).Value
        $MemberNames += "$MemberName (ID: $MemberID)"
        Write-Verbose "Member name: $($MemberName)"
      }
    }
    else
    {
      $MemberNames += 'No associated boundaries'
      Write-Verbose 'There are no boundaries associated with this Boundary Group.'
    }
    $BoundaryMembers = $MemberNames -join "--CRLF--"
    $BoundaryGroupRow = New-Object -TypeName psobject -Property @{Name = $BoundaryGroup.Name; Description = $BoundaryGroup.Description; 'Boundary Members' = "$BoundaryMembers"; 'Site Systems' = $BoundaryGroupSiteSystems};
    $BoundaryGroupTable += $BoundaryGroupRow
  }
  
  $BoundaryGroupTable = $BoundaryGroupTable|select 'Name','Description','Boundary Members','Site Systems'
  Write-HtmlTable -InputObject $BoundaryGroupTable -Level 4 -Border 1 -File $FilePath
}
else
{
  Write-HTMLParagraph -Level 3 -Text "There are no Boundary Groups configured. It is mandatory to configure a Boundary Group for Configuration Manger to work properly." -File $FilePath
}
#End User Defined Boundary Groups

#Default Boundary Group
Write-HTMLHeading -Level 3 -Text "Default Boundary Group" -File $FilePath

$DefaultBG = Get-CMDefaultBoundaryGroup
$DefaultBGID = $DefaultBG.GroupID
$BGSystems = @()
if ($DefaultBG.SiteSystemCount -gt 0)
{
    $CMSiteSystems = Get-WmiObject -Class SMS_BoundaryGroupSiteSystems -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where {$_.GroupID -eq "$DefaultBGID"}
    foreach($SS in $CMSiteSystems){
        $BGSystems +=[regex]::Match($SS.ServerNALPath,'\[\"Display=\\\\(.*)\\\"\]MSWNET').Groups[1].value
    }
    $BoundaryGroupSiteSystems = $BGSystems -join '--CRLF--'
}
Else
{
    $BoundaryGroupSiteSystems = "None"
}
$DefaultBGRelationship = Get-WmiObject -Class SMS_BoundaryGroupRelationships -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where {($_.SourceGroupID -eq "$DefaultBGID") -and ($_.DestinationGroupID -eq "$DefaultBGID")}
$FallbackSUP = $DefaultBGRelationship.FallbackSUP
$FallbackDP = $DefaultBGRelationship.FallbackDP
IF ($FallbackSUP -eq -1) {$FallbackSUP = "Never"}else{$FallbackSUP = "$FallbackSUP mins"}
IF ($FallbackDP -eq -1) {$FallbackDP = "Never"}else{$FallbackDP = "$FallbackDP mins"}


$DefaultBoundaryGroupRow = New-Object -TypeName psobject -Property @{Name = $DefaultBG.Name; 'Site Systems' = $BoundaryGroupSiteSystems; 'DP Fallback Time' = $FallbackDP; 'SUP Fallback Time' = $FallbackSUP};
$DefaultBoundaryGroupRow = $DefaultBoundaryGroupRow|select 'Name','Site Systems','DP Fallback Time','SUP Fallback Time'
Write-HtmlTable -InputObject $DefaultBoundaryGroupRow -Level 4 -Border 1 -File $FilePath

#End Default Boundary Group
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all Boundary Groups and their members


#region enumerating Client Policies
Write-Verbose "$(Get-Date):   Enumerating all Client/Device Settings"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Summary of Custom Client Device Settings' -File $FilePath

$AllClientSettings = Get-CMClientSetting | Where-Object -FilterScript {$_.SettingsID -ne '0'}

foreach ($ClientSetting in $AllClientSettings)
{
  $SettingInfo = @()
  $SettingInfo += "Client Settings Description: $($ClientSetting.Description)"
  $SettingInfo += "Client Settings ID: $($ClientSetting.SettingsID)"
  $SettingInfo += "Client Settings Priority: $($ClientSetting.Priority)"
  if ($ClientSetting.Type -eq '1')
  {
    $SettingDescription = 'This is a custom client Device Setting.'
  }
  else
  {
    $SettingDescription = 'This is a custom client User Setting.'
  }
  Write-HTMLHeading -Level 3 -Text "Client Settings Name: $($ClientSetting.Name)" -File $FilePath
  Write-HtmlList -InputObject $SettingInfo -Description $SettingDescription -Level 2 -File $FilePath
  Write-HTMLParagraph -Level 3 -Text "<u><b>Setting Configuration</b></u>:" -File $FilePath
  foreach ($AgentConfig in $ClientSetting.AgentConfigurations)
  {
    try
    {
      switch ($AgentConfig.AgentID)
      {
        1{
          $Config = 'Compliance Settings'
          $KnownProps = @("AgentID","Enabled","EnableUserStateManagement","EvaluationSchedule","PerProviderTimeout","PerScanDefaultPriority","PerScanTimeout","PerScanTTL","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $ConfigList = @()
          $ConfigList += "Enable compliance evaluation on clients: $($AgentConfig.Enabled)"
          $ConfigList += "Enable user data and profiles: $($AgentConfig.EnableUserStateManagement)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        2{
          $Config = 'Software Inventory'
          $KnownProps = @("AgentID","CollectableFileExclude","CollectableFileMaxSize","CollectableFilePaths","CollectableFiles","CollectableFileSubdirectories","Enabled","Exclude","ExcludeWindirAndSubfolders","InventoriableTypes","Path","PSComputerName","PSShowComputerName","QueryTimeout","ReportOptions","ReportTimeout","ScanInterval","Schedule","SmsProviderObjectPath","Subdirectories")
          $ConfigList = @()
          $ConfigList += "Enable software inventory on clients: $($AgentConfig.Enabled)"
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
          if ($Schedule.DaySpan -gt 0)
          {
            $ConfigList += "Schedule: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $ConfigList += "Schedule: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $ConfigList += "Schedule: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $ConfigList += "Schedule: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $ConfigList += "Schedule: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $ConfigList += "Schedule: Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $ConfigList += "Schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          switch ($AgentConfig.ReportOptions)
          {
            1 { $InvDetail = 'Product only' }
            2 { $InvDetail = 'File only' }
            7 { $InvDetail = 'Full details' }
          }
          $ConfigList += "Inventory reporting detail: $InvDetail"

          if ($AgentConfig.InventoriableTypes)
          {
            $counter = 0
            $InvFiles = @()
            Foreach ($type in $AgentConfig.InventoriableTypes) {
                $InvFilePath = $AgentConfig.Path[$counter]
                $InvFileSubF = If($AgentConfig.Subdirectories[$counter] -eq "true"){"Yes"}Else{"No"}
                $InvFileWinD = If($AgentConfig.ExcludeWindirAndSubfolders[$counter] -eq "true"){"Windows"}Else{""}
                $InvFileComp = If($AgentConfig.Exclude[$counter] -eq "true"){"Compressed"}Else{""}
                $InvFileExcl = ("$InvFileWinD,$InvFileComp").Trim(',')
                $InvFiles +=  New-Object -TypeName psobject -Property @{'Name' = "$type"; 'Path' = "$InvFilePath"; 'Subfolders' = "$InvFileSubF"; 'Exclude' =  "$InvFileExcl"}
                $counter++
            }
          }
          if ($AgentConfig.CollectableFiles) {
            $counter = 0
            $CollectedFiles = @()
            Foreach ($CollFile in $AgentConfig.CollectableFiles) {
                $CollFileName = $AgentConfig.CollectableFiles[$counter]
                $CollFilePath = $AgentConfig.CollectableFilePaths[$counter]
                $CollFileSubF = If($AgentConfig.CollectableFileSubdirectories[$counter] -eq "true"){"Yes"}Else{"No"}
                $CollFileSize = $AgentConfig.CollectableFileMaxSize[$counter]
                $CollFileExclude = If($AgentConfig.CollectableFileExclude[$counter] -eq "true"){"Compressed"}Else{"None"}
                $CollectedFiles +=  New-Object -TypeName psobject -Property @{'Name' = "$CollFileName"; 'Path' = "$CollFilePath"; 'Subfolders' = "$CollFileSubF"; 'Size' = "$CollFileSize"; 'Exclude' =  "$CollFileExclude"}
                $counter++
            }
          }
          $InvFiles = $InvFiles | Select-Object Name,Path,Subfolders,Exclude
          $CollectedFiles = $CollectedFiles | Select-Object Name,Path,Subfolders,Size,Exclude
          if ($InvFiles.count -gt 0) {
              $ConfigList += 'Inventory these file types:'
              Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
              Write-HtmlTable -InputObject $InvFiles -Level 7 -Border 1 -File $FilePath
          } else {
              $ConfigList += 'Inventory these file types: None'
              Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          }
          if ($CollectedFiles.count -gt 0) {
              Write-HtmlList -InputObject 'Collect Files:' -Level 3 -File $FilePath
              Write-HtmlTable -InputObject $CollectedFiles -Level 7 -Border 1 -File $FilePath
          } else {
              Write-HtmlList -InputObject 'Collect Files: None' -Level 3 -File $FilePath
          }
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        3{
          $KnownProps = @("AccessLevel","AgentID","AllowClientChange","AllowLocalAdminToDoRemoteControl","AllowRAUnsolicitedControl","AllowRAUnsolicitedView","AllowRemCtrlToUnattended","AudibleSignal","ClipboardAccessPermissionRequired","Enabled","EnableRA","EnableTS","EnforceRAandTSSettings","FirewallExceptionProfiles","ManageRA","ManageTS","PermissionRequired","PermittedViewers","PSComputerName","PSShowComputerName","RemCtrlConnectionBar","RemCtrlTaskbarIcon","SmsProviderObjectPath","TSUserAuthentication")
          $Config = 'Remote Tools'
          $ConfigList = @()
          #Remote Control enabled or not?  And are firewall exceptions configured?
          switch ($AgentConfig.FirewallExceptionProfiles)
          {
            0 { $RCState = 'Disabled' }
            8 { $RCState = 'Enabled - Firewall Profiles Configured: None' }
            9 { $RCState = 'Enabled - Firewall Profiles Configured: Public' }
            10 { $RCState = 'Enabled - Firewall Profiles Configured: Private' }
            11 { $RCState = 'Enabled - Firewall Profiles Configured: Private, Public' }
            12 { $RCState = 'Enabled - Firewall Profiles Configured: Domain' }
            13 { $RCState = 'Enabled - Firewall Profiles Configured: Domain, Public' }
            14 { $RCState = 'Enabled - Firewall Profiles Configured: Domain, Private' }
            15 { $RCState = 'Enabled - Firewall Profiles Configured: Domain, Private, Public' }
          }
          $ConfigList += "Enable Remote Control on clients: $RCState"
          $ConfigList += "Users can change policy or notification settings in Software Center: $($AgentConfig.AllowClientChange)"
          $ConfigList += "Allow Remote Control of an unattended computer: $($AgentConfig.AllowRemCtrlToUnattended)"
          $ConfigList += "Prompt user for Remote Control permission: $($AgentConfig.PermissionRequired)"
          $ConfigList += "Prompt user for permission to transfer content from shared clipboard: $($AgentConfig.ClipboardAccessPermissionRequired)"
          $ConfigList += "Grant Remote Control permission to local Administrators group: $($AgentConfig.AllowLocalAdminToDoRemoteControl)"
          switch ($AgentConfig.AccessLevel)
          {
            0 { $accesslevel = 'No access' }
            1 { $accesslevel = 'View only' }
            2 { $accesslevel = 'Full Control' }
          }
          $ConfigList += "Access level allowed: $accesslevel"
          if ($AgentConfig.PermittedViewers.count -gt 0) {
            $viewers = Write-HtmlList -InputObject ($AgentConfig.PermittedViewers)
            $ConfigList += "Permitted viewers of Remote Control and Remote Assistance: $viewers"
          } Else {
            $ConfigList += "Permitted viewers of Remote Control and Remote Assistance: None"
          }
          $ConfigList += "Show session notification icon on taskbar: $($AgentConfig.RemCtrlTaskbarIcon)"
          $ConfigList += "Show session connection bar: $($AgentConfig.RemCtrlConnectionBar)"
          Switch ($AgentConfig.AudibleSignal)
          {
            0 { $ClientSound = 'None.' }
            1 { $ClientSound = 'Beginning and end of session.' }
            2 { $ClientSound = 'Repeatedly during session.' }
          }
          $ConfigList += "Play a sound on client: $ClientSound"
          $ConfigList += "Manage unsolicited Remote Assistance settings: $($AgentConfig.ManageRA)"
          $ConfigList += "Manage solicited Remote Assistance settings: $($AgentConfig.EnforceRAandTSSettings)"
          #Level of access for Remote Assistance:
          if (($AgentConfig.AllowRAUnsolicitedView -ne 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -ne 'True'))
          {
            $RALevel = 'None'
          }
          elseif (($AgentConfig.AllowRAUnsolicitedView -eq 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -ne 'True'))
          {
            $RALevel = 'Remote viewing'
          }
          elseif (($AgentConfig.AllowRAUnsolicitedView -eq 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -eq 'True'))
          {
            $RALevel = 'Full Control'
          }
          $ConfigList += "Level of access for Remote Assistance: $RALevel"
          $ConfigList += "Manage Remote Desktop settings: $($AgentConfig.ManageTS)"
          if ($AgentConfig.ManageTS -eq 'True')
          {
            $ConfigList += "Allow permitted viewers to connect by using Remote Desktop connection: $($AgentConfig.EnableTS)"
            $ConfigList += "Require network level authentication on computers that run Windows Vista operating system and later versions: $($AgentConfig.TSUserAuthentication)"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        4{
          $KnownProps = @("AddPortalToTrustedSiteList","AgentID","AllowPortalToHaveElevatedTrust","BrandingTitle","DayReminderInterval","DisplayNewProgramNotification","EnableHealthAttestation","EnableThirdPartyOrchestration","GracePeriodHours","HourReminderInterval","InstallRestriction","OnPremHAServiceUrl","OSDBrandingSubTitle","PortalUrl","PowerShellExecutionPolicy","PSComputerName","PSShowComputerName","ReminderInterval","SmsProviderObjectPath","SUMBrandingSubTitle","SuspendBitLocker","SWDBrandingSubTitle","SystemRestartTurnaroundTime","UseNewSoftwareCenter","UseOnPremHAService")
          $Config = 'Computer Agent'
          $ConfigList = @()
          $ConfigList += "Deployment deadline greater than 24 hours, remind user every (hours): $([string]($AgentConfig.ReminderInterval) / 60 / 60)"
          $ConfigList += "Deployment deadline less than 24 hours, remind user every (hours): $([string]($AgentConfig.DayReminderInterval) / 60 / 60)"
          $ConfigList += "Deployment deadline less than 1 hour, remind user every (minutes): $([string]($AgentConfig.HourReminderInterval) / 60)"
          $ConfigList += "Default application catalog website point: $($AgentConfig.PortalUrl)"
          $ConfigList += "Add default Application Catalog website to Internet Explorer trusted sites zone: $($AgentConfig.AddPortalToTrustedSiteList)"
          $ConfigList += "Allow Silverlight applications to run in elevated trust mode: $($AgentConfig.AllowPortalToHaveElevatedTrust)"
          $ConfigList += "Organization name displayed in Software Center: $($AgentConfig.BrandingTitle)"
          $ConfigList += "Use New Software Center: $($AgentConfig.UseNewSoftwareCenter)"
          $ConfigList += "Enable communication with Health Attestation Service: $($AgentConfig.EnableHealthAttestation)"
          $ConfigList += "Use on-premises Health Attestation Service: $($AgentConfig.UseOnPremHAService)"
          switch ($AgentConfig.InstallRestriction)
          {
            0 { $InstallRestriction = 'All Users' }
            1 { $InstallRestriction = 'Only Administrators' }
            3 { $InstallRestriction = 'Only Administrators and primary Users'}
            4 { $InstallRestriction = 'No users' }
          }
          $ConfigList += "Install Permissions: $($InstallRestriction)"
          Switch ($AgentConfig.SuspendBitLocker)
          {
            0 { $SuspendBitlocker = 'Never' }
            1 { $SuspendBitlocker = 'Always' }
          }
          $ConfigList += "Suspend Bitlocker PIN entry on restart: $($SuspendBitlocker)"
          Switch ($AgentConfig.EnableThirdPartyOrchestration)
          {
            0 { $EnableThirdPartyTool = 'No' }
            1 { $EnableThirdPartyTool = 'Yes' }
          }
          $ConfigList += "Additional software manages the deployment of applications and software updates: $($EnableThirdPartyTool)"
          Switch ($AgentConfig.PowerShellExecutionPolicy)
          {
            0 { $ExecutionPolicy = 'All signed' }
            1 { $ExecutionPolicy = 'Bypass' }
            2 { $ExecutionPolicy = 'Restricted' }
          }
          $ConfigList += "Powershell execution policy: $($ExecutionPolicy)"
          switch ($AgentConfig.DisplayNewProgramNotification)
          {
            False { $DisplayNotifications = 'No' }
            True { $DisplayNotifications = 'Yes' }
          }
          $ConfigList += "Show notifications for new deployments: $($DisplayNotifications)"
          #The deadline randomization setting now appears in AgentID 25.  But since in GUI under 'Computer Agent', we will loop the config to find agent 25 and get the setting here.
          foreach ($AC in $ClientSetting.AgentConfigurations){
            If ($AC.AgentID -eq 25) {
                  switch ($AC.DisableGlobalRandomization)
                  {
                    False { $DisableGlobalRandomization = 'No' }
                    True { $DisableGlobalRandomization = 'Yes' }
                  }
            }
          }
          $ConfigList += "Disable deadline randomization: $($DisableGlobalRandomization)"
          $ConfigList += "Grace period for enforcement after deployment deadline (hours): $($AgentConfig.GracePeriodHours)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        5{
        #AgentID 5 is for Network Access Protection (NAP) and is no longer part of the product.
        }
        8{
          $KnownProps = @("AgentID","DataCollectionSchedule","Enabled","LastUpdateTimeOfRules","MaximumUsageInstancesPerReport","MeterRuleIDList","MRUAgeLimitInDays","MRURefreshInMinutes","PSComputerName","PSShowComputerName","ReportTimeout","SmsProviderObjectPath")
          $Config = 'Software Metering'
          $ConfigList = @()
          $ConfigList += "Enable software metering on clients: $($AgentConfig.Enabled)"
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.DataCollectionSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            $DCSched = " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $DCSched = " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $DCSched = " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $DCSched = " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $DCSched = " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $DCSched = " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $DCSched = " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          $ConfigList += "Schedule data collection: $DCSched"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        9{
          $KnownProps = @("AgentID","AlternateContentProviders","AssignmentBatchingTimeout","BrandingSubTitle","BrandingTitle","ContentDownloadTimeout","ContentLocationTimeout","DayReminderInterval","Enabled","EnableExpressUpdates","EvaluationSchedule","ExpressUpdatesPort","HourReminderInterval","MaxRandomDelayMinutes","MaxScanRetryCount","O365Management","PerDPInactivityTimeout","PSComputerName","PSShowComputerName","ReminderInterval","ScanRetryDelay","ScanSchedule","SmsProviderObjectPath","TotalInactivityTimeout","UpdateStatusRefreshIntervalDays","UserExperience","UserJobPerDPInactivityTimeout","UserJobTotalInactivityTimeout","WSUSLocationTimeout","WSUSScanRetryCodes","WUAMaxRebootsWhenOnInternet","WUASuccessCodes","WUfBEnabled","EnableThirdPartyUpdates")
          $Config = 'Software Updates'
          $ConfigList = @()
          $ConfigList += "Enable software updates on clients: $($AgentConfig.Enabled)"
          ##Software Update scan schedule:
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.ScanSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            $SoftScanSched = " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $SoftScanSched = " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $SoftScanSched = " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $SoftScanSched = " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $SoftScanSched = " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $SoftScanSched = " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $SoftScanSched = " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          $ConfigList += "Software Update scan schedule: $SoftScanSched"
          ##Schedule deployment re-evaluation:
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            $SoftReevalSched = " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $SoftReevalSched = " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $SoftReevalSched = " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $SoftReevalSched = " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $SoftReevalSched = " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $SoftReevalSched = " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $SoftReevalSched = " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          $ConfigList += "Schedule deployment re-evaluation: $SoftReevalSched"
          if ($AgentConfig.AssignmentBatchingTimeout -eq '0')
          {
            $ConfigList += "When any software update deployment deadline is reached, install all other software update deployments with deadline coming within a specified period of time: No"
          }
          else 
          {
            $ConfigList += "When any software update deployment deadline is reached, install all other software update deployments with deadline coming within a specified period of time: Yes"
            
            if ($AgentConfig.AssignmentBatchingTimeout -le '82800')
            {
              $hours = [string]$AgentConfig.AssignmentBatchingTimeout / 60 / 60 
              $gracetime = "$($hours) hours"
            }
            else 
            {
              $days = [string]$AgentConfig.AssignmentBatchingTimeout / 60 / 60 / 24
              $gracetime = "$($days) days"
            }
            $ConfigList += "Period of time for which all pending deployments with deadline in this time will also be installed: $gracetime"
          }
          if($AgentConfig.EnableExpressUpdates -eq $False)
          {
              $ConfigList += "Enable installation of Express installation files on clients: No"
          }
          else
          {
              $ConfigList += "Enable installation of Express installation files on clients: Yes"
              $ConfigList += "Port used to download content for Express installation files: $($AgentConfig.ExpressUpdatesPort)"
          }
          If($AgentConfig.O365Management -eq 1)
          {
              $ConfigList += "Enable management of the Office 365 Client Agent: Yes"
          }
          else
          {
              $ConfigList += "Enable management of the Office 365 Client Agent: No"
          }
          If($AgentConfig.EnableThirdPartyUpdates -eq "True")
          {
              $ConfigList += "Enable Third Party Software Updates: Yes"
          }
          else
          {
              $ConfigList += "Enable Third Party Software Updates: No"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        10{
          $KnownProps = @("AgentID","AllowUserAffinity","AllowUserAffinityAfterMinutes","AutoApproveAffinity","ConsoleMinutes","IntervalDays","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'User and Device Affinity'
          $ConfigList = @()
          if ($ClientSetting.Type -eq '1'){
              $ConfigList += "User device affinity usage threshold (minutes): $($AgentConfig.ConsoleMinutes)"
              $ConfigList += "User device affinity usage threshold (days): $($AgentConfig.IntervalDays)"
              if ($AgentConfig.AutoApproveAffinity -eq '0')
              {
                $AAAffinity = 'No'
              }
              else
              {
                $AAAffinity = 'Yes'
              }
              $ConfigList += "Automatically configure user device affinity from usage data: $AAAffinity"
          }Else{
              IF ($($AgentConfig.AllowUserAffinity) -eq '1'){
                $UserDefinedAffinity = 'Yes'
              }Else{
                $UserDefinedAffinity = 'No'
              }
              $ConfigList += "Allow user to define their primary devices: $UserDefinedAffinity"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        11{
          $KnownProps = @("AgentID","ApplyToAllClients","EnableBitsMaxBandwidth","EnableDownloadOffSchedule","MaxBandwidthValidFrom","MaxBandwidthValidTo","MaxTransferRateOffSchedule","MaxTransferRateOnSchedule","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'Background Intelligent Transfer'
          $ConfigList = @()
          $ConfigList += "Limit the maximum network bandwidth for BITS background transfers: $($AgentConfig.EnableBitsMaxBandwidth)"
          $ConfigList += "Throttling window start time: $($AgentConfig.MaxBandwidthValidFrom)"
          $ConfigList += "Throttling window end time: $($AgentConfig.MaxBandwidthValidTo)"
          $ConfigList += "Maximum transfer rate during throttling window (kbps): $($AgentConfig.MaxTransferRateOnSchedule)"
          $ConfigList += "Allow BITS downloads outside the throttling window: $($AgentConfig.EnableDownloadOffSchedule)"
          $ConfigList += "Maximum transfer rate outside the throttling window (Kbps): $($AgentConfig.MaxTransferRateOffSchedule)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        12{
          $KnownProps = @("AgentID","DeviceEnrollmentProfileID","EnableDeviceEnrollment","EnableFileCollection","EnableHardwareInventory","EnableModernDeviceEnrollment","EnableSoftwareDistribution","EnableSoftwareInventory","FailureRetryCount","FailureRetryInterval","FileCollectionExcludeCompressed","FileCollectionExcludeEncrypted","FileCollectionFilter","FileCollectionInterval","FileCollectionPath","FileCollectionSubdirectories","HardwareInventoryInterval","MDMPollInterval","ModernDeviceEnrollmentProfileID","PollingInterval","PollServer","PSComputerName","PSShowComputerName","SmsProviderObjectPath","SoftwareInventoryExcludeCompressed","SoftwareInventoryExcludeEncrypted","SoftwareInventoryFilter","SoftwareInventoryInterval","SoftwareInventoryPath","SoftwareInventorySubdirectories")
          $Config = 'Enrollment'
          $ConfigList = @()
          if ($ClientSetting.Type -eq '1'){
              $ConfigList += "Polling interval for modern devices (minutes): $($AgentConfig.MDMPollInterval)"
          } Else {
              If ($AgentConfig.EnableDeviceEnrollment -eq '1'){
                $ConfigList += 'Allow users to enroll mobile devices and Mac computers: Yes'
                $MacDEID = "$($AgentConfig.DeviceEnrollmentProfileID)"
                $MacDEName = (Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "Select * from SMS_DeviceEnrollmentProfile where ProfileID = `'$($AgentConfig.DeviceEnrollmentProfileID)`'").Name
                $ConfigList += "Enrollment Profile: $MacDEName (ID: $MacDEID)"
              }else{
                $ConfigList += 'Allow users to enroll mobile devices and Mac computers: Yes'
              }
              If ($AgentConfig.EnableModernDeviceEnrollment -eq '1'){
                $ConfigList += "Allow users to enroll modern devices: Yes"
                $ModernDEID = "$($AgentConfig.ModernDeviceEnrollmentProfileID)"
                $ConfigList += "Modern device enrollment profile: $ModernDEID"
              }else{
                $ConfigList += "Allow users to enroll modern devices: No"
              }
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        13{
          $KnownProps = @("AgentID","PolicyDownloadMethod","PolicyEnableUserAuthForAllUserPolicies","PolicyEnableUserGroupSupport","PolicyEnableUserPolicyOnInternet","PolicyEnableUserPolicyPolling","PolicyRequestAssignmentTimeout","PolicyTimeDelayBeforeUserPolicyRefreshAtLogonOrUnlock","PolicyTimeUntilAck","PolicyTimeUntilExpire","PolicyTimeUntilUpdateActualConfig","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'Client Policy'
          $ConfigList = @()
          $ConfigList += "Client policy polling interval (minutes): $($AgentConfig.PolicyRequestAssignmentTimeout)"
          $ConfigList += "Enable user policy on clients: $($AgentConfig.PolicyEnableUserPolicyPolling)"
          $ConfigList += "Enable user policy requests from Internet clients: $($AgentConfig.PolicyEnableUserPolicyOnInternet)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        15{
          $KnownProps = @("AgentID","Enabled","InventoryReportID","LastUpdateTime","Max3rdPartyMIFSize","MaxRandomDelayMinutes","MIFCollection","ProviderTimeout","PSComputerName","PSShowComputerName","Schedule","SmsProviderObjectPath")
          $Config = 'Hardware Inventory'
          $ConfigList = @()
          $ConfigList += "Enable hardware inventory on clients: $($AgentConfig.Enabled)"
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
          if ($Schedule.DaySpan -gt 0)
          {
            $ConfigList += "Hardware inventory schedule: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $ConfigList += "Hardware inventory schedule: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $ConfigList += "Hardware inventory schedule: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $ConfigList += "Hardware inventory schedule: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $ConfigList += "Hardware inventory schedule: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $ConfigList += "Hardware inventory schedule: Occurs on last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $ConfigList += "Hardware inventory schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          $ConfigList += "Maximum random delay (minutes): $($AgentConfig.MaxRandomDelayMinutes)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        16{
          $KnownProps = @("AgentID","BulkSendInterval","BulkSendIntervalHigh","BulkSendIntervalLow","CacheCleanoutInterval","CacheMaxAge","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'State Messaging'
          $ConfigList = @()
          $ConfigList += "State message reporting cycle (minutes): $($AgentConfig.BulkSendInterval)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        17{
          $KnownProps = @("AgentID","AlternateContentProviders","AppXInplaceUpgradeEnabled","Enabled","EvaluationSchedule","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'Software Deployment'
          $ConfigList = @()
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            $ConfigList += "Schedule re-evaluation for deployments: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            $ConfigList += "Schedule re-evaluation for deployments: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            $ConfigList += "Schedule re-evaluation for deployments: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            $ConfigList += "Schedule re-evaluation for deployments: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              $ConfigList += "Schedule re-evaluation for deployments: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              $ConfigList += "Schedule re-evaluation for deployments: Occurs on last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0)
            {
              switch ($Schedule.WeekOrder)
              {
                0 {$order = 'last'}
                1 {$order = 'first'}
                2 {$order = 'second'}
                3 {$order = 'third'}
                4 {$order = 'fourth'}
              }
              $ConfigList += "Schedule re-evaluation for deployments: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        18{
          $KnownProps = @("AgentID","AllowUserToOptOutFromPowerPlan","Enabled","EnableP2PWakeupSolution","EnableUserIdleMonitoring","EnableWakeupProxy","MaxCPU","MaxMachinesPerManager","MinimumServersNeeded","NumOfDaysToKeep","NumOfMonthsToKeep","Port","PSComputerName","PSShowComputerName","SmsProviderObjectPath","WakeupProxyDirectAccessPrefixList","WakeupProxyFirewallFlags","WolPort")
          $Config = 'Power Management'
          $ConfigList = @()
          $ConfigList += "Allow power management of clients: $($AgentConfig.Enabled)"
          $ConfigList += "Allow users to exclude their device from power management: $($AgentConfig.AllowUserToOptOutFromPowerPlan)"
          $ConfigList += "Enable wake-up proxy: $($AgentConfig.EnableWakeupProxy)"
          if ($AgentConfig.EnableWakeupProxy -eq 'True')
          {
            $ConfigList += "Wake-up proxy port number (UDP): $($AgentConfig.Port)"
            $ConfigList += "Wake On LAN port number (UDP): $($AgentConfig.WolPort)"
            switch ($AgentConfig.WakeupProxyFirewallFlags)
            {
              0 { $FirewallCfg = 'Disabled' }
              9 { $FirewallCfg = 'Enabled: Public.' }
              10 { $FirewallCfg = 'Enabled: Private.' }
              11 { $FirewallCfg = 'Enabled: Private, Public.' }
              12 { $FirewallCfg = 'Enabled: Domain.' }
              13 { $FirewallCfg = 'Enabled: Domain, Public.' }
              14 { $FirewallCfg = 'Enabled: Domain, Private.' }
              15 { $FirewallCfg = 'Enabled: Domain, Private, Public.' }
            }
            $ConfigList += "Windows Firewall exception for wake-up proxy: $FirewallCfg"
            If ($AgentConfig.WakeupProxyDirectAccessPrefixList -eq ""){
                $v6Prefixes = "None"
            }Else{
                $v6Prefixes = "$($AgentConfig.WakeupProxyDirectAccessPrefixList)"
            }
            $ConfigList += "IPv6 prefixes if required for DirectAccess or other intervening network devices. Use a comma to specifiy multiple entries: $v6Prefixes"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        20{
          $KnownProps = @("AgentID","DisableFirstSignatureUpdate","EnableBlueProvider","EnableEP","ForceRebootPeriod","InstallRetryPeriod","InstallSCEPClient","LicenseAgreed","OverrideMaintenanceWindow","PersistInstallation","PolicyEnforcePeriod","PSComputerName","PSShowComputerName","Remove3rdParty","SmsProviderObjectPath","SuppressReboot")
          $Config = 'Endpoint Protection'
          $ConfigList = @()
          $ConfigList += "Manage Endpoint Protection client on client computers: $($AgentConfig.EnableEP)"
          $ConfigList += "Install Endpoint Protection client on client computers: $($AgentConfig.InstallSCEPClient)"
          $ConfigList += "Automatically remove previously installed antimalware software before Endpoint Protection is installed: $($AgentConfig.Remove3rdParty)"
          $ConfigList += "Allow Endpoint Protection client installation and restarts outside maintenance windows. Maintenance windows must be at least 30 minutes long for client installation: $($AgentConfig.OverrideMaintenanceWindow)"
          $ConfigList += "For Windows Embedded devices with write filters, commit Endpoint Protection client installation (requires restart): $($AgentConfig.PersistInstallation)"
          $ConfigList += "Suppress any required computer restarts after the Endpoint Protection client is installed: $($AgentConfig.SuppressReboot)"
          If($AgentConfig.SuppressReboot -eq $false){
            $ConfigList += "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): $($AgentConfig.ForceRebootPeriod)"
          }Else{
            $ConfigList += "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): N/A"
          }
          $ConfigList += "Disable alternate sources (such as Microsoft Windows Update, Microsoft Windows Server Update Services, or UNC shares) for the initial definition update on client computers: $($AgentConfig.DisableFirstSignatureUpdate)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        21{
          $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","RebootLogoffNotificationCountdownDuration","RebootLogoffNotificationFinalWindow","SmsProviderObjectPath")
          $Config = 'Computer Restart'
          $ConfigList = @()
          $ConfigList += "Display a temporary notification to the user that indicates the interval before the user is logged of or the computer restarts (minutes): $($AgentConfig.RebootLogoffNotificationCountdownDuration)"
          $ConfigList += "Display a dialog box that the user cannot close, which displays the countdown interval before the user is logged of or the computer restarts (minutes): $([string]$AgentConfig.RebootLogoffNotificationFinalWindow / 60)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        22{
          $KnownProps = @("AADAuthFlags","AgentID","AllowCloudDP","AllowCMG","AutoAADJoin","AutoMDMEnrollment","CoManagementFlags","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'Cloud Services'
          $ConfigList = @()
          $ConfigList += "Allow access to Cloud Distribution Point: $($AgentConfig.AllowCloudDP)"
          $ConfigList += "Automatically register new Windows 10 domain joined devices with Azure Active Directory: $($AgentConfig.AutoAADJoin)"
          $ConfigList += "Enable clients to use a cloud management gateway: $($AgentConfig.AllowCMG)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        23{
          $KnownProps = @("AgentID","MeteredNetworkUsage","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
          $Config = 'Metered Internet Connections'
          $ConfigList = @()
          switch ($AgentConfig.MeteredNetworkUsage)
          {
            1 { $Usage = 'Allow' }
            2 { $Usage = 'Limit' }
            4 { $Usage = 'Block' }
          }
          $ConfigList += "Specifiy how clients communicate on metered network connections: $($Usage)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        25{
          #The deadline randomization setting now appears in AgentID 25.  But since in GUI under 'Computer Agent' (4), we will loop that into configuration agent 4.
        }
        27{
          $KnownProps = @("AgentID","BranchCacheEnabled","BroadcastPort","CachePartialContent","CanBeSuperPeer","ConfigureBranchCache","ConfigureCacheSize","HttpPort","HttpsEnabled","MaxAvgDiskQueueLength","MaxBranchCacheSizePercent","MaxCacheSizeMB","MaxCacheSizePercent","MaxConnectionCountOnClients","MaxConnectionCountOnServers","MaxPercentProcessorTime","PSComputerName","PSShowComputerName","RejectWhenBatteryLow","SmsProviderObjectPath","UsePartialSource")
          $Config = 'Client Cache Settings'
          $ConfigList = @()
          $ConfigList += "Configure BranchCache: $($AgentConfig.ConfigureBranchCache)"
          $ConfigList += "Enable BranchCache: $($AgentConfig.BranchCacheEnabled)"
          $ConfigList += "Maximum BranchCache cache size (percentage of disk): $($AgentConfig.MaxBranchCacheSizePercent)"
          $ConfigList += "Configure client cache size: $($AgentConfig.ConfigureCacheSize)"
          $ConfigList += "Maximum cache size (MB): $($AgentConfig.MaxCacheSizeMB)"
          $ConfigList += "Maximum cache size (percentage of disk): $($AgentConfig.MaxCacheSizePercent)"
          $ConfigList += "Enable Configuration Manager client in full OS to share content: $($AgentConfig.CanBeSuperPeer)"
          $ConfigList += "Port for initial network broadcast: $($AgentConfig.BroadcastPort)"
          $ConfigList += "Enable HTTPS for client peer communication: $($AgentConfig.HttpsEnabled)"
          $ConfigList += "Port for content download from peer (HTTP/HTTPS): $($AgentConfig.HttpPort)"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        29{
          $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","SmsProviderObjectPath","WACommercialID","WAEnable","WAIEOptInlevel","WAOptInDownlevel","WATelLevel")
          $Config = 'Windows Analytics'
          $ConfigList = @()
          If ($AgentConfig.WAEnable -eq 1){
            $WAEnabled = 'Yes'
          }Else{
            $WAEnabled = 'No'
          }
          $ConfigList += "Manage Windows telemetry settings with Configuration Manager: $WAEnabled"
          If ($WAEnabled -eq 'Yes'){
              $ConfigList += "Commercial ID key: $($AgentConfig.WACommercialID)"
              switch ($AgentConfig.WATelLevel)
              {
                1 { $Level = 'Basic' }
                2 { $Level = 'Enhanced' }
                3 { $Level = 'Full' }
              }
              $ConfigList += "Windows 10 telemetry: $Level"
              switch ($AgentConfig.WAOptInDownlevel)
              {
                0 { $Level = 'Disabled' }
                1 { $Level = 'Enable' }
              }
              $ConfigList += "Windows 8.1 and earlier telemetry: $Level"
              switch ($AgentConfig.WAIEOptInlevel)
              {
                0 { $Level = 'Disabled' }
                1 { $Level = 'Enable for local internet, trusted sites, and machine Local only' }
                2 { $Level = 'Enable for Internet and restricted sites only' }
                3 { $Level = 'Enable for all zones' }
              }
              $ConfigList += "Windows 8.1 and earlier Internet Explorer data collection: $Level"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        30{
          $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","SmsProviderObjectPath","SCBrandingColor","SCBrandingString","SCLogo","SCShowApplicationsTab","SCShowComplianceTab","SCShowInstallationTab","SCShowOptionsTab","SCShowOSDTab","SCShowUpdatesTab","SC_Old_Branding","SettingsXml")
          $Config = 'Software Center'
          $ConfigList = @()
          If ($AgentConfig.SC_Old_Branding -eq 1){
              $ConfigList += "Select these new settings to specify company information: Yes"
              $SCBrand = ([xml]$AgentConfig.SettingsXml).settings
              If (-not [string]::IsNullOrEmpty($SCBrand.'brand-orgname')){
                  $ConfigList += "Organization Name: $($SCBrand.'brand-orgname')"
              }
              If (-not [string]::IsNullOrEmpty($SCBrand.'brand-color')){
                  $BrandColor = [System.Web.HttpUtility]::HtmlEncode($($SCBrand.'brand-color'))
                  $ConfigList += "Color scheme for Software Center: <font Style=`"height: 20px; width: 20px; background-color: $BrandColor;  color: $BrandColor; border-radius: 50%;`">----</font>  $($SCBrand.'brand-color')"
              }
              If (-not [string]::IsNullOrEmpty($SCBrand.'brand-logo')){
                  $EncodedImage=$SCBrand.'brand-logo'
                  $ImageData="data:image/jpg;base64,$EncodedImage"
                  $ConfigList += "Organization Logo Defined:<br /><img src=`"$ImageData`">"
              }
              If ($SCBrand.'software-list'.'unapproved-applications-hidden' -eq 'true'){
                  $ConfigList += "Hide unapproved applications in Software Center: Selected"
              }else{
                  $ConfigList += "Hide unapproved applications in Software Center: Not Selected"
              }
              if ($SCBrand.'software-list'.'installed-applications-hidden' -eq 'true'){
                  $ConfigList += "Hide installed applications in Software Center: Selected"
              }else{
                  $ConfigList += "Hide installed applications in Software Center: Not Selected"
              }
              $tabvisibility = "Select which tabs should be exposed to the end user in Software Center:<br />"
              foreach ($tab in $SCBrand.'tab-visibility'.tab){
                  switch ($tab.name)
                  {
                    'AvailableSoftware' {$tabvisibility = $tabvisibility + " &bull;  Applications: $($tab.visible) <br />"}
                    'Updates' {$tabvisibility = $tabvisibility + " &bull;  Updates: $($tab.visible) <br />"}
                    'OSD' {$tabvisibility = $tabvisibility + " &bull;  Operating Systems: $($tab.visible) <br />"}
                    'InstallationStatus' {$tabvisibility = $tabvisibility + " &bull;  Installation Status: $($tab.visible) <br />"}
                    'Compliance' {$tabvisibility = $tabvisibility + " &bull;  Device Compliance: $($tab.visible) <br />"}
                    'Options' {$tabvisibility = $tabvisibility + " &bull;  Applications: $($tab.visible) <br />"}
                  }
              }
              $ConfigList += $tabvisibility.TrimEnd('<br />')
          }Else{
              $ConfigList += "Select these new settings to specify company information: No"
          }
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
        32{
          $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","SmsProviderObjectPath","EnableWindowsDO")
          $Config = 'Delivery Optimization'
          $ConfigList = @()
          If ($AgentConfig.EnableWindowsDO -eq 'True'){
            $WindowsDO = 'Yes'
          }Else{
            $WindowsDO = 'No'
          }
          $ConfigList += "Use Configuration Manager Boundary Groups for Delivery Optimization Group ID: $WindowsDO"
          Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
          If ($UnknownClientSettings) {
              $UnknownProps = @()
              $props = ($AgentConfig| Get-Member -Type Property).Name
              Foreach ($prop in $props) {
                if ($prop -notin $KnownProps) {$UnknownProps += "Property Name: $Prop -- Assigned Value: $($AgentConfig.$prop)"}
              }
              If ($UnknownProps -gt 0) {
                Write-HtmlList -InputObject $UnknownProps -Description "Unknown Properties:" -Level 3 -File $FilePath
              }
          }
        }
      }
    }
    catch [System.Management.Automation.PropertyNotFoundException] 
    {
      Write-Verbose "$(Get-Date):   Client Settings Property not found."
    }

  }
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating Client Policies

#region Security

Write-Verbose "$(Get-Date):   Collecting all administrative users"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Administrative Users' -File $FilePath

$Admins = Get-CMAdministrativeUser

Write-HTMLParagraph -Text 'Details on all administative accounts in the site:' -Level 2 -File $FilePath

$AdminArray = @();

foreach ($Admin in $Admins) 
{
  switch ($Admin.AccountType)
  {
    0 { $AccountType = 'User' }
    1 { $AccountType = 'Group' }
    2 { $AccountType = 'Machine' } 
  } 
  $AdminArray +=  New-Object -TypeName psobject -Property @{Name = $Admin.LogonName; 'Account Type' = $AccountType; 'Security Roles' = "$($($Admin.RoleNames) -join '--CRLF--')"; 'Security Scopes' = "$($($Admin.CategoryNames) -join '--CRLF--')"; Collections = "$($($Admin.CollectionNames) -join '--CRLF--')";}
}
$AdminArray = $AdminArray| Select-Object -Property 'Name','Account Type','Security Roles','Security Scopes','Collections'
Write-HtmlTable -InputObject $AdminArray -Border 1 -Level 3 -File $FilePath
#endregion Security


#region enumerating all custom Security roles
Write-Verbose "$(Get-Date):   enumerating all custom build security roles"
Write-HTMLHeading -Level 2 -Text 'Custom Security Roles' -File $FilePath
$SecurityRoles = Get-CMSecurityRole | Where-Object -FilterScript {-not $_.IsBuiltIn}
if (-not [string]::IsNullOrEmpty($SecurityRoles))
{
  $SecRoleArray = @();
  
  Write-HTMLParagraph -Text 'Details on all custom security roles in the site:' -Level 2 -File $FilePath
  
  foreach ($SecurityRole in $SecurityRoles)
  {
    if ($SecurityRole.NumberOfAdmins -gt 0)
    {
      $Members = $(Get-CMAdministrativeUser | Where-Object -FilterScript {$_.Roles -ilike "$($SecurityRole.RoleID)"}).LogonName
    }
    $SecRoleArray += New-Object -TypeName psobject -Property @{Name = $SecurityRole.RoleName; Description = $SecurityRole.RoleDescription; 'Copied From' = $((Get-CMSecurityRole -Id $SecurityRole.CopiedFromID).RoleName); Members = "$($Members -join '--CRLF--')"; 'Role ID' = $SecurityRole.RoleID;}
  }
  $SecRoleArray = $SecRoleArray | Select-Object -Property 'Name','Description','Copied From','Members','Role ID'
  Write-HtmlTable -InputObject $SecRoleArray -Border 1 -Level 3 -File $FilePath  
}
else
{
  Write-HTMLParagraph -Text 'There are no custom built security roles.' -Level 2 -File $FilePath
}
#endregion enumerating all custom Security roles


#region System Used Accounts

Write-Verbose "$(Get-Date):   Enumerating all used accounts"
Write-HTMLHeading -Level 2 -Text 'Configured Accounts' -File $FilePath
$Accounts = Get-CMAccount
Write-HTMLParagraph -Text 'List of all accounts used for specific tasks in the site:' -Level 2 -File $FilePath

If(-not [string]::IsNullOrEmpty($Accounts)){
    $AccountsArray = @();

    foreach ($Account in $Accounts)
    {
      $AccountsArray += New-Object -TypeName psobject -Property @{'User Name'= $Account.UserName; 'Account Usage' = if ([string]::IsNullOrEmpty($Account.AccountUsage)) {'not assigned'} else {"$($Account.AccountUsage)"}; 'Site Code' = $Account.SiteCode};
    }

    $AccountsArray = $AccountsArray | Select-Object -Property 'User Name','Account Usage','Site Code'
    Write-HtmlTable -InputObject $AccountsArray -Border 1 -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No accounts in use in this site.' -Level 3 -File $FilePath
}
#endregion System Used Accounts
Write-HtmliLink -ReturnTOC -File $FilePath

####
#region Assets and Compliance
####
Write-Verbose "$(Get-Date):   Done with Administration, next Assets and Compliance"
Write-HTMLHeading -Level 1 -PageBreak -Text 'Assets and Compliance' -File $FilePath

#region enumerating all User Collections
Write-HTMLHeading -Level 2 -Text 'Summary of User Collections' -File $FilePath
$UserCollections = Get-CMUserCollection
$BuiltinUserCollections = $UserCollections|where {$_.CollectionID -like "SMS*"}
$CustomUserCollections = $UserCollections|where {$_.CollectionID -notlike "SMS*"}
if ($ListAllInformation)
{
  $CustomUCArray = @();
  $BuiltInUCArray = @();
  Write-HTMLParagraph -Text "Configuration Manager comes with a few built-in user collections.  Any number of custom collections can be defined in the console.  Below is a summary of both types." -Level 2 -File $FilePath
  Write-HTMLHeading -Level 3 -Text 'Built-In User Collections' -File $FilePath
  Write-HTMLParagraph -Text "There are $($BuiltinUserCollections.count) built-in default user collections.  Their names and member counts are listed below:" -Level 3 -File $FilePath
  foreach ($UserCollection in $BuiltinUserCollections)
  {
    Write-Verbose "$(Get-Date):   Found Built-in User Collection: $($UserCollection.Name)"
    $BuiltInUCArray += New-Object -TypeName psobject -Property @{'Collection Name' = $UserCollection.Name; 'Collection ID' = $UserCollection.CollectionID; 'Member Count' = $UserCollection.MemberCount;};
  }
  $BuiltInUCArray = $BuiltInUCArray | Select-Object -Property 'Collection Name','Collection ID','Member Count'
  Write-HtmlTable -InputObject $BuiltInUCArray -Border 1 -Level 4 -File $FilePath

  Write-HTMLHeading -Level 3 -Text 'User Defined User Collections' -File $FilePath
  foreach ($UserCollection in $CustomUserCollections)
  {
    Write-Verbose "$(Get-Date):   Found Custom User Collection: $($UserCollection.Name)"
    $CustomUCArray += New-Object -TypeName psobject -Property @{'Collection Name' = $UserCollection.Name; 'Collection ID' = $UserCollection.CollectionID; 'Member Count' = $UserCollection.MemberCount;};
  }
  if($CustomUCArray.Count -gt 0){
      Write-HTMLParagraph -Text "There are $($CustomUserCollections.count) user defined user collections.  Their names and member counts are listed below:" -Level 3 -File $FilePath
      $CustomUCArray = $CustomUCArray | Select-Object -Property 'Collection Name','Collection ID','Member Count'
      Write-HtmlTable -InputObject $CustomUCArray -Border 1 -Level 4 -File $FilePath
  }else{
      Write-HTMLParagraph -Text "There are no user defined user collection." -Level 3 -File $FilePath
  }
}
else
{
  Write-HTMLParagraph -Text "There are $($CustomUserCollections.count) User Defined User Collections.  These are in addition to the $($BuiltinUserCollections.count) built-in default user collections." -Level 2 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all User Collections


#region enumerating all Device Collections
Write-Verbose "$(Get-Date):   Getting Device Collections."
Write-HTMLHeading -Level 2 -PageBreak -Text 'Summary of Device Collections' -File $FilePath
Write-HTMLParagraph -Text 'This section contains a brief summary of built-in device collections as well as a more detailed summary of custom device collections.' -Level 3 -File $FilePath
$DeviceCollections = Get-CMDeviceCollection
$BuiltInDeviceCollections = $DeviceCollections | where {$_.IsBuiltIn -eq $true}
$CustomDeviceCollections = $DeviceCollections | where {$_.IsBuiltIn -eq $false}
$IncUpCollCount = ($CustomDeviceCollections|where {($_.RefreshType -eq 4) -or ($_.RefreshType -eq 6)}).count
$ServiceWindowCollections = $CustomDeviceCollections|where {$_.ServiceWindowsCount -gt 0}
Write-HtmlList -InputObject "$IncUpCollCount of $($CustomDeviceCollections.Count) have incremental updates Enabled." -Description "Incremental Update Summary:" -Level 3 -File $FilePath

if ($ListAllInformation)
{
  Write-HTMLHeading -Level 3 -Text 'Built-In Device Collections' -File $FilePath
  $DevCols = @()
  foreach ($DeviceCollection in $BuiltInDeviceCollections)
  {
    Write-Verbose "$(Get-Date):   Found Built-in Device Collection: $($DeviceCollection.Name)"
    $DevCols += New-Object -TypeName psobject -Property @{'Name' = "$($DeviceCollection.Name)"; 'Collection ID' = "$($DeviceCollection.CollectionID)"; 'Member Count' = "$($DeviceCollection.MemberCount)";}
  }
  $DevCols = $DevCols | Select-Object 'Name','Collection ID','Member Count'
  Write-HTMLParagraph -Level 4 -File $FilePath -Text 'Summary of membership of the built-in device collections:'
  Write-HtmlTable -InputObject $DevCols -Border 1 -Level 5 -File $FilePath
  Write-HTMLHeading -Level 3 -Text 'User Defined Device Collections' -File $FilePath
  Write-HTMLParagraph -Level 3 -File $FilePath -Text "There are $($CustomDeviceCollections.count) custom defined collections in this site."
  Write-HTMLParagraph -Level 3 -File $FilePath -Text "There are $($ServiceWindowCollections.count) custom collections with defined service windows.  These are listed below:"
  $SWCollections = @()
  foreach ($collection in $ServiceWindowCollections){
    $LinkID = ($Collection.Name).Replace(' ','')
    $SWCollections +=  Write-HtmliLink -LinkID $LinkID -Text "$($Collection.Name) ($($Collection.CollectionID))"
  }
  If($SWCollections.count -eq 0){$SWCollections += "None Found"}
  Write-HtmlList -InputObject $SWCollections -Level 3 -File $FilePath
  foreach ($DeviceCollection in $CustomDeviceCollections)
  {
    Write-Verbose "$(Get-Date):   Found Custom Device Collection: $($DeviceCollection.Name)"
    $CollectionInfo = @()
    $CollectionName = "$($DeviceCollection.Name)"
    $CollectionDesc = "$($DeviceCollection.Comment)"
    $CollectionInfo += "Collection ID: $($DeviceCollection.CollectionID)"
    $CollectionInfo += "Total count of members: $($DeviceCollection.MemberCount)"
    $CollectionInfo += "Limited to Device Collection: $($DeviceCollection.LimitToCollectionName) ($($DeviceCollection.LimitToCollectionID))"
    Switch ($DeviceCollection.RefreshType)
    {
        1 {$UpdateSchedule = "No schedule configured"}
        2 {$UpdateSchedule = "Full update schedule only"}
        4 {$UpdateSchedule = "Incremental update only"}
        6 {$UpdateSchedule = "Full and Incremental updates configured"}
    }
    $CollectionInfo += "Selected Update Schedule: $UpdateSchedule"
    Write-HTMLHeading -Level 4 -Text $CollectionName -File $FilePath -ExcludeTOC
    If ($CollectionDesc){
        Write-HtmlList -InputObject $CollectionInfo -Description "Description: $CollectionDesc" -Level 4 -File $FilePath
    }else{
        Write-HtmlList -InputObject $CollectionInfo -Level 4 -File $FilePath
    }
    If ($DeviceCollection.ServiceWindowsCount -gt 0) {
        $ServiceWindows = Get-CMMaintenanceWindow -CollectionId $DeviceCollection.CollectionID
        Write-Verbose "$(Get-Date):   Enumerating Maintenance Windows for collection: $($DeviceCollection.Name)"
        $ServiceWindowArray = @()
        foreach ($ServiceWindow in $ServiceWindows)
            {
                $SWName = $ServiceWindow.Name
                $Schedule = Convert-CMSchedule -ScheduleString $ServiceWindow.ServiceWindowSchedules
                $StartTime = $Schedule.StartTime
                $HourLength = $Schedule.HourDuration
                $MinuteLength = $Schedule.MinuteDuration
                $Duration = "$($HourLength):$("{0:D2}" -f $MinuteLength)"
                switch ($ServiceWindow.IsGMT)
                    {
                        true {$UTCTime = 'Yes'}
                        false {$UTCTime = 'No'}
                    }
                switch ($ServiceWindow.ServiceWindowType)
                    {
                        0 {$WindowType = 'Task Sequences'}
                        1 {$WindowType = 'All Deployments'}
                        4 {$WindowType = 'Software Updates'}
                    }
                switch ($ServiceWindow.RecurrenceType)
                    {
                        1 {$WindowRecurence = "None"}
                        2 {
                            if ($Schedule.DaySpan -eq '1') {
                                $WindowRecurence = 'Daily'
                            } else {
                                $WindowRecurence = "Every $($Schedule.DaySpan) days"
                            }
                            }
                        3 {                                              
                            $WindowRecurence = "Every $($Schedule.ForNumberofWeeks) week(s) on $(Convert-WeekDay $Schedule.Day)"
                            }
                        4 {
                            switch ($Schedule.weekorder)
                                {
                                    0 {$order = 'last'}
                                    1 {$order = 'first'}
                                    2 {$order = 'second'}
                                    3 {$order = 'third'}
                                    4 {$order = 'fourth'}
                                }
                            $WindowRecurence = "Every $($Schedule.ForNumberofMonths) month(s) on every $($order) $(Convert-WeekDay $Schedule.Day)"
                            }

                        5 {
                            if ($Schedule.MonthDay -eq '0'){
                                $DayOfMonth = 'the last day of the month'
                            } else {
                                $DayOfMonth = "day $($Schedule.MonthDay)"
                            }
                            $WindowRecurence = "Every $($Schedule.ForNumberofMonths) month(s) on $($DayOfMonth)."
                            }
                    }
                switch ($ServiceWindow.IsEnabled)
                    {
                        true {$WindowEnabled = 'Yes'}
                        false {$WindowEnabled = 'No'}
                    }
                $ServiceWindowArray += New-Object -TypeName psobject -Property @{'Name' = $SWName; 'Start Time' = $StartTime; 'UTC' = $UTCTime; 'Duration' = $Duration; 'Recurance' = $WindowRecurence; 'Type' = $WindowType; 'Enabled' = $WindowEnabled}
            }
        Write-HTMLParagraph -Level 4 -File $FilePath -Text 'Collection Maintenance Windows:'
        $ServiceWindowArray = $ServiceWindowArray | Select-Object 'Name','Start Time','UTC','Duration','Recurance','Type','Enabled'
        Write-HtmlTable -InputObject $ServiceWindowArray -Border 1 -Level 5 -File $FilePath
    } else {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text 'No maintenance windows configured on this collection.'
    }
        ### enumerating the Collection Membership Rules
        Write-HTMLParagraph -Level 4 -File $FilePath -Text 'Collection Membership Rules:'
        $QueryRules = $Null
        $DirectRules = $Null
        $IncludeRules = $Null
        $CollectionRules = $DeviceCollection.CollectionRules #just for Direct and Query

        $Collection = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_Collection WHERE CollectionID = '$($DeviceCollection.CollectionID)'" -ComputerName $SMSProvider
        [wmi]$Collection = $Collection.__PATH
                    
        $OtherCollectionRules = $Collection.CollectionRules
        try {
            $DirectRules = $CollectionRules | where {$_.ResourceID} -ErrorAction SilentlyContinue
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Verbose "$(Get-Date):   Collection Rule info not found"
        }
        try {
            $QueryRules = $CollectionRules | where {$_.QueryExpression} -ErrorAction SilentlyContinue                            
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Verbose "$(Get-Date):   Collection Rule info not found"
        }
        try {
            $IncludeRules = $OtherCollectionRules | where {$_.IncludeCollectionID} -ErrorAction SilentlyContinue
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Verbose "$(Get-Date):   Collection Rule info not found"
        }

        if (-not [string]::IsNullOrEmpty($QueryRules)) {
            $QueryRulesArray = @();
            foreach ($QueryRule in $QueryRules) {
                $QueryRulesArray += New-Object -TypeName psobject -Property @{'Query Name'= $QueryRule.RuleName; 'Query Expression' = $($QueryRule.QueryExpression -replace ',',', ')}
            }
            Write-HtmlTable -InputObject $QueryRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (-not [string]::IsNullOrEmpty($DirectRules)) {
            $DirectRulesArray = @();
            foreach ($DirectRule in $DirectRules) {
                $DirectRulesArray += New-Object -TypeName psobject -Property @{'Resource Name'= $DirectRule.RuleName; 'Resource ID' = $DirectRule.ResourceId}
            }
            Write-HtmlTable -InputObject $DirectRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (-not [String]::IsNullOrEmpty($IncludeRules)) {
            $IncludeRulesArray = @()
            foreach ($IncludeRule in $IncludeRules) {
                $IncludeRulesArray += New-Object -TypeName psobject -Property @{'Collection Name'= $IncludeRule.RuleName; 'Collection ID' = $IncludeRule.IncludeCollectionId}
            }
            Write-HtmlTable -InputObject $IncludeRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (([String]::IsNullOrEmpty($IncludeRules)) -and ([string]::IsNullOrEmpty($DirectRules)) -and ([string]::IsNullOrEmpty($QueryRules))){
        Write-HTMLParagraph -Level 5 -File $FilePath -Text 'No collection membership rules defined.'
        }
    }
}else{
  Write-HTMLParagraph -Text "There are $($CustomDeviceCollections.count) User Defined Device collections.  These are in addition to the $($BuiltInDeviceCollections.count) built-in default device collections." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all Device Collections


#region enumerating all Compliance Settings
Write-Verbose "$(Get-Date):   Working on Compliance Settings"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Compliance Settings' -File $FilePath
Write-HTMLParagraph -Text 'This section contains a summary of all configuration items, baselines, Settings, Conditional Access, and other configurable resources.' -Level 3 -File $FilePath
#region enumerating all Configuration Items and baselines.
Write-HTMLHeading -Level 3 -Text 'Configuration Items' -File $FilePath
$CIs = Get-CMConfigurationItem -Fast | Where {$_.IsUserDefined -eq "true"}
if(-not [string]::IsNullOrEmpty($CIs)){
    $CIsArray = @()
    foreach ($CI in $CIs){
        $CIsArray += New-Object -TypeName psobject -Property @{'Name' = $CI.LocalizedDisplayName; 'Last modified' = $CI.DateLastModified; 'Last modified by' = $CI.LastModifiedBy; 'CI ID' = $CI.CI_ID}
    }
    $CIsArray = $CIsArray | Select-Object 'Name','Last modified','Last modified by','CI ID'
    Write-HtmlTable -InputObject $CIsArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'There are no Configuration Items configured.' -Level 3 -File $FilePath
}
Write-HTMLHeading -Level 3 -Text 'Configuration Baselines' -File $FilePath
$CBs = Get-CMBaseline | Where {$_.IsUserDefined -eq "true"}
if ($CBs){
    $CBsArray = @()
    foreach ($CB in $CBs){
        $CBsArray += New-Object -TypeName psobject -Property @{'Baseline Name' = $CB.LocalizedDisplayName; 'Last modified' = $CB.DateLastModified; 'Last modified by' = $CB.LastModifiedBy; 'Baseline ID' = $CB.CI_ID}
    }
    $CBsArray = $CBsArray|Select-Object 'Baseline Name','Last modified','Last modified by','Baseline ID'
    Write-HtmlTable -InputObject $CIsArray -Border 1 -Level 4 -File $FilePath
} else {
    Write-HTMLParagraph -Text 'There are no Configuration Baselines configured.' -Level 3 -File $FilePath
}
#endregion enumerating all Configuration Items and baselines.

#region enumerating Configuration Policies...
Write-Verbose "$(Get-Date):   Working on Configuration Policies..."
$CMPolicies=@()
$CMPolicies=Get-CMConfigurationPolicy -fast | select CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-Verbose "$(Get-Date):   $($CMPolicies.count) Configuration Policies found."

$ComplianceSettings = @()
$RemoteSettings = @()
$UserStateSettings = @()
$CommSettings = @()
$TandCSettings = @()
$EdUpgradeSettings = @()
$WinHelloSettings = @()
$WiFiProfileSettings = @()
$VpnSettings = @()
$CertSettings = @()

foreach ($CMPolicy in $CMPolicies){
    Switch ($CMPolicy){
        {'SettingsAndPolicy:SMS_CompliancePolicySettings' -in $_.CategoryInstance_UniqueIDs}{
            $ComplianceSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_RemoteConnectionSettings' -in $_.CategoryInstance_UniqueIDs}{
            $RemoteSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_UserStateManagementSettings' -in $_.CategoryInstance_UniqueIDs}{
            $UserStateSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_CommunicationsProvisioningSettings' -in $_.CategoryInstance_UniqueIDs}{
            $CommSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_TermsAndConditionsSettings' -in $_.CategoryInstance_UniqueIDs}{
            $TandCSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_EditionUpgradeSettings' -in $_.CategoryInstance_UniqueIDs}{
            $EdUpgradeSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_PassportForWorkProfileSettings' -in $_.CategoryInstance_UniqueIDs}{
            $WinHelloSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_WirelessProfileSettings' -in $_.CategoryInstance_UniqueIDs}{
            $WiFiProfileSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_VpnConnectionSettings' -in $_.CategoryInstance_UniqueIDs}{
            $VpnSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_TrustedRootCertificateSettings' -in $_.CategoryInstance_UniqueIDs}{
            $CertSettings += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
    }
}
Write-HTMLHeading -Level 3 -Text 'User Data and Profiles' -File $FilePath
if ($UserStateSettings.count -gt 0) {
    $UserStateSettings = $UserStateSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $UserStateSettings -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No User Data and Profiles defined in site.' -Level 4 -File $FilePath
}
Write-HTMLHeading -Level 3 -Text 'Remote Connection Profiles' -File $FilePath
if ($RemoteSettings.count -gt 0) {
    $RemoteSettings = $RemoteSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $RemoteSettings -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Remote Connection Profiles defined in site.' -Level 4 -File $FilePath
}
Write-HTMLHeading -Level 3 -Text 'Compliance Policies' -File $FilePath
if ($ComplianceSettings.count -gt 0) {
    $ComplianceSettings = $ComplianceSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $ComplianceSettings -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Compliance Policies defined in site.' -Level 4 -File $FilePath
}

Write-HTMLHeading -Level 3 -Text 'Company Resource Access' -File $FilePath

Write-HTMLHeading -Level 4 -Text 'Certificate Profiles' -File $FilePath
if ($CertSettings.count -gt 0) {
    $CertSettings = $CertSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $CertSettings -Border 1 -Level 5 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Certificate Profiles defined in site.' -Level 5 -File $FilePath
}
Write-HTMLHeading -Level 4 -Text 'Email Profiles' -File $FilePath
if ($CommSettings.count -gt 0) {
    $CommSettings = $CommSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $CommSettings -Border 1 -Level 5 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Email Profiles defined in site.' -Level 5 -File $FilePath
}
Write-HTMLHeading -Level 4 -Text 'VPN Profiles' -File $FilePath
if ($VpnSettings.count -gt 0) {
    $VpnSettings = $VpnSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $VpnSettings -Border 1 -Level 5 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No VPN Profiles defined in site.' -Level 5 -File $FilePath
}
Write-HTMLHeading -Level 4 -Text 'Wi-Fi Profiles' -File $FilePath
if ($WiFiProfileSettings.count -gt 0) {
    $WiFiProfileSettings = $WiFiProfileSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $WiFiProfileSettings -Border 1 -Level 5 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Wi-Fi Profiles defined in site.' -Level 5 -File $FilePath
}
Write-HTMLHeading -Level 4 -Text 'Windows Hello for Business Profiles' -File $FilePath
if ($WinHelloSettings.count -gt 0) {
    $WinHelloSettings = $WinHelloSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $WinHelloSettings -Border 1 -Level 5 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Windows Hello for Business Profiles defined in site.' -Level 5 -File $FilePath
}
Write-HTMLHeading -Level 3 -Text 'Terms and Conditions' -File $FilePath
if ($TandCSettings.count -gt 0) {
    $TandCSettings = $TandCSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $TandCSettings -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Terms and Conditions defined in site.' -Level 4 -File $FilePath
}
Write-HTMLHeading -Level 3 -Text 'Windows 10 Edition Upgrades' -File $FilePath
if ($EdUpgradeSettings.count -gt 0) {
    $EdUpgradeSettings = $EdUpgradeSettings | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $EdUpgradeSettings -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Windows 10 Edition Upgrades defined in site.' -Level 4 -File $FilePath
}

#endregion enumerating Configuration Policies.
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all Compliance Settings

#region Endpoint Protection
Write-Verbose "$(Get-Date):   Working on Endpoint Protection"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Endpoint Protection' -File $FilePath
Write-HTMLParagraph -Text 'This section contains a summary of all Endpoint Security configuration settings.  This includes System Center Endpoint Protection (Antimalware), Firewall, Windows Defender ATP, and Device Guard Policies.' -Level 3 -File $FilePath
#region Antimalware
Write-HTMLHeading -Level 3 -Text 'Antimalware Policies' -File $FilePath
if (-not ($(Get-CMEndpointProtectionPoint) -eq $Null)){
    $AntiMalwarePolicies = Get-CMAntimalwarePolicy
    if (-not [string]::IsNullOrEmpty($AntiMalwarePolicies)){
        foreach ($AntiMalwarePolicy in $AntiMalwarePolicies){
                if ($AntiMalwarePolicy.Name -eq 'Default Client Antimalware Policy'){
                    $AgentConfig = $AntiMalwarePolicy.AgentConfiguration
                    Write-HTMLHeading -Level 4 -Text "$($AntiMalwarePolicy.Name)" -File $FilePath
                    Write-HTMLParagraph -Text "Description: $($AntiMalwarePolicy.Description)" -Level 4 -File $FilePath
                    $listTitle = 'Scheduled Scans'
                    $listArray = @()
                    $listArray += "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                    if ($AgentConfig.EnableScheduledScan){
                        switch ($AgentConfig.ScheduledScanType)
                            {
                                1 { $ScheduledScanType = 'Quick Scan' }
                                2 { $ScheduledScanType = 'Full Scan' }
                            }
                        $listArray += "Scan type: $($ScheduledScanType)"
                        $listArray += "Scan day: $(Convert-WeekDay $AgentConfig.ScheduledScanWeekDay)"
                        $listArray += "Scan time: $(Convert-Time -time $AgentConfig.ScheduledScanTime)"
                        $listArray += "Run a daily quick scan on client computers: $($AgentConfig.EnableQuickDailyScan)"
                        $listArray += "Daily quick scan schedule time: $(Convert-Time -time $AgentConfig.ScheduledScanQuickTime)"
                        $listArray += "Check for the latest definition updates before running a scan: $($AgentConfig.CheckLatestDefinition)"
                        $listArray += "Start a scheduled scan only when the computer is idle: $($AgentConfig.ScanWhenClientNotInUse)"
                        $listArray += "Force a scan of the selected scan type if client computer is offline during two or more scheduled scans: $($AgentConfig.EnableCatchupScan)"
                        $listArray += "Limit CPU usage during scans to (%): $($AgentConfig.LimitCPUUsage)"
                    }
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Scan settings'
                    $listArray = @()
                    $listArray += "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                    $listArray += "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                    $listArray += "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                    $listArray += "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                    $listArray += "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                    switch ($AgentConfig.ScheduledScanUserControl)
                        {
                            0 { $UserControl = 'No control' }
                            1 { $UserControl = 'Scan time only' }
                            2 { $UserControl = 'Full control' }
                        }
                    $listArray += "User control of scheduled scans: $UserControl"
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Default Actions'
                    $listArray = @()
                    switch ($AgentConfig.DefaultActionSevere)
                        {
                            0 { $Action = 'Recommended' }
                            2 { $Action = 'Quarantine' }
                            3 { $Action = 'Remove' }
                            6 { $Action = 'Allow' }
                        }
                    $listArray += "Severe threats: $Action"
                    switch ($AgentConfig.DefaultActionHigh)
                        {
                            0 { $Action = 'Recommended' }
                            2 { $Action = 'Quarantine' }
                            3 { $Action = 'Remove' }
                            6 { $Action = 'Allow' }
                        }
                    $listArray += "High threats: $Action"
                    switch ($AgentConfig.DefaultActionMedium)
                        {
                            0 { $Action = 'Recommended' }
                            2 { $Action = 'Quarantine' }
                            3 { $Action = 'Remove' }
                            6 { $Action = 'Allow' }
                        }
                    $listArray += "Medium threats: $Action"
                    switch ($AgentConfig.DefaultActionLow)
                        {
                            0 { $Action = 'Recommended' }
                            2 { $Action = 'Quarantine' }
                            3 { $Action = 'Remove' }
                            6 { $Action = 'Allow' }
                        }
                    $listArray += "Low threats: $Action"
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Real-time protection'
                    $listArray = @()
                    $listArray += "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                    $listArray += "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                    switch ($AgentConfig.RealtimeScanOption){
                        0 { $SysFiles = 'Scan incoming and outgoing files' }
                        1 { $SysFiles = 'Scan incoming files only' }
                        2 { $SysFiles = 'Scan outgoing files only' }
                    }
                    $listArray += "Scan system files: $SysFiles"
                    $listArray += "Scan all downloaded files and enable exploit protection for Internet Explorer: $($AgentConfig.ScannAllDownloaded)"
                    $listArray += "Enable behavior monitoring: $($AgentConfig.UseBehaviorMonitor)"
                    $listArray += "Enable protection against network-based exploits: $($AgentConfig.NetworkProtectionAgainstExploits)"
                    $listArray += "Allow users on client computers to configure real-time protection settings: $($AgentConfig.AllowClientUserConfigRealtime)"
                    $listArray += "Enable protection against Potentially Unwanted Applications at download and prior to installation: $($AgentConfig.EnablePUAProtection)"
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Exclusion settings'
                    $listArray = @()
                    $filesArray = @()
                    foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths){
                        $filesArray += "$($ExcludedFileFolder)"
                    }
                    $listArray += Write-HtmlList -Description 'Excluded files and folders:' -InputObject $filesArray -Level 1
                    $filesArray = @()
                    foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes){
                        $filesArray += "$($ExcludedFileType)"
                    }
                    $listArray += Write-HtmlList -Description 'Excluded file types:' -InputObject $filesArray -Level 1
                    $ProcessArray = @()
                    foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses){
                        $ProcessArray += "$($ExcludedProcess)"
                    }
                    $listArray += Write-HtmlList -Description 'Excluded processes:' -InputObject $filesArray -Level 1
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Advanced'
                    $listArray = @()
                    $listArray += "Create a system restore point before computers are cleaned: $($AgentConfig.CreateSystemRestorePointBeforeClean)"
                    $listArray += "Disable the client user interface: $($AgentConfig.DisableClientUI)"
                    $listArray += "Show notifications messages on the client computer when the user needs to run a full scan, update definitions, or run Windows Defender Offline: $($AgentConfig.ShowNotificationMessages)"
                    $listArray += "Delete quarantined files after (days): $($AgentConfig.DeleteQuarantinedFilesPeriod)"
                    $listArray += "Allow users to configure the setting for quarantined file deletion: $($AgentConfig.AllowUserConfigQuarantinedFileDeletionPeriod)"
                    $listArray += "Allow users to exclude file and folders, file types and processes: $($AgentConfig.AllowUserAddExcludes)"
                    $listArray += "Allow all users to view the full History results: $($AgentConfig.AllowUserViewHistory)"
                    $listArray += "Enable reparse point scanning: $($AgentConfig.EnableReparsePointScanning)"
                    $listArray += "Randomize scheduled scan and definition update start time (within 30 minutes): $($AgentConfig.RandomizeScheduledScanStartTime)"
                    $listArray += "Enable auto sample file submission to help Microsoft determine whether certain detected items are malicious: $($AgentConfig.EnableAutoSampleSubmission)"
                    $listArray += "Allow users to modify auto sample file submission settings: $($AgentConfig.AllowUserConfigAutoSampleSubmission)"
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Threat overrides'
                    $listArray = @()
                    if (-not [string]::IsNullOrEmpty($AgentConfig.ThreatName)){
                        $listArray +='Threat name and override action: Threats specified'
                    }else{
                        $listArray +='Threat name and override action: none specified'
                    }
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Cloud Protection Service'
                    $listArray = @()
                    switch ($AgentConfig.JoinSpyNet){
                        0 { $CPSLevel =  'Do not join MAPS' }
                        1 { $CPSLevel =  'Basic membership' }
                        2 { $CPSLevel =  'Advanced membership' }
                    }
                    $listArray += "Cloud Protection Service membership type: $CPSLevel"
                    $listArray += "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"
                    switch ($AgentConfig.JoinSpyNet){
                        0 { $BSFLevel =  'Normal' }
                        1 { $BSFLevel =  'High' }
                        2 { $BSFLevel =  'High with extra protection' }
                        3 { $BSFLevel =  'Block unknown programs' }
                    }
                    $listArray += "Level for blocking suspicious files: $BSFLevel"
                    $listArray += "Allow extended cloud check to block and scan suspicious files for up to (seconds): $($AgentConfig.CloudTimeout)"
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    $listTitle = 'Definition Updates'
                    $listArray = @()
                    $listArray += "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                    $listArray += "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                    $listArray += "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                    $FallbackArray = @()
                    foreach ($Fallback in $AgentConfig.FallbackOrder){
                        $FallbackArray += "$($Fallback)"
                    }
                    $listArray += Write-HtmlList -Description 'Set sources and order for Endpoint Protection definition updates:' -InputObject $FallbackArray -Level 1
                    $listArray += "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                    $UNCShareArray = @()
                    foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources){
                        $UNCShareArray += "$($UNCShare)"
                    }
                    if ($UNCShareArray.count -gt 0){
                        $listArray += Write-HtmlList -Description 'If UNC file shares are selected as a definition update source, specify the UNC paths:' -InputObject $UNCShareArray -Level 1
                    }else{
                        $listArray += 'If UNC file shares are selected as a definition update source, specify the UNC paths: None'
                    }             
                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                }else{
                    $AgentConfig_custom = $AntiMalwarePolicy.AgentConfigurations
                    Write-HTMLHeading -Level 4 -Text "$($AntiMalwarePolicy.Name)" -File $FilePath
                    If("$($AntiMalwarePolicy.Description)" -ne ""){
                        Write-HTMLParagraph -Text "Description: $($AntiMalwarePolicy.Description)" -Level 4 -File $FilePath
                    }
                    foreach ($Agentconfig in $AgentConfig_custom){
                        switch ($AgentConfig.AgentID){
                            201 {
                                    $listTitle = 'Scheduled Scans'
                                    $listArray = @()
                                    $listArray += "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                    if ($AgentConfig.EnableScheduledScan){
                                        switch ($AgentConfig.ScheduledScanType)
                                            {
                                                1 { $ScheduledScanType = 'Quick Scan' }
                                                2 { $ScheduledScanType = 'Full Scan' }
                                            }
                                        $listArray += "Scan type: $($ScheduledScanType)"
                                        $listArray += "Scan day: $(Convert-WeekDay $AgentConfig.ScheduledScanWeekDay)"
                                        $listArray += "Scan time: $(Convert-Time -time $AgentConfig.ScheduledScanTime)"
                                        $listArray += "Run a daily quick scan on client computers: $($AgentConfig.EnableQuickDailyScan)"
                                        $listArray += "Daily quick scan schedule time: $(Convert-Time -time $AgentConfig.ScheduledScanQuickTime)"
                                        $listArray += "Check for the latest definition updates before running a scan: $($AgentConfig.CheckLatestDefinition)"
                                        $listArray += "Start a scheduled scan only when the computer is idle: $($AgentConfig.ScanWhenClientNotInUse)"
                                        $listArray += "Force a scan of the selected scan type if client computer is offline during two or more scheduled scans: $($AgentConfig.EnableCatchupScan)"
                                        $listArray += "Limit CPU usage during scans to (%): $($AgentConfig.LimitCPUUsage)"
                                    }
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            202 {
                                    $listTitle = 'Default Actions'
                                    $listArray = @()
                                    switch ($AgentConfig.DefaultActionSevere)
                                        {
                                            0 { $Action = 'Recommended' }
                                            2 { $Action = 'Quarantine' }
                                            3 { $Action = 'Remove' }
                                            6 { $Action = 'Allow' }
                                        }
                                    $listArray += "Severe threats: $Action"
                                    switch ($AgentConfig.DefaultActionHigh)
                                        {
                                            0 { $Action = 'Recommended' }
                                            2 { $Action = 'Quarantine' }
                                            3 { $Action = 'Remove' }
                                            6 { $Action = 'Allow' }
                                        }
                                    $listArray += "High threats: $Action"
                                    switch ($AgentConfig.DefaultActionMedium)
                                        {
                                            0 { $Action = 'Recommended' }
                                            2 { $Action = 'Quarantine' }
                                            3 { $Action = 'Remove' }
                                            6 { $Action = 'Allow' }
                                        }
                                    $listArray += "Medium threats: $Action"
                                    switch ($AgentConfig.DefaultActionLow)
                                        {
                                            0 { $Action = 'Recommended' }
                                            2 { $Action = 'Quarantine' }
                                            3 { $Action = 'Remove' }
                                            6 { $Action = 'Allow' }
                                        }
                                    $listArray += "Low threats: $Action"
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            203 {
                                    $listTitle = 'Exclusion settings'
                                    $listArray = @()
                                    $filesArray = @()
                                    foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths){
                                        $filesArray += "$($ExcludedFileFolder)"
                                    }
                                    $listArray += Write-HtmlList -Description 'Excluded files and folders:' -InputObject $filesArray -Level 1
                                    $filesArray = @()
                                    foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes){
                                        $filesArray += "$($ExcludedFileType)"
                                    }
                                    $listArray += Write-HtmlList -Description 'Excluded file types:' -InputObject $filesArray -Level 1
                                    $ProcessArray = @()
                                    foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses){
                                        $ProcessArray += "$($ExcludedProcess)"
                                    }
                                    $listArray += Write-HtmlList -Description 'Excluded processes:' -InputObject $filesArray -Level 1
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            204 {
                                    $listTitle = 'Real-time protection'
                                    $listArray = @()
                                    $listArray += "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                    $listArray += "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                                    switch ($AgentConfig.RealtimeScanOption){
                                        0 { $SysFiles = 'Scan incoming and outgoing files' }
                                        1 { $SysFiles = 'Scan incoming files only' }
                                        2 { $SysFiles = 'Scan outgoing files only' }
                                    }
                                    $listArray += "Scan system files: $SysFiles"
                                    $listArray += "Scan all downloaded files and enable exploit protection for Internet Explorer: $($AgentConfig.ScannAllDownloaded)"
                                    $listArray += "Enable behavior monitoring: $($AgentConfig.UseBehaviorMonitor)"
                                    $listArray += "Enable protection against network-based exploits: $($AgentConfig.NetworkProtectionAgainstExploits)"
                                    $listArray += "Allow users on client computers to configure real-time protection settings: $($AgentConfig.AllowClientUserConfigRealtime)"
                                    $listArray += "Enable protection against Potentially Unwanted Applications at download and prior to installation: $($AgentConfig.EnablePUAProtection)"
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            205 {
                                    $listTitle = 'Advanced'
                                    $listArray = @()
                                    $listArray += "Create a system restore point before computers are cleaned: $($AgentConfig.CreateSystemRestorePointBeforeClean)"
                                    $listArray += "Disable the client user interface: $($AgentConfig.DisableClientUI)"
                                    $listArray += "Show notifications messages on the client computer when the user needs to run a full scan, update definitions, or run Windows Defender Offline: $($AgentConfig.ShowNotificationMessages)"
                                    $listArray += "Delete quarantined files after (days): $($AgentConfig.DeleteQuarantinedFilesPeriod)"
                                    $listArray += "Allow users to configure the setting for quarantined file deletion: $($AgentConfig.AllowUserConfigQuarantinedFileDeletionPeriod)"
                                    $listArray += "Allow users to exclude file and folders, file types and processes: $($AgentConfig.AllowUserAddExcludes)"
                                    $listArray += "Allow all users to view the full History results: $($AgentConfig.AllowUserViewHistory)"
                                    $listArray += "Enable reparse point scanning: $($AgentConfig.EnableReparsePointScanning)"
                                    $listArray += "Randomize scheduled scan and definition update start time (within 30 minutes): $($AgentConfig.RandomizeScheduledScanStartTime)"
                                    $listArray += "Enable auto sample file submission to help Microsoft determine whether certain detected items are malicious: $($AgentConfig.EnableAutoSampleSubmission)"
                                    $listArray += "Allow users to modify auto sample file submission settings: $($AgentConfig.AllowUserConfigAutoSampleSubmission)"
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            206 {
                                    $listTitle = 'Threat overrides'
                                    $listArray = @()
                                    if (-not [string]::IsNullOrEmpty($AgentConfig.ThreatName)){
                                        $listArray +='Threat name and override action: Threats specified'
                                    }else{
                                        $listArray +='Threat name and override action: none specified'
                                    }
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            207 {
                                    $listTitle = 'Cloud Protection Service'
                                    $listArray = @()
                                    switch ($AgentConfig.JoinSpyNet){
                                        0 { $CPSLevel =  'Do not join MAPS' }
                                        1 { $CPSLevel =  'Basic membership' }
                                        2 { $CPSLevel =  'Advanced membership' }
                                    }
                                    $listArray += "Cloud Protection Service membership type: $CPSLevel"
                                    $listArray += "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"
                                    switch ($AgentConfig.JoinSpyNet){
                                        0 { $BSFLevel =  'Normal' }
                                        1 { $BSFLevel =  'High' }
                                        2 { $BSFLevel =  'High with extra protection' }
                                        3 { $BSFLevel =  'Block unknown programs' }
                                    }
                                    $listArray += "Level for blocking suspicious files: $BSFLevel"
                                    $listArray += "Allow extended cloud check to block and scan suspicious files for up to (seconds): $($AgentConfig.CloudTimeout)"
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            208 {
                                    $listTitle = 'Definition Updates'
                                    $listArray = @()
                                    $listArray += "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                                    $listArray += "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                                    $listArray += "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                                    $FallbackArray = @()
                                    foreach ($Fallback in $AgentConfig.FallbackOrder){
                                        $FallbackArray += "$($Fallback)"
                                    }
                                    $listArray += Write-HtmlList -Description 'Set sources and order for Endpoint Protection definition updates:' -InputObject $FallbackArray -Level 1
                                    $listArray += "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                                    $UNCShareArray = @()
                                    foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources){
                                        $UNCShareArray += "$($UNCShare)"
                                    }
                                    if ($UNCShareArray.count -gt 0){
                                        $listArray += Write-HtmlList -Description 'If UNC file shares are selected as a definition update source, specify the UNC paths:' -InputObject $UNCShareArray -Level 1
                                    }else{
                                        $listArray += 'If UNC file shares are selected as a definition update source, specify the UNC paths: None'
                                    }
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                            209 {
                                    $listTitle = 'Scan settings'
                                    $listArray = @()
                                    $listArray += "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                                    $listArray += "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                                    $listArray += "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                                    $listArray += "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                                    $listArray += "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                                    switch ($AgentConfig.ScheduledScanUserControl)
                                        {
                                            0 { $UserControl = 'No control' }
                                            1 { $UserControl = 'Scan time only' }
                                            2 { $UserControl = 'Full control' }
                                        }
                                    $listArray += "User control of scheduled scans: $UserControl"
                                    Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                                }
                        }
                    }
                }
            }
    }else{
        Write-HTMLParagraph -Text 'There are no Anti Malware Policies configured.' -Level 3 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text 'There is no Endpoint Protection Point enabled in this site.' -Level 3 -File $FilePath
}
#endregion Antimalware
#region firewall and Device Guard
$FWPolicies = Get-CMConfigurationPolicy -Fast | where {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_FirewallSettings' -or $_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_DeviceGuardSettings'} | select CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Windows Defender Firewall Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($FWPolicies)) {
    $FWArray = @()
    foreach ($FWP in $FWPolicies){
        $FWArray += New-Object -TypeName psobject -Property @{'Name'=$FWP.LocalizedDisplayName;'Modified By'=$FWP.LastModifiedBy;'Modified'=$FWP.DateLastModified;'Deployed'=$FWP.IsAssigned}
    }
    $FWArray = $FWArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $FWArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No firewall policies defined in site.' -Level 4 -File $FilePath
}
$DeviceGuardPolicies = Get-CMConfigurationPolicy -Fast | where {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_DeviceGuardSettings'} | select CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Device Guard Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($DeviceGuardPolicies)) {
    $DeviceGuardArray = @()
    foreach ($DGP in $DeviceGuardPolicies){
        $DeviceGuardArray += New-Object -TypeName psobject -Property @{'Name'=$DGP.LocalizedDisplayName;'Modified By'=$DGP.LastModifiedBy;'Modified'=$DGP.DateLastModified;'Deployed'=$DGP.IsAssigned}
    }
    $DeviceGuardArray = $DeviceGuardArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $DeviceGuardArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Device Guard policies defined in site.' -Level 4 -File $FilePath
}
#endregion firewall and Device Guard
#region Windows Defender ATP
    #TBD
#endregion Windows Defender ATP
#endregion Endpoint Protection

#region Corporate-owned Devices
#region iOS Enrollment Profiles
    #TBD
#endregion iOS Enrollment Profiles
#region Windows Enrollment Profiles
    #TBD
#endregion Windows Enrollment Profiles
#endregion Corporate-owned Devices

Write-HtmliLink -ReturnTOC -File $FilePath
Write-Verbose "$(Get-Date):   Done with Assets and Compliance, next Software Library"

####
#region Software Library
####

Write-HTMLHeading -Level 1 -PageBreak -Text 'Software Library' -File $FilePath

#region Application Management
Write-HTMLHeading -Level 2 -PageBreak -Text 'Application Management' -File $FilePath

#region Applications
Write-Verbose "$(Get-Date):   Processing CM Appications."
Write-HTMLHeading -Level 3 -Text 'Applications' -File $FilePath
#$Applications = Get-WmiObject -Class sms_applicationlatest -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider
#Get-CMApplication | select LocalizedDisplayName,LocalizedDescription,Manufacturer,SoftwareVersion,PackageID,ISExpired,ISDeployed,NumberOfDeploymentTypes
$Applications = Get-CMApplication
$Applications = $Applications|sort LocalizedDisplayName
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($Applications)) {
        Write-HTMLParagraph -Text "Below are a summary of all $($Applications.Count) application installers defined in this site. These are applications that are installed with the configuration manager application model.  Packages are covered later in the documentation." -Level 3 -File $FilePath
        foreach ($App in $Applications) {
            Write-Verbose "$(Get-Date):   Found App: $($App.LocalizedDisplayName)"
            Write-HTMLHeading -Level 4 -Text "$($App.LocalizedDisplayName)" -File $FilePath
            $AppList = @()
            if ($App.LocalizedDescription -ne ""){
                $ListDescription = "Description: $($App.LocalizedDescription)"
            }
            $AppList += "Created by: $($App.CreatedBy)"
            $AppList += "Date created: $($App.DateCreated)"
            $AppList += "Publisher: $($App.Manufacturer)"
            $AppList += "Software Version: $($App.SoftwareVersion)"
            $AppList += "CM Package ID: $($App.PackageID)"
            $AppList += "Enabled: $($App.ISEnabled)"
            $AppList += "Deployed: $($App.ISDeployed)"
            If ($ListDescription -ne "") {
                Write-HtmlList -InputObject $AppList -Description $ListDescription -Level 5 -File $FilePath
            }Else{
                Write-HtmlList -InputObject $AppList -Level 5 -File $FilePath
            }
            $ListDescription = ""
            Write-Verbose "$(Get-Date):   Processing deployment types for: $($App.LocalizedDisplayName)"
            $DTs = Get-CMDeploymentType -ApplicationName "$($App.LocalizedDisplayName)"
            [xml]$PackageXML = $App.SDMPackageXML
            if (-not [string]::IsNullOrEmpty($DTs)) {
                $DTsArray = @()
                foreach ($DT in $DTs) {
                    #$xmlDT = [xml]$DT.SDMPackageXML
                    foreach ($xl in $PackageXML.AppMgmtDigest.DeploymentType) {
                        if ($dt.ContentId -eq $xl.Installer.Contents.Content.ContentId) {
                            Write-Verbose "$(Get-Date):   Found Deployment Type:  $($xl.Title.'#text')"
                            $Content = "$($xl.Installer.contents.content.location)"
                            if (Test-Path "filesystem::$Content" -ErrorAction SilentlyContinue){
                                $Verified = "Verified"
                            }else{
                                $Verified = "Unverified!"
                            }
                            $InstallCL = "$($xl.installer.customdata.installcommandline)"
                            $UninstallCL = "$($xl.installer.customdata.uninstallcommandline)"
                        }
                        if ($xl.Technology -eq "Deeplink"){ #this is a windows store app.  There is no onprem content.
                            $Content = "$($xl.Installer.CustomData.PackageUriNew)"
                            $Verified = "N/A"
                            $InstallCL = "N/A"
                            $UninstallCL = "N/A"
                        }
                    }
                    $DTListArray =@()
                    #$DTsArray += New-Object -TypeName psobject -Property @{'Priority'= $DT.PriorityInLatestApp; 'Deployment Type Name' = "$($DT.LocalizedDisplayName)"; 'Technology' = "$($DT.Technology)"; 'Install Command' = "$InstallCL"; 'Uninstall Command' = "$UninstallCL"; 'Content Path' = "$Content"; 'Content Status'="$Verified"}
                    #$DTsArray = New-Object -TypeName psobject -Property @{'Priority'= $DT.PriorityInLatestApp; 'Deployment Type Name' = $DT.LocalizedDisplayName; 'Technology' = $DT.Technology; 'Commandline' = if (-not ($DT.Technology -like 'AppV*')){ $xmlDT.AppMgmtDigest.DeploymentType.Installer.CustomData.InstallCommandLine } }
                    $DTListTitle = "Deployment Type Name: $($DT.LocalizedDisplayName)"
                    $DTListArray += "Deployment Type Priority: $($DT.PriorityInLatestApp)"
                    $DTListArray += "Technology: $($DT.Technology)"
                    $DTListArray += "Install Command: $InstallCL"
                    $DTListArray += "Uninstall Command: $UninstallCL"
                    $DTListArray += "Content Path: $Content"
                    $DTListArray += "Content Status: $Verified"
                    Write-HtmlList -Title $DTListTitle -InputObject $DTListArray -Level 5 -File $FilePath
                    $InstallCL = ""
                    $UninstallCL = ""
                    $Content = ""
                    $Verified = ""
                }
                #$DTsArray = $DTsArray | sort Priority | Select-Object 'Priority','Deployment Type Name','Technology','Install Command','Uninstall Command','Content Path','Content Status'
                #Write-HtmlTable -InputObject $DTsArray -Border 1 -Level 5 -File $FilePath
                
            }
            else {
                Write-HTMLParagraph -Text 'There are no Deployment Types configured for this Application.' -Level 5 -File $FilePath
            }
            Write-Verbose "$(Get-Date):   Processing deployments for: $($App.LocalizedDisplayName)"
            $AppDeployments = Get-CMApplicationDeployment -ApplicationID $($App.CI_ID)
            Write-HTMLHeading -Level 5 -Text "Deployments for $($App.LocalizedDisplayName):" -File $FilePath
            if (-not [string]::IsNullOrEmpty($AppDeployments)) {
                $DeploymentsArray = @()
                foreach ($AppDeployment in $AppDeployments){
                    Switch ($AppDeployment.DesiredConfigType){
                        1{$Action = 'Install'}
                        2{$Action = 'Remove'}
                    }
                    Switch ($AppDeployment.OfferTypeID){
                        0{$Purpose = 'Required'}
                        2{$Purpose = 'Available'}
                    }
                    Switch ($AppDeployment.UserUIExperience){
                        False{$UserNotice = 'Hide in Software Center and all notifications'}
                        True{
                            Switch ($AppDeployment.NotifyUser){
                                True{$UserNotice = 'Display in Software Center and show all notifications'}
                                False{$UserNotice = 'Display in Software Center and only show notifications for computer restarts'}
                            }
                        }
                    }
                    Switch ($AppDeployment.UseGMTTimes){
                        True{$TimeZone = 'GMT'}
                        False{$TimeZone = 'Client Local Time'}
                    }
                    $DeploymentsArray += New-Object -TypeName psobject -Property @{'Collection'="$($AppDeployment.CollectionName)";'Action'="$Action";'Purpose'="$Purpose";'User Notification'="$UserNotice";'Available Time'="$($AppDeployment.StartTime)";'Deadline'="$($AppDeployment.EnforcementDeadline)";'Time Zone'="$TimeZone"}
                }
                $DeploymentsArray = $DeploymentsArray | Select-Object 'Collection','Action','Purpose','User Notification','Available Time','Deadline','Time Zone'
                Write-HtmlTable -InputObject $DeploymentsArray -Border 1 -Level 6 -File $FilePath
            }else{
                Write-HTMLParagraph -Text 'There are no deployments for this application.' -Level 6 -File $FilePath
            }
        }
    }else{
        Write-HTMLParagraph -Text 'There are no Applications configured in this site.' -Level 4 -File $FilePath
    }
}
elseif (-not [string]::IsNullOrEmpty($Applications)) {
    Write-HTMLParagraph -Text "There are $($Applications.count) applications configured." -Level 4 -File $FilePath
    $AppBasics = @()
    foreach ($App in $Applications){
        $AppBasics += New-Object -TypeName PSObject -Property @{'Name'="$($App.LocalizedDisplayName)"; 'Created by' = $($App.CreatedBy); 'Date created'=$($App.DateCreated)}
    }
    $AppBasics = $AppBasics | Select 'Name','Created by','Date Created'
    Write-HtmlTable -InputObject $AppBasics -Border 1 -Level 4 -File $FilePath
}
else {
    Write-HTMLParagraph -Text 'There are no Applications configured in this site.' -Level 4 -File $FilePath
}
Write-Verbose "$(Get-Date):   Applications Complete."
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Applications

#region Packages
Write-Verbose "$(Get-Date):   Processing Packages."
Write-HTMLHeading -PageBreak -Level 3 -Text 'Packages' -File $FilePath
$Packages = Get-CMPackage
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($Packages)){
        Write-HTMLParagraph -Text "There are $($Packages.count) packages configured." -Level 5 -File $FilePath
        Write-HTMLParagraph -Text "Below is a summary of all $($Packages.count) packages defined in this site. These are applications that are installed using traditional packages." -Level 3 -File $FilePath
        foreach ($Package in $Packages) {
            Write-Verbose "$(Get-Date):   Found Package: $($Package.Name)"
            Write-HTMLHeading -Level 4 -Text "$($Package.Name)" -File $FilePath
            $PackageDetailList = @()
            $PackageDetailList += "Description: $($Package.Description)"
            $PackageDetailList += "PackageID: $($Package.PackageID)"
            $PackageDetailList += "Package Source Files: $($Package.PkgSourcePath)"
            if (Test-Path "filesystem::$($Package.PkgSourcePath)" -ErrorAction SilentlyContinue){
                $Verified = "Path Verified"
            }else{
                $Verified = "Path not found"
            }
            $PackageDetailList += "Source Files exist: $Verified"
            #$PackageDetailList += 'The Package has the following Programs configured:'
            Write-HtmlList -InputObject $PackageDetailList -Level 5 -File $FilePath
            Write-HTMLHeading -Level 6 -Text 'The Package has the following Programs configured:' -File $FilePath
            $Programs = Get-WmiObject -Class SMS_Program -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "PackageID = '$($Package.PackageID)'"
            Write-Verbose "$(Get-Date):   Getting programs in package $($Package.Name)..."
            if (-not [string]::IsNullOrEmpty($Programs)){
                foreach ($Program in $Programs){
                    $ProgramList = @()
                    $ProgramTitle = "Program Name: $($Program.ProgramName)"
                    $ProgramList += "Command Line: $($Program.CommandLine)"
                    if ($Program.ProgramFlags -band 0x00000001){
                        $ProgramList += "Allow this program to be installed from the Install Package task sequence without being deployed: Enabled"
                    }
                    if ($Program.ProgramFlags -band 0x00000002){
                        $ProgramList += "The task sequence shows a custom progress user interface message: Enabled"
                    }
                    if ($Program.ProgramFlags -band 0x00000010){
                        $ProgramList += "This is a default program."
                    }
                    if ($Program.ProgramFlags -band 0x00000020){
                        $ProgramList += "Disables MOM alerts while the program runs."
                    }
                    if ($Program.ProgramFlags -band 0x00000040){
                        $ProgramList += 'Generates MOM alert if the program fails.'
                    }
                    if ($Program.ProgramFlags -band 0x00000080){
                        $ProgramList += "This program's immediate dependent should always be run."
                    }
                    if ($Program.ProgramFlags -band 0x00000100){
                        $ProgramList += 'A device program. The program is not offered to desktop clients.'
                    }
                    if ($Program.ProgramFlags -band 0x00000400){
                        $ProgramList += 'The countdown dialog is not displayed.'
                    }
                    if ($Program.ProgramFlags -band 0x00001000){
                        $ProgramList += 'The program is disabled.'
                    }
                    if ($Program.ProgramFlags -band 0x00002000){
                        $ProgramList += 'The program requires no user interaction.'
                    }
                    if ($Program.ProgramFlags -band 0x00004000){
                        $ProgramList += 'The program can run only when a user is logged on.'
                    }
                    if ($Program.ProgramFlags -band 0x00008000){
                        $ProgramList += 'The program must be run as the local Administrator account.'
                    }
                    if ($Program.ProgramFlags -band 0x00010000){
                        $ProgramList += 'The program must be run by every user for whom it is valid. Valid only for mandatory jobs.'
                    }
                    if ($Program.ProgramFlags -band 0x00020000){
                        $ProgramList += 'The program is run only when no user is logged on.'
                    }
                    if ($Program.ProgramFlags -band 0x00040000){
                        $ProgramList += 'The program will restart the computer.'
                    }
                    if ($Program.ProgramFlags -band 0x00080000){
                        $ProgramList += 'Configuration Manager restarts the computer when the program has finished running successfully.'
                    }
                    if ($Program.ProgramFlags -band 0x00100000){
                        $ProgramList += 'Use a UNC path (no drive letter) to access the distribution point.'
                    }
                    if ($Program.ProgramFlags -band 0x00200000){
                        $ProgramList += 'Persists the connection to the drive specified in the DriveLetter property. The USEUNCPATH bit flag must not be set.'
                    }
                    if ($Program.ProgramFlags -band 0x00400000){
                        $ProgramList += 'Run the program as a minimized window.'
                    }
                    if ($Program.ProgramFlags -band 0x00800000){
                        $ProgramList += 'Run the program as a maximized window.'
                    }
                    if ($Program.ProgramFlags -band 0x01000000){
                        $ProgramList += 'Hide the program window.'
                    }
                    if ($Program.ProgramFlags -band 0x02000000){
                        $ProgramList += 'Logoff user when program completes successfully.'
                    }
                    if ($Program.ProgramFlags -band 0x08000000){
                        $ProgramList += 'Override check for platform support.'
                    }
                    if ($Program.ProgramFlags -band 0x20000000){
                        $ProgramList += 'Run uninstall from the registry key when the advertisement expires.'
                    }
                    Write-HtmlList -Title $ProgramTitle -InputObject $ProgramList -Level 6 -File $FilePath
                }
            }else{
                Write-Verbose "$(Get-Date):   No programs found in package $($Package.Name)..."
                Write-HTMLParagraph -Text 'The Package has no Programs configured.' -Level 6 -File $FilePath
            }                       
        }
    }else{
        Write-HTMLParagraph -Text 'There are no Packages configured in this site.' -Level 5 -File $FilePath
    }
}
elseif (-not [string]::IsNullOrEmpty($Packages)){
    Write-HTMLParagraph -Text "There are $($Packages.count) packages configured." -Level 5 -File $FilePath
    $PackageBasics = @()
    foreach ($Package in $Packages){
        $PackageBasics += New-Object -TypeName PSObject -Property @{'Name'="$($Package.Name)"; 'Programs' = $($Package.NumOfPrograms); 'Content Date'=$($Package.SourceDate)}
    }
    $PackageBasics = $PackageBasics | Select 'Name','Programs','Content Date'
    Write-HtmlTable -InputObject $PackageBasics -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'There are no packages configured in this site.' -Level 5 -File $FilePath
}
Write-Verbose "$(Get-Date):   Completed processing Packages."
#endregion Packages
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Application Management


#region Software Updates
Write-HTMLHeading -Level 2 -PageBreak -Text 'Software Updates' -File $FilePath

#region Update Groups
Write-HTMLHeading -Level 3 -PageBreak -Text 'Software Update Groups' -File $FilePath
$UpdateGroups = Get-CMSoftwareUpdateGroup
If(-not [string]::IsNullOrEmpty($UpdateGroups)){
    Write-HTMLParagraph -Text "There are $($UpdateGroups.count) update groups defined in this site." -Level 3 -File $FilePath
    $UGs = $UpdateGroups|Sort LocalizedDisplayName|Select @{Name='Group Name';expression={$_.LocalizedDisplayName}},@{Name='ID';expression={$_.CI_ID}},@{Name='Update Count';expression={$_.NumberOfUpdates}},@{Name='Expired Updates';expression={$_.NumberOfExpiredUpdates}},@{Name='Created By';expression={$_.CreatedBy}},@{Name='Date Created';expression={$_.DateCreated}},@{Name='Deployed';expression={$_.IsDeployed}},@{Name='Compliance';expression={"$($_.PercentCompliant)%"}}
    Write-HtmlTable -InputObject $UGs -Border 1 -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "There are no update groups defined in this site." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Update Groups

#region Update Packages
Write-HTMLHeading -Level 3 -PageBreak -Text 'Software Update Packages' -File $FilePath
$UpdatePackages=Get-CMSoftwareUpdateDeploymentPackage
If(-not [string]::IsNullOrEmpty($UpdatePackages)){
    Write-HTMLParagraph -Text "There are $($UpdatePackages.count) update packages defined in this site." -Level 3 -File $FilePath
    $SUPackages = @()
    foreach($UP in $UpdatePackages){
        $Name = $UP.Name
        #binary differential replication
        If ($UP.PkgFlags -band 0x04000000){
            $BDR = "Enabled"
        }else{
            $BDR = "Disabled"
        }
        $SourcePath="$($UP.PkgSourcePath)"
        $PackageID = "$($UP.PackageID)"
        If(Test-Path -Path "filesystem::$SourcePath"){
            $SourceStatus = "Verified"
        }else{
            $SourceStatus = "Not Found"
        }
        $SUPackages += New-Object -TypeName PSObject -Property @{'Name'="$Name";'Package ID'="$PackageID";'BDR'="$BDR";'Source Path'="$SourcePath";'Source Status'="$SourceStatus"}
    }
    $SUPackages= $SUPackages|Sort Name|Select-Object 'Name','Package ID','BDR','Source Path','Source Status'
    Write-HtmlTable -InputObject $SUPackages -Border 1 -Level 3 -File $FilePath
    Write-HTMLParagraph -Text '(BDR = Binary Differential Replication)' -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text "There are no update packages defined in this site." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Update Packages


#region ADRs
Write-Verbose "$(Get-Date):   Beginning processing of ADRs..."
Write-HTMLHeading -Level 3 -PageBreak -Text 'Automatic Deployment Rules' -File $FilePath
$CMPSSuppressFastNotUsedCheck = $true
$ADRs=Get-CMSoftwareUpdateAutoDeploymentRule
foreach ($ADR in $ADRs){
    Write-Verbose "$(Get-Date):   Processing ADR $($ADR.Name)"
    $ADRListDetails = @()
    $ADRListTitle = "Name: $($ADR.Name)"
    Write-HTMLHeading -Level 4 -Text "$($ADR.Name)" -File $FilePath
    $ADRListDescription = $ADR.Description
    Remove-Variable languages -ErrorAction SilentlyContinue
    foreach ($locale in ([xml]$adr.ContentTemplate).ContentActionXML.ContentLocales.Locale){
        if ($locale -ne 'Locale:0'){$languages = "$languages, $((Get-CMCategory -Id $locale).LocalizedCategoryInstanceName)"}
    }
    $ADRListDetails += "Languages: $($languages.Trim(', '))"
    If (-not [string]::IsNullOrEmpty($ADR.Schedule)){
        $Schedule=Convert-CMSchedule $ADR.Schedule
        if ($Schedule.DaySpan -gt 0){
            $ADRListDetails += "Evaluation Schedule: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
        }
        elseif ($Schedule.HourSpan -gt 0){
            $ADRListDetails += "Evaluation Schedule: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
        }
        elseif ($Schedule.MinuteSpan -gt 0){
            $ADRListDetails += "Evaluation Schedule: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
        }
        elseif ($Schedule.ForNumberOfWeeks){
            $ADRListDetails += "Evaluation Schedule: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
        }
        elseif ($Schedule.ForNumberOfMonths){
            if ($Schedule.MonthDay -gt 0){
                $ADRListDetails += "Evaluation Schedule: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0){
                $ADRListDetails += "Evaluation Schedule: Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.WeekOrder -gt 0){
                switch ($Schedule.WeekOrder){
                    0 {$order = 'last'}
                    1 {$order = 'first'}
                    2 {$order = 'second'}
                    3 {$order = 'third'}
                    4 {$order = 'fourth'}
                }
                $ADRListDetails += "Evaluation Schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
        }
    }else{
        $ADRListDetails += "Evaluation Schedule: No schedule defined"
    }
    [xml]$rules=$ADR.UpdateRuleXML
    Remove-Variable Categories -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Web
    $UpdateRuleList = @()
    Write-Verbose "$(Get-Date):   Getting a list of configured update rules for this ADR: $($ADR.Name)"
    foreach ($UpdateRule in $rules.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem){
        Switch($UpdateRule.PropertyName){
            '_Product'{
                $Categories = ""
                foreach ($Id in $UpdateRule.MatchRules.string){
                    $UpdateCategory = Get-CMSoftwareUpdateCategory -UniqueId $($Id.trim("'")) -Fast
                    $Categories = "$Categories, $($UpdateCategory.LocalizedCategoryInstanceName)"
                }
                $UpdateRuleList += "Products: $($Categories.Trim(', '))"
            }
            '_UpdateClassification'{
                $UpdateClassification = @()
                foreach ($UC in $UpdateRule.MatchRules.string.Trim("'")){
                    $UPClasses = Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "SELECT LocalizedCategoryInstanceName FROM SMS_CIAllCategories WHERE CategoryTypeName=`'UpdateClassification`' and CategoryInstance_UniqueID=`'$UC`'" -ComputerName $SMSProvider
                    $UpdateClassification += $UPClasses[0].LocalizedCategoryInstanceName
                }
                $UC = "Update Classifications: $($UpdateClassification -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$UC")
            }
            'IsSuperseded'{
                If ($UpdateRule.MatchRules.string -eq $false){
                    $UpdateRuleList += 'Superseded: No'
                }else{
                    $UpdateRuleList += 'Superseded: Yes'
                }
            }
            'DateRevised'{
                $interval=($UpdateRule.MatchRules.string).split(':')
                If($interval[0] -gt 0) {$RevisedInterval = "Last $($interval[0]) year(s)"}
                If($interval[1] -gt 0) {$RevisedInterval = "Last $($interval[1]) month(s)"}
                If($interval[2] -gt 0) {$RevisedInterval = "Last $($interval[2]) days(s)"}
                If($interval[3] -gt 0) {$RevisedInterval = "Last $($interval[3]) hours(s)"}
                $UpdateRuleList += "Date Released or Revised: $RevisedInterval"
            }
            'ArticleID'{
                $ArticleID = "Article ID: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$ArticleID")
            }
            'BulletinID'{
                $BulletinID = "Bulletin ID: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$BulletinID")
            }
            'ContentSize'{
                $ContentSize = "Content Size: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$ContentSize")
            }
            'CustomSeverity'{
                #10-critical;8-Important;2-Low;6-Moderate;0-None
                $CustSev = @()
                foreach ($CS in $UpdateRule.MatchRules.string.Trim("'")){
                    Switch($CS){
                        10{$CustSev += 'Critical'}
                        8{$CustSev += 'Important'}
                        6{$CustSev += 'Moderate'}
                        2{$CustSev += 'Low'}
                        0{$CustSev += 'None'}
                    }
                }
                $UpdateRuleList += "Custom Severity: $($CustSev -join ' OR ')"
            }
            'LocalizedDescription'{
                $LocalizedDescription = "Description: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$LocalizedDescription")
            }
            'UpdateLocales'{
                $UpdateLanguages = @()
                foreach ($locale in $UpdateRule.MatchRules.string.Trim("'")){
                    $language = Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "SELECT LocalizedCategoryInstanceName FROM SMS_CIAllCategories WHERE CategoryTypeName=`'Locale`' and CategoryInstance_UniqueID=`'$locale`'" -ComputerName $SMSProvider
                    $UpdateLanguages += $language[0].LocalizedCategoryInstanceName
                }
                $ULs = "Languages: $($UpdateLanguages -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$ULs")
            }
            'NumMissing'{
                $NumMissing = "Required: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$NumMissing")
            }
            'Severity'{
                #10-critical;8-Important;2-Low;6-Moderate;0-None
                $Severity = @()
                foreach ($SV in $UpdateRule.MatchRules.string.Trim("'")){
                    Switch($SV){
                        10{$Severity += 'Critical'}
                        8{$Severity += 'Important'}
                        6{$Severity += 'Moderate'}
                        2{$Severity += 'Low'}
                        0{$Severity += 'None'}
                    }
                }
                $UpdateRuleList += "Custom Severity: $($Severity -join ' OR ')"
            }
            'LocalizedDisplayName'{
                $UTitle = "Title: $(($UpdateRule.MatchRules.string) -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$UTitle")
            }
            '_Company'{
                $Vendors = @()
                foreach ($Vendor in $UpdateRule.MatchRules.string.Trim("'")){
                    $Vend = Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "SELECT LocalizedCategoryInstanceName FROM SMS_CIAllCategories WHERE CategoryTypeName=`'Company`' and CategoryInstance_UniqueID=`'$Vendor`'" -ComputerName $SMSProvider
                    $UpdateVendor += $Vend[0].LocalizedCategoryInstanceName
                }
                $UpVendors = "Vendor: $($UpdateVendor -join ' OR ')"
                $UpdateRuleList += [System.Web.HttpUtility]::HtmlEncode("$UpVendors")
            }
        }
    }
    $ADRDeployments = Get-CMSoftwareUpdateAutoDeploymentRuleDeployment -ID $ADR.AutoDeploymentID
    $DeploymentPackage = Get-CMSoftwareUpdateDeploymentPackage -Id ([xml]$adr.ContentTemplate).ContentActionXML.PackageID
    $Package = New-Object -TypeName PSObject -Property @{'Package Name'="$($DeploymentPackage.Name)";'Description'="$($DeploymentPackage.Description)";'Package ID'="$($DeploymentPackage.PackageID)";'Source Location'="$($DeploymentPackage.PkgSourcePath)"}
    $Package = $package | Select-Object 'Package Name','Package ID','Description','Source Location'
    Write-HtmlList -InputObject $ADRListDetails -Description $ADRListDescription -Level 3 -File $FilePath
    #$UpdateRuleList
    Write-HtmlList -InputObject $UpdateRuleList -Description 'Software Update Property Filters (Update Rules):' -Level 3 -File $FilePath
    Write-HtmlTable -InputObject $Package -Level 5 -File $FilePath
    Write-HTMLHeading -Level 5 -Text "Deployments for ADR: $($ADR.Name)" -File $FilePath
    If ($ListAllInformation){
        Foreach ($Deployment in $ADRDeployments){
            $ADRDTListDetails = @()
            $DTxml=([xml]$Deployment.DeploymentTemplate).DeploymentCreationActionXML
            $ADRDTListTitle = "Deployment Collection: $($Deployment.CollectionName) ($($Deployment.CollectionID))"
            $ADRDTListDetails += "Enable the deployment after this rule is run: $($DTxml.EnableDeployment)"
            $ADRDTListDetails += "Use Wake-on-LAN to wake up clients for required deployments: $($DTxml.EnableWakeOnLan)"
            Switch ($($DTxml.StateMessageVerbosity)){
                1 {$StateMessages = 'Only error messages'}
                5 {$StateMessages = 'Only success and error messages'}
                10 {$StateMessages = 'All messages'}
            }
            $ADRDTListDetails += "Choose how much state detail you want clients to report back. Detail level: $StateMessages"
            Switch ($($DTxml.Utc)){
                false{$timebase = 'Client local time'}
                true{$timebase = 'UTC'}
            }
            $ADRDTListDetails += "Time based on: $timebase"
            If ($DTxml.AvailableDeltaDuration -eq 0){
                $ADRDTListDetails += "Software available time: As soon as possible"
            }else{
                $ADRDTListDetails += "Software available time: $($DTxml.AvailableDeltaDuration) $($DTxml.AvailableDeltaDurationUnits)"
            }
            If ($DTxml.Duration -eq 0){
                $ADRDTListDetails += "Installation Deadline: As soon as possible"
            }else{
                $ADRDTListDetails += "Installation Deadline: $($DTxml.Duration) $($DTxml.DurationUnits)"
            }
            $ADRDTListDetails += "Delay Enforcement of this deployment according to user preferences, up to the grace period defined in client settings: $($DTxml.SoftDeadlineEnabled)"
            Switch ($($DTxml.UserNotificationOption)){
                'DisplayAll'{$UserNotification = 'Display in Software Center and show all notifications'}
                'DisplaySoftwareCenterOnly'{$UserNotification = 'Display in Software Center, and only show nitifications for computer restarts'}
                'HideAll'{$UserNotification = 'Hide in Software Center and all notifications'}
            }
            $ADRDTListDetails += "User notifications: $UserNotification"
            Switch ($($DTxml.AllowInstallOutSW)){
                false{$InstallOutMW = 'Do not allow'}
                true{$InstallOutMW = 'Allow installations'}
            }
            $ADRDTListDetails += "Deadline behavior for Software Update installation outside of maintenance windows: $InstallOutMW"
            Switch ($($DTxml.AllowRestart)){
                false{$RestartOutMW = 'Do not allow'}
                true{$RestartOutMW = 'Allow restarts'}
            }
            $ADRDTListDetails += "Deadline behavior for System restarts outside of maintenance windows: $RestartOutMW"
            $ADRDTListDetails += "Suppress reboots on servers if update requires reboot: $($DTxml.SuppressServers)"
            $ADRDTListDetails += "Suppress reboots on workstations if update requires reboot: $($DTxml.SuppressWorkstations)"
            $ADRDTListDetails += "Windows Embedded devices, Commit changes at deadline: $($DTxml.PersistOnWriteFilterDevices)"
            $ADRDTListDetails += "If any update in this deployment requires a system restart, run updates deployment evaluation cycle after restart: $($DTxml.RequirePostRebootFullScan)"
            If($($DTxml.EnableAlert) -eq $false){
                $ADRDTListDetails += "Configuration Manager alerts.  Generate an alert when the following conditions are met: False"
            }else{
                $ADRDTListDetails += "Configuration Manager alerts.  Generate an alert when the following conditions are met: True<br />Client Compliance is below the following percent: $($DTxml.AlertThresholdPercentage)<br />Offset from the deadline: $($DTxml.AlertDuration)"
            }
            $ADRDTListDetails += "Disable Operations Manager alerts while software updates run: $($DTxml.DisableMomAlert)"
            $ADRDTListDetails += "Generate Operations Manager alert when a software update installation fails: $($DTxml.GenerateMomAlert)"
            switch ($DTxml.UseRemoteDP){
                false{$deploymentopt = 'Do not install software updates'}
                true{$deploymentopt = 'Download software updates from distribution point and install'}
            }
            $ADRDTListDetails += "Select deployment options to use when when client uses neighbor or default boundary group: $deploymentopt"
            switch ($DTxml.UseUnprotectedDP){
                false{$deploymentopt2 = 'Do not install software updates'}
                true{$deploymentopt2 = 'Download and install software updates from the distribution points in the site default boundary group'}
            }
            $ADRDTListDetails += "When software updates are not available on any distribution point in current or neighbor boundary group, download from default boundary group: $deploymentopt2"
            $ADRDTListDetails += "Allow clients to share content with other clients on the same subnet: $($DTxml.UseBranchCache)"
            $ADRDTListDetails += "If software updates are not available on distribution point in current, neighbor or site boundary groups, download content from Microsoft Updates: $($DTxml.AllowWUMU)"
            $ADRDTListDetails += "Allow clients on a metered Internet connection to download content after the installation deadline which might incur additional costs: $($DTxml.AllowUseMeteredNetwork)"
            Write-HtmlList -InputObject $ADRDTListDetails -Title $ADRDTListTitle -Level 5 -File $FilePath
            #$DTxml
        }
    }Else{
        Write-HtmlTable -InputObject ($ADRDeployments|select @{Name='Collection';expression={$_.CollectionName}},Enabled) -Level 5 -File $FilePath
    }
}
Write-HtmliLink -ReturnTOC -File $FilePath
Write-Verbose "$(Get-Date):   Completed processing of ADRs."
#endregion ADRs


#endregion Software Updates


#region Operating Systems
Write-HTMLHeading -Level 2 -PageBreak -Text 'Operating Systems' -File $FilePath

#region Driver Packages
Write-Verbose "$(Get-Date):   Processing Driver Packages."
Write-HTMLHeading -Level 3 -PageBreak -Text 'Driver Packages' -File $FilePath
$DriverPackages = Get-CMDriverPackage
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($DriverPackages)){
        Write-HTMLParagraph -Text 'The following Driver Packages are configured in your site:' -Level 4 -File $FilePath
        foreach ($DriverPackage in $DriverPackages){
            $DPackArray = @()
            $PackageName = "$($DriverPackage.Name)"
            $PackageDescription = ""
            if ($DriverPackage.Description){
                $PackageDescription = "Description: $($DriverPackage.Description)"
            }
            $DPackArray += "PackageID: $($DriverPackage.PackageID)"
            $DPackArray += "Source path: $($DriverPackage.PkgSourcePath)"
            if (Test-Path "filesystem::$($DriverPackage.PkgSourcePath)" -ErrorAction SilentlyContinue){
                $Verified = "Path Verified"
            }else{
                $Verified = "Path not found"
            }
            $DPackArray += "Source Files exist: $Verified"
            $DPackArray += 'This package consists of the following Drivers:'
            $Drivers = Get-CMDriver -DriverPackageId "$($DriverPackage.PackageID)"
            $DriverArray = @()
            foreach ($Driver in $Drivers){
                if (Test-Path "filesystem::$($Driver.ContentSourcePath)" -ErrorAction SilentlyContinue){
                    $Verified = "Path Verified"
                }else{
                    $Verified = "Path not found"
                }
                $DriverArray += New-Object -TypeName psobject -Property @{'Driver Name'="$($Driver.LocalizedDisplayName)";'Manufacturer'="$($Driver.DriverProvider)";'Source Path'="$($Driver.ContentSourcePath)";'Source Status' = "$Verified";'INF File'="$($Driver.DriverINFFile)"}
            }
            If ($PackageDescription){
                Write-HtmlList -Title $PackageName -Description $PackageDescription -InputObject $DPackArray -Level 4 -File $FilePath
            }else{
                Write-HtmlList -Title $PackageName -InputObject $DPackArray -Level 4 -File $FilePath
            }
            $DriverArray = $DriverArray | Select-Object 'Driver Name','Manufacturer','Source Path','Source Status','INF File'
            Write-HtmlTable -InputObject $DriverArray -Border 1 -Level 5 -File $FilePath
        }
    }else{
        Write-HTMLParagraph -Text 'There are no Driver Packages configured in this site.' -Level 4 -File $FilePath
    }
}else{
    if (-not [string]::IsNullOrEmpty($DriverPackages)){
        Write-HTMLParagraph -Text "There are $($DriverPackages.count) Driver Packages configured." -Level 4 -File $FilePath
    }else{
        Write-HTMLParagraph -Text 'There are no Driver Packages configured in this site.' -Level 4 -File $FilePath
    }
}
Write-Verbose "$(Get-Date):   Completed processing Driver Packages."
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Driver Packages

#region Operating System Upgrade Packages
Write-Verbose "$(Get-Date):   Processing Operating System Upgrade Packages."
Write-HTMLHeading -Level 3 -PageBreak -Text 'Operating System Upgrade Packages' -File $FilePath
$OSUgPacks = Get-CMOperatingSystemInstaller
if (-not [string]::IsNullOrEmpty($OSUgPacks)){
    Write-HTMLParagraph -Text 'The following Operating System Upgrade Packages are available in this CM Site:' -Level 4 -File $FilePath
    foreach ($OSUpgradePack in $OSUgPacks)
        {
            $UgPackList = @()
            $UgPackTitle = "$($OSUpgradePack.Name)"
            $UgPackDescription = ""
            if ($OSUpgradePack.Description -ne "")
                    {
                        $UgPackDescription = "Description/Comment: $($OSUpgradePack.Description)"
                    }
            $UgPackList += "Version: $($OSUpgradePack.PackageID)"
            $UgPackList += "Language: $($OSUpgradePack.Language)"
            $UgPackList += "Image OS Version: $($OSUpgradePack.ImageOSVersion)"
            $UgPackList += "Package ID: $($OSUpgradePack.PackageID)"
            $UgPackList += "Source Path: $($OSUpgradePack.PkgSourcePath)"
            if (Test-Path "filesystem::$($OSUpgradePack.PkgSourcePath)" -ErrorAction SilentlyContinue){
                $Verified = "Path exists"
            }else{
                $Verified = "Path not found"
            }
            $UgPackList += "Source Path Status: $Verified"
            If ($UgPackDescription){
                Write-HtmlList -Title $UgPackTitle -Description $UgPackDescription -InputObject $UgPackList -Level 4 -File $FilePath
            }else{
                Write-HtmlList -Title $UgPackTitle -InputObject $UgPackList -Level 4 -File $FilePath
            }
        }
}else{
    Write-HTMLParagraph -Text 'There are no Operating System Upgrade Packages found in this site.' -Level 4 -File $FilePath
}
Write-Verbose "$(Get-Date):   Completed processing Operating System Upgrade Packages."
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Operating System Upgrade Packages

#region Operating System Images
Write-Verbose "$(Get-Date):   Processing Operating System Images."
Write-HTMLHeading -Level 3 -PageBreak -Text 'Operating System Images' -File $FilePath
$OSImages = Get-CMOperatingSystemImage
if (-not [string]::IsNullOrEmpty($OSImages)){
    Write-HTMLParagraph -Text 'The following Operating System Images are available in this CM Site:' -Level 4 -File $FilePath
    foreach ($OSImage in $OSImages)
        {
            $OSImageList = @()
            $OSImageName = "$($OSImage.Name)"
            $OSImageDescription = ""
            if ($OSImage.Description -ne "")
                    {
                        $OSImageDescription = "Description/Comment: $($OSImage.Description)"
                    }
            $OSImageList += "Version: $($OSImage.PackageID)"
            $OSImageList += "Language: $($OSImage.Language)"
            $OSImageList += "Image OS Version: $($OSImage.ImageOSVersion)"
            $OSImageList += "Package ID: $($OSImage.PackageID)"
            $OSImageList += "Source Path: $($OSImage.PkgSourcePath)"
            if (Test-Path "filesystem::$($OSImage.PkgSourcePath)" -ErrorAction SilentlyContinue){
                $OSImageList += "Source Path Status: Path exists"
                $OSImageList += "Image Size: $([int]((Get-Item filesystem::$($OSImage.PkgSourcePath)).Length/1MB)) MB"
            }else{
                $OSImageList += "Source Path Status: Path not found"
            }
            If ($OSImageDescription){
                Write-HtmlList -Title $OSImageName -Description $OSImageDescription -InputObject $OSImageList -Level 4 -File $FilePath
            }else{
                Write-HtmlList -Title $OSImageName -InputObject $OSImageList -Level 4 -File $FilePath
            }
        }
}else{
    Write-HTMLParagraph -Text 'There are no Operating System Images found in this site.' -Level 4 -File $FilePath
}
Write-Verbose "$(Get-Date):   Completed processing Operating System Images."
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Operating System Images

#region Boot Images
Write-Verbose "$(Get-Date):   Processing Boot Images."
Write-HTMLHeading -Level 3 -PageBreak -Text 'Boot Images' -File $FilePath
$BootImages = Get-CMBootImage
if (-not [string]::IsNullOrEmpty($BootImages)){
    Write-HTMLParagraph -Text 'The following Boot Images are available in this site:' -Level 4 -File $FilePath
    foreach ($BootImage in $BootImages){
        $BootImageList = @()
        $BootImageName = "$($BootImage.Name)"
        $BootImageDescription = ""
        if ($BootImage.Description -ne ""){
            $BootImageDescription = "Description/Comment: $($BootImage.Description)"
        }
        $BootImageList += "Last Updated: $($BootImage.LastRefreshTime)"
        $BootImageList += "Source Path: $($BootImage.PkgSourcePath)"
        $BootImageList += "Package ID: $($BootImage.PackageID)"
        $BootImageList += "Boot Image OS Version: $($BootImage.ImageOSVersion)"
        switch ($BootImage.Architecture)
            {
                0 { $BootImageList += "Architecture: x86" }
                9 { $BootImageList += "Architecture: x64" }
            }
        If ($BootImage.PkgFlags -band 0x00000400){
            $BootImageList += "Deploy this boot image from the PXE-enabled distribution point: Enabled"
        }else{
            $BootImageList += "Deploy this boot image from the PXE-enabled distribution point: Disabled"
        }
        If ($BootImage.PkgFlags -band 0x04000000){
            $BootImageList += "Enable binary differential replication: Enabled"
        }else{
            $BootImageList += "Enable binary differential replication: Disabled"
        }
        if ($BootImage.BackgroundBitmapPath)
            {
                $BootImageList += "Custom Background: $($BootImage.BackgroundBitmapPath)"
            }
        Switch ($BootImage.EnableLabShell)
            {
                True { $BootImageList += 'Command line support is enabled' }
                False { $BootImageList += 'Command line support is not enabled' }
            }
        $BootImageList += 'The following drivers are imported into this WinPE:'
        If ($OSImageDescription){
            Write-HtmlList -Title $BootImageName -Description $BootImageDescription -InputObject $BootImageList -Level 4 -File $FilePath
        }else{
            Write-HtmlList -Title $BootImageName -InputObject $BootImageList -Level 4 -File $FilePath
        }
        if (-not [string]::IsNullOrEmpty($BootImage.ReferencedDrivers)){
            $DriverArray = @()
            $ImportedDriverIDs = ($BootImage.ReferencedDrivers).ID
            foreach ($ImportedDriverID in $ImportedDriverIDs){
                $ImportedDriver = Get-CMDriver -ID $ImportedDriverID
                $DriverArray += New-Object -TypeName psobject -Property @{'Driver Name'="$($ImportedDriver.LocalizedDisplayName)";'Driver Class'="$($ImportedDriver.DriverClass)";'Inf File'="$($ImportedDriver.DriverINFFile)"}
            }
            Write-HtmlTable -InputObject $DriverArray -Border 1 -Level 6 -File $FilePath
        }else{
            #$DriverArray += New-Object -TypeName psobject -Property @{'Driver Name'='There are no drivers imported into the Boot Image.'}
            Write-HTMLParagraph -Level 6 -Text 'There are no drivers imported into the Boot Image.' -File $FilePath
        }
        if (-not [string]::IsNullOrEmpty($BootImage.OptionalComponents)){
            $Component = $Null
            $OCList = @()
            $OCDescription = 'The following Optional Components are added to this Boot Image:'
            foreach ($Component in $BootImage.OptionalComponents){
                switch ($Component){
                    {($_ -eq '1') -or ($_ -eq '27')} { $OCList += 'WinPE-DismCmdlets' }
                    {($_ -eq '2') -or ($_ -eq '28')} { $OCList += 'WinPE-Dot3Svc' }
                    {($_ -eq '3') -or ($_ -eq '29')} { $OCList += 'WinPE-EnhancedStorage' }
                    {($_ -eq '4') -or ($_ -eq '30')} { $OCList += 'WinPE-FMAPI' }
                    {($_ -eq '5') -or ($_ -eq '31')} { $OCList += 'WinPE-FontSupport-JA-JP' }
                    {($_ -eq '6') -or ($_ -eq '32')} { $OCList += 'WinPE-FontSupport-KO-KR' }
                    {($_ -eq '7') -or ($_ -eq '33')} { $OCList += 'WinPE-FontSupport-ZH-CN' }
                    {($_ -eq '8') -or ($_ -eq '34')} { $OCList += 'WinPE-FontSupport-ZH-HK' }
                    {($_ -eq '9') -or ($_ -eq '35')} { $OCList += 'WinPE-FontSupport-ZH-TW' }
                    {($_ -eq '10') -or ($_ -eq '36')} { $OCList += 'WinPE-HTA' }
                    {($_ -eq '11') -or ($_ -eq '37')} { $OCList += 'WinPE-StorageWMI' }
                    {($_ -eq '12') -or ($_ -eq '38')} { $OCList += 'WinPE-LegacySetup' }
                    {($_ -eq '13') -or ($_ -eq '39')} { $OCList += 'WinPE-MDAC' }
                    {($_ -eq '14') -or ($_ -eq '40')} { $OCList += 'WinPE-NetFx4' }
                    {($_ -eq '15') -or ($_ -eq '41')} { $OCList += 'WinPE-PowerShell3' }
                    {($_ -eq '16') -or ($_ -eq '42')} { $OCList += 'WinPE-PPPoE' }
                    {($_ -eq '17') -or ($_ -eq '43')} { $OCList += 'WinPE-RNDIS' }
                    {($_ -eq '18') -or ($_ -eq '44')} { $OCList += 'WinPE-Scripting' }
                    {($_ -eq '19') -or ($_ -eq '45')} { $OCList += 'WinPE-SecureStartup' }
                    {($_ -eq '20') -or ($_ -eq '46')} { $OCList += 'WinPE-Setup' }
                    {($_ -eq '21') -or ($_ -eq '47')} { $OCList += 'WinPE-Setup-Client' }
                    {($_ -eq '22') -or ($_ -eq '48')} { $OCList += 'WinPE-Setup-Server' }
                    #{($_ -eq "23") -or ($_ -eq "49")} { $OCList += "Not applicable" }
                    {($_ -eq '24') -or ($_ -eq '50')} { $OCList += 'WinPE-WDS-Tools' }
                    {($_ -eq '25') -or ($_ -eq '51')} { $OCList += 'WinPE-WinReCfg' }
                    {($_ -eq '26') -or ($_ -eq '52')} { $OCList += 'WinPE-WMI' }
                }
                $Component = $Null    
            }
            Write-HtmlList -Description $OCDescription -InputObject $OCList -Level 5 -File $FilePath
        }

    }
}else{
    Write-HTMLParagraph -Text 'There are no Boot Images present in this site.' -Level 4 -File $FilePath
}
Write-Verbose "$(Get-Date):   Completed processing Boot Images."
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Boot Images


#region Task Sequences
Write-Verbose "$(Get-Date):   Enumerating Task Sequences"
Write-HTMLHeading -Level 3 -PageBreak -Text 'Task Sequences' -File $FilePath
$TaskSequences = Get-CMTaskSequence
Write-Verbose "$(Get-Date):   working on $($TaskSequences.count) Task Sequences"
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($TaskSequences)){
        foreach ($TaskSequence in $TaskSequences){
                Write-Verbose "$(Get-Date):   Detailing $($TaskSequence.Name) Task Sequences"
                Write-HTMLHeading -Level 4 -Text "$($TaskSequence.Name)" -File $FilePath
                $TSDetails = @()
                $TSDetails += "Package ID: $($TaskSequence.PackageID)"
                $BootImage = $TaskSequence.References.Package|foreach {(Get-CMBootImage -id $_).Name} 
                If([string]::IsNullOrEmpty($BootImage)){$BootImage="None"}
                $TSDetails += "Task Sequence Boot Image: $BootImage"
                $OSImage = $TaskSequence.References.Package|foreach {(Get-CMOperatingSystemImage -id $_).Name}
                If([string]::IsNullOrEmpty($OSImage)){$OSImage="None"}
                $TSDetails += "Task Sequence Operating System Image: $OSImage"
                $TSDetails += "Sequence Steps:"
                Write-HtmlList -InputObject $TSDetails -Level 4 -File $FilePath
                $Sequence = $Null
                $Sequence = ([xml]$TaskSequence.Sequence).sequence
                $AllSteps = Process-TSSteps -Sequence $Sequence
                $c = 0
                foreach ($Step in $AllSteps){$c++;$Step|Add-Member -MemberType NoteProperty -Name 'Step' -Value $c}
                $AllSteps = $AllSteps |Select-Object 'Step','Group Name','Step Name','Description','Action','Status'
                Write-HtmlTable -InputObject $AllSteps -Border 1 -Level 6 -File $FilePath
            }
    }else{
        Write-HTMLParagraph -Level 3 -Text 'There are no Task Sequences present in this environment.' -File $FilePath
    }
}else{
    if (-not [string]::IsNullOrEmpty($TaskSequences)){
        Write-HTMLParagraph -Level 3 -Text 'The following Task Sequences are configured:' -File $FilePath
        $TSList =@()
        foreach ($TaskSequence in $TaskSequences){
            $OSImage = $TaskSequence.References.Package|foreach {(Get-CMOperatingSystemImage -id $_).Name}
            $BootImage = $TaskSequence.References.Package|foreach {(Get-CMBootImage -id $_).Name} 
            $TSName = "$($TaskSequence.Name)"
            If([string]::IsNullOrEmpty($OSImage)){$OSImage="None"}
            If([string]::IsNullOrEmpty($BootImage)){$BootImage="None"}
            $TSList += New-Object -TypeName PSObject -Property @{'Name'="$TSName";'Operating System Image'="$OSImage";'Boot Image'="$BootImage"}
        }
        $TSList = $TSList | Select-Object 'Name','Operating System Image','Boot Image'
        Write-HtmlTable -InputObject $TSList -Level 3 -File $FilePath
    }else{
        Write-HTMLParagraph -Level 3 -Text 'There are no Task Sequences present in this environment.' -File $FilePath
    }
}
Write-Verbose "$(Get-Date):   Completed Task Sequences"
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Task Sequences

#endregion Operating Systems

#region Windows 10 Servicing

#region Servicing Plan

###Work in progress...

#$ServicingPlans = Get-CMWindowsServicingPlan
#$ServicingPlans=((Get-CMWindowsServicingPlan).UpdateRuleXML).updatexml.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem

foreach ($ServicingPlan in $ServicingPlans){
    [xml]$UR = $ServicingPlan.UpdateRuleXML
    #Name
    #Description
    #Target Collection
    #Deployment Ring: Channel, Deployment Delay
    ($UR.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem | where {$_.PropertyName -eq 'AfterDays'}).Matchrules.string
    #Evaluation Schedule
    #Deployment Schedule
    #Deployment Package
    #Language Selection
    #User Experience: User Notification
}
#endregion Servicing Plan

#region Windows Update for Business Policies
<#
 General
•	Name: Test WUB Policy
•	Description: Test Policy for deploying windows updates.
 Deferral Settings
•	Branch Readiness Level: Semi-Annual Channel
•	Feature Updates deferral period (in days): 30
•	Pause Feature Updates: No
•	Quality Updates deferral period (in days): 12
•	Pause Quality Updates: No
•	Install updates for other Microsoft products: Yes
•	Include drivers: No
#>
#endregion Windows Update for Business Policies

#endregion Windows 10 Servicing

#region Scripts
Write-Verbose "$(Get-Date):   Enumerating Configuration Manager Scripts"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Configuration Manager Scripts' -File $FilePath
$ScriptFeature = Get-CMSiteFeature|where{$_.FeatureGuid -like '566F8720-F415-4E10-9A51-CDE682BA2B2E'}
if (-not [string]::IsNullOrEmpty($ScriptFeature)){
    If ($ScriptFeature.Status -eq 1){
        $CMScripts = Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider -Query 'select ScriptName,Author,Approver,ApprovalState,ScriptType,LastUpdateTime from SMS_Scripts'
        Write-Verbose "$(Get-Date):   working on $($TaskSequences.count) Task Sequences"

        if ([string]::IsNullOrEmpty($CMScripts)){
            Write-HTMLParagraph -Text "No Scripts are defined in this site." -Level 3 -File $FilePath
        }else{
            $Scripts = @()
            foreach ($Script in $CMScripts){
                SWitch($script.ApprovalState){
                    0{$Approval = "Waiting for Approval"}
                    1{$Approval = "Declined"}
                    3{$Approval = "Approved"}
                    default{$Approval = "Unknown"}
                }
                $UpdateTime = [Management.ManagementDateTimeConverter]::ToDateTime($Script.LastUpdateTime)
                $Scripts += New-Object -TypeName PSObject -Property @{'Script Name'="$($Script.ScriptName)";'Author'="$($Script.Author)";'Approver'=$($Script.Approver);'Approval State'="$Approval";'Last Update Time' = "$UpdateTime"}
            }
            $Scripts = $Scripts | Select-Object 'Script Name','Author','Approver','Approval State','Last Update Time'
            Write-HtmlTable -InputObject $Scripts -Border 1 -Level 3 -File $FilePath
        }
    }else{
        Write-HTMLParagraph -Text "Scripts feature, `"$($ScriptFeature.Name)`", not enabled in this site." -Level 3 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text "Scripts feature not found in this site. Scripts were introduced with release 1706." -Level 3 -File $FilePath
}
Write-Verbose "$(Get-Date):   Completed Configuration Manager Scripts"
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Scripts

#endregion Software Library

Write-HTMLTOC -InputObject $Global:DocTOC -File $FilePath

Set-Location -Path "$StartingPath"
