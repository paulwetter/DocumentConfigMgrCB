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
.PARAMETER ListAllInformation
    Specifies whether the script should only output an overview of what is configured (like count of collections) or 
    a full output with verbose information. This includes: User & Device collections, Application, Packages, ADR Deployments, 
    Drivers in Driver packages, Task Squence Details
.PARAMETER ListAppDetails
    Like ListAllInformation, but instead of all details, lists only Application Details.
.PARAMETER NoSqlDetail
    Skip additional details from the SQL server.  Useful when you do not have full database access to the SQL server (Normally requires access via SQL to the Master and CM database).
.PARAMETER SMSProvider
    Some information rely on WMI queries that need to be executed against the SMS Provider directly. 
    Please specify as FQDN.
    If not specified, it assumes localhost.
.PARAMETER UnknownClientSettings
    With new releases of CM come new client settings.  If this parameter is added, it will display raw 
    information for these client settings.
.PARAMETER SQLTimeout
    The amount of time we should wait for a sql query to time out.  Default is 300 seconds (5 minutes)
.PARAMETER MaskAccounts
    This will mask about half of the account name in the documentation
.PARAMETER SQLCredential
    If The SQL server is on a remote system, you can pass SQL credentials here.
.PARAMETER SkipRemoteServerDetails
    Skip connecting directly to each remote site system server for hardware and OS details since this can take a long time on sites with many site systems
.PARAMETER StyleSheet
    This is the path to an external CSS file that will allow you to style the report in your own way.  The style sheet will be embedded into the report.
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
    VERSION: 3.61
    AUTHOR: Paul Wetter
    Based on original script developed by David O'Brien
    CONTRIBUTOR: Florian Valente (BlackCatDeployment), Skatterbrainz, ChadSimmons
    LASTEDIT: January 26, 2021
#>

#endregion

#region script parameters
[CmdletBinding()]

Param(
    [parameter(Mandatory=$True)] 
    [string]$CompanyName,
    
    [parameter(Mandatory=$False)] 
    [string]$CompanyLogo = "https://blog.cyberadvisors.com/hubfs/CAI_logo.jpg",

    [parameter(Mandatory=$False)] 
    [Switch]$ListAllInformation,

    [parameter(Mandatory=$False)] 
    [Switch]$ListAppDetails,

    [parameter(Mandatory=$False)] 
    [string]$Author="Paul Wetter",

    [parameter(Mandatory=$False)] 
    [string]$Vendor = "Cyber Advisors",

    [parameter(Mandatory=$False)] 
    [String]$Title = "Configuration Manager Site Documentation",

    [parameter(Mandatory=$False)]
    [ValidateScript({$_ -match '\.html$'})]  
    [String]$FilePath = "CMDocumentation.html",

    [parameter(Mandatory=$False)] 
    [string]$SMSProvider='localhost',

    [parameter(Mandatory=$False)] 
    [Switch]$AddDateTime,
	
    [parameter(Mandatory=$False)] 
    [switch]$UnknownClientSettings,
    
    [parameter(Mandatory=$False)] 
    [switch]$NoSqlDetail,

    [parameter(Mandatory=$False)]
    [int]$SQLTimeout = 300,

    [parameter(Mandatory=$False)] 
    [switch]$MaskAccounts,
    
    #[parameter(Mandatory=$False)] 
    #[System.Management.Automation.PSCredential]$SQLCredential = [System.Management.Automation.PSCredential]::Empty,

    [parameter(Mandatory=$False,HelpMessage="Skip connecting directly to each site system server for hardware and OS details")]
    [switch]$SkipRemoteServerDetails,

    [parameter(Mandatory=$False,HelpMessage="CSS file path")]
    [string]$StyleSheet = ""

	)
#endregion script parameters

$DocumenationScriptVersion = '3.61'


$CMPSSuppressFastNotUsedCheck = $true
$Global:DocTOC = @()
$ScriptStartTime = Get-date
Write-host "Beginning Execution of version $DocumenationScriptVersion at: $($ScriptStartTime.ToShortTimeString())"
Write-Verbose "Beginning Execution of version $DocumenationScriptVersion at: $($ScriptStartTime.ToShortTimeString())"

#region HTML Writing Functions
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
    if ($InputObject){
        $table = $InputObject|ConvertTo-Html -Fragment
        $table[0] = "<table cellpadding=$Padding cellspacing=$Spacing border=$Border style=`"margin-left:$($Indent)px;`">"
        $table = $table -replace "--CRLF--","<BR />"
        $table = $table -replace "--TAB--","&nbsp;&nbsp;"
        $table = $table -replace "--B--","<B>"
        $table = $table -replace "--/B--","</B>"
        $table = $table -replace "--I--","<I>"
        $table = $table -replace "--/I--","</I>"
        $table = $table -replace "--U--","<U>"
        $table = $table -replace "--/U--","</U>"
        $table = $table -replace "--CBOX--","&#9745;"
        $table = $table -replace "--UNCBOX--","&#9744;"
        $table = $table -replace "--BULLET--","&bull;"
    }else{
        Write-Verbose 'Input object was empty outputting empty object paragraph text...'
        Write-HTMLParagraph -Text 'There was no data to output from this query.' -Level $Level -File $file
    }
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
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the text that will appear inside the <P> tag")]
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
        HelpMessage="Custom Style sheet to apply to the domcumenation")]
        [string]$CssStyleFile,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the table will indent by")]
        [string]$File
    )
    $Header = @()
    $Header += "<html>"
    $Header += "<Head>"
    $Header += "<Title>$Title</Title>"
    # if custom stylesheet parameter is invoked, apply to output, otherwise use hard-coded style
    ##Default Style
    $DefaultStyle = @()
    $DefaultStyle += "<Style>"
    $DefaultStyle += 'H1  {background-color:royalblue; border-top: 1px solid black;}'
    $DefaultStyle += 'H2	{margin-left:10px;background-color:steelblue; border-top: 1px solid black;}'
    $DefaultStyle += 'H3	{margin-left:20px;background-color:lightblue; border-top: 1px solid black;}'
    $DefaultStyle += 'H4	{margin-left:30px;background-color:lightsteelblue; border-top: 1px solid black;}'
    $DefaultStyle += 'H5	{margin-left:40px;background-color:lightcyan; border-top: 1px solid black;}'
    $DefaultStyle += 'H6	{margin-left:50px;background-color:lavender; border-top: 1px solid black;}'
    $DefaultStyle += ".pagebreak { page-break-before: always; }"
    $DefaultStyle += "TH  {background-color:LightBlue;padding: 3px; border: 2px solid black;}"
    $DefaultStyle += "TD  {padding: 3px; border: 1px solid black;}"
    $DefaultStyle += "TABLE	{border-collapse: collapse;}"
    $DefaultStyle += "TABLE.Cover TR {background-color:white}"
    $DefaultStyle += ".CoverImage {width:auto; max-width:auto}"
    $DefaultStyle += "</Style>"
    ##End Default Style
    If($CssStyleFile){
        If(Test-Path -Path $CssStyleFile) {
            $StyleContent = Get-Content "$CssStyleFile"
            Write-Verbose "Applying custom style sheet to document: $CssStyleFile"
            $Header += "<Style>"
            $Header += $StyleContent
            $Header += "</Style>"
        }
        else {
            Write-Verbose "Custom style sheet not found [$CssStyleFile].  Using default style."
            $Header += $DefaultStyle
        }
    }Else{
        $Header += $DefaultStyle
    }
    $Header += "</Head>"
    $Header += "<Body>"
    If ($File) {IF (Test-Path -Path $File) {Remove-Item -Path $File -Force}}
    If ($File) {$header | Out-File -filepath $File -Append}
    Else {Return $Header}
}

Function Write-HTMLFooter{
<#
.SYNOPSIS
    This will write the end of the file for the document/HTML.
.PARAMETER File
    This is the file that the HTML will be written to.
.EXAMPLE
    Write-HTMLHeader
.EXAMPLE
    Write-HTMLHeader -file "C:\test.html"
.NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,
        HelpMessage="This is the amount of space that the table will indent by")]
        [string]$File
    )
    $Footer += "</body></html>"
    If ($File) {$Footer | Out-File -filepath $File -Append}
    Else {Return $Footer}
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
    $Cover += "<Table class=`"Cover`" border=0 cellspacing=0 cellpadding=0 style=`"width:100%;border: 0px`">"
    $Cover += "<TR><TD Height=50 VAlign=`"top`" align=`"left`" style=`"border: 0px;font-size:48pt`">$Title</TD></TR>"
    $Cover += "<TR><TD Height=20 VAlign=`"top`" align=`"left`" style=`"border: 0px;font-size:24pt;padding-left:10px`">Report Prepared for: $Org</TD></TR>"
    If ($ImagePath){
        $ImageData=Convert-Image2Base64 -Path $ImagePath
    }
    If ($ImageData){
        $Cover += "<TR><TD Height=680 VAlign=`"bottom`" align=`"right`" style=`"border: 0px`"><img class=`"CoverImage`" src=`"$ImageData`"></TD></TR>"
    }Else{
        $Cover += "<TR><TD Height=680 VAlign=`"top`" style=`"border: 0px`">&nbsp;</TD></TR>"
    }
    $Cover += "<TR><TD Height=30 VAlign=`"top`" Align=`"right`" style=`"border: 0px;font-size:18pt`">Report Prepared By: $Author</TD></TR>"
    If ($Vendor) {$Cover += "<TR><TD Height=30 VAlign=`"top`" Align=`"right`" style=`"border: 0px;font-size:24pt`">$Vendor</TD></TR>"}
    $Cover += "<TR><TD Height=20 VAlign=`"top`" align=`"right`" style=`"border: 0px;font-size:11pt;padding-left:10px`">Executed with script version: $DocumenationScriptVersion</TD></TR>"
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

##########################################################################################################################
###############################################HTML Format functions above################################################
##########################################################################################################################

#endregion HTML Writing Functions

#region Additional Functions

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
 
.PARAMETER QueryTimeout
    Specifies how long to wait until the SQL Query times out. default 300 Seconds
 
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
        
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false)]
        [int]$QueryTimeout = 300,

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
        $command.CommandTimeout = $QueryTimeout
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

#small function that will convert utc time to local time, option to ignore daylight savings time.
function Convert-UTCtoLocal{
    param(
        [parameter(Mandatory=$true)]
        [String]$UTCTimeString,
        [parameter(Mandatory=$false)]
        [Switch]$IgnoreDST
    )
    $UTCTime = ($UTCTimeString.Split('.'))[0]
    $dt = ([datetime]::ParseExact($UTCTime,'yyyyMMddhhmmss',$null))
    if ($IgnoreDST){
        $dt+([System.TimeZoneInfo]::Local).BaseUtcOffset
    }else{
        $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
        $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
        [System.TimeZoneInfo]::ConvertTimeFromUtc($dt, $TZ)
    }
}

Function Get-HumanReadableSchedule {
    <#
    .SYNOPSIS
    This reads the data from an CM schedule object and then converts the schedule into a readable text similar to what appears in the console dialogs
    
    .PARAMETER Schedule
    this is a schedule object from within CM.  Schedules are converted from a 16 digit hex string with Convert-CMSchedule
    
    .EXAMPLE
    Get-HumanReadableSchedule -Schedule (Convert-CMSchedule -ScheduleString 0001170000100038)
    
    .NOTES
    Author: Paul Wetter
    Website: 
    Email: tellwetter[at]gmail.com
    #>
    
    [CmdletBinding()]
    param(
        $Schedule
    )
    if ($Schedule.HourDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.HourDuration) hours"
    } 
    elseif ($Schedule.MinuteDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.MinuteDuration) minutes"
    } 
    elseif ($Schedule.DayDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.DayDuration) days"
    }

    if ($Schedule.DaySpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.HourSpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.MinuteSpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.ForNumberOfWeeks) {
        $HrSched = "Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.ForNumberOfMonths) {
        if ($Schedule.MonthDay -gt 0) {
            $HrSched = "Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
        elseif ($Schedule.MonthDay -eq 0) {
            $HrSched = "Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
        elseif ($Schedule.WeekOrder -gt 0) {
            switch ($Schedule.WeekOrder) {
                0 { $order = 'last' }
                1 { $order = 'first' }
                2 { $order = 'second' }
                3 { $order = 'third' }
                4 { $order = 'fourth' }
            }
            $HrSched = "Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
    }
    elseif ($Schedule.HourDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.HourDuration) hours"
    } 
    elseif ($Schedule.MinuteDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.MinuteDuration) minutes"
    } 
    elseif ($Schedule.DayDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.DayDuration) days"
    }
    else {
        $HrSched = "Could not calculate schedule"
    }
    return $HrSched
}

##Recursively processes through all the conditions in a task sequence step.
function Process-TSConditions{
    param ($condition,$Level = 0)
    $prefix = ""
    for ($x=0; $x -lt $Level; $x++){$prefix="--TAB--" + $prefix}
    If($condition.osConditionGroup){
        $OSCondition = $condition.osConditionGroup.osExpressionGroup.name -join ", $($condition.osConditionGroup.type) "
        "$($prefix)Operating System Equals: $OSCondition"
        Remove-Variable OSCondition -ErrorAction Ignore
    }
    If($condition.expression){
        $expressions = $condition.expression
        foreach ($expression in $expressions){
            #$expression
            switch ($expression.type){
                'SMS_TaskSequence_WMIConditionExpression' {
                    foreach ($pair in $expression.variable){
                        if ($pair.name -eq 'Query'){
                            if ($pair.'#text' -like 'SELECT OsLanguage FROM Win32_OperatingSystem WHERE OsLanguage*'){
                                $lang=[int](($pair.'#text').Split('='))[1].Trim("`'")
                                "$($prefix)Operating System Language: $(([System.Globalization.Cultureinfo]::GetCultureInfo($lang)).DisplayName) ($lang)"
                            }else{
                                "$($prefix)WMI Query: " + $pair.'#text'
                            }
                        }
                    }
                }
                'SMS_TaskSequence_VariableConditionExpression'{
                    foreach ($pair in $expression.variable){
                        if ($pair.name -eq 'Operator'){
                            $ExpOperator = $pair.'#text'
                        }
                        if ($pair.name -eq 'Value'){
                            $ExpValue = $pair.'#text'
                        }
                        if ($pair.name -eq 'Variable'){
                            $ExpVariable = $pair.'#text'
                        }
                    }
                    "$($prefix)Task Sequence Variable: $ExpVariable $ExpOperator $ExpValue"
                    Remove-Variable ExpVariable,ExpOperator,ExpValue -ErrorAction Ignore
                }
                'SMS_TaskSequence_FileConditionExpression'{
                    If(('Path' -in ($expression.variable).name) -and ('DateTimeOperator' -notin ($expression.variable).name) -and ('VersionOperator' -notin ($expression.variable).name)){
                        "$($prefix)File Exists: " + ($expression.variable).'#text'
                    }else{
                        foreach ($pair in $expression.variable){
                            switch ($pair.name){
                                'DateTime'{$FileDate = Convert-UTCtoLocal($pair.'#text')}
                                'DateTimeOperator'{$FileDateOperator = $pair.'#text'}
                                'Path'{$FilePath = $pair.'#text'}
                                'Version'{$FileVersion = $pair.'#text'}
                                'VersionOperator'{$FileVersionOperator = $pair.'#text'}
                            }
                        }
                        #'DateTimeOperator' -in ($expression.variable).name
                        #'VersionOperator' -in ($expression.variable).name
                        "$($prefix)File: $FilePath     File Version: $FileVersionOperator $FileVersion     File Date: $FileDateOperator $FileDate"
                        Remove-Variable FileDate,FileDateOperator,FilePath,FileVersion,FileVersionOperator -ErrorAction Ignore
                    }
                }
                'SMS_TaskSequence_FolderConditionExpression'{
                    If(('Path' -in ($expression.variable).name) -and ('DateTimeOperator' -notin ($expression.variable).name)){
                        "$($prefix)Folder Exists: " + ($expression.variable).'#text'
                    }else{
                        foreach ($pair in $expression.variable){
                            switch ($pair.name){
                                'DateTime'{$FolderDate = Convert-UTCtoLocal($pair.'#text')}
                                'DateTimeOperator'{$FolderDateOperator = $pair.'#text'}
                                'Path'{$FolderPath = $pair.'#text'}
                            }
                        }
                        #'DateTimeOperator' -in ($expression.variable).name
                        #'VersionOperator' -in ($expression.variable).name
                        "$($prefix)Folder: $FolderPath     Folder Date: $FolderDateOperator $FolderDate"
                        Remove-Variable FolderPath,FolderDateOperator,FolderDate -ErrorAction Ignore
                    }
                }
                'SMS_TaskSequence_RegistryConditionExpression'{
                    foreach ($pair in $expression.variable){
                        Switch ($pair.name){
                            'Operator'{$RegOperator = $pair.'#text'}
                            'KeyPath'{$RegKeyPath = $pair.'#text'}
                            'Data'{$RegData = $pair.'#text'}
                            'Value'{$RegValue = $pair.'#text'}
                            'Type'{$RegType = $pair.'#text'}
                        }
                    }
                    "$($prefix)Registry Value: $RegKeyPath $RegValue ($RegType) $RegOperator $RegData"
                    Remove-Variable RegKeyPath,RegValue,RegType,RegOperator,RegData -ErrorAction Ignore
                }
                'SMS_TaskSequence_SoftwareConditionExpression'{
                    foreach ($pair in $expression.variable){
                        Switch ($pair.name){
                            'Operator'{$AppOperator = $pair.'#text'}
                            'ProductCode'{$AppProductCode = $pair.'#text'}
                            'ProductName'{$AppProductName = $pair.'#text'}
                            #'UpgradeCode'{$AppUpgradeCode = $pair.'#text'}
                            'Version'{$AppVersion = $pair.'#text'}
                        }
                    }
                    If ($AppOperator -eq 'AnyVersion'){
                        "$($prefix)Installed Software: Any Version of `"$AppProductName`""
                    }else{
                        "$($prefix)Installed Software: Exact Version of `"$AppProductName`", Version: $AppVersion, Product Code: $AppProductCode"
                    }
                    Remove-Variable AppOperator,AppProductCode,AppProductName,AppUpgradeCode,AppVersion -ErrorAction Ignore
                }
            }
        }
    }
    If($condition.operator){
        Switch($condition.operator.type){
        'or'{"$($prefix)-If any of these conditions are true"}
        'and'{"$($prefix)-If all of these conditions are true"}
        'not'{"$($prefix)-If none of these conditions are true"}
        }
        $Level = $Level + 1
        Process-TSConditions -condition $condition.operator -Level $Level
    }
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
                if ($node.condition){
                    $Conditions=(Process-TSConditions -condition $node.condition) -join "--CRLF--"
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
                If ($node.continueOnError -eq "true"){
                    $StepContinueError = "Yes"
                }else{
                    $StepContinueError = "No"
                }
                if($GroupName){
                    #"$($GroupName):  $($node.name) $($node.action)"
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="$GroupName";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="$($node.Action)";'Continue on Error'="$StepContinueError";'Status'="$StepStatus";'Conditions'="$Conditions"}
                }else{
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="N/A";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="$($node.Action)";'Continue on Error'="$StepContinueError";'Status'="$StepStatus";'Conditions'="$Conditions"}
                    #"$($node.name) $($node.action)"
                }
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                $TSStep
            }
            'subtasksequence'{
                foreach ($item in $node.defaultVarList.variable){
                    Switch ($item.property){
                        'TsName'{$NestTSName = $item.'#text'}
                        'TsPackageID'{$NestTSPackage = $item.'#text'}
                    }
                }
                if (-not [string]::IsNullOrEmpty($node.Description)){
                    $StepDescription = "$($node.Description)"
                }
                if ($node.condition){
                    $Conditions=(Process-TSConditions -condition $node.condition) -join "--CRLF--"
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
                If ($node.continueOnError -eq "true"){
                    $StepContinueError = "Yes"
                }else{
                    $StepContinueError = "No"
                }
                if($GroupName){
                    #"$($GroupName):  $($node.name) $($node.action)"
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="$GroupName";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="Run Task Sequence ($($node.Action)):--CRLF--$NestTSName ($NestTSPackage)";'Continue on Error'="$StepContinueError";'Status'="$StepStatus";'Conditions'="$Conditions"}
                }else{
                    $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="N/A";'Step Name'="$($node.Name)";'Description'="$StepDescription";'Action'="Run Task Sequence ($($node.Action)):--CRLF--$NestTSName ($NestTSPackage)";'Continue on Error'="$StepContinueError";'Status'="$StepStatus";'Conditions'="$Conditions"}
                    #"$($node.name) $($node.action)"
                }
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                $TSStep
            }
            'group'{
                $TSStepNumber++
                if ($node.condition){
                    $Conditions=(Process-TSConditions -condition $node.condition) -join "--CRLF--"
                }
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
                If ($node.continueOnError -eq "true"){
                    $StepContinueError = "Yes"
                }else{
                    $StepContinueError = "No"
                }
                #"Group: $($node.Name)"
                $TSStep = New-Object -TypeName psobject -Property @{'Group Name'="$($node.Name)";'Step Name'="N/A";'Description'="$StepDescription";'Action'="N/A";'Continue on Error'="$StepContinueError";'Status'="$StepStatus";'Conditions'="$Conditions"}
                Remove-Variable Conditions -ErrorAction Ignore
                Remove-Variable StepDescription -ErrorAction Ignore
                $TSStep
                Process-TSSteps -Sequence $node -GroupName "$($node.Name)" -TSSteps $TSSteps -StepCounter $TSStepNumber
            }
            default{}
        }
    }
}

Function Set-AccountMask{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,
        HelpMessage="This is the Account that you want to mask the name with ***.")]
        [string]$Account
    )
    Begin{}
        Process{
            IF ($Account -match '\\'){
            $SamAcct = $Account.Split('\')[1]
            $MaskAcct = "$($SamAcct.substring(0,$SamAcct.Length/2))"+"*"*($SamAcct.Length/2+1)
            $MaskAcct = $Account -replace "$SamAcct","$MaskAcct"
        }else{
            $SamAcct = $Account
            $MaskAcct = "$($SamAcct.substring(0,$SamAcct.Length/2))"+"*"*($SamAcct.Length/2+1)
            $MaskAcct = $Account -replace "$SamAcct","$MaskAcct"        
        }
        $MaskAcct
    }
    End{}
}

Function Get-PWCMDiscoveryMethod {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteServer = $env:COMPUTERNAME,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteName = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM\CcmEval -Name LastSiteCode -ErrorAction SilentlyContinue).LastSiteCode,
        [Parameter(Mandatory = $false)]
        [ValidateSet('ActiveDirectoryForestDiscovery', 'ActiveDirectoryGroupDiscovery', 'ActiveDirectorySystemDiscovery', 'ActiveDirectoryUserDiscovery', 'NetworkDiscovery', 'HeartbeatDiscovery')]
        [ValidateNotNullOrEmpty()]
        $DiscoveryMethod
    )
    
    switch ($DiscoveryMethod) {
        'ActiveDirectoryForestDiscovery' { $DMX = 1 }
        'ActiveDirectoryGroupDiscovery' { $DMX = 2 }
        'ActiveDirectorySystemDiscovery' { $DMX = 3 }
        'ActiveDirectoryUserDiscovery' { $DMX = 4 }
        'NetworkDiscovery' { $DMX = 5 }
        'HeartbeatDiscovery' { $DMX = 6 }
        Default { $DMX = 7 }
    }

    #region AD Group Discovery
    If ($dmx -eq 2 -or $DMX -eq 7) {
        $ADGDHash = @{ }
        $ADGroupDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_SECURITY_GROUP_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        $ADGDHash.Add('DiscoveryMethod', 'Active Directory Group Discovery')
        foreach ($Prop in $ADGroupDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('IncrementalSyncState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('IncrementalSyncState', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    $IncSchedule = Get-HumanReadableSchedule -Schedule $schedule
                    $ADGDHash.Add('IncrementalSyncSchedule', "$IncSchedule")
                }
                'Full Sync Schedule' {
                    $GDFullSched = Convert-CMSchedule $prop.Value1
                    $FullSchedule = Get-HumanReadableSchedule -Schedule $GDFullSched
                    $ADGDHash.Add('FullSyncSchedule', "$FullSchedule")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADGDHash.Add('DiscoveryState', 'Enabled')
                    }
                    else {
                        $ADGDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Discover DG Membership' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('DistributionGroupDiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('DistributionGroupDiscoveryState', 'Disabled')
                    }
                }
                'Enable Filtering Expired Logon' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('FilterExpiredLogonState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('FilterExpiredLogonState', 'Disabled')
                    }
                }
                'Days Since Last Logon' {
                    $ADGDHash.Add('FilterExpiredLogonTime', $Prop.Value)
                }
                'Enable Filtering Expired Password' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('FilterExpiredPasswordState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('FilterExpiredPasswordState', 'Disabled')
                    }
                }
                'Days Since Last Password Set' {
                    $ADGDHash.Add('FilterExpiredPasswordTime', $Prop.Value)
                }
            }    
        }
        $adgroupd = @()
        $ADGroupSearch = @()
        $ADGroupSearchCred = @()
        $start = 0
        foreach ($List in $ADGroupDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { $two = $value }
                            3 { $three = $value }
                            4 {
                                $adgroupd += [pscustomobject]@{'Domain' = $one; 'val1' = $two; 'val2' = $three; 'val3' = $value }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'Search Bases:*' {
                    $SearchDomain = $List.PropertyListName -replace 'Search Bases:', ''
                    $ADGroupSearch += [pscustomobject]@{'Domain' = $SearchDomain; 'SearchBase' = "$($List.Values -join ';')" }
                }
                'AD Accounts:*' {
                    $AccountDomain = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADGroupSearchCred += [pscustomobject]@{'Domain' = $AccountDomain; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        $ADGroupSearchLocations = @()
        foreach ($ADGD in $adgroupd) {
            If ($ADGD.val1 -eq 0) {
                $ADGDType = 'Location'
            }
            else {
                $ADGDType = 'Groups'
            }
            if ($ADGD.val2 -eq 0) {
                $ADGDRecursive = 'Yes'
            }
            else {
                $ADGDRecursive = 'No'
                IF ($ADGDType -like 'Groups') { $ADGDRecursive = 'Not Applicable' }
            }
            $ADGDAccount = 'Site Server'
            foreach ($account in $ADGroupSearchCred) {
                If ($account.Domain -like $ADGD.Domain) {
                    $ADGDAccount = "$($account.Account)"
                }
            }
            foreach ($SB in $ADGroupSearch) {
                If ($SB.Domain -like $ADGD.Domain) {
                    $ADGDSearchBase = "$($SB.SearchBase)"
                }
            }
            $ADGroupSearchLocations += [pscustomobject]@{'Name' = "$($ADGD.Domain)"; 'Type' = "$ADGDType"; 'Recursive' = "$ADGDRecursive"; 'Account' = "$ADGDAccount"; 'SearchBase' = $ADGDSearchBase }
        }
        $ADGDHash.Add('SearchLocations', $ADGroupSearchLocations)
        [PSCustomObject]$ADGDHash
    }
    #endregion AD Group Discovery

    #region AD Forest Discovery
    If ($dmx -eq 1 -or $DMX -eq 7) {
        $ADFDHash = @{ }
        $ADFDHash.Add('DiscoveryMethod', 'Active Directory Forest Discovery')
        $ADForestDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_FOREST_DISCOVERY_MANAGER|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        foreach ($Prop in $ADForestDiscovery.Props) {
            switch ($Prop.PropertyName) {
                'Startup Schedule' {
                    $FDSchedule = Convert-CMSchedule $prop.Value1
                    $ADFDHash.Add('SyncSchedule', "$(Get-HumanReadableSchedule -Schedule $FDSchedule)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADFDHash.Add('DiscoveryState', 'Enabled')
                    }
                    else {
                        $ADFDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Enable AD Site Boundary Creation' {
                    If ($Prop.Value -eq 1) {
                        $ADFDHash.Add('ADSiteBoundaryCreation', 'Enabled')
                    }
                    Else {
                        $ADFDHash.Add('ADSiteBoundaryCreation', 'Disabled')
                    }
                }
                'Enable Subnet Boundary Creation' {
                    If ($Prop.Value -eq 1) {
                        $ADFDHash.Add('SubnetBoundaryCreation', 'Enabled')
                    }
                    Else {
                        $ADFDHash.Add('SubnetBoundaryCreation', 'Disabled')
                    }
                }
            }    
        }
        [PSCustomObject]$ADFDHash
    }
    #endregion AD Forest Discovery

    #region Heartbeat Discovery
    If ($dmx -eq 6 -or $DMX -eq 7) { 
        $HBDHash = @{ }
        $HBDHash.Add('DiscoveryMethod', 'Heartbeat Discovery')
        $HeartbeatDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_ClientConfig WHERE FileType=2 AND ItemName='Client Properties' AND ItemType='Client Configuration'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        Foreach ($prop in $HeartbeatDiscovery.Props) {
            switch ($prop.PropertyName) {
                'DDR Refresh Interval' {
                    $schedule = Convert-CMSchedule $prop.Value2
                    If ($schedule.DaySpan -ne 0){
                        If(($schedule.DaySpan % 7) -eq 0){
                            $weeks = $schedule.DaySpan/7
                            $DDRInterval = "Every $weeks week(s)"
                        } else {
                            $DDRInterval = "Every $($schedule.DaySpan) day(s)"
                        }
                    } elseif ($schedule.HourSpan -ne 0) {
                        $DDRInterval = "Every $($schedule.HourSpan) hour(s)"
                    }
                    $HBDHash.Add('DDRRefreshInterval', "$DDRInterval")
                }
                'Enable Heartbeat DDR' { 
                    If ($Prop.Value -eq 1) {
                        $HBDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $HBDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
            }
        }
        [PSCustomObject]$HBDHash
    }
    #endregion Heartbeat Discovery

    #region Network Discovery
    If ($dmx -eq 5 -or $DMX -eq 7) {
        $NDHash = @{ }
        $NDHash.Add('DiscoveryMethod', 'Network Discovery')
        $NetworkDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_NETWORK_DISCOVERY|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        Foreach ($prop in $NetworkDiscovery.Props) {
            switch ($prop.PropertyName) {
                'Discovery Enabled' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Subnet Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalSubnets', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalSubnets', 'Disabled')
                    }
                }
                'Domain Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalDomain', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalDomain', 'Disabled')
                    }
                }
                'Router Hop Count' {
                    $NDHash.Add('SNMPMaxHopCount', $prop.Value1)
                }
                'DHCP Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalDHCP', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalDHCP', 'Disabled')
                    }
                }
                'ICMP Ping Timeout' {
                    If ($Prop.Value1 -eq 2000) {
                        $NDHash.Add('SlowNetwork', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('SlowNetwork', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    If ($prop.Value1) {
                        $SchedVal = $prop.Value1 -split ''
                        $schedules = foreach ($bb in $SchedVal) {
                            $cc = "$cc$bb"
                            If ($cc.length -eq 16) {
                                $cc
                                $cc = ''
                            }
                        }
                        Remove-Variable cc, bb
                        $NDschedules = @()
                        foreach ($schedule in $schedules) {
                            $Sched = Convert-CMSchedule $schedule
                            $NDschedules += "$(Get-HumanReadableSchedule -Schedule $Sched)"
                        }
                        $NDHash.Add('DiscoverySchedule', $NDschedules)
                    }
                }
            }
        }
        $NDSubnets = @()
        $NDDomains = @()
        $NDSnmp = @()
        $NDSnmpDevices = @()
        $NDDHCPServers = @()
        foreach ($List in $NetworkDiscovery.PropLists) {
            Switch ($List.PropertyListName) {
                'Subnet Include' {
                    Foreach ($Value in $List.Values) {
                        $SN = $Value.Split(' ')
                        $NDSubnets += [pscustomobject]@{'Network' = "$($SN[0])"; 'Subnet Mask' = "$($SN[1])"; 'Search' = 'Enabled' }
                    }
                }
                'Subnet Exclude' {
                    Foreach ($Value in $List.Values) {
                        $SN = $Value.Split(' ')
                        $NDSubnets += [pscustomobject]@{'Network' = "$($SN[0])"; 'Subnet Mask' = "$($SN[1])"; 'Search' = 'Disabled' }
                    }
                }
                'Domain Include' {
                    Foreach ($Value in $List.Values) {
                        $NDDomains += [pscustomobject]@{'Domain' = "$Value"; 'Search' = 'Enabled' }
                    }
                }
                'Domain Exclude' {
                    Foreach ($Value in $List.Values) {
                        $NDDomains += [pscustomobject]@{'Domain' = "$Value"; 'Search' = 'Disabled' }
                    }
                }
                'Community Names' {
                    foreach ($Value in $List.Values) {
                        $NDSnmp += "$Value"
                    }
                }
                'Address Include' {
                    foreach ($Value in $List.Values) {
                        $NDSnmpDevices += "$Value"
                    }
                }
                'DHCP Include' {
                    foreach ($Value in $List.Values) {
                        $NDDHCPServers += "$Value"
                    }
                }
                'Network Discovery Protocols' {
                    #{DHCP, OSPF} = always set to this.
                }
                'Address Discovery Protocols' {
                    Switch ($List.values -join ',') {
                        'OSPF,RIP' { $NDADP = 1 }
                        'OSPF,RIP,NETBIOS,DHCP' { $NDADP = 2 }
                    }
                }
                'Address Validation Protocols' {
                    Switch ($List.values -join ',') {
                        'ICMP,NAME_RESOLVE' { $NDAVP = 1 }
                        'ICMP,NAME_RESOLVE,NETBIOS' { $NDAVP = 2 }
                    }
                }
            }
        }
        If ($NDADP -eq 1) {
            $NDType = 'Topology'
        }
        else {
            if ($NDADP -eq 2) {
                If ($NDAVP -eq 1) {
                    $NDType = 'Topology and client'
                }
                else {
                    if ($NDAVP -eq 2) {
                        $NDType = 'Topology, client, and client operating system'
                    }
                }
            }
        }
        $NDHash.Add('DiscoveryType', $NDType)
        $NDHash.Add('SearchSubnets', $NDSubnets)
        $NDHash.Add('SearchDomains', $NDDomains)
        $NDHash.Add('SearchSNMPCommunities', $NDSnmp)
        $NDHash.Add('SearchSNMPDevices', $NDSnmpDevices)
        $NDHash.Add('SearchDHCPServers', $NDDHCPServers)
        [PSCustomObject]$NDHash
    }
    #endregion Network Discovery

    #region AD System Discovery
    If ($dmx -eq 3 -or $DMX -eq 7) {
        $ADSDHash = @{ }
        $ADSDHash.Add('DiscoveryMethod', 'Active Directory System Discovery')
        $ADSystemDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_SYSTEM_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        foreach ($Prop in $ADSystemDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('IncrementalSync', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('IncrementalSync', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    If ($schedule.MinuteSpan -ne 0) {
                        $ADSDHash.Add('IncrementalSyncSchedule', $schedule.MinuteSpan)
                    }
                }
                'Full Sync Schedule' {
                    $SDFullSched = Convert-CMSchedule $prop.Value1
                    $ADSDHash.Add('FullSyncSchedule', "$(Get-HumanReadableSchedule -Schedule $SDFullSched)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADSDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Enable Filtering Expired Logon' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('FilterExpiredLogon', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('FilterExpiredLogon', 'Disabled')
                    }
                }
                'Days Since Last Logon' {
                    $ADSDHash.Add('FilterExpiredLogonTime', $Prop.Value)
                }
                'Enable Filtering Expired Password' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('FilterExpiredPassword', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('FilterExpiredPassword', 'Disabled')
                    }
                }
                'Days Since Last Password Set' {
                    $ADSDHash.Add('FilterExpiredPasswordTime', $Prop.Value)
                }
            }    
        }
        $ADContainerDiscovery = @()
        $AdditionalADAttributes = @()
        $ADContainerSearchCreds = @()
        $ADContainerExclusions = @()
        foreach ($List in $ADSystemDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { 
                                If ($value -eq 0) {
                                    $two = 'Yes'
                                }
                                else {
                                    $two = 'No'
                                }
                            }
                            3 {
                                If ($value -eq 1) {
                                    $three = 'Excluded'
                                }
                                else {
                                    $three = 'Included'
                                }
                                $ADContainerDiscovery += [pscustomobject]@{'Container' = $one; 'Recursive' = $two; 'Groups' = $three; 'Account' = 'Site Server'; 'Exclusions' = 'None' }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'AD Attributes' {
                    $AdditionalADAttributes += $List.values
                }
                'AD Containers Exclusions' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 {
                                $ADContainerExclusions += [pscustomobject]@{'Container' = $one; 'Exclusions' = $value }
                                Remove-Variable one
                                $start = 0
                            }
                        }
                    }
                }
                'AD Accounts:*' {
                    $AccountContainer = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADContainerSearchCreds += [pscustomobject]@{'Container' = $AccountContainer; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        foreach ($account in $ADContainerSearchCreds) {
            foreach ($Container in $ADContainerDiscovery) {
                If ($account.Container -like $Container.Container) {
                    $Container.Account = $account.Account
                }
            }
        }
        foreach ($Exclusion in $ADContainerExclusions) {
            foreach ($Container in $ADContainerDiscovery) {
                If ($Exclusion.Container -like $Container.Container) {
                    $Container.Exclusions = $Exclusion.Exclusions
                }
            }
        }
        $ADSDHash.Add('ActiveDirectoryContainers', $ADContainerDiscovery)
        $ADSDHash.Add('ActiveDirectoryAttributes', $AdditionalADAttributes)
        [PSCustomObject]$ADSDHash
    }
    #endregion AD System Discovery

    #region AD User Discovery
    If ($dmx -eq 4 -or $DMX -eq 7) {
        $ADUDHash = @{ }
        $ADUDHash.Add('DiscoveryMethod', 'Active Directory User Discovery')
        $ADUserDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_USER_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
        foreach ($Prop in $ADUserDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADUDHash.Add('IncrementalSync', 'Enabled')
                    }
                    Else {
                        $ADUDHash.Add('IncrementalSync', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    If ($schedule.MinuteSpan -ne 0) {
                        $ADUDHash.Add('IncrementalSyncSchedule', $schedule.MinuteSpan)
                    }
                }
                'Full Sync Schedule' {
                    $SDFullSched = Convert-CMSchedule $prop.Value1
                    $ADUDHash.Add('FullSyncSchedule', "$(Get-HumanReadableSchedule -Schedule $SDFullSched)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADUDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADUDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
            }    
        }
        $ADUserContainerDiscovery = @()
        $AdditionalADUserAttributes = @()
        $ADUserSearchCreds = @()
        $ADUserContainerExclusions = @()
        foreach ($List in $ADUserDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { 
                                If ($value -eq 0) {
                                    $two = 'Yes'
                                }
                                else {
                                    $two = 'No'
                                }
                            }
                            3 {
                                If ($value -eq 1) {
                                    $three = 'Excluded'
                                }
                                else {
                                    $three = 'Included'
                                }
                                $ADUserContainerDiscovery += [pscustomobject]@{'Container' = $one; 'Recursive' = $two; 'Groups' = $three; 'Account' = 'Site Server'; 'Exclusions' = 'None' }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'AD Attributes' {
                    $AdditionalADUserAttributes += $List.values
                }
                'AD Containers Exclusions' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 {
                                $ADUserContainerExclusions += [pscustomobject]@{'Container' = $one; 'Exclusions' = $value }
                                Remove-Variable one
                                $start = 0
                            }
                        }
                    }
                }
                'AD Accounts:*' {
                    $AccountContainer = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADUserSearchCreds += [pscustomobject]@{'Container' = $AccountContainer; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        foreach ($account in $ADUserSearchCreds) {
            foreach ($Container in $ADUserContainerDiscovery) {
                If ($account.Container -like $Container.Container) {
                    $Container.Account = $account.Account
                }
            }
        }
        foreach ($Exclusion in $ADUserContainerExclusions) {
            foreach ($Container in $ADUserContainerDiscovery) {
                If ($Exclusion.Container -like $Container.Container) {
                    $Container.Exclusions = $Exclusion.Exclusions
                }
            }
        }
        $ADUDHash.Add('ActiveDirectoryContainers', $ADUserContainerDiscovery)
        $ADUDHash.Add('ActiveDirectoryAttributes', $AdditionalADUserAttributes)
        [PSCustomObject]$ADUDHash
    }
    #endregion User Discovery
}

Function Get-PWCMPhasedDeployment {
    <#
    .SYNOPSIS
    Gets the Details of a phased deployment.
    
    .DESCRIPTION
    Gets the Details of a phased deployment in CM.  If no deployment specified, all are found and returned.
    
    .PARAMETER Name
    Finds the phased deployment by its Name.
    
    .PARAMETER PhasedDeploymentID
    Finds the phased deployment by its deployment ID, GUID.
    
    .PARAMETER DeploymentObjectID
    Finds the phased deployment by the object that this deployment is targeting (Application, Task Sequence, or Update Group).
    
    .PARAMETER SiteServer
    The primary site server (Where the WMI queries will be directed).
    
    .PARAMETER SiteName
    Your SCCM Site code (MMS).
    
    .EXAMPLE
    Get-PWCMPhasedDeployment -Name '7-Zip structured Deployment'

    .EXAMPLE
    Get-PWCMPhasedDeployment -PhasedDeploymentID 'FD114F32-752F-4DA9-AB2A-1B388B0E5807'
    
    .EXAMPLE
    Get-PWCMPhasedDeployment -DeploymentObjectID MMS0000213

    .NOTES
    General notes
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName='ByName')]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $false, ParameterSetName='ByDeploymentID')]
        [ValidateNotNullOrEmpty()]
        [string]$PhasedDeploymentID,
        [Parameter(Mandatory = $false, ParameterSetName='ByObjectId')]
        [ValidateNotNullOrEmpty()]
        [string]$DeploymentObjectID,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteServer = $env:COMPUTERNAME,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteName = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM\CcmEval -Name LastSiteCode -ErrorAction SilentlyContinue).LastSiteCode
    )
    If ($Name){
        $PDs = Get-WmiObject -Query "SELECT * FROM SMS_PhasedDeployment where Name Like '$Name'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
    }
    elseIf ($PhasedDeploymentID){
        $PDs = Get-WmiObject -Query "SELECT * FROM SMS_PhasedDeployment where PhasedDeploymentID Like '$PhasedDeploymentID'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
    }
    elseIf ($DeploymentObjectID){
        Write-Verbose "Get-WmiObject -Query `"SELECT * FROM SMS_PhasedDeployment where DeploymentObjectID Like `'$DeploymentObjectID`'`" -Namespace ROOT\SMS\site_$SiteName -ComputerName $SiteServer"
        $PDs = Get-WmiObject -Query "SELECT * FROM SMS_PhasedDeployment where DeploymentObjectID Like '$DeploymentObjectID'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
    } else {
        $PDs = Get-WmiObject -Query "SELECT * FROM SMS_PhasedDeployment" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer
    }
    foreach ($PD in $PDs){
        $PD.Get()
        switch ($PD.DeploymentObjectType) {
            0 {
                $Type = 'Task Sequence'
                $DeploymentObjectName = (Get-WmiObject -Query "SELECT * FROM SMS_TaskSequencePackage WHERE PackageID='$($PD.DeploymentObjectID)'" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer).Name
            }
            1 {
                $Type = 'Software Update'
                $DeploymentObjectName = (Get-WmiObject -Query "SELECT * FROM SMS_AuthorizationList WHERE CI_ID=$($PD.DeploymentObjectID)" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer).LocalizedDisplayName
            }
            2 {
                $Type = 'Application'
                $DeploymentObjectName = (Get-WmiObject -Query "SELECT * FROM SMS_ApplicationLatest WHERE CI_ID=$($PD.DeploymentObjectID)" -Namespace "ROOT\SMS\site_$SiteName" -ComputerName $SiteServer).LocalizedDisplayName
            }
            Default { $Type = 'Unknown' }
        }
        $PD | Select-Object @{Name = "SmsProviderObjectPath"; Expression = {$_.__RELPATH}},Action,DeploymentObjectID,@{Name = "DeploymentObjectName"; Expression = {$DeploymentObjectName}},@{Name = "DeploymentObjectTypeName"; Expression = {$Type}},DeploymentObjectType,Description,EvaluatePhasedDeployment,@{Name = "LastEvaluateTime"; Expression = {$(Convert-UTCtoLocal($_.LastEvaluateTime))}},@{Name = "LastModifiedTime"; Expression = {$(Convert-UTCtoLocal($_.LastModifiedTime))}},Name,PhasedDeploymentDigest,PhasedDeploymentID,Version
    }
}

Function Get-PWCMPDPhases {
    <#
    .SYNOPSIS
    Gets the details on all the defined phases of a phased deployment.
    
    .DESCRIPTION
    Gets the details on all the defined phases of a phased deployment.  Depending on the type, there are different details available.
    
    .PARAMETER Name
    Finds the phased deployment by its Name.
    
    .PARAMETER PhasedDeploymentID
    Finds the phased deployment by its deployment ID, GUID.
    
    .PARAMETER DeploymentObjectID
    Finds the phased deployment by the object that this deployment is targeting (Application, Task Sequence, or Update Group).
    
    .PARAMETER SiteServer
    The primary site server (Where the WMI queries will be directed).
    
    .PARAMETER SiteName
    Your SCCM Site code (MMS).
    
    .EXAMPLE
    Get-PWCMPDPhases -Name '7-Zip structured Deployment'

    .EXAMPLE
    Get-PWCMPDPhases -PhasedDeploymentID 'FD114F32-752F-4DA9-AB2A-1B388B0E5807'
    
    .EXAMPLE
    Get-PWCMPDPhases -DeploymentObjectID MMS0000213

    #>
    
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName='ByName')]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $true, ParameterSetName='ByDeploymentID')]
        [ValidateNotNullOrEmpty()]
        [string]$PhasedDeploymentID,
        [Parameter(Mandatory = $true, ParameterSetName='ByObjectId')]
        [ValidateNotNullOrEmpty()]
        [string]$DeploymentObjectID,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteServer = $env:COMPUTERNAME,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteName = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM\CcmEval -Name LastSiteCode -ErrorAction SilentlyContinue).LastSiteCode
    )
    If ($Name) {
        Try{
            $PhasedDeployment = Get-PWCMPhasedDeployment -Name $Name -SiteServer $SiteServer -SiteName $SiteName
            $phases = [XML]$PhasedDeployment.PhasedDeploymentDigest
            $DeploymentType = $PhasedDeployment.DeploymentObjectTypeName
        }
        Catch{
            Throw "Error : Could not find phased deployment by name: $Name"
        }
        if ([string]::IsNullOrEmpty($DeploymentType)) {
            Throw "Could not find phased deployment by name: $Name"
        }
    }
    If ($PhasedDeploymentID) {
        Try{
            $PhasedDeployment = Get-PWCMPhasedDeployment -PhasedDeploymentID $PhasedDeploymentID -SiteServer $SiteServer -SiteName $SiteName
            $phases = [XML]$PhasedDeployment.PhasedDeploymentDigest
            $DeploymentType = $PhasedDeployment.DeploymentObjectTypeName
        }
        Catch{
            Throw "Error : Could not find phased deployment by PhasedDeploymentID: $PhasedDeploymentID"
        }
        if ([string]::IsNullOrEmpty($DeploymentType)) {
            Throw "Could not find phased deployment by PhasedDeploymentID: $PhasedDeploymentID"
        }
    }
    If ($DeploymentObjectID) {
        try{
            $PhasedDeployment = Get-PWCMPhasedDeployment -DeploymentObjectID $DeploymentObjectID -SiteServer $SiteServer -SiteName $SiteName
            $phases = [XML]$PhasedDeployment.PhasedDeploymentDigest
            $DeploymentType = $PhasedDeployment.DeploymentObjectTypeName    
        }
        catch{
            Throw "Error : Could not find phased deployment by DeploymentObjectID: $DeploymentObjectID"
        }
        if ([string]::IsNullOrEmpty($DeploymentType)) {
            Throw "Could not find phased deployment by DeploymentObjectID: $DeploymentObjectID"
        }
    }
    Foreach ($Phase in $phases.PhasedDeployment.Phases.phase){
        IF ($Phase.SuccessCriterias.Criteria){
            $ht += @{'SuccessCriteria' = [pscustomobject](@{"$($Phase.SuccessCriterias.Criteria.CriteriaType)" = "$($Phase.SuccessCriterias.Criteria.CriteriaValue)"})}
        }
        If ($Phase.Targets.CollectionID){
            $ht += @{'Target' = "$($Phase.Targets.CollectionID)"}
        }
        IF ($Phase.PhaseId){
            $ht += @{'Name' = "$($Phase.PhaseId)"}
        }
        IF ($Phase.Order){
            $ht += @{'Order' = "$($Phase.Order)"}
        }
        If ($Phase.Condition.DaysAfterPrevPhaseSuccess){
            If ($Phase.Condition.DaysAfterPrevPhaseSuccess -ne -1){
                $ht += @{'DaysAfterPrevPhaseSuccess' = "$($Phase.Condition.DaysAfterPrevPhaseSuccess)"}
                $ht += @{'ManuallyStartPhase' = $false}
            } else {
                $ht += @{'DaysAfterPrevPhaseSuccess' = ''}
                $ht += @{'ManuallyStartPhase' = $true}
            }
        }
        If ($Phase.ThrottlingDays){
            $ht += @{'GraduallyAvailable' = "$($Phase.ThrottlingDays)"}
        }
        If ($Phase.Deployment){
            switch ($DeploymentType) {
                'Task Sequence' {
                    #Task sequence phased deployments do not have easy to use XML. So, let's pipe the configurable settings into a custom object.
                    $DeploymentXML = ([xml]$($Phase.Deployment.OuterXML)).Deployment
                    $Deploy += @{'DeadlineDelta' = "$($DeploymentXML.DeadlineDelta)"}
                    $Deploy += @{'DeadlineUnit' = "$($DeploymentXML.DeadlineUnit)"}
                    If ($DeploymentXML.OfferFlags -band 0x100000){
                        $Deploy += @{'MaintenanceWindowInstallOverride' = $true}
                    } else {
                        $Deploy += @{'MaintenanceWindowInstallOverride' = $false}
                    }
                    If ($DeploymentXML.OfferFlags -band 0x200000){
                        $Deploy += @{'MaintenanceWindowRestartOverride' = $true}
                    } else {
                        $Deploy += @{'MaintenanceWindowRestartOverride' = $false}
                    }
                    If ($DeploymentXML.OfferFlags -band 0x1000){
                        $Deploy += @{'PreDownloadContent' = $true}
                    } else {
                        $Deploy += @{'PreDownloadContent' = $false}
                    }
                    If ($DeploymentXML.OfferFlags -band 0x800000){
                        $Deploy += @{'UserExperience' = 'ShowAll'}
                    } else {
                        $Deploy += @{'UserExperience' = 'HideAll'}
                    }
                    If ($DeploymentXML.OfferFlags -band 0x20000){
                        $Deploy += @{'UseSiteDefaultDP' = $true}
                    } else {
                        $Deploy += @{'UseSiteDefaultDP' = $false}
                    }
                    If ($DeploymentXML.RemoteClientFlags -band 0x0100){
                        $Deploy += @{'UseNeighborDP' = $false}
                    } else {
                        $Deploy += @{'UseNeighborDP' = $true}
                    }
                    If ($DeploymentXML.RemoteClientFlags -band 0x0020){
                        $Deploy += @{'RemoteDeployOpt' = 'DownloadWhenNeeded'}
                    } else {
                        $Deploy += @{'RemoteDeployOpt' = 'DownloadBefore'}
                    }
                    If ($DeploymentXML.RemoteClientFlags -band 0x8000){
                        $Deploy += @{'WECommitAtDeadline' = $true}
                    } else {
                        $Deploy += @{'WECommitAtDeadline' = $false}
                    }
                    $ObjPD = [pscustomobject]$Deploy
                    $ht += @{'DeploymentXML' = $DeploymentXML}
                    $ht += @{'Deployment' = $ObjPD}
                    Remove-Variable Deploy
                    Remove-Variable ObjPD        
                }
                'Software Update' {
                    #Software updates phased deployments already have the data with them in XML.  So, just pass the xml.
                    $DeploymentXML = ([xml]$($Phase.Deployment.OuterXML)).Deployment
                    $ht += @{'DeploymentXML' = $DeploymentXML}
                    $ht += @{'Deployment' =  $DeploymentXML}
                }
                Default {}
            }
        }
        [pscustomobject]$ht
        Remove-Variable ht
    }
}

Function Get-pwCMCloudManagementGateway {
    <#
    .SYNOPSIS
    Gets the Details of a ConfigMgr Site's Cloud Management Gateway(s) 
    .PARAMETER SiteCode
    Your SCCM Site code (MMS).
    .PARAMETER FilePath
    This is the path of the documentation file.    
    .EXAMPLE
    Get-pwCMCloudManagementGateway -SiteCode 'MMS' -FilePath 'C:\Documents\CMDocumentation.html'
    .NOTES
    ========== Change Log History ==========
    - 2021/01/22 by Chad@ChadsTech.net / Chad.Simmons@CatapultSystems.com - Created
    === To Do / Proposed Changes ===
    #TODO: CMG: Role Endpoints
    #TODO: CMG: Validate multiple CMGs, multiple sites
    #TODO: CMG: Validate features added since ConfigMgr v1910
    #>
    param (
        [Parameter()][ValidateNotNullOrEmpty()][string]$SiteCode = $CMSite.SiteCode,
        [Parameter()][ValidateNotNullOrEmpty()][string]$FilePath = $FilePath
    )
    Write-ProgressEx -CurrentOperation 'Enumerating Configuration Manager Cloud Management Gateway(s)'
    Write-HTMLHeading -Text 'Cloud Management Gateway' -Level 3 -File $FilePath
    #select * from vSMS_Azure_CloudService Where SiteCode = "$SiteCode"
    $CMGs = Get-CMCloudManagementGateway | Where-Object { $_.SiteCode -eq "$SiteCode"}
    ForEach ($CMG in $CMGs) {
        Write-Debug -Message $CMG
        #transform values to meaningful names
        $CMG | Add-Member -MemberType NoteProperty -Name 'ResourceGroup' -Value $CMG.AffinityGroup
        $CMG | Add-Member -MemberType NoteProperty -Name 'DeploymentModelName' -Value $CMG.DeploymentModel
        Switch ($CMG.DeploymentModel) {
            1 { $CMG.DeploymentModelName = 'Azure Resource Manager'}
        }
        $CMG | Add-Member -MemberType NoteProperty -Name 'AzureEnvironment' -Value $CMG.AzureEnvironmentId
        Switch ($CMG.AzureEnvironmentId) {
            1 { $CMG.AzureEnvironment = 'AzurePublicCloud'}
        }
        $CMG | Add-Member -MemberType NoteProperty -Name 'StateName' -Value $CMG.State
        Switch ($CMG.State) {
            0 { $CMG.StateName = 'Ready'}
            4 { $CMG.StateName = 'Starting'}
            6 { $CMG.StateName = 'Stopped'}
        }
        #transform object to multicolumn array and seperate known configuration items from unknown
        $CMGarray = @()
        $CMGarrayAdditional = @()
        ForEach ($prop in $($CMG | Get-Member -MemberType Property,NoteProperty | Where-Object { $_.Name -notlike '_*'})) { 
            If ($Prop.Name -in @($('Name,Fqdn,ServiceCName,StateName,StatusDetails,NALPath,AADApplicationId,AzureServiceID,AzureEnvironment,SubscriptionID,ResourceGroup,CreationTime,DeploymentModelName,DeploymentSlot,NumberOfInstances,Region,RootCertificate,SubCACertificate,StorageServiceName,StorageUsage,StorageQuotaInGB,StorageQuotaGrow,StorageWarningThreshold,StorageCriticalThreshold,TrafficOutUsage,TrafficOutInGB,TrafficOutStopService,TrafficWarningThreshold,TrafficCriticalThreshold' -split ','))) {
                $CMGarray += $(New-Object -TypeName PSObject -Property $([ordered]@{Configuration = $prop.name; Value = $CMG.$($prop.name)}))
            } ElseIf ($Prop.Name -notIn @('SerializedServiceCertificate','ServiceCertificate','ProxyCertificateConfig','AffinityGroup','DeploymentModel','AzureEnvironmentId','State')) {
                $CMGarrayAdditional += $(New-Object -TypeName PSObject -Property $([ordered]@{Configuration = $prop.name; Value = $CMG.$($prop.name)}))
            }
        }
        Write-HTMLParagraph -Text 'Cloud Management Gateway details.' -Level 3 -File $FilePath
        Write-HtmlTable -Border 1 -Level 3 -File $FilePath -InputObject $CMGarray
        Write-HTMLParagraph -Text 'Cloud Management Gateway additional details.' -Level 3 -File $FilePath
        Write-HtmlTable -Border 1 -Level 3 -File $FilePath -InputObject $CMGarrayAdditional
    }
    If (-not($CMGs)) { Write-HTMLParagraph -Text 'Cloud Management Gateway not detected.' -Level 3 -File $FilePath }
    Write-ProgressEx -CurrentOperation 'Completed Configuration Manager Cloud Management Gateway(s)'
}

function Convert-Bool2Text {
    #Simple function to convert a 1 to true and a 0 to false.  Or, just return the value if it isn't 1 or 0.
    param (
        $Value
    )
    switch ($Value) {
        0 { 'False' }
        1 { 'True' }
        Default {$Value}
    }
}

Function Get-pwCMManagementPoints {
    <#
    .SYNOPSIS
    Gets the Details of a ConfigMgr Site's Management Point(s)
    .PARAMETER SiteCode
    Your SCCM Site code (MMS).
    .PARAMETER FilePath
    This is the path of the documentation file.    
    .EXAMPLE
    Get-pwCMManagementPoints -SiteCode 'MMS' -FilePath 'C:\Documents\CMDocumentation.html'
    .NOTES
    ========== Change Log History ==========
    - 2021/01/26 by Chad@ChadsTech.net / Chad.Simmons@CatapultSystems.com - Moved to function, added 
    - 2018/03/07 by Paul Wetter - Created
    === To Do / Proposed Changes ===
    #TODO: MP: mobile device and Mac computer access
    #>
    param (
        [Parameter()][ValidateNotNullOrEmpty()][string]$SiteCode = $CMSite.SiteCode,
        [Parameter()][ValidateNotNullOrEmpty()][string]$FilePath = $FilePath
    )
    Write-ProgressEx -CurrentOperation "Management Points"
    $CMManagementPoints = Get-CMManagementPoint -SiteCode $SiteCode
    Write-HTMLHeading -Text "Summary of Management Points for Site $SiteCode" -PageBreak -Level 2 -File $FilePath
    ForEach ($CMManagementPoint in $CMManagementPoints) {
        Write-Verbose "$(Get-Date):   Management Point: $($CMManagementPoint)"
    
        $MPText = @()
        $MPName = $CMManagementPoint.NetworkOSPath.Split('\\')[2]
        Write-ProgressEx -CurrentOperation "Management Point Name: $MPName" -Activity "Managent Points" -Status "Collecting from DB" -Id 2
        $MPSslState = $CMManagementPoint.SslState
        $MPText += "Client connections: $(If ($MPSslState -eq 0) { 'HTTP' } Else { 'HTTPS' })"
        Write-Debug "$(Get-Date): Management Point SslState: $MPSslState"
    
        $MPCMGTraffic = [bool]($CMManagementPoint.props|Where-Object {$_.PropertyName -eq 'AllowProxyTraffic'}).value
        $MPText += "Allow Configuration Manager cloud management gateway traffic: $MPCMGTraffic"
        Write-Debug "$(Get-Date): Management Point CMG traffic: $MPCMGTraffic"
        
        $MPIntranet = [bool]($CMManagementPoint.props|Where-Object {$_.PropertyName -eq "MPIntranetFacing"}).value
        $MPInternet = [bool]($CMManagementPoint.props|Where-Object {$_.PropertyName -eq "MPInternetFacing"}).value
        $MPText += "Allow intranet connections: $MPIntranet"
        $MPText += "Allow Internet connections: $MPInternet"
        Write-Debug "$(Get-Date): Management Point Internet: $MPInternet Intranet: $MPIntranet"
    
        $MPText += "Allow mobile devices and Mac computers to use this management point: unknown"
    
        $MPHealthAlert = [bool](Get-CMAlert -Name "Not healthy alert for site role: Management point on '$MPName'").Enabled
        $MPText += "Alert Generate alert when management point is not healthy: $MPHealthAlert"
        Write-Debug "$(Get-Date): Management Point healthy alert: $MPHealthAlert"
    
        If (($CMManagementPoint.props | Where-Object {$_.PropertyName -eq 'UseSiteDatabase'}).value -eq 1) { $UseSiteDB = 'Site DB' } Else { $UseSiteDB = 'database replica' }
        Write-Debug "$(Get-Date): Management Point database: $UseSiteDB"
        $MPText += "Database to use with this management point: $UseSiteDB"
        If ($UseSiteDB -eq 'database replica') {       
            $MPText += "Database server: $(($CMManagementPoint.props | Where-Object {$_.PropertyName -eq 'SQLServerName'}).Value1)"
            $MPText += "Database name: $(($CMManagementPoint.props | Where-Object {$_.PropertyName -eq 'DatabaseName'}).Value1)"
            If (-not [String]::IsNullOrEmpty(($CMManagementPoint.Props | Where-Object { $_.PropertyName -eq 'UserName' }).Value1)) {
                $MPText += "Connection account: $(($CMManagementPoint.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value1)"
            } Else {
                $MPText += "Connection account: computer account of the management point"
            }
        }
        Write-HtmlList -InputObject $MPText -Title "Management Point Name: <B>$MPName</B>" -Level 2 -File $FilePath
        Remove-Variable -Name MPText, MPIntranet, MPInternet, MPHealthAlert, UseSiteDB, MPCMGTraffic
        If (-not($PSBoundParameters.ContainsKey('SkipRemoteServerDetails'))) {
            Write-Debug "$(Get-Date):   Test-Path -Path `"filesystem::\\$MPName\C$`""
            Push-Location -Path 'C:'
            $PathTest = Test-Path -Path "filesystem::\\$MPName\C$"
            Write-Debug "$(Get-Date):   Testing Access to Management Point: $MPName -- $PathTest"
            If (Test-Path -Path "filesystem::\\$MPName\C$") {$CMMPServerName=$MPName}
            Pop-Location
        }
    }
    Write-Debug "$(Get-Date):   Default Management Point: $CMMPServerName"
    Write-HtmliLink -ReturnTOC -File $FilePath
    Write-ProgressEx -CurrentOperation "Complete" -Activity "Management Points" -Status "Complete" -Id 2 -Completed
}

Function Get-pwCMSiteServicingUpdates {
    <#
    .SYNOPSIS
    Gets the Details of a ConfigMgr Site Servicing Updates
    .PARAMETER SiteCode
    Your SCCM Site code (MMS).
    .PARAMETER FilePath
    This is the path of the documentation file.    
    .EXAMPLE
    Get-pwCMSiteServicingUpdates -SiteCode 'MMS' -FilePath 'C:\Documents\CMDocumentation.html'
    .NOTES
    ========== Change Log History ==========
    - 2021/01/22 by Chad@ChadsTech.net / Chad.Simmons@CatapultSystems.com - Moved to function, updated sort, scopped to SiteCode
    - 2018/03/07 by Paul Wetter - Created
    === To Do / Proposed Changes ===
    #>
    param (
        [Parameter()][ValidateNotNullOrEmpty()][string]$SiteCode = $CMSite.SiteCode,
        [Parameter()][ValidateNotNullOrEmpty()][string]$FilePath = $FilePath
    )
    Write-ProgressEx -CurrentOperation 'Enumerating Configuration Manager Update Status and History'
    Write-HTMLHeading -Text 'Update Status and History' -Level 3 -File $FilePath
    Write-HTMLParagraph -Text 'Below is a history of updates that have been made available to this Site.  It includes information for if, or when, they were installed.  Some older updates may be listed as ready to install, however, they were never installed nor will they be available to install as they are superseded by the newer updates.' -Level 3 -File $FilePath
    $SiteUpdateHistory = Get-CMSiteUpdateHistory | Where-Object { $_.ManagedObject -like "*\root\sms\site_$($CMSite.SiteCode)`:SMS_CM_UpdatePackagesHistory*"} | Sort-Object FullVersion | Select-Object Name, @{N='Version'; E={$_.FullVersion}}, State, @{N='Installed Date'; E={$_.LastUpdateTime}}, @{N='Release Date'; E={$_.DateReleased}}
    If (-not [string]::IsNullOrEmpty($SiteUpdateHistory)) {
        ForEach ($SiteUpdate in $SiteUpdateHistory) {
            Switch ($SiteUpdate.State) {
                65538 { $SiteUpdate.State = 'Checking prerequisites' }
               131074 { $SiteUpdate.State = 'Prerequisite check passed'}
               131075 { $SiteUpdate.State = 'Prerequisite check passed with warnings' }
               196607 { $SiteUpdate.State = 'Prerequisite check failed' }
               196609 { $SiteUpdate.State = 'Installing' }
               196612 { $SiteUpdate.State = 'Installed' }
               262146 { $SiteUpdate.State = 'Ready to Install' }
               327682 { $SiteUpdate.State = 'Available to Download' }
              default { $SiteUpdate.State = "Other ($($SiteUpdate.State))" }
            }
          If ($SiteUpdate.State -ne 'Installed') { $SiteUpdate.'Installed Date' = 'N/A' }
        }
        Write-HtmlTable -InputObject $SiteUpdateHistory -Border 1 -Level 3 -File $FilePath
    }
    Remove-Variable -Name SiteUpdateHistory -ErrorAction SilentlyContinue -Force -WhatIf:$false
    Write-ProgressEx -CurrentOperation 'Completed Configuration Manager Update Status and History'    
}

Function Write-ProgressEx {
    [CmdletBinding()]
    param(
        [string]$Activity = $Progress.Activity,
        [string]$Status = $Progress.Status,
        [int]$Id = 0,
        [string]$CurrentOperation = '...',
        [switch]$Completed,
        [switch]$NoDisplay,
        [switch]$NoLog
    )
    $RunTimeFormatted = "[$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') -- Running for $("{0:HH:mm:ss}" -f ([datetime]$($(get-date) - $ScriptStartTime).Ticks))]"
    Write-Verbose -Message "$RunTimeFormatted $($Progress.Status) :: $CurrentOperation"
    $Progress = @{
        Activity = "$Activity"
        Status = "$Status"
        Id = $Id
        CurrentOperation = "$RunTimeFormatted $CurrentOperation"
        Completed = $Completed
    }
    If ($NoDisplay -eq $false) { Write-Progress @Progress }
    #TODO: If ($NoDisplay -eq $false) { Write-CMTraceLog -Message "$($Progress.Status) :: $CurrentOperation"
}

#endregion Additional Functions

####################################################################################################################################################################
####################################################################################################################################################################
#####################################################################Starting#######################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
$Progress = @{'Activity'="$Title started $($ScriptStartTime.ToShortTimeString())"; 'Status'='Initializing'}
Write-ProgressEx

$StartingPath = (get-location).Path
if ($FilePath -notlike "$PSScriptRoot\CMDocumentation.html"){
    if ([System.IO.Path]::IsPathRooted("$FilePath")){
    }else{
        $FilePath = "$($StartingPath)\$($FilePath.TrimStart('.\'))"
    }
    Write-ProgressEx -CurrentOperation "File path is $FilePath"
}
if ($AddDateTime){
    $dirName = [io.path]::GetDirectoryName($FilePath)
    $filename = [io.path]::GetFileNameWithoutExtension($FilePath)
    $ext = [io.path]::GetExtension($FilePath)
    $FilePath = "$dirName\$($filename)_$(get-date -f yyyyMMddTHHmmss)$ext"
}
Write-host "Outputting documentation to: $FilePath"

Write-ProgressEx -CurrentOperation 'Start writing report data'

If ($StyleSheet){
    Write-ProgressEx -CurrentOperation 'Applying Custom style sheet'
    Write-HTMLHeader -Title $Title -File $FilePath -CssStyleFile $StyleSheet
}else{
    Write-HTMLHeader -Title $Title -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Writing cover page'
Write-HTMLCoverPage -Title $Title -Author $Author -Vendor $Vendor -Org $CompanyName -ImagePath $CompanyLogo -File $FilePath
Add-Type -AssemblyName System.Web

$SiteCode = Get-SiteCode

#Import the CM Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
{
    Write-ProgressEx -CurrentOperation 'CM PowerShell module has not been imported yet, will import it now.'
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
$Progress = @{'Activity'="$Title started $($ScriptStartTime.ToShortTimeString())"; 'Status'='Configuration Summary'}

Write-HTMLHeading -Text 'Table of Contents' -Level 1 -PageBreak -ExcludeTOC -File $FilePath
Write-HtmlComment -Text "TOC_Insert_Point" -File $FilePath
Write-HTMLHeading -Text 'Summary of all Sites in this Hierarchy' -Level 1 -PageBreak -File $FilePath
Write-ProgressEx -CurrentOperation 'Getting Sites Information'
$CMSites = Get-CMSite

$CAS                    = $CMSites | Where-Object {$_.Type -eq 4}
$ChildPrimarySites      = $CMSites | Where-Object {$_.Type -eq 3}
$StandAlonePrimarySite  = $CMSites | Where-Object {$_.Type -eq 2}
$SecondarySites         = $CMSites | Where-Object {$_.Type -eq 1}

#region CAS
Write-ProgressEx -CurrentOperation 'Checking CAS'
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
if (-not [string]::IsNullOrEmpty($ChildPrimarySites)){
    Write-ProgressEx -CurrentOperation "Enumerating all child Primary Site."
    Write-HTMLParagraph -Text 'The following child Primary Sites are installed:' -level 1 -File $FilePath
    $ChildSites = @()
    foreach ($ChildSite in $ChildPrimarySites){
        $ChildSites += New-Object -TypeName psobject -Property @{'Site Name' = "$($ChildSite.SiteName)"; 'Site Code' = "$($ChildSite.SiteCode)"; Version = "$($ChildSite.Version)"};
    }
    Write-HtmlTable -InputObject $ChildSites -Border 1 -Level 1 -File $FilePath
}
#endregion Child Primary Sites


#region Standalone Primary
if (-not [string]::IsNullOrEmpty($StandAlonePrimarySite)) {
    Write-ProgressEx -CurrentOperation "Enumerating a standalone Primary Site."
    Write-HTMLParagraph -Text 'The following Primary Site is installed:' -level 1 -File $FilePath
    $CMSiteID = Get-WmiObject -Class SMS_Identification -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider| where-object {$_.ThisSiteCode -eq "$SiteCode"}
    $ReleaseVersion = $CMSiteID.MonthlyReleaseVersion
    $StandAlonePrimarySite = New-Object -TypeName psobject -Property @{'Site Name' = $StandAlonePrimarySite.SiteName; 'Site Code' = $StandAlonePrimarySite.SiteCode; Version = $StandAlonePrimarySite.Version; 'Build' = $StandAlonePrimarySite.BuildNumber; 'Release Version' = $ReleaseVersion};
    $StandAlonePrimarySite = $StandAlonePrimarySite |Select-Object 'Site Name','Site Code','Release Version',Version,Build
    Write-HtmlTable -InputObject $StandAlonePrimarySite -Border 1 -Level 1 -File $FilePath
}
#endregion Standalone Primary

#region Secondary Sites
Write-ProgressEx -CurrentOperation "Enumerating all secondary sites."
if (-not [string]::IsNullOrEmpty($SecondarySites)){
    Write-HTMLParagraph -Text 'The following Secondary Sites are installed:' -level 1 -File $FilePath
    $SecSites = @()
    foreach ($SecondarySite in $SecondarySites){
        $SecSites += New-Object -TypeName psobject -Property @{'Site Name' = "$($SecondarySite.SiteName)"; 'Site Code' = "$($SecondarySite.SiteCode)"; Version = "$($SecondarySite.Version)"}
    }
    Write-HtmlTable -InputObject $SecSites -Border 1 -Level 1 -File $FilePath
}
#endregion Secondary Sites


#region Site Configuration report
foreach ($CMSite in $CMSites)
{  
  Write-ProgressEx -CurrentOperation "Checking each site's configuration"
  Write-HTMLHeading -Text "Configuration Summary for Site $($CMSite.SiteCode)" -level 1 -File $FilePath
  Write-HTMLHeading -Text "Updates and Servicing" -Level 2 -File $FilePath

  #region Site Servicing Updates
  Get-pwCMSiteServicingUpdates -SiteCode $CMSite.SiteCode -FilePath $FilePath
  #endregion Site Servicing Updates

  #region Site Features
  Write-ProgressEx -CurrentOperation "Enumerating Configuration Manager Site Features"
  Write-HTMLHeading -Text "Site Features" -Level 3 -File $FilePath
  $features=Get-CMSiteFeature
  #region release features
  $ReleaseFeatures = $features | Where-Object {$_.FeatureType -eq 1}|Sort-Object Name
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

  #region PreReleaseFeatures
  Write-ProgressEx -CurrentOperation "Enumerating Configuration Manager Site Pre-Release Features"
  $PreReleaseFeatures = $features | Where-Object {$_.FeatureType -eq 0}|Sort-Object Name
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
  #endregion PreReleaseFeatures

  Write-HtmliLink -ReturnTOC -File $FilePath
  Write-ProgressEx -CurrentOperation "Completed Configuration Manager Site Features"
  #endregion Site Features

  #region SiteRoles
Write-ProgressEx -CurrentOperation "Site Roles"
If ($ListAllInformation){
  $SiteRolesTable = @()  
  $SiteRoles = Get-CMSiteRole -SiteCode $CMSite.SiteCode

  Write-HTMLHeading -Text "Site Roles" -Level 2 -File $FilePath
  Write-HTMLParagraph  -Text "The following Site Roles are installed in this site:" -Level 2 -File $FilePath
  foreach ($SiteRole in $SiteRoles) {
    If (($SiteRole.RoleName -eq 'SMS Component Server') `
        -or ($SiteRole.RoleName -eq 'SMS Site Server') `
        -or ($SiteRole.RoleName -eq 'SMS Notification Server') `
        -or ($SiteRole.RoleName -eq 'SMS DM Enrollment Service') `
        -or ($SiteRole.RoleName -eq 'SMS Multicast Service Point')
    ) {
        # Nothing to do
        continue
    }

    $RoleName = ""
    # Get Role settings
    $RoleSettings = @()
    Switch ($SiteRole.RoleName) {
        #region RoleSiteSystem
        'SMS Site System' {
            $RoleName = "Site system"
            $RoleSettings += @("--B--General--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "Server Remote Public Name" }).Value1 -ne "") {
                $RoleSettings += @("- Specify FQDN for use on the Internet - CHECKED")
                $RoleSettings += @("--TAB--Internet FQDN: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "Server Remote Public Name" }).Value1)")
            }
            Else {
                $RoleSettings += @("- Specify FQDN for use on the Internet - UNCHECKED")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "FDMOperation" }).Value -eq 1) {
                $RoleSettings += @("- Require the site server to initiate connections - CHECKED")
            }
            Else {
                $RoleSettings += @("- Require the site server to initiate connections - UNCHECKED")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UseMachineAccount" }).Value -eq 1) {
                $RoleSettings += @("- Installation account: Site server's computer")
            }
            Else {
                $RoleSettings += @("- Installation account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value2)")
            }
            $RoleSettings += @("- Active Directory forest: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ForestFQDN" }).Value1)")
            $RoleSettings += @("- Active Directory domain: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "DomainFQDN" }).Value1)")
            $RoleSettings += @("--B--Proxy--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UseProxy" }).Value -eq 1) {
                $RoleSettings += @("- Proxy: Configured")
                $RoleSettings += @("--TAB--Proxy server name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ProxyName" }).Value2)")
                $RoleSettings += @("--TAB--Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ProxyServerPort" }).Value)")
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "ProxyUserName" }).Value2 -ne "") {
                    $RoleSettings += @("--TAB--Proxy account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ProxyUserName" }).Value2)")
                }
                Else {
                    $RoleSettings += @("--TAB--Proxy account: No account configured")
                }
            }
            Else {
                $RoleSettings += @("- Proxy: Not configured")
            }
        }
        #endregion RoleSiteSystem
        #region RoleDP
        'SMS Distribution Point' {
            $RoleName = "Distribution point"
            $RoleSettings += @("--B--General--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UpdateBranchCacheKey" }).Value -eq 1) {
                $RoleSettings += @("- BranchCache: Enabled")
            }
            Else {
                $RoleSettings += @("- BranchCache: Disabled")
            }
            $RoleSettings += @("- Description: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "Description" }).Value1)")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SslState" }).Value -eq 63) {
                $RoleSettings += @("- Communication: HTTPS")
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "TokenAuthEnabled" }).Value -eq 1) {
                    $RoleSettings += @("--TAB--Allow mobile devices to connect - CHECKED")
                }
                Else {
                    $RoleSettings += @("--TAB--Allow mobile devices to connect - UNCHECKED")
                }
            }
            Else {
                $RoleSettings += @("- Communication: HTTP")
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsAnonymousEnabled" }).Value -eq 1) {
                    $RoleSettings += @("--TAB--Allow clients to connect anonymously - CHECKED")
                }
                Else {
                    $RoleSettings += @("--TAB--Allow clients to connect anonymously - UNCHECKED")
                }
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "CertificateFile" }).Value1 -eq "") {
                $RoleSettings += @("- Certificate: Self-signed")
            }
            Else {
                $RoleSettings += @("- Certificate: Imported")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "PreStagingAllowed" }).Value -eq 1) {
                $RoleSettings += @("- Enable for prestaged content - CHECKED")
            }
            Else {
                $RoleSettings += @("- Enable for prestaged content - UNCHECKED")
            }
            $RoleSettings += @("--B--PXE--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsPXE" }).Value -eq 1) {
                $RoleSettings += @("- PXE support: Enabled")
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsActive" }).Value -eq 1) {
                    $RoleSettings += @("--TAB--Allow to respond to incoming PXE requests - CHECKED")
                }
                Else {
                    $RoleSettings += @("--TAB--Allow to respond to incoming PXE requests - UNCHECKED")
                }
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SupportUnknownMachines" }).Value -eq 1) {
                    $RoleSettings += @("--TAB--Enable unknow computer support - CHECKED")
                }
                Else {
                    $RoleSettings += @("--TAB--Enable unknow computer support - UNCHECKED")
                }
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "PXEPassword" }).Value1 -ne "") {
                    $RoleSettings += @("--TAB--Require a password when computers use PXE - CHECKED")
                }
                Else {
                    $RoleSettings += @("--TAB--Require a password when computers use PXE - UNCHECKED")
                }
                Switch (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UdaSetting" }).Value) {
                    0 {$RoleSettings += @("- User device affinity: Do not use user device affinity")}
                    1 {$RoleSettings += @("- User device affinity: Allow user device affinity with manual approval")}
                    2 {$RoleSettings += @("- User device affinity: Allow user device affinity woth automatic approval")}
                }
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "BindPolicy" }).Value -eq 0) {
                    $RoleSettings += @("- Respond to PXE requests on all network interfaces - CHECKED")
                }
                Else {
                    $RoleSettings += @("- Respond to PXE requests on specific network interfaces - CHECKED")
                }
                $RoleSettings += @("- PXE response delay (seconds): $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ResponseDelay" }).Value)")
            }
            Else {
                $RoleSettings += @("- PXE support: Disabled")
            }
            $RoleSettings += @("--B--Multicast--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsMulticast" }).Value -eq 1) {
                $RoleSettings += @("- Multicast support: Enabled")
                $MCSettings = (Get-CMMulticastServicePoint -SiteSystemServerName ($SiteRole.NALPath).ToString().Split('\\')[2]).Props
                If ($Null -eq $MCSettings) {
                    $RoleSettings += @("--TAB--Multicast configuration unavailable")
                }
                Else {
                    If (($MCSettings | Where-Object { $_.PropertyName -eq "AuthType" }).Value -eq 1) {
                        $RoleSettings += @("- Multicast connection account: $(($MCSettings | Where-Object { $_.PropertyName -eq "UserName" }).Value1)")
                    }
                    Else {
                        $RoleSettings += @("- Multicast connection account: DP's computer account")
                    }
                    If (($MCSettings | Where-Object { $_.PropertyName -eq "IpAddressSource" }).Value -eq 1) {
                        $RoleSettings += @("- Multicast address settings: Use IPv4 addresses within any range - CHECKED")
                    }
                    Else {
                        $RoleSettings += @("- Multicast address settings: Use IPv4 addresses from the following range - CHECKED")
                        $RoleSettings += @("--TAB--Address start range: $(($MCSettings | Where-Object { $_.PropertyName -eq "StartIpAddress" }).Value1)")
                        $RoleSettings += @("--TAB--Address end range: $(($MCSettings | Where-Object { $_.PropertyName -eq "EndIpAddress" }).Value1)")
                    }
                    $RoleSettings += @("- UDP settings:")
                    $RoleSettings += @("--TAB--Port range start: $(($MCSettings | Where-Object { $_.PropertyName -eq "StartPort" }).Value)")
                    $RoleSettings += @("--TAB--Port range end: $(($MCSettings | Where-Object { $_.PropertyName -eq "EndPort" }).Value)")
                    $RoleSettings += @("- Maximum clients: $(($MCSettings | Where-Object { $_.PropertyName -eq "Multicast Max Clients" }).Value)")
                    If (($MCSettings | Where-Object { $_.PropertyName -eq "Multicast Session Schedule Cast" }).Value -eq 1) {
                        $RoleSettings += @("- Scheduled multicast - CHECKED")
                        $RoleSettings += @("--TAB--Session start delay (minutes): $(($MCSettings | Where-Object { $_.PropertyName -eq "Multicast Session Start Delay" }).Value)")
                        $RoleSettings += @("--TAB--Minimum session size (clients): $(($MCSettings | Where-Object { $_.PropertyName -eq "Multicast Session Minimum Size" }).Value)")
                    }
                    Else {
                        $RoleSettings += @("- Scheduled multicast - UNCHECKED")
                    }
                }
            }
            Else {
                $RoleSettings += @("- Multicast support: Disabled")
            }
            $RoleSettings += @("--B--Content Validation--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "DPMonEnabled" }).Value -eq 1) {
                $RoleSettings += @("- Content validation schedule: Enabled")
                $Schedule = Convert-CMSchedule -ScheduleString ($SiteRole.Props | Where-Object { $_.PropertyName -eq "DPMonSchedule" }).Value1
                if ($Schedule.DaySpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.HourSpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.MinuteSpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.ForNumberOfWeeks) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.ForNumberOfMonths) {
                    if ($Schedule.MonthDay -gt 0) {
                        $RoleSettings += @("--TAB--Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                    elseif ($Schedule.MonthDay -eq 0) {
                        $RoleSettings += @("--TAB--Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                    elseif ($Schedule.WeekOrder -gt 0) {
                        switch ($Schedule.WeekOrder) {
                            0 {$order = 'last'}
                            1 {$order = 'first'}
                            2 {$order = 'second'}
                            3 {$order = 'third'}
                            4 {$order = 'fourth'}
                        }
                        $RoleSettings += @("--TAB--Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                }
                Switch (($SiteRole.Props | Where-Object { $_.PropertyName -eq "DPMonPriority" }).Value) {
                    4 { $ValidPriority = "Lowest" }
                    5 { $ValidPriority = "Low" }
                    6 { $ValidPriority = "Medium" }
                    7 { $ValidPriority = "High" }
                    8 { $ValidPriority = "Highest" }
                }
                $RoleSettings += @("- Content validation priority: $ValidPriority")
            }
            Else {
                $RoleSettings += @("- Content validation schedule: Disabled")
            }
            $RoleSettings += @("--B--Boundary Groups--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "DistributeOnDemand" }).Value -eq 1) {
                $RoleSettings += @("- Enable for on-demand distribution - CHECKED")
            }
            Else {
                $RoleSettings += @("- Enable for on-demand distribution - UNCHECKED")
            }
            # Get schedule and rate limit
            $DPInfo = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Query "select * from SMS_SCI_address where ItemName = '$(($SiteRole.NALPath).ToString().Split('\\')[2])|MS_LAN'"  -ComputerName $SMSProvider
            If ($DPInfo) {
                # Schedule
                $RoleSettings += @("--B--Schedule--/B--")
                $RoleSettings += @("- Legend: 1 means all Priorities, 2 means all but low, 3 is high only, 4 means none")
                $iDay = 1
                ForEach ($DPSchedule in $DPInfo.UsageSchedule) {
                    $RoleSettings += @("--TAB--$(Convert-WeekDay $iDay): $($DPSchedule.HourUsage)")
                    $iDay++
                }
                # Rate limits
                $RoleSettings += @("--B--Rate Limits--/B--")
                If ($DPInfo.UnlimitedRateForAll) {
                    $RoleSettings += @("- Rate limit: Unlimited")
                }
                ElseIf (-not [String]::IsNullOrEmpty($DPInfo.RateLimitingSchedule)) {
                    $RoleSettings += @("- Rate limit: Per hour (in %)")
                    $RoleSettings += @("--TAB--Schedule: $($DPInfo.RateLimitingSchedule)")
                }
                Else {
                    $RoleSettings += @("- Rate limit: Pulse mode")
                    $RoleSettings += @("--TAB--Size of data block (KB): $($DPInfo.PropLists.Values[1])")
                    $RoleSettings += @("--TAB--Delay between data blocks (seconds): $($DPInfo.PropLists.Values[2])")
                }
                $RoleSettings += @("--B--Pull Distribution Point--/B--")
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsPullDP" }).Value -eq 1) {
                    $RoleSettings += @("- Enable to pull content from other DP - CHECKED")
                }
                Else {
                    $RoleSettings += @("- Enable to pull content from other DP - UNCHECKED")
                }
            }
        }
        #endregion RoleDP
        #region RoleMP
        'SMS Management Point' {
            $RoleName = "Management point"
            $RoleSettings += @("--B--General--/B--")
            If ($SiteRole.SslState -eq 1) { $RoleSettings += @("- Client connections: HTTPS") }
            Else { $RoleSettings += @("- Client connections: HTTP") }
	    <# moved to Management Point role details / Get-pwCMManagementPoints
            If (Get-CMAlert -Name "Not healthy*Management point*$(($SiteRole.NALPath).ToString().Split('\\')[2])*") {
                $RoleSettings += @("- Generate alert when the MP is not healthy - CHECKED")
            }
            Else {
                $RoleSettings += @("- Generate alert when the MP is not healthy - UNCHECKED")
            }
	    #>
            $RoleSettings += @("--B--Management Point Database--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UseSiteDatabase" }).Value -eq 1) {
                $RoleSettings += @("- Database: Site database")
            }
            Else {
                $RoleSettings += @("- Database: Database replica")
	    <# moved to Management Point role details / Get-pwCMManagementPoints
                $RoleSettings += @("--TAB--Database server: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "SQLServerName" }).Value1)")
                $RoleSettings += @("--TAB--Database name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "DatabaseName" }).Value1)")
            }
            If (-not [String]::IsNullOrEmpty(($SiteRole.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value1)) {
                $RoleSettings += @("- Connection account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value1)")
            }
            Else {
                $RoleSettings += @("- Connection account: MP's computer account")
            #>
	    }
        }
        #endregion RoleMP
        #region RoleFSP
        'SMS Fallback Status Point' {
            $RoleName = "Fallback status point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Number of state messages: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "Throttle Count" }).Value)")
            $RoleSettings += @("- Throttle interval (seconds): $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "Throttle Interval" }).Value)")
        }
        #endregion RoleFSP
        #region RoleSQL
        'SMS SQL Server' {
            $RoleName = "Site database server"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Instance: $(($SiteRole.PropLists.Values -split ", ")[2])")
            $SqlSettings = Get-CMDatabaseProperty -SiteCode $SiteCode #I laugh this cmdlet
            $RoleSettings += @("- Service broker port: $((($SqlSettings -match "Service Broker") -split ",")[1])")
            If ((($SqlSettings -match "IsCompression") -split ",")[1] -eq "1") {
                $RoleSettings += @("- Enable data compression - CHECKED")
            }
            Else {
                $RoleSettings += @("- Enable data compression - UNCHECKED")
            }
            $RoleSettings += @("- Data retention (days): $((($SqlSettings -match "Retention") -split ",")[1])")
        }
        #endregion RoleSQL
        #region RoleRSP
        'SMS SRS Reporting Point' {
            $RoleName = "Reporting services point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Database server name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "DatabaseServerName" }).Value2)")
            $RoleSettings += @("- Database name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "DatabaseName" }).Value2)")
            $RoleSettings += @("- Folder name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "RootFolder" }).Value2)")
            $RoleSettings += @("- Reporting Services server instance: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ReportServerInstance" }).Value2)")
            $RoleSettings += @("- Reporting Services manager URI: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ReportManagerUri" }).Value2)")
            $RoleSettings += @("- Reporting Services server URI: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ReportServerUri" }).Value2)")
            $RoleSettings += @("- Reporting Services account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "Username" }).Value2)")
        }
        #endregion RoleRSP
        #region RoleSUP
        'SMS Software Update Point' {
            $RoleName = "Software update point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- WSUS port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "WSUSIISPort" }).Value)")
            $RoleSettings += @("- WSUS SSL port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "WSUSIISSSLPort" }).Value)")
            $RoleSettings += @("- WSUS DB name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "DBServerName" }).Value2)")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SSLWSUS" }).Value -eq 1) {
                $RoleSettings += @("- WSUS requires SSL - CHECKED")
            }
            Else {
                $RoleSettings += @("- WSUS requires SSL - UNCHECKED")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsIntranet" }).Value -eq 1) {
                If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "IsINF" }).Value -eq 1) {
                    $RoleSettings += @("- WSUS connection type: Allow Internet and Intranet client connections")
                }
                Else {
                    $RoleSettings += @("- WSUS connection type: Allow Intranet-only client connections")
                }
            }
            Else {
                $RoleSettings += @("- WSUS connection type: Allow Internet-only client connections")
            }
            $RoleSettings += @("--B--Proxy And Account Settings--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UseProxy" }).Value -eq 1) {
                $RoleSettings += @("- Use proxy when synchronizing software updates - CHECKED")
            }
            Else {
                $RoleSettings += @("- Use proxy when synchronizing software updates - UNCHECKED")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "UseProxyForADR" }).Value -eq 1) {
                $RoleSettings += @("- Use proxy when downloading content by using ADR - CHECKED")
            }
            Else {
                $RoleSettings += @("- Use proxy when downloading content by using ADR - UNCHECKED")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "WSUSAccessAccount" }).Value2 -ne "") {
                $RoleSettings += @("- WSUS connection account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "WSUSAccessAccount" }).Value2)")
            }
            Else {
                $RoleSettings += @("- WSUS connection account: No account defined")
            }
        }
        #endregion RoleSUP
        #region RoleSC
        'SMS Application Web Service' {
            $RoleName = "Application Catalog web service point"
            $RoleSettings += @("--B--General--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceIISWebSite" }).Value1 -eq "") {
                $RoleSettings += @("- IIS website: Default Web Site")
            }
            Else {
                $RoleSettings += @("- IIS website: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceIISWebSite" }).Value1)")
            }
            $RoleSettings += @("- Web application name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceName" }).Value1)")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SslState" }).Value -eq 0) {
                $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServicePort" }).Value) (HTTP)")
            }
            Else {
                $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServicePort" }).Value) (HTTPS)")
            }
        }
        #endregion RoleSC
        #region RoleSCWeb
        'SMS Portal Web Site' {
            $RoleName = "Application Catalog website point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- IIS web site: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "PortalIISWebSite" }).Value1)")
            $RoleSettings += @("- Web application name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "PortalName" }).Value1)")
            $RoleSettings += @("- NetBIOS name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "NetbiosName" }).Value1)")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SslState" }).Value -eq 0) {
                $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "PortalPort" }).Value) (HTTP)")
            }
            Else {
                $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "PortalSslPort" }).Value) (HTTPS)")
            }
            $RoleSettings += @("--B--Customization--/B--")
            $RoleSettings += @("Organization name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "BrandingString" }).Value1)")
            $RoleSettings += @("Website theme: #$(($SiteRole.Props | Where-Object { $_.PropertyName -eq "PortalThemeColor" }).Value1)")
        }
        #endregion RoleSCWeb
        #region RoleSCP
        'SMS Dmp Connector' {
            $RoleName = "Service connection point"
            $RoleSettings += @("--B--General--/B--")
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "OfflineMode" }).Value -eq 0) {
                $RoleSettings += @("- Mode: Online")
            }
            Else {
                $RoleSettings += @("- Mode: Offline")
            }
        }
        #endregion RoleSCP
        #region RoleAI
        'AI Update Service Point' {
            $RoleName = "Asset Intelligence synchronization point"
            $RoleSettings += @("--B--General--/B--")
            $AISettings = Get-CMAssetIntelligenceProxy
            If ($AISettings.ProxyEnabled) {
                $RoleSettings += @("- Use this AI synchronization point - CHECKED")
                $RoleSettings += @("- Port: $($AISettings.Port)")
            }
            Else {
                $RoleSettings += @("- Use this AI synchronization point - UNCHECKED")
            }
            If (-not [String]::IsNullOrEmpty($AISettings.ProxyCertPath)) {
                $RoleSettings += @("- Certificate path: $($AISettings.ProxyCertPath)")
            }
            Else {
                $RoleSettings += @("- Certificate path: None")
            }
            $RoleSettings += @("--B--Synchronization settings--/B--")
            If ($AISettings.PeriodicCatalogUpdateEnabled) {
                $RoleSettings += @("- Synchronization on a schedule: Enabled")
                $Schedule = Convert-CMSchedule -ScheduleString $AISettings.PeriodicCatalogUpdateSchedule
                if ($Schedule.DaySpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.HourSpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.MinuteSpan -gt 0) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.ForNumberOfWeeks) {
                    $RoleSettings += @("--TAB--Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)")
                }
                elseif ($Schedule.ForNumberOfMonths) {
                    if ($Schedule.MonthDay -gt 0) {
                        $RoleSettings += @("--TAB--Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                    elseif ($Schedule.MonthDay -eq 0) {
                        $RoleSettings += @("--TAB--Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                    elseif ($Schedule.WeekOrder -gt 0) {
                        switch ($Schedule.WeekOrder) {
                            0 {$order = 'last'}
                            1 {$order = 'first'}
                            2 {$order = 'second'}
                            3 {$order = 'third'}
                            4 {$order = 'fourth'}
                        }
                        $RoleSettings += @("--TAB--Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)")
                    }
                }
            }
            Else {
                $RoleSettings += @("- Synchronization on a schedule: Disabled")
            }
        }
        #endregion RoleAI
        #region RoleEP
        'SMS Endpoint Protection Point' {
            $RoleName = "Endpoint Protection point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- License Terms and Privacy Statement acknowledgement - CHECKED")
            $RoleSettings += @("--B--Cloud Protection Service--/B--")
            $RoleSettings += @("- Check manually the membership type selected")
        }
        #endregion RoleEP
        #region RoleEnroll
        'SMS Enrollment Server' {
            $RoleName = "Enrollment point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Website name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceIISWebSite" }).Value1)")
            $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServicePort" }).Value)")
            $RoleSettings += @("- Virtual application name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceName" }).Value1)")
            If (-not [String]::IsNullOrEmpty(($SiteRole.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value1)) {
                $RoleSettings += @("- Connection account: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "UserName" }).Value1)")
            }
            Else {
                $RoleSettings += @("- Connection account: Enrollment point's computer account")
            }
        }
        #endregion RoleEnroll
        #region RoleEnrollWeb
        'SMS Enrollment Web Site' {
            $RoleName = "Enrollment proxy point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Enrollment point: HTTPS://$(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceHostName" }).Value1):$(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServicePort" }).Value)/$(($SiteRole.Props | Where-Object { $_.PropertyName -eq "ServiceName" }).Value1)")
            $RoleSettings += @("- Website name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "EnrollWebIISWebSite" }).Value1)")
            $RoleSettings += @("- Port: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "EnrollWebPort" }).Value)")
            $RoleSettings += @("- Virtual application name: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "EnrollWebName" }).Value1)")
        }
        #endregion RoleEnrollWeb
        #region RoleSMP
        'SMS State Migration Point' {
            $RoleName = "State migration point"
            $RoleSettings += @("--B--General--/B--")
            $RoleSettings += @("- Folder details:")
            ($SiteRole.PropLists | Where-Object { $_.PropertyListName -eq "Directories" }).Values | ForEach-Object {
                $StateDirectory = $_ -split "=|;"
                Switch ($StateDirectory[7]) {
                    1 { $SpaceUnit = "MB" }
                    2 { $SpaceUnit = "GB" }
                    3 { $SpaceUnit = "%" }
                }
                $RoleSettings += @("--TAB--Storage folder: $($StateDirectory[1]) | Max clients: $($StateDirectory[3]) | Min free space: $($StateDirectory[5])$SpaceUnit")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SMPStoreDeletionCycleTimeInMinutes" }).Value -eq 0) {
                $RoleSettings += @("- Deletion policy: Immediatly")
            }
            Else {
                $RoleSettings += @("- Deletion policy: $(($SiteRole.Props | Where-Object { $_.PropertyName -eq "SMPStoreDeletionCycleTimeInMinutes" }).Value) minutes")
            }
            If (($SiteRole.Props | Where-Object { $_.PropertyName -eq "SMPQuiesceState" }).Value -eq 1) {
                $RoleSettings += @("- Enable restore-only mode - CHECKED")
            }
            Else {
                $RoleSettings += @("- Enable restore-only mode - UNCHECKED")
            }
        }
        #endregion RoleSMP
        <#
        TODO
            Certificate registration point
        #>
        Default {
            $RoleName = $SiteRole.RoleName
            $RoleSettings += @("No data available")
        }
    }
    $SiteRoleobject = New-Object -TypeName PSObject -Property @{'Server Name' = ($SiteRole.NALPath).ToString().Split('\\')[2]; 'Role' = $RoleName; 'Properties' = ($RoleSettings -join '--CRLF--')}
    $SiteRolesTable += $SiteRoleobject
  }
  $SiteRolesTable = $SiteRolesTable | Sort-Object -Property 'Server Name', 'Role' | Select-Object 'Server Name', 'Role', 'Properties'
}else{
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
}
  Write-HtmlTable -InputObject $SiteRolesTable -Border 1 -Level 2 -File $FilePath
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion SiteRoles

  #region Site Server Details
If (-Not($PSBoundParameters.ContainsKey('SkipRemoteServerDetails'))) {
  Write-ProgressEx -CurrentOperation "Collecting Site Server Information"
  $SiteServers = Get-CMSiteSystemServer | Where-Object {$_.NALType -notlike 'Windows Azure'} | Select-Object @{Name='ServerName';expression={$_.NetworkOSPath.trim('\')}}
  Write-HTMLHeading -Text "Site Server Information" -Level 2 -File $FilePath
  Write-HTMLParagraph -Text "The section contains basic information on each of the site servers in the environment.  The data is collected via remote WMI queries." -Level 2 -File $FilePath
  foreach ($SiteServer in $SiteServers){
    $ServerInfo = @()
    $Server = $SiteServer.ServerName
    Write-HTMLHeading -Text "$Server" -Level 3 -File $FilePath
    Write-ProgressEx -Id 1 -Activity "Server Details" -Status "Querying WMI" -CurrentOperation "Collecting basic information for server [$Server] via WMI."
    Try{
        $InstalledServerFeatures = ''
        $InstalledFeatures=Get-WmiObject -Query 'SELECT * FROM Win32_OptionalFeature where InstallState = 1' -ComputerName $Server -ErrorAction Stop |Select-Object Caption,Name
        foreach ($Feature in $InstalledFeatures){
            $InstalledServerFeatures = "$InstalledServerFeatures; $($Feature.Name)"
        }
    }
    Catch{
        Write-debug "$(Get-Date):   Unable to query WMI for Installed Features on server [$Server]."
        $InstalledFeatures=@("Feature Query Failed")
    }
    $InstalledServerFeatures = $InstalledServerFeatures.Trim(';')
    Try{
        $Drives = Get-WmiObject -Query 'SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3' -ComputerName $Server -ErrorAction Stop
        $DriveList = @()
        Foreach ($Drive in $Drives){
            $Letter = $Drive.DeviceID
            $Size=[math]::Round($($Drive.Size)/1024/1024/1024,1)
            $Free=[math]::Round($($Drive.FreeSpace)/1024/1024/1024,1)
            $DriveList += New-Object -TypeName psobject -Property @{'DriveLetter'="$Letter";'Size'="$Size GB";'FreeSpace'="$Free GB"}
        }
    }
    Catch{
        Write-debug "$(Get-Date):   Unable to query WMI for drive info on server [$Server]."
        $DriveList = New-Object -TypeName psobject -Property @{'DriveLetter'="NA";'Size'="NA";'FreeSpace'="NA"}
    }
    try{
        [int]$Capacity = 0
        Get-WmiObject -Class win32_physicalmemory -ComputerName $Server -ErrorAction Stop | ForEach-Object {[int64]$Capacity = $Capacity + [int64]$_.Capacity}
        $TotalMemory = $Capacity / 1024 / 1024 / 1024
    }
    catch{
        Write-debug "$(Get-Date):   Failed to collect memory information for server [$Server]."
        [string]$TotalMemory = "Memory query failed"
      }
    Try{
        $CPUs = Get-WmiObject -Class win32_processor -ComputerName $Server -ErrorAction Stop
        [int]$Cores=0
        foreach ($CPU in $CPUs) {
            $Cores = $Cores + $CPU.NumberOfCores
            $CPUModel = $CPU.Name
        }
    }
    Catch{
        Write-debug "$(Get-Date):   Failed to collect processor information for server [$Server]."
        [string]$Cores = "Unknown"
        $CPUModel = "Model Unknown"
    }
    Try{
        $OSInfo = Get-WmiObject -Class win32_OperatingSystem -Property Caption,BuildNumber -ComputerName $Server -ErrorAction Stop
        $OSName = $OSInfo.Caption
        $OSBuild = $Osinfo.BuildNumber
    }
    Catch{
        Write-debug "$(Get-Date):   Failed to collect OS information for server [$Server]."
        $OSName = "Unknown"
        $OSBuild = "Unknown"
    }
    $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="Operating System";'Value'="$OSName"}
    $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="Operating System Build";'Value'="$OSBuild"}
    $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="Installed OS Features";'Value'="$InstalledServerFeatures"}
    $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="Total Memory";'Value'="$TotalMemory"}
    $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="CPU(Cores)";'Value'="$CPUModel ($Cores)"}
    Foreach($Drive in $DriveList){
        $ServerInfo += New-Object -TypeName psobject -Property @{'Property'="Local Disk ($($Drive.DriveLetter))";'Value'="Capacity: $($Drive.Size)--CRLF--Free Space: $($Drive.FreeSpace)"}
    }
    $ServerInfo =  $ServerInfo|Select-Object 'Property','Value'
    Write-HtmlTable -InputObject $ServerInfo -Border 1 -Level 3 -File $FilePath
  }
  Write-ProgressEx -Id 1 -Activity "Server Details" -Status "Querying WMI" -CurrentOperation "End Collecting Site Server Information" -Completed
  Write-ProgressEx -CurrentOperation "End Collecting Site Server Information"
  Write-HtmliLink -ReturnTOC -File $FilePath
}
  #endregion Site Server Details

  #region SiteMaintenanceTasks
  Write-ProgressEx -CurrentOperation "Site Maintenance Tasks"
  $SiteMaintenanceTaskTable = @()
  # Use WMI instead of cmdlet because WMI is more accurate and easy to use
  #$SiteMaintenanceTasks = Get-CMSiteMaintenanceTask -SiteCode $CMSite.SiteCode
if ($ListAllInformation){
  $SiteMaintenanceTasks = Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Query "Select * from SMS_SCI_SQLTask" -ComputerName $SMSProvider
  Write-HTMLHeading -Text "Site Maintenance Tasks for Site $($CMSite.SiteCode)" -Level 2 -File $FilePath
  
  foreach ($SiteMaintenanceTask in $SiteMaintenanceTasks) {
    $DeleteOlderThan = ""
    $BeginTime = ""
    $LatestBeginTime = ""
    $ScheduleTask = ""
    $OtherDetails = ""

    If ($SiteMaintenanceTask.Enabled) {
        # Convert to binary DaysOfWeek integer
        # 1=Sunday, 2=Monday, 4=Tuesday, 8=Wednesday, 16=Thursday, 32=Friday, 64=Saturday
        # Example 1: DaysOfWeek = 64 -> 01000000 -> Display Saturday
        # Example 2: DaysOfWeek = 68 -> 01000100 -> Display Tuesday + Saturday
        $DaysToBinary = "{0:d8}" -f [Int32]([convert]::ToString($SiteMaintenanceTask.DaysOfWeek, 2))
        $DaysToDisplay = @()
        For ($i=$DaysToBinary.Length; $i -gt 0; $i--) {
            If ($DaysToBinary[$i] -eq '1') {
                $DaysToDisplay += @(Convert-WeekDay ($DaysToBinary.Length-$i))
            }
        }
        $ScheduleTask = ($DaysToDisplay -join ", ")
    
        If ($SiteMaintenanceTask.DeleteOlderThan -eq 0) { $DeleteOlderThan = "Not Applicable" }
        Else { $DeleteOlderThan = "$($SiteMaintenanceTask.DeleteOlderThan) days" }

        $BeginTime = ($SiteMaintenanceTask.BeginTime).Substring(8,2)+":"+($SiteMaintenanceTask.BeginTime).substring(10,2)
        $LatestBeginTime = ($SiteMaintenanceTask.LatestBeginTime).Substring(8,2)+":"+($SiteMaintenanceTask.LatestBeginTime).substring(10,2)

        $OtherDetails = "$($SiteMaintenanceTask.DeviceName)"
    }

    $SiteMaintenanceTaskRowHash = New-Object -TypeName PSObject -Property @{
        'Task Name' = $SiteMaintenanceTask.TaskName;
        'Enabled' = $SiteMaintenanceTask.Enabled;
        'Delete older than' = $DeleteOlderThan;
        'Start after' = $BeginTime;
        'Latest start' = $LatestBeginTime;
        'Schedule' = $ScheduleTask;
        'Other details' = $OtherDetails
    }
    $SiteMaintenanceTaskTable += $SiteMaintenanceTaskRowHash;
  }

  $SiteMaintenanceTaskTable = $SiteMaintenanceTaskTable | Sort-Object -Property 'Task Name' | Select-Object 'Task Name', 'Enabled', 'Delete older than', 'Start after', 'Latest start', 'Schedule', 'Other details'
}else{
  $SiteMaintenanceTasks = Get-CMSiteMaintenanceTask -SiteCode $CMSite.SiteCode
  Write-HTMLHeading -Text "Site Maintenance Tasks for Site $($CMSite.SiteCode)" -Level 2 -File $FilePath
  
  foreach ($SiteMaintenanceTask in $SiteMaintenanceTasks) {
    $SiteMaintenanceTaskRowHash = New-Object -TypeName PSObject -Property @{'Task Name' = $SiteMaintenanceTask.TaskName; Enabled = $SiteMaintenanceTask.Enabled};
    $SiteMaintenanceTaskTable += $SiteMaintenanceTaskRowHash;
  }

  $SiteMaintenanceTaskTable = $SiteMaintenanceTaskTable|Select-Object 'Task Name',Enabled
}
  Write-HtmlTable -InputObject $SiteMaintenanceTaskTable -Border 1 -Level 2 -File $FilePath
  Write-HtmliLink -ReturnTOC -File $FilePath
  Write-ProgressEx -CurrentOperation "Completed Site Maintenance Tasks"
  #endregion SiteMaintenanceTasks
  
  #region Site SQL Info
  Write-ProgressEx -CurrentOperation "Getting site SQL server and database information"
  Write-HTMLHeading -Text "Summary of SQL database info for Site $($CMSite.SiteCode)" -PageBreak -Level 2 -File $FilePath
  $SiteDef = Get-CMSiteDefinition -SiteCode $($CMSite.SiteCode)
  $SQLServer = $SiteDef.SQLServerName
  $CMDB = $SiteDef.SQLDatabaseName
  if (($CMDB.split('\')).count -eq 2){
    $CMDatabase = $CMDB.split('\')[1]
    $SQLInstance = $CMDB.split('\')[0]
  }else{
    $CMDatabase = $CMDB
    $SQLInstance = ''
  }
  Write-ProgressEx -CurrentOperation "SQL server: [$SQLServer]  SQL database: [$CMDatabase]  SQL Instance: [$SQLInstance]"
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
    $Drives=Get-WmiObject -Class win32_LogicalDisk -ComputerName $SQLServer | Where-Object {$_.DriveType -eq 3}
    Foreach($Drive in $Drives){
        $SQLHWInfo += "Drive $($Drive.DeviceID) size: $([math]::Round($Drive.size/1024/1024/1024,1)) GB ($([math]::Round($Drive.FreeSpace/1024/1024/1024,1)) GB Free)"
    }
  }
  catch {
    $SQLHWInfo += "Failed to access server: $SQLServer" 
  }
  Write-HtmlList -InputObject $SQLHWInfo -Description $SQLHWDesc -Level 2 -File $FilePath
  
  If ($NoSqlDetail){
    Write-ProgressEx -CurrentOperation "Skipping SQL detailed info."
  }Else{
    Write-ProgressEx -CurrentOperation "Getting SQL Database detailed info."
      Try{
          If ($SQLInstance){
                $SQLConnectString = "$SQLServer\$SQLInstance"
          }else{
                $SQLConnectString = "$SQLServer"
          }
          Write-HTMLParagraph -Text "SQL instance version information:" -Level 2 -File $FilePath
          $SQLVersion = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database Master -QueryTimeout $SQLTimeout -Query "SELECT SERVERPROPERTY (`'edition`') Edition, SERVERPROPERTY(`'productversion`') Version, SERVERPROPERTY (`'productlevel`') SP, SERVERPROPERTY (`'ProductUpdateLevel`') CU" #-Credential $SQLCredential
          $SQLVersion = $SQLVersion | Select-Object Edition,Version,SP,CU
          Write-HtmlTable -InputObject $SQLVersion -Border 1 -Level 3 -File $FilePath
          Write-HTMLParagraph -Text "The following are important global settings on the SQL server.  Typically, this SQL server should be dedicated to Configuration Manager." -Level 2 -File $FilePath
          $SQLConfig = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database Master -QueryTimeout $SQLTimeout -Query "SELECT name ServerSetting,value_in_use Value FROM sys.configurations where configuration_id = 1543 OR configuration_id = 1544 OR configuration_id = 1539" #-Credential $SQLCredential
          $SQLConfig = $SQLConfig | Select-Object @{Name='Server Setting';Expression={$_.ServerSetting}},Value
          Write-HtmlTable -InputObject $SQLConfig -Border 1 -Level 3 -File $FilePath
          Write-HTMLParagraph -Text "Site database information:" -Level 2 -File $FilePath
          $DatabaseInfo = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database Master -QueryTimeout $SQLTimeout -Query "SELECT name, compatibility_level, collation_name FROM sys.Databases WHERE name='$CMDatabase'" #-Credential $SQLCredential
          $DatabaseInfo = $DatabaseInfo | Select-Object @{Name='Database Name';Expression={$_.name}},@{Name='Compatibility Level';Expression={$_.compatibility_level}},@{Name='Collation';Expression={$_.collation_name}}
          Write-HtmlTable -InputObject $DatabaseInfo -Border 1 -Level 3 -File $FilePath
          Write-HTMLParagraph -Text "Below are the database files for the site database ($CMDatabase):" -Level 2 -File $FilePath
          $DatabaseFiles = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database Master -QueryTimeout $SQLTimeout -Query "SELECT db.name `'DatabaseName`',type_desc `'FileType`',physical_name `'FilePath`',mf.state_desc `'Status`',size*8/1024 `'FileSizeMB`',max_size `'MaximumSize`',growth `'GrowthRate`',(CASE WHEN is_percent_growth = 1 THEN `'Percent`' ELSE `'MB`' END) `'GrowthUnit`',create_date `'DateCreated`',compatibility_level `'DBLevel`',user_access_desc `'AccessMode`',recovery_model_desc `'RecoveryModel`' FROM sys.master_files mf INNER JOIN sys.databases db ON db.database_id = mf.database_id where db.name = `'$CMDatabase`'" #-Credential $SQLCredential
          If ([string]::IsNullOrEmpty($DatabaseFiles)) {
              Write-HTMLParagraph -Text "Permission to perform this action is not granted.  Run 'GRANT VIEW ANY DEFINITION TO [%UserOrGroupToGrantRightsTo%]' on the SQL instance.  See https://docs.microsoft.com/en-us/sql/relational-databases/system-catalog-views/sys-master-files-transact-sql?view=sql-server-2016#permissions for detail." -Level 7 -File $FilePath
          } Else {
            $DatabaseFiles = $DatabaseFiles | Select-Object @{Name='File Type';Expression={$_.FileType}},@{Name='File Path';Expression={$_.FilePath}},Status,@{Name='File Size MB';Expression={'{0:N0}' -f $_.FileSizeMB}},@{Name='Maximum Size';Expression={$(IF($_.MaximumSize -eq -1){"Unlimited"}else{'{0:N0}' -f ($_.MaximumSize/128)})}},@{Name='Growth Rate';Expression={"$(IF($_.GrowthUnit -eq "Percent"){"$($_.GrowthRate)%"}Else{"$($_.GrowthRate/128)MB"})"}},@{Name='Recovery Model';Expression={$_.RecoveryModel}}
            Write-HtmlTable -InputObject $DatabaseFiles -Border 1 -Level 3 -File $FilePath
          }
          Write-HTMLParagraph -Text "Below is a fragmentation summary (%) for indexes on the site database ($CMDatabase):" -Level 2 -File $FilePath
          try {
            $IndexFragmentation = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database $CMDatabase -QueryTimeout $SQLTimeout -Query "SELECT SUM(CASE WHEN indexstats.avg_fragmentation_in_percent > 75 THEN  1 ELSE 0 END) [Over 75],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 50 AND indexstats.avg_fragmentation_in_percent <= 75) THEN  1 ELSE 0 END) [50 to 75],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 25 AND indexstats.avg_fragmentation_in_percent <= 50) THEN  1 ELSE 0 END) [25 to 50],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 5 AND indexstats.avg_fragmentation_in_percent <= 25) THEN  1 ELSE 0 END) [5 to 25],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 1 AND indexstats.avg_fragmentation_in_percent <= 5) THEN  1 ELSE 0 END) [Under 5],SUM(CASE WHEN indexstats.avg_fragmentation_in_percent < 1 THEN  1 ELSE 0 END) [Not Fragmented] FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, NULL) AS indexstats JOIN sys.tables dbtables on dbtables.[object_id] = indexstats.[object_id] WHERE indexstats.database_id = DB_ID()" #-Credential $SQLCredential
            $IndexFragmentation = $IndexFragmentation | Select-Object 'Over 75','50 to 75','25 to 50','5 to 25','Under 5','Not Fragmented'
            Write-HtmlTable -InputObject $IndexFragmentation -Border 1 -Level 3 -File $FilePath
          } catch {
            If ($($error[0].Exception.Message) -eq 'Exception calling "Load" with "1" argument(s): "The user does not have permission to perform this action."') {
                Write-HTMLParagraph -Text "Error exception: $($Error[0].exception.Message)" -Level 2 -File $FilePath
                Write-HTMLParagraph -Text "Permission to perform this action is not granted.  Run 'GRANT VIEW SERVER STATE TO [%UserOrGroupToGrantRightsTo%]' on the SQL instance.  See https://docs.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-views/sys-dm-db-index-physical-stats-transact-sql?view=sql-server-2016#permissions for detail." -Level 7 -File $FilePath
            } Else {
                Write-HTMLParagraph -Text "Error exception: $($Error[0].exception.Message)" -Level 6 -File $FilePath
                Write-Debug "$(Get-Date):   Error exception: $($Error[0].exception)"
            }
          }
          Write-HTMLParagraph -Text "Below is a fragmentation summary (%) for indexes on the site database ($CMDatabase) for Tables over 10 MB in size:" -Level 2 -File $FilePath
          $IndexFragmentation10 = Invoke-SqlDataReader -ServerInstance $SQLConnectString -Database $CMDatabase -QueryTimeout $SQLTimeout -Query "SELECT SUM(CASE WHEN indexstats.avg_fragmentation_in_percent > 75 THEN  1 ELSE 0 END) [Over 75],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 50 AND indexstats.avg_fragmentation_in_percent <= 75) THEN  1 ELSE 0 END) [50 to 75],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 25 AND indexstats.avg_fragmentation_in_percent <= 50) THEN  1 ELSE 0 END) [25 to 50],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 5 AND indexstats.avg_fragmentation_in_percent <= 25) THEN  1 ELSE 0 END) [5 to 25],SUM(CASE WHEN (indexstats.avg_fragmentation_in_percent > 1 AND indexstats.avg_fragmentation_in_percent <= 5) THEN  1 ELSE 0 END) [Under 5],SUM(CASE WHEN indexstats.avg_fragmentation_in_percent < 1 THEN  1 ELSE 0 END) [Not Fragmented] FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, NULL) AS indexstats JOIN sys.tables dbtables on dbtables.[object_id] = indexstats.[object_id] WHERE indexstats.database_id = DB_ID() AND page_count > 1280" #-Credential $SQLCredential #1280 pages is 10 MB
          $IndexFragmentation10 = $IndexFragmentation10 | Select-Object 'Over 75','50 to 75','25 to 50','5 to 25','Under 5','Not Fragmented'
          Write-HtmlTable -InputObject $IndexFragmentation10 -Border 1 -Level 3 -File $FilePath
      }
      Catch{
          Write-Warning 'Failed to collect all detailed SQL data.'
          Write-Debug "$(Get-Date):   Error exception: $($Error[0].exception)"
      }
      Write-ProgressEx -CurrentOperation "SQL detailed info complete."
  }
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion Getting Site SQL Info

  #region Management Points
  Get-pwCMManagementPoints -SiteCode $CMSite.SiteCode
  #endregion Management Points

  
  #region Distribution Point details
  Write-ProgressEx -CurrentOperation "Distribution Point(s) Summary"
  Write-HTMLHeading -Text "Summary of Distribution Points for Site $($CMSite.SiteCode)" -Level 2 -PageBreak -File $FilePath
  $CMDistributionPoints = Get-CMDistributionPoint -SiteCode $CMSite.SiteCode
  
If (-not($PSBoundParameters.ContainsKey('SkipRemoteServerDetails'))) {
  foreach ($CMDistributionPoint in $CMDistributionPoints)
  {
    $CMDPServerName = $CMDistributionPoint.NetworkOSPath.Split('\\')[2]
    Write-Verbose "$(Get-Date):   Found DP: $($CMDPServerName)"
    Write-ProgressEx -CurrentOperation "Found DP: $($CMDPServerName)" -Activity "Distribution Points" -Status "Collecting from DB and WMI" -Id 3
    Write-HTMLHeading -Text "$CMDPServerName" -Level 3 -File $FilePath
    Write-Verbose "Trying to ping $($CMDPServerName)"
    $PingResult = Ping-Host $CMDPServerName
    if (-not ($PingResult))
    {
      Write-Debug "Ping Failed: $($CMDPServerName)"
      Write-HTMLParagraph -Text "The Distribution Point $($CMDPServerName) is not reachable. Check connectivity." -Level 3 -File $FilePath
    }
    else
    {
      Write-Debug "Ping Succeeded: $($CMDPServerName)"
      Write-HTMLParagraph -Text "<B>Disk Information:</B>" -Level 4 -File $FilePath
      $CMDPDrives = @(Get-WmiObject -Class SMS_DistributionPointDriveInfo -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider).Where({$PSItem.NALPath -like "*$CMDPServerName*"})
      $DPDrives = @()
      foreach ($CMDPDrive in $CMDPDrives){
            $Size = 0
            $Size = $CMDPDrive.BytesTotal / 1024 / 1024
            $Freesize = 0
            $Freesize = $CMDPDrive.BytesFree / 1024 / 1024
            $DPDrives +=New-Object -TypeName psobject -Property @{'Partition' = "$($CMDPDrive.Drive):"; 'Size' = "$([MATH]::Round($Size,2)) GB"; 'Free Space' = "$([MATH]::Round($Freesize,2)) GB"; 'Percent Free' = "$($CMDPDrive.PercentFree)%"}
      }
      $DPDrives = $DPDrives|Select-Object 'Partition','Size','Free Space','Percent Free'
      Write-HtmlTable -InputObject $DPDrives -Border 1 -Level 4 -File $FilePath
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
    $DPGroupMembers = (Get-WmiObject -class SMS_DPGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider) | Where-Object {$_.DPNALPath -ilike "*$($CMDPServerName)*"}
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
  Write-ProgressEx -CurrentOperation "Completed DPs" -Activity "Distribution Points" -Status "Collecting from DB and WMI" -Id 3 -Completed
  Write-ProgressEx -CurrentOperation "Completed Distribution Point(s) Summary"
  #endregion Distribution Point details

  #region enumerating Software Update Points and Configuration
  Write-ProgressEx -CurrentOperation "Software Update configuration for Site"
  Write-HTMLHeading -Text "Software Update configuration for Site $($CMSite.SiteCode)" -Level 2 -PageBreak -File $FilePath
  
  Write-HTMLHeading -Text "Software Update Point Component Settings for Site $($CMSite.SiteCode)" -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "This is a list of all of the software update classifications and products that are synchronized into the site as well as some of the general site configuration settings." -Level 3 -File $FilePath

  $cats=Get-CMSoftwareUpdateCategory -Fast
  $UpdatesClassifications = $cats|Where-Object {$_.CategoryTypeName -eq "UpdateClassification" -and $_.AllowSubscription -eq $true}
  $SubscribedUpdatesClassifications = $UpdatesClassifications|Where-Object {$_.IsSubscribed -eq $true}
  $Products = $cats|Where-Object {$_.CategoryTypeName -eq "Product" -and $_.AllowSubscription -eq $true}
  $SubscribedProducts = $Products|Where-Object {$_.IsSubscribed -eq $true}
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
  Write-HtmlList -InputObject ($SubscribedProducts.LocalizedCategoryInstanceName | Sort-Object) -Level 4 -File $FilePath


  Write-ProgressEx -CurrentOperation "Enumerating all Software Update Points"
  Write-HTMLHeading -Text "Software Update Point Servers for Site $($CMSite.SiteCode)" -Level 3 -File $FilePath
  Write-HTMLParagraph -Text "Below is the basic configuration and settings for each software update point in the site." -Level 4 -File $FilePath
  Write-debug "Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $SMSProvider | Where-Object {$_.rolename -eq `'SMS Software Update Point`'}"
  $CMSUPs = Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $SMSProvider | Where-Object {$_.rolename -eq 'SMS Software Update Point'}
  #$CMSUPs = (Get-CMSoftwareUpdatePoint).Where({$_.SiteCode -eq "$($CMSite.SiteCode)"})
  if (-not [string]::IsNullOrEmpty($CMSUPs))
  {
    foreach ($CMSUP in $CMSUPs) {
      $SUPPropertyTable = @();
      $CMSUPServerName = $CMSUP.NetworkOSPath.split('\\')[2]
      Write-Debug "$(Get-Date):   Found SUP: $($CMSUPServerName)"
      Write-HTMLHeading -Text "$($CMSUPServerName)" -Level 4 -File $FilePath
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'WSUS IIS Port'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'WSUSIISPORT'}).value)}
      #8530
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Database'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'DBServerName'}).value2)}
      #soup-cm1.soup.steamedsoup.com\MICROSOFT##WID
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Access Account'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'WSUSAccessAccount'}).value2)}
      #soup.steamedsoup.com\SVC-SCCM-RAA
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SSL Enabled'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'SSLWSUS'}).value)}
      #0
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SSL Port'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'WSUSIISSSLPORT'}).value)}
      #8531
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'SUP Enabled'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'Enabled'}).value)}
      #1
      $SUPPropertyTable += New-Object -TypeName psobject -Property @{Name = 'Proxy Enabled'; Value = (($CMSUP.props|Select-Object Propertyname,Value,Value1,Value2| Where-Object {$_.PropertyName -like 'UseProxy'}).value)}
      #0
      $SUPPropertyTable = $SUPPropertyTable|Select-Object Name,Value
      Write-HtmlTable -InputObject $SUPPropertyTable -Border 1 -Level 4 -File $FilePath
      }
  }
  else
  {
    Write-HTMLParagraph -Text "This site has no Software Update Points installed." -Level 3 -File $FilePath
  }
  Write-ProgressEx -CurrentOperation "Completed Enumeration of Software Update Points"
  Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion enumerating Software Update Points and Configuration
  #region enumerating client push settings
    Write-ProgressEx -CurrentOperation "Enumerating Client Push Settings for Site"
    Write-HTMLHeading -Text "Client Push Settings for Site $($CMSite.SiteCode)" -Level 2 -PageBreak -File $FilePath
    Write-HTMLParagraph -Text "Client push allows CM to install the client to computers in the domain directly from CM using admin credentials on the remote computer.  This is <a href=`"https://docs.microsoft.com/en-us/sccm/core/clients/deploy/plan/client-installation-methods`" target=`"_blank`">one of several</a> ways to install the CM client on computers." -Level 3 -File $FilePath
    #Client push settings are found in WMI.  They are all in the list of global properties: SMS_DISCOVERY_DATA_MANAGER
    $CPProps = (Get-WmiObject -Namespace ROOT\SMS\site_$($CMSite.SiteCode) -ComputerName $SMSProvider -Query "SELECT * FROM SMS_SCI_SCProperty where ItemType='SMS_DISCOVERY_DATA_MANAGER'")|Select-Object PropertyName,Value,Value1,Value2
    If(($CPProps | Where-Object {$_.PropertyName -eq 'SETTINGS'}).Value1 -eq 'Active'){
        Write-HTMLParagraph -Level 3 -Text "Automatic site-wide client push is enabled with the below settings." -File $FilePath
    }else{
        Write-HTMLParagraph -Level 3 -Text "Automatic site-wide client push is NOT enabled." -File $FilePath
    }

    $InstallClientsEnabledTitle = "Install Configuration Manager client software on the following computers"
    $InstallClientsEnabled = @()
    $InstallClientsDisabledTitle = "Exclude Configuration Manager client software from the following computers"
    $InstallClientsDisabled = @()

    #FILTERS Property has the following known values.  Evaluating in a Switch: 0=install on all systems; 1=Do not Install Workstations; 2=Do not install on DCs; 4=Do not install servers; 5=Do not install on servers and workstations 7=Do not install servers, workstations or DCs
    Switch(($CPProps | Where-Object {$_.PropertyName -eq 'FILTERS'}).Value){
        0{
            $InstallClientsEnabled += "Servers"
            $InstallClientsEnabled += "Workstations"
            $InstallDC = $true
        }
        1{
            $InstallClientsEnabled += "Servers"
            $InstallClientsDisabled += "Workstations"
            $InstallDC = $true
        }
        2{
            $InstallClientsEnabled += "Servers"
            $InstallClientsEnabled += "Workstations"
        }
        3{
            $InstallClientsEnabled += "Servers"
            $InstallClientsDisabled += "Workstations"
            $InstallDC = $false
        }
        4{
            $InstallClientsDisabled += "Servers"
            $InstallClientsEnabled += "Workstations"
            $InstallDC = $true
        }
        5{
            $InstallClientsDisabled += "Servers"
            $InstallClientsDisabled += "Workstations"
            $InstallDC = $true
        }
        7{
            $InstallClientsDisabled += "Servers"
            $InstallClientsDisabled += "Workstations"
            $InstallDC = $false
        }
    }
    #AutoInstallSiteSystem: 1=Enabled on Site system servers; 0=Disable install on site system servers
    switch(($CPProps|Where-Object {$_.PropertyName -eq 'AutoInstallSiteSystem'}).Value){
        1{$InstallClientsEnabled +="Configuration Manager site system servers"}
        0{$InstallClientsDisabled +="Configuration Manager site system servers"}
    }
    If(-not [string]::IsNullOrEmpty($InstallClientsEnabled)){
        Write-HtmlList -Title $InstallClientsEnabledTitle -InputObject $InstallClientsEnabled -Level 4 -File $FilePath
    }
    If(-not [string]::IsNullOrEmpty($InstallClientsDisabled)){
        Write-HtmlList -Title $InstallClientsDisabledTitle -InputObject $InstallClientsDisabled -Level 4 -File $FilePath
    }
    Switch($InstallDC){
        $true{$DCBehavior="Always install the Configuration Manager client on domain controllers"}
        $false{$DCBehavior="Never install the Configuration Manager client on domain controllers unless specified in the Client Push Installation Wizard"}
    }
    If(-not [string]::IsNullOrEmpty($DCBehavior)){
        Write-HtmlList -Title "Domain Controller Client Install Behavior" -InputObject $DCBehavior -Level 4 -File $FilePath
    }
    $CPSettings = Get-CMClientPushInstallation
    $CPAccounts=($CPSettings.PropLists|Where-Object {$_.PropertyListName -eq 'Reserved2'}).values
    $CPAccountsDescription="Even if client push is not enabled on the site, client push accounts are important/required for installing and reinstalling clients from the console."
    If(-not [string]::IsNullOrEmpty($CPAccounts)){
        if ($MaskAccounts){
            Write-HtmlList -Title "Defined Client Push Accounts" -Description $CPAccountsDescription -InputObject ($CPAccounts|Set-AccountMask) -Level 4 -File $FilePath
        }else{
            Write-HtmlList -Title "Defined Client Push Accounts" -Description $CPAccountsDescription -InputObject $CPAccounts -Level 4 -File $FilePath
        }
    }else{
        Write-HtmlList -Title "Defined Client Push Accounts" -Description $CPAccountsDescription -InputObject "None defined" -Level 4 -File $FilePath
    }
    $InstallProps=($CPSettings.props|Where-Object {$_.PropertyName -eq 'Advanced Client Command Line'}).Value1
    If(-not [string]::IsNullOrEmpty($InstallProps)){
        Write-HtmlList -Title "Client Push MSI Installation Properties" -InputObject $InstallProps -Level 4 -File $FilePath
    }else{
        Write-HtmlList -Title "Client Push MSI Installation Properties" -InputObject "None defined" -Level 4 -File $FilePath
    }
    Write-ProgressEx -CurrentOperation "Completed Enumeration of Client Push Settings for Site"
    Write-HtmliLink -ReturnTOC -File $FilePath
  #endregion enumerating client push settings
  Write-ProgressEx -CurrentOperation "Completed checking site configuration for $($CMSite.SiteCode)"
}
Write-ProgressEx -CurrentOperation "Completed Checking each site's configuration."
#endregion Site Configuration report


##### Hierarchy wide configuration
$Progress = @{'Activity'="$Title started $($ScriptStartTime.ToShortTimeString())"; 'Status'='Hierarchy Wide Configuration'}


Write-HTMLHeading -Level 1 -PageBreak -Text "Summary of Hierarchy Wide Configuration" -File $FilePath

#region Discovery Methods
Write-ProgressEx -CurrentOperation "Enumerating all Discovery Methods"
Write-HTMLHeading -Level 2 -Text "Discovery Methods" -File $FilePath
Write-ProgressEx -CurrentOperation "Enumerating AD Forest Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Active Directory Forest Discovery" -ExcludeTOC -File $FilePath
$DMADFD = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod ActiveDirectoryForestDiscovery
if ($DMADFD.DiscoveryState -eq "Enabled") {
$ForestDiscoverySettings = @()
$ForestDiscoverySettings += "Synchronization Schedule: $($DMADFD.SyncSchedule)"
$ForestDiscoverySettings += "Automatically Create Active Directory Site boundaries: $($DMADFD.ADSiteBoundaryCreation)"
$ForestDiscoverySettings += "Automatically IP address range boundaries for IP subnets: $($DMADFD.SubnetBoundaryCreation)"
Write-HtmlList -Description "Active Directory Forest Discovery is currently enabled with the following settings:" -InputObject $ForestDiscoverySettings -Level 3 -File $FilePath
} else {
    Write-HTMLParagraph -Text "Active Directory Forest Discovery is currently disabled." -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation "Enumerating AD Group Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Active Directory Group Discovery" -ExcludeTOC -File $FilePath
$DMADGD = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod ActiveDirectoryGroupDiscovery
if ($DMADGD.DiscoveryState -eq "Enabled") {
    $GroupDiscoverySettings = @()
    $GroupDiscoverySettings += "Full discovery schedule: $($DMADGD.FullSyncSchedule)"
    If ($DMADGD.IncrementalSyncState -eq "Enabled"){
        $GroupDiscoverySettings += "Delta discovery schedule: $($DMADGD.IncrementalSyncSchedule)"
    } else {
        $GroupDiscoverySettings += "Delta discovery schedule: Disabled"
    }
    If ($DMADGD.FilterExpiredLogonState -eq "Enabled"){
        $GroupDiscoverySettings += "Only discover computers that have logged on to the domain in a given period of time: Enabled"
        $GroupDiscoverySettings += "Time since last logon (days): $($DMADGD.FilterExpiredLogonTime)"
    } else {
        $GroupDiscoverySettings += "Only discover computers that have logged on to the domain in a given period of time: Disabled"
    }
    If ($DMADGD.FilterExpiredPasswordState -eq "Enabled"){
        $GroupDiscoverySettings += "Only discover computers that have updated their computer account password: Enabled"
        $GroupDiscoverySettings += "Time since last password update (days): $($DMADGD.FilterExpiredPasswordTime)"
    } else {
        $GroupDiscoverySettings += "Only discover computers that have updated their computer account password: Disabled"
    }
    If ($DMADGD.DistributionGroupDiscoveryState -eq "Enabled"){
        $GroupDiscoverySettings += "Discover the membership of distribution groups: Enabled"
    } else {
        $GroupDiscoverySettings += "Discover the membership of distribution groups: Disabled"
    }
    $GroupDiscoverySettings += "Search Locations:"
    Write-HtmlList -Description "Active Directory Group Discovery is currently enabled with the following settings:" -InputObject $GroupDiscoverySettings -Level 3 -File $FilePath
    Write-HtmlTable -InputObject $DMADGD.SearchLocations -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text "Active Directory Group Discovery is currently disabled." -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation "Enumerating AD System Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Active Directory System Discovery" -ExcludeTOC -File $FilePath
$DMADSD = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod ActiveDirectorySystemDiscovery
if ($DMADSD.DiscoveryState -eq "Enabled") {
    $SystemDiscoverySettings = @()
    $SystemDiscoverySettings += "Full discovery schedule: $($DMADSD.FullSyncSchedule)"
    If ($DMADSD.IncrementalSync -eq "Enabled"){
        $SystemDiscoverySettings += "Delta discovery schedule: $($DMADSD.IncrementalSyncSchedule)"
    } else {
        $SystemDiscoverySettings += "Delta discovery schedule: Disabled"
    }
    If ($DMADSD.FilterExpiredLogon -eq "Enabled"){
        $SystemDiscoverySettings += "Only discover computers that have logged on to the domain in a given period of time: Enabled"
        $SystemDiscoverySettings += "Time since last logon (days): $($DMADSD.FilterExpiredLogonTime)"
    } else {
        $SystemDiscoverySettings += "Only discover computers that have logged on to the domain in a given period of time: Disabled"
    }
    If ($DMADSD.FilterExpiredPassword -eq "Enabled"){
        $SystemDiscoverySettings += "Only discover computers that have updated their computer account password: Enabled"
        $SystemDiscoverySettings += "Time since last password update (days): $($DMADSD.FilterExpiredPasswordTime)"
    } else {
        $SystemDiscoverySettings += "Only discover computers that have updated their computer account password: Disabled"
    }
    $SystemDiscoverySettings += "Search Containers:"
    Write-HtmlList -Description "Active Directory System Discovery is currently enabled with the following settings:" -InputObject $SystemDiscoverySettings -Level 3 -File $FilePath
    Write-HtmlTable -InputObject $DMADSD.ActiveDirectoryContainers -Level 4 -File $FilePath
    Write-HtmlList -Description "The following additional attributes are synced from AD:" -InputObject $DMADSD.ActiveDirectoryAttributes -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "Active Directory System Discovery is currently disabled." -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation "Enumerating AD User Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Active Directory User Discovery" -ExcludeTOC -File $FilePath
$DMADUD = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod ActiveDirectoryUserDiscovery
if ($DMADUD.DiscoveryState -eq "Enabled") {
    $UserDiscoverySettings = @()
    $UserDiscoverySettings += "Full discovery schedule: $($DMADUD.FullSyncSchedule)"
    If ($DMADUD.IncrementalSync -eq "Enabled"){
        $UserDiscoverySettings += "Delta discovery schedule (minutes): $($DMADUD.IncrementalSyncSchedule)"
    } else {
        $UserDiscoverySettings += "Delta discovery schedule: Disabled"
    }
    $UserDiscoverySettings += "Search Containers:"
    Write-HtmlList -Description "Active Directory System Discovery is currently enabled with the following settings:" -InputObject $UserDiscoverySettings -Level 3 -File $FilePath
    Write-HtmlTable -InputObject $DMADUD.ActiveDirectoryContainers -Level 4 -File $FilePath
    Write-HtmlList -Description "The following additional attributes are synced from AD:" -InputObject $DMADUD.ActiveDirectoryAttributes -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "Active Directory User Discovery is currently disabled." -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation "Enumerating Network Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Network Discovery" -ExcludeTOC -File $FilePath
$DMND = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod NetworkDiscovery
if ($DMND.DiscoveryState -eq "Enabled") {
    $NetworkDiscoverySettings = @()
    $NetworkDiscoverySettings += "Type of discover: $($DMND.DiscoveryType)"
    $NetworkDiscoverySettings += "Slow networks: $($DMND.SlowNetwork)"
if (($DMND.SearchSubnets).count -gt 0){
    $NetworkDiscoverySettings += "Search Subnets:<br />$(Write-HtmlTable -InputObject $DMND.SearchSubnets)"
} else {
    $NetworkDiscoverySettings += "Search Subnets: None defined"
}
    $NetworkDiscoverySettings += "Include local subnets: $($DMND.IncludeLocalSubnets)"
    if (($DMND.SearchDomains).count -gt 0){
        $NetworkDiscoverySettings += "Search Domains:<br />$(Write-HtmlTable -InputObject $DMND.SearchDomains)"
    } else {
        $NetworkDiscoverySettings += "Search Domains: None defined"
    }
    $NetworkDiscoverySettings += "Search local domain: $($DMND.IncludeLocalDomain)"
    If (($DMND.SearchSNMPCommunities).count -gt 0){
        $NetworkDiscoverySettings += Write-HtmlList -Description "SNMP Communities:" -InputObject $DMND.SearchSNMPCommunities
    } else {
        $NetworkDiscoverySettings += "SNMP Communities: None defined"
    }
    $NetworkDiscoverySettings += "Maximum hops: $($DMND.SNMPMaxHopCount)"
    If (($DMND.SearchSNMPDevices).count -gt 0){
        $NetworkDiscoverySettings += Write-HtmlList -Description "SNMP Devices:" -InputObject $DMND.SearchSNMPDevices
    } else {
        $NetworkDiscoverySettings += "SNMP Devices: None defined"
    }
    If (($DMND.SearchDHCPServers).count -gt 0){
        $NetworkDiscoverySettings += Write-HtmlList -Description "DHCP Servers:" -InputObject $DMND.SearchDHCPServers
    } else {
        $NetworkDiscoverySettings += "DHCP Servers: None defined"
    }
    $NetworkDiscoverySettings += "Include Site Server DHCP server: $($DMND.IncludeLocalDHCP)"
    If (($DMND.DiscoverySchedule).count -gt 0) {
        $NetworkDiscoverySettings += Write-HtmlList -Description "Discovery Schedules:" -InputObject $DMND.DiscoverySchedule
    } else {
        $NetworkDiscoverySettings += "Discovery Schedules: None specified"
    }
    Write-HtmlList -Description "Network Discovery is currently enabled with the following settings:" -InputObject $NetworkDiscoverySettings -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "Network Discovery is currently disabled." -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation "Enumerating Heartbeat Discovery Method" -Activity 'Discovery Method' -Status 'Collecting settings' -Id 4
Write-HTMLHeading -Level 3 -Text "Heartbeat Discovery" -ExcludeTOC -File $FilePath
$DMHD = Get-PWCMDiscoveryMethod -SiteServer $SMSProvider -SiteName $SiteCode -DiscoveryMethod HeartbeatDiscovery
if ($DMHD.DiscoveryState -eq "Enabled") {
    $HBDiscoverySettings = @()
    $HBDiscoverySettings += "Heartbeat schedule, send every: $($DMHD.DDRRefreshInterval)"
    Write-HtmlList -Description "Heartbeat Discovery is currently enabled with the following settings:" -InputObject $HBDiscoverySettings -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "Heartbeat Discovery is currently disabled." -Level 3 -File $FilePath
}
Write-ProgressEx -CurrentOperation "Completed Discovery Methods" -Activity 'Discovery Method' -Status 'Completed' -Id 4 -Completed
Write-ProgressEx -CurrentOperation "Completed Enumerating all Discovery Methods"
#endregion Discovery Methods

#region enumerating Boundaries
Write-ProgressEx -CurrentOperation "Enumerating and Sorting Site Boundaries"
Write-HTMLHeading -Level 2 -Text "Summary of Site Boundaries" -File $FilePath

$SubnetBoundaryTable = @();
$ADBoundaryTable = @();
$IPv6BoundaryTable = @();
$IPRangeTable = @();

$Boundaries = Get-CMBoundary
    if (-not [string]::IsNullOrEmpty($Boundaries))
{
  ##Boundary Site Types: 0=IP Subnet; 1=AD Site; 2=IPv6 Prefix; 3=IP Address Range
  foreach ($Boundary in $Boundaries) {       
    if ($Boundary.BoundaryType -eq 0) {
      $BoundaryType = 'IP Subnet';
      $NamesOfBoundarySiteSystems = $Null
      if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
      {
        ForEach-Object -Begin {$BoundarySiteSystems = $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
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
        ForEach-Object -Begin {$BoundarySiteSystems = $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
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
        ForEach-Object -Begin {$BoundarySiteSystems = $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(',')} -End {$NamesOfBoundarySiteSystems} | Out-Null
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
Write-ProgressEx -CurrentOperation "Boundaries Collected and sorted. Writing to report"

#region IPv6 Boundaries Table
      Write-ProgressEx -CurrentOperation "Writing [$($IPv6BoundaryTable.count)] IPv6 boundaries"
      Write-HTMLHeading -Level 3 -Text "IPv6 Boundaries" -File $FilePath
      If ($IPv6BoundaryTable){
          $IPv6BoundaryTable = $IPv6BoundaryTable|Select-Object @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $IPv6BoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IPv6 boundaries defined." -Level 3 -File $FilePath
      }
#endregion IPv6 Boundaries Table
      Write-HTMLParagraph -Text '&nbsp;' -File $FilePath

#region IP Subnet Boundaries Table
      Write-ProgressEx -CurrentOperation "Writing [$($SubnetBoundaryTable.count)] subnet boundaries"
      Write-HTMLHeading -Level 3 -Text "IP Subnet Boundaries" -File $FilePath
      If ($SubnetBoundaryTable){
          $SubnetBoundaryTable = $SubnetBoundaryTable|Select-Object @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $SubnetBoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IP subnet boundaries defined." -Level 3 -File $FilePath
      }
#endregion IP Subnet Boundaries Table
      Write-HTMLParagraph -Text '&nbsp;' -File $FilePath

#region IP Range Boundaries Table
      Write-ProgressEx -CurrentOperation "Writing [$($IPRangeTable.count)] IP range boundaries"
      Write-HTMLHeading -Level 3 -Text "IP Range Boundaries" -File $FilePath
      If ($IPRangeTable){
          $IPRangeTable = $IPRangeTable|Select-Object @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $IPRangeTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No IP Range boundaries defined." -Level 3 -File $FilePath
      }
#endregion IP Range Boundaries Table

#region AD Site Boundaries Table
      Write-ProgressEx -CurrentOperation "Writing [$($ADBoundaryTable.count)] AD Site boundaries"
      Write-HTMLHeading -Level 3 -Text "AD Site Boundaries" -File $FilePath
      If ($ADBoundaryTable){
          $ADBoundaryTable = $ADBoundaryTable|Select-Object @{Name='Name';Expression={$_.Value}},'Description','Boundary Type','Default Site Code','Associated Site Systems'
          Write-HtmlTable -InputObject $ADBoundaryTable -Level 3 -Border 1 -File $FilePath
      } Else {
          Write-HTMLParagraph -Text "No AD Site boundaries defined." -Level 3 -File $FilePath
      }
#endregion AD Site Boundaries Table
    Write-HtmliLink -ReturnTOC -File $FilePath
    Write-ProgressEx -CurrentOperation "Completed Enumerating and Sorting Site Boundaries"
#endregion enumerating Boundaries

#region enumerating all Boundary Groups and their members
Write-ProgressEx -CurrentOperation "Enumerating all Boundary Groups and their members"
Write-HTMLHeading -Level 2 -Text "Site Boundary Groups" -PageBreak -File $FilePath

#region User Defined Boundary Groups
Write-ProgressEx -CurrentOperation "Enumerating user defined Boundary Groups"
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
        $CMSiteSystems = Get-WmiObject -Class SMS_BoundaryGroupSiteSystems -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.GroupID -eq "$($BoundaryGroup.GroupID)"}
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
        Write-debug "Member name: $($MemberName)"
      }
    }
    else
    {
      $MemberNames += 'No associated boundaries'
      Write-debug 'There are no boundaries associated with this Boundary Group.'
    }
    $BoundaryMembers = $MemberNames -join "--CRLF--"
    $BGRs = @()
    $BGRs += Get-CMBoundaryGroupRelationship -SourceGroupId $BoundaryGroup.GroupID
    $FallBackRelationships = ""
    if ($BGRs.count -gt 0){
        Foreach ($BGR in $BGRs){
            $BGName = "$($BGR.DestinationGroupName)"
            if ($BGR.FallbackDP -eq -1){
                $FallbackDP = "Never"
            }else{
                $FallbackDP = "$($BGR.FallbackDP) minutes"
            }
            if ($BGR.FallbackMP -eq -1){
                $FallbackMP = "Never"
            }else{
                $FallbackMP = "$($BGR.FallbackMP) minutes"
            }
            if ($BGR.FallbackSUP -eq -1){
                $FallbackSUP = "Never"
            }else{
                $FallbackSUP = "$($BGR.FallbackSUP) minutes"
            }
            $FallBackRelationships = "$($FallBackRelationships)--B----U--$BGName--/U----/B-- --CRLF--Fallback DP: $FallbackDP --CRLF--Fallback MP: $FallbackMP --CRLF--Fallback SUP: $FallbackSUP --CRLF--"
        }
    }
    Switch($BoundaryGroup.Flags){
        0{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----UNCBOX-- Prefer cloud DP over DP."}
        1{$BGOptions = "--UNCBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----UNCBOX-- Prefer cloud DP over DP."}
        2{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----CBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----UNCBOX-- Prefer cloud DP over DP."}
        4{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----CBOX-- Prefer DP over peers within same Subnet.--CRLF----UNCBOX-- Prefer cloud DP over DP."}
        6{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----CBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----CBOX-- Prefer DP over peers within same Subnet.--CRLF----UNCBOX-- Prefer cloud DP over DP."}
        8{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----CBOX-- Prefer cloud DP over DP."}
        9{$BGOptions = "--UNCBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----CBOX-- Prefer cloud DP over DP."}
        10{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----CBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----UNCBOX-- Prefer DP over peers within same Subnet.--CRLF----CBOX-- Prefer cloud DP over DP."}
        12{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----UNCBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----CBOX-- Prefer DP over peers within same Subnet.--CRLF----CBOX-- Prefer cloud DP over DP."}
        14{$BGOptions = "--CBOX-- Allow peer downloads in this Boundary Group.--CRLF----CBOX-- During peer downloads, only use peers within the same Subnet.--CRLF----CBOX-- Prefer DP over peers within same Subnet.--CRLF----CBOX-- Prefer cloud DP over DP."}
        default{$BGOptions = "Unknown Options"}
    }
    $BoundaryGroupRow = New-Object -TypeName psobject -Property @{Name = $BoundaryGroup.Name; Description = $BoundaryGroup.Description; 'Assigned Site' = $BoundaryGroup.DefaultSiteCode; 'Boundary Members' = "$BoundaryMembers"; 'Site Systems' = $BoundaryGroupSiteSystems;'Fallback Relationships' = "$FallBackRelationships"; 'Options' = "$BGOptions"};
    $BoundaryGroupTable += $BoundaryGroupRow
  }
  
  $BoundaryGroupTable = $BoundaryGroupTable|Select-Object 'Name','Description','Assigned Site','Boundary Members','Site Systems','Fallback Relationships','Options'
  Write-HtmlTable -InputObject $BoundaryGroupTable -Level 4 -Border 1 -File $FilePath
}
else
{
  Write-HTMLParagraph -Level 3 -Text "There are no Boundary Groups configured. It is mandatory to configure a Boundary Group for Configuration Manger to work properly." -File $FilePath
}
Write-ProgressEx -CurrentOperation "Completed user defined Boundary Groups"
#endregion User Defined Boundary Groups

#region Default Boundary Group
Write-ProgressEx -CurrentOperation "Enumerating Default Boundary Group"
Write-HTMLHeading -Level 3 -Text "Default Boundary Group" -File $FilePath

$DefaultBG = Get-CMDefaultBoundaryGroup
$DefaultBGID = $DefaultBG.GroupID
$BGSystems = @()
if ($DefaultBG.SiteSystemCount -gt 0)
{
    $CMSiteSystems = Get-WmiObject -Class SMS_BoundaryGroupSiteSystems -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.GroupID -eq "$DefaultBGID"}
    foreach($SS in $CMSiteSystems){
        $BGSystems +=[regex]::Match($SS.ServerNALPath,'\[\"Display=\\\\(.*)\\\"\]MSWNET').Groups[1].value
    }
    $BoundaryGroupSiteSystems = $BGSystems -join '--CRLF--'
}
Else
{
    $BoundaryGroupSiteSystems = "None"
}
$DefaultBGRelationship = Get-WmiObject -Class SMS_BoundaryGroupRelationships -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider | Where-Object {($_.SourceGroupID -eq "$DefaultBGID") -and ($_.DestinationGroupID -eq "$DefaultBGID")}
$FallbackSUP = $DefaultBGRelationship.FallbackSUP
$FallbackDP = $DefaultBGRelationship.FallbackDP
$FallbackMP = $DefaultBGRelationship.FallbackMP
IF (($FallbackSUP -eq -1) -or ($FallbackSUP -eq 0)) {$FallbackSUP = "Never"}else{$FallbackSUP = "$FallbackSUP mins"}
IF (($FallbackDP -eq -1) -or ($FallbackDP -eq 0)) {$FallbackDP = "Never"}else{$FallbackDP = "$FallbackDP mins"}
IF (($FallbackMP -eq -1) -or ($FallbackMP -eq 0)) {$FallbackMP = "Never"}else{$FallbackMP = "$FallbackMP mins"}

$DefaultBoundaryGroupRow = New-Object -TypeName psobject -Property @{Name = $DefaultBG.Name; 'Site Systems' = $BoundaryGroupSiteSystems; 'DP Fallback Time' = $FallbackDP; 'SUP Fallback Time' = $FallbackSUP; 'MP Fallback Time' = $FallbackMP};
$DefaultBoundaryGroupRow = $DefaultBoundaryGroupRow|Select-Object 'Name','Site Systems','DP Fallback Time','SUP Fallback Time','MP Fallback Time'
Write-HtmlTable -InputObject $DefaultBoundaryGroupRow -Level 4 -Border 1 -File $FilePath

Write-ProgressEx -CurrentOperation "Completed Enumerating Default Boundary Group"
#endregion Default Boundary Group
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all Boundary Groups and their members

#region Cloud Services
Get-pwCMCloudManagementGateway -SiteCode $CMSite.SiteCode -FilePath $FilePath
#endregion Cloud Services

#region enumerating Client Policies
Write-ProgressEx -CurrentOperation "Enumerating all Client Settings"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Summary of Custom Client Device Settings' -File $FilePath

$AllClientSettings = Get-CMClientSetting | Where-Object -FilterScript {$_.SettingsID -ne '0'}

If ($ListAllInformation){
    foreach ($ClientSetting in $AllClientSettings)
    {
      Write-ProgressEx -CurrentOperation "Client Setting: $($ClientSetting.Name)" -Activity 'Client Settings' -Status 'Collecting detailed settings' -Id 5
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
        If("$($ClientSetting.AssignmentCount)" -gt 0){
            $CSDeployments=Get-WmiObject -Query "SELECT * FROM SMS_ClientSettingsAssignment WHERE ClientSettingsID=$($ClientSetting.SettingsID)" -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider
            $CSDeploymentArray = @()
            foreach ($CSD in $CSDeployments){
                $CreationTime = [datetime]::ParseExact("$($CSD.CreationTime.Split('.')[0])",'yyyyMMddHHmmss',$null)
                $CSDeploymentArray += New-Object -TypeName psobject -Property @{'Collection ID'="$($CSD.CollectionID)";'Collection Name'="$($CSD.CollectionName)";'Date Created'="$CreationTime"}
            }
            Write-HTMLParagraph -Text "This client setting policy is deployed to the following $($CSDeploymentArray.count) collections:" -Level 3 -File $FilePath
            $CSDeploymentArray = $CSDeploymentArray | Select-Object 'Collection Name','Collection ID','Date Created'
            Write-HtmlTable -InputObject $CSDeploymentArray -Border 1 -Level 4 -File $FilePath
        }else{
            Write-HTMLParagraph -Text "This client setting policy is not deployed to any collections." -Level 3 -File $FilePath
        }
      Write-HTMLParagraph -Level 3 -Text "<u><b>Setting Configuration</b></u>:" -File $FilePath
      function Get-ClientSettingsConfigs {
          #This function captures all the client settings for a client setting and combines them into a readable format in the documentation.
          param (
              $AgentConfigs,
              [switch]$Default
          )
      foreach ($AgentConfig in $AgentConfigs)
      {
        try
        {
          switch ($AgentConfig.AgentID)
          {
            1{
              $Config = 'Compliance Settings'
              $KnownProps = @("AgentID","Enabled","EnableUserStateManagement","EvaluationSchedule","PerProviderTimeout","PerScanDefaultPriority","PerScanTimeout","PerScanTTL","PSComputerName","PSShowComputerName","SmsProviderObjectPath")
              $ConfigList = @()
              $ConfigList += "Enable compliance evaluation on clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              IF ($Default){
                $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
                $ConfigList += "Schedule Compliance Evaluation: " + (Get-HumanReadableSchedule -Schedule $Schedule)
              }
              $ConfigList += "Enable user data and profiles: " + $(Convert-Bool2Text $AgentConfig.EnableUserStateManagement)
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
              $ConfigList += "Enable software inventory on clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
              $ConfigList += "Schedule: " + (Get-HumanReadableSchedule -Schedule $Schedule)
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
                  Write-HtmlTable -InputObject $InvFiles -Level 6 -Border 1 -File $FilePath
              } else {
                  $ConfigList += 'Inventory these file types: None'
                  Write-HtmlList -InputObject $ConfigList -Description "<b>$Config</b>" -Level 3 -File $FilePath
              }
              if ($CollectedFiles.count -gt 0) {
                  Write-HtmlList -InputObject 'Collect Files:' -Level 3 -File $FilePath
                  Write-HtmlTable -InputObject $CollectedFiles -Level 6 -Border 1 -File $FilePath
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
              $ConfigList += "Users can change policy or notification settings in Software Center: " + $(Convert-Bool2Text $AgentConfig.AllowClientChange)
              $ConfigList += "Allow Remote Control of an unattended computer: " + $(Convert-Bool2Text $AgentConfig.AllowRemCtrlToUnattended)
              $ConfigList += "Prompt user for Remote Control permission: " + $(Convert-Bool2Text $AgentConfig.PermissionRequired)
              $ConfigList += "Prompt user for permission to transfer content from shared clipboard: " + $(Convert-Bool2Text $AgentConfig.ClipboardAccessPermissionRequired)
              $ConfigList += "Grant Remote Control permission to local Administrators group: " + $(Convert-Bool2Text $AgentConfig.AllowLocalAdminToDoRemoteControl)
              switch ($AgentConfig.AccessLevel)
              {
                0 { $accesslevel = 'No access' }
                1 { $accesslevel = 'View only' }
                2 { $accesslevel = 'Full Control' }
              }
              $ConfigList += "Access level allowed: $accesslevel"
              if ($AgentConfig.PermittedViewers.count -gt 0) {
                If($MaskAccounts){
                    $viewers = Write-HtmlList -InputObject (($AgentConfig.PermittedViewers)|Set-AccountMask)
                }else{
                    $viewers = Write-HtmlList -InputObject ($AgentConfig.PermittedViewers)
                }
                $ConfigList += "Permitted viewers of Remote Control and Remote Assistance: $viewers"
              } Else {
                $ConfigList += "Permitted viewers of Remote Control and Remote Assistance: None"
              }
              $ConfigList += "Show session notification icon on taskbar: " + $(Convert-Bool2Text $AgentConfig.RemCtrlTaskbarIcon)
              $ConfigList += "Show session connection bar: " + $(Convert-Bool2Text $AgentConfig.RemCtrlConnectionBar)
              Switch ($AgentConfig.AudibleSignal)
              {
                0 { $ClientSound = 'None.' }
                1 { $ClientSound = 'Beginning and end of session.' }
                2 { $ClientSound = 'Repeatedly during session.' }
              }
              $ConfigList += "Play a sound on client: $ClientSound"
              $ConfigList += "Manage unsolicited Remote Assistance settings: " + $(Convert-Bool2Text $AgentConfig.ManageRA)
              $ConfigList += "Manage solicited Remote Assistance settings: " + $(Convert-Bool2Text $AgentConfig.EnforceRAandTSSettings)
              #Level of access for Remote Assistance:
              if (($AgentConfig.AllowRAUnsolicitedView -notin @('True',1)) -and ($AgentConfig.AllowRAUnsolicitedControl -notin @('True',1)))
              {
                $RALevel = 'None'
              }
              elseif (($AgentConfig.AllowRAUnsolicitedView -in @('True',1)) -and ($AgentConfig.AllowRAUnsolicitedControl -notin @('True',1)))
              {
                $RALevel = 'Remote viewing'
              }
              elseif (($AgentConfig.AllowRAUnsolicitedView -in @('True',1)) -and ($AgentConfig.AllowRAUnsolicitedControl -in @('True',1)))
              {
                $RALevel = 'Full Control'
              }
              $ConfigList += "Level of access for Remote Assistance: $RALevel"
              $ConfigList += "Manage Remote Desktop settings: " + $(Convert-Bool2Text $AgentConfig.ManageTS)
              if ($AgentConfig.ManageTS -in @('True',1))
              {
                $ConfigList += "Allow permitted viewers to connect by using Remote Desktop connection: " + $(Convert-Bool2Text $AgentConfig.EnableTS)
                $ConfigList += "Require network level authentication on computers that run Windows Vista operating system and later versions: " + $(Convert-Bool2Text $AgentConfig.TSUserAuthentication)
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
              $KnownProps = @("AddPortalToTrustedSiteList","AgentID","AllowPortalToHaveElevatedTrust","BrandingTitle","DayReminderInterval","DisplayNewProgramNotification","EnableHealthAttestation","EnableThirdPartyOrchestration","GracePeriodHours","HourReminderInterval","InstallRestriction","OnPremHAServiceUrl","OSDBrandingSubTitle","PortalUrl","PowerShellExecutionPolicy","PSComputerName","PSShowComputerName","ReminderInterval","SmsProviderObjectPath","SUMBrandingSubTitle","SuspendBitLocker","SWDBrandingSubTitle","SystemRestartTurnaroundTime","UseNewSoftwareCenter","UseOnPremHAService","DevicesHealthCollectionEnabled")
              $Config = 'Computer Agent'
              $ConfigList = @()
              If ($Default){
                $ConfigList += "Deployment deadline greater than 24 hours, remind user every (hours): $([string]($AgentConfig.ReminderInterval))"
                $ConfigList += "Deployment deadline less than 24 hours, remind user every (hours): $([string]($AgentConfig.DayReminderInterval))"
                $ConfigList += "Deployment deadline less than 1 hour, remind user every (minutes): $([string]($AgentConfig.HourReminderInterval))"  
              } Else {
                $ConfigList += "Deployment deadline greater than 24 hours, remind user every (hours): $([string]($AgentConfig.ReminderInterval) / 60 / 60)"
                $ConfigList += "Deployment deadline less than 24 hours, remind user every (hours): $([string]($AgentConfig.DayReminderInterval) / 60 / 60)"
                $ConfigList += "Deployment deadline less than 1 hour, remind user every (minutes): $([string]($AgentConfig.HourReminderInterval) / 60)"  
              }
              $ConfigList += "Default application catalog website point: $($AgentConfig.PortalUrl)"
              $ConfigList += "Add default Application Catalog website to Internet Explorer trusted sites zone: " + $(Convert-Bool2Text $AgentConfig.AddPortalToTrustedSiteList)
              $ConfigList += "Allow Silverlight applications to run in elevated trust mode: " + $(Convert-Bool2Text $AgentConfig.AllowPortalToHaveElevatedTrust)
              $ConfigList += "Organization name displayed in Software Center: $($AgentConfig.BrandingTitle)"
              $ConfigList += "Use New Software Center: " + $(Convert-Bool2Text $AgentConfig.UseNewSoftwareCenter)
              $ConfigList += "Enable communication with Health Attestation Service: " + $(Convert-Bool2Text $AgentConfig.EnableHealthAttestation)
              $ConfigList += "Use on-premises Health Attestation Service: " + $(Convert-Bool2Text $AgentConfig.UseOnPremHAService)
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
              $ConfigList += "Suspend BitLocker PIN entry on restart: $($SuspendBitlocker)"
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
              $ConfigList += "PowerShell execution policy: $($ExecutionPolicy)"
              switch ($AgentConfig.DisplayNewProgramNotification)
              {
                False { $DisplayNotifications = 'No' }
                0 { $DisplayNotifications = 'No' }
                True { $DisplayNotifications = 'Yes' }
                1 { $DisplayNotifications = 'Yes' }
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
              switch ($AgentConfig.DevicesHealthCollectionEnabled) {
                    '1' {$EndPointAnalytics = "Yes"}
                    '0' {$EndPointAnalytics = "No"}
                    Default {$EndPointAnalytics = "Yes"}
              }
              $ConfigList += "Enable Endpoint Analytics data collection: " + $EndPointAnalytics
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
            6{
                $KnownProps = @("CacheTombstoneContentMinDuration")
                #CacheTombstoneContentMinDuration : 93300 Seconds -- Minimum duration before the client can remove cached content (minutes):
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
            8{
              $KnownProps = @("AgentID","DataCollectionSchedule","Enabled","LastUpdateTimeOfRules","MaximumUsageInstancesPerReport","MeterRuleIDList","MRUAgeLimitInDays","MRURefreshInMinutes","PSComputerName","PSShowComputerName","ReportTimeout","SmsProviderObjectPath")
              $Config = 'Software Metering'
              $ConfigList = @()
              $ConfigList += "Enable software metering on clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.DataCollectionSchedule
              $DCSched = Get-HumanReadableSchedule -Schedule $Schedule
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
              $KnownProps = @("AgentID","AlternateContentProviders","AssignmentBatchingTimeout","BrandingSubTitle","BrandingTitle","ContentDownloadTimeout","ContentLocationTimeout","DayReminderInterval","Enabled","EnableExpressUpdates","EvaluationSchedule","ExpressUpdatesPort","HourReminderInterval","MaxRandomDelayMinutes","MaxScanRetryCount","O365Management","PerDPInactivityTimeout","PSComputerName","PSShowComputerName","ReminderInterval","ScanRetryDelay","ScanSchedule","SmsProviderObjectPath","TotalInactivityTimeout","UpdateStatusRefreshIntervalDays","UserExperience","UserJobPerDPInactivityTimeout","UserJobTotalInactivityTimeout","WSUSLocationTimeout","WSUSScanRetryCodes","WUAMaxRebootsWhenOnInternet","WUASuccessCodes","WUfBEnabled","EnableThirdPartyUpdates","ServiceWindowManagement","NEOPriorityOption","DynamicUpdateOption")
              $Config = 'Software Updates'
              $ConfigList = @()
              $ConfigList += "Enable software updates on clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              ##Software Update scan schedule:
              $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.ScanSchedule
              $SoftScanSched = Get-HumanReadableSchedule -Schedule $Schedule
              $ConfigList += "Software Update scan schedule: $SoftScanSched"
              ##Schedule deployment re-evaluation:
              $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
              $SoftReevalSched = Get-HumanReadableSchedule -Schedule $Schedule
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
                  $ConfigList += "Allow clients to download delta content when available (`"Enable installation of Express installation files on clients`" - Before 1902): No"
              }
              else
              {
                  $ConfigList += "Allow clients to download delta content when available (Enable installation of Express installation files on clients - Before 1902): Yes"
                  $ConfigList += "Port used to download content for Express installation files: $($AgentConfig.ExpressUpdatesPort)"
              }
              switch ($AgentConfig.O365Management) {
                  0{$ConfigList += "Enable management of the Office 365 Client Agent: Not Configured"}
                  1{$ConfigList += "Enable management of the Office 365 Client Agent: Yes"}
                  2{$ConfigList += "Enable management of the Office 365 Client Agent: No"}
              }
              <# No longer appears this is listed in the client settings....
              If($AgentConfig.EnableThirdPartyUpdates -eq "True")
              {
                  $ConfigList += "Enable Third Party Software Updates: Yes"
              }
              else
              {
                  $ConfigList += "Enable Third Party Software Updates: No"
              }  #>
              If($AgentConfig.ServiceWindowManagement -eq "True"){
                  $ConfigList += "Enable installation of software updates in `"All deployments`" maintenance windows when `"Software update`" maintenance window is available: Yes"
              } else {
                  $ConfigList += "Enable installation of software updates in `"All deployments`" maintenance windows when `"Software update`" maintenance window is available: No"
              }
              switch ($AgentConfig.NEOPriorityOption) {
                  0 { $ThreadPriority = 'Not Configured' }
                  1 { $ThreadPriority = 'Normal' }
                  2 { $ThreadPriority = 'Low' }
              }
              switch ($AgentConfig.DynamicUpdateOption) {
                  0 { $DynamicUpdateOption = 'Not Configured' }
                  1 { $DynamicUpdateOption = 'Yes' }
                  2 { $DynamicUpdateOption = 'No' }
              }
              $ConfigList += "Specify thread priority for feature updates: $ThreadPriority"
              $ConfigList += "Enable Dynamic Update for feature updates: $DynamicUpdateOption"
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
              if ($Default -or $ClientSetting.Type -eq '1'){
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
              IF ($Default){
                IF ($($AgentConfig.AllowUserAffinity) -eq '1'){
                    $UserDefinedAffinity = 'Yes'
                  }Else{
                    $UserDefinedAffinity = 'No'
                  }
                  $ConfigList += "User Setting:  Allow user to define their primary devices: $UserDefinedAffinity"
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
              $ConfigList += "Limit the maximum network bandwidth for BITS background transfers: " + $(Convert-Bool2Text $AgentConfig.EnableBitsMaxBandwidth)
              $ConfigList += "Throttling window start time (24hr): $($AgentConfig.MaxBandwidthValidFrom):00"
              $ConfigList += "Throttling window end time (24hr): $($AgentConfig.MaxBandwidthValidTo):00"
              $ConfigList += "Maximum transfer rate during throttling window (kbps): $($AgentConfig.MaxTransferRateOnSchedule)"
              $ConfigList += "Allow BITS downloads outside the throttling window: " + $(Convert-Bool2Text $AgentConfig.EnableDownloadOffSchedule)
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
              if ($Default -or $ClientSetting.Type -eq '1'){
                  $ConfigList += "Polling interval for modern devices (minutes): $($AgentConfig.MDMPollInterval)"
              } Else {
                  If ($AgentConfig.EnableDeviceEnrollment -eq '1'){
                    $ConfigList += 'Allow users to enroll mobile devices and Mac computers: Yes'
                    $MacDEID = "$($AgentConfig.DeviceEnrollmentProfileID)"
                    $MacDEName = (Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "Select * from SMS_DeviceEnrollmentProfile where ProfileID = `'$($AgentConfig.DeviceEnrollmentProfileID)`'" -ComputerName $SMSProvider).Name
                    $ConfigList += "Enrollment Profile: $MacDEName (ID: $MacDEID)"
                  }else{
                    $ConfigList += 'Allow users to enroll mobile devices and Mac computers: No'
                  }
                  If ($AgentConfig.EnableModernDeviceEnrollment -eq '1'){
                    $ConfigList += "Allow users to enroll modern devices: Yes"
                    $ModernDEID = "$($AgentConfig.ModernDeviceEnrollmentProfileID)"
                    $ConfigList += "Modern device enrollment profile: $ModernDEID"
                  }else{
                    $ConfigList += "Allow users to enroll modern devices: No"
                  }
              }
              IF ($Default){
                If ($AgentConfig.EnableDeviceEnrollment -eq '1'){
                    $ConfigList += 'User Setting: Allow users to enroll mobile devices and Mac computers: Yes'
                    $MacDEID = "$($AgentConfig.DeviceEnrollmentProfileID)"
                    $MacDEName = (Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "Select * from SMS_DeviceEnrollmentProfile where ProfileID = `'$($AgentConfig.DeviceEnrollmentProfileID)`'" -ComputerName $SMSProvider).Name
                    $ConfigList += "User Setting: Enrollment Profile: $MacDEName (ID: $MacDEID)"
                  }else{
                    $ConfigList += 'User Setting: Allow users to enroll mobile devices and Mac computers: No'
                  }
                  If ($AgentConfig.EnableModernDeviceEnrollment -eq '1'){
                    $ConfigList += "User Setting: Allow users to enroll modern devices: Yes"
                    $ModernDEID = "$($AgentConfig.ModernDeviceEnrollmentProfileID)"
                    $ConfigList += "User Setting: Modern device enrollment profile: $ModernDEID"
                  }else{
                    $ConfigList += "User Setting: Allow users to enroll modern devices: No"
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
              $ConfigList += "Enable user policy on clients: " + $(Convert-Bool2Text $AgentConfig.PolicyEnableUserPolicyPolling)
              $ConfigList += "Enable user policy requests from Internet clients: " + $(Convert-Bool2Text $AgentConfig.PolicyEnableUserPolicyOnInternet)
              $ConfigList += "Enable user policy for multiple user sessions: " + $(Convert-Bool2Text $AgentConfig.PolicyEnableUserPolicyOnTS)
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
              $ConfigList += "Enable hardware inventory on clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
              $HardwareInvSched = Get-HumanReadableSchedule -Schedule $Schedule
              $ConfigList += "Hardware inventory schedule: $HardwareInvSched"
              $ConfigList += "Maximum random delay (minutes): $($AgentConfig.MaxRandomDelayMinutes)"
              IF ($Default){
                $ConfigList += "Maximum custom MIF file size (KB): $($AgentConfig.Max3rdPartyMIFSize)"
                switch ($AgentConfig.MIFCollection) {
                    0 { $MIFfiles = 'None' }
                    4 { $MIFfiles = 'Collect NOIDMIF files' }
                    8 { $MIFfiles = 'Collect IDMIF files' }
                    12 { $MIFfiles = 'Collect IDMIF and NOIDMIF files' }
                }
                $ConfigList += "Collect MIF files: $MIFfiles"
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
              $AgentEvalSched = Get-HumanReadableSchedule -Schedule $Schedule
              $ConfigList += "Schedule re-evaluation for deployments: $AgentEvalSched"
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
              $KnownProps = @("AgentID","AllowUserToOptOutFromPowerPlan","Enabled","EnableP2PWakeupSolution","EnableUserIdleMonitoring","EnableWakeupProxy","MaxCPU","MaxMachinesPerManager","MinimumServersNeeded","NumOfDaysToKeep","NumOfMonthsToKeep","Port","PSComputerName","PSShowComputerName","SmsProviderObjectPath","WakeupProxyDirectAccessPrefixList","WakeupProxyFirewallFlags","WolPort","AllowWakeup")
              $Config = 'Power Management'
              $ConfigList = @()
              $ConfigList += "Allow power management of clients: " + $(Convert-Bool2Text $AgentConfig.Enabled)
              $ConfigList += "Allow users to exclude their device from power management: " + $(Convert-Bool2Text $AgentConfig.AllowUserToOptOutFromPowerPlan)
              #AllowWakeup: 0 = Not Configured; 1 = Enabled; 2 = Disabled
              switch ($AgentConfig.AllowWakeup) {
                0 { $Wakeup = 'Not Configured' }
                1 { $Wakeup = 'Enabled' }
                2 { $Wakeup = 'Disabled' }
              }
              $ConfigList += "Allow network wake-up: $Wakeup"
              $ConfigList += "Enable wake-up proxy: " + $(Convert-Bool2Text $AgentConfig.EnableWakeupProxy)
              $ConfigList += "Wake On LAN port number (UDP): $($AgentConfig.WolPort)"
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
              $ConfigList += "Manage Endpoint Protection client on client computers: " + $(Convert-Bool2Text $AgentConfig.EnableEP)
              $ConfigList += "Install Endpoint Protection client on client computers: " + $(Convert-Bool2Text $AgentConfig.InstallSCEPClient)
              #$ConfigList += "Automatically remove previously installed antimalware software before Endpoint Protection is installed: $($AgentConfig.Remove3rdParty)"
              $ConfigList += "Allow Endpoint Protection client installation and restarts outside maintenance windows. Maintenance windows must be at least 30 minutes long for client installation: " + $(Convert-Bool2Text $AgentConfig.OverrideMaintenanceWindow)
              $ConfigList += "For Windows Embedded devices with write filters, commit Endpoint Protection client installation (requires restart): " + $(Convert-Bool2Text $AgentConfig.PersistInstallation)
              $ConfigList += "Suppress any required computer restarts after the Endpoint Protection client is installed: " + $(Convert-Bool2Text $AgentConfig.SuppressReboot)
              If($AgentConfig.SuppressReboot -eq $false){
                $ConfigList += "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): $($AgentConfig.ForceRebootPeriod)"
              }Else{
                $ConfigList += "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): N/A"
              }
              $ConfigList += "Disable alternate sources (such as Microsoft Windows Update, Microsoft Windows Server Update Services, or UNC shares) for the initial definition update on client computers: " + $(Convert-Bool2Text $AgentConfig.DisableFirstSignatureUpdate)
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
              $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","RebootLogoffNotificationCountdownDuration","RebootLogoffNotificationFinalWindow","SmsProviderObjectPath","CountdownSnoozeInterval","RebootNotificationsDialog","EnforceReboot")
              $Config = 'Computer Restart'
              $ConfigList = @()
              $ConfigList += "Configuration Manager can force a device to restart: " + $(Convert-Bool2Text $AgentConfig.EnforceReboot)
              $ConfigList += "Display a temporary notification to the user that indicates the interval before the user is logged of or the computer restarts (minutes): $($AgentConfig.RebootLogoffNotificationCountdownDuration)"
              if ($Default){
                $ConfigList += "Display a dialog box that the user cannot close, which displays the countdown interval before the user is logged of or the computer restarts (minutes): $([string]$AgentConfig.RebootLogoffNotificationFinalWindow)"
              } else {
                $ConfigList += "Display a dialog box that the user cannot close, which displays the countdown interval before the user is logged of or the computer restarts (minutes): $([string]$AgentConfig.RebootLogoffNotificationFinalWindow / 60)"
              }
              $ConfigList += "Specify the snooze duration for computer restart countdown notifications (minutes): $($AgentConfig.CountdownSnoozeInterval)"
              $ConfigList += "When a deployment requires a restart, show a dialog windows to the user instead of a toast notification: " + $(Convert-Bool2Text $AgentConfig.RebootNotificationsDialog)
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
              $ConfigList += "Allow access to Cloud Distribution Point: " + $(Convert-Bool2Text $AgentConfig.AllowCloudDP)
              $ConfigList += "Automatically register new Windows 10 domain joined devices with Azure Active Directory: " + $(Convert-Bool2Text $AgentConfig.AutoAADJoin)
              $ConfigList += "Enable clients to use a cloud management gateway: " + $(Convert-Bool2Text $AgentConfig.AllowCMG)
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
              $ConfigList += "Specify how clients communicate on metered network connections: $($Usage)"
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
              $ConfigList += "Configure BranchCache: " + $(Convert-Bool2Text $AgentConfig.ConfigureBranchCache)
              $ConfigList += "Enable BranchCache: " + $(Convert-Bool2Text $AgentConfig.BranchCacheEnabled)
              $ConfigList += "Maximum BranchCache cache size (percentage of disk): $($AgentConfig.MaxBranchCacheSizePercent)"
              $ConfigList += "Configure client cache size: " + $(Convert-Bool2Text $AgentConfig.ConfigureCacheSize)
              $ConfigList += "Maximum cache size (MB): $($AgentConfig.MaxCacheSizeMB)"
              $ConfigList += "Maximum cache size (percentage of disk): $($AgentConfig.MaxCacheSizePercent)"
              #Client Cache duration appears in AgentID 6.  But since in GUI under 'Client cache settings', we will loop the config to find agent 6 and get the setting here.
              foreach ($AC in $ClientSetting.AgentConfigurations){
                If ($AC.AgentID -eq 6) {
                    $ConfigList += "Minimum duration before the client can remove cached content (minutes): $($AC.CacheTombstoneContentMinDuration/60)"
                }
              }
              $ConfigList += "Enable as peer cache source (Enable Configuration Manager client in full OS to share content - 1902 and earlier): " + $(Convert-Bool2Text $AgentConfig.CanBeSuperPeer)
              $ConfigList += "Port for initial network broadcast: $($AgentConfig.BroadcastPort)"
              $ConfigList += "Port for content download from peer: $($AgentConfig.HttpPort)"
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
              $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","SmsProviderObjectPath","SCBrandingColor","SCBrandingString","SCLogo","SCShowApplicationsTab","SCShowComplianceTab","SCShowInstallationTab","SCShowOptionsTab","SCShowOSDTab","SCShowUpdatesTab","SC_Old_Branding","SettingsXml","SC_UserPortal")
              $Config = 'Software Center'
              $ConfigList = @()
              switch ($AgentConfig.SC_UserPortal) {
                  0 { $UserPortal = 'Software Center' }
                  1 { $UserPortal = 'Company Portal' }
              }
              $ConfigList += "Select the user portal: $UserPortal"
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
                  if ($SCBrand.'application-catalog-link-hidden' -eq 'true'){
                    $ConfigList += "Hide Application Catalog in Software Center: Selected"
                }else{
                    $ConfigList += "Hide Application Catalog in Software Center: Not Selected"
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
                        'Options' {$tabvisibility = $tabvisibility + " &bull;  Options: $($tab.visible) <br />"}
                      }
                  }
                  IF (-not [string]::IsNullOrEmpty($SCBrand.'custom-tab'.'custom-tab-name')){
                    $tabvisibility = $tabvisibility + " &bull;  $($SCBrand.'custom-tab'.'custom-tab-name'): $($SCBrand.'custom-tab'.'custom-tab-content') <br />"
                  }else{
                    $tabvisibility = $tabvisibility + " &bull;  Custom Tab not defined. <br />"
                  }
                  $ConfigList += $tabvisibility.TrimEnd('<br />')
                  If (($SCBrand.'defaults-list'.'required-filter-default' -eq 'false') -or ($null -eq $SCBrand.'defaults-list'.'required-filter-default')){
                    $ConfigList += 'Default application filter: All'
                  }else{
                    $ConfigList += 'Default application filter: Required'
                  }
                  if (($SCBrand.'defaults-list'.'list-view-default' -eq 'false') -or ($null -eq $SCBrand.'defaults-list'.'list-view-default')){
                    $ConfigList += 'Default application view: Tile view'
                  }else{
                    $ConfigList += 'Default application view: List view'
                  }
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
              $KnownProps = @("AgentID","PSComputerName","PSShowComputerName","SmsProviderObjectPath","EnableWindowsDO","StampDOINC")
              $Config = 'Delivery Optimization'
              $ConfigList = @()
              If ($AgentConfig.EnableWindowsDO -eq 'True'){
                $WindowsDO = 'Yes'
              }Else{
                $WindowsDO = 'No'
              }
              $ConfigList += "Use Configuration Manager Boundary Groups for Delivery Optimization Group ID: $WindowsDO"
              $ConfigList += "Enable devices managed by Configuration Manager to use Delivery Optimization In-Network Cache server (Beta) for content download: " + $(Convert-Bool2Text $AgentConfig.StampDOINC)
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
          Write-Debug "$(Get-Date):   Client Settings Property not found."
        }

      }
    }
    Get-ClientSettingsConfigs -AgentConfigs $ClientSetting.AgentConfigurations
    }

    #Region Default Client Settings
    Write-ProgressEx -CurrentOperation "Default Client Setting" -Activity 'Client Settings' -Status 'Building default client settings object' -Id 5
    #These next lines build a client setting much like other custom clinet settings do with the AgentConfigurations property
    $DefaultCS = @{}
    $dc = (Get-CMClientSetting -Id 0)
    $DefaultCS += @{Description = $dc.Description}
    $DefaultCS += @{SettingsID = $dc.SettingsID}
    $DefaultCS += @{Priority = $dc.Priority}
    $DefaultCS += @{Name = $dc.Name}
    $DCS = @() #Default Client settings
    $DCS += (Get-CMClientSetting -Id 0 -Setting ComplianceSettings) + @{AgentID=1}
    $DCS += (Get-CMClientSetting -Id 0 -Setting SoftwareInventory) + @{AgentID=2}
    $DCS += (Get-CMClientSetting -Id 0 -Setting RemoteTools) + @{AgentID=3}
    $DCS += (Get-CMClientSetting -Id 0 -Setting ComputerAgent) + @{AgentID=4}
    #5 -- NAP agent.  depricated.
    #7 client cache duration setting?
    $DCS += (Get-CMClientSetting -Id 0 -Setting SoftwareMetering) + @{AgentID=8}
    $DCS += (Get-CMClientSetting -Id 0 -Setting SoftwareUpdates) + @{AgentID=9}
    $DCS += (Get-CMClientSetting -Id 0 -Setting UserAndDeviceAffinity) + @{AgentID=10}
    $DCS += (Get-CMClientSetting -Id 0 -Setting BackgroundIntelligentTransfer) + @{AgentID=11}
    $DCS += (Get-CMClientSetting -Id 0 -Setting MobileDevice) + @{AgentID=12}
    $DCS += (Get-CMClientSetting -Id 0 -Setting ClientPolicy) + @{AgentID=13}
    $DCS += (Get-CMClientSetting -Id 0 -Setting HardwareInventory) + @{AgentID=15}
    $DCS += (Get-CMClientSetting -Id 0 -Setting StateMessaging) + @{AgentID=16}
    $DCS += (Get-CMClientSetting -Id 0 -Setting SoftwareDeployment) + @{AgentID=17}
    $DCS += (Get-CMClientSetting -Id 0 -Setting PowerManagement) + @{AgentID=18}
    $DCS += (Get-CMClientSetting -Id 0 -Setting EndpointProtection) + @{AgentID=20}
    $DCS += (Get-CMClientSetting -Id 0 -Setting ComputerRestart) + @{AgentID=21}
    $DCS += (Get-CMClientSetting -Id 0 -Setting Cloud) + @{AgentID=22}
    $DCS += (Get-CMClientSetting -Id 0 -Setting MeteredNetwork) + @{AgentID=23}
    #25#The deadline randomization setting now appears in AgentID 25.  But since in GUI under 'Computer Agent' (4), we will loop that into configuration agent 4.
    $DCS += (Get-CMClientSetting -Id 0 -Setting ClientCache) + @{AgentID=27}
    $DCS += (Get-CMClientSetting -Id 0 -Setting WindowsAnalytics) + @{AgentID=29}
    $DCS += (Get-CMClientSetting -Id 0 -Setting SoftwareCenter) + @{AgentID=30}
    $DCS += (Get-CMClientSetting -Id 0 -Setting DeliveryOptimization) + @{AgentID=32}
    $DefaultCS += @{AgentConfigurations = $DCS}
    Write-ProgressEx -CurrentOperation "Default Client Setting" -Activity 'Client Settings' -Status 'Documenting default client settings' -Id 5
    $SettingInfo = @()
    $SettingInfo += "Client Settings Description: $($DefaultCS.Description)"
    $SettingInfo += "Client Settings ID: $($DefaultCS.SettingsID)"
    $SettingInfo += "Client Settings Priority: $($DefaultCS.Priority)"
    $SettingDescription = 'This is default client Device Setting.  If no other client settings are assigned, these are used.'
    Write-HTMLHeading -Level 3 -Text "Client Settings Name: $($DefaultCS.Name)" -File $FilePath
    Write-HtmlList -InputObject $SettingInfo -Description $SettingDescription -Level 2 -File $FilePath
    Get-ClientSettingsConfigs -AgentConfigs $DefaultCS.AgentConfigurations -Default
    #EndRegion Default Client Settings

}else{
    foreach ($ClientSetting in $AllClientSettings){
      Write-ProgressEx -CurrentOperation "Client Setting: $($ClientSetting.Name)" -Activity 'Client Settings' -Status 'Collecting basic settings' -Id 5
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
      Write-HTMLHeading -Level 3 -Text "Client Settings Name: $($ClientSetting.Name)" -File $FilePath -ExcludeTOC
      Write-HtmlList -InputObject $SettingInfo -Description $SettingDescription -Level 2 -File $FilePath
        If("$($ClientSetting.AssignmentCount)" -gt 0){
            $CSDeployments=Get-WmiObject -Query "SELECT * FROM SMS_ClientSettingsAssignment WHERE ClientSettingsID=$($ClientSetting.SettingsID)" -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider
            $CSDeploymentArray = @()
            foreach ($CSD in $CSDeployments){
                $CreationTime = [datetime]::ParseExact("$($CSD.CreationTime.Split('.')[0])",'yyyyMMddHHmmss',$null)
                $CSDeploymentArray += New-Object -TypeName psobject -Property @{'Collection ID'="$($CSD.CollectionID)";'Collection Name'="$($CSD.CollectionName)";'Date Created'="$CreationTime"}
            }
            Write-HTMLParagraph -Text "This client setting policy is deployed to the following $($CSDeploymentArray.count) collections:" -Level 3 -File $FilePath
            $CSDeploymentArray = $CSDeploymentArray | Select-Object 'Collection Name','Collection ID','Date Created'
            Write-HtmlTable -InputObject $CSDeploymentArray -Border 1 -Level 4 -File $FilePath
        }else{
            Write-HTMLParagraph -Text "This client setting policy is not deployed to any collections." -Level 3 -File $FilePath
        }
    }
}
Write-ProgressEx -CurrentOperation "Client Setting: $($ClientSetting.Name)" -Activity 'Client Settings' -Status 'Collection Complete' -Id 5 -Completed

Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation "Completed enumerating all Client Settings"
#endregion enumerating Client Policies

#region Security

Write-ProgressEx -CurrentOperation "Collecting all administrative users"
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
  If ($MaskAccounts){
    $LogonName = Set-AccountMask $Admin.LogonName
  }else{
    $LogonName = $Admin.LogonName
  }
  $AdminArray +=  New-Object -TypeName psobject -Property @{Name = $LogonName; 'Account Type' = $AccountType; 'Security Roles' = "$($($Admin.RoleNames) -join '--CRLF--')"; 'Security Scopes' = "$($($Admin.CategoryNames) -join '--CRLF--')"; Collections = "$($($Admin.CollectionNames) -join '--CRLF--')";}
}
$AdminArray = $AdminArray| Select-Object -Property 'Name','Account Type','Security Roles','Security Scopes','Collections'
Write-HtmlTable -InputObject $AdminArray -Border 1 -Level 3 -File $FilePath
#endregion Security


#region enumerating all custom Security roles
Write-ProgressEx -CurrentOperation "enumerating all custom build security roles"
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
    If($MaskAccounts){
        $SecRoleArray += New-Object -TypeName psobject -Property @{Name = $SecurityRole.RoleName; Description = $SecurityRole.RoleDescription; 'Copied From' = $((Get-CMSecurityRole -Id $SecurityRole.CopiedFromID).RoleName); Members = "$(($Members|Set-AccountMask) -join '--CRLF--')"; 'Role ID' = $SecurityRole.RoleID;}
    }else{
        $SecRoleArray += New-Object -TypeName psobject -Property @{Name = $SecurityRole.RoleName; Description = $SecurityRole.RoleDescription; 'Copied From' = $((Get-CMSecurityRole -Id $SecurityRole.CopiedFromID).RoleName); Members = "$($Members -join '--CRLF--')"; 'Role ID' = $SecurityRole.RoleID;}
    }
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

Write-ProgressEx -CurrentOperation "Enumerating all used accounts"
Write-HTMLHeading -Level 2 -Text 'Configured Accounts' -File $FilePath
$Accounts = Get-CMAccount
Write-HTMLParagraph -Text 'List of all accounts used for specific tasks in the site:' -Level 2 -File $FilePath

If(-not [string]::IsNullOrEmpty($Accounts)){
    $AccountsArray = @();

    foreach ($Account in $Accounts)
    {
      If($MaskAccounts){
        $AccountsArray += New-Object -TypeName psobject -Property @{'User Name'= $($Account.UserName|Set-AccountMask); 'Account Usage' = if ([string]::IsNullOrEmpty($Account.AccountUsage)) {'not assigned'} else {"$($Account.AccountUsage)"}; 'Site Code' = $Account.SiteCode};
      }else{
        $AccountsArray += New-Object -TypeName psobject -Property @{'User Name'= $Account.UserName; 'Account Usage' = if ([string]::IsNullOrEmpty($Account.AccountUsage)) {'not assigned'} else {"$($Account.AccountUsage)"}; 'Site Code' = $Account.SiteCode};
      }
    }

    $AccountsArray = $AccountsArray | Select-Object -Property 'User Name','Account Usage','Site Code'
    Write-HtmlTable -InputObject $AccountsArray -Border 1 -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No accounts in use in this site.' -Level 3 -File $FilePath
}
#endregion System Used Accounts

Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation "Done with Administration, next Assets and Compliance"

####
#region Assets and Compliance
####
$Progress = @{'Activity'="$Title started $($ScriptStartTime.ToShortTimeString())"; 'Status'='Assets and Compliance'}
Write-ProgressEx

Write-HTMLHeading -Level 1 -PageBreak -Text 'Assets and Compliance' -File $FilePath

#region enumerating all User Collections
Write-ProgressEx -CurrentOperation "Collecting User Collections"
Write-HTMLHeading -Level 2 -Text 'Summary of User Collections' -File $FilePath
$UserCollections = Get-CMUserCollection
$BuiltinUserCollections = @()
$CustomUserCollections = @()
$BuiltinUserCollections = $UserCollections|Where-Object {$_.CollectionID -like "SMS*"}
$CustomUserCollections = $UserCollections|Where-Object {$_.CollectionID -notlike "SMS*"}
if ($ListAllInformation)
{
  $BuiltInUCArray = @();
  Write-ProgressEx -CurrentOperation "Built-In User Collections"
  Write-HTMLParagraph -Text "Configuration Manager comes with a few built-in user collections.  Any number of custom collections can be defined in the console.  Below is a summary of both types." -Level 2 -File $FilePath
  Write-HTMLHeading -Level 3 -Text 'Built-In User Collections' -File $FilePath
  Write-HTMLParagraph -Text "There are $($BuiltinUserCollections.count) built-in default user collections.  Their names and member counts are listed below:" -Level 3 -File $FilePath
  foreach ($UserCollection in $BuiltinUserCollections)
  {
    Write-ProgressEx -CurrentOperation "Built-in User Collection: $($UserCollection.Name)" -Activity 'Built-In User Collections' -Status 'Collecting details' -Id 6
    # Get collection folder (not visible from Get-CMUserCollection cmdlet)
    $CollectionFolder = (Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class "SMS_Collection" -Filter "CollectionId = '$($UserCollection.CollectionID)'" -ComputerName $SMSProvider).ObjectPath
    $BuiltInUCArray += New-Object -TypeName psobject -Property @{'Collection Name' = $UserCollection.Name; 'Collection ID' = $UserCollection.CollectionID; 'Member Count' = $UserCollection.MemberCount; 'Schedule' = "$(Get-HumanReadableSchedule $UserCollection.RefreshSchedule)"; 'Folder' = "Root$CollectionFolder";};
  }
  Write-ProgressEx -CurrentOperation "Completed Built-in User Collection" -Activity 'Built-In User Collections' -Status 'Collecting details' -Id 6 -Completed
  $BuiltInUCArray = $BuiltInUCArray | Select-Object -Property 'Collection Name','Collection ID','Member Count','Schedule','Folder'
  Write-HtmlTable -InputObject $BuiltInUCArray -Border 1 -Level 4 -File $FilePath
  Write-ProgressEx -CurrentOperation "Completed writing Built-in User Collection"

  Write-ProgressEx -CurrentOperation "User Defined User Collections"
  Write-HTMLHeading -Level 3 -Text 'User Defined User Collections' -File $FilePath
  foreach ($UserCollection in $CustomUserCollections)
  {
    Write-ProgressEx -CurrentOperation "Custom User Collection: $($UserCollection.Name)" -Activity 'User Defined User Collections' -Status 'Collecting details' -Id 7
    $CollectionInfo = @()
    $CollectionName = "$($UserCollection.Name)"
    # Get collection folder (not visible from Get-CMUserCollection cmdlet)
    $CollectionFolder = (Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class "SMS_Collection" -Filter "CollectionId = '$($UserCollection.CollectionID)'" -ComputerName $SMSProvider).ObjectPath
    $CollectionInfo += "Folder: Root$CollectionFolder"
    $CollectionInfo += "Description: $($UserCollection.Comment)"
    $CollectionInfo += "Collection ID: $($UserCollection.CollectionID)"
    $CollectionInfo += "Total count of members: $($UserCollection.MemberCount)"
    $CollectionInfo += "Limiting Collection: $($UserCollection.LimitToCollectionName) ($($UserCollection.LimitToCollectionID))"
    Switch ($UserCollection.RefreshType)
    {
        1 {$UpdateSchedule = "No schedule configured"}
        2 {$UpdateSchedule = "Full update schedule only"}
        4 {$UpdateSchedule = "Incremental update only"}
        6 {$UpdateSchedule = "Full and Incremental updates configured"}
    }
    $CollectionInfo += "Selected Update Schedule: $UpdateSchedule"
    Write-HTMLHeading -Level 4 -Text $CollectionName -File $FilePath -ExcludeTOC

    Write-HtmlList -InputObject $CollectionInfo -Description "<u><b>Collection Information:</b></u>" -Level 4 -File $FilePath

    ### enumerating the Collection Membership Rules
    Write-HTMLParagraph -Level 4 -File $FilePath -Text '<u><b>Collection Membership Rules:</b></u>'
    $QueryRules = $Null
    $DirectRules = $Null
    $IncludeRules = $Null
    $ExcludeRules = $Null

    try {
        $DirectRules = $UserCollection | Get-CMUserCollectionDirectMembershipRule -ErrorAction SilentlyContinue
    }
    catch [System.Management.Automation.PropertyNotFoundException] {
        Write-Debug "$(Get-Date):   Collection Direct Rule info not found"
    }
    try {
        $QueryRules = $UserCollection | Get-CMUserCollectionQueryMembershipRule -ErrorAction SilentlyContinue
    }
    catch [System.Management.Automation.PropertyNotFoundException] {
        Write-Debug "$(Get-Date):   Collection Query Rule info not found"
    }
    try {
        $IncludeRules = $UserCollection | Get-CMUserCollectionIncludeMembershipRule -ErrorAction SilentlyContinue
    }
    catch [System.Management.Automation.PropertyNotFoundException] {
        Write-Debug "$(Get-Date):   Collection Include Rule info not found"
    }
    try {
        $ExcludeRules = $UserCollection | Get-CMUserCollectionExcludeMembershipRule -ErrorAction SilentlyContinue
    }
    catch [System.Management.Automation.PropertyNotFoundException] {
        Write-Debug "$(Get-Date):   Collection Include Rule info not found"
    }

    if (-not [string]::IsNullOrEmpty($QueryRules)) {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Query Rule(s):</b>'
        $QueryRulesArray = @();
        foreach ($QueryRule in $QueryRules) {
            $QueryRulesArray += New-Object -TypeName psobject -Property @{'Query Name'= $QueryRule.RuleName; 'Query Expression' = $($QueryRule.QueryExpression -replace ',',', ')}
        }
        Write-HtmlTable -InputObject $QueryRulesArray -Border 1 -Level 5 -File $FilePath
    }
    if (-not [string]::IsNullOrEmpty($DirectRules)) {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Direct Rule(s):</b>'
        $DirectRulesArray = @();
        foreach ($DirectRule in $DirectRules) {
            $DirectRulesArray += New-Object -TypeName psobject -Property @{'Resource Name'= $DirectRule.RuleName; 'Resource ID' = $DirectRule.ResourceId}
        }
        Write-HtmlTable -InputObject $DirectRulesArray -Border 1 -Level 5 -File $FilePath
    }
    if (-not [String]::IsNullOrEmpty($IncludeRules)) {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Include Rule(s):</b>'
        $IncludeRulesArray = @()
        foreach ($IncludeRule in $IncludeRules) {
            $IncludeRulesArray += New-Object -TypeName psobject -Property @{'Collection Name'= $IncludeRule.RuleName; 'Collection ID' = $IncludeRule.IncludeCollectionId}
        }
        Write-HtmlTable -InputObject $IncludeRulesArray -Border 1 -Level 5 -File $FilePath
    }
    if (-not [String]::IsNullOrEmpty($ExcludeRules)) {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Exclude Rule(s):</b>'
        $ExcludeRulesArray = @()
        foreach ($ExcludeRule in $ExcludeRules) {
            $ExcludeRulesArray += New-Object -TypeName psobject -Property @{'Collection Name'= $ExcludeRule.RuleName; 'Collection ID' = $ExcludeRule.ExcludeCollectionId}
        }
        Write-HtmlTable -InputObject $ExcludeRulesArray -Border 1 -Level 5 -File $FilePath
    }
    if (([String]::IsNullOrEmpty($IncludeRules)) -and ([String]::IsNullOrEmpty($ExcludeRules)) -and ([string]::IsNullOrEmpty($DirectRules)) -and ([string]::IsNullOrEmpty($QueryRules))){
    Write-HTMLParagraph -Level 5 -File $FilePath -Text 'No collection membership rules defined.'
    }
  }
  Write-ProgressEx -CurrentOperation "Completed Custom User Collections" -Activity 'User Defined User Collections' -Status 'Collecting details' -Id 7 -Completed
}
else
{
  Write-HTMLParagraph -Text "There are $($CustomUserCollections.count) User Defined User Collections.  These are in addition to the $($BuiltinUserCollections.count) built-in default user collections." -Level 2 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion enumerating all User Collections
Write-ProgressEx -CurrentOperation "Completed all user collections"

#region enumerating all Device Collections
Write-ProgressEx -CurrentOperation "Getting Device Collections."
Write-HTMLHeading -Level 2 -PageBreak -Text 'Summary of Device Collections' -File $FilePath
if ($ListAllInformation){
    Write-HTMLParagraph -Text 'This section contains a brief summary of built-in device collections as well as a more detailed summary of custom device collections.' -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'This section contains a brief summary of device collections.' -Level 3 -File $FilePath
}
$DeviceCollections = Get-CMDeviceCollection
$BuiltInDeviceCollections = @()
$CustomDeviceCollections = @()
$IncUpCollCount = @()
$ServiceWindowCollections = @()
$BuiltInDeviceCollections = $DeviceCollections | Where-Object {$_.IsBuiltIn -eq $true}
$CustomDeviceCollections = $DeviceCollections | Where-Object {$_.IsBuiltIn -eq $false}
$IncUpCollCount = ($CustomDeviceCollections | Where-Object {($_.RefreshType -eq 4) -or ($_.RefreshType -eq 6)}).count
$ServiceWindowCollections = $CustomDeviceCollections | Where-Object {$_.ServiceWindowsCount -gt 0}
Write-HtmlList -InputObject "$IncUpCollCount of $($CustomDeviceCollections.Count) have incremental updates Enabled." -Description "Incremental Update Summary:" -Level 3 -File $FilePath

#region Built-in Device Collections
Write-ProgressEx -CurrentOperation "Getting Built-in Device Collections"
if ($ListAllInformation)
{
  Write-HTMLHeading -Level 3 -Text 'Built-In Device Collections' -File $FilePath
  $DevCols = @()
  foreach ($DeviceCollection in $BuiltInDeviceCollections)
  {
    Write-ProgressEx -CurrentOperation "Found Built-in Device Collection: $($DeviceCollection.Name)" -Activity 'Built-In Device Collections' -Status 'Collecting details' -Id 8
    # Get collection folder (not visible from Get-CMDeviceCollection cmdlet)
    $CollectionFolder = (Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class "SMS_Collection" -Filter "CollectionId = '$($DeviceCollection.CollectionID)'" -ComputerName $SMSProvider).ObjectPath
    $DevCols += New-Object -TypeName psobject -Property @{'Name' = "$($DeviceCollection.Name)"; 'Collection ID' = "$($DeviceCollection.CollectionID)"; 'Member Count' = "$($DeviceCollection.MemberCount)"; 'Schedule' = "$(Get-HumanReadableSchedule $DeviceCollection.RefreshSchedule)"; 'Folder' = "Root$CollectionFolder";}
  }
  $DevCols = $DevCols | Select-Object 'Name','Collection ID','Member Count','Schedule','Folder'
  Write-HTMLParagraph -Level 4 -File $FilePath -Text 'Summary of membership of the built-in device collections:'
  Write-HtmlTable -InputObject $DevCols -Border 1 -Level 5 -File $FilePath
  Write-ProgressEx -CurrentOperation 'Found Built-in Device Collection' -Activity 'Built-In Device Collections' -Status 'Collecting details' -Id 8 -Completed
  Write-ProgressEx -CurrentOperation 'Completed Built-in Device Collections'
#endregion Built-in Device Collections

#region User Defined Device Collections
  Write-ProgressEx -CurrentOperation 'User Defined Device Collections'
  Write-HTMLHeading -Level 3 -Text 'User Defined Device Collections' -File $FilePath
  Write-HTMLParagraph -Level 3 -File $FilePath -Text "There are $($CustomDeviceCollections.count) custom defined collections in this site."
#region Service Window Device Collections
  Write-HTMLParagraph -Level 3 -File $FilePath -Text "There are $($ServiceWindowCollections.count) custom collections with defined service windows.  These are listed below:"
  $SWCollections = @()
  foreach ($collection in $ServiceWindowCollections){
    $LinkID = ($Collection.Name).Replace(' ','')
    $SWCollections +=  Write-HtmliLink -LinkID $LinkID -Text "$($Collection.Name) ($($Collection.CollectionID))"
  }
  If($SWCollections.count -eq 0){$SWCollections += "None Found"}
  Write-HtmlList -InputObject $SWCollections -Level 3 -File $FilePath
#endregion Service Window Device Collections
  foreach ($DeviceCollection in $CustomDeviceCollections)
  {
    Write-ProgressEx -CurrentOperation "Found user defined device collection: $($DeviceCollection.Name)" -Activity 'Device collection' -Status 'Collecting details' -Id 9
    $CollectionInfo = @()
    $CollectionName = "$($DeviceCollection.Name)"
    # Get collection folder (not visible from Get-CMDeviceCollection cmdlet)
    $CollectionFolder = (Get-WmiObject -Namespace "root\sms\site_$SiteCode" -Class "SMS_Collection" -Filter "CollectionId = '$($DeviceCollection.CollectionID)'" -ComputerName $SMSProvider).ObjectPath
    $CollectionInfo += "Folder: Root$CollectionFolder"
    $CollectionInfo += "Description: $($DeviceCollection.Comment)"
    $CollectionInfo += "Collection ID: $($DeviceCollection.CollectionID)"
    $CollectionInfo += "Total count of members: $($DeviceCollection.MemberCount)"
    $CollectionInfo += "Limiting Collection: $($DeviceCollection.LimitToCollectionName) ($($DeviceCollection.LimitToCollectionID))"
    Switch ($DeviceCollection.RefreshType)
    {
        1 {$UpdateSchedule = "No schedule configured"}
        2 {$UpdateSchedule = "Full update schedule only"}
        4 {$UpdateSchedule = "Incremental update only"}
        6 {$UpdateSchedule = "Full and Incremental updates configured"}
    }
    $CollectionInfo += "Selected Update Schedule: $UpdateSchedule"
    Write-HTMLHeading -Level 4 -Text $CollectionName -File $FilePath -ExcludeTOC

    Write-HtmlList -InputObject $CollectionInfo -Description "<u><b>Collection Information:</b></u>" -Level 4 -File $FilePath
    Write-HTMLParagraph -Level 4 -File $FilePath -Text '<u><b>Collection Maintenance Windows:</b></u>'
    If ($DeviceCollection.ServiceWindowsCount -gt 0) {
        $ServiceWindows = Get-CMMaintenanceWindow -CollectionId $DeviceCollection.CollectionID
        Write-Debug "$(Get-Date):   Enumerating Maintenance Windows for collection: $($DeviceCollection.Name)"
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
                $ServiceWindowArray += New-Object -TypeName psobject -Property @{'Name' = $SWName; 'Start Time' = $StartTime; 'UTC' = $UTCTime; 'Duration' = $Duration; 'Recurrence' = $WindowRecurence; 'Type' = $WindowType; 'Enabled' = $WindowEnabled}
            }
        $ServiceWindowArray = $ServiceWindowArray | Select-Object 'Name','Start Time','UTC','Duration','Recurrence','Type','Enabled'
        Write-HtmlTable -InputObject $ServiceWindowArray -Border 1 -Level 5 -File $FilePath
    } else {
        Write-HTMLParagraph -Level 4 -File $FilePath -Text 'No maintenance windows configured on this collection.'
    }
        ### enumerating the Collection Membership Rules
        Write-HTMLParagraph -Level 4 -File $FilePath -Text '<u><b>Collection Membership Rules:</b></u>'
        $QueryRules = $Null
        $DirectRules = $Null
        $IncludeRules = $Null
        $ExcludeRules = $Null

        try {
            $DirectRules = $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule -ErrorAction Stop
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Debug "$(Get-Date):   Collection Direct Rule info not found"
        }
        try {
            $QueryRules = $DeviceCollection | Get-CMDeviceCollectionQueryMembershipRule -ErrorAction Stop
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Debug "$(Get-Date):   Collection Query Rule info not found"
        }
        try {
            $IncludeRules = $DeviceCollection | Get-CMDeviceCollectionIncludeMembershipRule -ErrorAction Stop
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Debug "$(Get-Date):   Collection Include Rule info not found"
        }
        try {
            $ExcludeRules = $DeviceCollection | Get-CMDeviceCollectionExcludeMembershipRule -ErrorAction Stop
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            Write-Debug "$(Get-Date):   Collection Include Rule info not found"
        }

        if (-not [string]::IsNullOrEmpty($QueryRules)) {
            Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Query Rule(s):</b>'
            $QueryRulesArray = @();
            foreach ($QueryRule in $QueryRules) {
                $QueryRulesArray += New-Object -TypeName psobject -Property @{'Query Name'= $QueryRule.RuleName; 'Query Expression' = $($QueryRule.QueryExpression -replace ',',', ')}
            }
            Write-HtmlTable -InputObject $QueryRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (-not [string]::IsNullOrEmpty($DirectRules)) {
            Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Direct Rule(s):</b>'
            $DirectRulesArray = @();
            foreach ($DirectRule in $DirectRules) {
                $DirectRulesArray += New-Object -TypeName psobject -Property @{'Resource Name'= $DirectRule.RuleName; 'Resource ID' = $DirectRule.ResourceId}
            }
            Write-HtmlTable -InputObject $DirectRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (-not [String]::IsNullOrEmpty($IncludeRules)) {
            Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Include Rule(s):</b>'
            $IncludeRulesArray = @()
            foreach ($IncludeRule in $IncludeRules) {
                $IncludeRulesArray += New-Object -TypeName psobject -Property @{'Collection Name'= $IncludeRule.RuleName; 'Collection ID' = $IncludeRule.IncludeCollectionId}
            }
            Write-HtmlTable -InputObject $IncludeRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (-not [String]::IsNullOrEmpty($ExcludeRules)) {
            Write-HTMLParagraph -Level 4 -File $FilePath -Text '<b>Exclude Rule(s):</b>'
            $ExcludeRulesArray = @()
            foreach ($ExcludeRule in $ExcludeRules) {
                $ExcludeRulesArray += New-Object -TypeName psobject -Property @{'Collection Name'= $ExcludeRule.RuleName; 'Collection ID' = $ExcludeRule.ExcludeCollectionId}
            }
            Write-HtmlTable -InputObject $ExcludeRulesArray -Border 1 -Level 5 -File $FilePath
        }
        if (([String]::IsNullOrEmpty($IncludeRules)) -and ([String]::IsNullOrEmpty($ExcludeRules)) -and ([string]::IsNullOrEmpty($DirectRules)) -and ([string]::IsNullOrEmpty($QueryRules))){
        Write-HTMLParagraph -Level 5 -File $FilePath -Text 'No collection membership rules defined.'
        }
    }
  Write-ProgressEx -CurrentOperation "Completed user defined device collections" -Activity 'Device collection' -Status 'Collecting details' -Id 9 -Completed
}else{
  Write-HTMLParagraph -Text "There are $($CustomDeviceCollections.count) User Defined Device collections.  These are in addition to the $($BuiltInDeviceCollections.count) built-in default device collections." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation 'Completed User Defined Device Collections'
#endregion User Defined Device Collections


#region enumerating all Compliance Settings
Write-ProgressEx -CurrentOperation "Working on Compliance Settings"
Write-HTMLHeading -Level 2 -PageBreak -Text 'Compliance Settings' -File $FilePath
Write-HTMLParagraph -Text 'This section contains a summary of all configuration items, baselines, Settings, Conditional Access, and other configurable resources.' -Level 3 -File $FilePath
#region enumerating all Configuration Items and baselines.
Write-ProgressEx -CurrentOperation 'Configuration Items'
Write-HTMLHeading -Level 3 -Text 'Configuration Items' -File $FilePath
$CIs = Get-CMConfigurationItem -Fast | Where-Object {$_.IsUserDefined -eq "true"}
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
Write-ProgressEx -CurrentOperation 'Completed Configuration Items'
Write-ProgressEx -CurrentOperation 'Configuration Baselines'
Write-HTMLHeading -Level 3 -Text 'Configuration Baselines' -File $FilePath
$CBs = Get-CMBaseline | Where-Object {$_.IsUserDefined -eq "true"}
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
Write-ProgressEx -CurrentOperation 'Completed Configuration Baselines'
#endregion enumerating all Configuration Items and baselines.

#region enumerating Configuration Policies
Write-ProgressEx -CurrentOperation 'Working on Configuration Policies'
$CMPolicies=@()
$CMPolicies=Get-CMConfigurationPolicy -fast | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-ProgressEx -CurrentOperation "$($CMPolicies.count) Configuration Policies found"

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
$EdgeBrowser = @()
$KFMSettings = @()

Write-ProgressEx -CurrentOperation "Processing [$($CMPolicies.count)] Configuration Policies"
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
        {'SettingsAndPolicy:SMS_EdgeBrowserSettings' -in $_.CategoryInstance_UniqueIDs}{
            $EdgeBrowser += New-Object -TypeName psobject -Property @{Name = "$($CMPolicy.LocalizedDisplayName)";'Modified By' = "$($CMPolicy.LastModifiedBy)";'Modified' = "$($CMPolicy.DateLastModified)"; Deployed = "$($CMPolicy.IsAssigned)"}
        }
        {'SettingsAndPolicy:SMS_OneDriveKnownFolderMigrationSettings' -in $_.CategoryInstance_UniqueIDs}{
            $KfmXml = [xml](Get-CMConfigurationPolicy -Id $CMPolicy.CI_ID).SDMPackageXML
            $KfmDeployed = $CMPolicy.IsAssigned
            $KfmName = $KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Annotation.DisplayName.text
            $KfmDescription = $KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Annotation.Description.text
            $KfmPlatforms = $KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.PlatformApplicabilityCondition.OperatingSystemExpression.Operands.RuleExpression.RuleId -join '; '
            $KfmSilentMove = ($KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Rules.rule| Where-Object {$_.ID -eq 'Rule_OneDrive_KFMSilentOptIn'}).expression.operands.constantvalue.value
            If (!([string]::IsNullOrEmpty($KfmSilentMove))){
                $KFMSilentOptIn = 'Enabled'
                $KFMOptInWithWizard = 'Disabled'
                $O365ID = $KfmSilentMove
            }
            $KfmPromptToMove = ($KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Rules.rule| Where-Object {$_.ID -eq 'Rule_OneDrive_KFMOptInWithWizard'}).expression.operands.constantvalue.value
            If (!([string]::IsNullOrEmpty($KfmPromptToMove))){
                $KFMSilentOptIn = 'Disabled'
                $KFMOptInWithWizard = 'Enabled'
                $O365ID = $KfmPromptToMove
            }
            $KfmNoteAfter = ($KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Rules.rule| Where-Object {$_.ID -eq 'Rule_OneDrive_KFMSilentOptInWithNotification'}).expression.operands.constantvalue.value #1 = Enabled
            If ($KfmNoteAfter -eq 1){
                $KfmNoteAfter = 'Enabled'
            } else {
                $KfmNoteAfter = 'Disabled'
            }
            $KfmBlockOptOut = ($KfmXml.DesiredConfigurationDigest.ConfigurationPolicy.Rules.rule| Where-Object {$_.ID -eq 'Rule_OneDrive_KFMBlockOptOut'}).expression.operands.constantvalue.value #1 = Enabled
            If ($KfmBlockOptOut -eq 1){
                $KfmBlockOptOut = 'Enabled'
            } else {
                $KfmBlockOptOut = 'Disabled'
            }
            $KFMSettings += [pscustomobject]@{
                Name           = $KfmName
                Description    = $KfmDescription
                Platforms      = $KfmPlatforms
                O365ID         = $O365ID
                PromptToMove   = $KFMOptInWithWizard
                SilentMove     = $KFMSilentOptIn
                NotifyComplete = $KfmNoteAfter
                BlockOptOut    = $KfmBlockOptOut
                Deployed       = $KfmDeployed
                CI_ID          = $CMPolicy.CI_ID
            }
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
Write-HTMLHeading -Level 3 -Text 'OneDrive For Business Profiles' -File $FilePath
if ($KFMSettings.count -gt 0) {
    foreach ($KFMsetting in $KFMSettings) {
        $listArray = @()
        $listArray += "Supported Platforms: $($KFMsetting.Platforms)"
        $listArray += "Office 365 TenantId: $($KFMsetting.O365ID)"
        $listArray += "Prompt users to move Windows known folders to OneDrive: $($KFMsetting.PromptToMove)"
        $listArray += "Silently move Windows known folders to OneDrive: $($KFMsetting.SilentMove)"
        $listArray += "Show notification to users after folders have been redirected: $($KFMsetting.NotifyComplete)"
        $listArray += "Prevent users from redirecting their Windows known folders back to their PC: $($KFMsetting.BlockOptOut)"
        Write-HtmlList -Title "$($KFMsetting.Name)" -Description "Description: $($KFMsetting.Description)" -InputObject $listArray -Level 4 -File $FilePath
        If ($KFMSetting.Deployed -eq "True"){
            $KFMDeployments = Get-CMConfigurationPolicyDeployment -SmsObjectId 16780575 | foreach {Get-CMCollection -Id $_.TargetCollectionID}|select CollectionID,Name
            Write-HTMLParagraph -Text 'Deployments:' -Level 4 -File $FilePath
            Write-HtmlTable -InputObject $KFMDeployments -Level 4 -File $FilePath
        } else {
            Write-HTMLParagraph -Text 'Deployments: None' -Level 4 -File $FilePath
        }
    }
}else{
    Write-HTMLParagraph -Text 'No OneDrive For Business Profiles defined in site.' -Level 4 -File $FilePath
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
Write-HTMLHeading -Level 3 -Text 'Microsoft Edge Browser Profiles' -File $FilePath
if ($EdgeBrowser.count -gt 0) {
    $EdgeBrowser = $EdgeBrowser | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $EdgeBrowser -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'Microsoft Edge Browser Profiles defined in site.' -Level 4 -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Completed Configuration Policies'
#endregion enumerating Configuration Policies

Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation 'Completed all Compliance Settings'
#endregion enumerating all Compliance Settings

#region Endpoint Protection
Write-ProgressEx -CurrentOperation 'Endpoint Protection settings'
Write-HTMLHeading -Level 2 -PageBreak -Text 'Endpoint Protection' -File $FilePath
Write-HTMLParagraph -Text 'This section contains a summary of all Endpoint Security configuration settings.  This includes System Center Endpoint Protection (Antimalware), Firewall, Windows Defender ATP, and Device Guard Policies.' -Level 3 -File $FilePath
#region Antimalware
Write-ProgressEx -CurrentOperation 'Antimalware Policies'
Write-HTMLHeading -Level 3 -Text 'Antimalware Policies' -File $FilePath
if (-not ($Null -eq $(Get-CMEndpointProtectionPoint))){
    $AntiMalwarePolicies = Get-CMAntimalwarePolicy
    if (-not [string]::IsNullOrEmpty($AntiMalwarePolicies)){
        foreach ($AntiMalwarePolicy in $AntiMalwarePolicies){
                if ($AntiMalwarePolicy.Name -eq 'Default Client Antimalware Policy'){
                    if($ListAllInformation){
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
                        $listArray += Write-HtmlList -Description 'Excluded processes:' -InputObject $ProcessArray -Level 1
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
                        $AgentConfig = $AntiMalwarePolicy.AgentConfiguration
                        Write-HTMLHeading -Level 4 -Text "$($AntiMalwarePolicy.Name)" -File $FilePath
                        Write-HTMLParagraph -Text "Description: $($AntiMalwarePolicy.Description)" -Level 4 -File $FilePath
                        $listTitle = 'Basic Protection Settings'
                        $listArray = @()
                        $listArray += "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                        $listArray += "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                        Write-HtmlList -Title $listTitle -InputObject $listArray -Level 4 -File $FilePath
                    }
                }else{
                    if($ListAllInformation){
                        $AgentConfig_custom = $AntiMalwarePolicy.AgentConfigurations
                        Write-HTMLHeading -Level 4 -Text "$($AntiMalwarePolicy.Name)" -File $FilePath
                        If("$($AntiMalwarePolicy.Description)" -ne ""){
                            Write-HTMLParagraph -Text "Description: $($AntiMalwarePolicy.Description)" -Level 4 -File $FilePath
                        }
                        If("$($AntiMalwarePolicy.AssignmentCount)" -gt 0){
                            $APDeployments=Get-WmiObject -Query "SELECT * FROM SMS_ClientSettingsAssignment WHERE ClientSettingsID=$($AntiMalwarePolicy.SettingsID)" -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider
                            $APDeploymentArray = @()
                            foreach ($APD in $APDeployments){
                                $CreationTime = [datetime]::ParseExact("$($APD.CreationTime.Split('.')[0])",'yyyyMMddHHmmss',$null)
                                $APDeploymentArray += New-Object -TypeName psobject -Property @{'Collection ID'="$($APD.CollectionID)";'Collection Name'="$($APD.CollectionName)";'Date Created'="$CreationTime"}
                            }
                            Write-HTMLParagraph -Text "This antimalware policy is deployed to the following $($APDeploymentArray.count) collections:" -Level 4 -File $FilePath
                            $APDeploymentArray = $APDeploymentArray | Select-Object 'Collection Name','Collection ID','Date Created'
                            Write-HtmlTable -InputObject $APDeploymentArray -Border 1 -Level 4 -File $FilePath
                        }else{
                            Write-HTMLParagraph -Text "This antimalware policy is not deployed to any collections." -Level 4 -File $FilePath
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
                    }else{
                        $AgentConfig_custom = $AntiMalwarePolicy.AgentConfigurations
                        Write-HTMLHeading -Level 4 -Text "$($AntiMalwarePolicy.Name)" -File $FilePath
                        If("$($AntiMalwarePolicy.Description)" -ne ""){
                            Write-HTMLParagraph -Text "Description: $($AntiMalwarePolicy.Description)" -Level 4 -File $FilePath
                        }
                        If("$($AntiMalwarePolicy.AssignmentCount)" -gt 0){
                            $APDeployments=Get-WmiObject -Query "SELECT * FROM SMS_ClientSettingsAssignment WHERE ClientSettingsID=$($AntiMalwarePolicy.SettingsID)" -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider
                            $APDeploymentArray = @()
                            foreach ($APD in $APDeployments){
                                $CreationTime = [datetime]::ParseExact("$($APD.CreationTime.Split('.')[0])",'yyyyMMddHHmmss',$null)
                                $APDeploymentArray += New-Object -TypeName psobject -Property @{'Collection ID'="$($APD.CollectionID)";'Collection Name'="$($APD.CollectionName)";'Date Created'="$CreationTime"}
                            }
                            Write-HTMLParagraph -Text "This antimalware policy is deployed to the following $($APDeploymentArray.count) collections:" -Level 4 -File $FilePath
                            $APDeploymentArray = $APDeploymentArray | Select-Object 'Collection Name','Collection ID','Date Created'
                            Write-HtmlTable -InputObject $APDeploymentArray -Border 1 -Level 4 -File $FilePath
                        }else{
                            Write-HTMLParagraph -Text "This antimalware policy is not deployed to any collections." -Level 4 -File $FilePath
                        }
                        $listArray = @()
                        foreach ($Agentconfig in $AgentConfig_custom){
                            switch ($AgentConfig.AgentID){
                                201 {
                                        $listArray += "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                    }
                                204 {
                                        $listArray += "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                    }
                            }
                        }
                        Write-HtmlList -Title 'Basic Protection Settings' -InputObject $listArray -Level 4 -File $FilePath
                    }
                }
            }
    }else{
        Write-HTMLParagraph -Text 'There are no Anti Malware Policies configured.' -Level 3 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text 'There is no Endpoint Protection Point enabled in this site.' -Level 3 -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Completed Antimalware Policies'
#endregion Antimalware
#region firewall and Device Guard
Write-ProgressEx -CurrentOperation 'Firewall and Device Guard'
$FWPolicies = Get-CMConfigurationPolicy -Fast | Where-Object {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_FirewallSettings'} | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
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
$ATPPolicies = Get-CMConfigurationPolicy -Fast | Where-Object {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_AdvancedThreatProtectionSettings'} | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Windows Defender ATP Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($ATPPolicies)) {
    $ATPArray = @()
    foreach ($ATP in $ATPPolicies){
        $ATPArray += New-Object -TypeName psobject -Property @{'Name'=$ATP.LocalizedDisplayName;'Modified By'=$ATP.LastModifiedBy;'Modified'=$ATP.DateLastModified;'Deployed'=$ATP.IsAssigned}
    }
    $ATPArray = $ATPArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $ATPArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No ATP Policies defined in site.' -Level 4 -File $FilePath
}
$ExploitGuardPolicies = Get-CMConfigurationPolicy -Fast | Where-Object {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_ExploitGuardSettings'} | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Windows Defender Exploit Guard Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($ExploitGuardPolicies)) {
    $ExploitGuardArray = @()
    foreach ($EGP in $ExploitGuardPolicies){
        $ExploitGuardArray += New-Object -TypeName psobject -Property @{'Name'=$EGP.LocalizedDisplayName;'Modified By'=$EGP.LastModifiedBy;'Modified'=$EGP.DateLastModified;'Deployed'=$EGP.IsAssigned}
    }
    $ExploitGuardArray = $ExploitGuardArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $ExploitGuardArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Exploit Guard Policies defined in site.' -Level 4 -File $FilePath
}
$ApplicationGuardPolicies = Get-CMConfigurationPolicy -Fast | Where-Object {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_WindowsDefenderApplicationGuard'} | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Windows Defender Application Guard Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($ApplicationGuardPolicies)) {
    $ApplicationGuardArray = @()
    foreach ($AGP in $ApplicationGuardPolicies){
        $ApplicationGuardArray += New-Object -TypeName psobject -Property @{'Name'=$AGP.LocalizedDisplayName;'Modified By'=$AGP.LastModifiedBy;'Modified'=$AGP.DateLastModified;'Deployed'=$AGP.IsAssigned}
    }
    $ApplicationGuardArray = $ApplicationGuardArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $ApplicationGuardArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Application Guard Policies defined in site.' -Level 4 -File $FilePath
}
$DeviceGuardPolicies = Get-CMConfigurationPolicy -Fast | Where-Object {$_.CategoryInstance_UniqueIDs -contains 'SettingsAndPolicy:SMS_DeviceGuardSettings'} | Select-Object CategoryInstance_UniqueIDs,LocalizedDisplayName,LocalizedCategoryInstanceNames,CI_ID,LastModifiedBy,DateLastModified,IsAssigned
Write-HTMLHeading -Level 3 -Text 'Windows Defender Application Control Policies' -File $FilePath
if (-not [string]::IsNullOrEmpty($DeviceGuardPolicies)) {
    $DeviceGuardArray = @()
    foreach ($DGP in $DeviceGuardPolicies){
        $DeviceGuardArray += New-Object -TypeName psobject -Property @{'Name'=$DGP.LocalizedDisplayName;'Modified By'=$DGP.LastModifiedBy;'Modified'=$DGP.DateLastModified;'Deployed'=$DGP.IsAssigned}
    }
    $DeviceGuardArray = $DeviceGuardArray | Select-Object 'Name','Modified By','Modified','Deployed'
    Write-HtmlTable -InputObject $DeviceGuardArray -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'No Application Control Policies defined in site.' -Level 4 -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Completed Firewall and Device Guard'
#endregion firewall and Device Guard
Write-ProgressEx -CurrentOperation 'Completed Endpoint Protection settings'
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
Write-ProgressEx -CurrentOperation "Done with Assets and Compliance, next Software Library"
#endregion Assets and Compliance

####
#region Software Library
####
$Progress = @{'Activity'="$Title started $($ScriptStartTime.ToShortTimeString())"; 'Status'='Software Library'}
Write-ProgressEx -CurrentOperation 'Beginning Software Library'

Write-HTMLHeading -Level 1 -PageBreak -Text 'Software Library' -File $FilePath

#region Application Management
Write-HTMLHeading -Level 2 -PageBreak -Text 'Application Management' -File $FilePath

#region Applications
Write-ProgressEx -CurrentOperation 'Processing Applications'
Write-HTMLHeading -Level 3 -Text 'Applications' -File $FilePath
#$Applications = Get-WmiObject -Class sms_applicationlatest -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider
#Get-CMApplication | select LocalizedDisplayName,LocalizedDescription,Manufacturer,SoftwareVersion,PackageID,ISExpired,ISDeployed,NumberOfDeploymentTypes
$Applications = Get-CMApplication|Sort-Object LocalizedDisplayName
if ($ListAllInformation -or $ListAppDetails){
    if (-not [string]::IsNullOrEmpty($Applications)) {
        Write-HTMLParagraph -Text "Below is a summary of all $($Applications.Count) application installers defined in this site. These are applications that are installed with the configuration manager application model.  Packages are covered later in the documentation." -Level 3 -File $FilePath
        foreach ($App in $Applications) {
            #region Application Details
            Write-ProgressEx -CurrentOperation "Application: $($App.LocalizedDisplayName)" -Activity 'Configuration Manager Application' -Status 'Application details' -Id 10
            Write-HTMLHeading -Level 4 -Text "$($App.LocalizedDisplayName)" -File $FilePath
            if ($App.LocalizedDescription -ne ""){
                Write-HTMLParagraph -Text "Applicaton Description: $($App.LocalizedDescription)" -Level 4 -File $FilePath
            }
            Write-Debug "$(Get-Date):   Processing General application Info.."
            $ListDescription = "General Application Information"
            $AppList = @()
            $AppList += "Created by: $($App.CreatedBy)"
            $AppList += "Date created: $($App.DateCreated)"
            $AppList += "Date published: $($app.EffectiveDate)"
            $AppList += "Modified date: $($app.DateLastModified)"
            $AppList += "Modified by: $($app.LastModifiedBy)"
            $AppList += "Revision number: $($app.CIVersion)"
            $AppList += "Publisher: $($App.Manufacturer)"
            $AppList += "Software version: $($App.SoftwareVersion)"
            $AppList += "CM Package ID: $($App.PackageID)"
            $AppList += "Retired: $($App.IsExpired)"
            $AppList += "Deployed: $($App.ISDeployed)"
            Write-HtmlList -InputObject $AppList -Description $ListDescription -Level 5 -File $FilePath
            $ListDescription = ""
            Write-Debug "$(Get-Date):   Completed General application Info.."

            $PackageXML = [xml]$App.SDMPackageXML

            Write-Debug "$(Get-Date):   Processing application Software Center info.."
            $ListDescription = "Software Center Details"
            $AppList = @()
            $AppList += "Localized App Name:  $($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.Title)"
            If (!([string]::IsNullOrEmpty($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.UserCategories.Tag))){
                $AppList += "User categories:   $($($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.UserCategories.Tag) -join '; ')"
            }
            $AppList += "User documentation:  $($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.InfoUrl)"
            $AppList += "Link text:  $($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.InfoUrlText)"
            $AppList += "Localized description:  $($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.Description)"
            $AppList += "Keywords:  $($($xml.AppMgmtDigest.Application.DisplayInfo.info.tags.tag) -join ';')"
            If ([string]::IsNullOrEmpty($PackageXML.AppMgmtDigest.Application.DisplayInfo.info.Icon)){
                $AppList += "Custom Icon:  None"
            } else {
                $AppList += "Custom Icon:  Yes"
            }
            $AppList += "Featured App: $($App.Featured)"
            Write-HtmlList -InputObject $AppList -Description $ListDescription -Level 5 -File $FilePath
            $ListDescription = ""
            Write-Debug "$(Get-Date):   Completed application Software Center info.."

            Write-Debug "$(Get-Date):   Processing application content locations.."
            #Content Locations:
            #SELECT * FROM SMS_DPGroupContentInfo
            #SELECT * FROM SMS_DPContentInfo
            Write-Debug "$(Get-Date):   Completed application content locations.."

            Write-Debug "$(Get-Date):   Processing application distribution settings.."
            $ListDescription = "Distribution Details"
            $AppList = @()
            switch ($xml.AppMgmtDigest.Application.HighPriority) {
                1 { $Priority = "High" }
                2 { $Priority = "Medium" }
                3 { $Priority = "Low" }
                Default {$Priority = "Medium"}
            }
            $AppList += "Distribution Priority:  $Priority"
            If ($xml.AppMgmtDigest.Application.SendToProtectedDP -eq 'true'){
                $AppList += "Enable for On-demand Distribution: True"
            } else {
                $AppList += "Enable for On-demand Distribution: False"
            }
            If ($xml.AppMgmtDigest.Application.AutoDistribute -eq 'true'){
                $AppList += "Prestaged Distribution Point settings: Automatically downloadcontent when packages are assigned to distribution points"
            }
            elseIf ($xml.AppMgmtDigest.Application.DownloadDelta -eq 'true'){
                $AppList += "Prestaged Distribution Point settings: Download only content changes to the distribution point"
            }
            ELSE{
                $AppList += "Prestaged Distribution Point settings: Manually copy the content in this package to the distribution point"
            }
            Write-HtmlList -InputObject $AppList -Description $ListDescription -Level 5 -File $FilePath
            $ListDescription = ""
            Write-Debug "$(Get-Date):   Completed application distribution settings.."

            Write-Debug "$(Get-Date):   Processing Supersedence info.."
            $ListDescription = "Supersedence details"
            $AppList = @()
            $AppList += "Supersedes other applications:  $($app.IsSuperseding)"
            $AppList += "Is superseded by another application:  $($app.IsSuperseded)"
            if ($xml.AppMgmtDigest.Application.DisplaySupersedes -eq 'true'){
                $AppList += "Allow user to see this app and all apps that it supersedes in Software Center: Yes"
            } else {
                $AppList += "Allow user to see this app and all apps that it supersedes in Software Center: No"
            }
            Write-HtmlList -InputObject $AppList -Description $ListDescription -Level 5 -File $FilePath
            $ListDescription = ""
            Write-Debug "$(Get-Date):   Completed supersedence details"
            #endregion Application Details
            
            #region deployment types
            Write-ProgressEx -CurrentOperation "Application: $($App.LocalizedDisplayName)" -Activity 'Configuration Manager Application' -Status 'Application deployment types' -Id 10
            Write-Debug "$(Get-Date):   Processing deployment types for: $($App.LocalizedDisplayName)"
            $DTs = $PackageXML.AppMgmtDigest.DeploymentType
            if (-not [string]::IsNullOrEmpty($DTs)) {
                foreach ($DT in $DTs) {
                    Write-HTMLHeading -Level 5 -Text "Deployment Type: $($DT.title.'#text')" -File $FilePath
                    $DTListData = @()
                    $DTSection = "General"
                    $DTListData += "Technology: $($DT.Technology)"
                    $DTListData += "Admin Comments: $($DT.Description.'#text')"
                    $DTListData += "Languages: $($($DT.Languages.lang) -join ',')"
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath

                    $DTListData = @()
                    $DTSection = "Content"
                    $InstallContent = $DT.Installer.Contents.Content | Where-Object {$_.ContentID -eq $DT.Installer.CustomData.InstallContent.ContentId}
                    $DTListData += "Content Location: $($InstallContent.Location)"
                    switch ($dt.Installer.CustomData.UninstallSetting) {
                        'Different' { $UninstallContentSetting = 'Different from install content' }
                        'SameAsInstall' { $UninstallContentSetting = 'Same as install content'}
                        'NoneRequired' { $UninstallContentSetting = 'No uninstall content'}
                        Default { $UninstallContentSetting = 'Same as install content' }
                    }
                    $DTListData += "Uninstall content settings: $UninstallContentSetting"
                    If ($UninstallContentSetting -eq 'Different from install content' ){
                        $UnInstallContent = $dt.Installer.Contents.Content | Where-Object {$_.ContentID -eq $dt.Installer.CustomData.UninstallContent.ContentId}
                        $DTListData += "Uninstall content location: $($UnInstallContent.Location)"
                    } else{
                        $DTListData += "Uninstall content location: N/A"
                    }
                    If ($InstallContent.PinOnClient -eq 'true'){
                        $DTListData += "Persist in the client cache: True"
                    } else {
                        $DTListData += "Persist in the client cache: False"
                    }
                    If ($InstallContent.FallbackToUnprotectedDP -eq 'true'){
                        $DTListData += "Allow clients to use DP from the default site: True"
                    } else {
                        $DTListData += "Allow clients to use DP from the default site: False"
                    }
                    switch ($InstallContent.OnSlowNetwork) {
                        'DoNothing' { $DTListData += "Deployment options for neighbor or default boundary group: Do not download content" }
                        'Download' { $DTListData += "Deployment options for neighbor or default boundary group: Download content from DP and run locally" }
                        Default { $DTListData += "Deployment options for neighbor or default boundary group: Download content from DP and run locally" }
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath

                    $DTListData = @()
                    $DTSection = "Programs"
                    $DTListData += "Install program: $(($DT.Installer.InstallAction.Args.arg | Where-Object {$_.name -eq 'InstallCommandLine'}).'#text')"
                    #$DTListData += "Content ID: $($DT.Installer.InstallAction.Contents.Content.ContentId)"
                    $DTListData += "Uninstall program: $(($DT.Installer.UninstallAction.Args.arg | Where-Object {$_.name -eq 'InstallCommandLine'}).'#text')"
                    #$DTListData += "Content ID: $($DT.Installer.UninstallAction.Contents.Content.ContentId)"
                    $DTListData += "Repair program: $(($DT.Installer.RepairAction.Args.arg | Where-Object {$_.name -eq 'InstallCommandLine'}).'#text')"
                    #$DTListData += "Content ID: $($DT.Installer.RepairAction.Contents.Content.ContentId)"
                    $DTListData += "Run installation and uninstall program as 32-bit process on 64-bit clients:  $(($DT.Installer.InstallAction.Args.arg | Where-Object {$_.name -eq 'RunAs32Bit'}).'#text')"
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath
                    
                    $DTListData = @()
                    $DTSection = "Detection Method"
                    If ($DT.Installer.DetectAction.Provider -like 'MSI'){
                        #MSI
                        $ProductCode = ($DT.Installer.DetectAction.args.arg | Where-Object {$_.Name -eq 'ProductCode'}).'#text'
                        $DTListData += "MSI detection method."
                        $DTListData += "MSI Product Code: $ProductCode"
                    } elseif ($DT.Installer.DetectAction.Provider -like 'Local') {
                        #Enhanced Detection Method
                        $DTListData += "Using enhanced detection method."
                        #$EDM = [xml]($DT.Installer.DetectAction.Args.arg | Where-Object {$_.Name -eq 'MethodBody'}).'#text'
                    } else {
                        #Unknown
                        $DTListData += "Other detection method."
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath

                    $DTListData = @()
                    $DTSection = "User Experience"
                    switch (($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'ExecutionContext'}).'#text') {
                        'System' { $InstallBehavior = 'Install for system' }
                        'User' { $InstallBehavior = 'Install for user' }
                        'Any' { $InstallBehavior = 'Install for system if resource is device; otherwise install for user' }
                    }
                    $DTListData += "Installation Behavior: $InstallBehavior"
                    switch (($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'RequiresLogOn'}).'#text') {
                        'True' { $LogonReq = 'Only when a user is logged on' }
                        'False' { $LogonReq = 'Only when no user is logged on' }
                        $null { $LogonReq = 'Whether or not a user is logged on' }
                        default { $LogonReq = 'Whether or not a user is logged on' }
                    }
                    $DTListData += "Logon requirement: $LogonReq"
                    switch (($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'UserInteractionMode'}).'#text') {
                        'Normal' { $AppVisibility = 'Normal' }
                        'Maximized' { $AppVisibility = 'Maximized' }
                        'Minimized' { $AppVisibility = 'Minimized' }
                        'Hidden' { $AppVisibility = 'Hidden' }
                        Default { $AppVisibility = 'Normal'}
                    }
                    $DTListData += "Installation program visibility: $AppVisibility"
                    if (($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'RequiresUserInteraction'}).'#text' -eq 'true'){
                        $DTListData += "Allow user to view and interact with the program Installation: true"
                    } else {
                        $DTListData += "Allow user to view and interact with the program Installation: false"
                    }
                    $MaxRunTime = ($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'MaxExecuteTime'}).'#text'
                    $DTListData += "Maximum allowed run time (minutes): $MaxRunTime"
                    $EstimatedRunTime = ($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'ExecuteTime'}).'#text'
                    $DTListData += "Estimated installation time (minutes): $EstimatedRunTime"
                    switch (($DT.Installer.InstallAction.args.arg | Where-Object {$_.Name -eq 'PostInstallBehavior'}).'#text') {
                        'BasedOnExitCode' { $CompletionBehavior = 'Determine behavior based on return codes' }
                        'NoAction' { $CompletionBehavior = 'No specific action' }
                        'ProgramReboot' { $CompletionBehavior = 'The Software install program might force a device restart' }
                        'ForceReboot' { $CompletionBehavior = 'Configuration Manager client will force a mandatory device restart' }
                        Default {}
                    }
                    $DTListData += "Installation Completion Behaviour: $CompletionBehavior"
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath

                    $DTListData = @()
                    $DTSection = "Requirements"
                    If([string]::IsNullOrEmpty($DT.Requirements)){
                        $DTListData += "No requirements defined."
                    } else {
                        #Any number of requirements
                        $DTListData += "One or more requirements defined."
                        If (!([string]::IsNullOrEmpty($DT.Requirements.Rule.OperatingSystemExpression))) {
                            $Operator = $DT.Requirements.Rule.OperatingSystemExpression.Operator
                            $OSes = $DT.Requirements.Rule.OperatingSystemExpression.Operands.RuleExpression.RuleId -join '; '
                            $DTListData += "Operating system: &#60;$Operator&#62; $OSes"
                        }
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath
                    
                    $DTListData = @()
                    $DTSection = "Return Codes"
                    foreach ($XC in $DT.Installer.CustomData.ExitCodes.ExitCode) {
                        $DTListData += "$($XC.Class) ($($XC.Code))"
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath

                    $DTListData = @()
                    $DTSection = "Dependencies"
                    If([string]::IsNullOrEmpty($DT.Dependencies)){
                        $DTListData += "No dependencies defined."
                    } else {
                        #Any number of dependencies
                        $DTListData += "One or more dependencies defined."
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath
                    
                    $DTListData = @()
                    $DTSection = "Install behavior - Executables that must be closed"
                    If([string]::IsNullOrEmpty($DT.Installer.CustomData.InstallProcessDetection)){
                        $DTListData += "No executables defined."
                    } else {
                        #Any number of dependencies
                        foreach ($exe in $DT.Installer.CustomData.InstallProcessDetection.ProcessList.ProcessInformation){
                            $DTListData += "$($exe.DisplayInfo.info.DisplayName) ($($exe.Name))"
                        }
                    }
                    Write-HtmlList -InputObject $DTListData -Description $DTSection -Level 6 -File $FilePath
                }
            }
            else {
                Write-HTMLParagraph -Text 'There are no Deployment Types configured for this Application.' -Level 5 -File $FilePath
            }
            #endregion deployment types

            #region deployments
            Write-ProgressEx -CurrentOperation "Application: $($App.LocalizedDisplayName)" -Activity 'Configuration Manager Application' -Status 'Application deployments' -Id 10
            $AppDeployments = Get-CMApplicationDeployment -ApplicationID $($App.CI_ID)
            Write-HTMLHeading -Level 5 -Text "Deployments for $($App.LocalizedDisplayName):" -File $FilePath
            if (-not [string]::IsNullOrEmpty($AppDeployments)) {
                $DeploymentsArray = @()
                foreach ($AppDeployment in $AppDeployments){
                    $AppDeploySettings = @()
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
                        True{$TimeZone = 'UTC'}
                        False{$TimeZone = 'Client Local Time'}
                    }
                    $AppDeploySettings += '--B--User Experience--/B--'
                    $AppDeploySettings += "--BULLET-- User notifications: $UserNotice"
                    If($AppDeployment.OfferFlags -band 0x00000020){ #Bit 32 = HIGHIMPACTDEPLOYMENT
                        $AppDeploySettings += '--TAB----CBOX-- When software changes are required, show a dialog window to the user instead of a toast notification'
                    } else {
                        $AppDeploySettings += '--TAB----UNCBOX-- When software changes are required, show a dialog window to the user instead of a toast notification'
                    }
                    $AppDeploySettings += '--BULLET-- Allow the following activies outside of maintinance windows:'
                    if ($AppDeployment.OverrideServiceWindows -eq 'True'){
                        $AppDeploySettings += '--TAB----CBOX-- Software Installation'
                    } else {
                        $AppDeploySettings += '--TAB----UNCBOX-- Software Installation'
                    }
                    if ($AppDeployment.RebootOutsideOfServiceWindows -eq 'True'){
                        $AppDeploySettings += '--TAB----CBOX-- System Restart'
                    } else {
                        $AppDeploySettings += '--TAB----UNCBOX-- System Restart'
                    }
                    $AppDeploySettings += '--CRLF----B--Deployment Settings--/B--'
                    If($AppDeployment.OfferFlags -band 0x00000008){ #Bit 8 = ALLOWUSERSTOREPAIRAPP
                        $AppDeploySettings += '--CBOX-- Allow end users to attempt to repair this application'
                    } else {
                        $AppDeploySettings += '--UNCBOX-- Allow end users to attempt to repair this application'
                    }
                    If($AppDeployment.OfferFlags -band 0x00000001){ #Bit 1 = PREDEPLOY
                        $AppDeploySettings += '--CBOX-- Pre-deploy software to the users primary device'
                    } else {
                        $AppDeploySettings += '--UNCBOX-- Pre-deploy software to the users primary device'
                    }
                    If($AppDeployment.RequireApproval -eq 'True'){
                        $AppDeploySettings += '--CBOX-- An administrator must approve a request for this application on the device'
                        $EmailSubscriptions = Get-CMAlertSubscription | Where-Object {(Get-CMAlert -TypeInstanceId "$($AppDeployment.AssignmentUniqueID)").ID -in $_.AlertIds}
                        If ($EmailSubscriptions){
                            $AppDeploySettings += "--CBOX-- Email Approvals: $($EmailSubscriptions.EmailAddress)"
                        } else {
                            $AppDeploySettings += "--UNCBOX-- Email Approvals"
                        }
                    } else {
                        $AppDeploySettings += '--UNCBOX-- An administrator must approve a request for this application on the device'
                    }
                    If($AppDeployment.WoLEnabled -eq 'True'){
                        $AppDeploySettings += '--CBOX-- Send wake-up packets'
                    } else {
                        $AppDeploySettings += '--UNCBOX-- Send wake-up packets'
                    }
                    If($AppDeployment.OfferFlags -band 0x00000004){ #Bit 4 = ENABLEPROCESSTERMINATION
                        $AppDeploySettings += '--CBOX-- Automatically close any running executables you specified in Install Behavior'
                    } else {
                        $AppDeploySettings += '--UNCBOX-- Automatically close any running executables you specified in Install Behavior'
                    }
                    If($AppDeployment.OfferFlags -band 0x00000002){ #Bit 2 = ONDEMAND
                        #'Use default distribution point group associated to this collection'
                    }
                    If($AppDeployment.OfferFlags -band 0x00000010){ #Bit 16 = RELATIVESCHEDULE
                        #'RELATIVESCHEDULE?'
                    }
                    if ($AppDeployment.SoftDeadlineEnabled -eq 'True'){
                        $AppDeploySettings += '--CRLF----CBOX-- Delay enforcement of this deployment according to user preferences, up to grace period..'
                    } else {
                        $AppDeploySettings += '--CRLF----UNCBOX-- Delay enforcement of this deployment according to user preferences, up to grace period..'
                    }
                    $AppDeploySettingsString = $AppDeploySettings -join '--CRLF--'
                    $DeploymentsArray += New-Object -TypeName psobject -Property @{'Collection'="$($AppDeployment.CollectionName)";'Action'="$Action";'Purpose'="$Purpose";'Deployment Details'="$AppDeploySettingsString";'Available Time'="$($AppDeployment.StartTime)";'Deadline'="$($AppDeployment.EnforcementDeadline)";'Time Zone'="$TimeZone";'Deployment ID'="$($AppDeployment.AssignmentUniqueID)"}
                }
                $DeploymentsArray = $DeploymentsArray | Select-Object 'Collection','Action','Purpose','Deployment Details','Available Time','Deadline','Time Zone','Deployment ID'
                Write-HtmlTable -InputObject $DeploymentsArray -Border 1 -Level 6 -File $FilePath
            }else{
                Write-HTMLParagraph -Text 'There are no deployments for this application.' -Level 6 -File $FilePath
            }
            #endregion deployments

            #region phaseddeployments
            Write-ProgressEx -CurrentOperation "Application: $($App.LocalizedDisplayName)" -Activity 'Configuration Manager Application' -Status 'Application phased deployments' -Id 10
            Write-HTMLHeading -Level 5 -Text "Phased Deployments for $($App.LocalizedDisplayName):" -File $FilePath
            $PDs = @()
            $PDs = Get-PWCMPhasedDeployment -DeploymentObjectID $($App.CI_ID) -SiteServer $SMSProvider -SiteName $SiteCode
            if (-not [string]::IsNullOrEmpty($PDs)) {
                foreach ($PD in $PDs){
                    Write-HTMLHeading -Text $PD.Name -Level 6 -ExcludeTOC -File $FilePath
                    Write-HTMLParagraph -Text "Description: $($PD.Description)" -Level 6 -File $FilePath
                    $AppPhases = @()
                    $AppPhases = Get-PWCMPDPhases -PhasedDeploymentID $PD.PhasedDeploymentID
                    $AppPhasesArray = @()
                    foreach ($Phase in $AppPhases){
                        If ($Phase.SuccessCriteria.Number){
                            $SuccessCriteria = "$($Phase.SuccessCriteria.Number) Computers"
                        }
                        If ($Phase.SuccessCriteria.Compliance){
                            $SuccessCriteria = "$($Phase.SuccessCriteria.Compliance) Percent"
                        }
                        $AppPhasesArray += New-Object -TypeName psobject -Property @{'Order'=$Phase.Order;'Name'="$($Phase.Name)";'Target'="$($Phase.Target)";'Manual Start'="$($Phase.ManuallyStartPhase)";'Start After Previous Phase'="$($Phase.DaysAfterPrevPhaseSuccess) Days";'Gradually Available Over'="$($Phase.GraduallyAvailable) Days";'SuccessCriteria'="$SuccessCriteria"}
                        Remove-Variable SuccessCriteria -ErrorAction Ignore
                    }
                    $AppPhasesArray = $AppPhasesArray |Select-Object 'Order','Name','Target','Manual Start','Start After Previous Phase','Gradually Available Over','SuccessCriteria'
                    Write-HtmlTable -InputObject $AppPhasesArray -Border 1 -Level 6 -File $FilePath
                    Remove-Variable AppPhasesArray -ErrorAction Ignore
                    Remove-Variable AppPhases -ErrorAction Ignore
                }
            }else{
                Write-HTMLParagraph -Text 'There are no phased deployments for this application.' -Level 6 -File $FilePath
            }
            #endregion phaseddeployments

        }
    }else{
        Write-HTMLParagraph -Text 'There are no Applications configured in this site.' -Level 4 -File $FilePath
    }
    Write-ProgressEx -CurrentOperation "Completed applications" -Activity 'Configuration Manager Application' -Status 'Application details' -Id 10 -Completed
}
elseif (-not [string]::IsNullOrEmpty($Applications)) {
    Write-HTMLParagraph -Text "There are $($Applications.count) applications configured." -Level 4 -File $FilePath
    $AppBasics = @()
    foreach ($App in $Applications){
        $AppBasics += New-Object -TypeName PSObject -Property @{'Name'="$($App.LocalizedDisplayName)"; 'Deployed' = $($App.IsDeployed); 'Created by' = $($App.CreatedBy); 'Date created'=$($App.DateCreated)}
    }
    $AppBasics = $AppBasics | Select-Object 'Name','Deployed','Created by','Date Created'
    Write-HtmlTable -InputObject $AppBasics -Border 1 -Level 4 -File $FilePath
}
else {
    Write-HTMLParagraph -Text 'There are no Applications configured in this site.' -Level 4 -File $FilePath
}
Write-ProgressEx -CurrentOperation "Applications Complete"
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Applications

#region Packages
Write-ProgressEx -CurrentOperation "Processing Packages"
Write-HTMLHeading -PageBreak -Level 3 -Text 'Packages' -File $FilePath
$Packages = Get-CMPackage
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($Packages)){
        Write-HTMLParagraph -Text "There are $($Packages.count) packages configured." -Level 5 -File $FilePath
        Write-HTMLParagraph -Text "Below is a summary of all $($Packages.count) packages defined in this site. These are applications that are installed using traditional packages." -Level 3 -File $FilePath
        foreach ($Package in $Packages) {
            Write-Debug "$(Get-Date):   Found Package: $($Package.Name)"
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
            Write-Debug "$(Get-Date):   Getting programs in package $($Package.Name)..."
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
                Write-Debug "$(Get-Date):   No programs found in package $($Package.Name)..."
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
    $PackageBasics = $PackageBasics | Select-Object 'Name','Programs','Content Date'
    Write-HtmlTable -InputObject $PackageBasics -Border 1 -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text 'There are no packages configured in this site.' -Level 5 -File $FilePath
}
Write-ProgressEx -CurrentOperation "Completed processing Packages"
#endregion Packages
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Application Management


#region Software Updates
Write-ProgressEx -CurrentOperation 'Software Updates' 
Write-HTMLHeading -Level 2 -PageBreak -Text 'Software Updates' -File $FilePath

#region Update Groups
Write-ProgressEx -CurrentOperation 'Software Update groups' 
Write-HTMLHeading -Level 3 -PageBreak -Text 'Software Update Groups' -File $FilePath
$UpdateGroups = Get-CMSoftwareUpdateGroup
If(-not [string]::IsNullOrEmpty($UpdateGroups)){
    Write-HTMLParagraph -Text "There are $($UpdateGroups.count) update groups defined in this site." -Level 3 -File $FilePath
    $UGs = $UpdateGroups|Sort-Object LocalizedDisplayName|Select-Object @{Name='Group Name';expression={$_.LocalizedDisplayName}},@{Name='ID';expression={$_.CI_ID}},@{Name='Update Count';expression={$_.NumberOfUpdates}},@{Name='Expired Updates';expression={$_.NumberOfExpiredUpdates}},@{Name='Created By';expression={$_.CreatedBy}},@{Name='Date Created';expression={$_.DateCreated}},@{Name='Deployed';expression={$_.IsDeployed}},@{Name='Compliance';expression={"$($_.PercentCompliant)%"}}
    Write-HtmlTable -InputObject $UGs -Border 1 -Level 3 -File $FilePath
}else{
    Write-HTMLParagraph -Text "There are no update groups defined in this site." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation 'Completed Software Update groups'
#endregion Update Groups

#region Update Packages
Write-ProgressEx -CurrentOperation 'Software Update packages'
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
    $SUPackages= $SUPackages|Sort-Object Name|Select-Object 'Name','Package ID','BDR','Source Path','Source Status'
    Write-HtmlTable -InputObject $SUPackages -Border 1 -Level 3 -File $FilePath
    Write-HTMLParagraph -Text '(BDR = Binary Differential Replication)' -Level 4 -File $FilePath
}else{
    Write-HTMLParagraph -Text "There are no update packages defined in this site." -Level 3 -File $FilePath
}
Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation 'Completed Software Update packages'
#endregion Update Packages


#region ADRs
Write-ProgressEx -CurrentOperation 'Processing Automatic Deployment Rules'
Write-HTMLHeading -Level 3 -PageBreak -Text 'Automatic Deployment Rules' -File $FilePath
$CMPSSuppressFastNotUsedCheck = $true
$ADRs=Get-CMSoftwareUpdateAutoDeploymentRule
foreach ($ADR in $ADRs){
    Write-ProgressEx -CurrentOperation "Processing ADR $($ADR.Name)" -Activity 'Automatic Deployment Rules' -Status 'Collecting details' -Id 11
    $ADRListDetails = @()
    #$ADRListTitle = "Name: $($ADR.Name)"
    Write-HTMLHeading -Level 4 -Text "$($ADR.Name)" -File $FilePath
    $ADRListDescription = $ADR.Description
    Remove-Variable languages -ErrorAction SilentlyContinue
    foreach ($locale in ([xml]$adr.ContentTemplate).ContentActionXML.ContentLocales.Locale){
        if ($locale -ne 'Locale:0'){$languages = "$languages, $((Get-CMCategory -Id $locale).LocalizedCategoryInstanceName)"}
    }
    If ($Null -ne $languages){
        $ADRListDetails += "Languages: $($languages.Trim(', '))"
    } else {
        $ADRListDetails += "Languages: N/A"
    }
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
                If ($Schedule.OffsetDay){
                    $ADRListDetails += "Evaluation Schedule: Occurs $($Schedule.OffsetDay) days after the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                }else{
                    $ADRListDetails += "Evaluation Schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                }
            }
        }
    }else{
        $ADRListDetails += "Evaluation Schedule: No schedule defined"
    }
    [xml]$rules=$ADR.UpdateRuleXML
    Remove-Variable Categories -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Web
    $UpdateRuleList = @()
    Write-Debug "$(Get-Date):   Getting a list of configured update rules for this ADR: $($ADR.Name)"
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
                    #$UPClasses = @(Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -Query "SELECT LocalizedCategoryInstanceName FROM SMS_CategoryInstance WHERE CategoryTypeName=`'UpdateClassification`' and CategoryInstance_UniqueID=`'$UC`'" -ComputerName $SMSProvider)
                    $UPClass = Get-CMSoftwareUpdateCategory -UniqueId $UC -Fast
                    $UpdateClassification += $UPClass.LocalizedCategoryInstanceName
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
                #$Vendors = @()
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
    Remove-Variable Package -ErrorAction Ignore
    If ($null -ne ([xml]$adr.ContentTemplate).ContentActionXML.PackageID){
        $DeploymentPackage = Get-CMSoftwareUpdateDeploymentPackage -Id ([xml]$adr.ContentTemplate).ContentActionXML.PackageID
        $Package = New-Object -TypeName PSObject -Property @{'Package Name'="$($DeploymentPackage.Name)";'Description'="$($DeploymentPackage.Description)";'Package ID'="$($DeploymentPackage.PackageID)";'Source Location'="$($DeploymentPackage.PkgSourcePath)"}
        $Package = $package | Select-Object 'Package Name','Package ID','Description','Source Location'
    } else {
        $UpdateRuleList += 'No software update package defined in this ADR.'
    }
    Write-HtmlList -InputObject $ADRListDetails -Description $ADRListDescription -Level 3 -File $FilePath
    #$UpdateRuleList
    Write-HtmlList -InputObject $UpdateRuleList -Description 'Software Update Property Filters (Update Rules):' -Level 3 -File $FilePath
    If ($Package){
        Write-HtmlTable -InputObject $Package -Level 5 -File $FilePath
    }
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
        Write-HtmlTable -InputObject ($ADRDeployments|Select-Object @{Name='Collection';expression={$_.CollectionName}},Enabled) -Level 5 -File $FilePath
    }
}
Write-HtmliLink -ReturnTOC -File $FilePath
Write-ProgressEx -CurrentOperation 'Completed Processing ADRs' -Activity 'Automatic Deployment Rules' -Status 'Collecting details' -Id 11 -Completed
Write-ProgressEx -CurrentOperation 'Completed processing of ADRs'
#endregion ADRs


#endregion Software Updates


#region Operating Systems
Write-ProgressEx -CurrentOperation 'Operating Systems'
Write-HTMLHeading -Level 2 -PageBreak -Text 'Operating Systems' -File $FilePath

#region Driver Packages
Write-ProgressEx -CurrentOperation 'Processing Driver Packages'
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
            If (-not [string]::IsNullOrEmpty($PackageDescription)){
                Write-HtmlList -Title $PackageName -Description $PackageDescription -InputObject $DPackArray -Level 4 -File $FilePath
            }else{
                Write-HtmlList -Title $PackageName -InputObject $DPackArray -Level 4 -File $FilePath
            }
            Write-Debug "$(Get-Date):   Getting Drivers in Package: $PackageName"
            $Drivers = @(Get-CMDriver -DriverPackageId "$($DriverPackage.PackageID)")
            if ($Drivers.Count -gt 0){
                Write-Debug "$(Get-Date):   Drivers found in package. Processing the drivers"
                $DriverArray = @()
                foreach ($Driver in $Drivers){
                    if (Test-Path "filesystem::$($Driver.ContentSourcePath)" -ErrorAction SilentlyContinue){
                        $Verified = "Path Verified"
                    }else{
                        $Verified = "Path not found"
                    }
                    $DriverArray += New-Object -TypeName psobject -Property @{'Driver Name'="$($Driver.LocalizedDisplayName)";'Manufacturer'="$($Driver.DriverProvider)";'Driver Version'="$($Driver.DriverVersion)";'Source Path'="$($Driver.ContentSourcePath)";'Source Status' = "$Verified";'INF File'="$($Driver.DriverINFFile)"}
                }
                $StartBase = $DriverArray[0].'source path'
                $BaseLenght = $StartBase.length
                $ArrayLenght = $DriverArray.Count
                $counter = 0
                Do{
                    $RelBase = $StartBase.Substring(0,$BaseLenght-$counter)
                    #$RelBase
                    $newcount = ($DriverArray.'source path'|Select-String -SimpleMatch "$RelBase").count
                    $counter++
                }While (($newcount -lt $ArrayLenght) -and ($counter -lt $BaseLenght))
                if ($StartBase.length -gt 15){
                    foreach ($Drvr in $DriverArray){
                        $SP = ($drvr.'source path').replace("$RelBase","***.\")
                        $drvr.'source path' = "$SP"
                    }
                }
                if (-not [string]::IsNullOrEmpty($DriverPackages)){
                    $DriverArray = $DriverArray | Select-Object 'Driver Name','Manufacturer','Driver Version','Source Path','Source Status','INF File'
                    if (-not [string]::IsNullOrEmpty($DriverArray)){
                        If (-not [string]::IsNullOrEmpty($RelBase)){
                            Write-HTMLParagraph -Text "Paths that start with ***. have drivers located in this root path:<br />$RelBase" -Level 5 -File $FilePath
                        }
                        Write-HtmlTable -InputObject $DriverArray -Border 1 -Level 5 -File $FilePath
                    }
                }
            }else{
                Write-HTMLParagraph -Text 'No drivers were found in the package.' -Level 5 -File $FilePath
            }
        }
    }else{
        Write-HTMLParagraph -Text 'There are no Driver Packages configured in this site.' -Level 4 -File $FilePath
    }
}else{
    if (-not [string]::IsNullOrEmpty($DriverPackages)){
        Write-HTMLParagraph -Text "There are $($DriverPackages.count) Driver Packages configured." -Level 4 -File $FilePath
        if (-not [string]::IsNullOrEmpty($DriverPackages)){
            Write-HtmlList -InputObject ($DriverPackages.Name) -Level 4 -File $FilePath
        }
    }else{
        Write-HTMLParagraph -Text 'There are no Driver Packages configured in this site.' -Level 4 -File $FilePath
    }
}
Write-ProgressEx -CurrentOperation 'Completed processing Driver Packages'
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Driver Packages

#region Operating System Upgrade Packages
Write-ProgressEx -CurrentOperation 'Processing Operating System Upgrade Packages'
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
            $UgPackList += "Version: $($OSUpgradePack.Version)"
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
Write-ProgressEx -CurrentOperation 'Completed Operating System Upgrade Packages'
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Operating System Upgrade Packages

#region Operating System Images
Write-ProgressEx -CurrentOperation 'Processing Operating System Images'
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
Write-ProgressEx -CurrentOperation 'Completed Operating System Images'
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Operating System Images

#region Boot Images
Write-ProgressEx -CurrentOperation 'Processing Boot Images'
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
                $DriverArray += New-Object -TypeName psobject -Property @{'Driver Name'="$($ImportedDriver.LocalizedDisplayName)";'Driver Class'="$($ImportedDriver.DriverClass)";'Inf File'="$($ImportedDriver.DriverINFFile)";'Driver Version'="$($ImportedDriver.DriverVersion)"}
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
Remove-Variable BootImage -ErrorAction Ignore
Write-ProgressEx -CurrentOperation 'Completed Boot Images'
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Boot Images


#region Task Sequences
Write-ProgressEx -CurrentOperation 'Processing Task Sequences'
Write-HTMLHeading -Level 3 -PageBreak -Text 'Task Sequences' -File $FilePath
$TaskSequences = Get-CMTaskSequence
Write-Debug "$(Get-Date):   working on $($TaskSequences.count) Task Sequences"
if ($ListAllInformation){
    if (-not [string]::IsNullOrEmpty($TaskSequences)){
        foreach ($TaskSequence in $TaskSequences){
                Write-ProgressEx -CurrentOperation "Detailing $($TaskSequence.Name) Task Sequences" -Activity 'Processing Task Seqeunce' -Status 'Collecting Steps' -Id 12
                Write-HTMLHeading -Level 4 -Text "$($TaskSequence.Name)" -File $FilePath
                $TSDetails = @()
                $TSDetails += "Package ID: $($TaskSequence.PackageID)"
                $TSBootImage = $TaskSequence.BootImageID
                If([string]::IsNullOrEmpty($TSBootImage)){
                    $BootImage="None"
                }else{
                    $BootImage = (Get-CMBootImage -id $TSBootImage -ErrorAction Ignore).Name
                }
                Write-Debug "$(Get-Date):   Task Sequence Boot Image: $BootImage"
                $TSDetails += "Task Sequence Boot Image: $BootImage"
                $TSRefs = $TaskSequence.References.Package
                $OSImages = Get-CMOperatingSystemImage
                If([string]::IsNullOrEmpty($TSRefs) -or [string]::IsNullOrEmpty($OSImages)){
                    $TSOSImage="None"
                }else{
                    foreach ($Ref in $TSRefs){
                        If($Ref -in $OSImages.PackageID){
                            Write-Verbose "$(Get-Date):   (Get-CMOperatingSystemImage -id $($Ref)).Name"
                            $TSOSImage = (Get-CMOperatingSystemImage -id $Ref).Name
                        }
                    }
                }
                If([string]::IsNullOrEmpty($TSOSImage)){$TSOSImage="None"}
                Write-Debug "$(Get-Date):   Task Sequence Operating System Image: $TSOSImage"
                $TSDetails += "Task Sequence Operating System Image: $TSOSImage"
                #Task Sequence References
                $TSReferences = Get-WmiObject -Namespace "Root\SMS\site_$($SiteCode)" -Query "SELECT * FROM SMS_TaskSequencePackageReference_Flat where PackageID = `'$($TaskSequence.PackageID)`'" -ComputerName $SMSProvider
                $TotalSize = 0
                $TotalTsRefs = 0
                $TotalTsRefs = $TSReferences.count
                $TSRefInfo = @()
                foreach($ref in $TSReferences){
                    switch ($ref.ObjectType){
                        0 {
                            $PackageType = "Package"
                        }
                        3 {
                            $PackageType = "Driver Package"
                        }
                        4 {
                            $PackageType = "Task Sequence"
                        }
                        257{
                            $PackageType = "OS Image"
                        }
                        258{
                            $PackageType = "Boot Image"
                        }
                        259{
                            $PackageType = "OS Upgrade Package"
                        }
                        512{
                            $PackageType = "Application"
                        }
                        default{
                            $PackageType = $ref.ObjectType
                        }
                    }
                    $SizeMB = [math]::Round($ref.SourceSize/1024)
                    $TotalSize = $TotalSize + $SizeMB
                    If ($SizeMB -eq 0){
                        $SizeMB = "<1 MB"
                    } else {
                        $SizeMB = "$SizeMB MB"
                    }
                    $TSRefInfo += New-Object -TypeName PSObject -Property @{'Name'="$($ref.ObjectName)";'Type'="$PackageType";'Package ID'="$($ref.RefPackageID)";'Size'="$SizeMB"}
                }
                $TSRefInfo = $TSRefInfo|Select-Object Type,Name,'Package ID',Size
                Write-HtmlList -InputObject $TSDetails -Level 4 -File $FilePath
                Write-HTMLHeading -Level 5 -Text "Task Sequence References" -File $FilePath
                Write-HtmlList -InputObject @("This task sequence references $TotalTsRefs packages.","The referenced packages total $TotalSize MB.") -Level 4 -File $FilePath
                If ($TotalTsRefs -ne 0){Write-HtmlTable -InputObject $TSRefInfo -Border 1 -Level 6 -File $FilePath}
                Write-HTMLHeading -Level 5 -Text "Task Sequence Steps" -File $FilePath
                $Sequence = $Null
                $Sequence = ([xml]$TaskSequence.Sequence).sequence
                $AllSteps = Process-TSSteps -Sequence $Sequence
                $c = 0
                foreach ($Step in $AllSteps){$c++;$Step|Add-Member -MemberType NoteProperty -Name 'Step' -Value $c}
                $AllSteps = $AllSteps |Select-Object 'Step','Group Name','Step Name','Description','Action','Status','Continue on Error','Conditions'
                If ($AllSteps.count -gt 0){
                    Write-HtmlTable -InputObject $AllSteps -Border 1 -Level 6 -File $FilePath
                }else{
                    Write-HtmlList -InputObject "This task sequence contains no steps" -Level 4 -File $FilePath
                }
                Remove-Variable TSOSImage -ErrorAction Ignore
                Remove-Variable BootImage -ErrorAction Ignore
            }
    }else{
        Write-HTMLParagraph -Level 3 -Text 'There are no Task Sequences present in this environment.' -File $FilePath
    }
}else{
    if (-not [string]::IsNullOrEmpty($TaskSequences)){
        Write-HTMLParagraph -Level 3 -Text 'The following Task Sequences are configured:' -File $FilePath
        $TSList =@()
        foreach ($TaskSequence in $TaskSequences){
            $OSImages = Get-CMOperatingSystemImage
            If([string]::IsNullOrEmpty($TSRefs) -or [string]::IsNullOrEmpty($OSImages)){
                $TSOSImage="None"
            }else{
                foreach ($Ref in $TSRefs){
                    If($Ref -in $OSImages.PackageID){
                        Write-Debug "$(Get-Date):   (Get-CMOperatingSystemImage -id $($Ref)).Name"
                        $TSOSImage = (Get-CMOperatingSystemImage -id $Ref).Name
                    }
                }
            }
            If([string]::IsNullOrEmpty($TSOSImage)){$TSOSImage="None"}
            Write-Debug "$(Get-Date):   Task Sequence Operating System Image: $TSOSImage"
            $TSBootImage = $TaskSequence.BootImageID
            If([string]::IsNullOrEmpty($TSBootImage)){
                $BootImage = "None"
            }else{
                $BootImage = (Get-CMBootImage -id $TSBootImage -ErrorAction Ignore).Name
            }
            Write-Debug "$(Get-Date):   Task Sequence Boot Image: $BootImage"
            $TSName = "$($TaskSequence.Name)"
            $TSList += New-Object -TypeName PSObject -Property @{'Name'="$TSName";'Operating System Image'="$TSOSImage";'Boot Image'="$BootImage"}
        }
        $TSList = $TSList | Select-Object 'Name','Operating System Image','Boot Image'
        Write-HtmlTable -InputObject $TSList -Level 3 -File $FilePath
    }else{
        Write-HTMLParagraph -Level 3 -Text 'There are no Task Sequences present in this environment.' -File $FilePath
    }
}
Write-ProgressEx -CurrentOperation "Task Sequence details complete" -Activity 'Processing Task Seqeunce' -Status 'Collecting Steps' -Id 12 -Completed
Write-ProgressEx -CurrentOperation "Completed Task Sequences"
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Task Sequences

#endregion Operating Systems

#region Windows 10 Servicing
Write-ProgressEx -CurrentOperation 'Windows 10 Servicing'
Write-HTMLHeading -Level 2 -PageBreak -Text 'Windows 10 Servicing' -File $FilePath

#region Servicing Plan
Write-ProgressEx -CurrentOperation 'Windows 10 Servicing Plans'
Write-HTMLHeading -Level 3 -Text 'Servicing Plans' -File $FilePath
$CMPSSuppressFastNotUsedCheck = $true
$ServicingPlans = Get-CMWindowsServicingPlan -WarningAction Ignore
if(-not [string]::IsNullOrEmpty($ServicingPlans)){
    Write-HTMLParagraph -Text 'Below are details on all servicing plans defined in this site.' -Level 3 -File $FilePath
    foreach ($ServicingPlan in $ServicingPlans){
        Write-HTMLHeading -Level 4 -Text "$($ServicingPlan.Name)" -File $FilePath
        Write-HTMLParagraph -Text "$($ServicingPlan.Description)" -Level 4 -File $FilePath
        $DTxml = ([xml]($ServicingPlan).DeploymentTemplate).DeploymentCreationActionXML
        $SPListDetails = @()
        $SPListDetails += "Use Wake-on-LAN to wake up clients for required deployments: $($DTxml.EnableWakeOnLan)"
        Switch ($($DTxml.StateMessageVerbosity)){
	        1 {$StateMessages = 'Only error messages'}
	        5 {$StateMessages = 'Only success and error messages'}
	        10 {$StateMessages = 'All messages'}
        }
        $SPListDetails += "Choose how much state detail you want clients to report back. Detail level: $StateMessages"
        Switch(([xml]($ServicingPlan).AutoDeploymentProperties).AutoDeploymentRule.NoEULAUpdates){
	        true{$EULAUpdates = "Automatically deploy only software updates found by this rule that do not include a license agreement, or for which the license agreement has already been approved."}
	        false{$EULAUpdates = "Automatically deploy all software updates found by this rule, and approve any license agreements."}
        }
        $SPListDetails += "Updates with License Agreements: $EULAUpdates"
        Write-HtmlList -Description 'Deployment Settings' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        $TargetCollection = (Get-CMCollection -id ($ServicingPlan.CollectionID)).Name
        $SPListDetails += "Target Collection: $TargetCollection"
        Write-HtmlList -Description 'Servicing Plan' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        Switch(([xml]($ServicingPlan).UpdateRuleXML).UpdateXML.OSBranch){
	        0{$DeployBranch="Semi-Annual Channel (Targeted)"}
	        1{$DeployBranch="Semi-Annual Channel"}
        }
        $WaitDays = (((([xml]($ServicingPlan).UpdateRuleXML).UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem| Where-Object {$_.PropertyName -eq 'AfterDays'}).Matchrules.string).Split(':'))[2]
        $SPListDetails += "Specify the Windows readiness state to which this Servicing Plan should apply: $DeployBranch"
        $SPListDetails += "How many days after Microsoft has published a new upgrade would you like to wait before deploying in your environment: $WaitDays"
        Write-HtmlList -Description 'Deployment Ring' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        Switch ($($DTxml.Utc)){
	        false{$timebase = 'Client local time'}
	        true{$timebase = 'UTC'}
        }
        $SPListDetails += "Time based on: $timebase"
        If ($DTxml.AvailableDeltaDuration -eq 0){
	        $SPListDetails += "Software available time: As soon as possible"
        }else{
	        $SPListDetails += "Software available time: $($DTxml.AvailableDeltaDuration) $($DTxml.AvailableDeltaDurationUnits)"
        }
        If ($DTxml.Duration -eq 0){
	        $SPListDetails += "Installation Deadline: As soon as possible"
        }else{
	        $SPListDetails += "Installation Deadline: $($DTxml.Duration) $($DTxml.DurationUnits)"
        }
        $SPListDetails += "Delay Enforcement of this deployment according to user preferences, up to the grace period defined in client settings: $($DTxml.SoftDeadlineEnabled)"
        Write-HtmlList -Description 'Deployment Schedule' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        Switch ($($DTxml.UserNotificationOption)){
	        'DisplayAll'{$UserNotification = 'Display in Software Center and show all notifications'}
	        'DisplaySoftwareCenterOnly'{$UserNotification = 'Display in Software Center, and only show nitifications for computer restarts'}
	        'HideAll'{$UserNotification = 'Hide in Software Center and all notifications'}
        }
        $SPListDetails += "User notifications: $UserNotification"
        Switch ($($DTxml.AllowInstallOutSW)){
	        false{$InstallOutMW = 'Do not allow'}
	        true{$InstallOutMW = 'Allow installations'}
        }
        $SPListDetails += "Deadline behavior for Software Update installation outside of maintenance windows: $InstallOutMW"
        Switch ($($DTxml.AllowRestart)){
	        false{$RestartOutMW = 'Do not allow'}
	        true{$RestartOutMW = 'Allow restarts'}
        }
        $SPListDetails += "Deadline behavior for System restarts outside of maintenance windows: $RestartOutMW"
        $SPListDetails += "Suppress reboots on servers if update requires reboot: $($DTxml.SuppressServers)"
        $SPListDetails += "Suppress reboots on workstations if update requires reboot: $($DTxml.SuppressWorkstations)"
        $SPListDetails += "Windows Embedded devices, Commit changes at deadline: $($DTxml.PersistOnWriteFilterDevices)"
        $SPListDetails += "If any update in this deployment requires a system restart, run updates deployment evaluation cycle after restart: $($DTxml.RequirePostRebootFullScan)"
        Write-HtmlList -Description 'User Experience' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        If($($DTxml.EnableAlert) -eq $false){
	        $SPListDetails += "Configuration Manager alerts.  Generate an alert when the following conditions are met: False"
        }else{
	        $SPListDetails += "Configuration Manager alerts.  Generate an alert when the following conditions are met: True<br />Client Compliance is below the following percent: $($DTxml.AlertThresholdPercentage)<br />Offset from the deadline: $($DTxml.AlertDuration) $($DTxml.AlertDurationUnits)"
        }
        $SPListDetails += "Disable Operations Manager alerts while software updates run: $($DTxml.DisableMomAlert)"
        $SPListDetails += "Generate Operations Manager alert when a software update installation fails: $($DTxml.GenerateMomAlert)"
        Write-HtmlList -Description 'Alerts' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        switch ($DTxml.UseRemoteDP){
	        false{$deploymentopt = 'Do not install software updates'}
	        true{$deploymentopt = 'Download software updates from distribution point and install'}
        }
        $SPListDetails += "Select deployment options to use when when client uses neighbor or default boundary group: $deploymentopt"
        switch ($DTxml.UseUnprotectedDP){
	        false{$deploymentopt2 = 'Do not install software updates'}
	        true{$deploymentopt2 = 'Download and install software updates from the distribution points in the site default boundary group'}
        }
        $SPListDetails += "When software updates are not available on any distribution point in current or neighbor boundary group, download from default boundary group: $deploymentopt2"
        $SPListDetails += "If software updates are not available on distribution point in current, neighbor or site boundary groups, download content from Microsoft Updates: $($DTxml.AllowWUMU)"
        $SPListDetails += "Allow clients on a metered Internet connection to download content after the installation deadline which might incur additional costs: $($DTxml.AllowUseMeteredNetwork)"
        Write-HtmlList -Description 'Download Settings' -InputObject $SPListDetails -Level 4 -File $FilePath
        $SPListDetails = @()
        If (-not [string]::IsNullOrEmpty($ServicingPlan.Schedule)){
	        $Schedule=Convert-CMSchedule $ServicingPlan.Schedule
	        if ($Schedule.DaySpan -gt 0){
		        $SPListDetails += "Evaluation Schedule: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
	        }
	        elseif ($Schedule.HourSpan -gt 0){
		        $SPListDetails += "Evaluation Schedule: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
	        }
	        elseif ($Schedule.MinuteSpan -gt 0){
		        $SPListDetails += "Evaluation Schedule: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
	        }
	        elseif ($Schedule.ForNumberOfWeeks){
		        $SPListDetails += "Evaluation Schedule: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
	        }
	        elseif ($Schedule.ForNumberOfMonths){
		        if ($Schedule.MonthDay -gt 0){
			        $SPListDetails += "Evaluation Schedule: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
		        }
		        elseif ($Schedule.MonthDay -eq 0){
			        $SPListDetails += "Evaluation Schedule: Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
		        }
		        elseif ($Schedule.WeekOrder -gt 0){
			        switch ($Schedule.WeekOrder){
				        0 {$order = 'last'}
				        1 {$order = 'first'}
				        2 {$order = 'second'}
				        3 {$order = 'third'}
				        4 {$order = 'fourth'}
			        }
			        $SPListDetails += "Evaluation Schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
		        }
	        }
        }elseif (([xml]$ServicingPlan.AutoDeploymentProperties).AutoDeploymentRule.AlignWithSyncSchedule -eq "true") {
	        $SPListDetails += "Evaluation Schedule: Run the rule after any software update point synchronization"
        }else{
	        $SPListDetails += "Evaluation Schedule: Do not run this rule automatically"
        }
        Write-HtmlList -Description 'Evaluation Schedule' -InputObject $SPListDetails -Level 4 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text 'No servicing plans defined in this site.' -Level 3 -File $FilePath
}

Write-ProgressEx -CurrentOperation 'Completed Windows 10 Servicing Plans'
#endregion Servicing Plan

#region Windows Update for Business Policies
Write-ProgressEx -CurrentOperation 'Windows Update for Business Policies'
Write-HTMLHeading -Level 3 -Text 'Windows Update for Business Policies' -File $FilePath
Write-HTMLParagraph -Text 'Below are details on all Windows Update for Business Policies defined in this site.' -Level 3 -File $FilePath
$WUBPolicies = Get-CMConfigurationPolicy | Where-Object {'SettingsAndPolicy:SMS_WindowsUpdateForBusinessConfigurationSettings' -in $_.CategoryInstance_UniqueIDs}
if(-not [string]::IsNullOrEmpty($WUBPolicies)){
    foreach($WUBPolicy in $WUBPolicies){
        Remove-Variable rules -ErrorAction Ignore
        $rules=([XML]($WUBPolicy).SDMPackageXML).DesiredConfigurationDigest.ConfigurationPolicy.rules.rule.Expression.Operands
        foreach($rule in $rules){
            Switch ($rule.SettingReference.SettingLogicalName){
                'WindowsUpdateForBusinessConfigurationSettings_BranchReadinessLevel'{
                    Switch($rule.ConstantValue.Value){
                        2{$BranchLevel = 'Branch readiness level: Windows Insider Build - Fast'}
                        4{$BranchLevel = 'Branch readiness level: Windows Insider Build - Slow'}
                        8{$BranchLevel = 'Branch readiness level: Release Windows Insider Build'}
                        16{$BranchLevel = 'Branch readiness level: Semi-Annual Channel (Targeted)'}
                        32{$BranchLevel = 'Branch readiness level: Semi-Annual Channel'}
                    }
                }
                'WindowsUpdateForBusinessConfigurationSettings_DeferFeatureUpdatesPeriodInDays'{$DFUDays = "Defer Feature Updates - Deferral Period (days): $($rule.ConstantValue.Value)"}
                'WindowsUpdateForBusinessConfigurationSettings_PauseFeatureUpdates'{
                    Switch($rule.ConstantValue.Value){
                        1{$PauseFU = "Pause Feature Updates Starting: True"}
                        0{$PauseFU = "Pause Feature Updates Starting: False"}
                    }
                }
                'WindowsUpdateForBusinessConfigurationSettings_PauseFeatureUpdatesStartTime'{$PauseFUDate = "Pause Feature Updates starting: $($rule.ConstantValue.Value) (only if enabled)"}
                'WindowsUpdateForBusinessConfigurationSettings_DeferQualityUpdatesPeriodInDays'{$DQUDays = "Defer Quality Updates - Deferral Period (days): $($rule.ConstantValue.Value)"}
                'WindowsUpdateForBusinessConfigurationSettings_PauseQualityUpdates'{
                    Switch($rule.ConstantValue.Value){
                        1{$PauseQU = "Pause Quality Updates Starting: True"}
                        0{$PauseQU = "Pause Quality Updates Starting: False"}
                    }
                }
                'WindowsUpdateForBusinessConfigurationSettings_PauseQualityUpdatesStartTime'{$PauseQUDate = "Pause Quality Updates starting: $($rule.ConstantValue.Value) (only if enabled)"}
                'WindowsUpdateForBusinessConfigurationSettings_ExcludeWUDriversInQualityUpdate'{
                    Switch($rule.ConstantValue.Value){
                        0{$WUDrivers = "Include drivers with Windows Updates: True"}
                        1{$WUDrivers = "Include drivers with Windows Updates: False"}
                    }
                }
                'WindowsUpdateForBusinessConfigurationSettings_AllowMUUpdateService'{
                    switch($rule.ConstantValue.Value){
                        0{$OtherProducts = "Install updates from other Microsoft Products: False"}
                        1{$OtherProducts = "Install updates from other Microsoft Products: True"}
                    }
                }
            }
        }
        Remove-Variable WUBList -ErrorAction Ignore
        $WUBList = @($BranchLevel,$DFUDays,$PauseFU,$(if(-not [string]::IsNullOrEmpty($PauseFUDate)){$PauseFUDate}),$DQUDays,$PauseQU,$(if(-not [string]::IsNullOrEmpty($PauseQUDate)){$PauseQUDate}),$WUDrivers,$OtherProducts)
        Write-HtmlList -InputObject $WUBList -Title "$($WUBPolicy.LocalizedDisplayName)" -Description "$($WUBPolicy.LocalizedDescription)" -Level 4 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text 'No Windows Update for Business Policies defined in this site.' -Level 3 -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Completed Windows Update for Business Policies'

#endregion Windows Update for Business Policies

Write-ProgressEx -CurrentOperation 'Completed Windows 10 Servicing'
#endregion Windows 10 Servicing

#region Scripts
Write-ProgressEx -CurrentOperation 'Collecting Configuration Manager Scripts'
Write-HTMLHeading -Level 2 -PageBreak -Text 'Configuration Manager Scripts' -File $FilePath
$ScriptFeature = Get-CMSiteFeature| Where-Object {$_.FeatureGuid -like '566F8720-F415-4E10-9A51-CDE682BA2B2E'}
if (-not [string]::IsNullOrEmpty($ScriptFeature)){
    If ($ScriptFeature.Status -eq 1){
        $CMScripts = Get-WmiObject -Namespace ROOT\SMS\site_$SiteCode -ComputerName $SMSProvider -Class SMS_Scripts
        Write-Debug "$(Get-Date):   Working on $($CMScripts.count) Scripts"
        if ([string]::IsNullOrEmpty($CMScripts)){
            Write-HTMLParagraph -Text "No Scripts are defined in this site." -Level 3 -File $FilePath
        }else{
            if ($ListAllInformation){
                Write-Debug "$(Get-Date):   Working on detailed script information"
                #Detailed Script Information
                foreach ($Script in $CMScripts){
                    #Exclude scripts with feature set to 1. These are system scripts like CMPivot.
                    if (($Script.Feature -ne 1)){
                        #Get details of each script here
                        $ScriptList = @()
                        $Script.Get()
                        Switch($script.ApprovalState){
                            0{$Approval = "Waiting for Approval"}
                            1{$Approval = "Declined"}
                            3{$Approval = "Approved"}
                            default{$Approval = "Unknown"}
                        }
                        $UpdateTime = [Management.ManagementDateTimeConverter]::ToDateTime($Script.LastUpdateTime)
                        Write-HTMLHeading -Level 3 -Text "$($Script.ScriptName)" -File $FilePath
                        #$ScriptTitle = "$($Script.ScriptName)"
                        $ScriptList += "Author: $($Script.Author)"
                        $ScriptList += "Approved By: $($Script.Approver)"
                        $ScriptList += "Approval State: $Approval"
                        $ScriptList += "Last Updated: $UpdateTime"
                        #Write-HtmlList -Title $ScriptTitle -InputObject $ScriptList -Level 3 -File $FilePath
                        Write-HtmlList -InputObject $ScriptList -Level 3 -File $FilePath
                        $ScriptText = ([System.Text.Encoding]::unicode.GetString([System.Convert]::FromBase64String($($Script.Script)))).Substring(1)
                        Write-HTMLParagraph -Text "<B>Script Contents:</B><pre style=`"margin-left:60px; background-color:#eeeeee;`">$ScriptText</pre>" -Level 4 -File $FilePath
                    }
                }
            }Else{
                $Scripts = @()
                foreach ($Script in $CMScripts){
                    Switch($script.ApprovalState){
                        0{$Approval = "Waiting for Approval"}
                        1{$Approval = "Declined"}
                        3{$Approval = "Approved"}
                        default{$Approval = "Unknown"}
                    }
                    $UpdateTime = [Management.ManagementDateTimeConverter]::ToDateTime($Script.LastUpdateTime)
                    if (($Script.Feature -ne 1)){
                        $Scripts += New-Object -TypeName PSObject -Property @{'Script Name'="$($Script.ScriptName)";'Author'="$($Script.Author)";'Approver'=$($Script.Approver);'Approval State'="$Approval";'Last Update Time' = "$UpdateTime"}
                    }
                }
                $Scripts = $Scripts | Select-Object 'Script Name','Author','Approver','Approval State','Last Update Time'
                if ($Scripts.Count -gt 0){
                    Write-HtmlTable -InputObject $Scripts -Border 1 -Level 3 -File $FilePath
                }else{
                    Write-HTMLParagraph -Text "No Non-System Scripts are defined in this site." -Level 3 -File $FilePath
                }
            }
        }
    }else{
        Write-HTMLParagraph -Text "Scripts feature, `"$($ScriptFeature.Name)`", not enabled in this site." -Level 3 -File $FilePath
    }
}else{
    Write-HTMLParagraph -Text "Scripts feature not found in this site. Scripts were introduced with release 1706." -Level 3 -File $FilePath
}
Write-ProgressEx -CurrentOperation 'Completed Configuration Manager Scripts'
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Scripts

#endregion Software Library
Write-ProgressEx -CurrentOperation 'Completed Software Library'

#region Statistics
Write-ProgressEx -CurrentOperation 'Writing final Script Statistics'
Write-HTMLHeading -Level 1 -PageBreak -Text 'Documentation Script Execution Details' -File $FilePath
Set-Location -Path "$StartingPath"
$ScriptEndTime = Get-date
$ExecTime = $ScriptEndTime - $ScriptStartTime
$ExecTimeString = $(if($ExecTime.Days -gt 0){"$($ExecTime.Days) day(s), "}) + $(if($ExecTime.Hours -gt 0){"$($ExecTime.Hours) hour(s), "}) + $(if($ExecTime.Minutes -gt 0){"$($ExecTime.Minutes) minute(s), and "}) + "$($ExecTime.Seconds) second(s)."
$ScriptDetailsList = @()
$ScriptDetailsList += "Execution Start Time: $($ScriptStartTime.ToString())"
$ScriptDetailsList += "Execution End Time: $($ScriptEndTime.ToString())"
$ScriptDetailsList += "Total Execution Time: $ExecTimeString"
$ScriptDetailsList += "Executed Script Version: $DocumenationScriptVersion"
Write-HtmlList -InputObject $ScriptDetailsList -Level 2 -Type UL -File $FilePath
Write-HtmliLink -ReturnTOC -File $FilePath
#endregion Statistics

Write-HTMLFooter -File $FilePath

Write-HTMLTOC -InputObject $Global:DocTOC -File $FilePath

Write-host "Completed execution at: $($ScriptEndTime.ToShortTimeString())."
Write-Host "Total execution time: $ExecTimeString"
