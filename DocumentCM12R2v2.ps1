#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Demonstrates script functionality using Microsoft Word, PDF, formatted text or HTML.
.DESCRIPTION
	Creates a sample report of various Word functionality using Microsoft Word, PDF, formatted text, HTML and PowerShell.
	Creates a document named Script_Template.docx (or .PDF or .TXT or .HTML).
	Word and PDF documents include a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	Look for the sections starting with ### to find the lines to either be replaced with your 
	script code or that need changing for your script needs
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010 and 2013 are supported.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2010 but Subtitle/Subject & Author fields need to be moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2010/2013. Doesn't work in 2013, works in 2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Sideline.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be ReportName_2014-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin or Local Administrator).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	This parameter is disabled by default.
.PARAMETER ComputerName
	Specifies a computer to use to run the script against.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual computer name.
	Default is localhost.
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -TEXT

	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -HTML

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\ScriptTemplate.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\ScriptTemplate.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2014 at 6PM is 2014-06-01_1800.
	Output filename will be Script_Template_2014-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -Hardware
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	localhost for running hardware inventory.
	localhost will be replaced by the actual computer name.
.EXAMPLE
	PS C:\PSScript > .\ScriptTemplate.ps1 -Hardware -ComputerName 192.168.1.51
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	192.168.1.51 for running hardware inventory.
	192.168.1.51 will be replaced by the actual computer name, if possible.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: DocumentCM12R2v2.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters, Barry Schiffer, David O'Brien
	LASTEDIT: April 06, 2015
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[Switch]$Software,

	[parameter(Mandatory=$False)] 
	[Switch]$ListAllInformation,

	[parameter(Mandatory=$False)] 
	[string]$SMSProvider='localhost'

	)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2014

#HTML functions and sample text contributed by Ken Avram October 2014
#Organized functions into logical units 16-Oct-2014
#Added regions 16-Oct-2014
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($PDF -eq $Null)
{
	$PDF = $False
}
If($Text -eq $Null)
{
	$Text = $False
}
If($MSWord -eq $Null)
{
	$MSWord = $False
}
If($HTML -eq $Null)
{
	$HTML = $False
}
If($AddDateTime -eq $Null)
{
	$AddDateTime = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}
If($ComputerName -eq $Null)
{
	$ComputerName = "LocalHost"
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:ComputerName))
{
	$ComputerName = "LocalHost"
}

If($MSWord -eq $Null)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Text -eq $Null)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($HTML -eq $Null)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}
#endregion

#region initialize variables for word html and text
If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 

	[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

If($HTML)
{
    Set htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set htmlbold        -Option AllScope -Value 1 4>$Null
    Set htmlitalics     -Option AllScope -Value 2 4>$Null
    Set htmlred         -Option AllScope -Value 4 4>$Null
    Set htmlcyan        -Option AllScope -Value 8 4>$Null
    Set htmlblue        -Option AllScope -Value 16 4>$Null
    Set htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set htmllightblue   -Option AllScope -Value 64 4>$Null
    Set htmlpurple      -Option AllScope -Value 128 4>$Null
    Set htmlyellow      -Option AllScope -Value 256 4>$Null
    Set htmllime        -Option AllScope -Value 512 4>$Null
    Set htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set htmlgray        -Option AllScope -Value 8192 4>$Null
    Set htmlolive       -Option AllScope -Value 16384 4>$Null
    Set htmlorange      -Option AllScope -Value 32768 4>$Null
    Set htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	$global:output = ""
}
#endregion

#region code for -hardware switch
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 4 0 "General Computer"
	}
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Results -ne $Null)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results
	}

	If($? -and $Results -ne $Null)
	{
		$Nics = $Results | Where {$_.ipaddress -ne $Null}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $ThisNic -ne $Null)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
	}

	$Results = $Null
	$ComputerItems = $Null
	$Drives = $Null
	$Processors = $Null
	$Nics = $Null
}

Function OutputComputerItem
{
	Param([object]$Item)
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{ Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{ Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{ Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$ItemInformation += @{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }
		$ItemInformation += @{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }
		$Table = AddWordTable -Hashtable $ItemInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t: " $Item.manufacturer
		Line 2 "Model`t`t: " $Item.model
		Line 2 "Domain`t`t: " $Item.domain
		Line 2 "Total Ram`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets): " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT): " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		FormatHTMLTable
		WriteHTMLLine 0 0 ""
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"}
		1	{$xDriveType = "No Root Directory"}
		2	{$xDriveType = "Removable Disk"}
		3	{$xDriveType = "Local Disk"}
		4	{$xDriveType = "Network Drive"}
		5	{$xDriveType = "Compact Disc"}
		6	{$xDriveType = "RAM Disk"}
		Default {$xDriveType = "Unknown"}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{ Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{ Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{ Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{ Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation += @{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		$DriveInformation += @{ Data = "Drive Type"; Value = $xDriveType; }
		$Table = AddWordTable -Hashtable $DriveInformation -Columns Data,Value -List -AutoFit $wdAutoFitContent;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		FormatHTMLTable
		WriteHTMLLine 0 0 ""
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{ Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{ Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }
		}
		$ProcessorInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $processor.name
		Line 2 "Description`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		FormatHTMLTable
		WriteHTMLLine 0 0 ""
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"}
		2	{$xAvailability = "Unknown"}
		3	{$xAvailability = "Running or Full Power"}
		4	{$xAvailability = "Warning"}
		5	{$xAvailability = "In Test"}
		6	{$xAvailability = "Not Applicable"}
		7	{$xAvailability = "Power Off"}
		8	{$xAvailability = "Off Line"}
		9	{$xAvailability = "Off Duty"}
		10	{$xAvailability = "Degraded"}
		11	{$xAvailability = "Not Installed"}
		12	{$xAvailability = "Install Error"}
		13	{$xAvailability = "Power Save - Unknown"}
		14	{$xAvailability = "Power Save - Low Power Mode"}
		15	{$xAvailability = "Power Save - Standby"}
		16	{$xAvailability = "Power Cycle"}
		17	{$xAvailability = "Power Save - Warning"}
		Default	{$xAvailability = "Unknown"}
	}

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"}
		Default	{$xTcpipNetbiosOptions = "Unknown"}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		If($ThisNic.Name -eq $nic.description)
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
		}
		Else
		{
			$NicInformation += @{ Data = "Name"; Value = $ThisNic.Name; }
			$NicInformation += @{ Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		$NicInformation += @{ Data = "Manufacturer"; Value = $Nic.manufacturer; }
		$NicInformation += @{ Data = "Availability"; Value = $xAvailability; }
		$NicInformation += @{ Data = "Physical Address"; Value = $Nic.macaddress; }
		$NicInformation += @{ Data = "IP Address"; Value = $xIPAddress[0]; }
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{ Data = ""; Value = $tmp; }
			}
		}
		$NicInformation += @{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
		$NicInformation += @{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{ Data = ""; Value = $tmp; }
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation += @{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += @{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }
		$NicInformation += @{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{ Data = "Scope ID"; Value = $Nic.winsscopeid; }
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		If($ThisNic.Name -eq $nic.description)
		{
			Line 2 "Name`t`t`t: " $ThisNic.Name
		}
		Else
		{
			Line 2 "Name`t`t`t: " $ThisNic.Name
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		Line 2 "Manufacturer`t`t: " $ThisNic.manufacturer
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "" $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t:" $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t:" $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		If($ThisNic.Name -eq $nic.description)
		{
			$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		}
		Else
		{
			$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
			$rowdata += @{ Data = "Description"; Value = $Nic.description; }
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway,$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		FormatHTMLTable
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2' }

			'da-'	{ 'Automatisk tabel 2' }

			'de-'	{ 'Automatische Tabelle 2' }

			'en-'	{ 'Automatic Table 2' }

			'es-'	{ 'Tabla automática 2' }

			'fi-'	{ 'Automaattinen taulukko 2' }

			'fr-'	{ 'Sommaire Automatique 2' }

			'nb-'	{ 'Automatisk tabell 2' }

			'nl-'	{ 'Automatische inhoudsopgave 2' }

			'pt-'	{ 'Sumário Automático 2' }

			'sv-'	{ 'Automatisk innehållsförteckning2' }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
					"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
					"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
					"Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
					"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Script:Word -eq $Null)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($BuildingBlocks -ne $Null)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($part -ne $Null)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Script:Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Script:Selection -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
	
	
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

    0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 ""

    Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

    Writes a line with 1 tab stop.
	
.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $null 0 $htmlitalics

    Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $null 0 $htmlbold

    Writes a line omitting font and font size and setting the bold attribute
	
.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $null 0 ($htmlbold -bor $htmlitalics)

    Writes a line omitting font and font size and setting both italics and bold options
	
.EXAMPLE	
    WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $null 2  # 10 point font

    Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 
    
	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

    Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

    Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
        0 - default which actually gives it a 2 or 10 point.
        1 - 7.5 point font size
        2 - 10 point
        3 - 13.5 point
        4 - 15 point
        5 - 18 point
        6 - 24 point
        7 - 36 point
            Any number larger than 7 defaults to 7

    Style - Refers to the headers that are used with output and resemble the headers in word, HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not give you
    a blue colored font, you will have to set that yourself.

    Colors and Bold/Italics Flags are:

        htmlbold       
        htmlitalics    
        htmlred        
        htmlcyan        
        htmlblue       
        htmldarkblue   
        htmllightblue   
        htmlpurple      
        htmlyellow      
        htmllime       
        htmlmagenta     
        htmlwhite       
        htmlsilver      
        htmlgray       
        htmlolive       
        htmlorange      
        htmlmaroon      
        htmlgreen       
        htmlblack       

    
    

#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
    [string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)
	
	
	#Build output style
	[string]$output = ""
	
    If([String]::IsNullOrEmpty($Name))	
    {
        $HTMLBody = "<p></p>"
    }
    Else
        {

        $color = CheckHTMLColor $options

	    #build # of tabs
        
	    While($tabs -gt 0)
	    { 
		    $output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
	    }
 
	    $HTMLFontName = $fontName		
	   
        $HTMLBody = ""

	    If($options -band $htmlitalics) 
	    {
		    $HTMLBody += "<i>"
	    } 
 
	    If($options -band $htmlbold) 
	    {
		    $HTMLBody += "<b>"
	    } 

	    #output the rest of the parameters.
	    $output += $name + $value
	

        $HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
        Switch ($style)
	    {
		
		    1 {$HTMLStyle = "<h1>"}
		    2 {$HTMLStyle = "<h2>"}
		    3 {$HTMLStyle = "<h3>"}
		    4 {$HTMLStyle = "<h4>"}
		    Default {$HTMLStyle = ""}
	    }
    
        $HTMLBody += $HTMLStyle + $output

        Switch ($style)
	    {
		
		    1 {$HTMLStyle = "</h1>"}
		    2 {$HTMLStyle = "</h2>"}
		    3 {$HTMLStyle = "</h3>"}
		    4 {$HTMLStyle = "</h4>"}
		    Default {$HTMLStyle = ""}
	    }

        $HTMLBody += $HTMLStyle +  "</font>"

        If($options -band $htmlitalics) 
	    {
		    $HTMLBody += "</i>"
	    } 
 
	    If($options -band $htmlbold) 
	    {
		    $HTMLBody += "</b>"
	    } 
	}
	
	#echo $HTMLBody >> $FileName1
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2)

	For($rowidx = $RowIndex;$rowidx -le $NumRows;$rowidx++)
	{
		$rd = @($rowdata[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$fontitalics = $False
			$fontbold = $false
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			$htmlbody += "<td bgcolor='" + $tmp + "'><font face='" + $fontname + "' size='" + $fontsize + "'>"
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($rd[$columnIndex] -ne $null)
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}

		$htmlbody += "</tr>"
	}
	#echo $HTMLBody >> $FileName1
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
	
	
.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
    defaults are used if not supplied.

    for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
    which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

    FormatHTMLTable "Table Heading" "auto"
    FormatHTMLTable "Table Heading" "25%
    FormatHTMLTable "Table Heading" "400px"

.NOTES
    In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

    First, initialize the table array

    $rowdata = @()

    Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
    and the second and subsequent lines go into the $rowdata table as shown below:

    $columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

    The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
    not format correctly.

    This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

    $rowdata = @()
    $columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
    $rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
    $rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
    $rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
    $rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
    $rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
    $rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
    $rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
    $rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
    $rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$RunningOS,$htmlwhite))
    $rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
    $rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
    FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table"

     Colors and Bold/Italics Flags are shown below:

        htmlbold       
        htmlitalics    
        htmlred        
        htmlcyan        
        htmlblue       
        htmldarkblue   
        htmllightblue   
        htmlpurple      
        htmlyellow      
        htmllime       
        htmlmagenta     
        htmlwhite       
        htmlsilver      
        htmlgray       
        htmlolive       
        htmlorange      
        htmlmaroon      
        htmlgreen       
        htmlblack     

#>

Function FormatHTMLTable
{
    Param([string]$tableheader,
    [string]$tablewidth="auto",
    [string]$fontName="Calibri",
	[int]$fontSize=2)

    $HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnHeaders.Length -eq 0 -or $columnHeaders -eq $null)
	{
		$NumCols = 2
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnHeaders.Length
	}  # need to add one for the color attrib
		
	If($rowdata -ne $null)
	{
		$NumRows = $rowdata.length + 1
	}
	Else
	{
		$NumRows = 1
	}
	
	$htmlbody += "<table border='1' width='" + $tablewidth + "'><tr>"
       
   	For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
	{
		$tmp = CheckHTMLColor $columnheaders[$columnIndex+1]
		
		$htmlbody += "<td bgcolor='" + $tmp + "'><font face='" + $fontname + "' size='" + $fontsize + "'>"

		If($columnheaders[$columnIndex+1] -band $htmlbold)
		{
			$htmlbody += "<b>"
		}
		If($columnheaders[$columnIndex+1] -band $htmlitalics)
		{
			$htmlbody += "<i>"
		}
		If($columnheaders[$columnIndex] -ne $null)
		{
			If($columnheaders[$columnIndex] -eq " " -or $columnheaders[$columnIndex].length -eq 0)
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			Else
			{
				$found = $false
				For($i=0;$i -lt $columnHeaders[$columnIndex].length;$i+=2)
				{
					If($columnheaders[$columnIndex][$i] -eq " ")
					{
						$htmlbody += "&nbsp;"
					}
					If($columnheaders[$columnIndex][$i] -ne " ")
					{
						Break
					}
				}
				$htmlbody += $columnHeaders[$columnIndex]
			}
		}
		Else
		{
			$htmlbody += "&nbsp;&nbsp;&nbsp;"
		}
		If($columnheaders[$columnIndex+1] -band $htmlbold)
		{
			$htmlbody += "</b>"
		}
		If($columnheaders[$columnIndex+1] -band $htmlitalics)
		{
			$htmlbody += "</i>"
		}
		$htmlbody += "</font></td>"
	}
		
	$htmlbody += "</tr>"
		
	$rowindex = 2
	If($RowData -ne $null)
	{
		AddHTMLTable $fontName $fontSize
		$rowdata = @()
	}
		
	$htmlbody = "</table>"
	#echo $HTMLBody >> $FileName1
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
    $columnHeaders = @()
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
    If($AddDateTime)
    {
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
    }

    $htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
    #echo $htmlhead > $FileName1
	out-file -FilePath $Script:FileName1 -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Columns -eq $Null) -and ($Headers -ne $Null)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Columns -ne $Null) -and ($Headers -ne $Null)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Columns -eq $Null) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $Null) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Columns -eq $Null) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Headers -ne $Null) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
				If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Font -ne $null) { $Cell.Range.Font.Name = $Font; }
					If($Color -ne $null) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($BackgroundColor -ne $null) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Company Name : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Cover Page   : $($CoverPage)"
	Write-Verbose "$(Get-Date): User Name    : $($UserName)"
	Write-Verbose "$(Get-Date): Save As PDF  : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD : $($MSWORD)"
	Write-Verbose "$(Get-Date): Save As HTML : $($HTML)"
	Write-Verbose "$(Get-Date): Add DateTime : $($AddDateTime)"
	Write-Verbose "$(Get-Date): HW Inventory : $($Hardware)"
	Write-Verbose "$(Get-Date): Filename1    : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2    : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): OS Detected  : $($RunningOS)"
	Write-Verbose "$(Get-Date): PSUICulture  : $($PSUICulture)"
	Write-Verbose "$(Get-Date): PSCulture    : $($PSCulture)"
	Write-Verbose "$(Get-Date): Word version : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): Word language: $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2010 and detected operating system $($RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running Word 2013 and detected operating system $($RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Deleting $($Script:FileName1) since only $($Script:FileName2) is needed"
		Remove-Item $Script:FileName1 4>$Null
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
    #echo "<p></p></body></html>" >> $FileName1
	out-file -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	$pwdpath = $pwd.Path

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
       SetupHTML
	}
}

Function TestComputerName
{
	Param([string]$Cname)
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $($CName) is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "`n`n`t`tComputer $($CName) is offline.`nScript cannot continue.`n`n"
			Exit
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($CName)"
		Return $CName
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Result -ne $Null)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($CName)"
			Return $CName
		}
		Else
		{
			Write-Warning "Unable to resolve $($CName) to a hostname"
		}
	}
	Else
	{
		#computer is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}
	Return $CName
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
#endregion

#region deletable functions
### If needed, you can delete the following functions ###
Function ProcessSampleText
{
	OutputSampleText
}

Function OutputSampleText
{
	If($MSWORD -or $PDF)
	{
		#we have a cover page and the page for the table of contents
		#insert a new page to start the report

		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "This is Heading 1"

		WriteWordLine 2 0 "This is Heading 2"

		WriteWordLine 3 0 "This is Heading 3"

		WriteWordLine 4 0 "This is Heading 4"

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops"
		#the next line is a blank line used for spacing
		WriteWordLine 0 0 ""

		WriteWordLine 0 1 "This is a regular line of text indented 1 tab stops"
		WriteWordLine 0 0 ""

		WriteWordLine 0 2 "This is a regular line of text indented 2 tab stops"
		WriteWordLine 0 0 ""

		WriteWordLine 0 3 "This is a regular line of text indented 3 tab stops"
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text in the default font in italics" "" $null 0 $true $false
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text in the default font in bold" "" $null 0 $false $true
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text in the default font in bold italics" "" $null 0 $true $true
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text in the default font in 14 point" "" $null 14 $false $false
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 $false $false
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data: " $env:computername 
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in bold: " $env:computername $null 0 $false $true
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in bold italics: " $env:computername $null 0 $true $true
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in 14 point bold italics: " $env:computername $null 14 $true $true
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in 8 point Courier New bold italics: " $env:computername "Courier New" 8 $true $true
	}
	ElseIf($Text)
	{
		Line 0 "Report Title: " $Script:Title
		Line 0 ""
		
		Line 0 "This are no insert new page or headings using the Text option"

		Line 0 "This is a regular line of text indented 0 tab stops"
		#the next line is a blank line used for spacing
		Line 0 ""

		Line 1 "This is a regular line of text indented 1 tab stops"
		Line 0 ""

		Line 2 "This is a regular line of text indented 2 tab stops"
		Line 0 ""

		Line 3 "This is a regular line of text indented 3 tab stops"
		Line 0 ""

		Line 0 "This are no fonts or italics or bold using the Text option"
		Line 0 ""

		Line 0 "This is a regular line of text indented 0 tab stops with the computer name as data: " $env:computername 
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "This is Heading 1"

		WriteHTMLLine 2 0 "This is Heading 2"

		WriteHTMLLine 3 0 "This is Heading 3"

		WriteHTMLLine 4 0 "This is Heading 4"

		WriteHTMLLine 0 0 "This is a regular line of text indented 0 tab stops"
		#the next line is a blank line used for spacing
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 2 "This is a regular line of text indented 2 tab stops"
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 3 "This is a regular line of text indented 3 tab stops"
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $null 0 $htmlitalics
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $null 0 $htmlbold
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $null 0 ($htmlbold -bor $htmlitalics)
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 7.5 point" "" $null 1   # 7.5 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $null 2   # 10 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 13.5 point" "" $null 3   # 13.5 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 15 point" "" $null 4   # 15 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 18 point" "" $null 5   # 18 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 24 point" "" $null 6   # 24 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in the default font in 36 point" "" $null 7   # 36 point font
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data: " $env:computername 
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in bold: " $env:computername $null 0 $htmlitalics
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in bold italics: " $env:computername $null 0 ($htmlbold -bor $htmlitalics)
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 13.5 point bold italics: " $env:computername $null 4 ($htmlred -bor $htmlbold -bor $htmlitalics)
		WriteHTMLLine 0 0 ""

		WriteHTMLLine 0 0 "This is a regular line of text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlitalics)
	}
}

Function ProcessServices
{
	Write-Verbose "$(Get-Date): `tGathering Computer services information"

	Try
	{
		#Iain Brighton optimization 5-Jun-2014
		#Replaced with a single call to retrieve services via WMI. The repeated
		## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
		## If we need to retrieve the StartUp type might as well just use WMI.
		$Script:Services = Get-WMIObject Win32_Service -ComputerName $ComputerName -EA 0 | Sort DisplayName
	}

	Catch
	{
		$Script:Services = $Null
	}

	If(!$? -or $Script:Services -eq $Null)
	{
		Write-Warning "No services were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Services were retrieved"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $htmlbold
		}
		Return $False
	}
	ElseIf($? -and $Script:Services -eq $Null)
	{
		Write-Warning "Services retrieval was successful but no services were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Services retrieval was successful but no services were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $htmlbold
		}
		Return $False
	}
	Else
	{
		If($Services -is [array])
		{
			[int]$Script:NumServices = $Services.count
		}
		Else
		{
			[int]$Script:NumServices = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($Script:NumServices) Services found"
		Return $True
	}
}

Function ProcessAutoFitHorizontalTable
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`tProcessing services information for AutoFit Horizontal Table with Grid Lines"
		WriteWordLine 1 0 "Example of AutoFit Horizontal Table"
		WriteWordLine 3 0 "Services"
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing services information"
		Line 0 "Services"
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing services information"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Example of AutoFitContents HTML Table"
		WriteHTMLLine 3 0 "Services"
		WriteHTMLLine 0 0 ""
	}
	
	OutputAutoFitHorizontalTable $Script:Services $Script:NumServices
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`tProcessing services information for AutoFit Horizontal Table with no Internal Grid Lines"
		OutputAutoFitHorizontalTableNoInternalGridLines $Script:Services $Script:NumServices
		Write-Verbose "$(Get-Date): `t`tProcessing services information for AutoFit Horizontal Table with no Grid Lines"
		OutputAutoFitHorizontalTableNoGridLines $Script:Services $Script:NumServices
	}
}

Function OutputAutoFitHorizontalTable
{
	Param([object]$Services, [int]$NumServices)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Services ($($NumServices) Services found)"

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ServicesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		Line 0 "Services ($NumServices Services found)"
		Line 0 ""
	}
	ElseIf($HTML)
	{
       $rowdata = @()
	}

	ForEach($Service in $Services) 
	{
		#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

			## Add the hash to the array
			$ServicesWordTable += $WordTableRowHash;

			## Store "to highlight" cell references
			If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
			{
				$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
			}
			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
			Line 0 "Display Name`t: " $Service.DisplayName
			Line 0 "Status`t`t: " $Service.State
			Line 0 "Start Mode`t: " $Service.StartMode
			Line 0 ""
		}
		ElseIf($HTML)
		{
			If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
			{
				$HighlightedCells = $htmlred
			}
			Else
			{
				$HighlightedCells = $htmlwhite
			} 
			$rowdata += @(,
			($Service.DisplayName,$htmlwhite,$Service.State,$HighlightedCells,$Service.StartMode,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ServicesWordTable `
		-Columns DisplayName, Status, StartMode `
		-Headers "Display Name", "Status", "Startup Type" `
		-Format -155 `
		-AutoFit $wdAutoFitContent;

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))
		$msg = "Services ($NumServices Services found)"
		FormatHTMLTable $msg "auto"
		WriteHTMLLine 0 0 ""
		WriteHTMLLine 1 0 "Example of AutoFitContents With Word Wrap HTML Table"
		WriteHTMLLine 0 0 ""
		FormatHTMLTable $msg "25%"
	}
}

Function OutputAutoFitHorizontalTableNoInternalGridLines
{
	Param([object]$Services, [int]$NumServices)
	
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Example of AutoFit Horizontal Table with no Internal Grid Lines"
		WriteWordLine 3 0 "Services"
		WriteWordLine 0 1 "Services ($($NumServices) Services found)"

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ServicesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}

	ForEach($Service in $Services) 
	{
		#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

			## Add the hash to the array
			$ServicesWordTable += $WordTableRowHash;

			## Store "to highlight" cell references
			If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
			{
				$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
			}
			$CurrentServiceIndex++;
		}
	}
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ServicesWordTable `
		-Columns DisplayName, Status, StartMode `
		-Headers "Display Name", "Status", "Startup Type" `
		-Format -155 `
		-NoInternalGridLines `
		-AutoFit $wdAutoFitContent;

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
}

Function OutputAutoFitHorizontalTableNoGridLines
{
	Param([object]$Services, [int]$NumServices)
	
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Example of AutoFit Horizontal Table with no Grid Lines"
		WriteWordLine 3 0 "Services"
		WriteWordLine 0 1 "Services ($($NumServices) Services found)"

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ServicesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}

	ForEach($Service in $Services) 
	{
		#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

			## Add the hash to the array
			$ServicesWordTable += $WordTableRowHash;

			## Store "to highlight" cell references
			If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
			{
				$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
			}
			$CurrentServiceIndex++;
		}
	}
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ServicesWordTable `
		-Columns DisplayName, Status, StartMode `
		-Headers "Display Name", "Status", "Startup Type" `
		-Format -155 `
		-NoGridLines `
		-AutoFit $wdAutoFitContent;

		## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		## IB - Set the required highlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
}

Function ProcessFixedWidthHorizontalTable
{
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		Write-Verbose "$(Get-Date): `t`tProcessing services information for Fixed Width Horizontal Table"
		WriteWordLine 1 0 "Example of Fixed Width Horizontal Table"
		WriteWordLine 3 0 "Services"
	}
	ElseIf($Text)
	{
		#there is no example of this for the text option
	}
	ElseIf($HTML)
	{
	}
	
	OutputFixedWidthHorizontalTable $Script:Services $Script:NumServices
}

Function OutputFixedWidthHorizontalTable
{
	Param([object]$Services, [int]$NumServices)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Services ($($NumServices) Services found)"

		## IB - replacement Services table generation utilising AddWordTable function

		## Create an array of hashtables to store our services
		[System.Collections.Hashtable[]] $ServicesWordTable = @();
		## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
		[System.Collections.Hashtable[]] $HighlightedCells = @();
		## Seed the $Services row index from the second row
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		#there is no example of this for the text option
	}
	ElseIf($HTML)
	{
	}

	ForEach($Service in $Services) 
	{
		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

			## Add the hash to the array
			$ServicesWordTable += $WordTableRowHash;

			## Store "to highlight" cell references
			If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
			{
				$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
			}
			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
		}
		ElseIf($HTML)
		{
		}
	}

	If($MSWORD -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ServicesWordTable `
		-Columns DisplayName,Status,StartMode `
		-Headers "Display Name","Status","Startup Type" `
		-Format -155 `
		-AutoFit $wdAutoFitFixed;

		## IB - Set alternating row color before we override the header
		#SetWordTableAlternateRowColor -Table $Table -BackgroundColor $wdColorGray05;
		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		## IB - Set the required hightlighted cells
		SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 100;

		#indent the entire table 1 tab stop
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
	}
}

Function ProcessScriptInformation
{
	## IB - Build Script information
	Write-Verbose "$(Get-Date): `tBuilding script information"
	[System.Collections.Hashtable[]] $Script:ScriptInformation = @()
	$Script:ScriptInformation += @{ Data = "Company Name"; Value = $Script:CoName; }
	$Script:ScriptInformation += @{ Data = "Cover Page"; Value = $CoverPage; }
	$Script:ScriptInformation += @{ Data = "User Name"; Value = $UserName; }
	$Script:ScriptInformation += @{ Data = "Save as PDF"; Value = $PDF; }
	$Script:ScriptInformation += @{ Data = "Save as TEXT"; Value = $TEXT; }
	$Script:ScriptInformation += @{ Data = "Save as WORD"; Value = $MSWORD; }
	$Script:ScriptInformation += @{ Data = "Save as HTML"; Value = $HTML; }
	$Script:ScriptInformation += @{ Data = "Add DateTime"; Value = $AddDateTime; }
	$Script:ScriptInformation += @{ Data = "Hardware Inventory"; Value = $Hardware; }
	$Script:ScriptInformation += @{ Data = "Computer Name"; Value = $ComputerName; }
	If($MSWORD -or $PDF)
	{
		$Script:ScriptInformation += @{ Data = "Title"; Value = $Script:Title; }
	}
	$Script:ScriptInformation += @{ Data = "Filename1"; Value = $Script:FileName1; }
	## IB - We only need Filename2 if it's a PDF (and we're no longer worried about the number of rows!)
	If($PDF) 
	{ 
		$Script:ScriptInformation += @{ Data = "Filename2"; Value = $Script:Filename2; } 
	}
	$Script:ScriptInformation += @{ Data = "OS Detected"; Value = $RunningOS; }
	$Script:ScriptInformation += @{ Data = "PSUICulture"; Value = $PSUICulture; }
	$Script:ScriptInformation += @{ Data = "PSCulture"; Value = $PSCulture; }
	$Script:ScriptInformation += @{ Data = "Word version"; Value = $WordProduct; }
	$Script:ScriptInformation += @{ Data = "Word language"; Value = $WordLanguageValue; }
	$Script:ScriptInformation += @{ Data = "PoSH version"; Value = $Host.Version; }
}

Function ProcessAutoFitVerticalTable
{
	Write-Verbose "$(Get-Date): `t`tProcessing script information for AutoFit Vertical Table"
	OutputAutoFitVerticalTable $Script:ScriptInformation
}

Function OutputAutoFitVerticalTable
{
	Param([object]$ScriptInformation)
	
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Example of Horizontal AutoFit Vertical Table"

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format -155 `
		-AutoFit $wdAutoFitContent;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Save as TEXT`t`t: " $TEXT
		Line 0 "Add DateTime`t`t: " $AddDateTime
		Line 0 "Hardware Inventory`t: " $Hardware
		Line 0 "Computer Name`t`t: " $ComputerName
		Line 0 "Title`t`t`t: " $Script:Title
		Line 0 "Filename1`t`t: " $Script:FileName1
		Line 0 "PSUICulture`t`t: " $PSUICulture
		Line 0 "PSCulture`t`t: " $PSCulture
		Line 0 "PoSH version`t`t: " $Host.Version
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
		$rowdata = @()
		$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
		$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
		$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
		$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
		$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
		$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
		$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
		$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
		$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
		$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$RunningOS,$htmlwhite))
		$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
		$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
		FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" "auto"
	}
}

Function ProcessFixedWidthVerticalTable
{
	Write-Verbose "$(Get-Date): `t`tProcessing script information for Fixed Width Vertical Table"
	OutputFixedWidthVerticalTable $Script:ScriptInformation
}

Function OutputFixedWidthVerticalTable
{
	Param([object]$ScriptInformation)
	
	If($MSWORD -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Example of Horizontal Fixed Width Vertical Table"

		## We already have a hashtable of script info, so just reutilise it!
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format -155 `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 170;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
		$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
		$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
		$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
		$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
		$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
		$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
		$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
		$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
		$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$RunningOS,$htmlwhite))
		$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
		$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
		FormatHTMLTable "Example of Horizontal Fixed Width HTML Table" "370"
	}
}
### If needed, you can delete the preceeding functions ###
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	$ComputerName = TestComputerName $ComputerName
}
#endregion

#region script core
#Script begins

ProcessScriptSetup


###REPLACE AFTER THIS SECTION WITH YOUR SCRIPT###

###The function SetFileName1andFileName2 needs your script output filename###
SetFileName1andFileName2 'InventoryScript'

###change title for your report###
[string]$Script:Title = 'System Center 2012 R2 Configuration Manager Documentation Script for {0}' -f $CompanyName;

###REPLACE AFTER THIS SECTION WITH YOUR SCRIPT###

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

function Get-ExecuteWqlQuery
{
  param
  (
    [System.Object]
    $siteServerName,
    
    [System.Object]
    $query
  )
  
  $returnValue = $null
  $connectionManager = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager
  
  if($connectionManager.Connect($siteServerName))
  {
    $result = $connectionManager.QueryProcessor.ExecuteQuery($query)
    
    foreach($i in $result.GetEnumerator())
    {
      $returnValue = $i
      break
    }
    
    $connectionManager.Dispose() 
  }
  
  $returnValue
}

function Get-ApplicationObjectFromServer
{
  param
  (
    [System.Object]
    $appName,
    
    [System.Object]
    $siteServerName
  )
  
  $resultObject = Get-ExecuteWqlQuery $siteServerName 'select thissitecode from sms_identification' 
  $siteCode = $resultObject['thissitecode'].StringValue
  
  $path = [string]::Format('\\{0}\ROOT\sms\site_{1}', $siteServerName, $siteCode)
  $scope = New-Object System.Management.ManagementScope -ArgumentList $path
  
  $query = [string]::Format("select * from sms_application where LocalizedDisplayName='{0}' AND ISLatest='true'", $appName.Trim())
  
  $oQuery = New-Object System.Management.ObjectQuery -ArgumentList $query
  $obectSearcher = New-Object System.Management.ManagementObjectSearcher -ArgumentList $scope,$oQuery
  $applicationFoundInCollection = $obectSearcher.Get()    
  $applicationFoundInCollectionEnumerator = $applicationFoundInCollection.GetEnumerator()
  
  if($applicationFoundInCollectionEnumerator.MoveNext())
  {
    $returnValue = $applicationFoundInCollectionEnumerator.Current
    $getResult = $returnValue.Get()        
    $sdmPackageXml = $returnValue.Properties['SDMPackageXML'].Value.ToString()
    [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($sdmPackageXml)
  }
}

function Load-ConfigMgrAssemblies()
{
  
  $AdminConsoleDirectory = Split-Path $env:SMS_ADMIN_UI_PATH -Parent
  $filesToLoad = 'Microsoft.ConfigurationManagement.ApplicationManagement.dll','AdminUI.WqlQueryEngine.dll', 'AdminUI.DcmObjectWrapper.dll' 
  
  Set-Location $AdminConsoleDirectory
  [System.IO.Directory]::SetCurrentDirectory($AdminConsoleDirectory)
  
  foreach($fileName in $filesToLoad)
  {
    $fullAssemblyName = [System.IO.Path]::Combine($AdminConsoleDirectory, $fileName)
    if([System.IO.File]::Exists($fullAssemblyName ))
    {   
      $FileLoaded = [Reflection.Assembly]::LoadFrom($fullAssemblyName )
    }
    else
    {
      Write-Output ([System.String]::Format('File not found {0}',$fileName )) -backgroundcolor 'red'
    }
  }
}

$SiteCode = Get-SiteCode

Write-Verbose "$(Get-Date): Start writing report data"

$LocationBeforeExecution = Get-Location

$Script:selection.InsertNewPage() | Out-Null

#Import the CM12 Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
{
  Write-Verbose "$(Get-Date):   CM12 module has not been imported yet, will import it now."
  Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null
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

WriteWordLine 1 0 'Summary of all Sites in this Hierarchy'
Write-Verbose "$(Get-Date):   Getting Site Information"
$CMSites = Get-CMSite

$CAS                    = $CMSites | Where-Object {$_.Type -eq 4}
$ChildPrimarySites      = $CMSites | Where-Object {$_.Type -eq 3}
$StandAlonePrimarySite  = $CMSites | Where-Object {$_.Type -eq 2}
$SecondarySites         = $CMSites | Where-Object {$_.Type -eq 1}

#region CAS
if (-not [string]::IsNullOrEmpty($CAS))
{
  WriteWordLine 0 1 'The following Central Administration Site is installed:'
  $CAS = @{'Site Name' = $CAS.SiteName; 'Site Code' = $CAS.SiteCode; Version = $CAS.Version };
  
  $Table = AddWordTable -Hashtable $CAS -Format -155 -AutoFit $wdAutoFitFixed;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
  ## IB - set column widths without recursion
  $Table.Columns.Item(1).Width = 100;
  $Table.Columns.Item(2).Width = 170;
  
  $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
  $Table.AutoFitBehavior($wdAutoFitFixed)
  
  FindWordDocumentEnd
  $Table = $Null
}
else {
    WriteWordLine 0 1 'No' -nonewline
    WriteWordLine 0 0 ' CAS' -boldface $true -nonewline
    WriteWordLine 0 1 ' detected. Continue with Primary Sites.'
}
#endregion CAS

#region Child Primary Sites
if (-not [string]::IsNullOrEmpty($ChildPrimarySites))
{
  Write-Verbose "$(Get-Date):   Enumerating all child Primary Site."
  WriteWordLine 0 1 'The following child Primary Sites are installed:'
  $StandAlonePrimarySite = @{'Site Name' = $ChildPrimarySites.SiteName; 'Site Code' = $ChildPrimarySites.SiteCode; Version = $ChildPrimarySites.Version };
  
  $Table = AddWordTable -Hashtable $ChildPrimarySites -Format -155 -AutoFit $wdAutoFitFixed;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
  ## IB - set column widths without recursion
  $Table.Columns.Item(1).Width = 100;
  $Table.Columns.Item(2).Width = 170;
  
  $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
  $Table.AutoFitBehavior($wdAutoFitFixed)
  
  FindWordDocumentEnd
  $Table = $Null
}
#endregion Child Primary Sites

#region Standalone Primary
if (-not [string]::IsNullOrEmpty($StandAlonePrimarySite))
{
  Write-Verbose "$(Get-Date):   Enumerating a standalone Primary Site."
  WriteWordLine 0 1 'The following Primary Site is installed:'
  $SiteCULevel = (Invoke-Command -ComputerName $(Get-CMSiteRole -RoleName 'SMS Site Server').NALPath.tostring().split('\\')[2] -ScriptBlock {Get-ItemProperty -Path registry::hklm\software\microsoft\sms\setup | Select-Object CULevel} -ErrorAction SilentlyContinue ).CULevel
  $StandAlonePrimarySite = @{'Site Name' = $StandAlonePrimarySite.SiteName; 'Site Code' = $StandAlonePrimarySite.SiteCode; Version = $StandAlonePrimarySite.Version; 'CU Installed' = $SiteCULevel };
  
  $Table = AddWordTable -Hashtable $StandAlonePrimarySite -Format -155 -AutoFit $wdAutoFitFixed;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
  ## IB - set column widths without recursion
  $Table.Columns.Item(1).Width = 100;
  $Table.Columns.Item(2).Width = 170;
  
  $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
  $Table.AutoFitBehavior($wdAutoFitFixed)
  
  FindWordDocumentEnd
  $Table = $Null
}

#endregion Standalone Primary

#region Secondary Sites
if (-not [string]::IsNullOrEmpty($SecondarySites))
{
  Write-Verbose "$(Get-Date):   Enumerating all secondary sites."
  WriteWordLine 0 1 'The following Secondary Sites are installed:'
  $SecondarySites = @{'Site Name' = $SecondarySites.SiteName; 'Site Code' = $SecondarySites.SiteCode; Version = $SecondarySites.Version };
  
  $Table = AddWordTable -Hashtable $SecondarySites -Format -155 -AutoFit $wdAutoFitFixed;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
  ## IB - set column widths without recursion
  $Table.Columns.Item(1).Width = 100;
  $Table.Columns.Item(2).Width = 170;
  
  $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
  $Table.AutoFitBehavior($wdAutoFitFixed)
  
  FindWordDocumentEnd
  $Table = $Null
}
#endregion Secondary Sites

#region Site Configuration report

foreach ($CMSite in $CMSites)
{  
  Write-Verbose "$(Get-Date):   Checking each site's configuration."
  WriteWordLine 1 0 "Configuration Summary for Site $($CMSite.SiteCode)"
  WriteWordLine 0 0 ''   
 
  $SiteRoleWordTable = @()  
  $SiteRoles = Get-CMSiteRole -SiteCode $CMSite.SiteCode | Select-Object -Property NALPath, rolename

  WriteWordLine 2 0 'Site Roles'
  WriteWordLine 0 1 'The following Site Roles are installed in this site:'
  foreach ($SiteRole in $SiteRoles) {
    if (-not (($SiteRole.rolename -eq 'SMS Component Server') -or ($SiteRole.rolename -eq 'SMS Site System'))) {
        $SiteRoleRowHash = @{'Server Name' = ($SiteRole.NALPath).ToString().Split('\\')[2]; 'Role' = $SiteRole.RoleName}
        $SiteRoleWordTable += $SiteRoleRowHash
    }
  }
  
  $Table = AddWordTable -Hashtable $SiteRoleWordTable -Format -155 -AutoFit $wdAutoFitContent
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray1 -Underline -Italic
  
  FindWordDocumentEnd
  $Table = $Null

  $SiteMaintenanceTaskWordTable = @()
  $SiteMaintenanceTasks = Get-CMSiteMaintenanceTask -SiteCode $CMSite.SiteCode
  WriteWordLine 2 0 "Site Maintenance Tasks for Site $($CMSite.SiteCode)"
  
  foreach ($SiteMaintenanceTask in $SiteMaintenanceTasks) {
    $SiteMaintenanceTaskRowHash = @{'Task Name' = $SiteMaintenanceTask.TaskName; Enabled = $SiteMaintenanceTask.Enabled};
    $SiteMaintenanceTaskWordTable += $SiteMaintenanceTaskRowHash;
  }
  
  $Table = AddWordTable -Hashtable $SiteMaintenanceTaskWordTable -Format -155 -AutoFit $wdAutoFitContent;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  <#
  ## IB - set column widths without recursion
  $Table.Columns.Item(1).Width = 100;
  $Table.Columns.Item(2).Width = 170;
  
  $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
  $Table.AutoFitBehavior($wdAutoFitFixed)
  #>
  FindWordDocumentEnd
  $Table = $Null
  
  $CMManagementPoints = Get-CMManagementPoint -SiteCode $CMSite.SiteCode
  WriteWordLine 2 1 "Summary of Management Points for Site $($CMSite.SiteCode)"
  foreach ($CMManagementPoint in $CMManagementPoints)
  {
    Write-Verbose "$(Get-Date):   Management Point: $($CMManagementPoint)"
    $CMMPServerName = $CMManagementPoint.NetworkOSPath.Split('\\')[2]
    WriteWordLine 0 1 "$($CMMPServerName)"
  }
  
  WriteWordLine 2 1 "Summary of Distribution Points for Site $($CMSite.SiteCode)"
  $CMDistributionPoints = Get-CMDistributionPoint -SiteCode $CMSite.SiteCode
  
  foreach ($CMDistributionPoint in $CMDistributionPoints)
  {
    $CMDPServerName = $CMDistributionPoint.NetworkOSPath.Split('\\')[2]
    Write-Verbose "$(Get-Date):   Found DP: $($CMDPServerName)"
    WriteWordLine 0 1 "$($CMDPServerName)" -boldface $true
    Write-Verbose "Trying to ping $($CMDPServerName)"
    $PingResult = Test-NetConnection -ComputerName $CMDPServerName
    if (-not ($PingResult.PingSucceeded))
    {
      WriteWordLine 0 2 "The Distribution Point $($CMDPServerName) is not reachable. Check connectivity."
    }
    else
    {
      WriteWordLine 0 2 'Disk information:'
      $CMDPDrives = (Get-WmiObject -Class SMS_DistributionPointDriveInfo -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider).Where({$PSItem.NALPath -like "*$CMDPServerName*"})
      foreach ($CMDPDrive in $CMDPDrives)
      {
        WriteWordLine 0 2 "Partition $($CMDPDrive.Drive):" -boldface $true
        $Size = ''
        $Size = $CMDPDrive.BytesTotal / 1024 / 1024
        $Freesize = ''
        $Freesize = $CMDPDrive.BytesFree / 1024 / 1024
        
        WriteWordLine 0 3 "$([MATH]::Round($Size,2)) GB Size in total"
        WriteWordLine 0 3 "$([MATH]::Round($Freesize,2)) GB Free Space"
        WriteWordLine 0 3 "Still $($CMDPDrive.PercentFree) percent free."
        WriteWordLine 0 0 ''
      }
      
      WriteWordLine -style 0 -tabs 2 -value 'Hardware Info:' -boldface $true
      try {
          $Capacity = 0
          Get-WmiObject -Class win32_physicalmemory -ComputerName $CMDPServerName | ForEach-Object {[int64]$Capacity = $Capacity + [int64]$_.Capacity}
          $TotalMemory = $Capacity / 1024 / 1024 / 1024
            WriteWordLine 0 3 "This server has a total of $($TotalMemory) GB RAM."
      }
      catch {
        WriteWordLine 0 3 "Failed to access server $CMDPServerName." 
        }
    }
    
    $DPInfo = $CMDistributionPoint.Props
    $IsPXE = ($DPInfo.Where({$_.PropertyName -eq 'IsPXE'})).Value
    $UnknownMachines = ($DPInfo.Where({$_.PropertyName -eq 'SupportUnknownMachines'})).Value
    switch ($IsPXE)
    {
      1 
      {
        WriteWordLine 0 2 'PXE Enabled'
        switch ($UnknownMachines)
        {
          1 { WriteWordLine 0 2 'Supports unknown machines: true' }
          0 { WriteWordLine 0 2 'Supports unknown machines: false' }
        }
      }
      0
      {
        WriteWordLine 0 2 'PXE disabled'
      }
    }
    
    $DPGroupMembers = $Null
    $DPGroupIDs = $Null
    $DPGroupMembers = (Get-WmiObject -class SMS_DPGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider).Where({$_.DPNALPath -ilike "*$($CMDPServerName)*"});
    if (-not [string]::IsNullOrEmpty($DPGroupMembers))
    {
      $DPGroupIDs = $DPGroupMembers.GroupID
    }
    
    #enumerating DP Group Membership
    if (-not [string]::IsNullOrEmpty($DPGroupIDs))
    {
      WriteWordLine 0 1 'This Distribution Point is a member of the following DP Groups:'
      foreach ($DPGroupID in $DPGroupIDs)
      {
        $DPGroupName = (Get-CMDistributionPointGroup -Id "$($DPGroupID)").Name
        WriteWordLine 0 2 "$($DPGroupName)"
      }
    }
    else
    {
      WriteWordLine 0 1 'This Distribution Point is not a member of any DP Group.'
    }
  }
  #enumerating Software Update Points
  Write-Verbose "$(Get-Date):   Enumerating all Software Update Points"
  WriteWordLine 2 1 "Summary of Software Update Point Servers for Site $($CMSite.SiteCode)"
  $CMSUPs = Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $CMMPServerName | Where-Object {$_.rolename -eq 'SMS Software Update Point'}
  #$CMSUPs = (Get-CMSoftwareUpdatePoint).Where({$_.SiteCode -eq "$($CMSite.SiteCode)"})
  if (-not [string]::IsNullOrEmpty($CMSUPs))
  {
    foreach ($CMSUP in $CMSUPs) {
      $SUPHashTable = @();
      $CMSUPServerName = $CMSUP.NetworkOSPath.split('\\')[2]
      Write-Verbose "$(Get-Date):   Found SUP: $($CMSUPServerName)"
      WriteWordLine 0 1 "$($CMSUPServerName)"
      foreach ($SUPProp in $CMSUP.Props) {
        $SUPHash = @{Value2 = $SUPProp.Value2; Value1 = $SUPProp.Value1; Value = $SUPProp.Value; 'Property Name' = $SUPProp.PropertyName};
        $SUPHashTable += $SUPHash;
      }
      $Table = AddWordTable -Hashtable $SUPHashTable -Format -155 -AutoFit $wdAutoFitFixed;
      
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
      
      ## IB - set column widths without recursion
      $Table.Columns.Item(1).Width = 100;
      $Table.Columns.Item(2).Width = 170;
      
      $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
      $Table.AutoFitBehavior($wdAutoFitFixed)
      
      FindWordDocumentEnd
      $Table = $Null
    }
  }
  else
  {
    WriteWordLine 0 1 'This site has no Software Update Points installed.'
  }
}

##### Hierarchy wide configuration
WriteWordLine 1 0 'Summary of Hierarchy Wide Configuration'

#region enumerating Boundaries
Write-Verbose "$(Get-Date): Enumerating all Site Boundaries"
WriteWordLine 2 0 'Summary of Site Boundaries'

$Boundaries = Get-CMBoundary
    if (-not [string]::IsNullOrEmpty($Boundaries))
{
  $SubnetHashTable  = @();
  $ADHashTable      = @();
  $IPv6HashTable    = @();
  $IPRangeHashTable = @();
  
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
      $SubnetHash = @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)"
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems"
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $SubnetHashTable += $SubnetHash;
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
      $ADHash = @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)"
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems";
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $ADHashTable += $ADHash;
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
      $IPv6Hash = @{'Boundary Type' = $BoundaryType; 
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)"
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems";
                    Description = $Boundary.DisplayName;
                    Value = $Boundary.Value;
                    }
      $IPv6HashTable += $IPv6Hash;
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
      $IPRangeHash = @{'Boundary Type' = $BoundaryType
                    'Default Site Code' = "$($Boundary.DefaultSiteCode)"
                    'Associated Site Systems' = "$NamesOfBoundarySiteSystems"
                    Description = $Boundary.DisplayName
                    Value = $Boundary.Value
                    }
      $IPRangeHashTable += $IPRangeHash
    }
  }
}
          
#region IPv6 Boundaries Table
      $Table = AddWordTable -Hashtable $IPv6HashTable -Format -155 -AutoFit $wdAutoFitContent;
      
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray1
      $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional) 
      FindWordDocumentEnd
      $Table = $Null
#endregion IPv6 Boundaries Table
WriteWordLine 0 0 ''
WriteWordLine 0 0 ''

#region IP Subnet Boundaries Table
      $Table = AddWordTable -Hashtable $SubnetHashTable -Format -155 -AutoFit $wdAutoFitContent
      
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15
      $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional) 
      FindWordDocumentEnd
      $Table = $Null
#endregion IP Subnet Boundaries Table

WriteWordLine 0 0 ''
WriteWordLine 0 0 ''

#region IP Range Boundaries Table
      $Table = AddWordTable -Hashtable $IPRangeHashTable -Format -155 -AutoFit $wdAutoFitContent ;
      
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
      $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional) 
      FindWordDocumentEnd
      $Table = $Null
#endregion IP Range Boundaries Table

WriteWordLine 0 0 ''
WriteWordLine 0 0 ''

#region AD Site Boundaries Table
      $Table = AddWordTable -Hashtable $ADHashTable -Format -155 -AutoFit $wdAutoFitContent;
      
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
      $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional) 
      FindWordDocumentEnd
      $Table = $Null
#endregion AD Site Boundaries Table

WriteWordLine 0 0 ''
WriteWordLine 0 0 ''

#endregion enumerating Boundaries

#region enumerating all Boundary Groups and their members
Write-Verbose "$(Get-Date):   Enumerating all Boundary Groups and their members"

$BoundaryGroups = Get-CMBoundaryGroup
WriteWordLine 2 0 'Summary of Site Boundary Groups'

$BoundaryGroupHashTable = @();
if (-not [string]::IsNullOrEmpty($BoundaryGroups))
{
  foreach ($BoundaryGroup in $BoundaryGroups) {
    $MemberNames = @();
    if ($BoundaryGroup.SiteSystemCount -gt 0)
    {
      $MemberIDs = (Get-WmiObject -Class SMS_BoundaryGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object -FilterScript {$_.GroupID -eq "$($BoundaryGroup.GroupID)"}).BoundaryID
      foreach ($MemberID in $MemberIDs)
      {
        $MemberName = (Get-CMBoundary -Id $MemberID).DisplayName
        $MemberNames += "$MemberName (ID: $MemberID); "
        Write-Verbose "Member names: $($MemberName)"
      }
    }
    else
    {
      $MemberNames += 'There are no Site Systems associated to this Boundary Group.'
      Write-Verbose 'There are no Site Systems associated to this Boundary Group.'
    }
    $BoundaryGroupRow = @{Name = $BoundaryGroup.Name; Description = $BoundaryGroup.Description; 'Boundary members' = "$MemberNames"};
    $BoundaryGroupHashTable += $BoundaryGroupRow;
  }
  
  $Table = AddWordTable -Hashtable $BoundaryGroupHashTable -Format -155 -AutoFit $wdAutoFitContent
  #-Columns Name, Description, 'Boundary Members' -Headers Name, Description, 'Boundary Members'
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15
  
  FindWordDocumentEnd
  $Table = $Null
}
else
{
  WriteWordLine 0 1 'There are no Boundary Groups configured. It is mandatory to configure a Boundary Group in order for CM12 to work properly.'
}

#endregion enumerating all Boundary Groups and their members

#region enumerating Client Policies
Write-Verbose "$(Get-Date):   Enumerating all Client/Device Settings"
WriteWordLine 2 0 'Summary of Custom Client Device Settings'

$AllClientSettings = Get-CMClientSetting | Where-Object -FilterScript {$_.SettingsID -ne '0'}
foreach ($ClientSetting in $AllClientSettings)
{
  WriteWordLine 0 1 "Client Settings Name: $($ClientSetting.Name)" -bold
  WriteWordLine 0 2 "Client Settings Description: $($ClientSetting.Description)"
  WriteWordLine 0 2 "Client Settings ID: $($ClientSetting.SettingsID)"
  WriteWordLine 0 2 "Client Settings Priority: $($ClientSetting.Priority)"
  if ($ClientSetting.Type -eq '1')
  {
    WriteWordLine 0 2 'This is a custom client Device Setting.'
  }
  else
  {
    WriteWordLine 0 2 'This is a custom client User Setting.'
  }
  WriteWordLine 0 1 'Configurations'
  foreach ($AgentConfig in $ClientSetting.AgentConfigurations)
  {
    try
    {
      switch ($AgentConfig.AgentID)
      {
        1
        {
          WriteWordLine 0 2 'Compliance Settings'
          WriteWordLine 0 2 "Enable compliance evaluation on clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 "Enable user data and profiles: $($AgentConfig.EnableUserStateManagement)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        2
        {
          WriteWordLine 0 2 'Software Inventory'
          WriteWordLine 0 2 "Enable software inventory on clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 'Schedule software inventory and file collection: ' -nonewline
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 0 " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 0 " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 2 'Inventory reporting detail: ' -nonewline
          switch ($AgentConfig.ReportOptions)
          {
            1 { WriteWordLine 0 0 'Product only' }
            2 { WriteWordLine 0 0 'File only' }
            7 { WriteWordLine 0 0 'Full details' }
          }
          
          WriteWordLine 0 2 'Inventory these file types: '
          if ($AgentConfig.InventoriableTypes)
          {
            WriteWordLine 0 3 "$($AgentConfig.InventoriableTypes)"
          }
          if ($AgentConfig.Path)
          {                               
            WriteWordLine 0 3 "$($AgentConfig.Path)"
          }
          if (($AgentConfig.InventoriableTypes) -and ($AgentConfig.ExcludeWindirAndSubfolders -eq 'true'))
          {
            WriteWordLine 0 3 'Exclude WinDir and Subfolders'
          }
          else 
          {
            WriteWordLine 0 3 'Do not exclude WinDir and Subfolders'
          }
          
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        3
        {
          WriteWordLine 0 2 'Remote Tools'
          WriteWordLine 0 2 'Enable Remote Control on clients: ' -nonewline
          switch ($AgentConfig.FirewallExceptionProfiles)
          {
            0 { WriteWordLine 0 0 'Disabled' }
            8 { WriteWordLine 0 0 'Enabled: No Firewall Profile.' }
            9 { WriteWordLine 0 2 'Enabled: Public.' }
            10 { WriteWordLine 0 2 'Enabled: Private.' }
            11 { WriteWordLine 0 2 'Enabled: Private, Public.' }
            12 { WriteWordLine 0 2 'Enabled: Domain.' }
            13 { WriteWordLine 0 2 'Enabled: Domain, Public.' }
            14 { WriteWordLine 0 2 'Enabled: Domain, Private.' }
            15 { WriteWordLine 0 2 'Enabled: Domain, Private, Public.' }
          }
          WriteWordLine 0 2 "Users can change policy or notification settings in Software Center: $($AgentConfig.AllowClientChange)"
          WriteWordLine 0 2 "Allow Remote Control of an unattended computer: $($AgentConfig.AllowRemCtrlToUnattended)"
          WriteWordLine 0 2 "Prompt user for Remote Control permission: $($AgentConfig.PermissionRequired)"
          WriteWordLine 0 2 "Grant Remote Control permission to local Administrators group: $($AgentConfig.AllowLocalAdminToDoRemoteControl)"
          WriteWordLine 0 2 'Access level allowed: ' -nonewline
          switch ($AgentConfig.AccessLevel)
          {
            0 { WriteWordLine 0 0 'No access' }
            1 { WriteWordLine 0 0 'View only' }
            2 { WriteWordLine 0 0 'Full Control' }
          }
          WriteWordLine 0 2 'Permitted viewers of Remote Control and Remote Assistance:'
          foreach ($Viewer in $AgentConfig.PermittedViewers)
          {
            WriteWordLine 0 3 "$($Viewer)"
          }
          WriteWordLine 0 2 "Show session notification icon on taskbar: $($AgentConfig.RemCtrlTaskbarIcon)"
          WriteWordLine 0 2 "Show session connection bar: $($AgentConfig.RemCtrlConnectionBar)"
          WriteWordLine 0 2 'Play a sound on client: ' -nonewline
          Switch ($AgentConfig.AudibleSignal)
          {
            0 { WriteWordLine 0 0 'None.' }
            1 { WriteWordLine 0 0 'Beginning and end of session.' }
            2 { WriteWordLine 0 0 'Repeatedly during session.' }
          }
          WriteWordLine 0 2 "Manage unsolicited Remote Assistance settings: $($AgentConfig.ManageRA)"
          WriteWordLine 0 2 "Manage solicited Remote Assistance settings: $($AgentConfig.EnforceRAandTSSettings)"
          WriteWordLine 0 2 'Level of access for Remote Assistance: ' -nonewline
          if (($AgentConfig.AllowRAUnsolicitedView -ne 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -ne 'True'))
          {
            WriteWordLine 0 0 'None.'
          }
          elseif (($AgentConfig.AllowRAUnsolicitedView -eq 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -ne 'True'))
          {
            WriteWordLine 0 0 'Remote viewing.'
          }
          elseif (($AgentConfig.AllowRAUnsolicitedView -eq 'True') -and ($AgentConfig.AllowRAUnsolicitedControl -eq 'True'))
          {
            WriteWordLine 0 0 'Full Control.'
          }
          WriteWordLine 0 2 "Manage Remote Desktop settings: $($AgentConfig.ManageTS)"
          if ($AgentConfig.ManageTS -eq 'True')
          {
            WriteWordLine 0 2 "Allow permitted viewers to connect by using Remote Desktop connection: $($AgentConfig.EnableTS)"
            WriteWordLine 0 2 "Require network level authentication on computers that run Windows Vista operating system and later versions: $($AgentConfig.TSUserAuthentication)"
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        4
        {
          WriteWordLine 0 2 'Computer Agent'
          WriteWordLine 0 2 "Deployment deadline greater than 24 hours, remind user every (hours): $([string]($AgentConfig.ReminderInterval) / 60 / 60)"
          WriteWordLine 0 2 "Deployment deadline less than 24 hours, remind user every (hours): $([string]($AgentConfig.DayReminderInterval) / 60 / 60)"
          WriteWordLine 0 2 "Deployment deadline less than 1 hour, remind user every (minutes): $([string]($AgentConfig.HourReminderInterval) / 60)"
          WriteWordLine 0 2 "Default application catalog website point: $($AgentConfig.PortalUrl)"
          WriteWordLine 0 2 "Add default Application Catalog website to Internet Explorer trusted sites zone: $($AgentConfig.AddPortalToTrustedSiteList)"
          WriteWordLine 0 2 "Allow Silverlight applications to run in elevated trust mode: $($AgentConfig.AllowPortalToHaveElevatedTrust)"
          WriteWordLine 0 2 "Organization name displayed in Software Center: $($AgentConfig.BrandingTitle)"
          switch ($AgentConfig.InstallRestriction)
          {
            0 { $InstallRestriction = 'All Users' }
            1 { $InstallRestriction = 'Only Administrators' }
            3 { $InstallRestriction = 'Only Administrators and primary Users'}
            4 { $InstallRestriction = 'No users' }
          }
          WriteWordLine 0 2 "Install Permissions: $($InstallRestriction)"
          Switch ($AgentConfig.SuspendBitLocker)
          {
            0 { $SuspendBitlocker = 'Never' }
            1 { $SuspendBitlocker = 'Always' }
          }
          WriteWordLine 0 2 "Suspend Bitlocker PIN entry on restart: $($SuspendBitlocker)"
          Switch ($AgentConfig.EnableThirdPartyOrchestration)
          {
            0 { $EnableThirdPartyTool = 'No' }
            1 { $EnableThirdPartyTool = 'Yes' }
          }
          WriteWordLine 0 2 "Additional software manages the deployment of applications and software updates: $($EnableThirdPartyTool)"
          Switch ($AgentConfig.PowerShellExecutionPolicy)
          {
            0 { $ExecutionPolicy = 'All signed' }
            1 { $ExecutionPolicy = 'Bypass' }
            2 { $ExecutionPolicy = 'Restricted' }
          }
          WriteWordLine 0 2 "Powershell execution policy: $($ExecutionPolicy)"
          switch ($AgentConfig.DisplayNewProgramNotification)
          {
            False { $DisplayNotifications = 'No' }
            True { $DisplayNotifications = 'Yes' }
          }
          WriteWordLine 0 2 "Show notifications for new deployments: $($DisplayNotifications)"
          switch ($AgentConfig.DisableGlobalRandomization)
          {
            False { $DisableGlobalRandomization = 'No' }
            True { $DisableGlobalRandomization = 'Yes' }
          }
          WriteWordLine 0 2 "Disable deadline randomization: $($DisableGlobalRandomization)"
          WriteWordLine 0 0 '---------------------'
        }
        5
        {
          WriteWordLine 0 2 'Network Access Protection (NAP)'
          WriteWordLine 0 2 "Enable Network Access Protection on clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 "Use UTC (Universal Time Coordinated) for evaluation time: $($AgentConfig.EffectiveTimeinUTC)"
          WriteWordLine 0 2 "Require a new scan for each evaluation: $($AgentConfig.ForceScan)"
          WriteWordLine 0 2 'NAP re-evaluation schedule:' -nonewline
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.ComputeComplianceSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 0 " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 0 " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 0 '---------------------'
        }
        8
        {
          WriteWordLine 0 2 'Software Metering'
          WriteWordLine 0 2 "Enable software metering on clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 'Schedule data collection: ' -nonewline
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.DataCollectionSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 0 " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 0 " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        9
        {
          WriteWordLine 0 2 'Software Updates'
          WriteWordLine 0 2 "Enable software updates on clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 'Software Update scan schedule: ' -nonewline
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.ScanSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 0 " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 0 " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 2 'Schedule deployment re-evaluation: ' -nonewline
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 0 " Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 0 " Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 0 " Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 2 'When any software update deployment deadline is reached, install all other software update deployments with deadline coming within a specified period of time: ' -nonewline
          if ($AgentConfig.AssignmentBatchingTimeout -eq '0')
          {
            WriteWordLine 0 0 'No.'
          }
          else 
          {
            WriteWordLine 0 0 'Yes.'    
            WriteWordLine 0 2 'Period of time for which all pending deployments with deadline in this time will also be installed: ' -nonewline
            if ($AgentConfig.AssignmentBatchingTimeout -le '82800')
            {
              $hours = [string]$AgentConfig.AssignmentBatchingTimeout / 60 / 60 
              WriteWordLine 0 0 "$($hours) hours"
            }
            else 
            {
              $days = [string]$AgentConfig.AssignmentBatchingTimeout / 60 / 60 / 24
              WriteWordLine 0 0 "$($days) days"
            }
          }
          
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        10
        {
          WriteWordLine 0 2 'User and Device Affinity'
          WriteWordLine 0 2 "User device affinity usage threshold (minutes): $($AgentConfig.ConsoleMinutes)"
          WriteWordLine 0 2 "User device affinity usage threshold (days): $($AgentConfig.IntervalDays)"
          WriteWordLine 0 2 'Automatically configure user device affinity from usage data: ' -nonewline 
          if ($AgentConfig.AutoApproveAffinity -eq '0')
          {
            WriteWordLine 0 0 'No'
          }
          else
          {
            WriteWordLine 0 0 'Yes'
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        11
        {
          WriteWordLine 0 2 'Background Intelligent Transfer'
          WriteWordLine 0 2 "Limit the maximum network bandwidth for BITS background transfers: $($AgentConfig.EnableBitsMaxBandwidth)"
          WriteWordLine 0 2 "Throttling window start time: $($AgentConfig.MaxBandwidthValidFrom)"
          WriteWordLine 0 2 "Throttling window end time: $($AgentConfig.MaxBandwidthValidTo)"
          WriteWordLine 0 2 "Maximum transfer rate during throttling window (kbps): $($AgentConfig.MaxTransferRateOnSchedule)"
          WriteWordLine 0 2 "Allow BITS downloads outside the throttling window: $($AgentConfig.EnableDownloadOffSchedule)"
          WriteWordLine 0 2 "Maximum transfer rate outside the throttling window (Kbps): $($AgentConfig.MaxTransferRateOffSchedule)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        12
        {
          WriteWordLine 0 2 'Enrollment'
          WriteWordLine 0 2 'Allow users to enroll mobile devices and Mac computers: ' -nonewline
          if ($AgentConfig.EnableDeviceEnrollment -eq '0')
          {
            WriteWordLine 0 0 'No'
          }
          else
          {
            WriteWordLine 0 0 'Yes'
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        13
        {
          WriteWordLine 0 2 'Client Policy'
          WriteWordLine 0 2 "Client policy polling interval (minutes): $($AgentConfig.PolicyRequestAssignmentTimeout)"
          WriteWordLine 0 2 "Enable user policy on clients: $($AgentConfig.PolicyEnableUserPolicyPolling)"
          WriteWordLine 0 2 "Enable user policy requests from Internet clients: $($AgentConfig.PolicyEnableUserPolicyOnInternet)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        15
        {
          WriteWordLine 0 2 'Hardware Inventory'
          WriteWordLine 0 2 "Enable hardware inventory on clients: $($AgentConfig.Enabled)"
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.Schedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 2 "Hardware inventory schedule: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 2 "Hardware inventory schedule: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 2 "Hardware inventory schedule: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 2 "Hardware inventory schedule: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 2 "Hardware inventory schedule: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 2 "Hardware inventory schedule: Occurs on last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 2 "Hardware inventory schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        16 
        {
          WriteWordLine 0 2 'State Messaging'
          WriteWordLine 0 2 "State message reporting cycle (minutes): $($AgentConfig.BulkSendInterval)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        17
        {
          WriteWordLine 0 2 'Software Deployment'
          $Schedule = Convert-CMSchedule -ScheduleString $AgentConfig.EvaluationSchedule
          if ($Schedule.DaySpan -gt 0)
          {
            WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.HourSpan -gt 0)
          {
            WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.MinuteSpan -gt 0)
          {
            WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfWeeks)
          {
            WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)"
          }
          elseif ($Schedule.ForNumberOfMonths)
          {
            if ($Schedule.MonthDay -gt 0)
            {
              WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
            elseif ($Schedule.MonthDay -eq 0)
            {
              WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs on last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
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
              WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
            }
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        18
        {
          WriteWordLine 0 2 'Power Management'
          WriteWordLine 0 2 "Allow power management of clients: $($AgentConfig.Enabled)"
          WriteWordLine 0 2 "Allow users to exclude their device from power management: $($AgentConfig.AllowUserToOptOutFromPowerPlan)"
          WriteWordLine 0 2 "Enable wake-up proxy: $($AgentConfig.EnableWakeupProxy)"
          if ($AgentConfig.EnableWakeupProxy -eq 'True')
          {
            WriteWordLine 0 2 "Wake-up proxy port number (UDP): $($AgentConfig.Port)"
            WriteWordLine 0 2 "Wake On LAN port number (UDP): $($AgentConfig.WolPort)"
            WriteWordLine 0 2 'Windows Firewall exception for wake-up proxy: ' -nonewline
            switch ($AgentConfig.WakeupProxyFirewallFlags)
            {
              0 { WriteWordLine 0 2 'disabled' }
              9 { WriteWordLine 0 2 'Enabled: Public.' }
              10 { WriteWordLine 0 2 'Enabled: Private.' }
              11 { WriteWordLine 0 2 'Enabled: Private, Public.' }
              12 { WriteWordLine 0 2 'Enabled: Domain.' }
              13 { WriteWordLine 0 2 'Enabled: Domain, Public.' }
              14 { WriteWordLine 0 2 'Enabled: Domain, Private.' }
              15 { WriteWordLine 0 2 'Enabled: Domain, Private, Public.' }
            }
            WriteWordLine 0 2 "IPv6 prefixes if required for DirectAccess or other intervening network devices. Use a comma to specifiy multiple entries: $($AgentConfig.WakeupProxyDirectAccessPrefixList)"
          }
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        20
        {
          WriteWordLine 0 2 'Endpoint Protection'
          WriteWordLine 0 2 "Manage Endpoint Protection client on client computers: $($AgentConfig.EnableEP)"
          WriteWordLine 0 2 "Install Endpoint Protection client on client computers: $($AgentConfig.InstallSCEPClient)"
          WriteWordLine 0 2 "Automatically remove previously installed antimalware software before Endpoint Protection is installed: $($AgentConfig.Remove3rdParty)"
          WriteWordLine 0 2 "Allow Endpoint Protection client installation and restarts outside maintenance windows. Maintenance windows must be at least 30 minutes long for client installation: $($AgentConfig.OverrideMaintenanceWindow)"
          WriteWordLine 0 2 "For Windows Embedded devices with write filters, commit Endpoint Protection client installation (requires restart): $($AgentConfig.PersistInstallation)"
          WriteWordLine 0 2 "Suppress any required computer restarts after the Endpoint Protection client is installed: $($AgentConfig.SuppressReboot)"
          WriteWordLine 0 2 "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): $($AgentConfig.ForceRebootPeriod)"
          WriteWordLine 0 2 "Disable alternate sources (such as Microsoft Windows Update, Microsoft Windows Server Update Services, or UNC shares) for the initial definition update on client computers: $($AgentConfig.DisableFirstSignatureUpdate)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        21
        {
          WriteWordLine 0 2 'Computer Restart'
          WriteWordLine 0 2 "Display a temporary notification to the user that indicates the interval before the user is logged of or the computer restarts (minutes): $($AgentConfig.RebootLogoffNotificationCountdownDuration)"
          WriteWordLine 0 2 "Display a dialog box that the user cannot close, which displays the countdown interval before the user is logged of or the computer restarts (minutes): $([string]$AgentConfig.RebootLogoffNotificationFinalWindow / 60)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        22
        {
          WriteWordLine 0 2 'Cloud Services'
          WriteWordLine 0 2 "Allow access to Cloud Distribution Point: $($AgentConfig.AllowCloudDP)"
          WriteWordLine 0 0 ''
          WriteWordLine 0 0 '---------------------'
        }
        23
        {
          WriteWordLine 0 2 'Metered Internet Connections'
          switch ($AgentConfig.MeteredNetworkUsage)
          {
            1 { $Usage = 'Allow' }
            2 { $Usage = 'Limit' }
            4 { $Usage = 'Block' }
          }
          WriteWordLine 0 2 "Specifiy how clients communicate on metered network connections: $($Usage)"
          WriteWordLine 0 0 ''
        }
        
      }
    }
    catch [System.Management.Automation.PropertyNotFoundException] 
    {
      WriteWordLine 0 0 ''
    }
  }
}
#endregion enumerating Client Policies

#region Security

Write-Verbose "$(Get-Date):   Collecting all administrative users"
WriteWordLine 2 0 'Administrative Users'
$Admins = Get-CMAdministrativeUser

WriteWordLine 0 1 'Enumerating administrative users:'

$AdminHashArray = @();

foreach ($Admin in $Admins) 
{
  switch ($Admin.AccountType)
  {
    0 { $AccountType = 'User' }
    1 { $AccountType = 'Group' }
    2 { $AccountType = 'Machine' } 
  } 
  
  $AdminRow = @{Name = $Admin.LogonName; 'Account Type' = $AccountType; 'Security Roles' = "$($Admin.RoleNames)"; 'Security Scopes' = "$($Admin.CategoryNames)"; Collections = "$($Admin.CollectionNames)";}
  $AdminHashArray += $AdminRow;
}

$Table = AddWordTable -Hashtable $AdminHashArray -Columns Name, 'Account Type', 'Security Roles', 'Security Scopes', Collections -Headers Name, 'Account Type', 'Security Roles', 'Security Scopes', Collections -Format -155 -AutoFit $wdAutoFitContent;
  
## Set first column format
SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

  
FindWordDocumentEnd
$Table = $Null

#endregion Security

#region enumerating all custom Security roles
Write-Verbose "$(Get-Date):   enumerating all custom build security roles"
WriteWordLine 2 0 'Custom Security Roles'
$SecurityRoles = Get-CMSecurityRole | Where-Object -FilterScript {-not $_.IsBuiltIn}
if (-not [string]::IsNullOrEmpty($SecurityRoles))
{
  $SRHashArray = @();
  
  WriteWordLine 0 1 'Enumerating all custom build security roles:'
  
  foreach ($SecurityRole in $SecurityRoles)
  {
    if ($SecurityRole.NumberOfAdmins -gt 0)
    {
      $Members = $(Get-CMAdministrativeUser | Where-Object -FilterScript {$_.Roles -ilike "$($SecurityRole.RoleID)"}).LogonName
    }
    $SRRow = @{Name = $SecurityRole.RoleName; Description = $SecurityRole.RoleDescription; 'Copied From' = $((Get-CMSecurityRole -Id $SecurityRole.CopiedFromID).RoleName); Members = "$Members"; 'Role ID' = $SecurityRole.RoleID;}
    $SRHashArray += $SRRow;
  }
  
  $Table = AddWordTable -Hashtable $SRHashArray -Columns Name, Description, 'Copied From', Members -Headers Name, Description, 'Copied From', Members -Format -155 -AutoFit $wdAutoFitContent;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
  FindWordDocumentEnd
  $Table = $Null
  
}
else
{
  WriteWordLine 0 1 'There are no custom built security roles.'
}

#endregion enumerating all custom Security roles

#region Used Accounts

Write-Verbose "$(Get-Date):   Enumerating all used accounts"
WriteWordLine 2 0 'Configured Accounts'
$Accounts = Get-CMAccount
WriteWordLine 0 1 'Enumerating all accounts used for specific tasks.'

$AccountsHashArray = @();

foreach ($Account in $Accounts)
{
  $AccountRow = @{'User Name'= $Account.UserName; 'Account Usage' = if ([string]::IsNullOrEmpty($Account.AccountUsage)) {'not assigned'} else {"$($Account.AccountUsage)"}; 'Site Code' = $Account.SiteCode};
  $AccountsHashArray += $AccountRow;
}

$Table = AddWordTable -Hashtable $AccountsHashArray -Columns 'User Name', 'Account Usage', 'Site Code' -Headers 'User Name', 'Account Usage', 'Site Code' -Format -155 -AutoFit $wdAutoFitContent;
  
## Set first column format
SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
FindWordDocumentEnd
$Table = $Null

#endregion Used Accounts

####
#region Assets and Compliance
####
Write-Verbose "$(Get-Date):   Done with Administration, next Assets and Compliance"
WriteWordLine 1 0 'Assets and Compliance'

#region enumerating all User Collections
WriteWordLine 2 0 'Summary of User Collections'
$UserCollections = Get-CMUserCollection
if ($ListAllInformation)
{
  $UserCollHashArray = @();
  
  foreach ($UserCollection in $UserCollections)
  {
    Write-Verbose "$(Get-Date):   Found User Collection: $($UserCollection.Name)"
    
    $UserCollRow = @{'Collection Name' = $UserCollection.Name; 'Collection ID' = $UserCollection.CollectionID; 'Member Count' = $UserCollection.MemberCount; 'Limited To' = "$($UserCollection.LimitToCollectionName) / $($UserCollection.LimitToCollectionID)";};
    $UserCollHashArray += $UserCollRow;
  }
  $Table = AddWordTable -Hashtable $UserCollHashArray -Columns 'Collection Name', 'Collection ID', 'Member Count', 'Limited To' -Headers 'Collection Name', 'Collection ID', 'Member Count', 'Limited To' -Format -155 -AutoFit $wdAutoFitContent;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15
  
  FindWordDocumentEnd
  $Table = $Null
}
else
{
  WriteWordLine 0 1 "There are $($UserCollections.count) User Collections." 
}

#endregion enumerating all User Collections

#region enumerating all Device Collections
WriteWordLine 2 0 'Summary of Device Collections'
$DeviceCollections = Get-CMDeviceCollection
if ($ListAllInformation)
{
  foreach ($DeviceCollection in $DeviceCollections)
  {
    Write-Verbose "$(Get-Date):   Found Device Collection: $($DeviceCollection.Name)"
    WriteWordLine 0 1 "Collection Name: $($DeviceCollection.Name)" -boldface $true
    WriteWordLine 0 1 "Collection ID: $($DeviceCollection.CollectionID)"
    WriteWordLine 0 1 "Total count of members: $($DeviceCollection.MemberCount)"
    WriteWordLine 0 1 "Limited to Device Collection: $($DeviceCollection.LimitToCollectionName) / $($DeviceCollection.LimitToCollectionID)"
    $CollSettings = Get-WmiObject -Class SMS_CollectionSettings -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.CollectionID -eq "$($DeviceCollection.CollectionID)"}
    if (-not [String]::IsNullOrEmpty($CollSettings))
        {
            $CollSettings = [wmi]$CollSettings.__PATH
            $ServiceWindows = $($CollSettings.ServiceWindows)
            if (-not [string]::IsNullOrEmpty($ServiceWindows))
                {
                    #$ServiceWindows
                    WriteWordLine 0 2 'Checking Maintenance Windows on Collection:' 
                    #$ServiceWindows = [wmi]$ServiceWindows.__PATH
                        
                    foreach ($ServiceWindow in $ServiceWindows)
                        {
                
                            $ScheduleString = Read-ScheduleToken
                            $startTime = $ScheduleString.TokenData.starttime
                            $startTime = Convert-NormalDateToConfigMgrDate -starttime $startTime
                            WriteWordLine 0 3 "Maintenance window name: $($ServiceWindow.Name)"
                            switch ($ServiceWindow.ServiceWindowType)
                                {
                                    0 {WriteWordLine 0 3 'This is a Task Sequence maintenance window'}
                                    1 {WriteWordLine 0 3 'This is a general maintenance window'}                        
                                }   
                            switch ($ServiceWindow.RecurrenceType)
                                {
                                    1 {WriteWordLine 0 3 "This maintenance window occurs only once on $($startTime) and lasts for $($ScheduleString.TokenData.HourDuration) hour(s) and $($ScheduleString.TokenData.MinuteDuration) minute(s)."}
                                    2 
                                        {
                                            if ($ScheduleString.TokenData.DaySpan -eq '1')
                                                {
                                                    $daily = 'daily'
                                                }
                                            else
                                                {
                                                    $daily = "every $($ScheduleString.TokenData.DaySpan) days"
                                                }
                        
                                            WriteWordLine 0 3 "This maintenance window occurs $($daily)."
                                        }
                                    3 
                                        {                                              
                                            WriteWordLine 0 3 "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofWeeks) week(s) on $(Convert-WeekDay $ScheduleString.TokenData.Day) and lasts $($ScheduleString.TokenData.HourDuration) hour(s) and $($ScheduleString.TokenData.MinuteDuration) minute(s) starting on $($startTime)."
                                        }
                                    4 
                                        {
                                            switch ($ScheduleString.TokenData.weekorder)
                                                {
                                                    0 {$order = 'last'}
                                                    1 {$order = 'first'}
                                                    2 {$order = 'second'}
                                                    3 {$order = 'third'}
                                                    4 {$order = 'fourth'}
                                                }
                                            WriteWordLine 0 3 "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofMonths) month(s) on every $($order) $(Convert-WeekDay $ScheduleString.TokenData.Day)"
                                        }

                                    5 
                                        {
                                            if ($ScheduleString.TokenData.MonthDay -eq '0')
                                                { 
                                                    $DayOfMonth = 'the last day of the month'
                                                }
                                            else
                                                {
                                                    $DayOfMonth = "day $($ScheduleString.TokenData.MonthDay)"
                                                }
                                            WriteWordLine 0 3 "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofMonths) month(s) on $($DayOfMonth)."
                                            WriteWordLine 0 3 "It lasts $($ScheduleString.TokenData.HourDuration) hours and $($ScheduleString.TokenData.MinuteDuration) minutes."
                                        }
                                }
                            switch ($ServiceWindow.IsEnabled)
                                {
                                    true {WriteWordLine 0 3 'The maintenance window is enabled'}
                                    false {WriteWordLine 0 3 'The maintenance window is disabled'}
                                }
                        }
                }
            else
                {
                    WriteWordLine 0 2 'No maintenance windows configured on this collection.'
                }
        }
        try {
            $CollVars = Get-CMDeviceCollectionVariable -CollectionId $DeviceCollection.CollectionID
            if ($CollVars) {
                $CollVarsHashArray = @();

                foreach ($CollVar in $CollVars)
                {
                  $CollVarRow = @{'Variable Name'= $CollVar.Name; 'Value' = $CollVar.Value; 'Hidden Value' = $CollVar.IsMasked};
                  $CollVarsHashArray += $CollVarRow;
                }

                $Table = AddWordTable -Hashtable $CollVarsHashArray -Columns 'Variable Name', 'Value', 'Hidden Value' -Headers 'Variable Name', 'Value', 'Hidden Value' -Format -155 -AutoFit $wdAutoFitContent
  
                ## Set first column format
                SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

                FindWordDocumentEnd
                $Table = $Null
            }
            else {
                WriteWordLine 0 1 'Enumerating device collection variables: No device collection variables configured!'
            }
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            WriteWordLine 0 0 ''
        }
        ### enumerating the Collection Membership Rules
        $QueryRules = $Null
        $DirectRules = $Null
        $IncludeRules = $Null
        $CollectionRules = $DeviceCollection.CollectionRules #just for Direct and Query
                    
        $Collection = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_Collection WHERE CollectionID = '$($DeviceCollection.CollectionID)'"
        [wmi]$Collection = $Collection.__PATH
                    
        $OtherCollectionRules = $Collection.CollectionRules
        try {
            $DirectRules = $CollectionRules | where {$_.ResourceID} -ErrorAction SilentlyContinue
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            WriteWordLine 0 0 ''
        }
        try {
            $QueryRules = $CollectionRules | where {$_.QueryExpression} -ErrorAction SilentlyContinue                            
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
            WriteWordLine 0 0 ''
        }
        try {
            $IncludeRules = $OtherCollectionRules | where {$_.IncludeCollectionID} -ErrorAction SilentlyContinue
        }
        catch [System.Management.Automation.PropertyNotFoundException] {
                    WriteWordLine 0 0 ''
        }

        if ($QueryRules) {            
            $QueryRulesHashArray = @();

            foreach ($QueryRule in $QueryRules)
            {
                $QueryRuleRow = @{'Query Name'= $QueryRule.RuleName; 'Query Expression' = $QueryRule.QueryExpression; 'Query ID' = $QueryRule.QueryId};
                $QueryRulesHashArray += $QueryRuleRow;
            }

            $Table = AddWordTable -Hashtable $QueryRulesHashArray -Columns 'Query Name', 'Query Expression', 'Query ID' -Headers 'Query Name', 'Query Expression', 'Query ID' -Format -155 -AutoFit $wdAutoFitContent;
  
            ## Set first column format
            SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
            FindWordDocumentEnd
            $Table = $Null
        }
        if ($DirectRules) {
            $DirectRulesHashArray = @();

            foreach ($DirectRule in $DirectRules)
            {
                $DirectRuleRow = @{'Resource Name'= $DirectRule.RuleName; 'Resource ID' = $DirectRule.ResourceId};
                $DirectRulesHashArray += $DirectRuleRow;
            }

            $Table = AddWordTable -Hashtable $DirectRulesHashArray -Columns 'Resource Name', 'Resource ID' -Headers 'Resource Name', 'Resource ID' -Format -155 -AutoFit $wdAutoFitContent;
  
            ## Set first column format
            SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

            FindWordDocumentEnd
            $Table = $Null           
        }
        else {
            WriteWordLine 0 1 'Enumerating device collection membership rules: No device collection direct membership rules configured!'
        }
        if ($IncludeRules) {
            $IncludeRulesHashArray = @();

            foreach ($IncludeRule in $IncludeRules)
            {
                $IncludeRuleRow = @{'Collection Name'= $IncludeRule.RuleName; 'Collection ID' = $IncludeRule.IncludeCollectionId};
                $IncludeRulesHashArray += $IncludeRuleRow;
            }

            $Table = AddWordTable -Hashtable $IncludeRulesHashArray -Columns 'Collection Name', 'Collection ID' -Headers 'Collection Name', 'Collection ID' -Format -155 -AutoFit $wdAutoFitContent;
  
            ## Set first column format
            SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
            $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
            $Table.AutoFitBehavior($wdAutoFitFixed)
  
            FindWordDocumentEnd
            $Table = $Null  
        }
        else {
            WriteWordLine 0 1 'Enumerating device collection membership rules: No device collection Include Collection membership rules configured!'
        }
    #move to the end of the current document
	Write-Verbose "$(Get-Date):   move to the end of the current document"
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	WriteWordLine 0 0 ''
    }
    }

else {
    WriteWordLine 0 1 "There are $($DeviceCollections.count) Device collections."
}

    Write-Verbose "$(Get-Date):   Working on Compliance Settings"
    WriteWordLine 2 0 'Compliance Settings'
    WriteWordLine 0 0 ''
    WriteWordLine 3 0 'Configuration Items'

    $CIs = Get-CMConfigurationItem
    WriteWordLine 0 1 'Enumerating Configuration Items:'

  $CIsHashArray = @();
  
  foreach ($CI in $CIs)
  {
    $CIRow = @{'Name' = $CI.LocalizedDisplayName; 'Last modified' = $CI.DateLastModified; 'Last modified by' = $CI.LastModifiedBy; 'CI ID' = $CI.CI_ID}
    $CIsHashArray += $CIRow
  }
  $Table = AddWordTable -Hashtable $CIsHashArray -Columns 'Name', 'Last modified', 'Last modified by', 'CI ID' -Headers 'Name', 'Last modified', 'Last modified by', 'CI ID' -Format -155 -AutoFit $wdAutoFitContent;
  
  ## Set first column format
  SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

  FindWordDocumentEnd
  $Table = $Null

    WriteWordLine 0 0 ''

    WriteWordLine 3 0 'Configuration Baselines'
    $CBs = Get-CMBaseline

    if ($CBs) {

      $CBsHashArray = @();
  
      foreach ($CB in $CBs)
      {
        $CBRow = @{'Name' = $CB.LocalizedDisplayName; 'Last modified' = $CB.DateLastModified; 'Last modified by' = $CB.LastModifiedBy; 'CI ID' = $CB.CI_ID}
        $CBsHashArray += $CBRow
      }
      $Table = AddWordTable -Hashtable $CBsHashArray -Columns 'Name', 'Last modified', 'Last modified by', 'CI ID' -Headers 'Name', 'Last modified', 'Last modified by', 'CI ID' -Format -155 -AutoFit $wdAutoFitContent;
  
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
      FindWordDocumentEnd
      $Table = $Null

      WriteWordLine 0 0 ''

    }
    else {
        WriteWordLine 0 1 'There are no Configuration Baselines configured.'
    }

    ### User Data and Profiles
    Write-Verbose "$(Get-Date):   Working on User Data and Profiles"
    WriteWordLine 3 0 'User Data and Profiles'
    $UserDataProfiles = Get-CMUserDataAndProfileConfigurationItem

    if (-not [string]::IsNullOrEmpty($UserDataProfiles)) {
      $UserDataProfilesHashArray = @();
  
      foreach ($UDP in $UserDataProfiles)
      {
        $UDPRow = @{'Name' = $UDP.LocalizedDisplayName; 'Last modified' = $UDP.DateLastModified; 'Last modified by' = $UDP.LastModifiedBy; 'CI ID' = $UDP.CI_ID}
        $UserDataProfilesHashArray += $UDPRow
      }
      $Table = AddWordTable -Hashtable $UserDataProfilesHashArray -Columns 'Name', 'Last modified', 'Last modified by', 'CI ID' -Headers 'Name', 'Last modified', 'Last modified by', 'CI ID' -Format -155 -AutoFit $wdAutoFitContent;
  
      ## Set first column format
      SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

      FindWordDocumentEnd
      $Table = $Null
      WriteWordLine 0 0 ''
    }
    else {
        WriteWordLine 0 1 'There are no User Data and Profile configurations configured.'
    }

        Write-Verbose "$(Get-Date):   Working on Endpoint Protection"
        WriteWordLine 2 0 'Endpoint Protection'
        if (-not ($(Get-CMEndpointProtectionPoint) -eq $Null))
            {
                WriteWordLine 3 0 'Antimalware Policies'
                $AntiMalwarePolicies = Get-CMAntimalwarePolicy
                if (-not [string]::IsNullOrEmpty($AntiMalwarePolicies))
                    {
                        foreach ($AntiMalwarePolicy in $AntiMalwarePolicies)
                            {
                                if ($AntiMalwarePolicy.Name -eq 'Default Client Antimalware Policy')
                                    {
                                        $AgentConfig = $AntiMalwarePolicy.AgentConfiguration
                                        WriteWordLine 0 1 "$($AntiMalwarePolicy.Name)" -boldface $true
                                        WriteWordLine 0 1 "Description: $($AntiMalwarePolicy.Description)"
                                        WriteWordLine 0 2 'Scheduled Scans' -boldface $true
                                        WriteWordLine 0 3 "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                        if ($AgentConfig.EnableScheduledScan)
                                            {
                                                switch ($AgentConfig.ScheduledScanType)
                                                    {
                                                        1 { $ScheduledScanType = 'Quick Scan' }
                                                        2 { $ScheduledScanType = 'Full Scan' }
                                                    }
                                                WriteWordLine 0 3 "Scan type: $($ScheduledScanType)"
                                                WriteWordLine 0 3 "Scan day: $(Convert-WeekDay $AgentConfig.ScheduledScanWeekDay)"
                                                WriteWordLine 0 3 "Scan time: $(Convert-Time -time $AgentConfig.ScheduledScanTime)"
                                                WriteWordLine 0 3 "Run a daily quick scan on client computers: $($AgentConfig.EnableQuickDailyScan)"
                                                WriteWordLine 0 3 "Daily quick scan schedule time: $(Convert-Time -time $AgentConfig.ScheduledScanQuickTime)"
                                                WriteWordLine 0 3 "Check for the latest definition updates before running a scan: $($AgentConfig.CheckLatestDefinition)"
                                                WriteWordLine 0 3 "Start a scheduled scan only when the computer is idle: $($AgentConfig.ScanWhenClientNotInUse)"
                                                WriteWordLine 0 3 "Force a scan of the selected scan type if client computer is offline during two or more scheduled scans: $($AgentConfig.EnableCatchupScan)"
                                                WriteWordLine 0 3 "Limit CPU usage during scans to (%): $($AgentConfig.LimitCPUUsage)"
                                            }
                                        WriteWordLine 0 0 ''
                                        WriteWordLine 0 2 'Scan settings' -boldface $true
                                        WriteWordLine 0 3 "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                                        WriteWordLine 0 3 "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                                        WriteWordLine 0 3 "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                                        WriteWordLine 0 3 "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                                        WriteWordLine 0 3 "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                                        WriteWordLine 0 3 'User control of scheduled scans: ' -nonewline
                                        switch ($AgentConfig.ScheduledScanUserControl)
                                            {
                                                0 { WriteWordLine 0 0 'No control' }
                                                1 { WriteWordLine 0 0 'Scan time only' }
                                                2 { WriteWordLine 0 0 'Full control' }
                                            }
                                        WriteWordLine 0 2 'Default Actions' -boldface $true
                                        WriteWordLine 0 3 'Severe threats: ' -nonewline
                                        switch ($AgentConfig.DefaultActionSevere)
                                            {
                                                0 { WriteWordLine 0 0 'Recommended' }
                                                2 { WriteWordLine 0 0 'Quarantine' }
                                                3 { WriteWordLine 0 0 'Remove' }
                                                6 { WriteWordLine 0 0 'Allow' }
                                            }
                                        WriteWordLine 0 3 'High threats: ' -nonewline
                                        switch ($AgentConfig.DefaultActionSevere)
                                            {
                                                0 { WriteWordLine 0 0 'Recommended' }
                                                2 { WriteWordLine 0 0 'Quarantine' }
                                                3 { WriteWordLine 0 0 'Remove' }
                                                6 { WriteWordLine 0 0 'Allow' }
                                            }
                                        WriteWordLine 0 3 'Medium threats: ' -nonewline
                                        switch ($AgentConfig.DefaultActionSevere)
                                            {
                                                0 { WriteWordLine 0 0 'Recommended' }
                                                2 { WriteWordLine 0 0 'Quarantine' }
                                                3 { WriteWordLine 0 0 'Remove' }
                                                6 { WriteWordLine 0 0 'Allow' }
                                            }
                                        WriteWordLine 0 3 'Low threats: ' -nonewline
                                        switch ($AgentConfig.DefaultActionSevere)
                                            {
                                                0 { WriteWordLine 0 0 'Recommended' }
                                                2 { WriteWordLine 0 0 'Quarantine' }
                                                3 { WriteWordLine 0 0 'Remove' }
                                                6 { WriteWordLine 0 0 'Allow' }
                                            }
                                        WriteWordLine 0 2 'Real-time protection' -boldface $true
                                        WriteWordLine 0 3 "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                        WriteWordLine 0 3 "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                                        WriteWordLine 0 3 'Scan system files: ' -nonewline
                                        switch ($AgentConfig.RealtimeScanOption)
                                            {
                                                0 { WriteWordLine 0 0 'Scan incoming and outgoing files' }
                                                1 { WriteWordLine 0 0 'Scan incoming files only' }
                                                2 { WriteWordLine 0 0 'Scan outgoing files only' }
                                            }
                                        WriteWordLine 0 2 'Exclusion settings' -boldface $true
                                        WriteWordLine 0 3 'Excluded files and folders: '
                                        foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths)
                                            {
                                                WriteWordLine 0 4 "$($ExcludedFileFolder)"
                                            }
                                        WriteWordLine 0 3 'Excluded file types: '
                                        foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes)
                                            {
                                                WriteWordLine 0 4 "$($ExcludedFileType)"
                                            }
                                        WriteWordLine 0 3 'Excluded processes: '
                                        foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses)
                                            {
                                                WriteWordLine 0 4 "$($ExcludedProcess)"
                                            }
                                        WriteWordLine 0 2 'Advanced' -boldface $true
                                        WriteWordLine 0 3 "Create a system restore point before computers are cleaned: $($AgentConfig.CreateSystemRestorePointBeforeClean)"
                                        WriteWordLine 0 3 "Disable the client user interface: $($AgentConfig.DisableClientUI)"
                                        WriteWordLine 0 3 "Show notifications messages on the client computer when the user needs to run a full scan, update definitions, or run Windows Defender Offline: $($AgentConfig.ShowNotificationMessages)"
                                        WriteWordLine 0 3 "Delete quarantined files after (days): $($AgentConfig.DeleteQuarantinedFilesPeriod)"
                                        WriteWordLine 0 3 "Allow users to configure the setting for quarantined file deletion: $($AgentConfig.AllowUserConfigQuarantinedFileDeletionPeriod)"
                                        WriteWordLine 0 3 "Allow users to exclude file and folders, file types and processes: $($AgentConfig.AllowUserAddExcludes)"
                                        WriteWordLine 0 3 "Allow all users to view the full History results: $($AgentConfig.AllowUserViewHistory)"
                                        WriteWordLine 0 3 "Enable reparse point scanning: $($AgentConfig.EnableReparsePointScanning)"
                                        WriteWordLine 0 3 "Randomize scheduled scan and definition update start time (within 30 minutes): $($AgentConfig.RandomizeScheduledScanStartTime)"
        
                                        WriteWordLine 0 2 'Threat overrides' -boldface $true
                                        if (-not [string]::IsNullOrEmpty($AgentConfig.ThreatName))
                                            {
                                                WriteWordLine 0 3 'Threat name and override action: Threats specified.'
                                            }
                                        WriteWordLine 0 2 'Microsoft Active Protection Service' -boldface $true
                                        WriteWordLine 0 3 'Microsoft Active Protection Service membership type: ' -nonewline
                                        switch ($AgentConfig.JoinSpyNet)
                                            {
                                                0 { WriteWordLine 0 0 'Do not join MAPS' }
                                                1 { WriteWordLine 0 0 'Basic membership' }
                                                2 { WriteWordLine 0 0 'Advanced membership' }
                                            }
                                        WriteWordLine 0 3 "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"

                                        WriteWordLine 0 2 'Definition Updates' -boldface $true
                                        WriteWordLine 0 3 "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                                        WriteWordLine 0 3 "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                                        WriteWordLine 0 3 "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                                        WriteWordLine 0 3 'Set sources and order for Endpoint Protection definition updates: '
                                        foreach ($Fallback in $AgentConfig.FallbackOrder)
                                            {
                                                WriteWordLine 0 3 "$($Fallback)"
                                            }
                                        WriteWordLine 0 3 "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                                        WriteWordLine 0 3 'If UNC file shares are selected as a definition update source, specify the UNC paths:' 
                                        foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources)
                                            {
                                                WriteWordLine 0 4 "$($UNCShare)"
                                            }
                                    }
                            else
                                {
                                    $AgentConfig_custom = $AntiMalwarePolicy.AgentConfigurations
                                    WriteWordLine 0 1 "$($AntiMalwarePolicy.Name)" -boldface $true
                                    WriteWordLine 0 1 "Description: $($AntiMalwarePolicy.Description)"
                                    foreach ($Agentconfig in $AgentConfig_custom)
                                        {
                                            switch ($AgentConfig.AgentID)
                                                {
                                                    201 
                                                        {
                                                            WriteWordLine 0 2 'Scheduled Scans' -boldface $true
                                                            WriteWordLine 0 2 "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                                            if ($AgentConfig.EnableScheduledScan)
                                                                {
                                                                    switch ($AgentConfig.ScheduledScanType)
                                                                        {
                                                                            1 { $ScheduledScanType = 'Quick Scan' }
                                                                            2 { $ScheduledScanType = 'Full Scan' }
                                                                        }
                                                                    WriteWordLine 0 3 "Scan type: $($ScheduledScanType)"
                                                                    WriteWordLine 0 3 "Scan day: $(Convert-WeekDay $AgentConfig.ScheduledScanWeekDay)"
                                                                    WriteWordLine 0 3 "Scan time: $(Convert-Time -time $AgentConfig.ScheduledScanTime)"
                                                                    WriteWordLine 0 3 "Run a daily quick scan on client computers: $($AgentConfig.EnableQuickDailyScan)"
                                                                    WriteWordLine 0 3 "Daily quick scan schedule time: $(Convert-Time -time $AgentConfig.ScheduledScanQuickTime)"
                                                                    WriteWordLine 0 3 "Check for the latest definition updates before running a scan: $($AgentConfig.CheckLatestDefinition)"
                                                                    WriteWordLine 0 3 "Start a scheduled scan only when the computer is idle: $($AgentConfig.ScanWhenClientNotInUse)"
                                                                    WriteWordLine 0 3 "Force a scan of the selected scan type if client computer is offline during two or more scheduled scans: $($AgentConfig.EnableCatchupScan)"
                                                                    WriteWordLine 0 3 "Limit CPU usage during scans to (%): $($AgentConfig.LimitCPUUsage)"
                                                                }
                                                        }
                                                    202
                                                        {
                                                            WriteWordLine 0 2 'Default Actions' -boldface $true
                                                            WriteWordLine 0 3 'Severe threats: ' -nonewline
                                                            switch ($AgentConfig.DefaultActionSevere)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Recommended' }
                                                                    2 { WriteWordLine 0 0 'Quarantine' }
                                                                    3 { WriteWordLine 0 0 'Remove' }
                                                                    6 { WriteWordLine 0 0 'Allow' }
                                                                }
                                                            WriteWordLine 0 3 'High threats: ' -nonewline
                                                            switch ($AgentConfig.DefaultActionSevere)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Recommended' }
                                                                    2 { WriteWordLine 0 0 'Quarantine' }
                                                                    3 { WriteWordLine 0 0 'Remove' }
                                                                    6 { WriteWordLine 0 0 'Allow' }
                                                                }
                                                            WriteWordLine 0 3 'Medium threats: ' -nonewline
                                                            switch ($AgentConfig.DefaultActionSevere)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Recommended' }
                                                                    2 { WriteWordLine 0 0 'Quarantine' }
                                                                    3 { WriteWordLine 0 0 'Remove' }
                                                                    6 { WriteWordLine 0 0 'Allow' }
                                                                }
                                                            WriteWordLine 0 3 'Low threats: ' -nonewline
                                                            switch ($AgentConfig.DefaultActionSevere)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Recommended' }
                                                                    2 { WriteWordLine 0 0 'Quarantine' }
                                                                    3 { WriteWordLine 0 0 'Remove' }
                                                                    6 { WriteWordLine 0 0 'Allow' }
                                                                }                                           
                                                        }
                                                    203
                                                        {
                                                            WriteWordLine 0 2 'Exclusion settings' -boldface $true
                                                            WriteWordLine 0 3 'Excluded files and folders: '
                                                            foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths)
                                                                {
                                                                    WriteWordLine 0 4 "$($ExcludedFileFolder)"
                                                                }
                                                            WriteWordLine 0 3 'Excluded file types: '
                                                            foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes)
                                                                {
                                                                    WriteWordLine 0 4 "$($ExcludedFileType)"
                                                                }
                                                            WriteWordLine 0 3 'Excluded processes: '
                                                            foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses)
                                                                {
                                                                    WriteWordLine 0 4 "$($ExcludedProcess)"
                                                                }                                            
                                                        }
                                                    204
                                                        {
                                                            WriteWordLine 0 2 'Real-time protection' -boldface $true
                                                            WriteWordLine 0 3 "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                                            WriteWordLine 0 3 "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                                                            WriteWordLine 0 3 'Scan system files: ' -nonewline
                                                            switch ($AgentConfig.RealtimeScanOption)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Scan incoming and outgoing files' }
                                                                    1 { WriteWordLine 0 0 'Scan incoming files only' }
                                                                    2 { WriteWordLine 0 0 'Scan outgoing files only' }
                                                                }                                            
                                                        }
                                                    205
                                                        {
                                                            WriteWordLine 0 2 'Advanced' -boldface $true
                                                            WriteWordLine 0 3 "Create a system restore point before computers are cleaned: $($AgentConfig.CreateSystemRestorePointBeforeClean)"
                                                            WriteWordLine 0 3 "Disable the client user interface: $($AgentConfig.DisableClientUI)"
                                                            WriteWordLine 0 3 "Show notifications messages on the client computer when the user needs to run a full scan, update definitions, or run Windows Defender Offline: $($AgentConfig.ShowNotificationMessages)"
                                                            WriteWordLine 0 3 "Delete quarantined files after (days): $($AgentConfig.DeleteQuarantinedFilesPeriod)"
                                                            WriteWordLine 0 3 "Allow users to configure the setting for quarantined file deletion: $($AgentConfig.AllowUserConfigQuarantinedFileDeletionPeriod)"
                                                            WriteWordLine 0 3 "Allow users to exclude file and folders, file types and processes: $($AgentConfig.AllowUserAddExcludes)"
                                                            WriteWordLine 0 3 "Allow all users to view the full History results: $($AgentConfig.AllowUserViewHistory)"
                                                            WriteWordLine 0 3 "Enable reparse point scanning: $($AgentConfig.EnableReparsePointScanning)"
                                                            WriteWordLine 0 3 "Randomize scheduled scan and definition update start time (within 30 minutes): $($AgentConfig.RandomizeScheduledScanStartTime)"                                            
                                                        }
                                                    206
                                                        {
                                            
                                                        }
                                                    207
                                                        {
                                                            WriteWordLine 0 2 'Microsoft Active Protection Service' -boldface $true
                                                            WriteWordLine 0 3 'Microsoft Active Protection Service membership type: ' -nonewline
                                                            switch ($AgentConfig.JoinSpyNet)
                                                                {
                                                                    0 { WriteWordLine 0 0 'Do not join MAPS' }
                                                                    1 { WriteWordLine 0 0 'Basic membership' }
                                                                    2 { WriteWordLine 0 0 'Advanced membership' }
                                                                }
                                                            WriteWordLine 0 3 "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"                                            
                                                        }
                                                    208
                                                        {
                                                            WriteWordLine 0 2 'Definition Updates' -boldface $true
                                                            WriteWordLine 0 3 "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                                                            WriteWordLine 0 3 "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                                                            WriteWordLine 0 3 "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                                                            WriteWordLine 0 3 'Set sources and order for Endpoint Protection definition updates: '
                                                            foreach ($Fallback in $AgentConfig.FallbackOrder)
                                                                {
                                                                    WriteWordLine 0 4 "$($Fallback)"
                                                                }
                                                            WriteWordLine 0 3 "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                                                            WriteWordLine 0 3 'If UNC file shares are selected as a definition update source, specify the UNC paths:' 
                                                            foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources)
                                                                {
                                                                    WriteWordLine 0 4 "$($UNCShare)"
                                                                }
                                                        }
                                                    209
                                                        {
                                                            WriteWordLine 0 2 'Scan settings' -boldface $true
                                                            WriteWordLine 0 3 "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                                                            WriteWordLine 0 3 "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                                                            WriteWordLine 0 3 "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                                                            WriteWordLine 0 3 "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                                                            WriteWordLine 0 3 "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                                                            WriteWordLine 0 3 'User control of scheduled scans: ' -nonewline
                                                            switch ($AgentConfig.ScheduledScanUserControl)
                                                                {
                                                                    0 { WriteWordLine 0 0 'No control' }
                                                                    1 { WriteWordLine 0 0 'Scan time only' }
                                                                    2 { WriteWordLine 0 0 'Full control' }
                                                                }
                                                        }
                                                }
                                        }
                                }
                        }
                    }
                else
                    {
                        WriteWordLine 0 1 'There are no Anti Malware Policies configured.'
                    }
            }
        else
            {
                WriteWordLine 0 1 'There is no Endpoint Protection Point enabled.'
            }

        WriteWordLine 0 0 ''

        Write-Verbose "$(Get-Date):   Working on Windows Firewall Policies"
        WriteWordLine 3 0 'Windows Firewall Policies'

        $FirewallPolicies = Get-CMWindowsFirewallPolicy
        if (-not [string]::IsNullOrEmpty($FirewallPolicies)) {

          $FirewallPolsHashArray = @()
  
          foreach ($FWP in $FirewallPolicies)
          {
            $FWPRow = @{'Name' = $FWP.LocalizedDisplayName; 'Last modified' = $FWP.DateLastModified; 'Last modified by' = $FWP.LastModifiedBy; 'CI ID' = $FWP.CI_ID}
            $FirewallPolsHashArray += $FWPRow
          }
          $Table = AddWordTable -Hashtable $FirewallPolsHashArray -Columns 'Name', 'Last modified', 'Last modified by', 'CI ID' -Headers 'Name', 'Last modified', 'Last modified by', 'CI ID' -Format -155 -AutoFit $wdAutoFitFixed;
  
          ## Set first column format
          SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
  
          ## IB - set column widths without recursion
          $Table.Columns.Item(1).Width = 100;
          $Table.Columns.Item(2).Width = 100;
          $Table.Columns.Item(3).Width = 100;
          $Table.Columns.Item(4).Width = 150;
          $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
          $Table.AutoFitBehavior($wdAutoFitFixed)
          FindWordDocumentEnd
          $Table = $Null
          WriteWordLine 0 0 ''           

        }
        else {
            WriteWordLine 0 1 'There are no Windows Firewall policies configured.'
        }
        
        #####
        ##### finished with Assets and Compliance, moving on to Software Library
        #####
        Write-Verbose "$(Get-Date):   Finished with Assets and Compliance."



#endregion Assets and Compliance

if ($Software)
    {
        Write-Verbose "$(Get-Date):   moving on to Software Library"
        WriteWordLine 1 0 'Software Library'

##### Applications
        
        WriteWordLine 2 0 'Applications'
        WriteWordLine 0 0 ''
        $Applications = Get-WmiObject -Class sms_applicationlatest -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider
        if ($ListAllInformation)
            {
                if (-not [string]::IsNullOrEmpty($Applications)) {
                    WriteWordLine 0 1 'The following Applications are configured in this site:'
                    

                    foreach ($App in $Applications) {
                        Write-Verbose 'Getting specific WMI instance for this App'
                        [wmi]$App = $App.__PATH
                        Write-Verbose "$(Get-Date):   Found App: $($App.LocalizedDisplayName)"
                        WriteWordLine 0 2 "$($App.LocalizedDisplayName)" -boldface $true
                        WriteWordLine 0 3 "Created by: $($App.CreatedBy)"
                        WriteWordLine 0 3 "Date created: $($App.DateCreated)"
                        $DTs = Get-CMDeploymentType -ApplicationName $App.LocalizedDisplayName
                        if (-not [string]::IsNullOrEmpty($DTs)) {
                            $DTsHashArray = @()
  
                            foreach ($DT in $DTs) {
                                $xmlDT = [xml]$DT.SDMPackageXML
                                $DTRow = @{'Deployment Type Name' = $DT.LocalizedDisplayName; 'Technology' = $DT.Technology; 'Commandline' = if (-not ($DT.Technology -like 'AppV*')){ $xmlDT.AppMgmtDigest.DeploymentType.Installer.CustomData.InstallCommandLine } }
                                $DTsHashArray += $DTRow
                            }
                            $Table = AddWordTable -Hashtable $DTsHashArray -Columns 'Deployment Type Name', 'Technology', 'Commandline' -Headers 'Deployment Type Name', 'Technology', 'Commandline' -Format -155 -AutoFit $wdAutoFitContent;
  
                            ## Set first column format
                            SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
 
                            FindWordDocumentEnd
                            $Table = $Null
                            WriteWordLine 0 0 ''
                        }
                        else {
                            WriteWordLine 0 3 'There are no Deployment Types configured for this Application.'
                        }
                    }
                }
                else {
                    WriteWordLine 0 1 'There are no Applications configured in this site.'
                }
            }
        elseif ($Applications) {
            WriteWordLine 0 1 "There are $($Applications.count) applications configured."
        }
        else {
            WriteWordLine 0 1 'There are no Applications configured in this site.'
        }
##### Packages
        
        WriteWordLine 2 0 'Packages'
        WriteWordLine 0 0 ''
        $Packages = Get-CMPackage
        if ($ListAllInformation)
            {
                if (-not [string]::IsNullOrEmpty($Packages))
                    {
                        WriteWordLine 0 1 'The following Packages are configured in this site:'
                        foreach ($Package in $Packages) {
                            WriteWordLine 0 0 ''
                            WriteWordLine 0 2 "$($Package.Name)" -boldface $true
                            WriteWordLine 0 3 "Description: $($Package.Description)"
                            WriteWordLine 0 3 "PackageID: $($Package.PackageID)"
                            $Programs = Get-WmiObject -Class SMS_Program -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "PackageID = '$($Package.PackageID)'"
                            if (-not [string]::IsNullOrEmpty($Programs))
                                {
                                    WriteWordLine 0 3 'The Package has the following Programs configured:'
                                    foreach ($Program in $Programs)
                                        {
                                            WriteWordLine 0 4 "Program Name: $($Program.ProgramName)" -boldface $true
                                            WriteWordLine 0 4 "Command Line: $($Program.CommandLine)"
                                            if ($Program.ProgramFlags -band 0x00000001)
                                                {
                                                    WriteWordLine 0 4 "`'Allow this program to be installed from the Install Package task sequence without being deployed`' enabled."
                                                }
                                            if ($Program.ProgramFlags -band 0x00000002)
                                                {
                                                    WriteWordLine 0 4 "`'The task sequence shows a custom progress user interface message.`' enabled."
                                                }
                                            if ($Program.ProgramFlags -band 0x00000010)
                                                {
                                                    WriteWordLine 0 4 'This is a default program.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00000020)
                                                {
                                                    WriteWordLine 0 4 'Disables MOM alerts while the program runs.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00000040)
                                                {
                                                    WriteWordLine 0 4 'Generates MOM alert if the program fails.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00000080)
                                                {
                                                    WriteWordLine 0 4 "This program's immediate dependent should always be run."
                                                }
                                            if ($Program.ProgramFlags -band 0x00000100)
                                                {
                                                    WriteWordLine 0 4 'A device program. The program is not offered to desktop clients.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00000400)
                                                {
                                                    WriteWordLine 0 4 'The countdown dialog is not displayed.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00001000)
                                                {
                                                    WriteWordLine 0 4 'The program is disabled.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00002000)
                                                {
                                                    WriteWordLine 0 4 'The program requires no user interaction.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00004000)
                                                {
                                                    WriteWordLine 0 4 'The program can run only when a user is logged on.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00008000)
                                                {
                                                    WriteWordLine 0 4 'The program must be run as the local Administrator account.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00010000)
                                                {
                                                    WriteWordLine 0 4 'The program must be run by every user for whom it is valid. Valid only for mandatory jobs.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00020000)
                                                {
                                                    WriteWordLine 0 4 'The program is run only when no user is logged on.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00040000)
                                                {
                                                    WriteWordLine 0 4 'The program will restart the computer.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00080000)
                                                {
                                                    WriteWordLine 0 4 'Configuration Manager restarts the computer when the program has finished running successfully.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00100000)
                                                {
                                                    WriteWordLine 0 4 'Use a UNC path (no drive letter) to access the distribution point.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00200000)
                                                {
                                                    WriteWordLine 0 4 'Persists the connection to the drive specified in the DriveLetter property. The USEUNCPATH bit flag must not be set.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00400000)
                                                {
                                                    WriteWordLine 0 4 'Run the program as a minimized window.'
                                                }
                                            if ($Program.ProgramFlags -band 0x00800000)
                                                {
                                                    WriteWordLine 0 4 'Run the program as a maximized window.'
                                                }
                                            if ($Program.ProgramFlags -band 0x01000000)
                                                {
                                                    WriteWordLine 0 4 'Hide the program window.'
                                                }
                                            if ($Program.ProgramFlags -band 0x02000000)
                                                {
                                                    WriteWordLine 0 4 'Logoff user when program completes successfully.'
                                                }
                                            if ($Program.ProgramFlags -band 0x08000000)
                                                {
                                                    WriteWordLine 0 4 'Override check for platform support.'
                                                }
                                            if ($Program.ProgramFlags -band 0x20000000)
                                                {
                                                    WriteWordLine 0 4 'Run uninstall from the registry key when the advertisement expires.'   
                                                }   
                                        }                             
                                }
                            else
                                {
                                    WriteWordLine 0 4 'The Package has no Programs configured.'
                                }                       
                        }
                    }
                else
                    {
                        WriteWordLine 0 1 'There are no Packages configured in this site.'
                    }
            }
        elseif ($Packages)
            {
                WriteWordLine 0 1 "There are $($Packages.count) packages configured."
            }
            else
                {
                    WriteWordLine 0 1 'There are no packages configured.'
                }
##### Driver Packages

    WriteWordLine 2 0 'Driver Packages'
    WriteWordLine 0 0 ''
    $DriverPackages = Get-CMDriverPackage
    if ($ListAllInformation)
        {
            if (-not [string]::IsNullOrEmpty($DriverPackages))
                    {
                        WriteWordLine 0 1 'The following Driver Packages are configured in your site:'
                        foreach ($DriverPackage in $DriverPackages)
                            {
                                WriteWordLine 0 0 ''
                                WriteWordLine 0 2 "Name: $($DriverPackage.Name)" -boldface $true
                                if ($DriverPackage.Description)
                                    {
                                        WriteWordLine 0 2 "Description: $($DriverPackage.Description)"
                                    }
                                WriteWordLine 0 2 "PackageID: $($DriverPackage.PackageID)"
                                WriteWordLine 0 2 "Source path: $($DriverPackage.PkgSourcePath)"
                                WriteWordLine 0 2 'This package consists of the following Drivers:'
                                $Drivers = Get-CMDriver -DriverPackageId "$($DriverPackage.PackageID)"
                                foreach ($Driver in $Drivers)
                                    {
                                        WriteWordLine 0 0 ''
                                        WriteWordLine 0 3 "Driver Name: $($Driver.LocalizedDisplayName)" -boldface $true
                                        WriteWordLine 0 3 "Manufacturer: $($Driver.DriverProvider)"
                                        WriteWordLine 0 3 "Source path: $($Driver.ContentSourcePath)"
                                        WriteWordLine 0 3 "INF File: $($Driver.DriverINFFile)"
                                    }
                                WriteWordLine 0 3 ''
                            }
                    }
                else
                    {
                        WriteWordLine 0 1 'There are no Driver Packages configured in this site.'
                    }
        }
    else
        {
            if (-not [string]::IsNullOrEmpty($DriverPackages))
                {
                    WriteWordLine 0 1 "There are $($DriverPackages.count) Driver Packages configured."                    
                }
            else
                {
                    WriteWordLine 0 1 'There are no Driver Packages configured in this site.'
                }
        }
##### Operating System Installers

    WriteWordLine 2 0 'Operating System Installers'
    WriteWordLine 0 0 ''
    $OSInstallers = Get-CMOperatingSystemInstaller
    if (-not [string]::IsNullOrEmpty($OSImages))
        {
            WriteWordLine 0 1 'The following OS Installers are imported into this environment:'
            foreach ($OSInstaller in $OSInstallers)
                {
                    WriteWordLine 0 2 "Name: $($OSInstaller.Name)" -boldface $true
                    if ($OSInstaller.Description)
                            {
                                WriteWordLine 0 2 "Description: $($OSInstaller.Description)"
                            }
                    WriteWordLine 0 2 "Package ID: $($OSInstaller.PackageID)"
                    WriteWordLine 0 2 "Source Path: $($OSInstaller.PkgSourcePath)"
                }
        }
    else
        {
            WriteWordLine 0 1 'There are no OS Installers imported into this environment.'
        }

####
####
#### Boot Images
    
WriteWordLine 2 0 'Boot Images'
WriteWordLine 0 0 ''
$BootImages = Get-CMBootImage
if (-not [string]::IsNullOrEmpty($BootImages))
    {
        WriteWordLine 0 1 'The following Boot Images are imported into this environment:'
        WriteWordLine 0 0 ''
        foreach ($BootImage in $BootImages)
            {
                WriteWordLine 0 2 "$($BootImage.Name)" -boldface $true
                if ($BootImage.Description)
                    {
                        WriteWordLine 0 2 "Description: $($BootImage.Description)"
                    }
                WriteWordLine 0 2 "Source Path: $($BootImage.PkgSourcePath)"
                WriteWordLine 0 2 "Package ID: $($BootImage.PackageID)"
                WriteWordLine 0 2 'Architecture: ' -nonewline
                switch ($BootImage.Architecture)
                    {
                        0 { WriteWordLine 0 0 'x86' }
                        9 { WriteWordLine 0 0 'x64' }
                    }
                if ($BootImage.BackgroundBitmapPath)
                    {
                        WriteWordLine 0 2 "Custom Background: $($BootImage.BackgroundBitmapPath)"
                    }
                Switch ($BootImage.EnableLabShell)
                    {
                        True { WriteWordLine 0 2 'Command line support is enabled' }
                        False { WriteWordLine 0 2 'Command line support is not enabled' }
                    }
                WriteWordLine 0 2 'The following drivers are imported into this WinPE'
                if (-not [string]::IsNullOrEmpty($BootImage.ReferencedDrivers))
                    {
                        $ImportedDriverIDs = ($BootImage.ReferencedDrivers).ID | Out-Null
                        foreach ($ImportedDriverID in $ImportedDriverIDs)
                            {
                                $ImportedDriver = Get-CMDriver -ID $ImportedDriverID
                                WriteWordLine 0 3 "Name: $($ImportedDriver.LocalizedDisplayName)" -boldface $true
                                WriteWordLine 0 3 "Inf File: $($ImportedDriver.DriverINFFile)"
                                WriteWordLine 0 3 "Driver Class: $($ImportedDriver.DriverClass)"
                                WriteWordLine 0 0 ''
                            }
                    }
                else
                    {
                        WriteWordLine 0 3 'There are no drivers imported into the Boot Image.'
                    }
            if (-not [string]::IsNullOrEmpty($BootImage.OptionalComponents))
                {
                    $Component = $Null
                    WriteWordLine 0 3 'The following Optional Components are added to this Boot Image:'
                    foreach ($Component in $BootImage.OptionalComponents)
                        {
                            switch ($Component)
                                {
                                    {($_ -eq '1') -or ($_ -eq '27')} { WriteWordLine 0 4 'WinPE-DismCmdlets' }                                    {($_ -eq '2') -or ($_ -eq '28')} { WriteWordLine 0 4 'WinPE-Dot3Svc' }                                    {($_ -eq '3') -or ($_ -eq '29')} { WriteWordLine 0 4 'WinPE-EnhancedStorage' }                                    {($_ -eq '4') -or ($_ -eq '30')} { WriteWordLine 0 4 'WinPE-FMAPI' }                                    {($_ -eq '5') -or ($_ -eq '31')} { WriteWordLine 0 4 'WinPE-FontSupport-JA-JP' }                                    {($_ -eq '6') -or ($_ -eq '32')} { WriteWordLine 0 4 'WinPE-FontSupport-KO-KR' }                                    {($_ -eq '7') -or ($_ -eq '33')} { WriteWordLine 0 4 'WinPE-FontSupport-ZH-CN' }                                    {($_ -eq '8') -or ($_ -eq '34')} { WriteWordLine 0 4 'WinPE-FontSupport-ZH-HK' }                                    {($_ -eq '9') -or ($_ -eq '35')} { WriteWordLine 0 4 'WinPE-FontSupport-ZH-TW' }                                    {($_ -eq '10') -or ($_ -eq '36')} { WriteWordLine 0 4 'WinPE-HTA' }                                    {($_ -eq '11') -or ($_ -eq '37')} { WriteWordLine 0 4 'WinPE-StorageWMI' }                                    {($_ -eq '12') -or ($_ -eq '38')} { WriteWordLine 0 4 'WinPE-LegacySetup' }                                    {($_ -eq '13') -or ($_ -eq '39')} { WriteWordLine 0 4 'WinPE-MDAC' }                                    {($_ -eq '14') -or ($_ -eq '40')} { WriteWordLine 0 4 'WinPE-NetFx4' }                                    {($_ -eq '15') -or ($_ -eq '41')} { WriteWordLine 0 4 'WinPE-PowerShell3' }                                    {($_ -eq '16') -or ($_ -eq '42')} { WriteWordLine 0 4 'WinPE-PPPoE' }                                    {($_ -eq '17') -or ($_ -eq '43')} { WriteWordLine 0 4 'WinPE-RNDIS' }                                    {($_ -eq '18') -or ($_ -eq '44')} { WriteWordLine 0 4 'WinPE-Scripting' }                                    {($_ -eq '19') -or ($_ -eq '45')} { WriteWordLine 0 4 'WinPE-SecureStartup' }                                    {($_ -eq '20') -or ($_ -eq '46')} { WriteWordLine 0 4 'WinPE-Setup' }                                    {($_ -eq '21') -or ($_ -eq '47')} { WriteWordLine 0 4 'WinPE-Setup-Client' }                                    {($_ -eq '22') -or ($_ -eq '48')} { WriteWordLine 0 4 'WinPE-Setup-Server' }                                    #{($_ -eq "23") -or ($_ -eq "49")} { WriteWordLine 0 4 "Not applicable" }                                    {($_ -eq '24') -or ($_ -eq '50')} { WriteWordLine 0 4 'WinPE-WDS-Tools' }                                    {($_ -eq '25') -or ($_ -eq '51')} { WriteWordLine 0 4 'WinPE-WinReCfg' }                                    {($_ -eq '26') -or ($_ -eq '52')} { WriteWordLine 0 4 'WinPE-WMI' }
                                } 
                            $Component = $Null    
                        }
                    }
                WriteWordLine 0 0 ''

            }
    }
else
    {
        WriteWordLine 0 1 'There are no Boot Images present in this environment.'
    }

####
####
#### Task Sequences
Write-Verbose "$(Get-Date):   Enumerating Task Sequences"
WriteWordLine 2 0 'Task Sequences'
WriteWordLine 0 0 ''

$TaskSequences = Get-CMTaskSequence
Write-Verbose "$(Get-Date):   working on $($TaskSequences.count) Task Sequences"
if ($ListAllInformation)
    {
        if (-not [string]::IsNullOrEmpty($TaskSequences))
            {
                foreach ($TaskSequence in $TaskSequences)
                    {
                        WriteWordLine 0 1 "Task Sequence name: $($TaskSequence.Name)" -boldface $true
                        WriteWordLine 0 1 "Package ID: $($TaskSequence.PackageID)"
                        if ($TaskSequence.BootImageID)
                            {
                                WriteWordLine 0 2 "Boot Image referenced in this Task Sequence: $((Get-CMBootImage -Id $TaskSequence.BootImageID -ErrorAction SilentlyContinue ).Name)"
                            }
        
                        $Sequence = $Null
                        [xml]$Sequence = $TaskSequence.Sequence
                        try
                            {
                                foreach ($Group in $Sequence.sequence.group)
                                    {
                                        WriteWordLine 0 1 "Group name: $($Group.Name)" -boldface $true
                                        if (-not [string]::IsNullOrEmpty($Group.Description))
                                            {
                                                WriteWordLine 0 1 "Description: $($Group.Description)"
                                            }
                                        WriteWordLine 0 1 'This Group has the following steps configured.'
                                        foreach ($Step in $Group.Step)
                                            {
                                                WriteWordLine 0 3 "$($Step.Name)" -boldface $true
                                                if (-not [string]::IsNullOrEmpty($Step.Description))
                                                    {
                                                        WriteWordLine 0 4 "$($Step.Description)"
                                                    }
                                                WriteWordLine 0 4 "$($Step.Action)"
                                                try 
                                                    {
                                                        if (-not [string]::IsNullOrEmpty($Step.disable))
                                                                {
                                                                    WriteWordLine 0 4 'This step is disabled.'
                                                                }
                                                    }   
                                                catch [System.Management.Automation.PropertyNotFoundException] 
                                                    {
                                                        WriteWordLine 0 4 'This step is enabled'
                                                    }
                                                WriteWordLine 0 0 ''
                                            }

                                    }
                            }
                        catch [System.Management.Automation.PropertyNotFoundException]
                            {
                                WriteWordLine 0 0 ''
                            }
                        try 
                            {
                                foreach ($Step in $Sequence.sequence.step)
                                    {
                                        WriteWordLine 0 3 "$($Step.Name)" -boldface $true
                                        if (-not [string]::IsNullOrEmpty($Step.Description))
                                            {
                                                WriteWordLine 0 4 "$($Step.Description)"
                                            }
                                        WriteWordLine 0 4 "$($Step.Action)"
                                        try 
                                            {
                                                if (-not [string]::IsNullOrEmpty($Step.disable))
                                                        {
                                                            WriteWordLine 0 4 'This step is disabled.'
                                                        }
                                            }   
                                        catch [System.Management.Automation.PropertyNotFoundException] 
                                            {
                                                WriteWordLine 0 4 'This step is enabled'
                                            }
                                        WriteWordLine 0 0 ''
                                    }
                            }
                        catch [System.Management.Automation.PropertyNotFoundException]
                            {
                                WriteWordLine 0 0 ''
                            }
                        #>
                        WriteWordLine 0 0 ''
                        WriteWordLine 0 0 '----------------------------------------------'
                    }
            }
        else
            {
                WriteWordLine 0 1 'There are no Task Sequences present in this environment.'
            }
    }
else
    {
        if (-not [string]::IsNullOrEmpty($TaskSequences))
            {
                WriteWordLine 0 1 'The following Task Sequences are configured:'
                foreach ($TaskSequence in $TaskSequences)
                    {
                        WriteWordLine 0 2 "$($TaskSequence.Name)"
                    }
            }
        else
            {
                WriteWordLine 0 1 'There are no Task Sequences present in this environment.'
            }
    }

} #End Software

#endregion Site Configuration report

Set-Location -Path $LocationBeforeExecution
$Script:ScriptInformation = $Null

###REPLACE BEFORE THIS SECTION WITH YOUR SCRIPT###
#endregion

###REPLACE BEFORE THIS SECTION WITH YOUR SCRIPT###
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script###
$AbstractTitle = "Template Script Report"
$SubjectTitle = "Sample Template Script Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
$ErrorActionPreference = $SaveEAPreference
#endregion