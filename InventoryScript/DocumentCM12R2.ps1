<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft System Center 2012 Configuration Manager R2 hierarchy using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Microsoft System Center 2012 Configuration Manager R2 hierarchy using Microsoft Word and PowerShell.
	Creates a Word document named after the customer's name.
	Document includes a Cover Page, Table of Contents and Footer. File will be saved in the folder from where the script is executed. Does not work with UNC paths.
.PARAMETER SMSProvider
    FQDN of a SMS Provider in this hierarchy. 
    This parameter is mandatory!
    This parameter has an alias of MP.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER ListAllInformation
    This parameter is a switch. If you use this switch, then you will get a lot more information regarding packages, applications, user and device collections.
    This parameter has an alias of LA.
.EXAMPLE
	PS C:\PSScript > .\DocumentCM12R2.ps1 -SMSProvider CM12.do.local
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="David O'Brien" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="David O'Brien"
	$env:username = adobrien

	David O'Brien for the Company Name.
	Motion for the Cover Page format.
	adobrien for the User Name.
    CM12.do.local for the SMS Provder.
.EXAMPLE
	PS C:\PSScript > .\DocumentCM12R2.ps1 -SMSProvider CM12.do.local -ListAllInformation
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="David O'Brien" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="David O'Brien"
	$env:username = adobrien

	David O'Brien for the Company Name.
	Motion for the Cover Page format.
	adobrien for the User Name.
    CM12.do.local for the SMS Provider.
    Will give you more information, because of the ListAllInformation switch.
.EXAMPLE
	PS C:\PSScript > .\DocumentCM12.ps1 -SMSProvider CM12.do.local -verbose
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="David O'Brien" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="David O'Brien"
	$env:username = adobrien

	David O'Brien for the Company Name.
	Motion for the Cover Page format.
	adobrien for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript > .\DocumentCM12R2.ps1 -SMSProvider CM12.do.local -CompanyName "David's company" -CoverPage "Motion" -UserName "David O'Brien"

	Will use:
		David's company for the Company Name.
		Motion for the Cover Page format.
		David O'Brien for the User Name.
        CM12.do.local for the SMS Provider.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.LINK
	http://www.david-obrien.net
.NOTES
	NAME: DocumentCM12.ps1
	VERSION: 1.0
	AUTHOR: David O'Brien (former script by Carl Webster! www.carlwebster.com ! with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: January, 28, 2014
    REQUIREMENTS: 
                    Local installation of Configuration Manager 2012 R2 console.
                    at least Powershell 4.0
                    at least Read-Only Analyst Permission in ConfigMgr site
                    Local installation of Microsoft Winword
    Change history:
        19.06.2013: added error checks, bitness of powershell process, has powershell module been loaded? (version 0.1)
        09.08.2013: corrected some spelling and grammar mistakes, added more error and sanity checks, added DP disk info  (version 0.2)
        11.08.2013: replaced Parameter "ManagementPoint" with "SMSProvider", replaced Word indenting with with bullet points
        13.08.2013: removed Parameter "SiteCode", evaluating it inside the script, parameter validation for SMS Provider, checking Site version running against and resulting bitness for cmdlets (version 0.3)
        25.01.2014: loads of fixes. Added more details to Packages. Added -PDF switch to generate PDF output.
        27.01.2014: DP info via site local WMI class instead of going out to the DP
        28.01.2014: more info out of Applications, DP still does not account for Cloud DP, check if running from UNC path, which might cause issues with saving the file, bug fixes
#>

[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param(	
    
    [parameter(
	Position = 0, 
	Mandatory=$true )
	] 
	[Alias("SMS")]
    [ValidateScript({
        $ping = New-Object System.Net.NetworkInformation.Ping
        $ping.Send("$_", 5000)})]
	[ValidateNotNullOrEmpty()]
	[string]$SMSProvider="",
    
	[parameter(
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Motion", 

	[parameter(
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

    [parameter(
	Position = 3, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName=$env:username,

    [parameter(
	Position = 4, 
	Mandatory=$false )
	] 
	[Alias("LA")]
	[ValidateNotNullOrEmpty()]
	[switch]$ListAllInformation,

	[parameter(
	Position = 5, 
	Mandatory=$false )
	] 
	[switch]$PDF
)

#check if running from UNC folder
if ($($MyInvocation.MyCommand.Path).ToString().Startswith("\\"))
    {
        Write-Output "Please execute this script from a local drive. Otherwise you might encounter issues with saving the Word file."
        $UserInput = Read-Host -Prompt "Would you still like to continue and save the file manually? (y)es or (n)o"
        if ($UserInput.ToString().ToLower() -inotlike "y*")
            {
                Write-Output "Aborting script by User's choice!"
                exit
            }
    }


#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
	
Set-StrictMode -Version 2

#the following values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
[int]$wdAlignPageNumberRight = 2
[long]$wdColorGray15 = 14277081
[int]$wdMove = 0
[int]$wdSeekMainDocument = 0
[int]$wdSeekPrimaryFooter = 4
[int]$wdStory = 6
[int]$wdColorRed = 255
[int]$wdColorBlack = 0
[int]$wdWord2007 = 12
[int]$wdWord2010 = 14
[int]$wdWord2013 = 15
[int]$wdSaveFormatPDF = 17
[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

$hash = @{}

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

Switch ($PSUICulture.Substring(0,3))
{
	'ca-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Taula automÃ¡tica 2';
			}
		}

	'da-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabel 2';
			}
		}

	'de-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische Tabelle 2';
			}
		}

	'en-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}

	'es-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Tabla automÃ¡tica 2';
			}
		}

	'fi-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automaattinen taulukko 2';
			}
		}

	'fr-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Sommaire Automatique 2';
			}
		}

	'nb-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabell 2';
			}
		}

	'nl-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
			}
		}

	'pt-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'SumÃ¡rio AutomÃ¡tico 2';
			}
		}

	'sv-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk innehÃ¥llsfÃ¶rteckning2';
			}
		}

	Default	{$hash.('en-US') = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}
}

# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
[int]$wdStyleHeading1 = -2
[int]$wdStyleHeading2 = -3
[int]$wdStyleHeading3 = -4
[int]$wdStyleHeading4 = -5
[int]$wdStyleNoSpacing = -158
[int]$wdTableGrid = -155

$myHash = $hash.$PSUICulture

If($myHash -eq $Null)
{
	$myHash = $hash.('en-US')
}

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4
$myHash.Word_TableGrid = $wdTableGrid

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP)
	
	$xArray = ""
	
	Switch ($PSUICulture.Substring(0,3))
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "IÃ³ (clar)", "IÃ³ (fosc)", "LÃ­nia lateral",
					"Moviment", "QuadrÃ­cula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "SemÃ for", "VisualitzaciÃ³", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "DiplomÃ tic", "ExposiciÃ³",
					"LÃ­nia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "QuadrÃ­cula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast",
					"Cubicles", "DiplomÃ tic", "En mosaic", "ExposiciÃ³", "LÃ­nia lateral",
					"Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevÃ¦gElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mÃ¸rk)", "Ion (mÃ¸rk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevÃ¦gElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "GÃ¥de",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"NÃ¥lestribet", "Ã…rlig", "Avispapir", "Tradionel")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Ã…rlig", "BevÃ¦gElse", "Eksponering",
					"Enkel", "Firkanter", "Fliser", "GÃ¥de", "Kontrast",
					"Mod", "NÃ¥lestribet", "Overskrid", "Sidelinje", "Stakke",
					"Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "RÃ¼ckblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "JÃ¤hrlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt",
					"JÃ¤hrlich", "Kacheln", "Kontrast", "Kubistisch", "Modern",
					"Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel",
					"Traditionell")
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
					"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
					"Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "SemÃ¡foro", "Retrospectiva", "CuadrÃ­cula",
					"Movimiento", "Cortar (oscuro)", "LÃ­nea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "CuadrÃ­cula", "CubÃ­culos", "ExposiciÃ³n", "LÃ­nea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periÃ³dico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador",
					"Contraste", "CubÃ­culos", "ExposiciÃ³n", "LÃ­nea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle",
					"Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "VaihtuvavÃ¤rinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti", "Kuvakkeet ja tiedot",
					"Liike" , "Liituraita" , "Mod" , "Palapeli", "Perinteinen", "Pinot",
					"Sivussa", "TyÃ¶pisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncÃ©)", "SÃ©maphore",
					"RÃ©trospective", "Ion (foncÃ©)", "Ion (clair)", "IntÃ©grale",
					"Filigrane", "Facette", "Secteur (clair)", "Ã€ bandes", "Austin",
					"Guide", "Whisp", "Lignes latÃ©rales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("MosaÃ¯ques", "Ligne latÃ©rale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilÃ©s",
					"Rayures fines", "AustÃ¨re", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisÃ©s", "Papier journal", "Austin", "Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "AustÃ¨re", "Blocs empilÃ©s", "Blocs superposÃ©s",
					"Classique", "Contraste", "Exposition", "Guide", "Ligne latÃ©rale", "Moderne",
					"MosaÃ¯ques", "Mots croisÃ©s", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mÃ¸rk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mÃ¸rk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Ã…rlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Ã…rlig", "Avlukker", "BevegElse", "Engasjement",
					"Enkel", "Fliser", "Konservativ", "Kontrast", "Mod", "Puslespill",
					"Sidelinje", "Smale striper", "Stabler", "Transcenderende")
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging",
					"Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks", "Krijtstreep",
					"Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("AnimaÃ§Ã£o", "Austin", "Em Tiras", "ExibiÃ§Ã£o Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
					"Grade", "Integral", "Ãon (Claro)", "Ãon (Escuro)", "Linha Lateral",
					"Retrospectiva", "SemÃ¡foro")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "AnimaÃ§Ã£o", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "ExposiÃ§Ã£o", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeÃ§a", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "AnimaÃ§Ã£o", "Anual", "Austero", "Baias", "Conservador",
					"Contraste", "ExposiÃ§Ã£o", "Ladrilhos", "Linha Lateral", "Listras", "Mod",
					"Pilhas", "Quebra-cabeÃ§a", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mÃ¶rkt)", "Knippe", "RutnÃ¤t", "RÃ¶rElse", "Sektor (ljus)", "Sektor (mÃ¶rk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Ã…terblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("AlfabetmÃ¶nster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "RutnÃ¤t",
					"RÃ¶rElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Ã…rligt",
					"Ã–vergÃ¥ende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("AlfabetmÃ¶nster", "Ã…rligt", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Ã–vergÃ¥ende", "Plattor", "Pussel", "RÃ¶rElse",
					"Sidlinje", "Sobert", "Staplat")
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
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
						"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
						"Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}
Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Output "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Output "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function CheckWord2007SaveAsPDFInstalled
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
	{
		Write-Output "Word 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
		Write-Output "The add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
		Write-Output "Install the SaveAs PDF add-in and rerun the script."
		Return $False
	}
	Return $True
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

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This function just gets $true or $false
function Test-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $null -ne $key.GetValue($name, $null)
}

# Gets the specified registry value or $null if it is missing
function Get-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    if ($key) {
        $key.GetValue($name, $null)
    }
} 
Function WriteWordLine
#function created by Ryan Revord
#@rsrevord on Twitter
#function created to make output to Word easy in this script
{
	Param( [int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [switch]$nonewline, [switch]$bold)
	$output=""
	#Build output style
	
	switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
        4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}

	<##build # of tabs
	While( $tabs -gt 0 ) { 
		$output += "`t"; $tabs--; 
	}
    #>
	# Rather than indenting text, let's apply a bullet style instead
    If($tabs -gt 1) {
        $Selection.Style = "List Bullet $tabs"
    }
	
	#output the rest of the parameters.
	$output += $name + $value
    
    if ($bold)
        {
            $Selection.Font.Bold = 1
        }
    else
        {
	        $Selection.Font.Bold = 0
        }

	$Selection.TypeText($output)
    
	#test for new WriteWordLine 0.
	If($nonewline){
		# Do nothing.
	} Else {
		$Selection.TypeParagraph()
	}   
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop=$properties | foreach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$null,$_,$null)
		if ($propname -eq $Name) 
		{
			Return $_
		}
	} #foreach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$null,$prop,$Value)
}

Function AbortScript
{
	$Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

$script:startTime = Get-Date
[string]$Title = "Inventory Report for System Center Configuration Manager"
[string]$filename1 = "$($pwd.path)\$($CompanyName).docx"
If($PDF)
    {
	    [string]$filename2 = "$($pwd.path)\$($CompanyName).pdf"
    }
CheckWordPreReq

Write-Verbose "$(Get-Date): Setting up Word"

# Setup word for output
Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	Write-Error "The Word object could not be created.  You may need to repair your Word installation.  Script cannot continue."
	Exit
}

[int]$WordVersion = [int]$Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	Write-Error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
	AbortScript
}

Write-Verbose "$(Get-Date): Running Microsoft $WordProduct"

If($PDF -and $WordVersion -eq $wdWord2007)
{
	Write-Verbose "$(Get-Date): Verify the Word 2007 Save As PDF add-in is installed"
	If(CheckWord2007SaveAsPDFInstalled)
	{
		Write-Verbose "$(Get-Date): The Word 2007 Save As PDF add-in is installed"
	}
	Else
	{
		AbortScript
	}
}

Write-Verbose "$(Get-Date): Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Warning "Company Name cannot be blank."
		Write-Warning "Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
		Write-Error "Script cannot continue.  See messages above."
		AbortScript
	}
}

Write-Verbose "$(Get-Date): Check Default Cover Page for language specific version"
[bool]$CPChanged = $False
Switch ($PSUICulture.Substring(0,3))
{
	'ca-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "LÃ­nia lateral"
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
				$CoverPage = "LÃ­nea lateral"
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
				If($WordVersion -eq $wdWord2013)
				{
					$CoverPage = "Lignes latÃ©rales"
					$CPChanged = $True
				}
				Else
				{
					$CoverPage = "Ligne latÃ©rale"
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

Write-Verbose "$(Get-Date): Validate cover page"
[bool]$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	Write-Error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
	AbortScript
}

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Company Name : $CompanyName"
Write-Verbose "$(Get-Date): Cover Page   : $CoverPage"
Write-Verbose "$(Get-Date): User Name    : $UserName"
Write-Verbose "$(Get-Date): Save As PDF  : $PDF"
Write-Verbose "$(Get-Date): Title        : $Title"
Write-Verbose "$(Get-Date): Filename1    : $filename1"
If($PDF)
{
	Write-Verbose "$(Get-Date): Filename2    : $filename2"
}
Write-Verbose "$(Get-Date): OS Detected  : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture  : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture    : $PSCulture"
Write-Verbose "$(Get-Date): Word version : $WordProduct"
Write-Verbose "$(Get-Date): Word language: $($Word.Language)"
Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

$Word.Visible = $true

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $configlog = $False is from Jeff Hicks
Write-Verbose "$(Get-Date): Load Word Templates"

[bool]$CoverPagesExist = $False
[bool]$BuildingBlocksExist = $False

$word.Templates.LoadBuildingBlocks()
If($WordVersion -eq $wdWord2007)
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
$part = $Null

If($BuildingBlocks -ne $Null)
{
	$BuildingBlocksExist = $True

	Try 
		{$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)}

	Catch
		{$part = $Null}

	If($part -ne $Null)
	{
		$CoverPagesExist = $True
	}
}

#cannot continue if cover page does not exist
If(!$CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Error "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist.  Script cannot continue."
	Write-Verbose "$(Get-Date): Closing Word"
	AbortScript
}

Write-Verbose "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "An empty Word document could not be created.  Script cannot continue."
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "An unknown error happened selecting the entire Word document for default formatting options.  Script cannot continue."
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Verbose "$(Get-Date): Disable grammar and spell checking"
$Word.Options.CheckGrammarAsYouType = $False
$Word.Options.CheckSpellingAsYouType = $False

If($BuildingBlocksExist)
{
	#insert new page, getting ready for table of contents
	Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

	#table of contents
	Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
	$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
	If($toc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
		Write-Warning "This report will not have a Table of Contents."
	}
	Else
	{
		$toc.insert($selection.Range,$True) | out-null
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
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
#get the footer and format font
$footers = $doc.Sections.Last.Footers
ForEach($footer in $footers) 
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
$selection.HeaderFooter.Range.Text = $footerText

#add page numbering
Write-Verbose "$(Get-Date): Add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
Write-Verbose "$(Get-Date): Return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date): Move to the end of the current document"
Write-Verbose "$(Get-Date):"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 
Function Convert-NormalDateToConfigMgrDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$starttime
    )

    return [System.Management.ManagementDateTimeconverter]::ToDateTime($starttime)
}

Function Read-ScheduleToken {

$SMS_ScheduleMethods = "SMS_ScheduleMethods"
$class_SMS_ScheduleMethods = [wmiclass]""
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
        1 {$weekday = "Sunday"}
        2 {$weekday = "Monday"}
        3 {$weekday = "Tuesday"}
        4 {$weekday = "Wednesday"}
        5 {$weekday = "Thursday"}
        6 {$weekday = "Friday"}
        7 {$weekday = "Saturday"}
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
    $wqlQuery = “SELECT * FROM SMS_ProviderLocation”
    $a = Get-WmiObject -Query $wqlQuery -Namespace “root\sms” -ComputerName $SMSProvider
    $a | ForEach-Object {
        if($_.ProviderForLocalSite)
            {
                $script:SiteCode = $_.SiteCode
            }
    }
return $SiteCode
}

function Get-ExecuteWqlQuery($siteServerName, $query)
{
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

function Get-ApplicationObjectFromServer($appName,$siteServerName)
{    
    $resultObject = Get-ExecuteWqlQuery $siteServerName "select thissitecode from sms_identification" 
    $siteCode = $resultObject["thissitecode"].StringValue
    
    $path = [string]::Format("\\{0}\ROOT\sms\site_{1}", $siteServerName, $siteCode)
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
        $sdmPackageXml = $returnValue.Properties["SDMPackageXML"].Value.ToString()
        [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($sdmPackageXml)
    }
}


 function Load-ConfigMgrAssemblies()
 {
     
     $AdminConsoleDirectory = Split-Path $env:SMS_ADMIN_UI_PATH -Parent
     $filesToLoad = "Microsoft.ConfigurationManagement.ApplicationManagement.dll","AdminUI.WqlQueryEngine.dll", "AdminUI.DcmObjectWrapper.dll" 
     
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
              Write-Output ([System.String]::Format("File not found {0}",$fileName )) -backgroundcolor "red"
         }
      }
 }

$SiteCode = Get-SiteCode

# Set Styles
Write-Verbose "$(Get-Date):  Setting your table style"
#$TableStyle = $doc.Styles | ?{$_.namelocal -eq "Grid Table 4 - Accent 5" }
$TableStyle = $myHash.Word_TableGrid
#$HeadingStyle = 

##################### MAIN SCRIPT STARTS HERE #######################

$LocationBeforeExecution = Get-Location

$selection.InsertNewPage() | Out-Null

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

#Load-ConfigMgrAssemblies 

#### Administration
#### Site Configuration

WriteWordLine 1 0 "Summary of all Sites in this Hierarchy"
Write-Verbose "$(Get-Date):   Getting Site Information"
$CMSites = Get-CMSite

$CAS = $CMSites | Where-Object {$_.Type -eq "4"}
$StandAlonePrimarySites = $CMSites | Where-Object {$_.Type -eq "2"}
$ChildPrimarySites = $CMSites | Where-Object {$_.Type -eq "3"}
$SecondarySites = $CMSites | Where-Object {$_.Type -eq "1"}

if (-not [string]::IsNullOrEmpty($CAS))
    {
        WriteWordLine 0 1 "The following Central Administration Site is installed:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
		$Columns = 3
        [int]$Rows = $CAS.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $TableStyle 
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		[int]$xRow = 1
		
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = "Site Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = "Site Code"
        
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = "Version"                      
        $xRow++							
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = $CAS.SiteName
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = $CAS.SiteCode
		$Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = $CAS.Version
		
		$Table.Rows.SetLeftIndent(50,1) | Out-Null
		$table.AutoFitBehavior(1) | Out-Null

		#return focus back to document
		Write-Verbose "$(Get-Date):   return focus back to document"
        $selection.EndOf(15) | Out-Null        $selection.MoveDown() | Out-Null
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
		Write-Verbose "$(Get-Date):   move to the end of the current document"
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
    }

if (-not [string]::IsNullOrEmpty($ChildPrimarySites))
    {
        Write-Verbose "$(Get-Date):   Enumerating all Primary Sites"
        WriteWordLine 0 1 "The following Primary Sites are installed:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
		$Columns = 3
        [int]$Rows = $ChildPrimarySites.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $TableStyle 
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		[int]$xRow = 1
		
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = "Site Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = "Site Code"
        
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = "Version"                      
        foreach ($ChildPrimarySite in $ChildPrimarySites)
            {
                $xRow++							
		        $Table.Cell($xRow,1).Range.Font.Size = "10"
		        $Table.Cell($xRow,1).Range.Text = $ChildPrimarySite.SiteName
		        $Table.Cell($xRow,2).Range.Font.Size = "10"
		        $Table.Cell($xRow,2).Range.Text = $ChildPrimarySite.SiteCode
                $Table.Cell($xRow,3).Range.Font.Size = "10"
		        $Table.Cell($xRow,3).Range.Text = $ChildPrimarySite.Version
            }				
		$Table.Rows.SetLeftIndent(50,1) | Out-Null
		$table.AutoFitBehavior(1) | Out-Null
 
		#return focus back to document
		Write-Verbose "$(Get-Date):   return focus back to document"
        $selection.EndOf(15) | Out-Null        $selection.MoveDown() | Out-Null
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
		Write-Verbose "$(Get-Date):   move to the end of the current document"
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
    }

if (-not [string]::IsNullOrEmpty($StandAlonePrimarySites))
    {
        Write-Verbose "$(Get-Date):   Enumerating all standalone Primary Sites."
        WriteWordLine 0 1 "The following Primary Sites are installed:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
		$Columns = 3
        [int]$Rows = $StandAlonePrimarySites.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $TableStyle 
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		[int]$xRow = 1
		
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = "Site Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = "Site Code"
        
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = "Version"                      
        $xRow++
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = $StandAlonePrimarySites.SiteName
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = $StandAlonePrimarySites.SiteCode
        $Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = $StandAlonePrimarySites.Version				
		$Table.Rows.SetLeftIndent(50,1) | Out-Null
		$table.AutoFitBehavior(1) | Out-Null

		#return focus back to document
		Write-Verbose "$(Get-Date):   return focus back to document"
        $selection.EndOf(15) | Out-Null        $selection.MoveDown() | Out-Null
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
		Write-Verbose "$(Get-Date):   move to the end of the current document"
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
    }
if (-not [string]::IsNullOrEmpty($SecondarySites))
    {
        Write-Verbose "$(Get-Date):   Enumerating all secondary sites."
        WriteWordLine 0 1 "The following Secondary Sites are installed:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
		$Columns = 3
        [int]$Rows = $SecondarySites.count + 1
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $TableStyle 
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		[int]$xRow = 1
		
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = "Site Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = "Site Code"
        
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = "Version"                      
        foreach ($SecondarySite in $SecondarySites)
            {
                $xRow++
		        $Table.Cell($xRow,1).Range.Font.Size = "10"
		        $Table.Cell($xRow,1).Range.Text = $SecondarySite.SiteName
		        $Table.Cell($xRow,2).Range.Font.Size = "10"
		        $Table.Cell($xRow,2).Range.Text = $SecondarySite.SiteCode
                $Table.Cell($xRow,3).Range.Font.Size = "10"
		        $Table.Cell($xRow,3).Range.Text = $SecondarySite.Version
            }				
		$Table.Rows.SetLeftIndent(50,1) | Out-Null
		$table.AutoFitBehavior(1) | Out-Null

		#return focus back to document
		Write-Verbose "$(Get-Date):   return focus back to document"
        $selection.EndOf(15) | Out-Null        $selection.MoveDown() | Out-Null
		$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
		Write-Verbose "$(Get-Date):   move to the end of the current document"
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
    }

foreach ($CMSite in $CMSites)
    {
            Write-Verbose "$(Get-Date):   Checking each site's configuration."
            WriteWordLine 1 0 "Configuration Summary for Site $($CMSite.SiteCode)"
            WriteWordLine 0 0 ""   
            $SiteMaintenanceTasks = Get-CMSiteMaintenanceTask -SiteCode $CMSite.SiteCode
            WriteWordLine 2 1 "Site Maintenance Tasks for Site $($CMSite.SiteCode)"
            $Table = $Null
            $TableRange = $Null
            $TableRange = $doc.Application.Selection.Range
			$Columns = 3
            [int]$Rows = $SiteMaintenanceTasks.count + 1
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $TableStyle 
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Font.Size = "10"
			$Table.Cell($xRow,1).Range.Text = "Task Name"
			
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Font.Size = "10"
			$Table.Cell($xRow,2).Range.Text = "State"

			$Table.Cell($xRow,3).Range.Font.Bold = $True
			$Table.Cell($xRow,3).Range.Font.Size = "10"
			$Table.Cell($xRow,3).Range.Text = "Location"                                  
            foreach ($SiteMaintenanceTask in $SiteMaintenanceTasks)
				{
					$xRow++							
					$Table.Cell($xRow,1).Range.Font.Size = "10"
					$Table.Cell($xRow,1).Range.Text = $SiteMaintenanceTask.TaskName
					$Table.Cell($xRow,2).Range.Font.Size = "10"
					$Table.Cell($xRow,2).Range.Text = $SiteMaintenanceTask.Enabled
					if ($SiteMaintenanceTask.TaskName -eq "Backup SMS Site Server")
                        {
                            $Table.Cell($xRow,3).Range.Font.Size = "10"
					        $Table.Cell($xRow,3).Range.Text = $SiteMaintenanceTask.DeviceName
                        }
				}
				
			$Table.Rows.SetLeftIndent(50,1) | Out-Null
			$table.AutoFitBehavior(1) | Out-Null

			#return focus back to document
			Write-Verbose "$(Get-Date):   return focus back to document"
            $selection.EndOf(15) | Out-Null            $selection.MoveDown() | Out-Null
			$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

            $CMManagementPoints = Get-CMManagementPoint -SiteCode $CMSite.SiteCode

            $selection.EndKey($wdStory,$wdMove) | Out-Null

            WriteWordLine 2 1 "Summary of Management Points for Site $($CMSite.SiteCode)"
            foreach ($CMManagementPoint in $CMManagementPoints)
                {
                    Write-Verbose "$(Get-Date):   Management Point: $($CMManagementPoint)"
                    $CMMPServerName = $CMManagementPoint.NetworkOSPath.Split("\\")[2]
                    WriteWordLine 0 1 "$($CMMPServerName)"
                }

    WriteWordLine 2 1 "Summary of Distribution Points for Site $($CMSite.SiteCode)"
    $CMDistributionPoints = Get-CMDistributionPoint -SiteCode $CMSite.SiteCode
    foreach ($CMDistributionPoint in $CMDistributionPoints)
        {
            $CMDPServerName = $CMDistributionPoint.NetworkOSPath.Split("\\")[2]
            Write-Verbose "$(Get-Date):   Found DP: $($CMDPServerName)"
            WriteWordLine 0 1 "$($CMDPServerName)" -bold
            Write-Verbose "Trying to ping $($CMDPServerName)"
            $PingResult = Test-NetConnection -ComputerName $CMDPServerName
            if (-not ($PingResult.PingSucceeded))
                {
                    WriteWordLine 0 2 "The Distribution Point $($CMDPServerName) is not reachable. Check connectivity."
                }
            else
                {
                    WriteWordLine 0 2 "Disk information:"
                    $CMDPDrives = Get-WmiObject -Class SMS_DistributionPointDriveInfo -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.NALPath -like "*$CMDPServerName*"}
                    foreach ($CMDPDrive in $CMDPDrives)
                        {
                            WriteWordLine 0 2 "Partition $($CMDPDrive.Drive):" -bold
                            $Size = ""
                            $Size = $CMDPDrive.BytesTotal / 1024 / 1024
                            $Freesize = ""
                            $Freesize = $CMDPDrive.BytesFree / 1024 / 1024

                            WriteWordLine 0 3 "$([MATH]::Round($Size,2)) GB Size in total"
                            WriteWordLine 0 3 "$([MATH]::Round($Freesize,2)) GB Free Space"
                            WriteWordLine 0 3 "Still $($CMDPDrive.PercentFree) percent free."
                            WriteWordLine 0 0
                        }

                    WriteWordLine 0 2 "Hardware Info:" -bold
                    $Capacity = ""
                    Get-WmiObject -Class win32_physicalmemory -ComputerName $CMDPServerName | foreach {$Capacity += $_.Capacity}
                    $TotalMemory = $Capacity / 1024 / 1024 / 1024
                    WriteWordLine 0 3 "This server has a total of $($TotalMemory) GB RAM."
                }

            $DPInfo = $CMDistributionPoint.Props
            $IsPXE = ($DPInfo | where {$_.PropertyName -eq "IsPXE"}).Value
            $UnknownMachines = ($DPInfo | where {$_.PropertyName -eq "SupportUnknownMachines"}).Value
            switch ($IsPXE)
                {
                    1 
                        {
                            WriteWordLine 0 2 "PXE Enabled"
                            switch ($UnknownMachines)
                                {
                                    1 { WriteWordLine 0 2 "Supports unknown machines: true" }
                                    0 { WriteWordLine 0 2 "Supports unknown machines: false" }
                                }
                        }
                    0
                        {
                            WriteWordLine 0 2 "PXE disabled"
                        }
                }

            $DPGroupMembers = $Null
            $DPGroupIDs = $Null
            $DPGroupMembers = Get-WmiObject -class SMS_DPGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.DPNALPath -ilike "*$($CMDPServerName)*"}
            if (-not [string]::IsNullOrEmpty($DPGroupMembers))
                {
                    $DPGroupIDs = $DPGroupMembers.GroupID
                }
            
            #enumerating DP Group Membership
            if (-not [string]::IsNullOrEmpty($DPGroupIDs))
                {
                    WriteWordLine 0 1 "This Distribution Point is a member of the following DP Groups:"
                    foreach ($DPGroupID in $DPGroupIDs)
                        {
                            $DPGroupName = (Get-CMDistributionPointGroup -Id "$($DPGroupID)").Name
                            WriteWordLine 0 2 "$($DPGroupName)"
                        }
                }
            else
                {
                    WriteWordLine 0 1 "This Distribution Point is not a member of any DP Group."
                }
        }

    #enumerating Software Update Points
    Write-Verbose "$(Get-Date):   Enumerating all Software Update Points"
    WriteWordLine 2 1 "Summary of Software Update Point Servers for Site $($CMSite.SiteCode)"
    #$CMSUPs = Get-WmiObject -Class sms_sci_sysresuse -Namespace root\sms\site_$($CMSite.SiteCode) -ComputerName $CMMPServerName | Where-Object {$_.rolename -eq "SMS Software Update Point"}
    $CMSUPs = Get-CMSoftwareUpdatePoint | Where-Object {$_.SiteCode -eq "$($CMSite.SiteCode)"}
    if (-not [string]::IsNullOrEmpty($CMSUPs))
        {
            foreach ($CMSUP in $CMSUPs)
                {
                    $CMSUPServerName = $CMSUP.NetworkOSPath.split("\\")[2]
                    Write-Verbose "$(Get-Date):   Found SUP: $($CMSUPServerName)"
                    WriteWordLine 0 1 "$($CMSUPServerName)"
                    $Table = $Null
                    $TableRange = $Null
                    $TableRange = $doc.Application.Selection.Range
			        $Columns = 4
                    [int]$Rows = $($CMSUP.Props).count + 1
			        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			        $table.Style = $TableStyle 
			        $table.Borders.InsideLineStyle = 1
			        $table.Borders.OutsideLineStyle = 1
			        [int]$xRow = 1
			        
			        $Table.Cell($xRow,1).Range.Font.Bold = $True
			        $Table.Cell($xRow,1).Range.Font.Size = "10"
			        $Table.Cell($xRow,1).Range.Text = "Property Name"
			        
			        $Table.Cell($xRow,2).Range.Font.Bold = $True
			        $Table.Cell($xRow,2).Range.Font.Size = "10"
			        $Table.Cell($xRow,2).Range.Text = "Value"
			        
			        $Table.Cell($xRow,3).Range.Font.Bold = $True
			        $Table.Cell($xRow,3).Range.Font.Size = "10"
			        $Table.Cell($xRow,3).Range.Text = "Value 1" 
			        
			        $Table.Cell($xRow,4).Range.Font.Bold = $True
			        $Table.Cell($xRow,4).Range.Font.Size = "10"
			        $Table.Cell($xRow,4).Range.Text = "Value 2"                                  
                    foreach ($SUPProp in $CMSUP.Props)
				        {
					        $xRow++							
					        $Table.Cell($xRow,1).Range.Font.Size = "10"
					        $Table.Cell($xRow,1).Range.Text = $SUPProp.PropertyName
					        $Table.Cell($xRow,2).Range.Font.Size = "10"
					        $Table.Cell($xRow,2).Range.Text = $SUPProp.Value
					        $Table.Cell($xRow,3).Range.Font.Size = "10"
					        $Table.Cell($xRow,3).Range.Text = $SUPProp.Value1
					        $Table.Cell($xRow,4).Range.Font.Size = "10"
					        $Table.Cell($xRow,4).Range.Text = $SUPProp.Value2
				        }
				
			        $Table.Rows.SetLeftIndent(50,1) | Out-Null
			        $table.AutoFitBehavior(1) | Out-Null

			        #return focus back to document
			        Write-Verbose "$(Get-Date):   return focus back to document"
                    $selection.EndOf(15) | Out-Null                    $selection.MoveDown() | Out-Null
			        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                   
                    
                }
        }
    else
        {
            WriteWordLine 0 1 "This site has no Software Update Points installed."
        }
    $selection.EndKey($wdStory,$wdMove) | Out-Null
}

$selection.EndKey($wdStory,$wdMove) | Out-Null
##### Hierarchy wide configuration
WriteWordLine 1 0 "Summary of Hierarchy Wide Configuration"

### enumerating Boundaries
Write-Verbose "$(Get-Date):   Enumerating all Site Boundaries"
WriteWordLine 2 0 "Summary of Site Boundaries"

$Boundaries = Get-CMBoundary
    if (-not [string]::IsNullOrEmpty($Boundaries))
        {
            $Table = $Null
            $TableRange = $Null
            $TableRange = $doc.Application.Selection.Range
	        $Columns = 5
            [int]$Rows = $Boundaries.count + 1
	        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	        $table.Style = $TableStyle 
	        $table.Borders.InsideLineStyle = 1
	        $table.Borders.OutsideLineStyle = 1
	        [int]$xRow = 1
	        
	        $Table.Cell($xRow,1).Range.Font.Bold = $True
	        $Table.Cell($xRow,1).Range.Font.Size = "10"
	        $Table.Cell($xRow,1).Range.Text = "Boundary Name"
	        
	        $Table.Cell($xRow,2).Range.Font.Bold = $True
	        $Table.Cell($xRow,2).Range.Font.Size = "10"
	        $Table.Cell($xRow,2).Range.Text = "Boundary Type"
            
	        $Table.Cell($xRow,3).Range.Font.Bold = $True
	        $Table.Cell($xRow,3).Range.Font.Size = "10"
	        $Table.Cell($xRow,3).Range.Text = "Associated Site Systems"
            
	        $Table.Cell($xRow,4).Range.Font.Bold = $True
	        $Table.Cell($xRow,4).Range.Font.Size = "10"
	        $Table.Cell($xRow,4).Range.Text = "Value" 
            
	        $Table.Cell($xRow,5).Range.Font.Bold = $True
	        $Table.Cell($xRow,5).Range.Font.Size = "10"
	        $Table.Cell($xRow,5).Range.Text = "Assigned Site"                                  
            foreach ($Boundary in $Boundaries)
		        {
			        $BoundarySiteSystems = $Null
                    $xRow++							
			        $Table.Cell($xRow,1).Range.Font.Size = "10"
			        $Table.Cell($xRow,1).Range.Text = $Boundary.DisplayName
                    switch ($Boundary.BoundaryType)
                        {
                            0 { $BoundaryType = "IP Subnet" }
                            1 { $BoundaryType = "Active Directory Site" }
                            2 { $BoundaryType = "IPv6 Prefix" }
                            3 { $BoundaryType = "IP Range" }
                        }
                    $Table.Cell($xRow,2).Range.Font.Size = "10"
			        $Table.Cell($xRow,2).Range.Text = $BoundaryType
                    $Table.Cell($xRow,3).Range.Font.Size = "10"
			        $BoundarySiteSystems = $Null
                    $NamesOfBoundarySiteSystems = $Null
                    if (-not [string]::IsNullOrEmpty($Boundary.SiteSystems))
                        {
                            ForEach-Object -Begin {$BoundarySiteSystems= $Boundary.SiteSystems} -Process {$NamesOfBoundarySiteSystems += $BoundarySiteSystems.split(",")} -End {$NamesOfBoundarySiteSystems} | Out-Null
                        }
                    else 
                        {
                            $NamesOfBoundarySiteSystems = "n/a"
                        }
                    $Table.Cell($xRow,3).Range.Text = $NamesOfBoundarySiteSystems
                    $Table.Cell($xRow,4).Range.Font.Size = "10"
			        $Table.Cell($xRow,4).Range.Text = $Boundary.Value
                    $Table.Cell($xRow,5).Range.Font.Size = "10"
			        $Table.Cell($xRow,5).Range.Text = $Boundary.DefaultSiteCode
		        }
				
	        $Table.Rows.SetLeftIndent(50,1) | Out-Null
	        $table.AutoFitBehavior(1) | Out-Null

	        #return focus back to document
	        Write-Verbose "$(Get-Date):   return focus back to document"
            $selection.EndOf(15) | Out-Null            $selection.MoveDown() | Out-Null
	        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        }

### enumerating all Boundary Groups and their members
Write-Verbose "$(Get-Date):   Enumerating all Boundary Groups and their members"

$BoundaryGroups = Get-CMBoundaryGroup
WriteWordLine 2 0 "Summary of Site Boundary Groups"
if (-not [string]::IsNullOrEmpty($BoundaryGroups))
    {
        foreach ($BoundaryGroup in $BoundaryGroups)
            {
                WriteWordLine 0 1 "$($BoundaryGroup.Name)" -bold
                WriteWordLine 0 2 "Description: $($BoundaryGroup.Description)"

                if ($BoundaryGroup.SiteSystemCount -gt 0)
                    {
                        $MemberIDs = (Get-WmiObject -Class SMS_BoundaryGroupMembers -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider | Where-Object {$_.GroupID -eq "$($BoundaryGroup.GroupID)"}).BoundaryID
                        foreach ($MemberID in $MemberIDs)
                            {
                                $MemberName = (Get-CMBoundary -Id $MemberID).DisplayName
                                WriteWordLine 0 2 "Member names: $($MemberName)"
                            }
                    }
                else
                    {
                        WriteWordLine 0 2 "There are no Site Systems associated to this Boundary Group."
                    }
            }
    }
else
    {
        WriteWordLine 0 1 "There are no Boundary Groups configured. It is mandatory to configure a Boundary Group in order for CM12 to work properly."
    }

### enumerating Client Policies
Write-Verbose "$(Get-Date):   Enumerating all Client/Device Settings"
WriteWordLine 2 0 "Summary of Custom Client Device Settings"

$AllClientSettings = Get-CMClientSetting | Where-Object {$_.SettingsID -ne "0"}
foreach ($ClientSetting in $AllClientSettings)
    {
        WriteWordLine 0 1 "Client Settings Name: $($ClientSetting.Name)" -bold
        WriteWordLine 0 2 "Client Settings Description: $($ClientSetting.Description)"
        WriteWordLine 0 2 "Client Settings ID: $($ClientSetting.SettingsID)"
        WriteWordLine 0 2 "Client Settings Priority: $($ClientSetting.Priority)"
        if ($ClientSetting.Type -eq "1")
            {
                WriteWordLine 0 2 "This is a custom client Device Setting."
            }
        else
            {
                WriteWordLine 0 2 "This is a custom client User Setting."
            }
        WriteWordLine 0 1 "Configurations"
        foreach ($AgentConfig in $ClientSetting.AgentConfigurations)
            {
                try
                    {
                        switch ($AgentConfig.AgentID)
                            {
                                1
                                    {
                                        WriteWordLine 0 2 "Compliance Settings"
                                        WriteWordLine 0 2 "Enable compliance evaluation on clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Enable user data and profiles: $($AgentConfig.EnableUserStateManagement)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                2
                                    {
                                        WriteWordLine 0 2 "Software Inventory"
                                        WriteWordLine 0 2 "Enable software inventory on clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Schedule software inventory and file collection: " -nonewline
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 2 "Inventory reporting detail: " -nonewline
                                        switch ($AgentConfig.ReportOptions)
                                            {
                                                1 { WriteWordLine 0 0 "Product only" }
                                                2 { WriteWordLine 0 0 "File only" }
                                                7 { WriteWordLine 0 0 "Full details" }
                                            }
                                
                                        WriteWordLine 0 2 "Inventory these file types: "
                                        if ($AgentConfig.InventoriableTypes)
                                            {
                                                WriteWordLine 0 3 "$($AgentConfig.InventoriableTypes)"
                                            }
                                        if ($AgentConfig.Path)
                                            {                               
                                                WriteWordLine 0 3 "$($AgentConfig.Path)"
                                            }
                                        if (($AgentConfig.InventoriableTypes) -and ($AgentConfig.ExcludeWindirAndSubfolders -eq "true"))
                                            {
                                                WriteWordLine 0 3 "Exclude WinDir and Subfolders"
                                            }
                                        else 
                                            {
                                                WriteWordLine 0 3 "Do not exclude WinDir and Subfolders"
                                            }
                                 
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                3
                                    {
                                        WriteWordLine 0 2 "Remote Tools"
                                        WriteWordLine 0 2 "Enable Remote Control on clients: " -nonewline
                                        switch ($AgentConfig.FirewallExceptionProfiles)
                                            {
                                                0 { WriteWordLine 0 0 "Disabled" }
                                                8 { WriteWordLine 0 0 "Enabled: No Firewall Profile." }
                                                9 { WriteWordLine 0 2 "Enabled: Public." }
                                                10 { WriteWordLine 0 2 "Enabled: Private." }
                                                11 { WriteWordLine 0 2 "Enabled: Private, Public." }
                                                12 { WriteWordLine 0 2 "Enabled: Domain." }
                                                13 { WriteWordLine 0 2 "Enabled: Domain, Public." }
                                                14 { WriteWordLine 0 2 "Enabled: Domain, Private." }
                                                15 { WriteWordLine 0 2 "Enabled: Domain, Private, Public." }
                                            }
                                        WriteWordLine 0 2 "Users can change policy or notification settings in Software Center: $($AgentConfig.AllowClientChange)"
                                        WriteWordLine 0 2 "Allow Remote Control of an unattended computer: $($AgentConfig.AllowRemCtrlToUnattended)"
                                        WriteWordLine 0 2 "Prompt user for Remote Control permission: $($AgentConfig.PermissionRequired)"
                                        WriteWordLine 0 2 "Grant Remote Control permission to local Administrators group: $($AgentConfig.AllowLocalAdminToDoRemoteControl)"
                                        WriteWordLine 0 2 "Access level allowed: " -nonewline
                                        switch ($AgentConfig.AccessLevel)
                                            {
                                                0 { WriteWordLine 0 0 "No access" }
                                                1 { WriteWordLine 0 0 "View only" }
                                                2 { WriteWordLine 0 0 "Full Control" }
                                            }
                                        WriteWordLine 0 2 "Permitted viewers of Remote Control and Remote Assistance:"
                                        foreach ($Viewer in $AgentConfig.PermittedViewers)
                                            {
                                                WriteWordLine 0 3 "$($Viewer)"
                                            }
                                        WriteWordLine 0 2 "Show session notification icon on taskbar: $($AgentConfig.RemCtrlTaskbarIcon)"
                                        WriteWordLine 0 2 "Show session connection bar: $($AgentConfig.RemCtrlConnectionBar)"
                                        WriteWordLine 0 2 "Play a sound on client: " -nonewline
                                        Switch ($AgentConfig.AudibleSignal)
                                            {
                                                0 { WriteWordLine 0 0 "None." }
                                                1 { WriteWordLine 0 0 "Beginning and end of session." }
                                                2 { WriteWordLine 0 0 "Repeatedly during session." }
                                            }
                                        WriteWordLine 0 2 "Manage unsolicited Remote Assistance settings: $($AgentConfig.ManageRA)"
                                        WriteWordLine 0 2 "Manage solicited Remote Assistance settings: $($AgentConfig.EnforceRAandTSSettings)"
                                        WriteWordLine 0 2 "Level of access for Remote Assistance: " -nonewline
                                        if (($AgentConfig.AllowRAUnsolicitedView -ne "True") -and ($AgentConfig.AllowRAUnsolicitedControl -ne "True"))
                                            {
                                                WriteWordLine 0 0 "None."
                                            }
                                        elseif (($AgentConfig.AllowRAUnsolicitedView -eq "True") -and ($AgentConfig.AllowRAUnsolicitedControl -ne "True"))
                                            {
                                                WriteWordLine 0 0 "Remote viewing."
                                            }
                                        elseif (($AgentConfig.AllowRAUnsolicitedView -eq "True") -and ($AgentConfig.AllowRAUnsolicitedControl -eq "True"))
                                            {
                                                WriteWordLine 0 0 "Full Control."
                                            }
                                        WriteWordLine 0 2 "Manage Remote Desktop settings: $($AgentConfig.ManageTS)"
                                        if ($AgentConfig.ManageTS -eq "True")
                                            {
                                                WriteWordLine 0 2 "Allow permitted viewers to connect by using Remote Desktop connection: $($AgentConfig.EnableTS)"
                                                WriteWordLine 0 2 "Require network level authentication on computers that run Windows Vista operating system and later versions: $($AgentConfig.TSUserAuthentication)"
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                4
                                    {
                                        WriteWordLine 0 2 "Computer Agent"
                                        WriteWordLine 0 2 "Deployment deadline greater than 24 hours, remind user every (hours): $([string]($AgentConfig.ReminderInterval) / 60 / 60)"
                                        WriteWordLine 0 2 "Deployment deadline less than 24 hours, remind user every (hours): $([string]($AgentConfig.DayReminderInterval) / 60 / 60)"
                                        WriteWordLine 0 2 "Deployment deadline less than 1 hour, remind user every (minutes): $([string]($AgentConfig.HourReminderInterval) / 60)"
                                        WriteWordLine 0 2 "Default application catalog website point: $($AgentConfig.PortalUrl)"
                                        WriteWordLine 0 2 "Add default Application Catalog website to Internet Explorer trusted sites zone: $($AgentConfig.AddPortalToTrustedSiteList)"
                                        WriteWordLine 0 2 "Allow Silverlight applications to run in elevated trust mode: $($AgentConfig.AllowPortalToHaveElevatedTrust)"
                                        WriteWordLine 0 2 "Organization name displayed in Software Center: $($AgentConfig.BrandingTitle)"
                                        switch ($AgentConfig.InstallRestriction)
                                            {
                                                0 { $InstallRestriction = "All Users" }
                                                1 { $InstallRestriction = "Only Administrators" }
                                                3 { $InstallRestriction = "Only Administrators and primary Users"}
                                                4 { $InstallRestriction = "No users" }
                                            }
                                        WriteWordLine 0 2 "Install Permissions: $($InstallRestriction)"
                                        Switch ($AgentConfig.SuspendBitLocker)
                                            {
                                                0 { $SuspendBitlocker = "Never" }
                                                1 { $SuspendBitlocker = "Always" }
                                            }
                                        WriteWordLine 0 2 "Suspend Bitlocker PIN entry on restart: $($SuspendBitlocker)"
                                        Switch ($AgentConfig.EnableThirdPartyOrchestration)
                                            {
                                                0 { $EnableThirdPartyTool = "No" }
                                                1 { $EnableThirdPartyTool = "Yes" }
                                            }
                                        WriteWordLine 0 2 "Additional software manages the deployment of applications and software updates: $($EnableThirdPartyTool)"
                                        Switch ($AgentConfig.PowerShellExecutionPolicy)
                                            {
                                                0 { $ExecutionPolicy = "All signed" }
                                                1 { $ExecutionPolicy = "Bypass" }
                                                2 { $ExecutionPolicy = "Restricted" }
                                            }
                                        WriteWordLine 0 2 "Powershell execution policy: $($ExecutionPolicy)"
                                        switch ($AgentConfig.DisplayNewProgramNotification)
                                            {
                                                False { $DisplayNotifications = "No" }
                                                True { $DisplayNotifications = "Yes" }
                                            }
                                        WriteWordLine 0 2 "Show notifications for new deployments: $($DisplayNotifications)"
                                        switch ($AgentConfig.DisableGlobalRandomization)
                                            {
                                                False { $DisableGlobalRandomization = "No" }
                                                True { $DisableGlobalRandomization = "Yes" }
                                            }
                                        WriteWordLine 0 2 "Disable deadline randomization: $($DisableGlobalRandomization)"
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                5
                                    {
                                        WriteWordLine 0 2 "Network Access Protection (NAP)"
                                        WriteWordLine 0 2 "Enable Network Access Protection on clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Use UTC (Universal Time Coordinated) for evaluation time: $($AgentConfig.EffectiveTimeinUTC)"
                                        WriteWordLine 0 2 "Require a new scan for each evaluation: $($AgentConfig.ForceScan)"
                                        WriteWordLine 0 2 "NAP re-evaluation schedule:" -nonewline
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                8
                                    {
                                        WriteWordLine 0 2 "Software Metering"
                                        WriteWordLine 0 2 "Enable software metering on clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Schedule data collection: " -nonewline
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                9
                                    {
                                        WriteWordLine 0 2 "Software Updates"
                                        WriteWordLine 0 2 "Enable software updates on clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Software Update scan schedule: " -nonewline
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 2 "Schedule deployment re-evaluation: " -nonewline
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 0 " Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 2 "When any software update deployment deadline is reached, install all other software update deployments with deadline coming within a specified period of time: " -nonewline
                                        if ($AgentConfig.AssignmentBatchingTimeout -eq "0")
                                            {
                                                WriteWordLine 0 0 "No."
                                            }
                                        else 
                                            {
                                                WriteWordLine 0 0 "Yes."    
                                                WriteWordLine 0 2 "Period of time for which all pending deployments with deadline in this time will also be installed: " -nonewline
                                                if ($AgentConfig.AssignmentBatchingTimeout -le "82800")
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

                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                10
                                    {
                                        WriteWordLine 0 2 "User and Device Affinity"
                                        WriteWordLine 0 2 "User device affinity usage threshold (minutes): $($AgentConfig.ConsoleMinutes)"
                                        WriteWordLine 0 2 "User device affinity usage threshold (days): $($AgentConfig.IntervalDays)"
                                        WriteWordLine 0 2 "Automatically configure user device affinity from usage data: " -nonewline 
                                        if ($AgentConfig.AutoApproveAffinity -eq "0")
                                            {
                                                WriteWordLine 0 0 "No"
                                            }
                                        else
                                            {
                                                WriteWordLine 0 0 "Yes"
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                11
                                    {
                                        WriteWordLine 0 2 "Background Intelligent Transfer"
                                        WriteWordLine 0 2 "Limit the maximum network bandwidth for BITS background transfers: $($AgentConfig.EnableBitsMaxBandwidth)"
                                        WriteWordLine 0 2 "Throttling window start time: $($AgentConfig.MaxBandwidthValidFrom)"
                                        WriteWordLine 0 2 "Throttling window end time: $($AgentConfig.MaxBandwidthValidTo)"
                                        WriteWordLine 0 2 "Maximum transfer rate during throttling window (kbps): $($AgentConfig.MaxTransferRateOnSchedule)"
                                        WriteWordLine 0 2 "Allow BITS downloads outside the throttling window: $($AgentConfig.EnableDownloadOffSchedule)"
                                        WriteWordLine 0 2 "Maximum transfer rate outside the throttling window (Kbps): $($AgentConfig.MaxTransferRateOffSchedule)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                12
                                    {
                                        WriteWordLine 0 2 "Enrollment"
                                        WriteWordLine 0 2 "Allow users to enroll mobile devices and Mac computers: " -nonewline
                                        if ($AgentConfig.EnableDeviceEnrollment -eq "0")
                                            {
                                                WriteWordLine 0 0 "No"
                                            }
                                        else
                                            {
                                                WriteWordLine 0 0 "Yes"
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                13
                                    {
                                        WriteWordLine 0 2 "Client Policy"
                                        WriteWordLine 0 2 "Client policy polling interval (minutes): $($AgentConfig.PolicyRequestAssignmentTimeout)"
                                        WriteWordLine 0 2 "Enable user policy on clients: $($AgentConfig.PolicyEnableUserPolicyPolling)"
                                        WriteWordLine 0 2 "Enable user policy requests from Internet clients: $($AgentConfig.PolicyEnableUserPolicyOnInternet)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                15
                                    {
                                        WriteWordLine 0 2 "Hardware Inventory"
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 2 "Hardware inventory schedule: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                16 
                                    {
                                        WriteWordLine 0 2 "State Messaging"
                                        WriteWordLine 0 2 "State message reporting cycle (minutes): $($AgentConfig.BulkSendInterval)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                17
                                    {
                                        WriteWordLine 0 2 "Software Deployment"
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 2 "Schedule re-evaluation for deployments: Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)"
                                                    }
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                18
                                    {
                                        WriteWordLine 0 2 "Power Management"
                                        WriteWordLine 0 2 "Allow power management of clients: $($AgentConfig.Enabled)"
                                        WriteWordLine 0 2 "Allow users to exclude their device from power management: $($AgentConfig.AllowUserToOptOutFromPowerPlan)"
                                        WriteWordLine 0 2 "Enable wake-up proxy: $($AgentConfig.EnableWakeupProxy)"
                                        if ($AgentConfig.EnableWakeupProxy -eq "True")
                                            {
                                                WriteWordLine 0 2 "Wake-up proxy port number (UDP): $($AgentConfig.Port)"
                                                WriteWordLine 0 2 "Wake On LAN port number (UDP): $($AgentConfig.WolPort)"
                                                WriteWordLine 0 2 "Windows Firewall exception for wake-up proxy: " -nonewline
                                                switch ($AgentConfig.WakeupProxyFirewallFlags)
                                                    {
                                                        0 { WriteWordLine 0 2 "disabled" }
                                                        9 { WriteWordLine 0 2 "Enabled: Public." }
                                                        10 { WriteWordLine 0 2 "Enabled: Private." }
                                                        11 { WriteWordLine 0 2 "Enabled: Private, Public." }
                                                        12 { WriteWordLine 0 2 "Enabled: Domain." }
                                                        13 { WriteWordLine 0 2 "Enabled: Domain, Public." }
                                                        14 { WriteWordLine 0 2 "Enabled: Domain, Private." }
                                                        15 { WriteWordLine 0 2 "Enabled: Domain, Private, Public." }
                                                    }
                                                WriteWordLine 0 2 "IPv6 prefixes if required for DirectAccess or other intervening network devices. Use a comma to specifiy multiple entries: $($AgentConfig.WakeupProxyDirectAccessPrefixList)"
                                            }
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                20
                                    {
                                        WriteWordLine 0 2 "Endpoint Protection"
                                        WriteWordLine 0 2 "Manage Endpoint Protection client on client computers: $($AgentConfig.EnableEP)"
                                        WriteWordLine 0 2 "Install Endpoint Protection client on client computers: $($AgentConfig.InstallSCEPClient)"
                                        WriteWordLine 0 2 "Automatically remove previously installed antimalware software before Endpoint Protection is installed: $($AgentConfig.Remove3rdParty)"
                                        WriteWordLine 0 2 "Allow Endpoint Protection client installation and restarts outside maintenance windows. Maintenance windows must be at least 30 minutes long for client installation: $($AgentConfig.OverrideMaintenanceWindow)"
                                        WriteWordLine 0 2 "For Windows Embedded devices with write filters, commit Endpoint Protection client installation (requires restart): $($AgentConfig.PersistInstallation)"
                                        WriteWordLine 0 2 "Suppress any required computer restarts after the Endpoint Protection client is installed: $($AgentConfig.SuppressReboot)"
                                        WriteWordLine 0 2 "Allowed period of time users can postpone a required restart to complete the Endpoint Protection installation (hours): $($AgentConfig.ForceRebootPeriod)"
                                        WriteWordLine 0 2 "Disable alternate sources (such as Microsoft Windows Update, Microsoft Windows Server Update Services, or UNC shares) for the initial definition update on client computers: $($AgentConfig.DisableFirstSignatureUpdate)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                21
                                    {
                                        WriteWordLine 0 2 "Computer Restart"
                                        WriteWordLine 0 2 "Display a temporary notification to the user that indicates the interval before the user is logged of or the computer restarts (minutes): $($AgentConfig.RebootLogoffNotificationCountdownDuration)"
                                        WriteWordLine 0 2 "Display a dialog box that the user cannot close, which displays the countdown interval before the user is logged of or the computer restarts (minutes): $([string]$AgentConfig.RebootLogoffNotificationFinalWindow / 60)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                22
                                    {
                                        WriteWordLine 0 2 "Cloud Services"
                                        WriteWordLine 0 2 "Allow access to Cloud Distribution Point: $($AgentConfig.AllowCloudDP)"
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 0 "---------------------"
                                    }
                                23
                                    {
                                        WriteWordLine 0 2 "Metered Internet Connections"
                                        switch ($AgentConfig.MeteredNetworkUsage)
                                            {
                                                1 { $Usage = "Allow" }
                                                2 { $Usage = "Limit" }
                                                4 { $Usage = "Block" }
                                            }
                                        WriteWordLine 0 2 "Specifiy how clients communicate on metered network connections: $($Usage)"
                                        WriteWordLine 0 0 ""
                                    }

                            }
            }
        catch [System.Management.Automation.PropertyNotFoundException] 
            {
                WriteWordLine 0 0 ""
            }
    }
}
#### Security
Write-Verbose "$(Get-Date):   Collecting all administrative users"
WriteWordLine 2 0 "Administrative Users"
$Admins = Get-CMAdministrativeUser

    WriteWordLine 0 1 "Enumerating administrative users:"
    $Table = $Null
    $TableRange = $Null
    $TableRange = $doc.Application.Selection.Range
	$Columns = 5
    [int]$Rows = $Admins.count + 1
	Write-Verbose "$(Get-Date):   add Admin properties to table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $TableStyle 
	$table.Borders.InsideLineStyle = 1
	$table.Borders.OutsideLineStyle = 1
	[int]$xRow = 1
	Write-Verbose "$(Get-Date):   format first row with column headings"
	
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Font.Size = "10"
	$Table.Cell($xRow,1).Range.Text = "Account name"
	
	$Table.Cell($xRow,2).Range.Font.Bold = $True
	$Table.Cell($xRow,2).Range.Font.Size = "10"
	$Table.Cell($xRow,2).Range.Text = "Account Type"
	
	$Table.Cell($xRow,3).Range.Font.Bold = $True
	$Table.Cell($xRow,3).Range.Font.Size = "10"
	$Table.Cell($xRow,3).Range.Text = "Security Roles"
    
	$Table.Cell($xRow,4).Range.Font.Bold = $True
	$Table.Cell($xRow,4).Range.Font.Size = "10"
	$Table.Cell($xRow,4).Range.Text = "Security Scopes"
    
	$Table.Cell($xRow,5).Range.Font.Bold = $True
	$Table.Cell($xRow,5).Range.Font.Size = "10"
	$Table.Cell($xRow,5).Range.Text = "Collections"                      
    foreach ($Admin in $Admins)
		{
			$xRow++							
			$Table.Cell($xRow,1).Range.Font.Size = "10"
			$Table.Cell($xRow,1).Range.Text = $Admin.LogonName
			$Table.Cell($xRow,2).Range.Font.Size = "10"
            switch ($Admin.AccountType)
                {
                    0 { $Table.Cell($xRow,2).Range.Text = "User" }
                    1 { $Table.Cell($xRow,2).Range.Text = "Group" }
                    2 { $Table.Cell($xRow,2).Range.Text = "Machine" } 
                } 
			$Table.Cell($xRow,3).Range.Font.Size = "10"
			$Table.Cell($xRow,3).Range.Text = $Admin.RoleNames
            $Table.Cell($xRow,4).Range.Font.Size = "10"
			$Table.Cell($xRow,4).Range.Text = $Admin.CategoryNames
            $Table.Cell($xRow,5).Range.Font.Size = "10"
			$Table.Cell($xRow,5).Range.Text = $Admin.CollectionNames
		}
				
	$Table.Rows.SetLeftIndent(50,1) | Out-Null
	$table.AutoFitBehavior(1) | Out-Null

	#return focus back to document
	Write-Verbose "$(Get-Date):   return focus back to document"
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
    
#move to the end of the current document
Write-Verbose "$(Get-Date):   move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
WriteWordLine 0 0 ""

#### enumerating all custom Security roles
Write-Verbose "$(Get-Date):   enumerating all custom build security roles"
WriteWordLine 2 0 "Custom Security Roles"
$SecurityRoles = Get-CMSecurityRole | Where-Object {-not $_.IsBuiltIn}
if (-not [string]::IsNullOrEmpty($SecurityRoles ))
    {
        WriteWordLine 0 1 "Enumerating all custom build security roles:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
	    $Columns = 5
        [int]$Rows = $SecurityRoles.count + 1
	    Write-Verbose "$(Get-Date):   add security role properties to table"
	    $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	    $table.Style = $TableStyle 
	    $table.Borders.InsideLineStyle = 1
	    $table.Borders.OutsideLineStyle = 1
	    [int]$xRow = 1
	    Write-Verbose "$(Get-Date):   format first row with column headings"
	    
	    $Table.Cell($xRow,1).Range.Font.Bold = $True
	    $Table.Cell($xRow,1).Range.Font.Size = "10"
	    $Table.Cell($xRow,1).Range.Text = "Name"
	    
	    $Table.Cell($xRow,2).Range.Font.Bold = $True
	    $Table.Cell($xRow,2).Range.Font.Size = "10"
	    $Table.Cell($xRow,2).Range.Text = "Description"
	    
	    $Table.Cell($xRow,3).Range.Font.Bold = $True
	    $Table.Cell($xRow,3).Range.Font.Size = "10"
	    $Table.Cell($xRow,3).Range.Text = "Copied from"
        
	    $Table.Cell($xRow,4).Range.Font.Bold = $True
	    $Table.Cell($xRow,4).Range.Font.Size = "10"
	    $Table.Cell($xRow,4).Range.Text = "Members"
        
	    $Table.Cell($xRow,5).Range.Font.Bold = $True
	    $Table.Cell($xRow,5).Range.Font.Size = "10"
	    $Table.Cell($xRow,5).Range.Text = "Role ID"                      
        foreach ($SecurityRole in $SecurityRoles)
		    {
			    $xRow++							
			    $Table.Cell($xRow,1).Range.Font.Size = "10"
			    $Table.Cell($xRow,1).Range.Text = $SecurityRole.RoleName
			    $Table.Cell($xRow,2).Range.Font.Size = "10"
                $Table.Cell($xRow,2).Range.Text = $SecurityRole.RoleDescription
			    $Table.Cell($xRow,3).Range.Font.Size = "10"
			    $Table.Cell($xRow,3).Range.Text = (Get-CMSecurityRole -Id $SecurityRole.CopiedFromID).RoleName
                $Table.Cell($xRow,4).Range.Font.Size = "10"
                if ($SecurityRole.NumberOfAdmins -gt 0)
			        {
                        $Table.Cell($xRow,4).Range.Text = (Get-CMAdministrativeUser | Where-Object {$_.Roles -ilike "$($SecurityRole.RoleID)"}).LogonName
                    }
                $Table.Cell($xRow,5).Range.Font.Size = "10"
			    $Table.Cell($xRow,5).Range.Text = $SecurityRole.RoleID
		    }
				
	    $Table.Rows.SetLeftIndent(30,1) | Out-Null
	    $table.AutoFitBehavior(1) | Out-Null

	    #return focus back to document
	    Write-Verbose "$(Get-Date):   return focus back to document"
	    $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
    
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    }
else
    {
        WriteWordLine 0 1 "There are no custom built security roles."
    }

#### Used Accounts
Write-Verbose "$(Get-Date):   Enumerating all used accounts"
WriteWordLine 2 0 "Configured Accounts"
$Accounts = Get-CMAccount
WriteWordLine 0 1 "Enumerating all accounts used for specific tasks."
    $Table = $Null
    $TableRange = $Null
    $TableRange = $doc.Application.Selection.Range
	$Columns = 3
    [int]$Rows = $Accounts.count + 1
	Write-Verbose "$(Get-Date):   add security role properties to table"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $TableStyle 
	$table.Borders.InsideLineStyle = 1
	$table.Borders.OutsideLineStyle = 1
	[int]$xRow = 1
	Write-Verbose "$(Get-Date):   format first row with column headings"
	
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Font.Size = "10"
	$Table.Cell($xRow,1).Range.Text = "User Name"
	
	$Table.Cell($xRow,2).Range.Font.Bold = $True
	$Table.Cell($xRow,2).Range.Font.Size = "10"
	$Table.Cell($xRow,2).Range.Text = "Account Usage"
	
	$Table.Cell($xRow,3).Range.Font.Bold = $True
	$Table.Cell($xRow,3).Range.Font.Size = "10"
	$Table.Cell($xRow,3).Range.Text = "Site Code"                     
    foreach ($Account in $Accounts)
		{
			$xRow++							
			$Table.Cell($xRow,1).Range.Font.Size = "10"
			$Table.Cell($xRow,1).Range.Text = $Account.UserName
			$Table.Cell($xRow,2).Range.Font.Size = "10"
            $Table.Cell($xRow,2).Range.Text = $Account.AccountUsage
			$Table.Cell($xRow,3).Range.Font.Size = "10"
			$Table.Cell($xRow,3).Range.Text = $Account.SiteCode
		}
				
	$Table.Rows.SetLeftIndent(30,1) | Out-Null
	$table.AutoFitBehavior(1) | Out-Null

	#return focus back to document
	Write-Verbose "$(Get-Date):   return focus back to document"
	$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
    
#move to the end of the current document
Write-Verbose "$(Get-Date):   move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
WriteWordLine 0 0 ""
############################################################################################

####
#### Assets and Compliance
####
Write-Verbose "$(Get-Date):   Done with Administration, next Assets and Compliance"
WriteWordLine 1 0 "Assets and Compliance"

#### enumerating all User Collections
WriteWordLine 2 0 "Summary of User Collections"
$UserCollections = Get-CMUserCollection
if ($ListAllInformation)
    {
        foreach ($UserCollection in $UserCollections)
            {
                Write-Verbose "$(Get-Date):   Found User Collection: $($UserCollection.Name)"
                WriteWordLine 0 1 "Collection Name: $($UserCollection.Name)" -bold
                WriteWordLine 0 1 "Collection ID: $($UserCollection.CollectionID)"
                WriteWordLine 0 1 "Total count of members: $($UserCollection.MemberCount)"
                WriteWordLine 0 1 "Limited to User Collection: $($UserCollection.LimitToCollectionName) / $($UserCollection.LimitToCollectionID)"
                WriteWordLine 0 0 ""
            }
    }
else
    {
     WriteWordLine 0 1 "There are $($UserCollections.count) User Collections." 
    }

####
#### enumerating all Device Collections
WriteWordLine 2 0 "Summary of Device Collections"
$DeviceCollections = Get-CMDeviceCollection
if ($ListAllInformation)
    {
        foreach ($DeviceCollection in $DeviceCollections)
            {
                Write-Verbose "$(Get-Date):   Found Device Collection: $($DeviceCollection.Name)"
                WriteWordLine 0 1 "Collection Name: $($DeviceCollection.Name)" -bold
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
                                WriteWordLine 0 2 "Checking Maintenance Windows on Collection:" 
                                #$ServiceWindows = [wmi]$ServiceWindows.__PATH
                        
                                foreach ($ServiceWindow in $ServiceWindows)
                                    {
                
                                        $ScheduleString = Read-ScheduleToken
                                        $startTime = $ScheduleString.TokenData.starttime
                                        $startTime = Convert-NormalDateToConfigMgrDate -starttime $startTime
                                        WriteWordLine 0 3 "Maintenance window name: $($ServiceWindow.Name)"
                                        switch ($ServiceWindow.ServiceWindowType)
                                            {
                                                0 {WriteWordLine 0 3 "This is a Task Sequence maintenance window"}
                                                1 {WriteWordLine 0 3 "This is a general maintenance window"}                        
                                            }   
                                        switch ($ServiceWindow.RecurrenceType)
                                            {
                                                1 {WriteWordLine 0 3 "This maintenance window occurs only once on $($startTime) and lasts for $($ScheduleString.TokenData.HourDuration) hour(s) and $($ScheduleString.TokenData.MinuteDuration) minute(s)."}
                                                2 
                                                    {
                                                        if ($ScheduleString.TokenData.DaySpan -eq "1")
                                                            {
                                                                $daily = "daily"
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
                                                                0 {$order = "last"}
                                                                1 {$order = "first"}
                                                                2 {$order = "second"}
                                                                3 {$order = "third"}
                                                                4 {$order = "fourth"}
                                                            }
                                                        WriteWordLine 0 3 "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofMonths) month(s) on every $($order) $(Convert-WeekDay $ScheduleString.TokenData.Day)"
                                                    }

                                                5 
                                                    {
                                                        if ($ScheduleString.TokenData.MonthDay -eq "0")
                                                            { 
                                                                $DayOfMonth = "the last day of the month"
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
                                                true {WriteWordLine 0 3 "The maintenance window is enabled"}
                                                false {WriteWordLine 0 3 "The maintenance window is disabled"}
                                            }
                                    }
                            }
                        else
                            {
                                WriteWordLine 0 2 "No maintenance windows configured on this collection."
                            }
                    }  
                        try
                            {
                                $CollVars = $CollSettings.CollectionVariables               
                                if (-not [string]::IsNullOrEmpty($CollVars))
                                    {
                                        WriteWordLine 0 1 "Enumerating device collection variables:"
                                        $Table = $Null
                                        $TableRange = $Null
                                        $TableRange = $doc.Application.Selection.Range
				                        $Columns = 3
                                        [int]$Rows = $CollVars.count + 1
				                        Write-Verbose "$(Get-Date):   add Collection variables to table"
				                        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				                        $table.Style = $TableStyle 
				                        $table.Borders.InsideLineStyle = 1
				                        $table.Borders.OutsideLineStyle = 1
				                        [int]$xRow = 1
				                        Write-Verbose "$(Get-Date):   format first row with column headings"
				                        
				                        $Table.Cell($xRow,1).Range.Font.Bold = $True
				                        $Table.Cell($xRow,1).Range.Font.Size = "10"
				                        $Table.Cell($xRow,1).Range.Text = "Variable name"
				                        
				                        $Table.Cell($xRow,2).Range.Font.Bold = $True
				                        $Table.Cell($xRow,2).Range.Font.Size = "10"
				                        $Table.Cell($xRow,2).Range.Text = "Value"
				                        
				                        $Table.Cell($xRow,3).Range.Font.Bold = $True
				                        $Table.Cell($xRow,3).Range.Font.Size = "10"
				                        $Table.Cell($xRow,3).Range.Text = "Is Masked"                      
                                        foreach ($CollVar in $CollVars)
				                            {
					                            $xRow++							
					                            $Table.Cell($xRow,1).Range.Font.Size = "10"
					                            $Table.Cell($xRow,1).Range.Text = $CollVar.Name
					                            $Table.Cell($xRow,2).Range.Font.Size = "10"
					                            $Table.Cell($xRow,2).Range.Text = $CollVar.Value
					                            $Table.Cell($xRow,3).Range.Font.Size = "10"
					                            $Table.Cell($xRow,3).Range.Text = $CollVar.IsMasked
					                        }
				
				                        $Table.Rows.SetLeftIndent(50,1) | Out-Null
				                        $table.AutoFitBehavior(1) | Out-Null

				                        #return focus back to document
				                        Write-Verbose "$(Get-Date):   return focus back to document"
				                        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                                        #move to the end of the current document
				                        Write-Verbose "$(Get-Date):   move to the end of the current document"
				                        $selection.EndKey($wdStory,$wdMove) | Out-Null
				                        WriteWordLine 0 0 ""
                                    }
                                else 
                                    {
                                        WriteWordLine 0 1 "Enumerating device collection variables: No device collection variables configured!"
                                    }
                        }
                    catch [System.Management.Automation.PropertyNotFoundException] 
                            {
                                WriteWordLine 0 0 ""
                            }
            ### enumerating the Collection Membership Rules
                    $QueryRules = $Null
                    $DirectRules = $Null
                    $IncludeRules = $Null
                    $CollectionRules = $DeviceCollection.CollectionRules #just for Direct and Query
                    
                    $Collection = Get-WmiObject -Namespace root\sms\site_$SiteCode -Query "SELECT * FROM SMS_Collection WHERE CollectionID = '$($DeviceCollection.CollectionID)'"
                    [wmi]$Collection = $Collection.__PATH
                    
                    $OtherCollectionRules = $Collection.CollectionRules
                    try 
                        {
                            $DirectRules = $CollectionRules | where {$_.ResourceID} -ErrorAction SilentlyContinue
                        }
                    catch [System.Management.Automation.PropertyNotFoundException] 
                            {
                                WriteWordLine 0 0 ""
                            }
                    try
                        {
                            $QueryRules = $CollectionRules | where {$_.QueryExpression} -ErrorAction SilentlyContinue                            
                        }
                    catch [System.Management.Automation.PropertyNotFoundException] 
                        {
                            WriteWordLine 0 0 ""
                        }
                    try 
                        {
                            $IncludeRules = $OtherCollectionRules | where {$_.IncludeCollectionID} -ErrorAction SilentlyContinue
                        }
                    catch [System.Management.Automation.PropertyNotFoundException] 
                            {
                                WriteWordLine 0 0 ""
                            }
            if ($QueryRules)
                    {
                        
                        WriteWordLine 0 1 "Enumerating device collection query membership rules:"
                        $Table = $Null
                        $TableRange = $Null
                        $TableRange = $doc.Application.Selection.Range
				        $Columns = 3
                        [int]$Rows = $QueryRules.count + 1
				        Write-Verbose "$(Get-Date):   add Collection variables to table"
				        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				        $table.Style = $TableStyle 
				        $table.Borders.InsideLineStyle = 1
				        $table.Borders.OutsideLineStyle = 1
				        [int]$xRow = 1
				        Write-Verbose "$(Get-Date):   format first row with column headings"
				        
				        $Table.Cell($xRow,1).Range.Font.Bold = $True
				        $Table.Cell($xRow,1).Range.Font.Size = "10"
				        $Table.Cell($xRow,1).Range.Text = "Query name"
				        
				        $Table.Cell($xRow,2).Range.Font.Bold = $True
				        $Table.Cell($xRow,2).Range.Font.Size = "10"
				        $Table.Cell($xRow,2).Range.Text = "Query Expression"
				        
				        $Table.Cell($xRow,3).Range.Font.Bold = $True
				        $Table.Cell($xRow,3).Range.Font.Size = "10"
				        $Table.Cell($xRow,3).Range.Text = "Query ID"
                        foreach ($QueryRule in $QueryRules)
                            {
                                $xRow++							
					            $Table.Cell($xRow,1).Range.Font.Size = "10"
					            $Table.Cell($xRow,1).Range.Text = $QueryRule.RuleName
					            $Table.Cell($xRow,2).Range.Font.Size = "10"
					            $Table.Cell($xRow,2).Range.Text = $QueryRule.QueryExpression
					            $Table.Cell($xRow,3).Range.Font.Size = "10"
					            $Table.Cell($xRow,3).Range.Text = $QueryRule.QueryID    
                            }				
				        $Table.Rows.SetLeftIndent(50,1) | Out-Null
				        $table.AutoFitBehavior(1) | Out-Null
				        #return focus back to document
				        Write-Verbose "$(Get-Date):   return focus back to document"
				        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                        #move to the end of the current document
			            Write-Verbose "$(Get-Date):   move to the end of the current document"
			            $selection.EndKey($wdStory,$wdMove) | Out-Null
			            WriteWordLine 0 0 ""
                    }
            if ($DirectRules)
                    {
                        WriteWordLine 0 1 "Enumerating device collection direct membership rules:"
                        $Table = $Null
                        $TableRange = $Null
                        $TableRange = $doc.Application.Selection.Range
				        $Columns = 2
                        [int]$Rows = $DirectRules.count + 1
				        Write-Verbose "$(Get-Date):   add Collection variables to table"
				        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				        $table.Style = $TableStyle 
				        $table.Borders.InsideLineStyle = 1
				        $table.Borders.OutsideLineStyle = 1
				        [int]$xRow = 1
				        Write-Verbose "$(Get-Date):   format first row with column headings"
				            
				        $Table.Cell($xRow,1).Range.Font.Bold = $True
				        $Table.Cell($xRow,1).Range.Font.Size = "10"
				        $Table.Cell($xRow,1).Range.Text = "Resource name"
				            
				        $Table.Cell($xRow,2).Range.Font.Bold = $True
				        $Table.Cell($xRow,2).Range.Font.Size = "10"
				        $Table.Cell($xRow,2).Range.Text = "Resource ID"
                        foreach ($DirectRule in $DirectRules)
                            {
                                $xRow++							
					            $Table.Cell($xRow,1).Range.Font.Size = "10"
					            $Table.Cell($xRow,1).Range.Text = $DirectRule.RuleName
					            $Table.Cell($xRow,2).Range.Font.Size = "10"
					            $Table.Cell($xRow,2).Range.Text = $DirectRule.ResourceID   
                            }				
				        $Table.Rows.SetLeftIndent(50,1) | Out-Null
				        $table.AutoFitBehavior(1) | Out-Null
				        #return focus back to document
				        Write-Verbose "$(Get-Date):   return focus back to document"
				        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                        #move to the end of the current document
			            Write-Verbose "$(Get-Date):   move to the end of the current document"
			            $selection.EndKey($wdStory,$wdMove) | Out-Null
			            WriteWordLine 0 0 ""
                    }
                else 
                    {
                        WriteWordLine 0 1 "Enumerating device collection membership rules: No device collection direct membership rules configured!"
                    }
                if ($IncludeRules)
                    {
                        WriteWordLine 0 1 "Enumerating device collection Include Collection membership rules:"
                        $Table = $Null
                        $TableRange = $Null
                        $TableRange = $doc.Application.Selection.Range
				        $Columns = 2
                        [int]$Rows = $IncludeRules.count + 1
				        Write-Verbose "$(Get-Date):   add Collection variables to table"
				        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				        $table.Style = $TableStyle 
				        $table.Borders.InsideLineStyle = 1
				        $table.Borders.OutsideLineStyle = 1
				        [int]$xRow = 1
				        Write-Verbose "$(Get-Date):   format first row with column headings"
				            
				        $Table.Cell($xRow,1).Range.Font.Bold = $True
				        $Table.Cell($xRow,1).Range.Font.Size = "10"
				        $Table.Cell($xRow,1).Range.Text = "Collection ID"
				            
				        $Table.Cell($xRow,2).Range.Font.Bold = $True
				        $Table.Cell($xRow,2).Range.Font.Size = "10"
				        $Table.Cell($xRow,2).Range.Text = "Collection Name"
                        foreach ($IncludeRule in $IncludeRules)
                            {
                                $xRow++							
					            $Table.Cell($xRow,1).Range.Font.Size = "10"
					            $Table.Cell($xRow,1).Range.Text = $IncludeRule.IncludeCollectionID
					            $Table.Cell($xRow,2).Range.Font.Size = "10"
					            $Table.Cell($xRow,2).Range.Text = $IncludeRule.RuleName   
                            }				
				        $Table.Rows.SetLeftIndent(50,1) | Out-Null
				        $table.AutoFitBehavior(1) | Out-Null
				        #return focus back to document
				        Write-Verbose "$(Get-Date):   return focus back to document"
				        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                        #move to the end of the current document
			            Write-Verbose "$(Get-Date):   move to the end of the current document"
			            $selection.EndKey($wdStory,$wdMove) | Out-Null
			            WriteWordLine 0 0 ""
                    }
                else 
                    {
                        WriteWordLine 0 1 "Enumerating device collection membership rules: No device collection Include Collection membership rules configured!"
                    }
    
			        #move to the end of the current document
			        Write-Verbose "$(Get-Date):   move to the end of the current document"
			        $selection.EndKey($wdStory,$wdMove) | Out-Null
			        WriteWordLine 0 0 ""
            }
    }

else
    {
        WriteWordLine 0 1 "There are $($DeviceCollections.count) Device collections."
    }

Write-Verbose "$(Get-Date):   Working on Compliance Settings"
WriteWordLine 2 0 "Compliance Settings"
WriteWordLine 0 0 ""
WriteWordLine 3 0 "Configuration Items"

$CIs = Get-CMConfigurationItem
WriteWordLine 0 1 "Enumerating Configuration Items:"
$Table = $Null
$TableRange = $Null
$TableRange = $doc.Application.Selection.Range
$Columns = 4
[int]$Rows = $CIs.count + 1
Write-Verbose "$(Get-Date):   add configuration items to table"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = $TableStyle 
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1
[int]$xRow = 1
Write-Verbose "$(Get-Date):   format first row with column headings"
    
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Font.Size = "10"
$Table.Cell($xRow,1).Range.Text = "Name"
    
$Table.Cell($xRow,2).Range.Font.Bold = $True
$Table.Cell($xRow,2).Range.Font.Size = "10"
$Table.Cell($xRow,2).Range.Text = "Last modified"
    
$Table.Cell($xRow,3).Range.Font.Bold = $True
$Table.Cell($xRow,3).Range.Font.Size = "10"
$Table.Cell($xRow,3).Range.Text = "Last modified by"
    
$Table.Cell($xRow,4).Range.Font.Bold = $True
$Table.Cell($xRow,4).Range.Font.Size = "10"
$Table.Cell($xRow,4).Range.Text = "CI ID"
foreach ($CI in $CIs)
    {
        $xRow++							
		$Table.Cell($xRow,1).Range.Font.Size = "10"
		$Table.Cell($xRow,1).Range.Text = $CI.LocalizedDisplayName
		$Table.Cell($xRow,2).Range.Font.Size = "10"
		$Table.Cell($xRow,2).Range.Text = $CI.DateLastModified
        $Table.Cell($xRow,3).Range.Font.Size = "10"
		$Table.Cell($xRow,3).Range.Text = $CI.LastModifiedBy
        $Table.Cell($xRow,4).Range.Font.Size = "10"
		$Table.Cell($xRow,4).Range.Text = $CI.CI_ID   
    }				
$Table.Rows.SetLeftIndent(50,1) | Out-Null
$table.AutoFitBehavior(1) | Out-Null
#return focus back to document
Write-Verbose "$(Get-Date):   return focus back to document"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
#move to the end of the current document
Write-Verbose "$(Get-Date):   move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
WriteWordLine 0 0 ""
    
#move to the end of the current document
Write-Verbose "$(Get-Date):   move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
WriteWordLine 0 0 ""

WriteWordLine 3 0 "Configuration Baselines"
$CBs = Get-CMBaseline
if (-not [string]::IsNullOrEmpty($CBs))
    {
        WriteWordLine 0 1 "Enumerating Configuration Baselines:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
        $Columns = 4
        [int]$Rows = $CBs.count + 1
        Write-Verbose "$(Get-Date):   add configuration items to table"
        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
        $table.Style = $TableStyle 
        $table.Borders.InsideLineStyle = 1
        $table.Borders.OutsideLineStyle = 1
        [int]$xRow = 1
        Write-Verbose "$(Get-Date):   format first row with column headings"
        
        $Table.Cell($xRow,1).Range.Font.Bold = $True
        $Table.Cell($xRow,1).Range.Font.Size = "10"
        $Table.Cell($xRow,1).Range.Text = "Name"
        
        $Table.Cell($xRow,2).Range.Font.Bold = $True
        $Table.Cell($xRow,2).Range.Font.Size = "10"
        $Table.Cell($xRow,2).Range.Text = "Last modified"
        
        $Table.Cell($xRow,3).Range.Font.Bold = $True
        $Table.Cell($xRow,3).Range.Font.Size = "10"
        $Table.Cell($xRow,3).Range.Text = "Last modified by"
        
        $Table.Cell($xRow,4).Range.Font.Bold = $True
        $Table.Cell($xRow,4).Range.Font.Size = "10"
        $Table.Cell($xRow,4).Range.Text = "CI ID"
        foreach ($CB in $CBs)
            {
                $xRow++							
		        $Table.Cell($xRow,1).Range.Font.Size = "10"
		        $Table.Cell($xRow,1).Range.Text = $CB.LocalizedDisplayName
		        $Table.Cell($xRow,2).Range.Font.Size = "10"
		        $Table.Cell($xRow,2).Range.Text = $CB.DateLastModified
                $Table.Cell($xRow,3).Range.Font.Size = "10"
		        $Table.Cell($xRow,3).Range.Text = $CB.LastModifiedBy
                $Table.Cell($xRow,4).Range.Font.Size = "10"
		        $Table.Cell($xRow,4).Range.Text = $CB.CI_ID   
            }				
        $Table.Rows.SetLeftIndent(50,1) | Out-Null
        $table.AutoFitBehavior(1) | Out-Null
        #return focus back to document
        Write-Verbose "$(Get-Date):   return focus back to document"
        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    }
else
    {
        WriteWordLine 0 1 "There are no Configuration Baselines configured."
    }

### User Data and Profiles
Write-Verbose "$(Get-Date):   Working on User Data and Profiles"
WriteWordLine 3 0 "User Data and Profiles"
$UserDataProfiles = Get-CMUserDataAndProfileConfigurationItem
if (-not [string]::IsNullOrEmpty($UserDataProfiles))
    {
        WriteWordLine 0 1 "Enumerating User Data and Profiles:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
        $Columns = 4
        [int]$Rows = $UserDataProfiles.count + 1
        Write-Verbose "$(Get-Date):   add configuration items to table"
        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
        $table.Style = $TableStyle 
        $table.Borders.InsideLineStyle = 1
        $table.Borders.OutsideLineStyle = 1
        [int]$xRow = 1
        Write-Verbose "$(Get-Date):   format first row with column headings"
        
        $Table.Cell($xRow,1).Range.Font.Bold = $True
        $Table.Cell($xRow,1).Range.Font.Size = "10"
        $Table.Cell($xRow,1).Range.Text = "Name"
        
        $Table.Cell($xRow,2).Range.Font.Bold = $True
        $Table.Cell($xRow,2).Range.Font.Size = "10"
        $Table.Cell($xRow,2).Range.Text = "Last modified"
        
        $Table.Cell($xRow,3).Range.Font.Bold = $True
        $Table.Cell($xRow,3).Range.Font.Size = "10"
        $Table.Cell($xRow,3).Range.Text = "Last modified by"
        
        $Table.Cell($xRow,4).Range.Font.Bold = $True
        $Table.Cell($xRow,4).Range.Font.Size = "10"
        $Table.Cell($xRow,4).Range.Text = "CI ID"
        foreach ($UserDataProfile in $UserDataProfiles)
            {
                $xRow++							
		        $Table.Cell($xRow,1).Range.Font.Size = "10"
		        $Table.Cell($xRow,1).Range.Text = $UserDataProfile.LocalizedDisplayName
		        $Table.Cell($xRow,2).Range.Font.Size = "10"
		        $Table.Cell($xRow,2).Range.Text = $UserDataProfile.DateLastModified
                $Table.Cell($xRow,3).Range.Font.Size = "10"
		        $Table.Cell($xRow,3).Range.Text = $UserDataProfile.LastModifiedBy
                $Table.Cell($xRow,4).Range.Font.Size = "10"
		        $Table.Cell($xRow,4).Range.Text = $UserDataProfile.CI_ID   
            }				
        $Table.Rows.SetLeftIndent(50,1) | Out-Null
        $table.AutoFitBehavior(1) | Out-Null
        #return focus back to document
        Write-Verbose "$(Get-Date):   return focus back to document"
        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    }
else
    {
        WriteWordLine 0 1 "There are no User Data and Profile configurations configured."
    }


Write-Verbose "$(Get-Date):   Working on Endpoint Protection"
WriteWordLine 2 0 "Endpoint Protection"
if (-not ($(Get-CMEndpointProtectionPoint) -eq $Null))
    {
        WriteWordLine 3 0 "Antimalware Policies"
        $AntiMalwarePolicies = Get-CMAntimalwarePolicy
        if (-not [string]::IsNullOrEmpty($AntiMalwarePolicies))
            {
                foreach ($AntiMalwarePolicy in $AntiMalwarePolicies)
                    {
                        if ($AntiMalwarePolicy.Name -eq "Default Client Antimalware Policy")
                            {
                                $AgentConfig = $AntiMalwarePolicy.AgentConfiguration
                                WriteWordLine 0 1 "$($AntiMalwarePolicy.Name)" -bold
                                WriteWordLine 0 1 "Description: $($AntiMalwarePolicy.Description)"
                                WriteWordLine 0 2 "Scheduled Scans" -bold
                                WriteWordLine 0 3 "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                if ($AgentConfig.EnableScheduledScan)
                                    {
                                        switch ($AgentConfig.ScheduledScanType)
                                            {
                                                1 { $ScheduledScanType = "Quick Scan" }
                                                2 { $ScheduledScanType = "Full Scan" }
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
                                WriteWordLine 0 0 ""
                                WriteWordLine 0 2 "Scan settings" -bold
                                WriteWordLine 0 3 "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                                WriteWordLine 0 3 "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                                WriteWordLine 0 3 "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                                WriteWordLine 0 3 "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                                WriteWordLine 0 3 "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                                WriteWordLine 0 3 "User control of scheduled scans: " -nonewline
                                switch ($AgentConfig.ScheduledScanUserControl)
                                    {
                                        0 { WriteWordLine 0 0 "No control" }
                                        1 { WriteWordLine 0 0 "Scan time only" }
                                        2 { WriteWordLine 0 0 "Full control" }
                                    }
                                WriteWordLine 0 2 "Default Actions" -bold
                                WriteWordLine 0 3 "Severe threats: " -nonewline
                                switch ($AgentConfig.DefaultActionSevere)
                                    {
                                        0 { WriteWordLine 0 0 "Recommended" }
                                        2 { WriteWordLine 0 0 "Quarantine" }
                                        3 { WriteWordLine 0 0 "Remove" }
                                        6 { WriteWordLine 0 0 "Allow" }
                                    }
                                WriteWordLine 0 3 "High threats: " -nonewline
                                switch ($AgentConfig.DefaultActionSevere)
                                    {
                                        0 { WriteWordLine 0 0 "Recommended" }
                                        2 { WriteWordLine 0 0 "Quarantine" }
                                        3 { WriteWordLine 0 0 "Remove" }
                                        6 { WriteWordLine 0 0 "Allow" }
                                    }
                                WriteWordLine 0 3 "Medium threats: " -nonewline
                                switch ($AgentConfig.DefaultActionSevere)
                                    {
                                        0 { WriteWordLine 0 0 "Recommended" }
                                        2 { WriteWordLine 0 0 "Quarantine" }
                                        3 { WriteWordLine 0 0 "Remove" }
                                        6 { WriteWordLine 0 0 "Allow" }
                                    }
                                WriteWordLine 0 3 "Low threats: " -nonewline
                                switch ($AgentConfig.DefaultActionSevere)
                                    {
                                        0 { WriteWordLine 0 0 "Recommended" }
                                        2 { WriteWordLine 0 0 "Quarantine" }
                                        3 { WriteWordLine 0 0 "Remove" }
                                        6 { WriteWordLine 0 0 "Allow" }
                                    }
                                WriteWordLine 0 2 "Real-time protection" -bold
                                WriteWordLine 0 3 "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                WriteWordLine 0 3 "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                                WriteWordLine 0 3 "Scan system files: " -nonewline
                                switch ($AgentConfig.RealtimeScanOption)
                                    {
                                        0 { WriteWordLine 0 0 "Scan incoming and outgoing files" }
                                        1 { WriteWordLine 0 0 "Scan incoming files only" }
                                        2 { WriteWordLine 0 0 "Scan outgoing files only" }
                                    }
                                WriteWordLine 0 2 "Exclusion settings" -bold
                                WriteWordLine 0 3 "Excluded files and folders: "
                                foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths)
                                    {
                                        WriteWordLine 0 4 "$($ExcludedFileFolder)"
                                    }
                                WriteWordLine 0 3 "Excluded file types: "
                                foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes)
                                    {
                                        WriteWordLine 0 4 "$($ExcludedFileType)"
                                    }
                                WriteWordLine 0 3 "Excluded processes: "
                                foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses)
                                    {
                                        WriteWordLine 0 4 "$($ExcludedProcess)"
                                    }
                                WriteWordLine 0 2 "Advanced" -bold
                                WriteWordLine 0 3 "Create a system restore point before computers are cleaned: $($AgentConfig.CreateSystemRestorePointBeforeClean)"
                                WriteWordLine 0 3 "Disable the client user interface: $($AgentConfig.DisableClientUI)"
                                WriteWordLine 0 3 "Show notifications messages on the client computer when the user needs to run a full scan, update definitions, or run Windows Defender Offline: $($AgentConfig.ShowNotificationMessages)"
                                WriteWordLine 0 3 "Delete quarantined files after (days): $($AgentConfig.DeleteQuarantinedFilesPeriod)"
                                WriteWordLine 0 3 "Allow users to configure the setting for quarantined file deletion: $($AgentConfig.AllowUserConfigQuarantinedFileDeletionPeriod)"
                                WriteWordLine 0 3 "Allow users to exclude file and folders, file types and processes: $($AgentConfig.AllowUserAddExcludes)"
                                WriteWordLine 0 3 "Allow all users to view the full History results: $($AgentConfig.AllowUserViewHistory)"
                                WriteWordLine 0 3 "Enable reparse point scanning: $($AgentConfig.EnableReparsePointScanning)"
                                WriteWordLine 0 3 "Randomize scheduled scan and definition update start time (within 30 minutes): $($AgentConfig.RandomizeScheduledScanStartTime)"
        
                                WriteWordLine 0 2 "Threat overrides" -bold
                                if (-not [string]::IsNullOrEmpty($AgentConfig.ThreatName))
                                    {
                                        WriteWordLine 0 3 "Threat name and override action: Threats specified."
                                    }
                                WriteWordLine 0 2 "Microsoft Active Protection Service" -bold
                                WriteWordLine 0 3 "Microsoft Active Protection Service membership type: " -nonewline
                                switch ($AgentConfig.JoinSpyNet)
                                    {
                                        0 { WriteWordLine 0 0 "Do not join MAPS" }
                                        1 { WriteWordLine 0 0 "Basic membership" }
                                        2 { WriteWordLine 0 0 "Advanced membership" }
                                    }
                                WriteWordLine 0 3 "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"

                                WriteWordLine 0 2 "Definition Updates" -bold
                                WriteWordLine 0 3 "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                                WriteWordLine 0 3 "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                                WriteWordLine 0 3 "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                                WriteWordLine 0 3 "Set sources and order for Endpoint Protection definition updates: "
                                foreach ($Fallback in $AgentConfig.FallbackOrder)
                                    {
                                        WriteWordLine 0 3 "$($Fallback)"
                                    }
                                WriteWordLine 0 3 "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                                WriteWordLine 0 3 "If UNC file shares are selected as a definition update source, specify the UNC paths:" 
                                foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources)
                                    {
                                        WriteWordLine 0 4 "$($UNCShare)"
                                    }
                            }
                    else
                        {
                            $AgentConfig_custom = $AntiMalwarePolicy.AgentConfigurations
                            WriteWordLine 0 1 "$($AntiMalwarePolicy.Name)" -bold
                            WriteWordLine 0 1 "Description: $($AntiMalwarePolicy.Description)"
                            foreach ($Agentconfig in $AgentConfig_custom)
                                {
                                    switch ($AgentConfig.AgentID)
                                        {
                                            201 
                                                {
                                                    WriteWordLine 0 2 "Scheduled Scans" -bold
                                                    WriteWordLine 0 2 "Run a scheduled scan on client computers: $($AgentConfig.EnableScheduledScan)"
                                                    if ($AgentConfig.EnableScheduledScan)
                                                        {
                                                            switch ($AgentConfig.ScheduledScanType)
                                                                {
                                                                    1 { $ScheduledScanType = "Quick Scan" }
                                                                    2 { $ScheduledScanType = "Full Scan" }
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
                                                    WriteWordLine 0 2 "Default Actions" -bold
                                                    WriteWordLine 0 3 "Severe threats: " -nonewline
                                                    switch ($AgentConfig.DefaultActionSevere)
                                                        {
                                                            0 { WriteWordLine 0 0 "Recommended" }
                                                            2 { WriteWordLine 0 0 "Quarantine" }
                                                            3 { WriteWordLine 0 0 "Remove" }
                                                            6 { WriteWordLine 0 0 "Allow" }
                                                        }
                                                    WriteWordLine 0 3 "High threats: " -nonewline
                                                    switch ($AgentConfig.DefaultActionSevere)
                                                        {
                                                            0 { WriteWordLine 0 0 "Recommended" }
                                                            2 { WriteWordLine 0 0 "Quarantine" }
                                                            3 { WriteWordLine 0 0 "Remove" }
                                                            6 { WriteWordLine 0 0 "Allow" }
                                                        }
                                                    WriteWordLine 0 3 "Medium threats: " -nonewline
                                                    switch ($AgentConfig.DefaultActionSevere)
                                                        {
                                                            0 { WriteWordLine 0 0 "Recommended" }
                                                            2 { WriteWordLine 0 0 "Quarantine" }
                                                            3 { WriteWordLine 0 0 "Remove" }
                                                            6 { WriteWordLine 0 0 "Allow" }
                                                        }
                                                    WriteWordLine 0 3 "Low threats: " -nonewline
                                                    switch ($AgentConfig.DefaultActionSevere)
                                                        {
                                                            0 { WriteWordLine 0 0 "Recommended" }
                                                            2 { WriteWordLine 0 0 "Quarantine" }
                                                            3 { WriteWordLine 0 0 "Remove" }
                                                            6 { WriteWordLine 0 0 "Allow" }
                                                        }                                           
                                                }
                                            203
                                                {
                                                    WriteWordLine 0 2 "Exclusion settings" -bold
                                                    WriteWordLine 0 3 "Excluded files and folders: "
                                                    foreach ($ExcludedFileFolder in $AgentConfig.ExcludedFilePaths)
                                                        {
                                                            WriteWordLine 0 4 "$($ExcludedFileFolder)"
                                                        }
                                                    WriteWordLine 0 3 "Excluded file types: "
                                                    foreach ($ExcludedFileType in $AgentConfig.ExcludedFileTypes)
                                                        {
                                                            WriteWordLine 0 4 "$($ExcludedFileType)"
                                                        }
                                                    WriteWordLine 0 3 "Excluded processes: "
                                                    foreach ($ExcludedProcess in $AgentConfig.ExcludedProcesses)
                                                        {
                                                            WriteWordLine 0 4 "$($ExcludedProcess)"
                                                        }                                            
                                                }
                                            204
                                                {
                                                    WriteWordLine 0 2 "Real-time protection" -bold
                                                    WriteWordLine 0 3 "Enable real-time protection: $($AgentConfig.RealtimeProtectionOn)"
                                                    WriteWordLine 0 3 "Monitor file and program activity on your computer: $($AgentConfig.MonitorFileProgramActivity)"
                                                    WriteWordLine 0 3 "Scan system files: " -nonewline
                                                    switch ($AgentConfig.RealtimeScanOption)
                                                        {
                                                            0 { WriteWordLine 0 0 "Scan incoming and outgoing files" }
                                                            1 { WriteWordLine 0 0 "Scan incoming files only" }
                                                            2 { WriteWordLine 0 0 "Scan outgoing files only" }
                                                        }                                            
                                                }
                                            205
                                                {
                                                    WriteWordLine 0 2 "Advanced" -bold
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
                                                    WriteWordLine 0 2 "Microsoft Active Protection Service" -bold
                                                    WriteWordLine 0 3 "Microsoft Active Protection Service membership type: " -nonewline
                                                    switch ($AgentConfig.JoinSpyNet)
                                                        {
                                                            0 { WriteWordLine 0 0 "Do not join MAPS" }
                                                            1 { WriteWordLine 0 0 "Basic membership" }
                                                            2 { WriteWordLine 0 0 "Advanced membership" }
                                                        }
                                                    WriteWordLine 0 3 "Allow users to modify Microsoft Active Protection Service settings: $($AgentConfig.AllowUserChangeSpyNetSettings)"                                            
                                                }
                                            208
                                                {
                                                    WriteWordLine 0 2 "Definition Updates" -bold
                                                    WriteWordLine 0 3 "Check for Endpoint Protection definitions at a specific interval (hours): (0 disable check on interval) $($AgentConfig.SignatureUpdateInterval)"
                                                    WriteWordLine 0 3 "Check for Endpoint Protection definitions daily at: (Only configurable if interval-based check is disabled) $(Convert-Time -time $AgentConfig.SignatureUpdateTime)"
                                                    WriteWordLine 0 3 "Force a definition update if the client computer is offline for more than two consecutive scheduled updates: $($AgentConfig.EnableSignatureUpdateCatchupInterval)"
                                                    WriteWordLine 0 3 "Set sources and order for Endpoint Protection definition updates: "
                                                    foreach ($Fallback in $AgentConfig.FallbackOrder)
                                                        {
                                                            WriteWordLine 0 4 "$($Fallback)"
                                                        }
                                                    WriteWordLine 0 3 "If Configuration Manager is used as a source for definition updates, clients will only update from alternative sources if definition is older than (hours): $($AgentConfig.AuGracePeriod / 60)"
                                                    WriteWordLine 0 3 "If UNC file shares are selected as a definition update source, specify the UNC paths:" 
                                                    foreach ($UNCShare in $AgentConfig.DefinitionUpdateFileSharesSources)
                                                        {
                                                            WriteWordLine 0 4 "$($UNCShare)"
                                                        }
                                                }
                                            209
                                                {
                                                    WriteWordLine 0 2 "Scan settings" -bold
                                                    WriteWordLine 0 3 "Scan email and email attachments: $($AgentConfig.ScanEmail)"
                                                    WriteWordLine 0 3 "Scan removable storage devices such as USB drives: $($AgentConfig.ScanRemovableStorage)"
                                                    WriteWordLine 0 3 "Scan network drives when running a full scan: $($AgentConfig.ScanNetworkDrives)"
                                                    WriteWordLine 0 3 "Scan archived files: $($AgentConfig.ScanArchivedFiles)"
                                                    WriteWordLine 0 3 "Allow users to configure CPU usage during scans: $($AgentConfig.AllowClientUserConfigLimitCPUUsage)"
                                                    WriteWordLine 0 3 "User control of scheduled scans: " -nonewline
                                                    switch ($AgentConfig.ScheduledScanUserControl)
                                                        {
                                                            0 { WriteWordLine 0 0 "No control" }
                                                            1 { WriteWordLine 0 0 "Scan time only" }
                                                            2 { WriteWordLine 0 0 "Full control" }
                                                        }
                                                }
                                        }
                                }
                        }
                }
            }
        else
            {
                WriteWordLine 0 1 "There are no Anti Malware Policies configured."
            }
    }
else
    {
        WriteWordLine 0 1 "There is no Endpoint Protection Point enabled."
    }

WriteWordLine 0 0 ""

Write-Verbose "$(Get-Date):   Working on Windows Firewall Policies"
WriteWordLine 3 0 "Windows Firewall Policies"

$FirewallPolicies = Get-CMWindowsFirewallPolicy
if (-not [string]::IsNullOrEmpty($FirewallPolicies))
    {
        WriteWordLine 0 1 "Enumerating Windows Firewall Policies:"
        $Table = $Null
        $TableRange = $Null
        $TableRange = $doc.Application.Selection.Range
        $Columns = 4
        [int]$Rows = $FirewallPolicies.count + 1
        Write-Verbose "$(Get-Date):   add configuration items to table"
        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
        $table.Style = $TableStyle 
        $table.Borders.InsideLineStyle = 1
        $table.Borders.OutsideLineStyle = 1
        [int]$xRow = 1
        Write-Verbose "$(Get-Date):   format first row with column headings"
        
        $Table.Cell($xRow,1).Range.Font.Bold = $True
        $Table.Cell($xRow,1).Range.Font.Size = "10"
        $Table.Cell($xRow,1).Range.Text = "Name"
        
        $Table.Cell($xRow,2).Range.Font.Bold = $True
        $Table.Cell($xRow,2).Range.Font.Size = "10"
        $Table.Cell($xRow,2).Range.Text = "Last modified"
        
        $Table.Cell($xRow,3).Range.Font.Bold = $True
        $Table.Cell($xRow,3).Range.Font.Size = "10"
        $Table.Cell($xRow,3).Range.Text = "Last modified by"
        
        $Table.Cell($xRow,4).Range.Font.Bold = $True
        $Table.Cell($xRow,4).Range.Font.Size = "10"
        $Table.Cell($xRow,4).Range.Text = "CI ID"
        foreach ($FirewallPolicy in $FirewallPolicies)
            {
                $xRow++							
		        $Table.Cell($xRow,1).Range.Font.Size = "10"
		        $Table.Cell($xRow,1).Range.Text = $FirewallPolicy.LocalizedDisplayName
		        $Table.Cell($xRow,2).Range.Font.Size = "10"
		        $Table.Cell($xRow,2).Range.Text = $FirewallPolicy.DateLastModified
                $Table.Cell($xRow,3).Range.Font.Size = "10"
		        $Table.Cell($xRow,3).Range.Text = $FirewallPolicy.LastModifiedBy
                $Table.Cell($xRow,4).Range.Font.Size = "10"
		        $Table.Cell($xRow,4).Range.Text = $FirewallPolicy.CI_ID   
            }				
        $Table.Rows.SetLeftIndent(50,1) | Out-Null
        $table.AutoFitBehavior(1) | Out-Null
        #return focus back to document
        Write-Verbose "$(Get-Date):   return focus back to document"
        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    
        #move to the end of the current document
        Write-Verbose "$(Get-Date):   move to the end of the current document"
        $selection.EndKey($wdStory,$wdMove) | Out-Null
        WriteWordLine 0 0 ""
    }
else
    {
        WriteWordLine 0 1 "There are no Windows Firewall policies configured."
    }
    

#####
##### finished with Assets and Compliance, moving on to Software Library
#####
        Write-Verbose "$(Get-Date):   Finished with Assets and Compliance, moving on to Software Library"
        WriteWordLine 1 0 "Software Library"

##### Applications
        
        WriteWordLine 2 0 "Applications"
        WriteWordLine 0 0 ""
        $Applications = Get-WmiObject -Class sms_application -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider
        if ($ListAllInformation)
            {
                if (-not [string]::IsNullOrEmpty($Applications))
                    {
                        WriteWordLine 0 1 "The following Applications are configured in this site:"
                        foreach ($Application in $Applications)
                            {                                
                                Write-Verbose "Getting specific WMI instance for this App"
                                [wmi]$Application = $Application.__PATH
                                Write-Verbose "$(Get-Date):   Found App: $($Application.LocalizedDisplayName)"
                                WriteWordLine 0 2 "$($Application.LocalizedDisplayName)" -bold
                                WriteWordLine 0 3 "Created by: $($Application.CreatedBy)"
                                WriteWordLine 0 3 "Date created: $($Application.DateCreated)"
                                WriteWordLine 0 3 "PackageID: $($Application.PackageID)"
                                $DTs = Get-CMDeploymentType -ApplicationName $Application.LocalizedDisplayName
                                if (-not [string]::IsNullOrEmpty($DTs))
                                    {                                       
                                        WriteWordLine 0 0 ""
                                        $Table = $Null
                                        $TableRange = $Null
                                        $TableRange = $doc.Application.Selection.Range
				                        $Columns = 3
                                        [int]$Rows = $DTs.count + 1
				                        Write-Verbose "$(Get-Date):   add Deployment Types to table"
				                        $Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				                        $table.Style = $TableStyle 
				                        $table.Borders.InsideLineStyle = 1
				                        $table.Borders.OutsideLineStyle = 1
				                        [int]$xRow = 1
				                        Write-Verbose "$(Get-Date):   format first row with column headings"
				                        
				                        $Table.Cell($xRow,1).Range.Font.Bold = $True
				                        $Table.Cell($xRow,1).Range.Font.Size = "10"
				                        $Table.Cell($xRow,1).Range.Text = "Deployment Type name"
				                
				                        $Table.Cell($xRow,2).Range.Font.Bold = $True
				                        $Table.Cell($xRow,2).Range.Font.Size = "10"
				                        $Table.Cell($xRow,2).Range.Text = "Technology"

				                        $Table.Cell($xRow,3).Range.Font.Bold = $True
				                        $Table.Cell($xRow,3).Range.Font.Size = "10"
				                        $Table.Cell($xRow,3).Range.Text = "Commandline"
                                        foreach ($DT in $DTs)
                                            {
                                                #[wmi]$DT = $DT.__PATH
                                                $xml = [xml]$DT.SDMPackageXML
                                                $xRow++							
					                            $Table.Cell($xRow,1).Range.Font.Size = "10"
					                            $Table.Cell($xRow,1).Range.Text = $DT.LocalizedDisplayName
					                            $Table.Cell($xRow,2).Range.Font.Size = "10"
					                            $Table.Cell($xRow,2).Range.Text = $DT.Technology
                                                if (-not ($DT.Technology -like "AppV*"))
                                                    { 
					                                    $Table.Cell($xRow,3).Range.Font.Size = "10"
					                                    $Table.Cell($xRow,3).Range.Text = $xml.AppMgmtDigest.DeploymentType.Installer.CustomData.InstallCommandLine
                                                    }
                                            }				
				                        $Table.Rows.SetLeftIndent(50,1) | Out-Null
				                        $table.AutoFitBehavior(1) | Out-Null
				                        #return focus back to document
				                        Write-Verbose "$(Get-Date):   return focus back to document"
				                        $doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument
                                        #move to the end of the current document
			                            Write-Verbose "$(Get-Date):   move to the end of the current document"
			                            $selection.EndKey($wdStory,$wdMove) | Out-Null
			                            WriteWordLine 0 0 ""
                           
                                    }
                                else
                                    {
                                        WriteWordLine 0 3 "There are no Deployment Types configured for this Application."
                                    }
                            }
                    }
                else
                    {
                        WriteWordLine 0 1 "There are no Applications configured in this site."
                    }
            }
        elseif ($Applications)
            {
                WriteWordLine 0 1 "There are $($Applications.count) applications configured."
            }
            else
                {
                    WriteWordLine 0 1 "There are no Applications configured."
                }

##### Packages
        
        WriteWordLine 2 0 "Packages"
        WriteWordLine 0 0 ""
        $Packages = Get-CMPackage
        if ($ListAllInformation)
            {
                if (-not [string]::IsNullOrEmpty($Packages))
                    {
                        WriteWordLine 0 1 "The following Packages are configured in this site:"
                        foreach ($Package in $Packages)
                        {
                        WriteWordLine 0 0 ""
                        WriteWordLine 0 2 "$($Package.Name)" -bold
                        WriteWordLine 0 3 "Description: $($Package.Description)"
                        WriteWordLine 0 3 "PackageID: $($Package.PackageID)"
                        $Programs = Get-WmiObject -Class SMS_Program -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "PackageID = '$($Package.PackageID)'"
                        if (-not [string]::IsNullOrEmpty($Programs))
                            {
                                WriteWordLine 0 3 "The Package has the following Programs configured:"
                                foreach ($Program in $Programs)
                                    {
                                        WriteWordLine 0 4 "Program Name: $($Program.ProgramName)" -bold
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
                                                WriteWordLine 0 4 "This is a default program."
                                            }
                                        if ($Program.ProgramFlags -band 0x00000020)
                                            {
                                                WriteWordLine 0 4 "Disables MOM alerts while the program runs."
                                            }
                                        if ($Program.ProgramFlags -band 0x00000040)
                                            {
                                                WriteWordLine 0 4 "Generates MOM alert if the program fails."
                                            }
                                        if ($Program.ProgramFlags -band 0x00000080)
                                            {
                                                WriteWordLine 0 4 "This program's immediate dependent should always be run."
                                            }
                                        if ($Program.ProgramFlags -band 0x00000100)
                                            {
                                                WriteWordLine 0 4 "A device program. The program is not offered to desktop clients."
                                            }
                                        if ($Program.ProgramFlags -band 0x00000400)
                                            {
                                                WriteWordLine 0 4 "The countdown dialog is not displayed."
                                            }
                                        if ($Program.ProgramFlags -band 0x00001000)
                                            {
                                                WriteWordLine 0 4 "The program is disabled."
                                            }
                                        if ($Program.ProgramFlags -band 0x00002000)
                                            {
                                                WriteWordLine 0 4 "The program requires no user interaction."
                                            }
                                        if ($Program.ProgramFlags -band 0x00004000)
                                            {
                                                WriteWordLine 0 4 "The program can run only when a user is logged on."
                                            }
                                        if ($Program.ProgramFlags -band 0x00008000)
                                            {
                                                WriteWordLine 0 4 "The program must be run as the local Administrator account."
                                            }
                                        if ($Program.ProgramFlags -band 0x00010000)
                                            {
                                                WriteWordLine 0 4 "The program must be run by every user for whom it is valid. Valid only for mandatory jobs."
                                            }
                                        if ($Program.ProgramFlags -band 0x00020000)
                                            {
                                                WriteWordLine 0 4 "The program is run only when no user is logged on."
                                            }
                                        if ($Program.ProgramFlags -band 0x00040000)
                                            {
                                                WriteWordLine 0 4 "The program will restart the computer."
                                            }
                                        if ($Program.ProgramFlags -band 0x00080000)
                                            {
                                                WriteWordLine 0 4 "Configuration Manager restarts the computer when the program has finished running successfully."
                                            }
                                        if ($Program.ProgramFlags -band 0x00100000)
                                            {
                                                WriteWordLine 0 4 "Use a UNC path (no drive letter) to access the distribution point."
                                            }
                                        if ($Program.ProgramFlags -band 0x00200000)
                                            {
                                                WriteWordLine 0 4 "Persists the connection to the drive specified in the DriveLetter property. The USEUNCPATH bit flag must not be set."
                                            }
                                        if ($Program.ProgramFlags -band 0x00400000)
                                            {
                                                WriteWordLine 0 4 "Run the program as a minimized window."
                                            }
                                        if ($Program.ProgramFlags -band 0x00800000)
                                            {
                                                WriteWordLine 0 4 "Run the program as a maximized window."
                                            }
                                        if ($Program.ProgramFlags -band 0x01000000)
                                            {
                                                WriteWordLine 0 4 "Hide the program window."
                                            }
                                        if ($Program.ProgramFlags -band 0x02000000)
                                            {
                                                WriteWordLine 0 4 "Logoff user when program completes successfully."
                                            }
                                        if ($Program.ProgramFlags -band 0x08000000)
                                            {
                                                WriteWordLine 0 4 "Override check for platform support."
                                            }
                                        if ($Program.ProgramFlags -band 0x20000000)
                                            {
                                                WriteWordLine 0 4 "Run uninstall from the registry key when the advertisement expires."   
                                            }   
                                    }                             
                            }
                        else
                            {
                                WriteWordLine 0 4 "The Package has no Programs configured."
                            }                       
                    }
                    }
                else
                    {
                        WriteWordLine 0 1 "There are no Packages configured in this site."
                    }
            }
        elseif ($Packages)
            {
                WriteWordLine 0 1 "There are $($Packages.count) packages configured."
            }
            else
                {
                    WriteWordLine 0 1 "There are no packages configured."
                }

##### Driver Packages

    WriteWordLine 2 0 "Driver Packages"
    WriteWordLine 0 0 ""
    $DriverPackages = Get-CMDriverPackage
    if ($ListAllInformation)
        {
            if (-not [string]::IsNullOrEmpty($DriverPackages))
                    {
                        WriteWordLine 0 1 "The following Driver Packages are configured in your site:"
                        foreach ($DriverPackage in $DriverPackages)
                            {
                                WriteWordLine 0 0 ""
                                WriteWordLine 0 2 "Name: $($DriverPackage.Name)" -bold
                                if ($DriverPackage.Description)
                                    {
                                        WriteWordLine 0 2 "Description: $($DriverPackage.Description)"
                                    }
                                WriteWordLine 0 2 "PackageID: $($DriverPackage.PackageID)"
                                WriteWordLine 0 2 "Source path: $($DriverPackage.PkgSourcePath)"
                                WriteWordLine 0 2 "This package consists of the following Drivers:"
                                $Drivers = Get-CMDriver -DriverPackageId "$($DriverPackage.PackageID)"
                                foreach ($Driver in $Drivers)
                                    {
                                        WriteWordLine 0 0 ""
                                        WriteWordLine 0 3 "Driver Name: $($Driver.LocalizedDisplayName)" -bold
                                        WriteWordLine 0 3 "Manufacturer: $($Driver.DriverProvider)"
                                        WriteWordLine 0 3 "Source path: $($Driver.ContentSourcePath)"
                                        WriteWordLine 0 3 "INF File: $($Driver.DriverINFFile)"
                                    }
                                WriteWordLine 0 3 ""
                            }
                    }
                else
                    {
                        WriteWordLine 0 1 "There are no Driver Packages configured in this site."
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
                    WriteWordLine 0 1 "There are no Driver Packages configured in this site."
                }
        }
 
 ##### Operating System Images

    WriteWordLine 2 0 "Operating System Images"
    WriteWordLine 0 0 ""
    $OSImages = Get-CMOperatingSystemImage
    if (-not [string]::IsNullOrEmpty($OSImages))
        {
            WriteWordLine 0 1 "The following OS Images are imported into your site:"
            foreach ($OSImage in $OSImages)
                {
                    WriteWordLine 0 0 ""
                    WriteWordLine 0 2 "Name: $($OSImage.Name)" -bold
                    if ($OSImage.Description)
                            {
                                WriteWordLine 0 2 "Description: $($OSImage.Description)"
                            }
                    WriteWordLine 0 2 "Package ID: $($OSImage.PackageID)"
                    WriteWordLine 0 2 "Source Path: $($OSImage.PkgSourcePath)"
                }
        }
    else
        {
            WriteWordLine 0 1 "There are no OS Images imported into this environment."
        }


##### Operating System Installers

    WriteWordLine 2 0 "Operating System Installers"
    WriteWordLine 0 0 ""
    $OSInstallers = Get-CMOperatingSystemInstaller
    if (-not [string]::IsNullOrEmpty($OSImages))
        {
            WriteWordLine 0 1 "The following OS Installers are imported into this environment:"
            foreach ($OSInstaller in $OSInstallers)
                {
                    WriteWordLine 0 2 "Name: $($OSInstaller.Name)" -bold
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
            WriteWordLine 0 1 "There are no OS Installers imported into this environment."
        }
        
####
####
#### Boot Images
    
WriteWordLine 2 0 "Boot Images"
WriteWordLine 0 0 ""
$BootImages = Get-CMBootImage
if (-not [string]::IsNullOrEmpty($BootImages))
    {
        WriteWordLine 0 1 "The following Boot Images are imported into this environment:"
        WriteWordLine 0 0 ""
        foreach ($BootImage in $BootImages)
            {
                WriteWordLine 0 2 "$($BootImage.Name)" -bold
                if ($BootImage.Description)
                    {
                        WriteWordLine 0 2 "Description: $($BootImage.Description)"
                    }
                WriteWordLine 0 2 "Source Path: $($BootImage.PkgSourcePath)"
                WriteWordLine 0 2 "Package ID: $($BootImage.PackageID)"
                WriteWordLine 0 2 "Architecture: " -nonewline
                switch ($BootImage.Architecture)
                    {
                        0 { WriteWordLine 0 0 "x86" }
                        9 { WriteWordLine 0 0 "x64" }
                    }
                if ($BootImage.BackgroundBitmapPath)
                    {
                        WriteWordLine 0 2 "Custom Background: $($BootImage.BackgroundBitmapPath)"
                    }
                Switch ($BootImage.EnableLabShell)
                    {
                        True { WriteWordLine 0 2 "Command line support is enabled" }
                        False { WriteWordLine 0 2 "Command line support is not enabled" }
                    }
                WriteWordLine 0 2 "The following drivers are imported into this WinPE"
                if (-not [string]::IsNullOrEmpty($BootImage.ReferencedDrivers))
                    {
                        $ImportedDriverIDs = ($BootImage.ReferencedDrivers).ID | Out-Null
                        foreach ($ImportedDriverID in $ImportedDriverIDs)
                            {
                                $ImportedDriver = Get-CMDriver -ID $ImportedDriverID
                                WriteWordLine 0 3 "Name: $($ImportedDriver.LocalizedDisplayName)" -bold
                                WriteWordLine 0 3 "Inf File: $($ImportedDriver.DriverINFFile)"
                                WriteWordLine 0 3 "Driver Class: $($ImportedDriver.DriverClass)"
                                WriteWordLine 0 0 ""
                            }
                    }
                else
                    {
                        WriteWordLine 0 3 "There are no drivers imported into the Boot Image."
                    }
            if (-not [string]::IsNullOrEmpty($BootImage.OptionalComponents))
                {
                    $Component = $Null
                    WriteWordLine 0 3 "The following Optional Components are added to this Boot Image:"
                    foreach ($Component in $BootImage.OptionalComponents)
                        {
                            switch ($Component)
                                {
                                    {($_ -eq "1") -or ($_ -eq "27")} { WriteWordLine 0 4 "WinPE-DismCmdlets" }                                    {($_ -eq "2") -or ($_ -eq "28")} { WriteWordLine 0 4 "WinPE-Dot3Svc" }                                    {($_ -eq "3") -or ($_ -eq "29")} { WriteWordLine 0 4 "WinPE-EnhancedStorage" }                                    {($_ -eq "4") -or ($_ -eq "30")} { WriteWordLine 0 4 "WinPE-FMAPI" }                                    {($_ -eq "5") -or ($_ -eq "31")} { WriteWordLine 0 4 "WinPE-FontSupport-JA-JP" }                                    {($_ -eq "6") -or ($_ -eq "32")} { WriteWordLine 0 4 "WinPE-FontSupport-KO-KR" }                                    {($_ -eq "7") -or ($_ -eq "33")} { WriteWordLine 0 4 "WinPE-FontSupport-ZH-CN" }                                    {($_ -eq "8") -or ($_ -eq "34")} { WriteWordLine 0 4 "WinPE-FontSupport-ZH-HK" }                                    {($_ -eq "9") -or ($_ -eq "35")} { WriteWordLine 0 4 "WinPE-FontSupport-ZH-TW" }                                    {($_ -eq "10") -or ($_ -eq "36")} { WriteWordLine 0 4 "WinPE-HTA" }                                    {($_ -eq "11") -or ($_ -eq "37")} { WriteWordLine 0 4 "WinPE-StorageWMI" }                                    {($_ -eq "12") -or ($_ -eq "38")} { WriteWordLine 0 4 "WinPE-LegacySetup" }                                    {($_ -eq "13") -or ($_ -eq "39")} { WriteWordLine 0 4 "WinPE-MDAC" }                                    {($_ -eq "14") -or ($_ -eq "40")} { WriteWordLine 0 4 "WinPE-NetFx4" }                                    {($_ -eq "15") -or ($_ -eq "41")} { WriteWordLine 0 4 "WinPE-PowerShell3" }                                    {($_ -eq "16") -or ($_ -eq "42")} { WriteWordLine 0 4 "WinPE-PPPoE" }                                    {($_ -eq "17") -or ($_ -eq "43")} { WriteWordLine 0 4 "WinPE-RNDIS" }                                    {($_ -eq "18") -or ($_ -eq "44")} { WriteWordLine 0 4 "WinPE-Scripting" }                                    {($_ -eq "19") -or ($_ -eq "45")} { WriteWordLine 0 4 "WinPE-SecureStartup" }                                    {($_ -eq "20") -or ($_ -eq "46")} { WriteWordLine 0 4 "WinPE-Setup" }                                    {($_ -eq "21") -or ($_ -eq "47")} { WriteWordLine 0 4 "WinPE-Setup-Client" }                                    {($_ -eq "22") -or ($_ -eq "48")} { WriteWordLine 0 4 "WinPE-Setup-Server" }                                    #{($_ -eq "23") -or ($_ -eq "49")} { WriteWordLine 0 4 "Not applicable" }                                    {($_ -eq "24") -or ($_ -eq "50")} { WriteWordLine 0 4 "WinPE-WDS-Tools" }                                    {($_ -eq "25") -or ($_ -eq "51")} { WriteWordLine 0 4 "WinPE-WinReCfg" }                                    {($_ -eq "26") -or ($_ -eq "52")} { WriteWordLine 0 4 "WinPE-WMI" }
                                } 
                            $Component = $Null    
                        }
                    }
                WriteWordLine 0 0 ""

            }
    }
else
    {
        WriteWordLine 0 1 "There are no Boot Images present in this environment."
    }

####
####
#### Task Sequences
Write-Verbose "$(Get-Date):   Enumerating Task Sequences"
WriteWordLine 2 0 "Task Sequences"
WriteWordLine 0 0 ""

$TaskSequences = Get-CMTaskSequence
Write-Verbose "$(Get-Date):   working on $($TaskSequences.count) Task Sequences"
if ($ListAllInformation)
    {
        if (-not [string]::IsNullOrEmpty($TaskSequences))
            {
                foreach ($TaskSequence in $TaskSequences)
                    {
                        WriteWordLine 0 1 "Task Sequence name: $($TaskSequence.Name)" -bold
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
                                        WriteWordLine 0 1 "Group name: $($Group.Name)" -bold
                                        if (-not [string]::IsNullOrEmpty($Group.Description))
                                            {
                                                WriteWordLine 0 1 "Description: $($Group.Description)"
                                            }
                                        WriteWordLine 0 1 "This Group has the following steps configured."
                                        foreach ($Step in $Group.Step)
                                            {
                                                WriteWordLine 0 3 "$($Step.Name)" -bold
                                                if (-not [string]::IsNullOrEmpty($Step.Description))
                                                    {
                                                        WriteWordLine 0 4 "$($Step.Description)"
                                                    }
                                                WriteWordLine 0 4 "$($Step.Action)"
                                                try 
                                                    {
                                                        if (-not [string]::IsNullOrEmpty($Step.disable))
                                                                {
                                                                    WriteWordLine 0 4 "This step is disabled."
                                                                }
                                                    }   
                                                catch [System.Management.Automation.PropertyNotFoundException] 
                                                    {
                                                        WriteWordLine 0 4 "This step is enabled"
                                                    }
                                                WriteWordLine 0 0 ""
                                            }

                                    }
                            }
                        catch [System.Management.Automation.PropertyNotFoundException]
                            {
                                WriteWordLine 0 0 ""
                            }
                        try 
                            {
                                foreach ($Step in $Sequence.sequence.step)
                                    {
                                        WriteWordLine 0 3 "$($Step.Name)" -bold
                                        if (-not [string]::IsNullOrEmpty($Step.Description))
                                            {
                                                WriteWordLine 0 4 "$($Step.Description)"
                                            }
                                        WriteWordLine 0 4 "$($Step.Action)"
                                        try 
                                            {
                                                if (-not [string]::IsNullOrEmpty($Step.disable))
                                                        {
                                                            WriteWordLine 0 4 "This step is disabled."
                                                        }
                                            }   
                                        catch [System.Management.Automation.PropertyNotFoundException] 
                                            {
                                                WriteWordLine 0 4 "This step is enabled"
                                            }
                                        WriteWordLine 0 0 ""
                                    }
                            }
                        catch [System.Management.Automation.PropertyNotFoundException]
                            {
                                WriteWordLine 0 0 ""
                            }
                        
                        WriteWordLine 0 0 ""
                        WriteWordLine 0 0 "----------------------------------------------"
                    }
            }
        else
            {
                WriteWordLine 0 1 "There are no Task Sequences present in this environment."
            }
    }
else
    {
        if (-not [string]::IsNullOrEmpty($TaskSequences))
            {
                WriteWordLine 0 1 "The following Task Sequences are configured:"
                foreach ($TaskSequence in $TaskSequences)
                    {
                        WriteWordLine 0 2 "$($TaskSequence.Name)"
                    }
            }
        else
            {
                WriteWordLine 0 1 "There are no Task Sequences present in this environment."
            }
    }

######################## END OF MAIN SCRIPT ######################
Set-Location $LocationBeforeExecution
Write-Verbose "$(Get-Date):   Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "XenApp 6.5 Farm Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp = $doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract = "Citrix XenApp 6.5 Inventory for $CompanyName"
	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date): Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}

Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
If($WordVersion -eq $wdWord2007)
{
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2"))
	{
		$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
		$doc.SaveAs($filename1, $SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$SaveFormat = $wdSaveFormatPDF
			$doc.SaveAs($filename2, $SaveFormat)
		}
	}
	Else
	{
		#works for Server 2008 and Windows 7
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
		}
	}
}
Else
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
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Now saving as PDF"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
		$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
	}
}

Write-Verbose "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
If($PDF)
{
	Write-Verbose "$(Get-Date): Deleting $($filename1) since only $($filename2) is needed"
	Remove-Item $filename1 -EA 0
}
Write-Verbose "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word -Scope Global -EA 0
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	Write-Verbose "$(Get-Date): $($filename2) is ready for use"
}
Else
{
	Write-Verbose "$(Get-Date): $($filename1) is ready for use"
}
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