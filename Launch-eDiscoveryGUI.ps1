<#
    .NOTES
	
Needed a Graphical User Interface for eDiscovery tasks, for both Search-mailbox and New-MailboxSearch.
Needed a way to track search progress for New-MailboxSearch


BIG THANKS TO PRATEEK SINGH
His script gave be the base for this application:
	Based this tool on a script from the great Prateek Singh : http://en.gravatar.com/almostit - great guy !
	And the page of his script is here:
		https://geekeefy.wordpress.com/2016/02/26/powershell-gui-based-email-search-and-removal-in-exchange/

I developped around it to include:
- Bilingual Graphical Interface (logging in a file is still in English)
- Added the code for New-MailboxSearch and related options in addition to Search-Mailbox
- Powershell Cmdlet generator: as we click and type, we see the generation of the corresponding Powershell code
- Added the option to estimate only the search
- Added multiple tabs to include 
	- Connection button to Exchange Online or On-Premise environment
	- Exchange rights status displayed on the tool for the currently logged on user: Exchange Discovery Management + Exchange Import Export
	- Tabs that leverage Get-MailboxSearch, and permits to remove previous New-MailboxSearch generated searches


Credits for the Logging function goes to SAM BOUTROS
https://social.technet.microsoft.com/Profile/sam%20boutros
Initial link for the Log function here:
https://gallery.technet.microsoft.com/scriptcenter/Scriptfunction-to-display-7f7f36a9
	
    .DESCRIPTION
	
IMPORTANT - This script requires Exchange 201X commandlets - which you get either by clicking on the "Connect to Exchange" button, or by launching this script from an Exchange Management Shell.

Exchange Management Shell is actually a powershell console that we open, and inside which we launch a couple of command lines which effect is to connect to the Exchange environment, and to import additional commands that enable you to manage Exchange with the command line (like "Get-Mailbox" which will give you a list of mailboxes, "Search-Mailbox", "New-MailboxSearch", etc....)

This script is a GUI wrapper that launches Powershell commands: it pops up a Windows dialog box with a tabs, and from where you can launch the Mailbox Search cmdlets (Search-Mailbox, New-MailboxSearch) just with the click of a button – after having populated the things you wish to search (dates, keywords, mailboxes to search, etc...).
So normally, you won’t have to mess up with Powershell commands, even if the script outputs some of these – it’s just if you want to use Powershell if you’re curious about it 

NOTE: To remove the "Connect to Exchange" button, just comment the .Controls.Add() control on the tabPage0 configuration section

Release notes:
==============
1- Complete version, GUI bilingual French/English (Powershell underlying console in English), switch to connect to On-Premise or O365 Exchange
1.1 - Added a check on Search-Mailbox -DeleteContent parameter and grey out "Delete Mailbox" if we don't have the right to Delete
1.2 - normalized strings
1.3 - Added test on Discovery Mailbox - if it doesn't exist, inform the user - except for Search-Mailbox -EstimateOnly, which doesn't need a Discovery Mailbox.
1.401 - added and optimized the Welcome/Readme text in English and French
Fix: removed Get-Credentials -Message as the "-Message" parameter doesn't work on Powershell v2.0 (Windows 2008 R2 Powershell).
1.402 - removed "check spelling of Discovery mailbox" when error when one mailbox is not found. Discovery Mailbox check is done separately
1.5 - Added logging into a file
LOG File is always created under User's profile \ Documents folder. Log file name hard coded (search for string:
"$($env:userprofile)\Documents\eDiscoveryLog.txt"
Can always change the log file destination by using the "-LogFile" parameter after each call of the "Log" function
1.5.1 - case Source Mailboxes to search in is blank
Updated Update-CmdLine when SourceMailboxes is blank
Add confirmation to continue searching all mailboxes (Continue ? Yes/No)
Updated btn_Run function -> using the command line built by Update-CmdLine function (reusing $richTextBox.Text) - NOTE: btn_Run has additionnal tests
1.5.2 - FR and US translation for "search in all mailboxes" confirm box.
1.5.3 - added filter on Get-mailbox to exclude search in Discovery Mailbox when using Search-Mailbox
1.5.4 - fixed Right to Delete variable $RightToDelete - was set in the Test-ExchCx function but was local to
this function only. Fix : make it global for the whole application.
1.5.4.1 - 09 MAR 2017 - changed a bit the Log function to add default string in case no String parameter is specified
1.5.4.2 - 09 MAR 2017 - removed the "-LogLevel full" if "-DeleteContent" is checked...

1.5.5 - Review code and remove a few useless comments before publishing
1.5.6 - Adding suggestions (no code modification) - Publishing version
18AUG2017 - incrementing minor version to get rid of minor-minor version
Suggestion : add time in date selection for searches.
1.5.6.1 - fixing Get-Mailbox -REsultSize Unlimited | Search-Mailbox / changing to Get-MailboxDatabase | ForEach {Get-Mailbox $_ -resultsize Unlmited | Search-Mailbox blablablabla}
1.5.7 - Change to Delete e-mails option using Search-Mailbox : instead of loading all mailboxes of an Org and piping all on Search-Mailbox, 
I load all MAilboxDatabases first, and using ForEach ($Database in $DatabaseList) {$MBXBatch = Get-Mailbox -Database $Database; and piping $MBXBatch to Search-Mailbox instead.
1.5.7.1 - Encoding issue - re-saved the file with UTF-8 encoding - French accents were messing up with the whole script

Intention for vNext:
- Add CSV file picker to choose .CSV file as input for list of mailboxes to search in
Would also like to add function to test if each mailbox in the list exist before launching the search
- Add date picker, and if no date selected, remove date filters in Search-Mailbox and New-MailboxSearch generated command lines
- Use -Force for Search-Mailbox instead of -Confirm $false which is not taken into account (Search-Mailbox with 
"-DeleteContent" continues to prompt for deletions when using Get-Mailbox | Search-Mailbox...


#>
#Switch the $TestMode variable to $true for tool testing, and to $false when finished testing
$TestMode = $false


$OrganizationNameEN = "Microsoft Exchange Search"
$OrganizationNameFR = "Recherche Microsoft Exchange"
$CurrentVersion = "1.5.7.1"

#------------------------------------------------
#region Application Functions
#------------------------------------------------
Function Split-ListSemicolon
{
    param(
        [string]$StringToSplit
    )
    $TargetEmailSplit = $StringToSplit.Split(';')
    $SourceMailboxes = ""
    For ($i = 0; $i -lt $TargetEmailSplit.Count - 1; $i++) {$SourceMailboxes += ("""") + $TargetEmailSplit[$i] + (""",")}
    $SourceMailboxes += ("""") + $TargetEmailSplit[$TargetEmailSplit.Count - 1] + ("""")
    Return $SourceMailboxes
}
	
Function Validate-Mailboxes ($Mailboxes)
{
    # Placeholder function for future version .... test existence of mailboxes before launching the search function
    # (can save time ... have to test)
}
	
function Log
{
    <# 
 .Synopsis
  Function to log input string to file and display it to screen

 .Description
  Function to log input string to file and display it to screen. Log entries in the log file are time stamped. Function allows for displaying text to screen in different colors.

 .Parameter String
  The string to be displayed to the screen and saved to the log file

 .Parameter Color
  The color in which to display the input string on the screen
  Default is White
  Valid options are
    Black
    Blue
    Cyan
    DarkBlue
    DarkCyan
    DarkGray
    DarkGreen
    DarkMagenta
    DarkRed
    DarkYellow
    Gray
    Green
    Magenta
    Red
    White
    Yellow

 .Parameter LogFile
  Path to the file where the input string should be saved.
  Example: c:\log.txt
  If absent, the input string will be displayed to the screen only and not saved to log file

 .Example
  Log -String "Hello World" -Color Yellow -LogFile c:\log.txt
  This example displays the "Hello World" string to the console in yellow, and adds it as a new line to the file c:\log.txt
  If c:\log.txt does not exist it will be created.
  Log entries in the log file are time stamped. Sample output:
    2014.08.06 06:52:17 AM: Hello World

 .Example
  Log "$((Get-Location).Path)" Cyan
  This example displays current path in Cyan, and does not log the displayed text to log file.

 .Example 
  "$((Get-Process | select -First 1).name) process ID is $((Get-Process | select -First 1).id)" | log -color DarkYellow
  Sample output of this example:
    "MDM process ID is 4492" in dark yellow

 .Example
  log "Found",(Get-ChildItem -Path .\ -File).Count,"files in folder",(Get-Item .\).FullName Green,Yellow,Green,Cyan .\mylog.txt
  Sample output will look like:
    Found 520 files in folder D:\Sandbox - and will have the listed foreground colors

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/06/2014
  v1.1 - 12/01/2014 - added multi-color display in the same line

#>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')] 
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipeLineByPropertyName = $true,
            Position = 0)]
        [AllowNull()]
        # added default "No Log Specified"
        [String[]]$String = "No Log Specified", 
        [Parameter(Mandatory = $false,
            Position = 1)]
        [ValidateSet("Black", "Blue", "Cyan", "DarkBlue", "DarkCyan", "DarkGray", "DarkGreen", "DarkMagenta", "DarkRed", "DarkYellow", "Gray", "Green", "Magenta", "Red", "White", "Yellow")]
        [String[]]$Color = "DarkGreen", 
        [Parameter(Mandatory = $false,
            Position = 2)]
        [String]$LogFile = "$($env:userprofile)\Documents\eDiscoveryLog.txt",
        [Parameter(Mandatory = $false,
            Position = 3)]
        [Switch]$NoNewLine
    )

    if ($String.Count -gt 1)
    {
        $i = 0
        foreach ($item in $String)
        {
            if (($item -eq $null) -or ($item -eq "")) { $item = 'null' }
            if ($Color[$i]) { $col = $Color[$i] } else { $col = "White" }
            Write-Host "$item " -ForegroundColor $col -NoNewline
            $i++
        }
        if (-not ($NoNewLine)) { Write-Host " " }
    }
    else
    { 
        if ($NoNewLine) { Write-Host $String -ForegroundColor $Color[0] -NoNewline }
        else { Write-Host $String -ForegroundColor $Color[0] }
    }

    if ($LogFile.Length -gt 2)
    {
        "$(Get-Date -format "yyyy.MM.dd hh:mm:ss tt"): $($String -join " ")" | Out-File -Filepath $Logfile -Append 
    }
    else
    {
        Write-Verbose "Log: Missing -LogFile parameter. Will not save input string to log file.."
    }
}
	
#------------------------------------------------
#endregion Application Functions
#------------------------------------------------

#------------------------------------------------
# Form Function - This is the main tool
#------------------------------------------------
function Search-MailboxGUI
{
    #Parameters - only language selection in this case
    param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$language
    )

    #------------------------------------------------
    #region ########### Manage the locale strings here ! ################
    #------------------------------------------------
    Switch ($language)
    {	
        "EN"
        {
            #region Locale = EN
            #Menus, buttons, labels ...
            $Str000 = "eDiscovery Tool"
            $Str001 = "Welcome !"
            $Str002 = "$OrganizationNameEN"
            $Str003 = "Exchange 2010, 2013, 2016, Exchange Online"
            $Str004 = "Ready !"
            $Str004a = "Please wait while working..."
            $Str004b = "Please wait while closing existing sessions ..."
            $Str005 = "Connect to Exchange"
            $Str006 = "Connection Status:"
            $Str007 = "Not Connected to Exchange"
            $Str007a = "Connected to Exchange"
            $Str007b = "Able to search !"
            $Str007c = "Discovery Rights not present !"
            $Str007d = "Connecting..."
            $Str007e = "Can delete mails from source."
            $Str007f = "Cannot delete mails from source."
            $Str008 = "Search in mailboxes"
            $Str009 = "Recipient E-mail, separated by semi-colons:"
            $Str010 = "Sender E-mail"
            $Str011 = "Attachment name"
            $Str012 = "Search Keyword"
            $Str013 = "Search Only in Subject Line"
            $Str014 = "Start Date (MM/DD/YYYY)"
            $Str015 = "End Date (MM/DD/YYYY)"
            $Str016 = "Delete E-mail from Source ? NOTE: 'Mailbox Import Export' right required"
            $Str017 = "Search a bit quicker using New-MailboxSearch (cannot delete mail)"
            $Str018 = "Estimate only (don't store results)"
            $Str019 = "Mailbox to store E-mails in:"
            $Str020 = "Folder:"
            $Str021 = "Launch Search"
            $Str022 = "Close the tool"
            $Str023 = "Retrieve Mailbox Searches"
            $Str024 = "This area is there to be able to request and view existing searches launched with the ""New-MailboxSearch"" option checked..."
            $Str025 = "Get previous mailbox search"
            $Str026 = "Get ALL mailbox searches"
            $Txt002 = "Working to get mailbox search"
            $Txt002a = "defined on tab ""Search In Mailboxes"" please wait..."
            $Txt002b = "Working to get ALL mailbox searches still on the system, please wait..."
            $Txt003 = "Cannot find any search named"
            $Txt003a = "... try another ""Folder:"" name on the ""Search Mailbox"" tab, or click on ""Get ALL mailbox searches..."" button."
            $TxtLbl001 = "Name"
            $TxtLbl002 = "% Progress"
            $TxtLbl003 = "Nb findings"
            $Str027 = "Retrieve Mailbox Searches Statistics"
            $Str028 = "Populate or update drop down list"
            $Str029 = "Click the below link to go to the mailbox where the results are stored:"
            $Str029a = "Click the below link to fo to the mailbox ("
            $Str029b = ") where the results are stored for"
            $Str030 = "Results Link"
            $Txt005 = "Please click on the ""Populate or update drop down list"" button to begin retrieving statistics about previousliy run Mailbox Searches (NOTE: only previous searches done with ""New-MailboxSearch"" are displayed)"
            $Txt006 = "Working on getting Mailbox Search names currently on the system..."
            $Txt007 = "Found"
            $Txt007a = "searches on the system !"
            $Txt008 = "Use the drop down list above to get more statistical information about a particular Mailbox Search !"
            $Txt009 = "Please wait, we're gathering the mailbox search statistics..."
            $Txt010 = "Showing the search statistics for"
            $Txt011 = "No mailbox searches found ... do some and come back later on this tab !"
            $Txt012 = "No link for this search. Probably an estimate."
            $Txt013 = "No valid link for this search..."
            $Txt014 = "Search interface terminated."
            $Err001 = "An error occurred - most likely one of the mailboxes doesn't exist or the there is a typo in one of the mailboxes name or address..."
            $Err002 = "A Mailbox Search with the same name already exists. Please type another one in the ""Folder"" field of the GUI on the ""Research Mailboxes"" tab...`nor remove the existing mailbox search, and run the script again..."
            $Err003 = "Wrong creds or no creds entered ..."
            $Str031 = "Do you really want to remove that MailboxSearch ?"
            $Str032 = "not deleted, returning to the search tool..."
            $Txt015 = "Deleting selected mailbox search, please wait..."
            $Str033 = "was deleted !"
            $Str034 = "On-Premise Exchange"
            $Str035 = "O365 Exchange"
            $Err004 = "On-Premise Exchange connection failed - wrong credentials or server URL connection... please try again or connect to an O365 Exchange environment."
            $Txt001 = "Welcome to the eDiscovery Graphical Interface tool.`n`nHere is a quick description of the tabs:`n`n`n1- the ""$Str001"" tab`n=================================================`n>> it's this tab.`n`nConnect to Exchange Online or to Exchange on-premise. You will also see the status of your Powershell session:`n- ""$Str007""`n--->if you are not in a Powershell session with Exchange cmdlets`n- ""$Str007a""`n--->if you are in a Powershell session with Exchange cmdlets`n- ""$Str007b""`n--->if you have the rights to execute the ""Search-Mailbox"" and the ""New-MailboxSearch"" cmdlets, which are required to use this tool`n- ""$Str007c""`n--->if you don't have the ""Search-Mailbox"" and the ""New-MailboxSearch"" rights brought by the ""Discovery Management"" RBAC role`n- ""$Str007d""`n--->when you are importing an Exchange Powershell session O365 or Exchange on-premise`n- ""$Str007e""`n--->if you have the ""Mailbox Import Export"" rights which are required to be able to remove e-mails searched`n- ""$Str007f""`n--->if you don't have the ""Mailbox Import Export"" rights`n`n2- the ""$Str008"" tab`n=================================================`n>> it's the tab where you search, estimate searches, or purge searched e-mails from mailboxes (malicious, sensitive, ...)`n`n> This tab provides the Powershell cmdlet that correspond to the search options and parameters filled. This cmdlet is automatically generated, in real time. You can either launch the search using the ""$Str021"" button, or just copy and paste the generated cmdlet on another Exchange Powershell session.`n> Also, this tab enabled you to choose whether you prefer to use Search-Mailbox or New-MailboxSearch to perform your search.`n*** NOTE1:  Search-Mailbox has an option to purge the source mailbox, while New-MailboxSearch is read-only.`n*** NOTE2: Also note that New-MailboxSearch keep the searches statistics on the Exchange server, while Search-Mailbox doesn't.`nHowever, both cmdlets also store their search results on the Discovery Mailbox (except Search-Mailbox -EstimateOnly) – New-MailboxSearch also stores the search results statistics on the Exchange servers, as well as on the Discovery Mailbox.`n`n3- the ""$Str023"" tab`n==================================================`n>> it's the tab where you can followup with the Search status of the mailboxes.`n*** NOTE: searches performed with ""New-MailboxSearch"" cmdlet are performed on the remote Exchange server, and you don't see the search progress right away. You have to use this tab.`n`n4- the ""$Str027"" tab`n==================================================`n>> this tab enables you to view more detailed search statistics for searches performed with ""New-MailboxSearch"". `n`nThis tab also provides you with the ability to:`n> Access the Discovery Mailbox directly by a provided URL, which updates upon the selected Mailbox Search`n> Delete previous searches, including the results stored in the corresponding Discovery Mailbox`n`nEnjoy this tool !"
            $Err005 = "The discovery mailbox for storing search results and statistics does not exist - please put an existing name or E-mail address for a discovery mailbox which you have access to."
            $Str036 = "No e-mail addresses specified => will scan all mailboxes, do you want to continue ?"
            #endregion Locale = EN
        }
        "FR"
        {	
            #region Locale = FR
            $Str000 = "Outil de recherche"
            $Str001 = "Bienvenue !"
            $Str002 = "$OrganizationNameFR"
            $Str003 = "Exchange 2010, 2013, 2016, Exchange en-ligne"
            $Str004 = "Prêt !"
            $Str004a = "Veuillez patienter, opération en cours..."
            $Str004b = "Veuillez patienter pendant la fermeture des sessions existentes..."
            $Str005 = "Se connecter à Exchange"
            $Str006 = "Statut de connexion:"
            $Str007 = "Non connecté à Exchange"
            $Str007a = "Connecté à Exchange"
            $Str007b = "Vous avez les droits de recherche !"
            $Str007c = "Vous n'avez pas les droits de recherche."
            $Str007d = "Connexion en cours..."
            $Str007e = "Vous avez les droits de suppression de courriels."
            $Str007f = "Vous ne pouvez pas supprimer de courriels."
            $Str008 = "Rechercher dans les boîtes aux lettres"
            $Str009 = "Adresses courriel, séparées par des point-virgules:"
            $Str010 = "Adresse de l'expéditeur"
            $Str011 = "Nom de pièce jointe"
            $Str012 = "Mot-clé de recherche"
            $Str013 = "Rechercher dans le sujet seulement"
            $Str014 = "Date de début (MM/JJ/AAAA)"
            $Str015 = "Date de fin (MM/JJ/AAAA)"
            $Str016 = "Supprimer le courriel de la boîte aux lettres ? NOTE: la permission 'Mailbox Import Export' est requise"
            $Str017 = "Rechercher plus rapidement avec New-MailboxSearch (ne peut supprimer de courriel)"
            $Str018 = "Estimer seulement (ne pas stocker les résultats)"
            $Str019 = "Boîte aux lettres de résultats:"
            $Str020 = "Répertoire:"
            $Str021 = "Lancer la recherche"
            $Str022 = "Fermer l'outil"
            $Str023 = "Retrouver les recherches précédentes"
            $Str024 = "Cette partie de l'outil permet le contrôle de l'état de la dernière recherche effectuée ainsi que de toutes les dernières recherches effectuées à l'aide de l'option ""New-MailboxSearch""..."
            $Str025 = "Afficher recherche précédente"
            $Str026 = "Afficher TOUTES les recherches"
            $Txt002 = "Récupération de la recherche"
            $Txt002a = "définie dans l'onglet ""Recherche dans les boîtes aux lettres"", veuillez patienter..."
            $Txt002b = "Récupération des données d'état de TOUTES les recherches disponibles dans le système, veuillez patienter..."
            $Txt003 = "Impossible de trouver une recherche nommée"
            $Txt003a = "... essayez un autre nom de ""Répertoire"" dans l'onglet ""Recherche dans les boîtes aux lettres"", ou cliquer sur le bouton ""Afficher TOUTES les recherches""."
            $TxtLbl001 = "Nom"
            $TxtLbl002 = "% Progression"
            $TxtLbl003 = "Nb résultats"
            $Str027 = "Statistiques des recherches"
            $Str028 = "Mettre à jour liste déroulante"
            $Str029 = "Clicker sur le lien ci-dessous pour ouvrir la boite aux lettres où sont stockés les résultats de la recherche:"
            $Str029a = "Clicker sur le lien ci-dessous pour ouvrir la boite ("
            $Str029b = ") dans laquelle les résultats sont stockés pour"
            $Str030 = "Lien vers la boîte aux lettres des résultats"
            $Txt005 = "Veuillez clicker sur le bouton ""Mettre à jour liste déroulante"" pour commencer à retrouver les statistiques des recherches précédentes (NOTE: seules les recherches effectuées à l'aide de ""New-MailboxSearch"" sont affichées)"
            $Txt006 = "Recherche des précédentes recherches actuellement dans le système..."
            $Txt007 = "On a trouvé"
            $Txt007a = "recherches sur le système !"
            $Txt008 = "Utilisez la liste déroulante ci-dessus afin de récupérer les statistiques détaillées sur une recherche en particulier !"
            $Txt009 = "Veuillez patienter, nous récupérons les statistiques de recherches ..."
            $Txt010 = "Statistiques trouvées pour"
            $Txt011 = "Aucune recherche trouvée ... faites quelques recherches puis revenez plus tard dans cet onglet !"
            $Txt012 = "Aucun lien disponible pout cette recherche... il s'agit probablement d'une recherche avec ""Estimation"" seulement."
            $Txt013 = "Pas de lien valide pour cette recherche..."
            $Txt014 = "Interface de recherche fermée."
            $Err001 = "Une erreur est survenue - il se peut que l'une des boîtes aux lettres n'existe pas ou qu'il y ait une faute de frappe dans l'adresse courriel ou le nom de l'une des boîtes recherchées... "
            $Err002 = "Une recherche de boîtes aux lettres portant le même nom existe déjà. Veuillez utiliser un autre nom dans le champ ""Répertoire:"" de l'onglet ""Recherche...""...`nou veuillez supprimer la recherche du même nom puis réessayez la recherche à nouveau..."
            $Err003 = "Mauvais compte et/ou mot de passe entrés ..."
            $Str031 = "Voulez-vous vraiment supprimer cette recherche ?"
            $Str032 = "pas supprimée, retour à l'outil de recherche..."
            $Txt015 = "Suppression de la recherche sélectionnée, veuillez patienter..."
            $Str033 = "a été supprimée !"
            $Str034 = "Exchange privé ou local"
            $Str035 = "Exchange O365"
            $Err004 = "Erreur de connexion à un environnement Exchange privé - mauvais login&mot de passe ou mauvaise URL de connexion... veuillez ré-essayer ou connectez-vous à un environnement O365 Exchange."
            $Txt001 = "Bienvenue dans l’application de recherche de courriel.`n`nVoici une description rapide des onglets :`n`n`n1- Onglet ""$Str001""`n=================================================`n>> vous y êtes.`n`nConnectez-vous à Exchange en-ligne (O365) ou à une infrastructure Exchange locale. Vous verrez le statut de votre session Powershell:`n- ""$Str007""`n-----> Aucune commande Exchange n’est disponible`n- ""$Str007a""`n-----> Les commandes principales Exchange sont disponibles`n- ""$Str007b""`n-----> En plus d’avoir les commandes Exchanges, vous avez accès aux commandes de recherche (nécessite les droits ""Search-Mailbox"" et ""New-MailboxSearch"" conférés par le rôle RBAC ""Discovery Management"")`n- ""$Str007c""`n-----> Vous avez les commandes Exchange disponibles, mais vous n’avez pas de droits de recherche`n- ""$Str007d""`n-----> L’outil est en cours de tentative de connextion à un environnement Exchange`n- ""$Str007e""`n-----> Vous avez les droits de suppression de mails à l’aide des commandes de recherche ! (Droits ""Mailbox Import Export"")`n- ""$Str007f""`n-----> Vous n’avez pas les droits de suppression de mails conférés par le rôle d’administration ""Mailbox Import Export""`n`n2- Onglet ""$Str008""`n=================================================`n>> onglet de recherche, estimations, ou purge de courriels recherchés (malicieux, sensibles, ...)`n`n> cet onglet génère en temps réel la ligne de commande Powershell complète. Vous pouvez lancer une recherche directement depuis l’interface avec le bouton ""$Str021"" ou simplement copier et coller la ligne de commande générée par l’application dans une nouvelle fenêtre Powershell ou Exchange Management Shell.`n> vous pouvez choisir d’utiliser soit la méthode ""New-MailboxSearch"" (activée par défaut), soit la méthode ""Search-Mailbox"".`n*** NOTE1 : ""Search-Mailbox"" est la seule méthode permettant de purger une boîte de certains messages, alors que ""New-MailboxSearch"" est une méthode de recherche en lecture seule.`n*** NOTE2 : les deux commandes enregistrent leurs résultats de recherche dans une boîte aux lettres de découverte (sauf Search-Mailbox -EstimateOnly) - ""New-MailboxSearch"" sauvegarde également les statistiques de recherches sur le serveur Exchange ainsi que dans la boîte de recherche.`n`n3- Onglet ""$Str023""`n=================================================`n>> permet de suivre la progression des recherches effectuées à l’aide de ""New-MailboxSearch"".`n*** NOTE : ""New-MailboxSearch"" envoie une demande de recherche dans une file d’attente au niveau des serveurs Exchange. Cet onglet permet de suivre et d’afficher le statut et les statistiques de ces recherches.`n`n4- Onglet ""$Str027""`n=================================================`n>> permet d’afficher les détails des recherches effectuées à l’aide de ""New-MailboxSearch""`nCet onglet permet également :`n> le lancement d’une session Outlook Web App pour l’accès à la boîte de découverte utilisée pour la recherche dont vous avez affiché les informations.`n> la suppression de recherches précédentes, incluant les résultats de ces recherches stockés dans la boîte de recherche correspondante.`n`nBonnes recherches !"
            $Err005 = "La boîte aux lettre de découverte pour le stockage des résultats et statistiques de recherches n'existe pas - veuillez mettre un nom de boîte aux lettres de découverte qui existe et à laquelle vous avez accès..."
            $Str036 = "Aucune addresse e-mail spécifiée => recherche dans toutes les boîtes aux lettres, on continue ?"
            #endregion Locale FR
        }
    }

    $Str000 += " - v" + $CurrentVersion

    Log "***********************************************************************************"
    Log "Search Tool $Str000 Welcome !"
    Log "Logging Started"
    Log "***********************************************************************************"

    #------------------------------------------------
    #endregion ########### Manage the locale strings here ! ################
    #------------------------------------------------

    #----------------------------------------------
    #region Import the Assemblies
    #----------------------------------------------
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  | Out-Null
    #----------------------------------------------
    #endregion Import Assemblies
    #------------------------------------------------

    #----------------------------------------------
    #region Form Objects and Elements Instantiation
    #----------------------------------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    ##Form object instantiation
    $frmSearchForm = New-Object 'System.Windows.Forms.Form'
    ## Error Box message 
    $MsgBoxError = [System.Windows.Forms.MessageBox]
    ##tabbing - step 1 - create Tab Control object to put in the form, and also the Tab Pages objects to put in Tab Control later on
    ##Tab Control object instantiation - it's the Binder for the tabs
    $tabcontrol = New-Object 'System.Windows.Forms.TabControl'
    ##Tab pages instantiation - these are the Tabs we put in the Binder=tab control
    $tabPage0 = New-Object 'System.Windows.Forms.TabPage'
    $tabPage1 = New-Object 'System.Windows.Forms.TabPage'
    $tabPage2 = New-Object 'System.Windows.Forms.TabPage'
    $tabPage3 = New-Object 'System.Windows.Forms.TabPage'
    ####Tab 0 items - Title, Subtitle, text box to put some welcome text####
    $lblabout = New-Object System.Windows.Forms.Label
    $lblBigTitle2 = New-Object System.Windows.Forms.Label
    $lblBigTitle1 = New-Object System.Windows.Forms.Label
    $richtxtWelcome = New-Object System.Windows.Forms.RichTextBox
    $btnTab0ConnectExch = New-Object System.Windows.Forms.Button
    $lblTab0CxStatus = New-Object System.Windows.Forms.Label
    $lblTab0CxStatusUpdate = New-Object System.Windows.Forms.Label
    $txtTab0ConnectionURI = New-Object System.Windows.Forms.TextBox
    $lblSwitchOnPremOnCloud = New-Object System.Windows.Forms.Label
    $chkOnPrem = New-Object System.Windows.Forms.CheckBox
    ####Tab 1 items####
    $txtAttachment = New-Object 'System.Windows.Forms.TextBox'
    $chkAttachment = New-Object 'System.Windows.Forms.CheckBox'
    $btnCancel = New-Object 'System.Windows.Forms.Button'
    $btnRun = New-Object 'System.Windows.Forms.Button'
    $txtDiscoveryMailbox = New-Object 'System.Windows.Forms.TextBox'
    $lblDiscoveryMailbox = New-Object 'System.Windows.Forms.Label'
    $txtDiscoveryMailboxFolder = New-Object 'System.Windows.Forms.TextBox'
    $lblDiscoveryMailboxFolder = New-Object 'System.Windows.Forms.Label'
    $chkDeleteMail = New-Object 'System.Windows.Forms.CheckBox'
    $txtEndDate = New-Object 'System.Windows.Forms.TextBox'
    $lblEndDate = New-Object 'System.Windows.Forms.Label'
    $txtStartDate = New-Object 'System.Windows.Forms.TextBox'
    $lblStartDate = New-Object 'System.Windows.Forms.Label'
    $chkSubject = New-Object 'System.Windows.Forms.CheckBox'
    $lblKeyword = New-Object 'System.Windows.Forms.Label'
    $txtKeyword = New-Object 'System.Windows.Forms.TextBox'
    $txtSender = New-Object 'System.Windows.Forms.TextBox'
    $chkSender = New-Object 'System.Windows.Forms.CheckBox'
    $chkUseNewMailboxSearch = New-Object 'System.Windows.Forms.CheckBox'
    $txtRecipient = New-Object 'System.Windows.Forms.TextBox'
    $lblRecipient = New-Object 'System.Windows.Forms.Label'
    $richTxtCurrentCmdlet = New-Object System.Windows.Forms.RichTextBox
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
    $chkEstimateOnly = New-Object System.Windows.Forms.CheckBox
    $global:RightToDelete = $true
    ####Tab 2 items####
    $lblTab2ExistingSearches = New-Object 'System.Windows.Forms.Label'
    $btnTab2GetMbxSearches = New-Object 'System.Windows.Forms.Button'
    $btnTab2Get1MbxSearch = New-Object System.Windows.Forms.Button
    $richtxtGetMbxSearch = New-Object 'System.Windows.Forms.RichTextBox'
    ####Tab 3 items####
    $lblTab4URL = New-Object System.Windows.Forms.Label
    $lblTab4URLLink = New-Object System.Windows.Forms.LinkLabel
    $txtTab4MbxSearchStats = New-Object System.Windows.Forms.RichTextBox
    $comboTab4MbxSearches = New-Object System.Windows.Forms.ComboBox
    $btnTab4PopList = New-Object System.Windows.Forms.Button
    $btnDelSearch = New-Object System.Windows.Forms.Button
	
    $StatusBar = New-Object System.Windows.Forms.StatusBar
    #----------------------------------------------
    #endregion Form Objects Instantiation
    #----------------------------------------------

    #----------------------------------------------
    #region Application Functions using variables defined within the form function
    #----------------------------------------------
    function Update-CmdLine()
    {
        $TargetEmail = $txtRecipient.Text
        $DiscoveryMailbox = $txtDiscoveryMailbox.Text
        $DiscoveryFolder = $txtDiscoveryMailboxFolder.Text
        $SenderEmail = $txtSender.Text
        $Keyword = ('''') + $txtKeyword.Text + ('''')
        $StartDate = $txtStartDate.Text
        $EndDate = $txtEndDate.Text
        $DateInt = $StartDate + ".." + $EndDate
        $richTxtCurrentCmdlet.ForeColor = 'Green'
				
        if ($chkSender.Checked -eq $true)
        { 
            $From = (" AND from:") + ('''') + $txtSender.Text + ('''')
            $FromMultiSenders = Split-ListSemicolon($txtSender.Text); $FromMulti = (" -Senders ") + $FromMultiSenders
        }
        else 
        { 
            $From = "" 
            $FromMulti = ""
        }
		
        if ($chkAttachment.Checked -eq $true) { $Attachment = (" AND attachment:") + ('''') + $txtAttachment.Text + ('''')}
        else { $Attachment = "" }
        if ($chkSubject.Checked -eq $true) { $Keyword = ("Subject:") + $Keyword }

        $SearchQuery = ("""") + "$Keyword$From$Attachment" + " AND Received:$DateInt" + ("""")
        $SearchQueryMulti = ("""") + "$Keyword$Attachment" + ("""")
		
        # Introducing $BlankSource boolean -> if no e-mail addresses specified, search all mailboxes 
        # To search all mailboxes: remove -SourceMailboxes parameters for New-SeachMailbox, and add "Get-Mailboxes" before the Search-Mailbox
        if (($TargetEmail -eq "") -or ($TargetEmail -eq $null)) {$BlankSource = $true} else {$BlankSource = $false; $SourceMailboxes = Split-ListSemicolon($TargetEmail)}

        If ($chkUseNewMailboxSearch.Checked -eq $true)
        {
            If ($BlankSource -eq $false)
            {
                $CommandMulti = "New-MailboxSearch -Name '$DiscoveryFolder' -SourceMailboxes $SourceMailboxes -TargetMailbox '$DiscoveryMailbox' -StartDate '$StartDate' -EndDate '$EndDate' -SearchQuery $SearchQueryMulti -ExcludeDuplicateMessages `$false -ErrorAction Continue" + $FromMulti
            }
            else # $BlankSource -eq $true
            { 
                $CommandMulti = "New-MailboxSearch -Name '$DiscoveryFolder' -TargetMailbox '$DiscoveryMailbox' -StartDate '$StartDate' -EndDate '$EndDate' -SearchQuery $SearchQueryMulti -ExcludeDuplicateMessages `$false -ErrorAction Continue" + $FromMulti
            }
            If ($chkEstimateOnly.Checked -eq $true)
            {
                $CommandMulti += " -EstimateOnly"
            }
            $richTxtCurrentCmdlet.Text = $CommandMulti
        }
        Else #If ($chkUseNewMailboxSearch.Checked -eq $false)
        {
            If ($BlankSource -eq $false)
            {
                $Command = $SourceMailboxes + (" | ")
            }
            else
            {
                # $blankSource -eq $true
                $Command = '$AllDatabases = get-MailboxDatabase; ForEach ($DB in $AllDatabases) {$MBXBatch = get-mailbox -ResultSize unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"} -Database $DB ; $MBXBatch | '
            }
            If ($chkDeleteMail.Checked -eq $true)
            {
                $Command += "Search-Mailbox -TargetMailbox '$DiscoveryMailbox' -TargetFolder '$DiscoveryFolder' -SearchQuery $SearchQuery -Verbose -DeleteContent -Confirm:`$false -Force"
                If ($BlankSource) {$Command += "}"}
                # Update status bar
                $StatusBar.Text = $Str004
            }
            Else # $chkDeleteMail -eq $false
            {
                If ($chkEstimateOnly.Checked -eq $true)
                {
                    $Command += "Search-Mailbox -SearchQuery $SearchQuery -Verbose -EstimateResultOnly"
                    If ($BlankSource) {$Command += "}"}
                }
                Else
                {
                    #chkEstimateOnly -eq $false
                    $Command += "Search-Mailbox -TargetMailbox '$DiscoveryMailbox' -TargetFolder '$DiscoveryFolder' -LogLevel full -SearchQuery $SearchQuery -Verbose"
                    If ($BlankSource) {$Command += "}"}
                }	
            }
            $richTxtCurrentCmdlet.Text = $Command
        }
    }
	
    function Populate-DropDownList
    {
        $btnDelSearch.Enabled = $false
        $lblTab4URL.Text = $Str029
        $lblTab4URLLink.Text = $Str030
        $lblTab4URLLink.Enabled = $false
        $StatusBar.Text = $Str004a
        $txtTab4MbxSearchStats.Text = $Txt006
        Log $Txt006
        #region Gather mailbox searches
        $ListOfSearchNames = Get-MailboxSearch | Select Name | Out-String
        $ListOfSearchNames = $ListOfSearchNames.Split("`n")
        $ListOfSearchNamesOffset = @(0) * ($ListOfSearchNames.Count - 6)
        For ($i = 3; $i -le ($ListOfSearchNames.Count - 4); $i++)
        {
            $ListOfSearchNamesOffset[$i - 3] = $ListOfSearchNames[$i].trim()
            # For debugging only #Write-Host "Item $i => $($ListOfSearchNames[$i]) -- Offset Item $($i-3) => $($ListOfSearchNamesOffset[$i-3])`n"
        }
        #endregion gather mailbox searches
        If ($ListOfSearchNames.Count -gt 0)
        {
            $combotab4MbxSearches.Items.Clear()
            $comboTab4MbxSearches.Items.AddRange($ListOfSearchNamesOffset)
            $LogString = $Txt007 + (" $($ListOfSearchNamesOffset.count) ") + $Txt007a
            $txtTab4MbxSearchStats.Text = $LogString
            Log $LogString
            $txtTab4MbxSearchStats.Text += ("`n`n") + $Txt008
			
        }
        Else
        {
            $txtTab4MbxSearchStats.Text = $Txt011
            Log $Txt011
        }
        $StatusBar.Text = $Str004
    }
	
    Function Test-ExchCx()
    {
        Try
        {
            Get-command Get-mailbox -ErrorAction Stop
            $lblTab0CxStatusUpdate.ForeColor = 'Green'
            $lblTab0CxStatusUpdate.Text = $Str007a
            Try
            {
                Get-command Get-MailboxSearch -ErrorAction Stop
                $lblTab0CxStatusUpdate.Text += ("`n") + $Str007b
                $tabPage1.enabled = $true
                $tabPage2.enabled = $true
                $tabPage3.enabled = $true
				
                Try
                {
                    Get-Command Search-Mailbox -ParameterName DeleteContent -ErrorAction Stop | Out-Null
                    $lblTab0CxStatusUpdate.Text += ("`n") + $Str007e
                }
                Catch
                {
                    $lblTab0CxStatusUpdate.ForeColor = 'Blue'
                    $lblTab0CxStatusUpdate.Text += ("`n") + $Str007f
                    $Global:RightToDelete = $false
                    $chkDeleteMail.Enabled = $false
                }	
            }
            Catch [System.SystemException]
            {
                $lblTab0CxStatusUpdate.ForeColor = 'Orange'
                $lblTab0CxStatusUpdate.Text += ("`n") + $Str007c
                $tabPage1.enabled = $false
                $tabPage2.enabled = $false
                $tabPage3.enabled = $false
            }
        }
        Catch [System.SystemException]
        {
            $lblTab0CxStatusUpdate.ForeColor = 'Red'
            $lblTab0CxStatusUpdate.Text = $Str007
            $tabPage1.enabled = $false
            $tabPage2.enabled = $false
            $tabPage3.enabled = $false
        }
    }

    Function Test-ExchCxTest()
    {
        Try
        {
            Get-command Get-mailbox -ErrorAction Stop
            $lblTab0CxStatusUpdate.ForeColor = 'Green'
            $lblTab0CxStatusUpdate.Text = $Str007a
            Try
            {
                Get-command Get-MailboxSearch -ErrorAction Stop
                $lblTab0CxStatusUpdate.Text += ("`n") + $Str007b
                $tabPage1.enabled = $true
                $tabPage2.enabled = $true
                $tabPage3.enabled = $true
				
                Try
                {
                    Get-Command Search-Mailbox -ParameterName DeleteContent -ErrorAction Stop | Out-Null
                    $lblTab0CxStatusUpdate.Text += ("`n") + $Str007e
                }
                Catch
                {
                    $lblTab0CxStatusUpdate.ForeColor = 'Blue'
                    $lblTab0CxStatusUpdate.Text += ("`n") + $Str007f
                    $Global:RightToDelete = $false
                    #					$chkDeleteMail.Enabled = $false
                }	
            }
            Catch [System.SystemException]
            {
                $lblTab0CxStatusUpdate.ForeColor = 'Orange'
                $lblTab0CxStatusUpdate.Text += ("`n") + $Str007c
                #				$tabPage1.enabled = $false
                $tabPage2.enabled = $false
                $tabPage3.enabled = $false
            }
        }
        Catch [System.SystemException]
        {
            $lblTab0CxStatusUpdate.ForeColor = 'Red'
            $lblTab0CxStatusUpdate.Text = $Str007
            #			$tabPage1.enabled = $false
            $tabPage2.enabled = $false
            $tabPage3.enabled = $false
        }
    }
	
    Function Connect-ExchOnPrem
    {
        param(
            [Parameter( Mandatory = $false)]
            [string]$URL = $txtTab0ConnectionURI.text
        )
        Try
        {
            # for Powershell v3+  - $Credentials = Get-Credential -Message "Enter your Exchange admin credentials"
            # for Powershell v2+  -
            $Credentials = Get-Credential
            $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos -Credential $Credentials -ErrorAction Stop
            Import-PSSession $ExOPSession
        }
        Catch
        {
            Log $Err004 Red
            $MsgBoxError::Show($Err004, $Str000, "OK", "Error")
        }
    }
	
    Function Connect-ExchOnline
    {
        try
        {
            $ExchOnlineCred = Get-Credential -ErrorAction Continue
            #save previous session - we'll remove it if login successful with new session
            $OldSessions = Get-PSSession
            #Create remote Powershell session with Exchange Online
            $ExchOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $ExchOnlineCred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
            #Import the remote PowerShell session
            Import-PSSession $ExchOnlineSession -AllowClobber | Out-Null
            Log "Ok, you're in !" Green
            #Remove previous sessions (and keep the current one)
            $OldSessions | Remove-PSSession
        }
        catch
        {
            Log $Err003 Red
            $MsgBoxError::Show($Err003, $Str000, "OK", "Error")
        }
		
    }
	
    #endregion  Application Functions using variables from the form
    #----------------------------------------------
	
    #----------------------------------------------
    #region Events handling (don't forget to add the event on the controls definition Control_Name.add_Click($control_name_Click)
    #----------------------------------------------
    $frmSearchForm_Load = {
        #Action ->
        $btnDelSearch.Enabled = $false
        $txtSender.Enabled = $false
        $chkSender.Checked = $false
        $chkOnPrem.Checked = $false
        $txtTab0ConnectionURI.Enabled = $chkOnPrem.Checked
        #$chkUseNewMailboxSearch.Checked = $false
        $txtAttachment.Enabled = $false
        $chkAttachment.Checked = $false
        #$chkEstimateOnly.Checked = $True
        #$chkUseNewMailboxSearch.Checked = $true
        Log $Str000 DarkGreen
        $txtExchangeDefaultConnectionURI = "email.contoso.ca"
        Update-CmdLine
        $StatusBar.Text = $Str004a
        if ($TestMode -eq $true) { Test-ExchCxTest } Else {	Test-ExchCx }
        $StatusBar.Text = $Str004
    } 

    $btnTab0ConnectExch_Click = {
        $StatusBar.Text = $Str004a
        $lblTab0CxStatusUpdate.Text = $Str007d
        $lblTab0CxStatusUpdate.Forecolor = 'Red'
        IF ($chkOnPrem.Checked)
        {
            Connect-ExchOnPrem $txtTab0ConnectionURI.text
            if ($TestMode -eq $true) { Test-ExchCxTest } Else {	Test-ExchCx }
        }
        Else
        {
            Connect-ExchOnline
            if ($TestMode -eq $true) { Test-ExchCxTest } Else {	Test-ExchCx }
        }
        $StatusBar.Text = $Str004
    }
	
    $lblabout_Click = {
        switch ($Language)
        {
            "EN"
            {
                $systemst = "QXV0aG9yOiBTYW0gRHJleQ0Kc2FtZHJleUBtaWNyb3NvZnQuY29tDQpzYW1teUBob3RtYWlsLmZyDQpNaWNyb3NvZnQgRW`
			5naW5lZXIgc2luY2UgT2N0IDE5OTkNCjE5OTktMjAwMDogUHJlc2FsZXMgRW5naW5lZXIgKEZyYW5jZSkNCjIwMDAtMjAwMzogU3VwcG9yd`
			CBFbmdpbmVlciAoRnJhbmNlKQ0KMjAwMy0yMDA2OiB2ZXJ5IGZpcnN0IFBGRSBpbiBGcmFuY2UNCjIwMDYtMjAwOTogTUNTIENvbnN1bHRhb`
			nQgKEZyYW5jZSkNCjIwMDktMjAxMDogVEFNIChGcmFuY2UpDQoyMDEwLW5vdyA6IENvbnN1bHRhbnQgKENhbmFkYSkNCk11c2ljaWFuLCBjb`
			21wb3NlciAoS2V5Ym9hcmQsIEd1aXRhcikNClBsYW5lIHBpbG90IHNpbmNlIDE5OTUNCkZvciBTaGFyZWQgU2VydmljZXMgQ2FuYWRh"
            } 
            "FR"
            {
                $systemst = "QXV0ZXVyOiBTYW0gRHJleQ0Kc2FtZHJleUBtaWNyb3NvZnQuY29tDQpzYW1teUBob3RtYWlsLmZyDQpJbmfDqW5pZXVyIGNo`
			ZXogTWljcm9zb2Z0IGRlcHVpcyBPY3QgMTk5OQ0KMTk5OS0yMDAwOiBJbmfDqW5pZXVyIEF2YW50LVZlbnRlIChGcmFuY2UpDQoyMDAwLTIwMD`
			M6IFNww6ljaWFsaXN0ZSBUZWNobmlxdWUgKEZyYW5jZSkNCjIwMDMtMjAwNjogUHJlbWllciBQRkUgZW4gRnJhbmNlDQoyMDA2LTIwMDk6IENv`
			bnN1bHRhbnQgTUNTIChGcmFuY2UpDQoyMDA5LTIwMTA6IFJlc3BvbnNhYmxlIFRlY2huaXF1ZSBkZSBDb21wdGUgKEZyYW5jZSkNCjIwMTAtMjA`
			xNiA6IENvbnN1bHRhbnQgKENhbmFkYSkNCk11c2ljaWVuLCBjb21wb3NpdGV1ciAoQ2xhdmllciwgR3VpdGFyZSkNCkJyZXZldCBkZSBQaWxvdGU`
			gUHJpdsOpIGRlcHVpcyAxOTk1DQpQb3VyIFNlcnZpY2VzIFBhcnRhZ8OpcyBDYW5hZGE="
            }
        }
        $systemst = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($systemst))
        $MsgBoxError::Show($systemst, $Str000, "Ok", "Information")
    }
	
    $chkOnPrem_CheckedChanged = {
        if ($chkOnPrem.Checked) {$lblSwitchOnPremOnCloud.Text = $Str034} Else {$lblSwitchOnPremOnCloud.Text = $Str035}
        $txtTab0ConnectionURI.Enabled = $chkOnPrem.Checked
    }

    $handler_btnRun_Click = {
        #cls
		
        if (($txtRecipient.Text -eq "") -or ($txtRecipient.Text -eq $null))
        {
            $ChoiceBlankMailContinueYN = $MsgBoxError::Show($Str036, $Str000, "YesNo", "Warning")
        }
	
        If ($ChoiceBlankMailContinueYN -eq "No")
        {
            Write-Host "Returning back to GUI ..." 
        }
        Else
        {
            $StatusBar.Text = $Str004a
            #Action ->
            #Extracting Main User data into variables
            #		$TargetEmail = $txtRecipient.Text
            $DiscoveryMailbox = $txtDiscoveryMailbox.Text
            $DiscoveryFolder = $txtDiscoveryMailboxFolder.Text
            #		$SenderEmail = $txtSender.Text
            #		$Keyword = ('''') + $txtKeyword.Text + ('''')
            #		$StartDate = $txtStartDate.Text
            #		$EndDate = $txtEndDate.Text
            #Using Date interval Powershell notation for Search Query instead of received :> and received :<
            #		$DateInt = $StartDate + ".." + $EndDate
            $Command = $richTxtCurrentCmdlet.Text

            #First things first : checking if the Discovery Mailbox exists if it doesn't, then display error message to get an existing one, bypass all the rest of the code below, and return to GUI
            $mailboxTest = $null 			# Initialize $mailboxTest variable (used to test if a mailbox exists)
            $DiscoverMailboxTest = $true 	# Initialize $DiscoverMailboxTest variable (used as a boolean to say if the mailbox exists or not)
            If ($TestMode -eq $True) {Write-Host "Test mode, doing Get-Mailbox $DiscoveryMailbox" -BackgroundColor (Get-Random("Blue", "White", "Red")); $MailboxTest = "Dummy Mailbox"} 
            Else {$mailboxTest = Get-Mailbox $DiscoveryMailbox -ErrorAction SilentlyContinue}
            if (($mailboxTest -eq $null) -or ($mailboxTest -eq ""))
            {
                $DiscoverMailboxTest = $false
            }


            If ($chkUseNewMailboxSearch.Checked -eq $false)
            {
                $MustDoDiscoveryMailboxTestSearchMailbox = $false
                If ($chkDeleteMail.Checked -eq $true)
                {
                    $MustDoDiscoveryMailboxTestSearchMailbox = $true
                }
                Else
                {
                    If ($chkEstimateOnly.Checked -eq $true)
                    {
                        Log $Command Blue
                        If ($TestMode) {Write-Host $Command} Else {$output = invoke-expression $Command | Out-String}
                        Log $output | Select Identity, TargetMailbox, Success, ResultItemscount, ResultItemsSize | Ft
                        # Update status bar
                        $StatusBar.Text = $Str004
                    }
                    Else
                    {
                        # $ChkEstimateOnly.checked -eq $false
                        $MustDoDiscoveryMailboxTestSearchMailbox = $true
                    }	
                }
                If ($MustDoDiscoveryMailboxTestSearchMailbox)
                {
                    Log "Does $DiscoveryMailbox exist ? -->", $DiscoverMailboxTest Green, white
                    If ($DiscoverMailboxTest)
                    {
                        Log "$DiscoveryMailbox exists" Green
                        Log $Command
                        If ($TestMode) {Write-Host $Command} Else {$output = invoke-expression $Command | Out-String}
                        Log $output | Select Identity, TargetMailbox, Success, ResultItemscount, ResultItemsSize | Ft
                        # Update status bar
                        $StatusBar.Text = $Str004
                        #$frmSearchForm.Close()
                    }
                    Else
                    {
                        $MsgBoxError::Show($Err005, $Str000, "Ok", "Error")
                    }
                }
            }
            Else
            {
                #chkUserNewMailboxSearch -eq $true
                Log "Does $DiscoveryMailbox exist ? -->", $DiscoverMailboxTest, "it is", $DiscoveryFolder Green, white
                If ($DiscoverMailboxTest)
                {
                    Log "$DiscoveryMailbox exists, now because we're using New-MailboxSearch we have to check if ""$DiscoveryFolder"" already exists because New-MailboxSearch cannot use an already existing Search name / folder..." Green
                    #Clearing errors first
                    $Error.Clear()
                    #Trying Get-MailboxSearch with name of Discovery Folder...
                    If ($TestMode) {Write-Host "$DiscoveryFolder test => Get-MailboxSearch '$DiscoveryFolder'"} Else {invoke-expression "Get-MailboxSearch '$DiscoveryFolder'"}
                    #If there is no error, that means the folder already exist, close form and exit script...
                    If (($Error.Count -eq 0) -and ($TestMode -eq $false))
                    {
                        Log "A Mailbox Search with the same name already exists. Please type another one in the 'Folder' field of the GUI..." DarkRed
                        Log "or remove the existing mailbox search by typing the following command or using the GUI:" DarkRed
                        Log "Remove-MailboxSearch ""$DiscoveryFolder""" DarkGreen
                        Log "and run the script again..." DarkRed
                        # Message box for displaying error in the GUI
                        $MsgBoxError::Show($Err002, $Str000, "OK", "Error")
                        # Update status bar
                        $StatusBar.Text = $Str004
                        #$frmSearchForm.Close()
                    }
                    Else
                    {
                        # $Error.Count > 0 => Get-MailboxSearch $DiscoveryFolder returned "Not found" => $DiscoveryFolder doesn't exist
                        Log "WARNING: THE ABOVE MESSAGE IS NORMAL, and is just here to confirm the Search Name does not already exist- PLEASE IGNORE IT" Yellow
                        Log "We can continue, the folder does not exist" DarkYellow
                        #Write-host $Error -BackgroundColor Blue -ForegroundColor Yellow
                        Log "Launching" Green
                        Log $Command DarkRed
                        Log "`nplease wait..." Green
                        $Error.Clear()
					
                        If ($TestMode) {Write-Host $Command} Else {$output = Invoke-Expression $Command}
                        If ($Error.Count -eq 0) 
                        {
                            Log "Wait while we start your Search request ..." DarkYellow
                            If ($TestMode) {Write-Host "Start-MailboxSearch $DiscoveryFolder"} Else {Start-MailboxSearch $DiscoveryFolder}
                            Write-Host "Waiting 10 seconds to let time to Exchange to build some preliminary stats about your search..."
                            for ($i = 10; $i -gt 0; $i--) {sleep 1; Write-Host $i}
                            If ($TestMode)
                            {
                                Write-Host "Get-MailboxSearch $DiscoveryFolder"; $output = $null		
                                # Update status bar
                                $StatusBar.Text = $Str004
                            } 
                            Else
                            {
                                $output = Get-MailboxSearch $DiscoveryFolder
                                Log "Name of the mailbox search:.......................", $output.Name
                                Log "Success status:...................................", $output.Status
                                Log "Percent Complete:.................................", $output.PercentComplete
                                Log "Number of mailboxes to search:....................", $output.NumberMailboxesToSearch
                                Log "Estimated number of results:......................", $output.REsultNumberEstimate
                                Log "Estimated total size of results:..................", $output.ResultSizeEstimate
                                Log "Link to view the results (OWA):...................", $output.ResultsLink
                                Log "`nPlease run the following command after a few more seconds to view the progress:" DarkYellow 
                                Log "Get-MailboxSearch ""$DiscoveryFolder"" | fl Name,Status,PercentComplete,ResultNumberEstimate" DarkRed
                                Log "after a few more seconds to view the progress" DarkYellow
                                Log "Or you can just use the Graphical Interface as well (See other Tabs of the GUI) that does the above using just one click" DarkYellow
                                Log "and run the following command:" DarkYellow
                                Log "Get-MailboxSearch ""$DiscoveryFolder"" | fl ResultsLink" DarkRed
                                Log "to view and copy/paste the URL to access the Discovery Mailbox directly..." DarkYellow
                                Log "Or you can also use the Graphical Interface (See other Tabs of the GUI) that gives the URL results as a clickable URL..." DarkYellow
                                # Update status bar
                                $StatusBar.Text = $Str004
                                #$frmSearchForm.Close()
                            }
                        }
                        Else 
                        {
                            Log $Err001 Red
                            $MsgBoxError::Show($Err001, $Str000, "OK", "Error")
                            # Update status bar
                            $StatusBar.Text = $Str004
                        }
                    } 						
                }
                Else
                {
                    # $DiscoveryMailboxTest -eq $false (Discovery Mailbox Does not exist)
                    $MsgBoxError::Show($Err005, $Str000, "Ok", "Error")			
                }
            }
        } # (closing when user says "Cancel" if no e-mail address specified and don't want to search ALL mailboxes)
        Log "Returning back to Graphical User interface..." Green
    } # (closing the handle)
	
    $btnCancel_Click = {
        #Action ->
        $frmSearchForm.Close()
    }
	
    $chkSender_CheckedChanged = {
        #Action ->
        $txtSender.Enabled = $chkSender.Checked
        Update-CmdLine	
    }
	
    $chkAttachment_CheckedChanged = {
        #Action ->
        $txtAttachment.Enabled = $chkAttachment.Checked
        Update-CmdLine
    }
	
    $chkUserNewMailboxSearch_CheckedChanged = {
        #Action ->
        $chkDeleteMail.Checked = $false
        if ($global:RightToDelete)
        {
            $chkDeleteMail.Enabled = (-not($chkUseNewMailboxSearch.Checked)) -and (-not($chkEstimateOnly.Checked))
        }
        Else
        {
            $chkDeleteMail.Enabled = $False
        }
        Update-CmdLine
    }
		
    $chkEstimateOnly_CheckedChanged = {
        $chkDeleteMail.Checked = $false
        if ($global:RightToDelete)
        {
            $chkDeleteMail.Enabled = (-not($chkUseNewMailboxSearch.Checked)) -and (-not($chkEstimateOnly.Checked))
        } 
        Else {$chkDeleteMail.Enabled = $False}
        Update-CmdLine
    }
		
    $txtSender_TextChanged = {
        Update-CmdLine
    }
	
    $txtAttachment_TextChanged = {
        Update-CmdLine
    }
	
    $txtKeyword_TextChanged = {
        Update-CmdLine
    }
	
    $chkSubject_CheckedChanged = {
        Update-CmdLine
    }
		
    $txtStartDate_TextChanged = {
        Update-CmdLine
    }
	
    $txtEndDate_TextChanged = {
        Update-CmdLine
    }
	
    $chkDeleteMail_CheckedChanged = {
        Update-CmdLine
    }
	
    $txtDiscoveryMailbox_TextChanged = {
        Update-CmdLine
    }
	
    $txtDiscoveryMailboxFolder_TextChanged = {
        Update-CmdLine
    }
	
    $btnTab2GetMbxSearches_Click = {
        $StatusBar.Text = $Str004a
        $richtxtGetMbxSearch.Text = ""
        $richtxtGetMbxSearch.SelectionColor = 'Black'
        $richtxtGetMbxSearch.SelectedText = $Txt002b
        Log $Txt002b
        $MailboxSearchResults = Get-MailboxSearch | Select @{Label = $TxtLbl001; Expression = {$_.Name}}, @{Label = $TxtLbl002; Expression = {$_.percentComplete}}, @{Label = $TxtLbl003; Expression = {$_.ResultNumberEstimate}} | ft -AutoSize | Out-String
        $richtxtGetMbxSearch.Text = $MailboxSearchResults
        Log "Got all search results"
        $StatusBar.Text = $Str004
    }

    $btnTab2Get1MbxSearch_Click = {
        $StatusBar.Text = $Str004a
        $DiscoveryFolder = $txtDiscoveryMailboxFolder.Text
        $richtxtGetMbxSearch.Text = ""
        $richtxtGetMbxSearch.SelectionColor = 'black'
        $richtxtGetMbxSearch.SelectedText = $Txt002 + (" ""$DiscoveryFolder"" ") + $Txt002a
        Log $Txt002, $DiscoveryFolder, $Txt002a
        $MailboxSearchResults = Get-MailboxSearch $DiscoveryFolder -ErrorAction SilentlyContinue | Select @{Label = $TxtLbl001; Expression = {$_.Name}}, @{Label = $TxtLbl002; Expression = {$_.percentComplete}}, @{Label = $TxtLbl003; Expression = {$_.ResultNumberEstimate}} | ft -AutoSize | Out-String
        if ($MailboxSearchResults -eq "")
        {
            $richtxtGetMbxSearch.Text = ""
            $richtxtGetMbxSearch.SelectionColor = 'Red'
            $richtxtGetMbxSearch.SelectedText = $Txt003 + (" ""$($DiscoveryFolder)"" ") + $Txt003a
            Log $Txt003, $DiscoveryFolder, $Txt003a
        }
        Else
        {
            $richtxtGetMbxSearch.Text = $MailboxSearchResults
            Log "Got 1 search result ..."
        }
        $StatusBar.Text = $Str004
    }

    $btnTab4PopList_Click = {
        Populate-DropDownList
    }
	
    $comboTab4MbxSearches_TextChanged = {
        $StatusBar.Text = $Str004a
        $txtTab4MbxSearchStats.AppendText("`n`n" + $Txt009)
        $comboselectedItem = $comboTab4MbxSearches.selectedItem
        $GetMailboxSearchTab4 = Get-MailboxSearch $comboselectedItem
        $txtTab4MbxSearchStats.Text = $Txt010 + ("""$($GetMailboxSearchTab4.Name)"" `n")
        $txtTab4MbxSearchStats.Text += $GetMailboxSearchTab4 | Select LastRunBy, SourceMailboxes, SearchQuery, StartDate, EndDate, TargetMailbox, Status, LastStartTime, NumberMailboxesToSearch, PercentComplete, ResultNumber, ResultNumberEstimate, ResultSize, ResultSizeEstimate, ResultSizeCopied, Errors | Out-String
        $lblTab4URL.Text = $Str029a + ("""$($GetMailboxSearchTab4.TargetMailbox)""") + $Str029b + (" ""$comboselectedItem"":")
        $resultsLink = $GetMailboxSearchTab4.ResultsLink
        $lblTab4URLLink.text = $resultsLink
        If (($ResultsLink -eq $null) -or ($ResultsLink -eq ""))
        {
            $lblTab4URLLink.Text = $Txt012
            $lblTab4URLLink.Enabled = $false	
        }
        Else
        {
            try
            {
                $lblTab4URLLink.add_Click( {[system.Diagnostics.Process]::start($lblTab4URLLink.text)}) 
                $lblTab4URLLink.Enabled = $true
            }
            catch
            {
                Log $Txt013
            }
        }
        $btnDelSearch.Enabled = $true
        $StatusBar.Text = $Str004
    }
	
    $btnDelSearch_OnClick = {
        $SelectedToDelete = $comboTab4MbxSearches.selectedItem
        $ClickResult = $MsgBoxError::Show($Str031 + (" `n") + ($comboTab4MbxSearches.selectedItem), $Str000, [System.Windows.Forms.MessageBoxButtons]::YesNo , "Warning")
        Switch ($ClickResult)
        {
            "Yes"
            {
                $StatusBar.Text = $Str004a
                $lblTab4URLLink.Enabled = $false
                $txtTab4MbxSearchStats.Text = $Txt015
                Log $Txt015
                Remove-MailboxSearch $comboTab4MbxSearches.selectedItem -Confirm:$false
                Populate-DropDownList
                $txtTab4MbxSearchStats.Text += ("`n`n") + $SelectedToDelete + " " + $Str033
            }
            "No"
            {
                Log $SelectedToDelete, " ", $Str032
                $MsgBoxError::Show($SelectedToDelete + " " + $Str032, $Str000, [System.Windows.Forms.MessageBoxButtons]::Ok, "Information")
			
            }
        }
    }
	
    $Form_StateCorrection_Load = {
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $frmSearchForm.WindowState = $InitialFormWindowState
    }

    $Form_Cleanup_FormClosed = {
        #Remove all event handlers from the controls
        $MsgBoxError::Show($Txt014, $Str000, "OK", "Information")
        try
        {
            $StatusBar.Text = $Str004b
            #Remove-PSSession $ExOPSession
            #Get-PSSession | Remove-PSSession
            Log $Txt014
            $chkAttachment.remove_CheckedChanged($chkAttachment_CheckedChanged)
            $btnCancel.remove_Click($btnCancel_Click)
            $btnRun.remove_Click($handler_btnRun_Click)
            $txtEndDate.remove_TextChanged($handler_textBox6_TextChanged)
            $lblStartDate.remove_Click($handler_label3_Click)
            $chkSubject.remove_CheckedChanged($handler_checkBox3_CheckedChanged)
            $lblKeyword.remove_Click($handler_lblSearchKeyword_Click)
            $txtKeyword.remove_TextChanged($handler_txtSearchKeyword_TextChanged)
            $txtSender.remove_TextChanged($handler_txtSender_TextChanged)
            $chkSender.remove_CheckedChanged($chkSender_CheckedChanged)
            $lblRecipient.remove_Click($handler_lblRecipient_Click)
            $frmSearchForm.remove_Load($frmSearchForm_Load)
            $frmSearchForm.remove_Load($Form_StateCorrection_Load)
            $frmSearchForm.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        catch [Exception]
        { }
    }
	
    #endregion  Events handling (don't forget to add the event on the controls definition Control_Name.add_Click($control_name_Click)
    #----------------------------------------------

    #----------------------------------------------
    #region Suspending layouts - resume layout needed later (this technique is used to avoid form flickering while adding the controls)
    #----------------------------------------------
    $frmSearchForm.SuspendLayout()
    #tabbing - step 2 - Suspend Layouts ...
    $tabcontrol.SuspendLayout()
    $tabPage0.SuspendLayout()
    $tabPage1.SuspendLayout()
    $tabPage2.SuspendLayout()
    $tabPage3.SuspendLayout()
    #----------------------------------------------
    #endregion Suspending layouts
    #----------------------------------------------

    #----------------------------------------------
    #region Setting and configuring each Tab, adding each controls (Controls=buttons, labels, texts, etc... and are configured later)
    #----------------------------------------------
    # frmSearchForm
    #
    #region Main form
    #tabbing - step 3 - Add Tab Control to the form...
    $frmSearchForm.Controls.Add($tabcontrol)
    $frmSearchForm.Controls.Add($StatusBar)
    #		#Testing AutoScale vs Autosize ...
    #		$frmSearchForm.AutoScaleMode = 2
    #		$frmSearchForm.AutoSizeMode = 0
    $frmSearchForm.StartPosition = "CenterScreen"
    $frmSearchForm.ClientSize = '630, 650'
    $frmSearchForm.Name = 'frmSearchForm'
    if ($TestMode) {$frmSearchForm.Text = $Str000 + ("  !!! TEST MODE !!!")} Else {$frmSearchForm.Text = $Str000}
    $frmSearchForm.FormBorderStyle = 'FixedDialog'
    $frmSearchForm.add_Load($frmSearchForm_Load)
    $frmSearchForm.MaximizeBox = $false
    #endregion main Form

    #region StatusBar
    $StatusBar.Name = 'StatusBar'
    $StatusBar.DataBindings.DefaultDataSourceUpdateMode = 0
    $StatusBar.TabIndex = 0
    $StatusBar.Location = '0,553'
    $StatusBar.Size = '630,22'
    # Update status bar
    $StatusBar.Text = $Str004
    #endregion Statusbar

    #region TabControl
    #tabbing - step 4 - Add Tabs to the Tab Control...
    $tabcontrol.Controls.Add($tabPage0)
    $tabcontrol.Controls.Add($tabPage1)
    $tabcontrol.Controls.Add($tabPage2)
    $tabcontrol.Controls.Add($tabPage3)
    $tabcontrol.Location = '2, 2'
    $tabcontrol.Name = 'tabControl'
    $tabcontrol.SelectedIndex = 0
    $tabcontrol.Size = '629, 620'
    $tabcontrol.TabIndex = 0
    #endregion TabControl

    #region TabPage0
    #tabbing - step 5 - Add traditional form controls to the Tab Page instead of to the form...
    $tabPage0.Controls.Add($lblBigTitle1)
    $tabPage0.Controls.Add($lblabout)
    $tabPage0.Controls.Add($lblBigTitle2)
    $tabPage0.Controls.Add($richtxtWelcome)
    $tabPage0.Controls.Add($btnTab0ConnectExch)
    $tabPage0.Controls.Add($lblTab0CxStatus)
    $tabPage0.Controls.Add($lblTab0CxStatusUpdate)
    $tabPage0.Controls.Add($txtTab0ConnectionURI)
    $tabPage0.Controls.Add($lblSwitchOnPremOnCloud)
    $tabPage0.Controls.Add($chkOnPrem)
    $tabpage0.Location = '4, 22'
    $tabpage0.Name = 'tabWelcome'
    $tabpage0.Padding = '1, 1, 1, 1'
    $tabpage0.Size = '621, 524'
    $tabpage0.TabIndex = 1
    $tabpage0.Text = $Str001
    $tabpage0.UseVisualStyleBackColor = $True
    #endregion TabPage0
	
    #region TabPage1
    #tabbing - step 5 - Add traditional form controls to the Tab Page instead of to the form...
    $tabPage1.Controls.Add($txtAttachment)
    $tabPage1.Controls.Add($chkAttachment)
    $tabPage1.Controls.Add($btnCancel)
    $tabPage1.Controls.Add($btnRun)
    $tabPage1.Controls.Add($txtDiscoveryMailbox)
    $tabPage1.Controls.Add($lblDiscoveryMailbox)
    $tabPage1.Controls.Add($lblDiscoveryMailboxFolder)
    $tabPage1.Controls.Add($txtDiscoveryMailboxFolder)
    $tabPage1.Controls.Add($chkUseNewMailboxSearch)
    $tabPage1.Controls.Add($chkDeleteMail)
    $tabPage1.Controls.Add($txtEndDate)
    $tabPage1.Controls.Add($lblEndDate)
    $tabPage1.Controls.Add($txtStartDate)
    $tabPage1.Controls.Add($lblStartDate)
    $tabPage1.Controls.Add($chkSubject)
    $tabPage1.Controls.Add($lblKeyword)
    $tabPage1.Controls.Add($txtKeyword)
    $tabPage1.Controls.Add($txtSender)
    $tabPage1.Controls.Add($chkSender)
    $tabPage1.Controls.Add($txtRecipient)
    $tabPage1.Controls.Add($lblRecipient)
    $tabPage1.Controls.Add($richTxtCurrentCmdlet)
    $tabPage1.Controls.Add($chkEstimateOnly)
    $tabpage1.Location = '4, 22'
    $tabpage1.Name = 'tabSearchMailbox'
    $tabpage1.Padding = '3, 3, 3, 3'
    $tabpage1.Size = '621, 524'
    $tabpage1.TabIndex = 1
    $tabpage1.Text = $Str008
    $tabpage1.UseVisualStyleBackColor = $True
    #endregion TabPage1

    #region tabPage2
    $tabPage2.Controls.Add($lblTab2ExistingSearches)
    $tabPage2.Controls.Add($btnTab2Get1MbxSearch)
    $tabPage2.Controls.Add($btnTab2GetMbxSearches)
    $tabPage2.Controls.Add($richtxtGetMbxSearch)
    $tabpage2.Location = '4, 22'
    $tabpage2.Name = 'tabGetExistingRequests'
    $tabpage2.Padding = '3, 3, 3, 3'
    $tabpage2.Size = '621, 524'
    $tabpage2.TabIndex = 2
    $tabpage2.Text = $Str023
    $tabpage2.UseVisualStyleBackColor = $True
    #endregion tabPage2
	
    #region tabPage3
    $tabPage3.Controls.Add($lblTab4URL)
    $tabPage3.Controls.Add($lblTab4URLLink)
    $tabPage3.Controls.Add($txtTab4MbxSearchStats)
    $tabPage3.Controls.Add($comboTab4MbxSearches)
    $tabPage3.Controls.Add($btnTab4PopList)
    $tabPage3.Controls.Add($btnDelSearch)
    $tabPage3.UseVisualStyleBackColor = $True
    $tabPage3.Text = $Str027
    $tabPage3.DataBindings.DefaultDataSourceUpdateMode = 0
    $tabPage3.TabIndex = 3
    $tabPage3.Name = 'tabPage3'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 621
    $System_Drawing_Size.Height = 524
    $tabPage3.Size = $System_Drawing_Size
    $System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
    $System_Windows_Forms_Padding.All = 3
    $System_Windows_Forms_Padding.Bottom = 3
    $System_Windows_Forms_Padding.Left = 3
    $System_Windows_Forms_Padding.Right = 3
    $System_Windows_Forms_Padding.Top = 3
    $tabPage3.Padding = $System_Windows_Forms_Padding
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 4
    $System_Drawing_Point.Y = 22
    $tabPage3.Location = $System_Drawing_Point
    #endregion tabPage3

    #endregion Setting and configuring each Tab, adding each controls (Controls are configured later)

    #region							TAB 0 (=Welcome! tab) ELEMENTS DEFINITION   		 			#
    #################################################################################################
    #																								#
    # 							TAB - TAB 0 (=Welcome! tab) ELEMENTS DEFINITION						#
    #																								#
    #################################################################################################
    #
    # lblBigTitle1
    #
    $lblBigTitle1.Text = $Str002
    $lblBigTitle1.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblBigTitle1.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 0)
    $lblBigTitle1.TabIndex = 1
    $lblBigTitle1.TextAlign = 32
    $lblBigTitle1.Name = 'lblBigTitle1'
    $lblBigTitle1.Size = '535,52'
    $lblBigTitle1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 25, 1, 3, 1)
    $lblBigTitle1.Location = '35,20'
    #
    # lblSystem
    #
    $lblabout.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblabout.Name = 'lblabout'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 10
    $System_Drawing_Size.Height = 10
    $lblabout.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 611
    $System_Drawing_Point.Y = 0
    $lblabout.Location = $System_Drawing_Point
    $lblabout.add_click($lblabout_Click)
    #
    #lblBigTitle2
    #
    $lblBigTitle2.Text = $Str003
    $lblBigTitle2.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblBigTitle2.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 255)
    $lblBigTitle2.TabIndex = 2
    $lblBigTitle2.TextAlign = 32
    $lblBigTitle2.Name = 'lblBigTitle2'
    $lblBigTitle2.Size = '457,35'
    $lblBigTitle2.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12, 0, 3, 0)
    $lblBigTitle2.Location = '79,72'
    #
    # richTxtWelcome
    #
    $richtxtWelcome.Text = $Txt001
    $richtxtWelcome.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 192, 192)
    $richtxtWelcome.TabIndex = 3
    $richtxtWelcome.Name = 'richtxtWelcome'
    $richtxtWelcome.Font = New-Object System.Drawing.Font("Lucida console", 8, 0, 3, 0)
    $richtxtWelcome.Location = '9,125'
    $richtxtWelcome.Size = '606,328'
    $richtxtWelcome.DataBindings.DefaultDataSourceUpdateMode = 0
    $richtxtWelcome.add_TextChanged($richtxtWelcome_TextChanged)
    #
    # Button Launch Exchange Connection
    #
    $btnTab0ConnectExch.UseVisualStyleBackColor = $True
    $btnTab0ConnectExch.Text = $Str005
    $btnTab0ConnectExch.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnTab0ConnectExch.TabIndex = 4
    $btnTab0ConnectExch.Name = 'btnTab0ConnectExch'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 163
    $System_Drawing_Size.Height = 23
    $btnTab0ConnectExch.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 35
    $System_Drawing_Point.Y = 482
    $btnTab0ConnectExch.Location = $System_Drawing_Point
    $btnTab0ConnectExch.add_Click($btnTab0ConnectExch_Click)
    #
    # $lblSwitchOnPremOnCloud
    #
    $lblSwitchOnPremOnCloud.Text = $Str035
    $lblSwitchOnPremOnCloud.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblSwitchOnPremOnCloud.TabIndex = 9
    $lblSwitchOnPremOnCloud.TextAlign = 2
    $lblSwitchOnPremOnCloud.Name = 'lblSwitchOnPremOnCloud'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 163
    $System_Drawing_Size.Height = 23
    $lblSwitchOnPremOnCloud.Size = $System_Drawing_Size
    $lblSwitchOnPremOnCloud.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, 2, 3, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 35
    $System_Drawing_Point.Y = 508
    $lblSwitchOnPremOnCloud.Location = $System_Drawing_Point
    #
    # $chkOnPrem
    #
    $chkOnPrem.UseVisualStyleBackColor = $True
    $chkOnPrem.Text = $Str034
    $chkOnPrem.DataBindings.DefaultDataSourceUpdateMode = 0
    $chkOnPrem.TabIndex = 8
    $chkOnPrem.Name = 'chkOnPrem'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 153
    $System_Drawing_Size.Height = 24
    $chkOnPrem.Size = $System_Drawing_Size
    $chkOnPrem.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, 2, 3, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 35
    $System_Drawing_Point.Y = 570
    $chkOnPrem.Location = $System_Drawing_Point
    $chkOnPrem.add_CheckedChanged($chkOnPrem_CheckedChanged)
	
    #
    # $lblTab0CxStatus
    #
    $lblTab0CxStatus.Text = $Str006
    $lblTab0CxStatus.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblTab0CxStatus.TabIndex = 5
    $lblTab0CxStatus.Name = 'lblTab0CxStatus'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 135
    $System_Drawing_Size.Height = 23
    $lblTab0CxStatus.Size = $System_Drawing_Size
    $lblTab0CxStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, 0, 3, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 222
    $System_Drawing_Point.Y = 504
    $lblTab0CxStatus.Location = $System_Drawing_Point
    #
    # $lblTab0CxStatusUpdate
    #
    $lblTab0CxStatusUpdate.Text = '--'
    $lblTab0CxStatusUpdate.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblTab0CxStatusUpdate.TabIndex = 4
    $lblTab0CxStatusUpdate.Name = 'lblTab0CxStatusUpdate'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 249
    $System_Drawing_Size.Height = 80
    $lblTab0CxStatusUpdate.Size = $System_Drawing_Size
    $lblTab0CxStatusUpdate.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, 0, 3, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 363
    $System_Drawing_Point.Y = 502
    $lblTab0CxStatusUpdate.Location = $System_Drawing_Point
    #
    # $txtTab0ConnectionURI
    #
    $txtTab0ConnectionURI.Text = "email.contoso.ca"
    $txtTab0ConnectionURI.Name = 'txtTab0ConnectionURI'
    $txtTab0ConnectionURI.TabIndex = 6
    $txtTab0ConnectionURI.Enabled = $False
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 163
    $System_Drawing_Size.Height = 20
    $txtTab0ConnectionURI.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 35
    $System_Drawing_Point.Y = 550
    $txtTab0ConnectionURI.Location = $System_Drawing_Point
    $txtTab0ConnectionURI.DataBindings.DefaultDataSourceUpdateMode = 0
    #
    #endregion TAB 0 (=Welcome! tab) ELEMENTS DEFINITION

    #region				TAB 1 (Search Mailbox tab) ELEMENTS DEFINITION			   		 			#
    #################################################################################################
    #																								#
    # 							TAB - TAB 1 ELEMENTS DEFINITION										#
    #																								#
    #################################################################################################
    #
    # txtAttachment
    #
    $txtAttachment.Location = '276, 172'
    $txtAttachment.Name = 'txtAttachment'
    $txtAttachment.Size = '202, 20'
    $txtAttachment.TabIndex = 19
    $txtAttachment.add_TextChanged($txtAttachment_TextChanged)
    #
    # chkAttachment
    #
    $chkAttachment.Font = 'Microsoft Sans Serif, 8pt'
    $chkAttachment.Location = '13, 170'
    $chkAttachment.Name = 'chkAttachment'
    $chkAttachment.Size = '257, 24'
    $chkAttachment.TabIndex = 18
    $chkAttachment.Text = $Str011
    $chkAttachment.UseVisualStyleBackColor = $True
    $chkAttachment.add_CheckedChanged($chkAttachment_CheckedChanged)
    #
    #$richTxtCurrentCmdlet
    #
    $richTxtCurrentCmdlet.Font = 'Microsoft Sans Serif, 8pt'
    $richTxtCurrentCmdlet.Location = '6,449'
    $richTxtCurrentCmdlet.Name = 'lblCurrentCmdlet'
    $richTxtCurrentCmdlet.Size = '609,93'
    $richTxtCurrentCmdlet.TabIndex = 17
    $richTxtCurrentCmdlet.Text = "COMMANDLET RUN WILL APPEAR HERE - IT CAN BE VERY LONG - COMMANDLET RUN WILL APPEAR HERE - IT CAN BE VERY LONG - COMMANDLET RUN WILL APPEAR HERE - IT CAN BE VERY LONG - COMMANDLET RUN WILL APPEAR HERE - IT CAN BE VERY LONG - "
    $richTxtCurrentCmdlet.ReadOnly = $true
    #
    # btnCancel
    #
    $btnCancel.Font = 'Microsoft Sans Serif, 8pt'
    $btnCancel.FlatAppearance.MouseDownBackColor = 'Red'
    $btnCancel.FlatAppearance.MouseOverBackColor = '255,0,0'
    $btnCancel.Location = '339, 556'
    $btnCancel.Name = 'btnCancel'
    $btnCancel.Size = '223, 23'
    $btnCancel.TabIndex = 17
    $btnCancel.Text = $Str022
    $btnCancel.UseVisualStyleBackColor = $true
    $btnCancel.add_Click($btnCancel_Click)
    #
    # btnRun
    #
    $btnRun.Font = 'Microsoft Sans Serif, 8pt'
    $btnRun.FlatAppearance.MouseDownBackColor = 'Red'
    $btnRun.FlatAppearance.MouseOverBackColor = '255,0,0'
    $btnRun.Location = '41, 556'
    $btnRun.Name = 'btnRun'
    $btnRun.Size = '223, 23'
    $btnRun.TabIndex = 16
    $btnRun.Text = $Str021
    $btnRun.UseVisualStyleBackColor = $true
    $btnRun.add_Click($handler_btnRun_Click)
    #
    # chkUseNewMailboxSearch
    #
    $chkUseNewMailboxSearch.Font = 'Microsoft Sans Serif, 8pt'
    $chkUseNewMailboxSearch.BackColor = 'cyan'
    $chkUseNewMailboxSearch.ForeColor = 'Black'
    $chkUseNewMailboxSearch.Location = '13, 347'
    $chkUseNewMailboxSearch.Name = 'chkDeleteMail'
    $chkUseNewMailboxSearch.Size = '600, 24'
    $chkUseNewMailboxSearch.TabIndex = 13
    $chkUseNewMailboxSearch.Text = $Str017
    $chkUseNewMailboxSearch.UseVisualStyleBackColor = $False
    $chkUseNewMailboxSearch.Checked = $true
    $chkUseNewMailboxSearch.add_CheckedChanged($chkUserNewMailboxSearch_CheckedChanged)
    #
    #$chkEstimateOnly
    #
    $chkEstimateOnly.UseVisualStyleBackColor = $True
    $chkEstimateOnly.Text = $Str018
    $chkEstimateOnly.DataBindings.DefaultDataSourceUpdateMode = 0
    $chkEstimateOnly.TabIndex = 14
    $chkEstimateOnly.Name = 'chkEstimateOnly'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 206
    $System_Drawing_Size.Height = 34
    $chkEstimateOnly.Size = $System_Drawing_Size
    $chkEstimateOnly.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, 0, 3, 0)
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 406
    $System_Drawing_Point.Y = 378
    $chkEstimateOnly.Location = $System_Drawing_Point
    $chkEstimateOnly.Checked = $true
    $chkEstimateOnly.Add_CheckedChanged($chkEstimateOnly_CheckedChanged)
    #
    # txtDiscoveryMailbox
    #
    $txtDiscoveryMailbox.Location = '181, 386'
    $txtDiscoveryMailbox.Name = 'txtDiscoveryMailbox'
    $txtDiscoveryMailbox.Size = '202, 20'
    $txtDiscoveryMailbox.TabIndex = 15
    $txtDiscoveryMailbox.Text = 'DiscoveryMailbox'
    $txtDiscoveryMailbox.add_TextChanged($txtDiscoveryMailbox_TextChanged)
    #
    # lblDiscoveryMailbox
    #
    $lblDiscoveryMailbox.Font = 'Microsoft Sans Serif, 8pt'
    $lblDiscoveryMailbox.Location = '13, 389'
    $lblDiscoveryMailbox.Name = 'lblDiscoveryMailbox'
    $lblDiscoveryMailbox.Size = '162, 23'
    $lblDiscoveryMailbox.TabIndex = 14
    $lblDiscoveryMailbox.Text = $Str019
    #
    # txtDiscoveryMailboxFolder
    #
    $txtDiscoveryMailboxFolder.Location = '181, 412'
    $txtDiscoveryMailboxFolder.Name = 'txtDiscoveryMailboxFolder'
    $txtDiscoveryMailboxFolder.Size = '202, 20'
    $txtDiscoveryMailboxFolder.TabIndex = 18
    $txtDiscoveryMailboxFolder.Text = 'Seach Number 01'
    $txtDiscoveryMailboxFolder.add_TextChanged($txtDiscoveryMailboxFolder_TextChanged)
    #
    # lblDiscoveryMailboxFolder
    #
    $lblDiscoveryMailboxFolder.Font = 'Microsoft Sans Serif, 8pt'
    $lblDiscoveryMailboxFolder.Location = '103, 415'
    $lblDiscoveryMailboxFolder.Name = 'lblDiscoveryMailboxFolder'
    $lblDiscoveryMailboxFolder.Size = '72, 20'
    $lblDiscoveryMailboxFolder.TabIndex = 17
    $lblDiscoveryMailboxFolder.Text = $Str020
	
    #
    # chkDeleteMail
    #
    $chkDeleteMail.Font = 'Microsoft Sans Serif, 8pt'
    $chkDeleteMail.BackColor = 'Yellow'
    $chkDeleteMail.ForeColor = 'Red'
    $chkDeleteMail.Location = '13, 317'
    $chkDeleteMail.Name = 'chkDeleteMail'
    $chkDeleteMail.Size = '600, 24'
    $chkDeleteMail.TabIndex = 13
    $chkDeleteMail.Text = $Str016
    $chkDeleteMail.UseVisualStyleBackColor = $False
    $chkDeleteMail.Enabled = $False
    $chkDeleteMail.add_CheckedChanged($chkDeleteMail_CheckedChanged)
    #
    # txtEndDate
    #
    $txtEndDate.Location = '276, 282'
    $txtEndDate.Name = 'txtEndDate'
    $txtEndDate.Size = '130, 20'
    $txtEndDate.TabIndex = 11
    $txtEndDate.Text = '01/01/2100'
    $txtEndDate.add_TextChanged($txtEndDate_TextChanged)
    #
    # lblEndDate
    #
    $lblEndDate.Font = 'Microsoft Sans Serif, 8pt'
    $lblEndDate.Location = '13, 285'
    $lblEndDate.Name = 'lblEndDate'
    $lblEndDate.Size = '257, 20'
    $lblEndDate.TabIndex = 10
    $lblEndDate.Text = $Str015
    #
    # txtStartDate
    #
    $txtStartDate.Location = '276, 252'
    $txtStartDate.Name = 'txtStartDate'
    $txtStartDate.Size = '130, 20'
    $txtStartDate.TabIndex = 9
    $txtStartDate.Text = '01/01/2000'
    $txtStartDate.add_TextChanged($txtStartDate_TextChanged)
    #
    # lblStartDate
    #
    $lblStartDate.Font = 'Microsoft Sans Serif, 8pt'
    $lblStartDate.Location = '13, 255'
    $lblStartDate.Name = 'lblStartDate'
    $lblStartDate.Size = '257, 20'
    $lblStartDate.TabIndex = 8
    $lblStartDate.Text = $Str014
    #
    # ADD-ON v1.5.5 - Time Start and Time End going with Date Start and Date End
    # TimeStart Box
    $txtStartTime = New-Object system.windows.Forms.TextBox
    $txtStartTime.Width = 100
    $txtStartTime.Height = 20
    $txtStartTime.location = new-object system.drawing.point(560, 251)
    $txtStartTime.Font = "Microsoft Sans Serif,10"

    # TimeEnd Box
    $txtEndTime = New-Object system.windows.Forms.TextBox
    $txtEndTime.Width = 100
    $txtEndTime.Height = 20
    $txtEndTime.location = new-object system.drawing.point(560, 280)
    $txtEndTime.Font = "Microsoft Sans Serif,10"
    $frmSearchForm.controls.Add($txtEndTime)
    # TimeStart Label
    $lblTimeStart = New-Object system.windows.Forms.Label
    $lblTimeStart.Text = "Time Start:"
    $lblTimeStart.AutoSize = $true
    $lblTimeStart.Width = 25
    $lblTimeStart.Height = 10
    $lblTimeStart.location = new-object system.drawing.point(457, 252)
    $lblTimeStart.Font = "Microsoft Sans Serif,10"
    $frmSearchForm.controls.Add($lblTimeStart)
    # TimeEnd Label
    $lblTimeEnd = New-Object system.windows.Forms.Label
    $lblTimeEnd.Text = "Time End:"
    $lblTimeEnd.AutoSize = $true
    $lblTimeEnd.Width = 25
    $lblTimeEnd.Height = 10
    $lblTimeEnd.location = new-object system.drawing.point(457, 281)
    $lblTimeEnd.Font = "Microsoft Sans Serif,10"
    $frmSearchForm.controls.Add($lblTimeEnd)
    #
    # chkSubject
    #
    $chkSubject.Font = 'Microsoft Sans Serif, 8pt'
    $chkSubject.Location = '13, 220'
    $chkSubject.Name = 'chkSubject'
    $chkSubject.Size = '257, 24'
    $chkSubject.TabIndex = 7
    $chkSubject.Text = $Str013
    $chkSubject.UseVisualStyleBackColor = $True
    $chkSubject.add_CheckedChanged($chkSubject_CheckedChanged)
    #
    # lblKeyword
    #
    $lblKeyword.Font = 'Microsoft Sans Serif, 8pt'
    $lblKeyword.Location = '13, 200'
    $lblKeyword.Name = 'lblKeyword'
    $lblKeyword.Size = '257, 20'
    $lblKeyword.TabIndex = 6
    $lblKeyword.Text = $Str012
    #
    # txtKeyword
    #
    $txtKeyword.Location = '276, 200'
    $txtKeyword.Name = 'txtKeyword'
    $txtKeyword.Size = '202, 20'
    $txtKeyword.TabIndex = 5
    $txtKeyword.Text = 'Keyword1 OR Keyword2 OR Keyword3'
    $txtKeyword.add_TextChanged($txtKeyword_TextChanged)
    #
    # txtSender
    #
    $txtSender.Location = '276, 142'
    $txtSender.Name = 'txtSender'
    $txtSender.Size = '202, 20'
    $txtSender.TabIndex = 4
    $txtSender.add_TextChanged($txtAttachment_TextChanged)
    #
    # chkSender
    #
    $chkSender.Font = 'Microsoft Sans Serif, 8pt'
    $chkSender.Location = '13, 140'
    $chkSender.Name = 'chkSender'
    $chkSender.Size = '257, 24'
    $chkSender.TabIndex = 3
    $chkSender.Text = $Str010
    $chkSender.UseVisualStyleBackColor = $True
    $chkSender.add_CheckedChanged($chkSender_CheckedChanged)
    #
    # txtRecipient
    #
    $txtRecipient.Location = '13, 50'
    $txtRecipient.Multiline = $True
    $txtRecipient.Name = 'txtRecipient'
    $txtRecipient.Size = '589, 66'
    $txtRecipient.TabIndex = 2
    $txtRecipient.Text = 'Sammy Krosoft;sam02.drey02;sam03.drey03'
    $txtRecipient.add_TextChanged($txtAttachment_TextChanged)
    #
    # lblRecipient
    #
    $lblRecipient.Font = 'Microsoft Sans Serif, 8pt'
    $lblRecipient.Location = '13, 23'
    $lblRecipient.Name = 'lblRecipient'
    $lblRecipient.Size = '611, 23'
    $lblRecipient.TabIndex = 0
    $lblRecipient.Text = $Str009
    #
    #
    #endregion	-	TAB 1 ELEMENTS DEFINITION
	
    #region				TAB 2 (Get-MailboxSearch Status) ELEMENTS DEFINITION			 			#
    #################################################################################################
    #																								#
    # 							TAB - TAB 2 ELEMENTS DEFINITION										#
    #																								#
    #################################################################################################
    #
    # lblTab2ExistingSearches
    #
    $lblTab2ExistingSearches.Font = 'Microsoft Sans Serif, 8pt'
    $lblTab2ExistingSearches.Location = '13, 23'
    $lblTab2ExistingSearches.Name = 'lblTab2ExistingSearches'
    $lblTab2ExistingSearches.Size = '560 ,50'
    $lblTab2ExistingSearches.TabIndex = 0
    $lblTab2ExistingSearches.Text = $Str024
    #
    # btnTab2Launch1GetMbxSearch
    $btnTab2Get1MbxSearch.UseVisualStyleBackColor = $True
    $btnTab2Get1MbxSearch.Text = $Str025
    $btnTab2Get1MbxSearch.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnTab2Get1MbxSearch.TabIndex = 3
    $btnTab2Get1MbxSearch.Name = 'btnTab2Get1MbxSearch'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 174
    $System_Drawing_Size.Height = 23
    $btnTab2Get1MbxSearch.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 64
    $System_Drawing_Point.Y = 76
    $btnTab2Get1MbxSearch.Location = $System_Drawing_Point
    $btnTab2Get1MbxSearch.add_Click($btnTab2Get1MbxSearch_Click)
    #
    # btnTab2LaunchGetMbxSearch
    $btnTab2GetMbxSearches.FlatAppearance.MouseDownBackColor = 'Gray'
    $btnTab2GetMbxSearches.FlatAppearance.MouseOverBackColor = '255, 128, 0'
    $btnTab2GetMbxSearches.Location = '353,76'
    $btnTab2GetMbxSearches.Name = 'btnTab2GetMbxSearch'
    $btnTab2GetMbxSearches.Size = '174,23'
    $btnTab2GetMbxSearches.TabIndex = 16
    $btnTab2GetMbxSearches.Text = $Str026
    $btnTab2GetMbxSearches.UseVisualStyleBackColor = $True
    $btnTab2GetMbxSearches.add_Click($btnTab2GetMbxSearches_Click)
    #
    # richtxtGetMbxSearch
    #
    $richtxtGetMbxSearch.Text = ''
    $richtxtGetMbxSearch.TabIndex = 2
    $richtxtGetMbxSearch.Name = 'richtxtGetMbxSearch'
    $richtxtGetMbxSearch.Font = New-Object System.Drawing.Font("Lucida Console", 9.75, 0, 3, 0)
    $richtxtGetMbxSearch.Size = '606,384'
    $richtxtGetMbxSearch.DataBindings.DefaultDataSourceUpdateMode = 0
    $richtxtGetMbxSearch.Location = '6,134'
    #
    #
    #endregion							TAB 2 ELEMENTS DEFINITION					   		 		#

    #region				TAB 3 (Get-MailboxSearch Details) ELEMENTS DEFINITION	   		 			#
    #################################################################################################
    #																								#
    # 							TAB - TAB 3 ELEMENTS DEFINITION										#
    #																								#
    #################################################################################################
    #
    # $lblTab4URL
    #
    $lblTab4URL.Text = $Str029
    $lblTab4URL.DataBindings.DefaultDataSourceUpdateMode = 0
    $lblTab4URL.TabIndex = 4
    $lblTab4URL.TextAlign = 32
    $lblTab4URL.Name = 'lblTab4URL'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 615
    $System_Drawing_Size.Height = 37
    $lblTab4URL.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 3
    $System_Drawing_Point.Y = 476
    $lblTab4URL.Location = $System_Drawing_Point
    #
    # $lblTab4URLLink
    #
    $lblTab4URLLink.TabIndex = 3
    $lblTab4URLLink.Text = $Str030
    $lblTab4URLLink.TabStop = $True
    $lblTab4URLLink.Name = 'lblTab4URLLink'
    $lblTab4URLLink.Enabled = $false
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 3
    $System_Drawing_Point.Y = 513
    $lblTab4URLLink.Location = $System_Drawing_Point
    $lblTab4URLLink.TextAlign = 2
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 615
    $System_Drawing_Size.Height = 78
    $lblTab4URLLink.Size = $System_Drawing_Size
    $lblTab4URLLink.DataBindings.DefaultDataSourceUpdateMode = 0
    #
    # $txtTab4MbxSearchStats
    #
    $txtTab4MbxSearchStats.Font = New-Object System.Drawing.Font("Lucida Console", 8.25, 0, 3, 0)
    $txtTab4MbxSearchStats.Text = $Txt005
    $txtTab4MbxSearchStats.TabIndex = 2
    $txtTab4MbxSearchStats.Name = 'txtTab4MbxSearchStats'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 567
    #Liamichou
    $System_Drawing_Size.Height = 309
    $txtTab4MbxSearchStats.Size = $System_Drawing_Size
    $txtTab4MbxSearchStats.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 26
    $System_Drawing_Point.Y = 150
    $txtTab4MbxSearchStats.Location = $System_Drawing_Point
    #
    # $comboTab4MbxSearches
    #
    $comboTab4MbxSearches.Name = 'comboTab4MbxSearches'
    $comboTab4MbxSearches.DropDownStyle = 'DropDownList'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 234
    $System_Drawing_Size.Height = 21
    $comboTab4MbxSearches.Size = $System_Drawing_Size
    $comboTab4MbxSearches.FormattingEnabled = $True
    $comboTab4MbxSearches.TabIndex = 1
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 76
    $comboTab4MbxSearches.Location = $System_Drawing_Point
    $comboTab4MbxSearches.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboTab4MbxSearches.add_TextChanged($comboTab4MbxSearches_TextChanged)
    #
    # $btnTab4PopList
    #
    $btnTab4PopList.UseVisualStyleBackColor = $True
    $btnTab4PopList.Text = $Str028
    $btnTab4PopList.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnTab4PopList.TabIndex = 0
    $btnTab4PopList.Name = 'btnTab4PopList'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 234
    $System_Drawing_Size.Height = 23
    $btnTab4PopList.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 172
    $System_Drawing_Point.Y = 21
    $btnTab4PopList.Location = $System_Drawing_Point
    $btnTab4PopList.add_Click($btnTab4PopList_Click)
    #
    # $btnDelSearch
    #
    $btnDelSearch.UseVisualStyleBackColor = $True
    $btnDelSearch.Text = 'Delete Search'
    $btnDelSearch.DataBindings.DefaultDataSourceUpdateMode = 0
    $btnDelSearch.TabIndex = 5
    $btnDelSearch.Name = 'btnDelSearch'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 109
    $System_Drawing_Size.Height = 23
    $btnDelSearch.Size = $System_Drawing_Size
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 484
    $System_Drawing_Point.Y = 76
    $btnDelSearch.Location = $System_Drawing_Point
    $btnDelSearch.add_Click($btnDelSearch_OnClick)	#
    #
    #endregion							TAB 3 ELEMENTS DEFINITION					   		 			#

    #region Resume Layout (this technique is used to avoid form flickering while adding the controls)
    $frmSearchForm.ResumeLayout()
    #tabbing - step 6 - Resume Layout for the Tab Control and for Tabs, just like the Form
    #Note: Tab Control, Tab PAges are like Forms under the Form...
    $tabcontrol.ResumeLayout()
    $tabPage0.ResumeLayout()
    $tabPage1.ResumeLayout()
    $tabPage2.ResumeLayout()
    $tabPage3.ResumeLayout()
    #endregion Resume layout

    #region Final form app operations (cleanup and launch)
    #Save the initial state of the form
    $InitialFormWindowState = $frmSearchForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $frmSearchForm.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $frmSearchForm.add_FormClosed($Form_Cleanup_FormClosed)
    #Show the Form
    $frmSearchForm.ShowDialog() | Out-Null
    #endregion Final form app operations (cleanup and launch)

}

#----------------------------------------------
#region Launcher and language selection function
#----------------------------------------------
function Search-MailboxGUILauncher
{

    #region Import the Assemblies
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    #endregion Import the Assemblies

    #region Form Objects instantiation
    $frmSelectLanguage = New-Object System.Windows.Forms.Form
    $comboLang = New-Object System.Windows.Forms.ComboBox
    #endregion Form Objects instantiation

    #region Events handling actions configuration
    $comboLang_TextChanged = {
        switch ($comboLang.SelectedItem)
        {
            "English (EN)"	{$Script:SelectedLanguage = "EN"}
            "Français (FR)"	{$Script:SelectedLanguage = "FR"}
        }
        $frmSelectLanguage.Close()
    }
	
    $frmSelectLanguage_Close = {
        If (($comboLang.SelectedItem -eq $null) -or ($comboLang.SelectedItem -eq ""))
        {
            $Script:SelectedLanguage = "None"
        }
    }
    #endregion Events handling actions configuration

    #region Form Configuration
    $frmSelectLanguage.StartPosition = "CenterScreen"
    $frmSelectLanguage.Name = 'frmSelectLanguage'
    $frmSelectLanguage.Text = 'Langue/Language'
    $frmSelectLanguage.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 308
    $System_Drawing_Size.Height = 47
    $frmSelectLanguage.ClientSize = $System_Drawing_Size
    $frmSelectLanguage.add_Closed($frmSelectLanguage_Close)
    #endregionForm Configuration

    #region Configuring and adding form objects
    #
    $comboLang.Items.Add('English (EN)')|Out-Null
    $comboLang.Items.Add('Français (FR)')|Out-Null
    $comboLang.Text = 'Choose/Choisissez'
    $comboLang.Name = 'comboLang'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 214
    $System_Drawing_Size.Height = 21
    $comboLang.Size = $System_Drawing_Size
    $comboLang.FormattingEnabled = $True
    $comboLang.TabIndex = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 42
    $System_Drawing_Point.Y = 12
    $comboLang.Location = $System_Drawing_Point
    $comboLang.DataBindings.DefaultDataSourceUpdateMode = 0
    $comboLang.add_TextChanged($comboLang_TextChanged)
    # Then add the defined form object to the form using Controls.Add()
    $frmSelectLanguage.Controls.Add($comboLang)
    #
    #endregion Configuring and adding form objects

    #region Final form operations (show and return value)
    #Show the Form
    $frmSelectLanguage.ShowDialog()| Out-Null
    #Return the selected language with a scope "global" to the script
    $script:SelectedLanguage
    #endregion Final form operations (show and return value)
} 
#----------------------------------------------
#endregion Launcher and language selection function
#----------------------------------------------


#Clear the screen and call the language selection function, followed by the Search GUI function
cls
$Language = Search-MailboxGUILauncher
Switch ($Language)
{
    "None" {Log "Ok_"}
    default {Search-MailboxGUI $language}
}
