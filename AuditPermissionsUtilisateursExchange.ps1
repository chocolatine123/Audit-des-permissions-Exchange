############## Lister les Administrateurs ainsi que les droits Exchange ######################
#### Auteur : Pierre-Alban Maurin
#### Nécessite Microsoft Exchange Web Services Managed API 1.2 ou versions supérieures
#### Lien de téléchargement : https://www.microsoft.com/en-us/download/details.aspx?id=42951
#### Nécessite les Outils d’administration de serveur distant pour Windows 10
#### Lien de téléchargement : https://www.microsoft.com/fr-FR/download/details.aspx?id=45520
#### Nécessite d'être membre de Organization Management
#### Nécessite en outre de posséder le rôle ApplicationImpersonation (présent dans le groupe Hygiene Management)
#### Nécessite d'être exécuté à partir d'une machine membre du domaine
#### OutputDir doit se trouver dans le même répertoire que ce script, ainsi que Get-MailboxFolderPermissionEWS.ps1
#### Ce dernier est à télécharger via ce lien : https://gallery.technet.microsoft.com/scriptcenter/Get-MailboxFolderPermission-fdf1f90f
####


################ Initialisation des variables ################
Write-Host "Initialisation des variables" -BackgroundColor Red
$dossier = Read-Host "Output Dir "
$ex = Read-Host "Entrer le nom de votre serveur Exchange "
$date = Get-Date
$mois = $date.Month
$an = $date.Year
$jour = $date.Day



################ Importation des modules nécessaires ################
Write-Host "Importation des modules nécessaires" -BackgroundColor Red
Import-Module ActiveDirectory
$Params = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri = "http://$ex/PowerShell/"
    Credential = ( $credential )
    Authentication = 'Kerberos'
    Name = 'ExchangeSession'
}
$WarningPreference = 'SilentlyContinue'
Import-PSSession -Session ( New-PSSession @Params )
$d = Get-ADDomainController
$dc = $d.HostName
$dom = $env:userdomain
cd $dossier
Write-Host "Done"


################ Lister les admins de Domaine ################
Write-Host "Lister les admins de Domaine" -BackgroundColor Red
$membresadda = Get-ADGroupMember -identity "Domain Admins" -Server $dc -Recursive | Select samaccountname,distinguishedName,SID
$membresadea = Get-ADGroupMember -identity "Enterprise Admins" -Server $dc -Recursive | Select samaccountname,distinguishedName,SID
$membresadadm = Get-ADGroupMember -identity Administrators -Server $dc -Recursive | Select samaccountname,distinguishedName,SID
Write-Host "Done"

################ Lister les admins Exchanges ################
Write-Host "Lister les admins Exchange" -BackgroundColor Red
$membresex = Get-ADGroupMember -identity "Organization Management" -Server $dc -Recursive | Select samaccountname,distinguishedName,SID
Write-Host "Done"

############### Export des noms en fichier texte #####################
Write-Host "Export des noms en fichier texte" -BackgroundColor Red
New-Item -Name "DroitsAdministrateurs" -ItemType directory
cd .\DroitsAdministrateurs
$datejour1 = "AdminsExhange."+$an+"."+$mois+"."+$jour+""
$membresadda + " " + $membresadea + " " + $groupeadadm + " " + $membresex | Out-File "temp-$datejour1.txt"
Type "*.txt" | Select -Unique > "$datejour1.txt"
(gc "$datejour1.txt") | ? {$_.trim() -ne "" } | set-content "$datejour1.txt"
Write-Host "Done"
Write-Host "Votre fichier recensant les admins Exchange se trouve dans le repertoire suivant : $dossier\DroitsAdministrateurs" -BackgroundColor Black


################ Récupération des données ################
Write-Host "Récupération des données" -BackgroundColor Red

###Droits Admins
$datejour = "DroitsExhange."+$an+"."+$mois+"."+$jour+""
Get-Mailbox |Get-MailboxPermission |select User,AccessRights,Identity |Export-Csv "$datejour.csv"
Write-Host "Done"
Write-Host "Votre fichier recensant les droits Exchange se trouve dans le repertoire suivant : $dossier\DroitsAdministrateurs" -BackgroundColor Black
Remove-Item .\Temp-*
cd ..


###Droits Utilisateurs
New-Item -Name "DroitsUsers" -ItemType directory
cd .\DroitsUsers
$datejour = "Droits-Users."+$an+"."+$mois+"."+$jour+""
Get-Mailbox -ResultSize unlimited | ..\Get-MailboxFolderPermissionEWS.ps1 -Server $ex -Impersonate -TrustAnySSL -Threads 20 -MultiThread |select User,Permissions,FolderName,Mailbox |Export-Csv "$datejour.csv"       #############
Write-Host "Done"
Write-Host "Votre fichier recensant les Permissions des utilisateurs Exchange se trouve dans le repertoire suivant : $dossier\DroitsUsers" -BackgroundColor Black
cd ..\DroitsAdministrateurs


######Partie 2 :
################ Tri des droits ################
Write-Host "Tri des données Admins" -BackgroundColor Red



####### Suppression des lignes vides ou en double #######
Get-Content .\DroitsExhange* | where { $_ -ne "$null" } |Out-File .\temp-droits-1.csv
Type .\temp-droits-1.csv | Select -Unique > .\temp-droits-2.csv


####### Suppression des droits NT AUTHORITY #######
## On supprimme SELF car ce dernier pointe vers les droits de l'utilsiateur concerné, donc pas de problèmes à ce niveau
## Idem pour SYSTEM et NETWORK SERVICE car il s'agit de comptes spéciaux de sécurité ;
## Etant donné qu'ils ont beaucoup de droits, ils apparaissent logiquement ici, mais on peut les éliminer d'office.
Get-Content .\temp-droits-2.csv | Where-Object {$_ -notmatch "NT AUTHORITY"} |Out-File .\temp-droits-3.csv

####### Suppression des comptes administrateurs #######
## Les comptes administrateurs étant déjà référencés dans le fichier texte .\AdminsExhange.*, on les référence à part.
## Traitons d'abord lefichier .\AdminsExhange.* pour convertir en variable les noms des comptes admin, tant Exchange qu'Active Directory
Get-Content .\AdminsExhange.*| Select-Object -Skip 2 | Set-Content .\temp-admin-1.txt
$fichier = Get-Content .\temp-admin-1.txt
$file = foreach ($noms in $fichier) 
{
    $Split = $noms.Split(" ")
    $noms.Split(" ")[+0]
}

$file | Out-File .\temp-admin-2.txt 
$NomsAdmin = Get-Content .\temp-admin-2.txt
$t = "$dom\\" + $NomsAdmin[0]
Get-Content .\temp-droits-3.csv | Where-Object {$_ -notmatch "$t"} |Out-File .\temp-droits-4-0.csv

for ($i=1; $i -lt $file.Count; $i ++)
{
    $NomsAdmin = Get-Content .\temp-admin-2.txt
    $t = "$dom\\" + $NomsAdmin[$i]
    $h = $i - 1
    Get-Content .\temp-droits-4-$h.csv | Where-Object {$_ -notmatch "$t"} |Out-File .\temp-droits-4-$i.csv 
    $t = Get-Content .\temp-droits-4-$i.csv
    $TemFile = new-item "temp-droits-4.txt" –type file -force
    ADD-content -path temp-droits-4.csv -value $t
}

Remove-Item -Path .\temp-droits-4-*.csv


####### Suppression des groupse d'administration #######
## Sur le même principe que précédement, on dégage tous les groupes d'administration AD et Exchange
Get-Content .\temp-droits-4.csv | Where-Object {$_ -notmatch "EXCH\\Domain Admins"} |Out-File .\temp-droits-5.csv
Get-Content .\temp-droits-5.csv | Where-Object {$_ -notmatch "EXCH\\Enterprise Admins"} |Out-File .\temp-droits-6.csv
Get-Content .\temp-droits-6.csv | Where-Object {$_ -notmatch "EXCH\\Organization Management"} |Out-File .\temp-droits-7.csv
Get-Content .\temp-droits-7.csv | Where-Object {$_ -notmatch "EXCH\\Administrators"} |Out-File .\temp-droits-8.csv

##Idem avec les groupes de gestions des serveurs Exchange
Get-Content .\temp-droits-8.csv | Where-Object {$_ -notmatch "EXCH\\Exchange Servers"} |Out-File .\temp-droits-9.csv
Get-Content .\temp-droits-9.csv | Where-Object {$_ -notmatch "EXCH\\Exchange Trusted Subsystem"} |Out-File .\temp-droits-10.csv
Get-Content .\temp-droits-10.csv | Where-Object {$_ -notmatch "EXCH\\Managed Availability Servers"} |Out-File .\temp-droits-11.csv

##Idem avec les dossiers publiques, Delegated Setup et Discovery Management
Get-Content .\temp-droits-11.csv | Where-Object {$_ -notmatch "EXCH\\Public Folder Management"} |Out-File .\temp-droits-12.csv
Get-Content .\temp-droits-12.csv | Where-Object {$_ -notmatch "EXCH\\Delegated Setup"} |Out-File .\temp-droits-13.csv
Get-Content .\temp-droits-13.csv | Where-Object {$_ -notmatch "EXCH\\Discovery Management"} |Out-File .\temp-droits-14.csv

Get-Content .\temp-droits-14.csv | Select -Unique >  Droits-Admin-Def.csv
Remove-Item -Path .\temp-*
Write-Host "Done"

########################   Traitement Droits Users   ################################
################ Tri des droits ################
cd ..\DroitsUsers
Write-Host "Tri des données Users" -BackgroundColor Red
####### Suppression des lignes vides ou en double #######
Get-Content .\Droits-Users* | where { $_ -ne "$null" } |Out-File .\temp-droits-User-1.csv
Type .\temp-droits-User-1.csv | Select -Unique > .\temp-droits-User-2.csv

Get-Content .\temp-droits-User-2.csv | Where-Object {$_ -notmatch "None"} |Out-File .\temp-droits-User-3.csv
#Get-Content .\temp-droits-User-3.csv | Where-Object {$_ -notmatch "Default"} |Out-File .\temp-droits-User-4.csv
#Get-Content .\temp-droits-User-4.csv | Where-Object {$_ -notmatch "Anonymous"} |Out-File .\temp-droits-User-5.csv
Get-Content .\temp-droits-User-3.csv | Where-Object {$_ -notmatch "FreeBusyTimeOnly"} |Out-File .\temp-droits-User-4.csv
Get-Content .\temp-droits-User-4.csv |select -Unique > Droits-Users-Def.csv
Remove-Item -Path .\temp-*
Write-Host "Done"

cd ..\
Remove-Item -Path .\temp-*


####Partie 3 :
################ Exportation du rapport en HTML ################
Write-Host "Exportation du rapport en HTML" -BackgroundColor Red


$header=@"
<head>
<title>Audit Exchange Rapport</title>
</head><body>

<table>
<colgroup><col/><col/><col/><col/></colgroup>
<tr><th>User</th><th>AccessRights</th><th>Identity</th></tr>

<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>

"@
$footer=@"
</table>
</body></html>
"@
$body=@"
<h1>Droits Admin Exchange</h1>
<h5>Généré le $(Get-Date)
"@

Import-Csv .\DroitsAdministrateurs\Droits-Admin-Def.csv | ForEach{
    $body+="<tr><td>$($_.User)</td><td>$($_.AccessRights)</td><td>$($_.Identity)</td></tr>"
}

    $body+="
        <table>
    <br><br>
    <h1>Droits Users Exchange</h1>
<h5>Généré le $(Get-Date)

<colgroup><col/><col/><col/><col/></colgroup>
<tr><th>User</th><th>Permissions</th><th>FolderName</th><th>Mailbox</th></tr>
    "

Import-Csv .\DroitsUsers\Droits-Users-Def.csv | ForEach{


    $body+="<tr><td>$($_.User)</td><td>$($_.Permissions)</td><td>$($_.FolderName)</td><td>$($_.Mailbox)</td></tr>"
}

New-Item -Name "Rapport" -ItemType directory
-join $header,$body,$footer | Out-File .\Rapport\RapportAudit.html
Invoke-Expression .\Rapport\RapportAudit.html
Write-Host "Votre rapport final se trouve dans le repertoire suivant : $dossier\Rapport" -BackgroundColor Black
Write-Host "Fin"
###### FIN ###### 
