#######################################################################
#            SCRIPT D'EXPORT CSV D'INFOS COUPLE PC/MONITEUR           #
#                         CLEMENT FOURSANS                            #
#                         03 JUIN 2019, 0.1                           #
#######################################################################

#Début de Script

#Saisie dans l'Input du ComputerName de la machine distante
$InputComputer = Read-Host " Le script affichera et exportera automatiquement : 
    - Le nom Netbios de l'ordinateur interrogé
    - Le constructeur, numéro de série, FriendlyName de la machine et de l'écran qui lui est associé
Saisissez le nom de la machine à interroger, laisser vide pour interroger pour la machine locale (Le résultat sera sauvegardé dans le répertoire courant)"


#Les informations des moniteurs sont retournées en Hexadecimal, fonction de conversion###############
Function ConvertTo-Char                                                                             
(	
	$Array
)
{
	$Output = ""
	ForEach($char in $Array)
	{	$Output += [char]$char -join ""
	}
	return $Output
}
######################################################################################################


#Début de Condition 1 : Si le nom de machine rentré est vide, exécution de la requête localement
if (!$InputComputer) {

#Requête WMI relative aux informations du moniteur
$Query = Get-WmiObject -Query "Select * FROM WMIMonitorID" -Namespace root\wmi -Impersonation 3
#Requête WMI relative aux informations de l'ordinateur
$Query2 = Get-WmiObject win32_bios

#Début de boucle, pour chaque $Monitor retourné par la requête $Query, création d'un objet dont les propriétés sont en ascii
$Results = ForEach ($Monitor in $Query)

{    
	New-Object PSObject -Property @{
		ComputerName = $env:ComputerName
		Active = $Monitor.Active
		Manufacturer = ConvertTo-Char($Monitor.ManufacturerName)
		UserFriendlyName = ConvertTo-Char($Monitor.userfriendlyname)
		SerialNumber = ConvertTo-Char($Monitor.serialnumberid)
	}
}
#Fin de boucle

#Création d'un objet dont les propriétés sont les valeurs extraites des requêtes précédentes
$report = New-Object PSObject
$report | Add-Member -MemberType NoteProperty -name Ordinateur -Value $env:ComputerName
$report | Add-Member -MemberType NoteProperty -name Constructeur_Ordinateur -Value $Query2.Manufacturer
$report | Add-Member -MemberType NoteProperty -name Nom_Ordinateur -Value $Query2.Name
$report | Add-Member -MemberType NoteProperty -name Serial_Ordinateur -Value $Query2.SerialNumber
$report | Add-Member -MemberType NoteProperty -name Constructeur_Moniteur -Value $Results.Manufacturer
$report | Add-Member -MemberType NoteProperty -name Nom_Moniteur -Value $Results.UserFriendlyName
$report | Add-Member -MemberType NoteProperty -name Serial_Moniteur -Value $Results.SerialNumber


#Export de l'objet dans un fichier CSV
$Report | Export-Csv .\$env:ComputerName.csv -NoTypeInformation

#Export à l'écran
Write-Host $Report
"Le résultat a été sauvegardé dans le répertoire courant sous $env:ComputerName.csv."
[void](Read-Host 'Appuyer sur Entrée pour quitter le script.')

}

#Fin de Condition 1

#Début de Condition 2, si un nom de machine est rentré, ajout de -ComputerName $InputComputer aux requêtes pour exécution à distance
else {

try {
$Query = Get-WmiObject -Query "Select * FROM WMIMonitorID" -Namespace root\wmi -Impersonation 3 -ComputerName $InputComputer
$Query2 = Get-WmiObject win32_bios -ComputerName $InputComputer

#Début de boucle
$Results = ForEach ($Monitor in $Query)

{    
	New-Object PSObject -Property @{
		ComputerName = $env:ComputerName
		Active = $Monitor.Active
		Manufacturer = ConvertTo-Char($Monitor.ManufacturerName)
		UserFriendlyName = ConvertTo-Char($Monitor.userfriendlyname)
		SerialNumber = ConvertTo-Char($Monitor.serialnumberid)
	}
}
#Fin de boucle


$report = New-Object PSObject
$report | Add-Member -MemberType NoteProperty -name Ordinateur -Value $env:ComputerName
$report | Add-Member -MemberType NoteProperty -name Constructeur_Ordinateur -Value $Query2.Manufacturer
$report | Add-Member -MemberType NoteProperty -name Nom_Ordinateur -Value $Query2.Name
$report | Add-Member -MemberType NoteProperty -name Serial_Ordinateur -Value $Query2.SerialNumber
$report | Add-Member -MemberType NoteProperty -name Constructeur_Moniteur -Value $Results.Manufacturer
$report | Add-Member -MemberType NoteProperty -name Nom_Moniteur -Value $Results.UserFriendlyName
$report | Add-Member -MemberType NoteProperty -name Serial_Moniteur -Value $Results.SerialNumber

$Report | Export-Csv .\$InputComputer.csv -NoTypeInformation
Write-Host $Report
"Le résultat a été sauvegardé dans le répertoire courant sous $InputComputer.csv."
[void](Read-Host 'Appuyer sur Entrée pour quitter le script.')

}

#Gestion d'erreur
catch {

Write-Host "Impossible d'accéder à l'Hôte Distant, assurez - vous que vous disposez de privilèges suffisants sur la machine d'exécution." -ForegroundColor Red


}

}

#Fin de condition 2
#Fin de Script