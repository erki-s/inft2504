#Kobler til tjenestene som kreves for å utføre scriptet
Connect-MgGraph -Scopes "User.ReadWrite.All"
Connect-AzureAD

###################################################################################################
Connect-ExchangeOnline
$groups = Import-Csv ./groups.csv -Delimiter ";"

foreach ($group in $groups) {
    New-DistributionGroup -Name $group.groupName -DisplayName $group.groupName -PrimarySmtpAddress ($group.groupName + "@suspiciousdomain.onmicrosoft.com")
}

# Lager gruppe for alle ansatte
New-UnifiedGroup -DisplayName "Alle ansatte"

$rooms = Import-Csv ./rooms.csv -Delimiter ";"
# Lager et rom for de ansatte for å booke
foreach ($room in $rooms) {
    New-Mailbox -Name ($room.name + "@suspiciousdomain.onmicrosoft.com") `
        -DisplayName $room.name `
        -Alias $room.alias `
        -Room -EnableRoomMailboxAccount $true `
        -RoomMailboxPassword (ConvertTo-SecureString -String $room.password -AsPlainText -Force)
}

###################################################################################################
# Lager Teams
Connect-MicrosoftTeams
$teams = Import-Csv ./teams.csv -Delimiter ";"

foreach ($team in $teams) {
    New-Team -DisplayName $team.teamName -Description $team.teamDesc -Visibility $team.teamVisibility -MailNickName ($group.name + "Team")
}

###################################################################################################

Connect-SPOService -url "https://suspiciousdomain.sharepoint.com" -Credential erki@suspiciousdomain.onmicrosoft.com
$sites = Import-Csv ./sharepoint.csv -Delimiter ";"

foreach ($site in $sites) {
    New-SPOSite -Url $site.url -Owner erki@suspiciousdomain.onmicrosoft.com -StorageQuota $site.storage -Title $site.title
}

###################################################################################################
#Oppretter brukere og legger dem i gruppene de skal være i

#Angir domene
$Domain = "@suspiciousdomain.onmicrosoft.com"
#Importerer Csv filen med brukerne
$users = Import-Csv ./users.csv -Delimiter ";"


foreach ($user in $users) {
    #Henter passordet fra Csv
    $PasswordProfile = @{Password = $user.Password}
    #Lager brukeren
    New-MgUser `
        -DisplayName ($user.FirstName + " " + $user.LastName) `
        -GivenName $user.FirstName `
        -Surname $user.LastName `
        -UserPrincipalName ($user.FirstName + $user.LastName + $Domain) `
        -PasswordProfile $PasswordProfile `
        -MailNickname ($user.FirstName + "." + $user.LastName) `
        -AccountEnabled

    $username = $user.FirstName + $user.LastName + $Domain

    #Legger til bruker i gruppe for sin avdeling
    Add-DistributionGroupMember -Identity $user.Avdeling -Member ($user.FirstName + " " + $user.LastName)

    #Legger til bruker i Alle ansatte
    Add-UnifiedGroupLinks -Identity "Alle ansatte" -LinkType Members -Links $username

    #Finner IDen til Team
    $teamId = Get-Team -DisplayName $user.Avdeling

    #Legger brukeren inn i team
    Add-TeamUser -GroupId $teamId.GroupId -User $username -role "Member"

    #Legger til brukeren i SharePoint
    Add-SPOUser -Site ("https://suspiciousdomain.sharepoint.com/sites/" + $user.Avdeling) -LoginName $username -Group $user.Avdeling
}

# Kobler fra tjenestene
Disconnect-MgGraph
Disconnect-AzureAD
Disconnect-ExchangeOnline
Disconnect-MicrosoftTeams
Disconnect-SPOService
###################################################################################################