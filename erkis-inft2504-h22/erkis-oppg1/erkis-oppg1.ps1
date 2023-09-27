#Kobler til tjenestene som kreves for å utføre scriptet
Connect-MgGraph -Scopes "User.ReadWrite.All"
Connect-AzureAD

#Angir domene
$Domain = "@suspiciousdomain.onmicrosoft.com"
#Importerer Csv filen med brukerne
$users = Import-Csv ./test.csv -Delimiter ";"


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
    #Finner IDen til brukeren
    $userObj =  Get-MgUser -Filter "UserPrincipalName eq '$username'" | Select-Object -ExpandProperty Id
    #Finner IDen til gruppen
    $groupObj =  Get-AzureADGroup -SearchString $user.GroupName | Select-Object -ExpandProperty ObjectID
    #Legger brukeren inn i gruppen
    Add-AzureADGroupMember -ObjectId $groupObj -RefObjectId $userObj
}