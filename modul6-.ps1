# Import ExchangeOnlineManagement module and connect to Exchange Online
Find-Module -Name ExchangeOnlineManagement | Install-Module -Force
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Create dynamic distribution groups
New-DynamicDistributionGroup -Name "Alle Trondheim" -IncludedRecipients "MailboxUsers" -ConditionalStateOrProvince "Trondheim"
New-DynamicDistributionGroup -Name "Alle Oslo" -IncludedRecipients "MailboxUsers" -ConditionalStateOrProvince "Oslo"
New-DynamicDistributionGroup -Name "Alle HR" -IncludedRecipients "MailboxUsers" -ConditionalDepartment "HR"

# Create new users and assign licenses (example users)
# Replace <UserPrincipalName>, <DisplayName>, <FirstName>, <LastName>, and <LicenseSkuId> with actual values
New-MsolUser -UserPrincipalName "trondheim.user@edudev365.onmicrosoft.com" -DisplayName "Trondheim User" -FirstName "Trondheim" -LastName "User" -UsageLocation "NO"
Set-MsolUserLicense -UserPrincipalName "trondheim.user@edudev365.onmicrosoft.com" -AddLicenses "<LicenseSkuId>"

New-MsolUser -UserPrincipalName "oslo.user@edudev365.onmicrosoft.com" -DisplayName "Oslo User" -FirstName "Oslo" -LastName "User" -UsageLocation "NO"
Set-MsolUserLicense -UserPrincipalName "oslo.user@edudev365.onmicrosoft.com" -AddLicenses "<LicenseSkuId>"

New-MsolUser -UserPrincipalName "hr.user@edudev365.onmicrosoft.com" -DisplayName "HR User" -FirstName "HR" -LastName "User" -UsageLocation "NO"
Set-MsolUserLicense -UserPrincipalName "hr.user@edudev365.onmicrosoft.com" -AddLicenses "<LicenseSkuId>"

# Create distribution list "Alle ansatte" and add all users
New-DistributionGroup -Name "Alle ansatte" -DisplayName "Alle ansatte" -PrimarySmtpAddress "alle.ansatte@edudev365.onmicrosoft.com"

# Add users to "Alle ansatte" distribution list using a CSV file or a variable
# Example using a CSV file
$users = Import-Csv -Path "C:/03-02-Users.csv"
foreach ($user in $users) {
    Add-DistributionGroupMember -Identity "Alle ansatte" -Member $user.UserPrincipalName
}

# Example using a variable
$allMailUsers = Get-Mailbox -RecipientTypeDetails UserMailbox
foreach ($user in $allMailUsers) {
    Add-DistributionGroupMember -Identity "Alle ansatte" -Member $user.Alias
}

# Create meeting rooms
New-Mailbox -Name "MeetingRoom1" -DisplayName "Meeting Room 1" -Alias "MeetingRoom1" -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String "YourSecurePassword1!" -AsPlainText -Force)
New-Mailbox -Name "MeetingRoom2" -DisplayName "Meeting Room 2" -Alias "MeetingRoom2" -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String "YourSecurePassword2!" -AsPlainText -Force)

# Set resource capacity and equipment for the meeting rooms
Get-Mailbox -Identity "MeetingRoom1" | Set-Mailbox -ResourceCapacity 12
Get-Mailbox -Identity "MeetingRoom2" | Set-Mailbox -ResourceCapacity 8

# Optionally set additional properties for meeting rooms, like equipment
Set-CalendarProcessing -Identity "MeetingRoom1" -ResourceDelegates "resource.delegate@edudev365.onmicrosoft.com"
Set-CalendarProcessing -Identity "MeetingRoom2" -ResourceDelegates "resource.delegate@edudev365.onmicrosoft.com"

# Example: Book a meeting in one of the rooms (replace with actual values)
$startTime = [datetime]::Parse("2023-06-01T10:00:00")
$endTime = [datetime]::Parse("2023-06-01T11:00:00")
New-Event -Subject "Team Meeting" -Start $startTime -End $endTime -Location "Meeting Room 1" -Attendees "user1@edudev365.onmicrosoft.com", "user2@edudev365.onmicrosoft.com"

# Note: The New-Event cmdlet is a placeholder; use appropriate cmdlets to book meetings via Outlook/Exchange APIs
