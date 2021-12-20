
function Eliza-Run-PartyInvite($partyList, $inviteFolder) {

    if (($partyList | Test-Path) -eq $True) {

        Write-Host "Hey, It's a party " + $inviteFolder

        $inviteList = Import-Csv $partyList

        $inviteList | ForEach-Object {

            $room = $_.RoomNum
            $partyName = $_.PartyName
            $partyCode = $_.PartyCode
            $arrival = $_.ArrivalTime
            $email = $_.Email

            Write-Host "Hey, $room , What is your party code? Answer is $partyCode"

        }


    } else {
        
        Write-Host "Hey, you stole my code"
    }

    
}

function Eliza-Check-PartyInvite($roomNum, $partyList) {

    $inviteList = Import-Csv $partyList

    $inviteList | ForEach-Object {
            
            $room = $_.RoomNum
            $partyName = $_.PartyName
            $partyCode = $_.PartyCode
            $arrival = $_.ArrivalTime
            $email = $_.Email

            if ($roomNum -eq $room) {
                $checkPartyCode = Read-Host "What is your Party Code? "

                Write-Host "Answer Is: " + $partyCode
                Write-Host "You Typed: " + $checkPartyCode
                Write-Host "Your Name: " + $partyName
                Write-Host "Your Email: " + $email

                if ($partyCode -eq $checkPartyCode) {
                    Write-Host "Welcome to the Party at " + $arrival
                    Break
                } else {
                    Write-Host "Sorry, we can't let you in"
                    Break
                }
            }
    }



    #Write-Host "Hello, World " + $inviteList.Count

}

function Eliza-Generate-PartyInvites($partyList, $inviteFolder) {

    $i = 0

    $inviteList = Import-Csv $partyList

    $email_count = $inviteList.count

    $inviteList | ForEach-Object {
        $i += 1;
        Write-Progress -Activity "Generating Emails for $i of $email_count" -Status "Percentage created" -PercentComplete ($i/$email_count*100);

        $room = $_.RoomNum
        $partyName = $_.PartyName
        $partyCode = $_.PartyCode
        $arrival = $_.ArrivalTime
        $email = $_.Email

        Write-Host "Generating Email for $email"
        
        $emailBody =  "Hey <b>" + $partyName + "</b> from room # <b>" + $room + "</b>,<br />"
        $emailBody += "You are invite to my awesome party on November 5th, 2022 at the Eliza Dorm Room <br />"
        $emailBody += "You'd need your secret party code <b>" + $partyCode + "</b> to be able to gain entry. So please remember it before coming to the party<br /><br />"
        $emailBody += ""
        $emailBody += "Your, Party Master - Eliza"

        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress("elizabeth@smagmedia.net")
        $mailMessage.To.Add($email)
        $mailMessage.Subject = "Welcome to Eliza's Dashing Party"
        $mailMessage.IsBodyHtml = $true
        $mailMessage.Body = $emailBody

        $smtpClient = New-Object System.Net.Mail.SmtpClient
        $smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory;
        $smtpClient.PickupDirectoryLocation = $inviteFolder
        $smtpClient.Send($mailMessage)
        $smtpClient.Dispose()

        $mailMessage.Dispose()
         

        #$i += 1;
        #Write-Progress -Activity "Generating Email File $i of $file_count" -Status "Percentage created" -PercentComplete ($i/$file_count*100);

    }
}