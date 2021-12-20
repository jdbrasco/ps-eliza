CD C:\TFS\SCRAP\ps-eliza

Import-Module .\Eliza-PartyInvite.psd1

Clear

Eliza-Run-PartyInvite -partyList "Party.csv" -inviteFolder "invite"

Eliza-Check-PartyInvite -roomNum 38 -partyList "Party500s.csv"

Eliza-Generate-PartyInvites -partyList "Party500s.csv" -inviteFolder "C:\TFS\SCRAP\ps-eliza\Invites"

Remove-Module Eliza-PartyInvite