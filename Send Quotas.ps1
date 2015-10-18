# Creating a remote powershell session to the FQDN of your exchange server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ExchangeFQDN/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

#gets all mailboxes in the Exchange organization. 
$u = Get-Mailbox 


#Threshold size in MB
$CloseToThreshold = 512
$AdminEmail = "Administrator@domain.com"
#Maximum mailbox size
$MaxMailboxSize = 3072 


Function SendAlertToUser
{
	
	#Write-Host $emailAddress $prohibitSendQuota $mailboxSize
	
	### E-mail message values 
	$FromAddress = "Exchange@domain.com"
	$ToAddress = [String] $emailAddress
	$MessageSubject = "MAILBOX QUOTA WARNING: You are reaching your mailbox quota!"
	$MessageBody = "You are about to hit your 'Prohibit Send Quota' of " + [String] $prohibitSendQuota + " MB. Once you reach this, you will not be able to send emails. Your current mailbox size is " + [String] $mailboxSize + " MB."
	$SendingServer = "FQDNOfYourExchangeServer"

	### Create the mail message and add the statistics text file as an attachment
	$SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $MessageBody
	
	$SMTPMessage.CC.Add($adminEmail )
	
	### Send the message
	$SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer
	$SMTPClient.Send($SMTPMessage)
}

Function SendAlertToAdmin
{
	
	#Write-Host $emailAddress $prohibitSendQuota $mailboxSize
	
	### E-mail message values 
	$FromAddress = "SendingAddress@domain.com"
	$ToAddress = [String] $emailAddress
	$MessageSubject = "MAILBOX " + [String] $emailAddress + " REACHED PROHIBIT SEND QUOTA!"
	$MessageBody = "Mailbox 'Prohibit Send Quota' size is " + [String] $prohibitSendQuota + " MB and current mailbox size is " + [String] $mailboxSize + " MB. User will not be able to send emails."
	$SendingServer = "FQDNOfYourExchangeServer"

	### Create the mail message and add the statistics text file as an attachment
	$SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $MessageBody
	
	$SMTPMessage.CC.Add($adminEmail )
	
	$SMTPMessage.Priority = [System.Net.Mail.MailPriority]::High	
	
	### Send the message
	$SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer
	$SMTPClient.Send($SMTPMessage)
}



foreach ($m in $u) 
	{
		
		#Finds the size of each mailbox and prints the name and mailbox size.
		$mailboxSize = (Get-MailboxStatistics -Identity $m).TotalItemSize.Value.ToMB()                                                                         
	
		Write-Host Mailbox: $m
		Write-Host Mailbox Size: $mailboxSize
		
		#Finds the current send quota.
		$prohibitSendQuota = (Get-Mailbox $m).ProhibitSendQuota.Value.ToMB()
        Write-Host prohibitSendQuota: $prohibitSendQuota                              
		
		#Gets the user's email address and displays it.
		$emailAddress = (Get-Mailbox $m).PrimarySmtpAddress
		Write-Host Email Address: $emailAddress
		
		
		#Calculates the buffer between the current mailbox size and the quota.
		$MailboxSizeQuotaDifference = $prohibitSendQuota - $mailboxSize

        #If the sendQuota is unlimited and the mailbox is less than the MaxMailboxSize, set the quota to $MaxMailboxSize
		if (($prohibitSendQuota -eq "Unlimited") -and ($mailboxSize -le $MaxMailboxSize))
			{
			Write-Host "User does not have a send quota. Setting quota to defined max mailbox size"
			set-mailbox $m -prohibitSendQuota $MaxMailboxSize
				}

        #If the sendQuota is Unlimited and the mailbox is greater than the defined max mailbox, set the new quota to 
        #mailboxsize plus threshold
		Elseif (($prohibitSendQuota -eq "Unlimited") -and ($mailboxSize -gt $MaxMailboxSize))
			{
			Write-Host "User does not have a send quota. Setting quota to mailbox size plus defined threshold"
                $NewQuota= ($mailboxSize + $CloseToThreshold)/1024
				write-host "User's new quota is now" $NewQuota
				set-mailbox $m -prohibitSendQuota $NewQuota"GB"
				}
		
		#Checks if the mailbox size difference is less than or equal to the quota threshold
		#if the buffer is less than or equal to the threshold, alert the users.

		Elseif (($MailboxSizeQuotaDifference -le $closeToThreshold) -and ($mailboxSize -lt $MaxMailboxSize))
			{		
				SendAlertToUser
				SendAlertToAdmin
				Write-Host "User mailbox is close to quota. Alert sent to user and Admins."
				} 
				
		#Else if the mailbox buffer is greater than the threshold ($CloseToThreshold) and the mailbox is bigger than $MaxMailboxSize, 
		#then add $CloseToThreshold to the current mailbox size and set that as the new prohibit send 
		Elseif (($MailboxSizeQuotaDifference -gt $closeToThreshold) -and ($mailboxSize -ge $MaxMailboxSize))
			{
				Write-Host "Space user has before they encounter send quota limit" $MailboxSizeQuotaDifference"."
                $NewQuota= ($mailboxSize + $CloseToThreshold)/1024
				write-host "User's new quota is now" $NewQuota
				set-mailbox $m -prohibitSendQuota $NewQuota"GB"

            }

        Elseif (($MailboxSizeQuotaDifference -lt $closeToThreshold) -and ($mailboxSize -ge $MaxMailboxSize))
        {
                SendAlertToUser
				SendAlertToAdmin
				Write-Host "User mailbox is close to quota. Alert sent to user and Admins."
        }

		#Otherwise the mailbox is fine.
		Else
			{
				Write-host "Mailbox is fine." 
				}
		}