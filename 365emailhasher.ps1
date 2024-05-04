# https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/generating-md5-hashes-from-text
Function Get-StringHash 
{ 
    param
    (
        [String] $String,
        $HashName = "MD5"
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
    $algorithm = [System.Security.Cryptography.HashAlgorithm]::Create('MD5')
    $StringBuilder = New-Object System.Text.StringBuilder 
  
    $algorithm.ComputeHash($bytes) | 
    ForEach-Object { 
        $null = $StringBuilder.Append($_.ToString("x2")) 
    } 
  
    $StringBuilder.ToString() 
}

if ($args.length -eq 0) {
	Write-Host "Add Graph-accessible bearer token as argument!"
	exit(1)
} else {
	$token = $args[0]
}

$headers = @{
	"Authorization" = "Bearer $token"
}

# Get user information
$usersResponse = Invoke-RestMethod -Method GET -Headers $headers -Uri https://graph.microsoft.com/v1.0/users 
foreach ($userInfo in $usersResponse.value) {
	# Get messages
	$userId = $userInfo.id
	$userPrincipalName = $userInfo.userPrincipalName
	try {
		$messageInfo = Invoke-RestMethod -Method GET -Headers $headers -Uri https://graph.microsoft.com/v1.0/users/$userId/messages -ErrorAction SilentlyContinue
		foreach($message in $messageInfo.value) {
			$messageSubject = $message.subject
			$messageDate = $message.receivedDateTime
			$sender = $message.sender.emailAddress.address
			if ($message.hasAttachments) {
				$messageId = $message.id
				$attachmentInfo = Invoke-RestMethod -Method GET -Headers $headers -Uri https://graph.microsoft.com/v1.0/users/$userId/messages/$messageId/attachments
				foreach($attachment in $attachmentInfo.value) {
					$decodedMsg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($attachment.ContentBytes))
					$md5hash = Get-StringHash $decodedMsg
					$output = [ordered]@{
						"Sender" = $sender
						"Recipient" =  $userPrincipalName
						"Message Subject" = $messageSubject
						"Received time" = $messageDate
						"Attachment Hash" = $md5hash
					}
					$output
				}
			}
		}
	} catch {}
}
