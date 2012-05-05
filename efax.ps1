[Reflection.Assembly]::LoadFile("c:\program files\microsoft\exchange\web services\1.1\Microsoft.Exchange.WebServices.dll")
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$s.Credentials = new-object net.networkcredential('xxxxxx', 'xxxxx', 'xxxxxxxxx.domain.com')
$s.AutoDiscoverUrl("xxxxx@domain.com")
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$softdel = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete

#$properties = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
#$properties.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$infLoop = 10

while ($infLoop -eq 10) 
{
	$inbox.FindItems(1) | % {
		
		#Check to see if first digit is '1' and also that the subject contains only digits.
		$trimSubj = $_.Subject.trim()

		if ($trimSubj[0] -eq '1' -AND $trimSubj -match "^\d+$")
		{
			$faxNum = $trimSubj
			$fwdEmail = $faxNum + "@efaxsend.com"
			$_.Forward("PeopleMatter eFax Service.", $fwdEmail)
			$replyMsg = "Your fax has been forwarded to the fax service using fax number: " + $faxNum
			$_.Reply($replyMsg, $false)
		
		}

		#Check to see if subject contains only digits and is the correct length
		elseif ($trimSubj -match "^\d+$" -AND ($trimSubj.Length -eq 10))
		{
			$faxNum = "1" + $trimSubj
			$fwdEmail = $faxNum + "@efaxsend.com"
			$_.Forward("PeopleMatter eFax Service.", $fwdEmail)
			$replyMsg = "Your fax has been forwarded to the fax service using fax number: " + $faxNum
			$_.Reply($replyMsg, $false)
			
		}

		else
		{
			# remove parenthesis, hyphens, and periods
			$faxNum = $trimSubj.replace("(", "").replace(")","").replace("-", "").replace(".","")

			# check to see if number is valid
			if ($faxNum -match "^\d+$")
			{
				if ($faxNum.Length -eq 10)
				{
					$faxNum = "1" + $faxNum
				
				}
				
				write-host "Fax num used: " $faxNum
				$fwdEmail = $faxNum + "@efaxsend.com"
				$_.Forward("PeopleMatter eFax Service.", $fwdEmail)
				$replyMsg = "Your fax has been forwarded to the fax service using fax number: " + $faxNum
				$_.Reply($replyMsg, $false)
			}
			
			else
			{
				$_.Forward("Error recognizing fax number.. Might be incoming fax", "xxxxxx@domain.com")
			}
		}
			
			

	$_.Delete($softdel)

	}
			

	start-sleep 20
}
