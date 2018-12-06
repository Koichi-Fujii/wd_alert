
Param($arg1)

############################################################
#
# [History]
#	date		rev	name	comment
#	2018/12/03	1.0	Fujii	New
#
############################################################

$rev=1.0
$ret=0

$First5 = Get-WinEvent -LogName "Microsoft-Windows-Windows Defender/Operational" | Where-Object { $_.Id -eq $arg1 } | Select-Object -first 5
$First1 = $First5 | Select-Object -first 1
$Body = $First1 | Format-List -property * | Out-String
$Footer = $First5 | Format-Table -autosize | Out-String

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "admin@peptistar.site"
$Mail.Subject = "[Alert-"+$First1.Id+"] Microsoft-Windows-Windows Defender "+$First1.MachineName
$Mail.BodyFormat = 1
$Mail.Importance = 2
$Mail.Body = $Body+$Footer
$Mail.Send()

if(!($?)){$ret=1}

exit $ret
