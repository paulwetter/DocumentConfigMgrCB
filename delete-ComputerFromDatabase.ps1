$Resource = [wmi](Get-WmiObject -Class SMS_R_SYSTEM -Namespace root\sms\site_pri -ComputerName cm12 -Filter "Name = 'XA002'").__PATH

Remove-WmiObject -InputObject $Resource.__PATH