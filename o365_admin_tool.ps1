Function welcome{

        Clear		
        Write-Host "******************************************************"  -foreground Red 
        Write-Host "     Welcome to Office365 Administrator Tool          " -foreground Red 
        Write-Host "******************************************************"  -foreground Red 
}

Function login{
		while(1){
				welcome;
				Import-Module MsOnline
				Write-Host "Enter Office365 account : " -nonewline -foreground "yello"
				$global:admaccount = Read-Host 
				 Write-Host "Enter Password : " -nonewline -foreground "yello"
				$global:admpassword_plain = Read-Host 
				$global:admpassword_encrypt = convertto-securestring  $admpassword_plain -asplaintext -force

				$global:cred=New-Object System.Management.Automation.PSCredential($admaccount,$admpassword_encrypt)
				$report = Connect-MsolService -credential $cred 2>&1
				$err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
				if($err){
					Write-Host "$report" -background Black -foreground Magenta	
					Read-Host
				}
				else{
					Write-Host "Login Success!" -BackgroundColor Black -ForegroundColor Green
					Start-Sleep -s 1
					break;
				}
		}	
}
Function Check_users_status{

        Write-Host "Enter user account you want to check : " -nonewline
        $careaccount = Read-Host
		$report = Get-MsolUser -UserPrincipalName $careaccount 2>&1
		$err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
		Write-Host "$report" -foreground "red"
		if($err){
			#Do Error Thing
		}
		else{
			Get-MsolUser -UserPrincipalName $careaccount | fl  #(format list) 
		}
}
Function Change_Password{
        Write-Host "1. Table mode 2. csv mode : " -NoNewline -ForegroundColor Yellow ;
        $choose = Read-Host
        if($choose -eq 1){
                $NewPass = Read-Host "Enter New Password "
                Get-MsolUser -all | Out-GridView -PassThru | ForEach-Object {
                        Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -NewPassword $NewPass -ForceChangePassword $false
                        Set-MsolUser -UserPrincipalName $_.UserPrincipalName -PasswordNeverExpires $true
                }
                Write-Host "Change Password Done!" -BackgroundColor Black -ForegroundColor Green
        }
        elseif ($choose -eq 2){
                Get-MsolUser -UserPrincipalName $admaccount | Select-Object UserPrincipalName,NewPassword | export-csv password.csv -encoding "utf8"
                $CurrentPath = $(Get-Location).ToString();
                Write-Host "We create a reference sample at $CurrentPath\password.csv"  -background Black -foreground Magenta
                while(1){
                        $Importfile = Read-Host "Please enter the csv filename ";
                        $report =  Get-Content -path $CurrentPath"\"$Importfile  2>&1;
                        $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
			            if($err){
					               Write-Host "Can't open file." -background Black -foreground Red	
			            }
			            else{
                                   Write-Host "Reading $Importfile ..." -background Black -foreground Magenta	
                                   Import-Csv $Importfile | ForEach-Object {
                                            Set-MsolUserPassword –UserPrincipalName $_.UserPrincipalName -NewPassword $_.NewPassword –ForceChangePassword $false
                                            Set-MsolUser -UserPrincipalName $_.UserPrincipalName -PasswordNeverExpires $true
                                   }
                                   Write-Host "Change Password Done!" -BackgroundColor Black -ForegroundColor Green
                                   Break;
			            }
                }
        }
}

Function Create_user{

        $UserPrincipalName = Read-Host "Account ";
        $NewPassword = Read-Host "Password";
        $FirstName = Read-Host "FirstName ";
        $LastName = Read-Host "LastName ";
        $DisplayName = Read-Host "DisplayName ";
        New-MsolUser -FirstName $FirstName -LastName $LastName -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -Password $NewPassword -ForceChangePassword $false
        Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation TW
        Get-MsolAccountSku |ft AccountSkuId,ActiveUnits,ConsumedUnits
        Write-Host "Choose the Licenses to Active " -nonewline
        Write-Host "1.STUDENT 2.FACULTY : " -nonewline -foreground Red
        $Licenses_no = Read-Host ;
        if($Licenses_no -eq "1"){
                $ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "ENTERPRISEPACK_STUDENT"};
        }
        elseif($Licenses_no -eq "2"){
                 $ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "ENTERPRISEPACK_FACULTY"} ;                
        }
        ForEach ($item in $ServicePlans){                         #ForEach ($<個別的項目或物件> in $<集合物件>)
                 Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $item.AccountSkuId
        }
        Write-Host "Create User Done!"; -BackgroundColor Black -ForegroundColor Green
}
Function Remove_user{

        Get-MsolUser -all | Out-GridView -PassThru | ForEach-Object { Remove-MsolUser -UserPrincipalName  $_.UserPrincipalName -force}
        Write-Host "Remove User Done!"; -BackgroundColor Black -ForegroundColor Green
}
Function Recover_user{

         Get-MsolUser -ReturnDeletedUsers | Out-GridView -PassThru | ForEach-Object {Restore-MsolUser -UserPrincipalName  $_.UserPrincipalName -AutoReconcileProxyConflicts}    
         Write-Host "Recover User Done!"; -BackgroundColor Black -ForegroundColor Green
}

Function  Send_mail
{
        $SMTPServer = "smtp.office365.com";
        $SMTPPort = "587";
        $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
        $smtp.EnableSSL = $true;
        $smtp.Credentials = New-Object System.Net.NetworkCredential($admaccount, $admpassword_plain);

        $From = $admaccount;
        Write-Host "From : $From"
        $To = Read-Host "To ";
        $subject = Read-Host "Subject ";
        $body = Read-Host "Content  ";
        $smtp.Send($From, $To, $subject, $body); 
}
Function See_mail_log{
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic –AllowRedirection
        Import-PSSession $Session
        Get-StaleMailboxReport | Out-GridView	
}
####Main Begin####

		 
		##直接登入
        <#Welcome;		
        Import-Module MsOnline
        $global:admaccount = "test@domain.com"
        $global:admpassword_plain = "P@ssw0rd"
        $admpassword_encrypt = convertto-securestring  $admpassword_plain -asplaintext -force
        $global:cred=New-Object System.Management.Automation.PSCredential($admaccount,$admpassword_encrypt)
        Connect-MsolService -credential $cred
		#>
		
		Login	##手動登入

		$i = 1;
		while ($i) {
				welcome;
				$ident = Get-MsolUser -UserPrincipalName $admaccount
				Write-Host "$ident "   -BackgroundColor Black -ForegroundColor Magenta
				Write-Host "Login : $admaccount "   -BackgroundColor Black  -ForegroundColor Magenta
				Write-Host "  1.Check users status"  -foreground Yellow 
				Write-Host "  2.Change password" -foreground Yellow 
				Write-Host "  3.Send mail"-foreground Yellow 
				Write-Host "  4.See mail log"-foreground Yellow 
				Write-Host "  5.Create user"-foreground Yellow
				Write-Host "  7.Remove user"-foreground Yellow
				Write-Host "  8.Recover user"-foreground Yellow  
				Write-Host "  9.Logout" -foreground Yellow
				Write-Host "  0.Exit" -foreground Yellow
				 
				$choose = Read-Host "Please choose the number "

				switch ($choose) 
				{ 
						1{ Check_users_status } 
						2{ Change_Password  }
						3{ Send_mail }
						4{ See_mail_log }
						5{ Create_user }
						7{ Remove_user}
						8{ Recover_user}
						9{ Remove-Module MSOnline;
							Login 
						}
						default { $i = 0; }
				}
				if($choose -ne 9){
					Pause
				}
		}
		Remove-Module MSOnline
####Main  End####