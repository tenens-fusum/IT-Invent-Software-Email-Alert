########################################################################################################################################################
#
#
#  Email:  		    me@andreevti.ru
#  GIT:             https://github.com/tenens-fusum/Windows-FileServer-Monitoring
#  Date created:   	26/04/2019
#  Version: 		1.0
#  Description: 	Скрипт создает Email оповещение об истечении срока действия лицензий на приложения в разделе "Программы" ПО IT-Invent.  
#                   Поле "ответственный за ПО" создается вручную в ПО ИТ-Инвент (Справочники - Программы - Свойства) , в качестве аргумента запроса в поле "значения списка" указывается #employees.
#
########################################################################################################################################################


#DB Parametrs
$DBServerName = "DBServerName"
$DBNAme = "ITInvent"

#SMTP parametrs
$EmailFrom = "Alert@domain.ru"
$SMTPServer = "smtp.domain.ru"

#Parametrs
$AlarmAfter = "30"                                          #Количество дней до истечения срока действия лицензии, после которых высылаются уведомления
$SentToManagerForAllSoftware = "Yes"                        #Параметр, указывающий высылать ли агрегированные уведомления на "ManagerForAllSoftware" или нет. Может принимать значения Yes или No
$SentOnlyToManagerForAllSoftware = "No"                     #Параметр, указывающий высылать ли персональные уведомления на ответственного за ПО или нет. Может принимать значения Yes или No
$ManagerForAllSoftware = "Software-manager@domain.ru"       #Пользователь, контролирующий все лицензии на ПО. Получает агрегированные уведомления

#Design for Email table 
 $Header = '
                        <style>
                        TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
                        TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED; text-align:center; color:white}
                        TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black; text-align:center}
                        </style>
            '



$Programs = Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select * From [$DBNAme].[dbo].[ITEMS] Where LICENCE_DATE is not NULL"

$AllExpiredSoftware = @()
Foreach ($Program in $Programs)
{
    $LicenseEndDate = $Program.LICENCE_DATE
    $LicenseEndDateN = $Program.LICENCE_DATE.ToString("dd.MM.yyyy")
    $NowDate = Get-Date
    $TypeNO = $Program.Type_NO
    $ModelNO = $Program.MODEL_NO
    $PrID = $Program.ID
    $SerialNumber = $Program.SERIAL_NO
    $DaysToExpire = [math]::Round((New-TimeSpan -Start $NowDate -End $LicenseEndDate).TotalDays, 0)
    $SoftwareName = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select Type_Name,Vendor_NO From [$DBNAme].[dbo].[CI_TYPES] Where TYPE_NO = '$TYPENO' AND CI_Type = 2").Type_Name
    $VendorNO = $SoftwareName.Vendor_NO
    $SoftwareModel = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select Model_Name From [$DBNAme].[dbo].[CI_MODELS] Where MODEL_NO = '$MODELNO' AND CI_Type = 2").Model_Name
    $VendorName = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select Vendor_Name From [$DBNAme].[dbo].[VENDORS] Where Vendor_NO = 42").Vendor_Name
    $Manager = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select Field_Value From [$DBNAme].[dbo].[FIELDS_VALUES] Where Item_ID = '$SerialNumber'").Field_Value   
    $LastChangeBy = $Program.CH_User
    $SerialNumber = $Program.SERIAL_NO
    $ExpiredLicenses = @()

       If (($ManagerEmail = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select Owner_Email From [$DBNAme].[dbo].[Owners] Where Owner_Display_Name = '$Manager'").Owner_Email) -notlike "[a-z0-9!#$%&'*+/=?^_{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_{|}~-]+)*@(?:a-z0-9?.)+a-z0-9?")      
          {
           Write-Host $ManagerEmail
           $ManagerEmail = (Invoke-Sqlcmd -ServerInstance $DBServerName -Database "$DBNAme" -Query "Select User_Name,User_Email From [$DBNAme].[dbo].[Users] Where User_Name = '$LastChangeBy'").User_Email
           }
           Else
                {
                $ManagerEmail = $Manager.Owner_Email
                }
        
        If (((New-TimeSpan -Start $NowDate -End $LicenseEndDate).TotalDays) -le $AlarmAfter)
           {
              $ExpiredSoftware = @()
              $ExpiredSoftware = @([PSCustomObject]@{SoftwareName = "$SoftwareName"; SerialNumber = "$SerialNumber";  SoftwareModel = "$SoftwareModel";"Program ID" = "$PrID"; VendorName = "$VendorName"; Expire_Date = "$LicenseEndDateN"; "Days to expire" = $DaysToExpire; "Manager" = $Manager})
              $AllExpiredSoftware += $ExpiredSoftware
              If ($SentOnlyToManagerForAllSoftware -contains "No")
                   {
                                 $Mailbody = '<p style="font-family: Verdana; font-size: 12pt;">Срок действия лицензии истекает в ближайшие ' + $AlarmAfter + ' дней</p>'
                                 $ExpiredSoftwareTable = $ExpiredSoftware | ConvertTo-Html -property @{label="ИД программы";expression={$($_."Program ID")}},@{label="Серийный номер ПО";expression={$($_."SerialNumber")}},@{label="Разработчик";expression={$($_."VendorName")}},@{label="Название";expression={$($_."SoftwareName")}},@{label="Версия";expression={$($_."SoftwareModel")}},@{label="Дата окончания действия лицензии";expression={$($_."Expire_Date")}},@{label="Осталос дней до окончания действия лицензии";expression={$($_."Days To Expire")}},@{label="Ответственный за ПО";expression={$($_."Manager")}} -Head $Header
                                 $Mailbody = $Mailbody + $ExpiredSoftwareTable
                                 Send-MailMessage -To "$ManagerEmail" -From $EmailFrom -SmtpServer $SMTPServer -Subject "test" -Body $Mailbody -BodyAsHtml -Encoding unicode
                    }
                    

                    }
        }

If ($SentToManagerForAllSoftware -contains "Yes")
    {
     $Mailbody = '<p style="font-family: Verdana; font-size: 12pt;">Срок действия лицензий истекает в ближайшие ' + $AlarmAfter + ' дней</p>'
       $AllExpiredSoftwareTable = $AllExpiredSoftware | ConvertTo-Html -Property @{label="ИД программы";expression={$($_."Program ID")}},@{label="Серийный номер ПО";expression={$($_."SerialNumber")}},@{label="Разработчик";expression={$($_."VendorName")}},@{label="Название";expression={$($_."SoftwareName")}},@{label="Версия";expression={$($_."SoftwareModel")}},@{label="Дата окончания действия лицензии";expression={$($_."Expire_Date")}},@{label="Осталос дней до окончания действия лицензии";expression={$($_."Days To Expire")}},@{label="Ответственный за ПО";expression={$($_."Manager")}} -Head $Header
       $Mailbody = $Mailbody + $AllExpiredSoftwareTable
       Send-MailMessage -To $ManagerForAllSoftware -From $EmailFrom -SmtpServer $SMTPServer -Subject "test" -Body $Mailbody -BodyAsHtml -Encoding unicode
     }
