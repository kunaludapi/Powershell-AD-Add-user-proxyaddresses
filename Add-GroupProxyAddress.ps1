<#  
  .Synopsis  
   Add smtp id to existing active directory Group proxyaddress.
  .Description  
   Run this script on domain controller. It will add addition record to proxy addresses in Group properties, and keep the existing as it is.
  .Example  
   Add-GroupProxyAddress -CSVFile c:\tenp\Group.csv
     
   It takes input from CSV file and add the smtp records in respective user proxy address attributes.
  .Example
   CSV file data format and example
   ----------------------------------------------
   | Group      | emailid                        |
   | --------------------------------------------
   | Group1    | Group1@vcloud-lab.com           |
   | Group2    | Group2@vcloud-lab.com           |
   ----------------------------------------------
  .OutPuts  
   GroupName ProxyAddresses
   -------- --------------
   Group1   {sip:Group1@testaccount.com, smtp:Group1@vcloud-lab.com}
   Group2   {sip:Group2@testaccount.com, smtp:Group2@vcloud-lab.com}
   
  .Notes  
   NAME: Add-GroupProxyAddress
   AUTHOR: Kunal Udapi
   CREATIONDATE: 01 DECEMBER 2016
   LASTEDIT: 3 February 2017  
   KEYWORDS: Add or update proxyaddress smtp on active directory Group account  
  .Link  
   #Check Online version: http://kunaludapi.blogspot.com
   #Check Online version: http://vcloud-lab.com
   #Requires -Version 3.0  
  #>  
#requires -Version 3   
[CmdletBinding()]
param(  
    [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true)]
    [alias('FilePath','File','CSV','CSVPath')]
    [String]$Path) #param
Begin {  
    Import-Module ActiveDirectory
} #Begin

Process {
    $Groups = Import-Csv -Path $Path
    #$Groups = Get-ADGroup -Filter * -SearchBase "OU=TestOu,DC=Rageframeworks,DC=com" -Properties ProxyAddresses

    Foreach ($u in $Groups) {
        #$smtpid = "smtp: {0}.{1}@kumarthegreat.com" -f $u.givenName, $u.Surname
        Try {
            $Group = Get-ADGroup -Identity $u.Group -ErrorAction Stop
            Write-Host "$($Group.SamAccountName) exists, Processing it..." -BackgroundColor DarkGray -NoNewline 
            $emailid = "SMTP:{0}" -f $u.emailid
            Set-ADGroup -Identity $u.Group -Add @{Proxyaddresses=$emailid} 
            #$cpemailid = "smtp:{0}" -f $u.cpemailid
            #Set-ADGroup -Identity $u.Group -Add @{Proxyaddresses=$cpemailid} 
            Write-Host "...ProxyAddress added" -BackgroundColor DarkGreen
        } #Try
        catch {
            Write-Host "$($Group.SamAccountName) does not exists" -BackgroundColor DarkRed
        } #catch
    } #foreach ($u in $Groups) 
    #Get-ADUser -Filter * -SearchBase "OU=TestOu,DC=Rageframeworks,DC=com" -Properties ProxyAddresses | select username, ProxyAddresses
    $TempFile = [System.IO.Path]::GetTempFileName()
    $Groups | foreach {
        $Group = $_.Group
        Try {
            Get-ADGroup -filter {Name -eq $Group} -Properties mail, proxyAddresses -ErrorAction Stop | select Name, Mail, GroupCategory, @{N='ProxyAddresses'; E={$($_.proxyAddresses -split ", ") -join "`n"}}
        } #try
        catch {
            Write-Host "$Group does not exists" -BackgroundColor DarkRed
        }
    } | Out-File $TempFile #foreach
} #Process
end {
    Get-Content -Path $TempFile
}
