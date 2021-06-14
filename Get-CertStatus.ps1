#############################################################################
# Author  : Tyler Cox
# 
# Special Thanks: David Jones for his function: https://www.powershellgallery.com/packages/PKITools/1.6
#
# Version : 1.2
# Created : 06/01/2021
# Modified : 06/14/2021
#
# Purpose : This script queries all certs from the CA and publishes it to a webpage as well as emails
#           the results of expiration days
#
# Requirements: None
#             
# Change Log:   Ver 1.2 - Added 5 day warning. Updated warning colors to include new 5 day category. 
#
#               Ver 1.1 - Added ability to send email summary only on the first of the month,
#                         but to also send email on odd days for certs expiring in 15 days or less
#
#               Ver 1.0 - Initial release
#
#############################################################################

#Declare some variables
$CurrentMonth = Get-Date -UFormat %b #Current month abbreviated
$CurrentYear = (Get-Date).year #Get the year
$CurrentDay = (Get-Date).Day #Get the day
$CurrentDayofWeek = (Get-Date).DayofWeek #Get the day of the week
[string]$emailSubject = "$($CurrentMonth)-$($CurrentYear) - Certificates Status"
$table = $null #null out our variable
$html = $null #null out our variable
$5dayCerts = 0 #Count used for emailing
$15dayCerts = 0 #Count used for emailing
$30dayCerts = 0 #Count used for emailing
$60dayCerts = 0 #Count used for emailing

#Variables to Edit
[string]$emailFrom = ""  #This is who the email address will be from. i.e. DoNotReply@contoso.com
[string]$emailTo = ""  #This is who the email will be sent to
[string]$emailSMTPserver = ''  #This is the SMTP Server
$pageURL = "" #This will be the DNS name of the local webpage URL. i.e. https://CertStatus.contoso.com
$IISLocation = "C:\inetpub\certstatus\" #This is the location to the IIS folder 
$CAServerName = ''  #The name of the server that hosts the Certificate Authority
$CAname = '' #The name of the Certificate Authority instance.

#get our cert expiry dates
$warninglowdate = (Get-Date).AddDays(60)
$warninghighdate = (Get-Date).AddDays(30)
$warningseveredate = (Get-Date).AddDays(15)
$warningcriticaldate = (Get-Date).AddDays(5)

#Get-IssuedCertificate from: https://www.powershellgallery.com/packages/PKITools/1.6
function Get-IssuedCertificate 
{
    <#
        .SYNOPSIS
        Get Issued Certificate data from one or more certificate athorities.
 
        .DESCRIPTION
        Can get various certificate fileds from the Certificate Authority database. Usfull for exporting certificates or checking what is about to expire
 
        .PARAMETER ExpireInDays
        Maximum number of days from now that a certificate will expire. (Default: 21900 = 60 years) Can be a negative numbe to check for recent expirations
 
        .PARAMETER CAlocation
        Certificate Authority location string "computername\CAName" (Default gets location strings from Current Domain)
 
        .PARAMETER Properties
        Fields in the Certificate Authority Database to Export
 
        .PARAMETER CertificateTemplateOid
        Filter on Certificate Template OID (use Get-CertificateTemplateOID)
 
        .PARAMETER CommonName
        Filter by Issued Common Name
 
        .EXAMPLE
        Get-IssuedCertificate -ExpireInDays 14
        Gets all Issued Certificates Expireing in the next two weeks
 
        .EXAMPLE
        Get-IssuedCertificate -ExpireInDays -7
        Gets all Issued Certificates that Expired last week
 
        .EXAMPLE
        Get-IssuedCertificate -CAlocation CA1\MyCA
        Gets all Certificates Issued by CA1
 
        .EXAMPLE
        Get-IssuedCertificate -Properties 'Issued Common Name', 'Certificate Hash'
        Gets all Issued Certificates and outputs only the Common name and thumbprint
 
        .EXAMPLE
        Get-IssuedCertificate -CommonName S1, S2.contoso.com
        Gets Certificats issued to S1 and S2.contoso.com
 
        .EXAMPLE
        $DSCCerts = Get-IssuedCertificate -CertificateTemplateOid (Get-CertificateTemplateOID -Name 'DSCTemplate') -Properties 'Issued Common Name', 'Binary Certificate'
        foreach ($cert in $DSCCerts)
        {
            set-content -path "c:\certs\$($cert.'Issued Common Name').cer" -Value $cert.'Binary Certificate' -Encoding Ascii
        }
        Get all certificates issued useing the DSCTemplate template and save them to the folder c:\certs named for the Common name of the certificate
 
   #>

 
    [CmdletBinding()]
    Param (
        
        # Maximum number of days from now that a certificate will expire. (Default: 21900 = 60 years)
        [Int]
        $ExpireInDays = 21900,
        
        # Certificate Authority location string "computername\CAName" (Default gets location strings from Current Domain)
        [String[]]
        $CAlocation = (get-CaLocationString),

        # Fields in the Certificate Authority Database to Export
        [String[]]
        $Properties = (
            'Issued Common Name', 
            'Certificate Expiration Date', 
            'Certificate Effective Date', 
            'Certificate Template', 
            #'Issued Email Address',
            'Issued Request ID', 
            'Certificate Hash', 
            #'Request Disposition',
            'Request Disposition Message', 
            'Requester Name', 
            'Binary Certificate' ),

        # Filter on Certificate Template OID (use Get-CertificateTemplateOID)
        [AllowNull()]
        [String]
        $CertificateTemplateOid,

        # Filter by Issued Common Name
        [AllowNull()]
        [String]
        $CommonName
    ) 
    
    foreach ($Location in $CAlocation) 
    {
        $CaView = New-Object -ComObject CertificateAuthority.View
        $null = $CaView.OpenConnection($Location)
        $CaView.SetResultColumnCount($Properties.Count)
    
        #region SetOutput Colum
        foreach ($item in $Properties)
        {
            $index = $CaView.GetColumnIndex($false, $item)
            $CaView.SetResultColumn($index)
        }
        #endregion

        #region Filters
        $CVR_SEEK_EQ = 1
        $CVR_SEEK_LT = 2
        $CVR_SEEK_GT = 16
    
        #region filter expiration Date
        $index = $CaView.GetColumnIndex($false, 'Certificate Expiration Date')
        $now = Get-Date
        $expirationdate = $now.AddDays($ExpireInDays)
        if ($ExpireInDays -gt 0)
        { 
            $CaView.SetRestriction($index,$CVR_SEEK_GT,0,$now)
            $CaView.SetRestriction($index,$CVR_SEEK_LT,0,$expirationdate)
        }
        else 
        {
            $CaView.SetRestriction($index,$CVR_SEEK_LT,0,$now)
            $CaView.SetRestriction($index,$CVR_SEEK_GT,0,$expirationdate)
        }
        #endregion filter expiration date

        #region Filter Template
        if ($CertificateTemplateOid)
        {
            $index = $CaView.GetColumnIndex($false, 'Certificate Template')
            $CaView.SetRestriction($index,$CVR_SEEK_EQ,0,$CertificateTemplateOid)
        }
        #endregion

        #region Filter Issued Common Name
        if ($CommonName)
        {
            $index = $CaView.GetColumnIndex($false, 'Issued Common Name')
            $CaView.SetRestriction($index,$CVR_SEEK_EQ,0,$CommonName)
        }
        #endregion

        #region Filter Only issued certificates
        # 20 - issued certificates
        $CaView.SetRestriction($CaView.GetColumnIndex($false, 'Request Disposition'),$CVR_SEEK_EQ,0,20)
        #endregion

        #endregion

        #region output each retuned row
        $CV_OUT_BASE64HEADER = 0 
        $CV_OUT_BASE64 = 1 
        $RowObj = $CaView.OpenView() 

        while ($RowObj.Next() -ne -1)
        {
            $Cert = New-Object -TypeName PsObject
            $ColObj = $RowObj.EnumCertViewColumn()
            $null = $ColObj.Next()
            do 
            {
                $displayName = $ColObj.GetDisplayName()
                # format Binary Certificate in a savable format.
                if ($displayName -eq 'Binary Certificate') 
                {
                    $Cert | Add-Member -MemberType NoteProperty -Name $displayName -Value $($ColObj.GetValue($CV_OUT_BASE64HEADER)) -Force
                } else 
                {
                    $Cert | Add-Member -MemberType NoteProperty -Name $displayName -Value $($ColObj.GetValue($CV_OUT_BASE64)) -Force
                }
            }
            until ($ColObj.Next() -eq -1)
            Clear-Variable -Name ColObj

            $Cert
        }
    }
}

#Get all certs from the CA 
$AllCerts = Get-IssuedCertificate -CAlocation ($CAServerName + '\' + $CAname) -Properties 'Issued Common Name', 'Certificate Expiration Date', 'Certificate Effective Date', 'Certificate Template', 'Requester Name', 'Certificate Hash'

#Parse by cert templates [Edit this if you want different templates to show up]
$MonitoredCerts = $AllCerts | Where {$_.'Certificate Template' -eq 'ACS' -OR $_.'Certificate Template' -eq 'ADFSCertificate' -OR $_.'Certificate Template' -eq 'CodeSigning' -OR $_.'Certificate Template' -eq 'DirectAccess' -OR $_.'Certificate Template' -eq 'EFSRecovery' -OR $_.'Certificate Template' -eq 'GeneralCodeSigning' -OR $_.'Certificate Template' -eq 'SubCA' -OR $_.'Certificate Template' -eq 'HospiraMedNet' -OR $_.'Certificate Template' -eq 'SCCMWebServerCertificate' -OR $_.'Certificate Template' -eq 'WebServer'}

#Loop through the certs to set the status of each one as well as get our counts
foreach ($Cert in $MonitoredCerts)
    {
        If ($Cert.'Certificate Expiration Date' -gt $warninglowdate)
            {
                #$Status = "Good - Cert expires in more than 60 days!"
                $CertStatus = '<a href=" " title="Good - Cert expires in more than 60 days!"><img src="images/checks/green.png" alt="Green - Good">' 
            }
        If ($Cert.'Certificate Expiration Date' -le $warninglowdate -AND $Cert.'Certificate Expiration Date' -gt $warninghighdate)
            {
                #$Status = "Warning Low - Cert expires in less than 60 days!"
                $60dayCerts += 1 #Add to our count
                $CertStatus = '<a href=" " title="Warning Low - Cert expires in less than 60 days!"><img src="images/checks/blue.png" alt="blue - Low">'
            }
        If ($Cert.'Certificate Expiration Date' -le $warninghighdate -AND $Cert.'Certificate Expiration Date' -gt $warningseveredate)
            {
                #$status = "Warning High - Cert expires in less than 30 days!"
                $30dayCerts += 1 #Add to our count
                $CertStatus = '<a href=" " title="Warning High - Cert expires in less than 30 days!"><img src="images/checks/Yellow.png" alt="Yellow - High">'
            }
        If ($Cert.'Certificate Expiration Date' -le $warningseveredate -AND $Cert.'Certificate Expiration Date' -gt $warningcriticaldate)
            {
                #$status = "Warning Severe - Cert expires in less than 15 days!"
                $15dayCerts += 1 #Add to our count
                $CertStatus = '<a href=" " title="Warning Severe - Cert expires in less than 15 days!"><img src="images/checks/Orange.png" alt="Orange - Severe">'
            }
        If ($Cert.'Certificate Expiration Date' -le $warningcriticaldate)
            {
                #$status = "Warning Critical - Cert expires in less than 5 days!"
                $5dayCerts += 1 #Add to our count
                $

        #Build our html table
        $table = '
        <tr class="row100">
            <td class="column100 column2" data-column="column2">'+ $Cert.'Issued Common Name' +'</td>
            <td class="column100 column3" data-column="column3">'+ $Cert.'Certificate Hash' +'</td>
            <td class="column100 column4" data-column="column4">'+ $Cert.'Certificate Effective Date' +'</td>
            <td class="column100 column5" data-column="column5">'+ $Cert.'Certificate Expiration Date'+'</td>
            <td class="column100 column6" data-column="column6">'+ $Cert.'Certificate Template' +'</td>
            <td class="column100 column8" data-column="column7">'+ $Cert.'Requester Name'+'</td>
            <td class="column100 column8" data-column="column8">'+ $CertStatus +'</td>
        </tr>'

        $html += $table

    }

#Get the template file
$template = (Get-Content -Path ($IISLocation + "template.html") -raw)
#Place variables and new $html into the template file and rename it as index.html
Invoke-Expression "@`"`r`n$template`r`n`"@" | Set-Content -Path ($IISLocation + 'index.html')

#Handle the emails
If ($CurrentDay -eq 1) #Send out the monthly email on the first of the month
    {
        [string]$emailBody = "<p><span style='color: red;'><strong>Certificate Status Monthly Update</strong></span><br /><br />Please use the link below to review the Certificates Status for this month.<br /><br />$($PageURL)<br /><br />Certs Expiring in 5 days or less:<span style='color: red;'><strong> $5dayCerts</strong></span><br />Certs Expiring in 15 days or less:<span style='color: Orange;'><strong> $15dayCerts</strong></span><br />Certs Expiring in 30 days or less:<span style='color: GOld;'><strong> $30dayCerts</strong></span><br />Certs Expiring in 60 days or less:<span style='color: dodgerblue;'><strong> $60dayCerts</strong></span><br /><br />This email is sent automatically.</p>"
        Send-MailMessage -To $emailTo -From $emailFrom -BodyAsHtml $emailBody -Subject $emailSubject -SmtpServer $emailSMTPserver #-Priority High    
    }
ElseIf (($CurrentDay -ne 1) -AND ($15dayCerts -ge 1) -AND ($CurrentDayofWeek -eq "Monday"))#Send a severe warning email if we detected one or more certs expiring in 15 days AND it's a monday as well as not the first day of the month. This is done so we don't get bombarded by cert emails.
    {
        [string]$emailBody = "<p><span style='color: red;'><strong>Severe Warning! Cert(s) expiring in 15 days or less!</strong></span><br /><br />Please use the link below to identify the cert(s) expiring:<br /><br />$($PageURL)<br /><br />Certs Expiring in 15 days or less:<span style='color: red;'><strong> $15dayCerts</strong></span><br />Certs Expiring in 30 days or less:<span style='color: orange;'><strong> $30dayCerts</strong></span><br />Certs Expiring in 60 days or less:<span style='color: dodgerblue;'><strong> $60dayCerts</strong></span><br /><br />This email is sent automatically.</p>"
        Send-MailMessage -To $emailTo -From $emailFrom -BodyAsHtml $emailBody -Subject "Certificate Status - Severe Warning" -SmtpServer $emailSMTPserver -Priority High    
    }
ElseIf (($CurrentDay -ne 1) -AND ($5dayCerts -ge 1) -AND ($CurrentDayofWeek -ne "Monday"))#Send a Critical warning email if we detected one or more certs expiring in 5 days AND it's not the first day of the month or a Monday. This is done so we don't get bombarded by cert emails.
    {
        [string]$emailBody = "<p><span style='color: red;'><strong>Critical Warning! Cert(s) expiring in 5 days or less!</strong></span><br /><br />Please use the link below to identify the cert(s) expiring:<br /><br />$($PageURL)<br /><br />Certs Expiring in 15 days or less:<span style='color: red;'><strong> $15dayCerts</strong></span><br />Certs Expiring in 30 days or less:<span style='color: orange;'><strong> $30dayCerts</strong></span><br />Certs Expiring in 60 days or less:<span style='color: dodgerblue;'><strong> $60dayCerts</strong></span><br /><br />This email is sent automatically.</p>"
        Send-MailMessage -To $emailTo -From $emailFrom -BodyAsHtml $emailBody -Subject "Certificate Status - Critical Warning" -SmtpServer $emailSMTPserver -Priority High    
    }
