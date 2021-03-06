﻿#Script to automatically import mails into the ticket system (over Exchange)

$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path D:\essentials\exchange-output.txt -append

#Credentials file
. .\authority.ps1
#Path file
. .\conf.ps1

#Load assembly file
#We need to do this the fucked up way, because there is no chance in hell we'll a current version of Powershell on the server
#Location of Exchange WebServices API DLL File (needs to be at least Version 1.2)
[Reflection.Assembly]::LoadFile($path.exchangeassy)

#SHA2 512bit hash convenience function, returns string with base16 hash
function hash($m) {
    $hasher = [System.Security.Cryptography.SHA512]::Create()
    $a = "0x"
    $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($m)) | % { $a += $_.ToString("X2"); }
    return $a
}

function Query($q) {
    $connection = new-object system.data.sqlclient.sqlconnection
    $connection.ConnectionString = $cstring
    $connection.Open()

    $command = new-object system.data.sqlclient.sqlcommand
    $command.Connection = $connection
    $command.CommandText = $q
        
    $result = $command.ExecuteReader() 
    $table = new-object “System.Data.DataTable”
    $table.Load($result) 
    $connection.Close()
  
    return $table
}

$_folder = [Microsoft.Exchange.WebServices.Data.Folder]
$_wellknownfolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
$resolv = [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve

#Mailbox information
$email    = $authority.servicedeskmailbox.email
$username = $authority.servicedeskmailbox.username
$password = $authority.servicedeskmailbox.password
$domain   = $authority.servicedeskmailbox.domain

#Targeted category name
$catname = "IMPORT"

#SQL server connection string
$cstring = $authority.valuemationdb.connstr

#Instantiate service object, unparameterized defaults to Exchange 2010 mode
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

#Use first option if you want to impersonate, otherwise, grabs current user credentials
$s.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)
#$s.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#$s.UseDefaultCredentials = $true

#Discover the url from your email address
$s.AutodiscoverUrl($email)

#Get a handle to the inbox, mailbox root
$inbox = $_folder::Bind($s,$_wellknownfolder::Inbox)
$root = $_folder::Bind($s,$_wellknownfolder::MsgFolderRoot)

function findMailboxFolder($foldername) {
    return $root.FindFolders($(New-Object Microsoft.Exchange.WebServices.Data.FolderView(100))) | ? { $_.DisplayName -eq $foldername }
}

function getFolderContent($folder) {
    return $folder.FindItems($folder.TotalCount)
}

#Identify folder where processed tickets will be moved to
$target = findMailboxFolder("Aktenschrank")

#Create a property set (to let us access the body & other details not available from the FindItems call)
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;


#Message rules
#!! Order rules from more specific to less specific !!

#$_.Subject -match "(SR|IN)-[0-9]{7}" -and

$rules = @(
#    @{
#        Predicate = {
#            param($m)
#            return $($m.Sender.Name -eq "JOBCONTROL@SILHOUETTE.COM") -and $($m.Subject -eq "I5 QSYSOPR Nachricht unbeantwortet")
#        };
#        Action = {
#            param($m)
#            $hash = Hash $($_.From.Address + $_.Subject + $_.DateTimeReceived.tostring() + $_.Body.Text); #generate hash
#            Query "insert into X_SIL_INCOMINGMAIL values ('servicedesk@silhouette.com','Servicedesk','$($m.Subject + ' ' + $m.Body.Text)','','$($m.DateTimeReceived)',$hash)" #sql insert
#        }
#    };
    @{ 
        Predicate = { 
            param($m) 
            return $m.Categories.Contains($catname) 
        };
        Action = {
            param($m) 
            $hash = Hash $($_.From.Address + $_.Subject + $_.DateTimeReceived.tostring() + $_.Body.Text); #generate hash
            Query "insert into X_SIL_INCOMINGMAIL values ('$($m.From.Address)','$($m.From.Name)','$($m.Subject)','$($m.Body.Text)','$($m.DateTimeReceived)',$hash)" #sql insert
        }
    }
)

#Move through inbox items

#Replies for mail-in
Start-Sleep -s 5 | getFolderContent($inbox) |  % {
    foreach($rule in $rules) {
        if ( $rule.Predicate.Invoke($_)) {
        # Load the property set to allow us to get to the body
        $_.load($psPropertySet)
        $rule.Action.Invoke($_)
        write-host "> Item processed"
        $_.Categories.Clear() #clear mail categories
        $_.isRead = $true #clear unread status
        $_.Update($resolv) #update mail properties
        $_.Move($(findMailboxFolder("Aktenschrank")).Id); #drop in target folder
        }
    }
}

Stop-Transcript