#ad helper functions, needs at least (probably, unsure)
#> $host.version
#
#Major  Minor  Build  Revision
#-----  -----  -----  --------
#4      0      -1     -1      

function ad_name($name)
{
    $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()            
    $root = $dom.GetDirectoryEntry()            
    $search = [System.DirectoryServices.DirectorySearcher]$root            
    $search.filter = "(&(objectclass=user)(name=$name))"
    $a = @()           
    $search.findall() | %{ $a += [ADSI]$_.Path }
    return $a
}

function ad_group($name)
{
    $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()            
    $root = $dom.GetDirectoryEntry()            
    $search = [System.DirectoryServices.DirectorySearcher]$root            
    $search.filter = "(&(objectclass=group)(name=$name))"
    $a = @()           
    $search.findall() | %{ $a += $_ }
    return $a
}

function ad($filter)
{
    $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()            
    $root = $dom.GetDirectoryEntry()            
    $search = [System.DirectoryServices.DirectorySearcher]$root        
    $search.filter = $filter
    $a = @()           
    $search.findall() | %{ $a += $_ }
    return $a
}