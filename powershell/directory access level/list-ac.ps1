param (
    [Parameter(Mandatory = $true)]
    [string]$path,
    [switch]$report,
    [switch]$showInheritedPerms
)

function FileSysRightText($sysRights) {
    $t = $sysRights.ToString()
    if ([System.Text.RegularExpressions.Regex]::IsMatch($t, "^[-0-9]")) {
        return "Special"
    }
    return $t
}
function Bool2YN($bool) {
    if ($bool) { return "Y" }
    return "N"
}

function GetAclRecord([string]$path, [bool]$incInheritedPerms) {
    Write-Progress -Activity "$path"
    $acl = Get-Acl $path
    $access = $acl.Access
    if (!$incInheritedPerms) {
        $access = $access | Where-Object { !$_.IsInherited }
    }
    if ($access.Count -gt 0) {
        [PSCustomObject]@{
            Path   = $path;
            Access = ($access | ForEach-Object {
                    "$($_.IdentityReference)`t$($_.AccessControlType.ToString()[0])`t$(FileSysRightText $_.FileSystemRights)`t$(Bool2YN $_.IsInherited)"
                })
        }
    }
}

Clear-Host

$list = New-Object System.Collections.ArrayList

$list.Add((GetAclRecord $path $true)) | Out-Null
Get-ChildItem -Path $path -Directory -Recurse | ForEach-Object {
    $list.Add((GetAclRecord $_.FullName $showInheritedPerms)) | Out-Null
}

Clear-Host

if ($report) {
@"
    <!DOCTYPE html>
    <html><head><title>Permission Checklist for $path</title>
    <style>
        html,body,table { font-size: 10pt; font-family: 微軟正黑體; }
        table { border-collapse: collapse; }
        thead td { text-align: center; background-color: royalblue; color: white; }
        td { padding: 3px; border: 1px solid #bbb; }
        td[rowspan] { vertical-align: top; }
        td.Y { color: gray; }
        td.N { color: blue; }
    </style>
    </head><body><h3>$path Permission Checklist</h3>
    <table><thead><tr><td>Path</td><td>User</td><td>Rights</td></tr></thead>
    <tbody>
"@
    $list | Where-Object { $_.Path } | ForEach-Object {
        $acl = $_
        $first = $true
        $acl.Access | ForEach-Object {
            "<tr>"
            if ($first) {
                "<td rowspan=$($acl.Access.Count)>$($acl.Path)</td>"
                $first = $false
            }
            $p = $_.Split("`t")        
            "<td class='$($p[1]) $($p[3])'>$($p[0])</td><td class='$($p[1]) $($p[3])'>$($p[2])</td>"
            "</tr>"
        }
    }
    "</tbody></table></body></html>"
}
else {
    Write-Output $list
}