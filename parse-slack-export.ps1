$exportPath = "C:\Users\Daniel\Downloads\slack\" #this should be were your extracted your slack export

$dms = Get-Content ($exportPath + "dms.json") | ConvertFrom-Json
$users = Get-Content ($exportPath + "users.json") | ConvertFrom-Json
$hash = @{}

#region Unused API stuff
<#
Was using the APIs for info before I noticed the files in the root export folder.
$token = Get-Content "C:\Users\Daniel\Desktop\token.txt"
$userListAPI = "https://slack.com/api/users.list?token=" + $token + "&pretty=1"
$teamListAPI = "https://slack.com/api/team.list?token=" + $token + "&pretty=1"
$global:users = Invoke-RestMethod -Method Post -Uri $userListAPI
$global:teams = Invoke-RestMethod -Method Post -Uri $teamListAPI
#>
#endregion

#region Functions
function GetUserRealName($id) {
    $user = $global:users.members.Where({$_.id -eq $id})
    return $user.profile.real_name
}

function ConvertUnixTime($timestamp) {
    $date = (Get-Date 01.01.1970).AddSeconds($timestamp)
    return $date
}
#endregion

#region Rename DM folders
foreach($dm in $dms){
    $user1 = GetUserRealName($dm.members[0])
    $user2 = GetUserRealName($dm.members[1])
    $names = $user1 + "-" + $user2
    $hash.Add($dm.id,$names)
}

foreach($item in $hash.keys) {
    $path = $exportPath + $item
    $newPath = $exportPath + $item + "-" + $hash.$item
    $newPath = $newPath.Replace("contractor","con")
    $newPath = $newPath.Replace("contracto","con")
    $newPath = $newPath.Replace("contract","con")
    $newPath = $newPath.Replace("contrac","con")
    $newPath = $newPath.Replace("contra","con")
    $newPath = $newPath.Replace("contr","con")
    $newPath = $newPath.Replace("_cont","_con")
    if(Test-Path $path) {
        Rename-Item -Path $path -NewName $newPath
    }
}
#endregion

#region Combine and convert JSON to HTML
$directories = Get-ChildItem $exportPath

foreach($directory in $directories) {
    if(Get-ChildItem $directory -Filter *.html) {
        Write-Output "Skipping " $directory
    } else {
        $htmlFile = $directory.FullName + "\all.html"
        New-Item -Force $htmlFile
        $jsonFiles = Get-ChildItem $directory -Exclude *.html
        foreach($jsonFile in $jsonFiles) {
            $jsonObjs = Get-Content $jsonFile.FullName | ConvertFrom-Json
            $hash = @{}
            $hash.Add("a","a")
            foreach($jsonObj in $jsonObjs) {
                $message = $null
                $timestamp = $null
                $user = $null
                $message = $null
                $title = $null
                $timestamp = ConvertUnixTime $jsonObj.ts
                $user = GetUserRealName $jsonObj.user
                if ($user -eq $null) {
                    $user = $jsonObj.username
                }
                #  If the message is an Edit append Edit and bold it.
                $message = $jsonObj.text
                if ($message -eq $null) {
                    $message = "<b>Edit</b> " + $jsonObj.message.text
                }
                $title = $jsonObj.attachments.title
                $text = $timestamp.ToString() + " " + $user + " " + $message + " " + $title + "<br>"
                $p = '<@\w+>'
                [regex]$regex = '<@\w+>'
                $matchValue = ([regex]::matches($text, $p) | %{$_.value})
                foreach($match in $matchValue) {
                    $length = $match.Length - 3
                    $replacement = "@" + (GetUserRealName $match.substring(2,$length))
                    $text = $text.Replace($match,$replacement)
                }
                Add-Content -Path $htmlFile -Value $text |Out-Null
            }
        }
    }
}
#endregion
