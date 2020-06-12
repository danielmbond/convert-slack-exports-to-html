$exportPath = "C:\Users\Daniel\Downloads\slack\" #this should be were your extracted your slack export
Clear-Host
$directories = $null

$dms = Get-Content ($exportPath + "dms.json") | ConvertFrom-Json
$users = Get-Content ($exportPath + "users.json") | ConvertFrom-Json

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

function RenameDirectories($path,$newPath) {
    if((Test-Path $path) -and $newPath) {
        $newPath = $newPath.Replace("_contractor","_con")
        $newPath = $newPath.Replace("_contracto","_con")
        $newPath = $newPath.Replace("_contract","_con")
        $newPath = $newPath.Replace("_contrac","_con")
        $newPath = $newPath.Replace("_contra","_con")
        $newPath = $newPath.Replace("_contr","_con")
        $newPath = $newPath.Replace("_cont","_con")
        $newPath = $newPath.Replace("--","-")
        if($path -ne $newPath) {
            Rename-Item -Path $path -NewName $newPath
            $newPath
        }
    }
}
#endregion

#region Rename DM folders
Write-Output "Building DM user hash--this could take a while."
$hash = @{}
foreach($dm in $dms){
    $dmUser1 = GetUserRealName($dm.members[0])
    $dmUser2 = GetUserRealName($dm.members[1])
    $dmUsers = $dmUser1 + "-" + $dmUser2
    $hash.Add($dm.id,$dmUsers)
}
Write-Output "Renaming DM folders."
foreach($item in $hash.keys) {
    $path = $exportPath + $item
    $newPath = $exportPath + $item + "-" + $hash.$item
    RenameDirectories $path, $newPath

}
#  Rename mpdm folders
Write-Output "Renaming mpdm folders."
$directories = Get-ChildItem $exportPath

foreach($directory in $directories) {
    RenameDirectories $directory.fullname $directory.fullname
}
#endregion
$directories = $null
#region Combine and convert JSON to HTML
$directories = Get-ChildItem $exportPath

foreach($directory in $directories) {
    #  Skip folder that already have an html file in them
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
