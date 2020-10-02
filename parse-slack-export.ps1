#  Parse Slack exports into html files.

$exportPath = "C:\Users\Daniel\Downloads\slack\" #this should be were your extracted your slack export

$pictures = $false #put the picutures in the HTML
$overwrite = $false #overwrite existing html files
$rebuildhash = $false #rebuild the dm hash even if it's not empty

Clear-Host
$directories = $null

$dms = Get-Content ($exportPath + "dms.json") | ConvertFrom-Json
$Global:users = Get-Content ($exportPath + "users.json") | ConvertFrom-Json
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
    $user = $Global:users.Where({$_.id -eq $id})
    return $user.profile.real_name
}

function GetUserPicture($id) {
    $user = $Global:users.Where({$_.id -eq $id})
    return $user.profile.image_24
}

function ConvertUnixTime($timestamp) {
    $date = (Get-Date 01.01.1970).AddSeconds($timestamp)
    return $date
}

function RenameDirectories($path,$newPath) {
    if((Test-Path $path) -and $newPath) {
        $newPath = $newPath.Replace("_contractor","_c")
        $newPath = $newPath.Replace("_contracto","_c")
        $newPath = $newPath.Replace("_contract","_c")
        $newPath = $newPath.Replace("_contrac","_c")
        $newPath = $newPath.Replace("_contra","_c")
        $newPath = $newPath.Replace("_contr","_c")
        $newPath = $newPath.Replace("_cont","_c")
        $newPath = $newPath.Replace("--","-")
        $newPath = $newPath.Replace("/","-")
        if($path -ne $newPath) {
            Rename-Item -Path $path -NewName $newPath
        }
    }
}
#endregion

#region Rename DM folders
Write-Output "Building DM user hash--this could take a while."

$hash = @{}
foreach($dm in $dms){
    $dmUser1 = GetUserRealName $dm.members[0]
    $dmUser2 = GetUserRealName $dm.members[1]
    $dmUsers = $dmUser1 + "-" + $dmUser2
    $hash.Add($dm.id,$dmUsers)
}

Write-Output "Renaming DM folders."
foreach($item in $hash.keys) {
    $path = $exportPath + $item
    $newPath = $exportPath + $item + "-" + $hash.$item
    RenameDirectories $path $newPath
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
$directories = Get-ChildItem $exportPath | ?{ $_.PSIsContainer }

foreach($directory in $directories) {
    #  Skip folder that already have an html file in them
    $htmlExists = Get-ChildItem $directory.FullName -Filter *.html 
    if($htmlExists -and ($htmlExists.Length -ne 0) -and $overwrite -eq $false) {
        Write-Output "Skipping " $directory
    } else {
        $htmlFile = $directory.FullName + "\all.html"
        New-Item -Force $htmlFile
        $jsonFiles = Get-ChildItem $directory.FullName -Exclude *.html
        foreach($jsonFile in $jsonFiles) {
            $jsonObjs = Get-Content $jsonFile.FullName | ConvertFrom-Json
            foreach($jsonObj in $jsonObjs) {
                $message = $null
                $timestamp = $null
                $user = $null
                $message = $null
                $title = $null
                $userPicture = $null
                $timestamp = ConvertUnixTime $jsonObj.ts
                $user = $jsonObj.user_profile.real_name
                if ($user -eq $null) {
                    $user = GetUserRealName $jsonObj.user
                }
                if ($user -eq $null -or $user -eq "") {
                    $user = $jsonObj.username
                }
                if($pictures) {
                    $userPicture = GetUserPicture $jsonObj.user
                    if($userPicture) {
                        $userPicture = "<img src=`"$userPicture`">"
                    } else {
                        $userPicture = $null
                    }
                }
                #  If the message is an Edit append Edit and bold it.
                $message = $jsonObj.text
                if ($message -eq $null) {
                    $message = "<b>Edit</b> " + $jsonObj.message.text
                }
                $title = $jsonObj.attachments.title
                $text = $timestamp.ToString() + " " + $user + $userPicture + " " + $message + " " + $title + "<br>"
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
