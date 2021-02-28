#  Parse Slack exports into html files.

$exportPath = "slack\" #this should be were your extracted your slack export

$pictures = $true #put the picutures in the HTML
$overwrite = $true #overwrite existing html files
$rebuildhash = $true #rebuild the dm hash even if it's not empty
$file_pictures = $true # Pictures from Files Array.
$build_index = $true # Build an HTML Index with Links to Results

Clear-Host
$directories = $null

$dms = Get-Content -Encoding "UTF8" ($exportPath + "dms.json") | ConvertFrom-Json
$Global:users = Get-Content -Encoding "UTF8" ($exportPath + "users.json") | ConvertFrom-Json
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
    $hashNames = $hash.$item
    [System.IO.Path]::GetInvalidFileNameChars() | % {$hashNames = $hashNames.replace($_,' ')}
    $newPath = $exportPath + $item + "-" + $hashNames
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

if ($build_index) {
	$menuHTMLFile = ".\menu.html"
	New-Item -Force $menuHTMLFile
}
foreach($directory in $directories) {
    #  Skip folder that already have an html file in them
    $htmlExists = Get-ChildItem $directory.FullName -Filter *.html 
    if($htmlExists -and ($htmlExists.Length -ne 0) -and $overwrite -eq $false) {
        Write-Output "Skipping " $directory
    } else {
        $htmlFile = $directory.FullName + "\all.html"
        New-Item -Force $htmlFile
		
		if ($build_index) {
			# Master HTML File
			Write-Output "Dir: " + $menuHTMLFile
			Write-Output $directory.Name
			$link = "<a href=`"/" + $htmlFile + "`" target=`"iframe_main`">#" + $directory.Name + "</a><br>"
			Add-Content -Path $menuHTMLFile -Value $link |Out-Null
		}
		
        $jsonFiles = Get-ChildItem $directory.FullName -Exclude *.html
        foreach($jsonFile in $jsonFiles) {
            $jsonObjs = Get-Content -Encoding "UTF8" $jsonFile.FullName | ConvertFrom-Json
            foreach($jsonObj in $jsonObjs) {
                $message = $null
                $timestamp = $null
                $user = $null
                $message = $null
                $title = $null
                $userPicture = $null
				$filePictures = $null
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
                        $userPicture = "<img src=`"$userPicture`" style=`"padding-right: 10px; padding-left: 10px`">"
                    } else {
                        $userPicture = $null
                    }
                }
				if ($file_pictures) {
					if ($jsonObj.files) {
						foreach ($file_picture in $jsonObj.files) {
							$filePictures = -join($filePictures, " ", "<a href=`"" + $file_picture.url_private + "`" target=`"_blank`"><img src=`"" + $file_picture.thumb_480 + "`"></a>")
						}
					}
				}
                #  If the message is an Edit append Edit and bold it.
                $message = $jsonObj.text
                if ($message -eq $null) {
                    $message = "<b>Edit</b> " + $jsonObj.message.text
                }
                $title = $jsonObj.attachments.title
                $text = $userPicture + $timestamp.ToString() + " " + $user + " " + $message + " " + $title + "<br>"
				if ($filePictures) {
					$text = -join($text, "<br>", $filePictures + "<br>")
				}
                $p = '<@\w+>'
                [regex]$regex = '<@\w+>'
                $matchValue = ([regex]::matches($text, $p) | %{$_.value})
                foreach($match in $matchValue) {
                    $length = $match.Length - 3
                    $replacement = "@" + (GetUserRealName $match.substring(2,$length))
                    $text = $text.Replace($match,$replacement)
                }
				$text = "<span style=`"text-align: center;`">" + $text + "</span>"
                Add-Content -Path $htmlFile -Value $text |Out-Null
            }
        }
    }
}
#endregion
