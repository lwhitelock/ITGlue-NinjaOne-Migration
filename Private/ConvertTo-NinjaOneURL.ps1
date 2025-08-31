# This will be used to remake the ITGlue Links to NinjaOne.


$EscapedITGURL = [regex]::Escape($ITGURL)

if ($environmentSettings.ITGCustomDomains) {
    $combinedEscapedURLs = ($environmentSettings.ITGCustomDomains -split "," | ForEach-Object { [regex]::Escape($_) }) -join "|"
    $EscapedITGURL = "(?:$EscapedITGURL|$combinedEscapedURLs)"
}

$RichRegexPatternToMatchSansAssets = "<(A|a) href=\S$EscapedITGURL/([0-9]{1,20})/(docs|passwords|configurations)/([0-9]{1,20})\S.*?</(A|a)>"
$RichRegexPatternToMatchWithAssets = "<(A|a) href=\S$EscapedITGURL/([0-9]{1,20})/(assets)/.*?/([0-9]{1,20})\S.*?</(A|a)>"
$ImgRegexPatternToMatch = @"
$EscapedITGURL/([0-9]{1,20}/docs/([0-9]{1,20})/(images)/([0-9]{1,20}).*?)(?=")
"@
$RichDocLocatorUrlPatternToMatch = @"
<(A|a) href=\S$EscapedITGURL/(DOC-.*?)(?=")\S.*?</(A|a)>
"@
$RichDocLocatorRelativeURLPatternToMatch = @"
<(A|a) href=\S/(DOC-.*?)(?=")\S.*?</(A|a)>
"@

$TextRegexPatternToMatchSansAssets = "$EscapedITGURL/([0-9]{1,20})/(docs|passwords|configurations)/([0-9]{1,20})"
$TextRegexPatternToMatchWithAssets = "$EscapedITGURL/([0-9]{1,20})/(assets)/.*?/([0-9]{1,20})"
$TextDocLocatorUrlPatternToMatch = "$EscapedITGURL/(DOC-[0-9]{0,20}-[0-9]{0,20}).*(?= )"

function Update-StringWithCaptureGroups {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$inputString,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$pattern,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$type
    )
  
    $regex = [regex]::new($pattern)
    
    $matchesPattern = $regex.Matches($inputString)

    Write-Host "Found $($matchesPattern.count) matches to replace"
  
    foreach ($match in $matchesPattern) {

        # Compare the 3rd Group to identify where to find the new content

        switch ($match.groups[3].value) {

            "docs" {
                Write-Host "Found an $($match.groups[3].value) URL to replace for ITGID $($match.groups[4].value)..." -ForegroundColor 'Blue'
                $MatchedItem = $MatchedArticles | Where-Object { $_.ITGID -eq $match.groups[4].value }
                if (($MatchedItem | Measure-Object).count -eq 1) {

                    $N1Org = $($MatchedItem.NinjaOneObject.organizationId)
                    $N1PFolder = $($MatchedItem.NinjaOneObject.parentFolderId)
                    $N1ID = $MatchedItem.NinjaOneID

                    if ($Null -ne $N1Org) {
                        $NinjaOneurl = "https://$($NinjaOneInstance)/#/customerDashboard/$($N1Org)/documentation/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                    } else {
                        $NinjaOneurl = "https://$($NinjaOneInstance)/#/systemDashboard/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                    }
                    $NinjaOneName = $MatchedItem.Name

                } else {
                    $NinjaOneUrl = $null
                    $NinjaOneName = $Null
                }

                if ($null -ne $NinjaOneurl) {
                    Write-Host "Matched $($match.groups[3].value) URL to NinjaOne KB Article: $NinjaOneName" -ForegroundColor 'Cyan'
                } else { Write-Warning "The matched regex: $($match.groups[4].value) did not resolve to a NinjaOne article" }
               
            }

            "a" {
                Write-Host "Found a DOC Locator link for locator $($match.groups[2].value)" -ForegroundColor 'Blue'
                $MatchedItem = $MatchedArticles | Where-Object { $_.ITGID -eq $match.groups[4].value }
                if (($MatchedItem | Measure-Object).count -eq 1) {

                    $N1Org = $($MatchedItem.NinjaOneObject.organizationId)
                    $N1PFolder = $($MatchedItem.NinjaOneObject.parentFolderId)
                    $N1ID = $MatchedItem.NinjaOneID

                    if ($Null -ne $N1Org) {
                        $NinjaOneurl = "https://$($NinjaOneInstance)/#/customerDashboard/$($N1Org)/documentation/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                    } else {
                        $NinjaOneurl = "https://$($NinjaOneInstance)/#/systemDashboard/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                    }
                    $NinjaOneName = $MatchedItem.Name

                } else {
                    $NinjaOneUrl = $null
                    $NinjaOneName = $Null
                }

                if ($null -ne $NinjaOneurl) {
                    Write-Host "Matched $($match.groups[2].value) Locator to NinjaOne doc: $NinjaOneName" -ForegroundColor 'Cyan'
                } else { Write-Warning "The matched regex: $($match.groups[4].value) did not resolve to a NinjaOne article" }

            }

            "passwords" {
                Write-Host "Found an $($match.groups[3].value) URL to replace" -ForegroundColor 'Blue'
                $MatchedItem = ($MatchedPasswords | Where-Object { $_.ITGID -eq $match.groups[4].value -and $_.Matched -eq $true })
                if (($MatchedItem | Measure-Object).count -eq 1) {
                    $N1Org = $($MatchedItem.NinjaOneObject.organizationId)
                    $N1ID = $MatchedItem.NinjaOneID
                    $N1Template = $($MatchedItem.NinjaOneObject.documentTemplateId)

                    $NinjaOneUrl = "https://$($NinjaOneInstance)/#/customerDashboard/$($N1Org)/documentation/appsAndServices/$($N1Template)/$($N1ID)"
                    $NinjaOneName = $MatchedItem.Name

                } else {
                    $NinjaOneUrl = $null
                    $NinjaOneName = $Null
                }

                if ($null -ne $NinjaOneurl) {
                    Write-Host "Matched $($match.groups[3].value) URL to NinjaOne Passsword: $NinjaOneName" -ForegroundColor 'Cyan'
                } else { Write-Warning "The matched regex: $($match.groups[4].value) did not resolve to a NinjaOne Password" }
            }

            "configurations" {
                Write-Host "Found an $($match.groups[3].value) URL to replace" -ForegroundColor 'Blue'
                $MatchedItem = $MatchedDevices | Where-Object { $_.ITGID -eq $match.groups[4].value -and $_.Matched -eq $true }
                
                if (($MatchedItem | Measure-Object).count -eq 1) {

                    switch ($MatchedItem.NinjaOneObject.nodeClass) {
                        'UNMANAGED_DEVICE' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/unmanagedDeviceDashboard/$($MatchedItem.NinjaOneID)/overview" }
                        'MANAGED_DEVICE' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'WINDOWS_SERVER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'WINDOWS_WORKSTATION' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'LINUX' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'MAC' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'VMWARE_VM_HOST' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/vmDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'VMWARE_VM_GUEST' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/vmGuestDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'LINUX_SERVER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'MAC_SERVER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'CLOUD_MONITOR_TARGET' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/cloudMonitorDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_SWITCH' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_ROUTER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_FIREWALL' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_PRIVATE_NETWORK_GATEWAY' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_PRINTER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_SCANNER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_DIAL_MANAGER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_WAP' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_IPSLA' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_COMPUTER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_VM_HOST' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_APPLIANCE' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_OTHER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_SERVER' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_PHONE' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_VIRTUAL_MACHINE' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'NMS_NETWORK_MANAGEMENT_AGENT' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/nmsDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'HYPERV_VMM_HOST' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/vmDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'HYPERV_VMM_GUEST' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/vmGuestDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'ANDROID' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/mobileDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'APPLE_IOS' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/mobileDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        'APPLE_IPADOS' { $NinjaOneURL = "https://$($NinjaOneInstance)/#/mobileDashboard/$($MatchedItem.NinjaOneID)/overview)" }
                        default { $NinjaOneURL = "https://$($NinjaOneInstance)/#/deviceDashboard/$($MatchedItem.NinjaOneID)/overview)" }

                    }
                    $NinjaOneName = $MatchedItem.name
                } else {
                    $NinjaOneUrl = $null
                    $NinjaOneName = $Null
                }

                if ($null -ne $NinjaOneurl) {
                    Write-Host "Matched $($match.groups[3].value) URL to NinjaOne Device: $NinjaOneName" -ForegroundColor 'Cyan'
                } else { 
                    Write-Warning "The matched regex: $($match.groups[4].value) did not resolve to a NinjaOne Device" 
                }
            }

            "assets" {
                Write-Host "Found an $($match.groups[3].value) URL to replace" -ForegroundColor 'Blue'
                $MatchedItem = $MatchedAssets | Where-Object { $_.ITGID -eq $match.groups[4].value -and $_.Matched -eq $true }
                if (($MatchedItem | Measure-Object).count -eq 1) {

                    $N1Org = $($MatchedItem.NinjaOneObject.organizationId)
                    $N1ID = $MatchedItem.NinjaOneID
                    $N1Template = $($MatchedItem.NinjaOneObject.documentTemplateId)

                    $NinjaOneUrl = "https://$($NinjaOneInstance)/#/customerDashboard/$($N1Org)/documentation/appsAndServices/$($N1Template)/$($N1ID)"
                    $NinjaOneName = $MatchedItem.Name

                } else {
                    $NinjaOneUrl = $null
                    $NinjaOneName = $Null
                }

                if ($null -ne $NinjaOneurl) {
                    Write-Host "Matched $($match.groups[3].value) URL to NinjaOne Document: $NinjaOneName" -ForegroundColor 'Cyan'
                } else { Write-Warning "The matched regex: $($match.groups[4].value) did not resolve to a NinjaOne Document" }
            }

            "images" {
                Write-Host "Found an external image using a Direct ITGlue link" -ForegroundColor 'Blue'
                $OriginalArticle = ($MatchedArticles | Where-Object { $_.ITGID -eq $match.groups[2].value }).Path
                $ImagePath = $match.groups[1].value.replace('/', '\')
                $FullImagePath = Join-Path -Path $OriginalArticle -ChildPath $ImagePath
                $ImageItem = Get-Item -Path "$FullImagePath*" -ErrorAction SilentlyContinue
                if ($ImageItem) {
                    Return [pscustomobject]@{
                        "path" = $ImageItem.FullName
                        "url"  = $match.Groups[1]
                    }
                } else { return $false }
            }
            default {
                if ($match.groups[1].value -like 'DOC-*') {
                    Write-Host "Found a DOC Locator link for locator $($match.groups[1].value)" -ForegroundColor 'Blue'
                    $MatchedItem = $MatchedArticles | Where-Object { $_.ITGLocator -eq $match.groups[1].value }
                    if (($MatchedItem | Measure-Object).count -eq 1) {

                        $N1Org = $($MatchedItem.NinjaOneObject.organizationId)
                        $N1PFolder = $($MatchedItem.NinjaOneObject.parentFolderId)
                        $N1ID = $MatchedItem.NinjaOneID

                        if ($Null -ne $N1Org) {
                            $NinjaOneurl = "https://$($NinjaOneInstance)/#/customerDashboard/$($N1Org)/documentation/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                        } else {
                            $NinjaOneurl = "https://$($NinjaOneInstance)/#/systemDashboard/knowledgeBase/" + $(if ($N1PFolder) { "$($N1PFolder)/" }) + "$($N1ID)/file"
                        }
                        $NinjaOneName = $MatchedItem.Name

                    } else {
                        $NinjaOneUrl = $null
                        $NinjaOneName = $Null
                    }

                    if ($null -ne $NinjaOneurl) {
                        Write-Host "Matched $($match.groups[1].value) Locator to NinjaOne doc: $NinjaOneName" -ForegroundColor 'Cyan'
                    } else { Write-Warning "The matched regex: $($_.ITGLocator -eq $match.groups[1].value) did not resolve to a NinjaOne article" }
                }
            }


        
        }
    
        if ($null -ne $NinjaOneurl) {
            if ($type -eq 'rich') {
                $ReplacementString = @"
<a href="$NinjaOneurl">$NinjaOneName</a>
"@
            } else {
                $ReplacementString = $NinjaOneurl
            }

            $inputString = $inputString -replace [regex]::Escape([string]$match.Value), [string]$ReplacementString
        }

      

    }
  
    return $inputString
}
  

function ConvertTo-NinjaOneURL {
    param(
        $Content
    )
    $NewContent = Update-StringWithCaptureGroups -inputString $Content -pattern $RegexPatternToMatchSansAssets
    $NewContent = Update-StringWithCaptureGroups -inputString $NewContent -pattern $RegexPatternToMatchWithAssets

    return $NewContent

}