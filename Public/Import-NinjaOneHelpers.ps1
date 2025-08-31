function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret ) {
        if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
            return $Script:NinjaToken
        } else {

            if ($Script:NinjaOneRefreshToken) {
                $Body = @{
                    'grant_type'    = 'refresh_token'
                    'client_id'     = $Script:NinjaOneClientID
                    'client_secret' = $Script:NinjaOneClientSecret
                    'refresh_token' = $Script:NinjaOneRefreshToken
                }
            } else {

                $body = @{
                    grant_type    = 'client_credentials'
                    client_id     = $Script:NinjaOneClientID
                    client_secret = $Script:NinjaOneClientSecret
                    scope         = 'monitoring management'
                }
            }

            $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    
            $Script:NinjaTokenExpiry = (Get-Date).AddSeconds(3000)
            $Script:NinjaToken = $token
            

            Write-Host 'Fetched New Token'
            return $token
        } else {
            Throw 'Please run Connect-NinjaOne first'
        }
    }

}
function Get-OAuthCode {
    param (
        [System.UriBuilder]$AuthURL,
        [string]$RedirectURL
    )
    $HTTP = [System.Net.HttpListener]::new()
    $HTTP.Prefixes.Add($RedirectURL)
    $HTTP.Start()
    Start-Process $AuthURL.ToString()
    $Result = @{}
    while ($HTTP.IsListening) {
        $Context = $HTTP.GetContext()
        if ($Context.Request.QueryString -and $Context.Request.QueryString['Code']) {
            $Result.Code = $Context.Request.QueryString['Code']
            if ($null -ne $Result.Code) {
                $Result.GotAuthorisationCode = $True
            }
            [string]$HTML = @"
            <html lang="en">
<head>
    <meta charset="UTF-8">
    <title>NinjaOne Authorization Code</title>
    <style>
        body {
            background-color: #f0f2f5;
            font-family: Arial, sans-serif;
            margin: 0;
        }
        .card {
            max-width: 500px;
            margin: 100px auto;
            background-color: #ffffff;
            padding: 40px 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            text-align: center;
        }
        .card h1 {
            margin-bottom: 20px;
            font-size: 24px;
            color: #333333;
        }
        .card p {
            margin: 10px 0;
            font-size: 16px;
            color: #555555;
        }
        .checkmark {
            font-size: 60px;
            color: #28a745;
            margin-bottom: 20px;
        }
    </style>
</head>
<body>
    <div class="card">
        <div class="checkmark">&#10004;</div>
        <h1>NinjaOne Login Successful</h1>
        <p>An authorization code has been received successfully.</p>
        <p>Please close this tab and return to the import tool Window.</p>
    </div>
</body>
</html>
"@
            $Response = [System.Text.Encoding]::UTF8.GetBytes($HTML)
            $Context.Response.ContentLength64 = $Response.Length
            $Context.Response.OutputStream.Write($Response, 0, $Response.Length)
            $Context.Response.OutputStream.Close()
            Start-Sleep -Seconds 1
            $HTTP.Stop()
        }
    }
    Return $Result
}

function Convert-ToHttpsUrl {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Url
    )

    process {
        # Trim whitespace
        $Url = $Url.Trim()

        # If it already starts with https://, return as-is
        if ($Url -match '^https://') {
            return $Url
        }

        # If it starts with http://, replace with https://
        elseif ($Url -match '^http://') {
            return $Url -replace '^http://', 'https://'
        }

        # If no scheme is provided, prepend https://
        else {
            return "https://$Url"
        }
    }
}

function Connect-NinjaOne {
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true)]
        $NinjaOneInstance,
        [Parameter(mandatory = $true)]
        $NinjaOneClientID,
        [Parameter(mandatory = $true)]
        $NinjaOneClientSecret,
        $NinjaOneRefreshToken
    )

    $AuthURL = "https://$NinjaOneInstance/ws/oauth/authorize?response_type=code&client_id=$NinjaOneClientID&redirect_uri=$NinjaOneRedirectURL&scope=monitoring%20management%20offline_access&state=STATE"

    $Result = Get-OAuthCode -AuthURL $AuthURL -RedirectURL $NinjaOneRedirectURL

    $AuthBody = @{
        'grant_type'    = 'authorization_code'
        'client_id'     = $NinjaOneClientID
        'client_secret' = $NinjaOneClientSecret
        'code'          = $Result.code
        'redirect_uri'  = $NinjaOneRedirectURL 
    }

    $Result = Invoke-WebRequest -uri "https://$($NinjaOneInstance)/ws/oauth/token" -Method POST -Body $AuthBody -ContentType 'application/x-www-form-urlencoded'

    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $Script:NinjaOneRefreshToken = ($Result.content | ConvertFrom-Json).refresh_token
    

    try {
        $Null = Get-NinjaOneToken -ea Stop
    } catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }

}

function Invoke-NinjaOneRequest {
    param(
        $Method,
        $Body,
        $InputObject,
        $Path,
        $QueryParams,
        [Switch]$Paginate,
        [Switch]$AsArray
    )

    $Token = Get-NinjaOneToken

    if ($InputObject) {
        if ($AsArray) {
            $Body = $InputObject | ConvertTo-Json -depth 100
            if (($InputObject | Measure-Object).count -eq 1 ) {
                $Body = '[' + $Body + ']'
            }
        } else {
            $Body = $InputObject | ConvertTo-Json -depth 100
        }
    }

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {           
                $After = 0
                $PageSize = 1000
                $NinjaResult = do {
                    $Result = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)?pageSize=$PageSize&after=$After$(if ($QueryParams){"&$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing).content | ConvertFrom-Json
                    $Result
                    $ResultCount = ($Result.id | Measure-Object -Maximum)
                    $After = $ResultCount.maximum
                } while (($Result | Measure-Object).count -eq $PageSize)

                Return $NinjaResult

            } else {
                $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json; charset=utf-8' -UseBasicParsing
            }       

        } elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -Body $Body -ContentType 'application/json; charset=utf-8' -UseBasicParsing
        } else {
            Throw 'Unknown Method'
        }
    } catch {
        Throw "Error Occured: $_"
    }

    try {
        return $NinjaResult.content | ConvertFrom-Json -ea stop
    } catch {
        return $NinjaResult.content
    }

}

function Invoke-NinjaOneDocumentTemplate {
    [CmdletBinding()]
    param (
        $Template,
        $ID
    )

    if (!$ID) {
        $DocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates" -Method GET -QueryParams "templateName=$($Template.name)" | where-object { $_.name -eq $Template.name }
    } else {
        $DocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates/$($ID)" -Method GET
    }
    
    $PatchTemplate = $False

    $MatchedCount = ($DocumentTemplate | Measure-Object).count
    if ($MatchedCount -eq 1) {
        # Matched a single document template
        # Check fields are correct
        foreach ($Field in $Template.Fields) {
            if ($Field.fieldName) {
                $MatchedField = $DocumentTemplate.Fields | Where-Object { $_.fieldName -eq $Field.fieldName -and $_.fieldType -eq $Field.fieldType }
                if (($MatchedField | Measure-Object).count -ne 1) {
                    $MatchedField = $DocumentTemplate.Fields | Where-Object { $_.fieldName -eq $Field.fieldName }
                    $MatchCount = ($MatchedField | Measure-Object).count
                    if ($MatchCount -eq 1 ) {
                        Throw "$($Field.fieldName) exists with the wrong type. Please manually edit the template $($Template.name) to set it to a $($Field.fieldType) field."
                    } elseif ($MatchCount -eq 0) {
                        $PatchTemplate = $True
                    } else {
                        Throw "Mutliple Fields exists for $($Field.fieldName) in $($Template.name)"
                    }
                }
            } elseif ($field.uiElementType) {
                continue
            } else {
                Throw "$Field had no fieldName or uiElementName"
            }
        }

        if ($PatchTemplate -eq $True) {
            Write-Host "Updating Template"
            $Template.Fields = $Template.Fields + ($DocumentTemplate.Fields | Where-Object { $_.fieldName -notin $Template.Fields.fieldName })
            $NinjaDocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates/$($DocumentTemplate.id)" -Method PUT -InputObject ($Template | Select-Object * -ExcludeProperty allowMultiple)
        }

        $NinjaDocumentTemplate = $DocumentTemplate


    } elseif ($MatchedCount -eq 0) {
        # Create a new Document Template
        Write-Host "Creating Template"
        $NinjaDocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates" -Method POST -InputObject $Template
    } else {
        # Matched multiple templates. Should be impossible but lets check anyway :D
        Throw "Multiiple Documents Matched the Provided Criteria"
    }

    return $NinjaDocumentTemplate

}

function Get-NinjaOneTime {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [DateTime]$Date,
        [Switch]$Seconds
    )
    try {
        $unixEpoch = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        $timeSpan = $Date.ToUniversalTime() - $unixEpoch

        if ($Seconds) {
            return [int64]([math]::Round($timeSpan.TotalSeconds))
        } else {
            return [int64]([math]::Round($timeSpan.TotalMilliSeconds))
        }
    } catch {
        return ""
    }
}

function Invoke-UploadNinjaOneFile($FileName, $FilePath, $ContentType, $EntityType) {

    try {
        $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
        $FileStream = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open)
        $fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
        $fileHeader.Name = 'files'
        $fileHeader.FileName = $FileName
        $fileContent = [System.Net.Http.StreamContent]::new($FileStream)
        $fileContent.Headers.ContentDisposition = $fileHeader
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($ContentType)
        $multipartContent.Add($fileContent)
        if ($EntityType) {
            $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/attachments/temp/upload?entityType=$EntityType"
        } else {
            $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/attachments/temp/upload"
        }
        write-host "$(Get-Date) - Starting upload"
        $Result = (Invoke-WebRequest -Uri $URI -Body $multipartContent -Method 'POST' -Headers @{Authorization = "Bearer $(($(Get-NinjaOneToken)).access_token)" }).content | ConvertFrom-Json -Depth 100
        Write-host "$(Get-Date) - Upload finished"
        $FileStream.close()
        return $Result
    } catch {
        $FileStream.close()
        Throw "Failed to upload file: $_"
    }

}

function Get-MimeType {
    param([string]$Path)

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($ext) {
        ".jpg" { "image/jpeg" }
        ".jpeg" { "image/jpeg" }
        ".png" { "image/png" }
        ".gif" { "image/gif" }
        ".cab" { "application/vnd.ms-cab-compressed" }
        ".txt" { "text/plain" }
        ".log" { "text/plain" }
        ".pdf" { "application/pdf" }
        ".csv" { "text/csv" }
        ".mp3" { "audio/mpeg" }
        ".eml" { "message/rfc822" }
        ".dot" { "application/msword" }
        ".wbk" { "application/msword" }
        ".doc" { "application/msword" }
        ".docx" { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        ".rtf" { "application/rtf" }
        ".xls" { "application/vnd.ms-excel" }
        ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        ".ods" { "application/vnd.oasis.opendocument.spreadsheet" }
        ".ppt" { "application/vnd.ms-powerpoint" }
        ".pptx" { "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
        ".pps" { "application/vnd.ms-powerpoint" }
        ".ppsx" { "application/vnd.openxmlformats-officedocument.presentationml.slideshow" }
        ".sldx" { "application/vnd.openxmlformats-officedocument.presentationml.slide" }
        ".vsd" { "application/vnd.visio" }
        ".vsdx" { "application/vnd.ms-visio.drawing.main+xml" }
        ".xml" { "application/xml" }
        ".html" { "text/html" }
        ".zip" { "application/zip" }
        ".rar" { "application/vnd.rar" }
        ".tar" { "application/x-tar" }
        default { "application/octet-stream" }
    }
}

function Invoke-UploadNinjaOneKBArticle {
    param (
        $FileName,
        $FilePath,
        $FolderPath,
        $OrganizationID,
        $FailCount = 0
    )

    try {
        $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
        # Only add fields that were provided
        if ($null -ne $OrganizationID) {
            $multipartContent.Add([System.Net.Http.StringContent]::new($OrganizationId), 'organizationId')
        }
        if ($FolderPath -ne '') {
            $multipartContent.Add([System.Net.Http.StringContent]::new($FolderPath), 'folderPath')
        }

        $MimeType = Get-MimeType $FilePath

        $FileStream = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open)
        $fileHeader = [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")
        $fileHeader.Name = 'files'
        $fileHeader.FileName = $FileName
        $fileContent = [System.Net.Http.StreamContent]::new($FileStream)
        $fileContent.Headers.ContentDisposition = $fileHeader
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($MimeType)
        $multipartContent.Add($fileContent)

        $URI = "https://$($Script:NinjaOneInstance)/ws/api/v2/knowledgebase/articles/upload"
        $Result = (Invoke-WebRequest -Uri $URI -Body $multipartContent -Method 'POST' -Headers @{Authorization = "Bearer $(($(Get-NinjaOneToken)).access_token)" }).content | ConvertFrom-Json -Depth 100
        $FileStream.close()
        return $Result
    } catch {
        $FileStream.close()
        if ($Failcount -le 9) {
            $Failcount++
            Write-Host "Upload failed retrying: $Failcount"
            try {
                $Result = Invoke-UploadNinjaOneKBArticle -FileName $FileName -FilePath $FilePath -FolderPath $FolderPath -OrganizationID $OrganizationID -Failcount $Failcount -ea stop
                return $result
            } catch {
                Throw $_
            }
        } else {
            $FileStream.close()
            Throw "Failed to upload file: $_"
        }     
    }

}
