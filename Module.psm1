#requires -module Microsoft.Graph.Sites,Microsoft.Graph.Users,Microsoft.Graph.Groups,Microsoft.Graph.Files
using namespace Microsoft.Graph.PowerShell.Models
using namespace Microsoft.Graph.PowerShell.Runtime
using namespace System.Management.Automation
using namespace System.Collections.Generic
Update-FormatData -PrependPath $PSScriptRoot\Formats\*.Format.ps1xml

#Reference: https://www.youtube.com/watch?v=YYMFP8xcNOQ

filter Get-JMgDrive {
    <#
.SYNOPSIS
A Universal method of getting the drive of a resource (site, onedrive, user, email, url, etc.) by piping it to this command
.EXAMPLE
Get-JMGDrive -Root
.EXAMPLE
'user@principalname.com' | Get-JMGDrive
.EXAMPLE
'https://mysite.sharepoint.com' | Get-JMGDrive
.EXAMPLE
'https://mysite.sharepoint.com/sites/OurSite' | Get-JMGDrive
.EXAMPLE
Get-MgSite -Property "siteCollection,webUrl" -Filter "siteCollection/root ne null" | Get-JMGDrive
.EXAMPLE
Get-MgSite 'SiteSearchKeyword'
#>
    [OutputType('MicrosoftGraphDrive1[]')]
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    param(
        #Retrieve a drive by user UPN, sharepoint site URI, or a user/onedrive/site/team object
        [Parameter(Position = 0, ValueFromPipeline, ParameterSetName = 'InputObject')]$InputObject,
        #Specify this to get the root Sharepoint Site
        [Parameter(Mandatory, ParameterSetName = 'Root')][Switch]$Root
    )
    if (-not (Get-MgContext)) {
        throw 'You are not connected to Microsoft Graph. Please run Connect-MgGraph first.'
    }
    if ($null -eq $InputObject) {
        #Fetch the "my" drive by default
        Write-Verbose 'No input specified, fetching the /me drive.'
        $drive = try {
            [MicrosoftGraphDrive1[]](Invoke-MgGraphRequest -Method Get 'v1.0/me/drives').value
        } catch {
            $PSItem
            | Set-StatusCodeErrorMessage 'NotFound' 'Tried to fetch your OneDrive by default, but you are not licensed or provisioned for OneDrive. Please choose another search option.'
            | Write-Error
            return
        }
        return $drive
    }
    if ($root) {
        return Get-MgDrive
    }
    switch ($InputObject.GetType().Name) {
        'Uri' { Get-MgSiteByUri $Uri | Get-JMgDrive }
        'MailAddress' { Get-MgUser -UserId $InputObject | Where-Object { $PSItem } | Get-JMgDrive }
        'MicrosoftGraphSite1' { Get-MgSiteDrive -SiteId $InputObject.Id }
        'MicrosoftGraphUser1' {
            try {
                Get-MgUserDrive -UserId $InputObject.Id -ErrorAction stop
            } catch [RestException] {
                if ($PSItem -match 'Access Denied') {
                    $PSItem.ErrorDetails = "You don't have access to this user's OneDrive. This is the default setting in O365 for privacy. To access these files as someone other than the user, you must create an access link: https://docs.microsoft.com/en-us/microsoft-365/admin/add-users/remove-former-employee-step-5?view=o365-worldwide"
                }
                Write-Error -ErrorRecord $PSItem; return
            }
        }
        'MicrosoftGraphGroup1' { Get-MgGroupDrive -GroupId $InputObject.Id }
        'MicrosoftGraphTeam1' { Get-MgGroupDrive -GroupId $InputObject.Id }
        default {
            [uri]$uri = $null
            if ([Uri]::TryCreate($InputObject, 'Absolute', [ref]$uri)) {
                return $uri | Get-JMgDrive
            }

            [mailaddress]$upn = $null
            if ([MailAddress]::TryCreate($InputObject, [ref]$upn)) {
                return $upn | Get-JMgDrive
            }

            # Last resort is a keyword search of Sharepoint Sites
            $sites = Get-MgSite -Search ([String]$InputObject)
            if (-not $sites) { throw "Site Search returned no results for $InputObject" }
            $sites | Get-JMgDrive
        }
    }
}

filter Get-JMgSiteByUri {
    param(
        [Uri]$Uri
    )
    $siteId = $Uri.Host, $Uri.AbsolutePath -join ':'
    Get-MgSite -SiteId $siteId
}

filter Get-JMgDriveItem {
    [CmdletBinding()]
    param(
        #The Id of the drive. You can pipe from Get-JMg
        [Parameter(ValueFromPipeline)]
        [Microsoft.Graph.Powershell.Models.MicrosoftGraphDrive1]$Drive,
        #The path to the file. If not specified, it gets the root folder
        [String]$Path
    )
    if (-not (Get-MgContext)) {
        throw 'You are not connected to Microsoft Graph. Please run Connect-MgGraph first.'
    }
    if (-not $Drive) {
        Write-Verbose 'No ID specified, fetching the contents of the /me drive.'
        $Drive = Get-JMgDrive
    }
    if (-not $Drive) {
        Write-Error 'No drive found.'
        return
    }
    $DriveId = $Drive.Id

    #Replace leading slashes
    $Path = $Path -replace '^/+'

    if (-not $Path) {
        return Get-MgDriveRoot -DriveId $DriveId
    } else {
        try {
            [MicrosoftGraphDriveItem1](Invoke-MgGraphRequest -Method GET "v1.0/drives/$DriveId/root:/$Path" -ErrorAction stop)
        } catch {
            $PSItem
            | Set-StatusCodeErrorMessage 'NotFound' "The file or folder '$Path' does not exist in drive $($Drive.Name). Paths should be specified in folder/folder/file.txt format"
            | Write-Error
            return
        }
    }
}

filter Get-JMgDriveChildItem {
    [CmdletBinding()]
    param(
        #The Id of the drive. You can pipe from Get-JMg
        [Parameter(ValueFromPipeline)]$Drive,
        #The path to the file. If not specified, it gets the root folder
        [String]$Path
    )

    #DriveItem casts to Drive so we cant do a standard type parameter for -Drive
    if ($Drive.GetType().Name -eq 'MicrosoftGraphDriveItem1') {
        Write-Error -Exception ([NotImplementedException]'You cant pass DriveItems to this command yet, only drives. Use -Path if you want a subfolder')
        return
    }
    $Drive = [MicrosoftGraphDrive1]$Drive

    if (-not (Get-MgContext)) {
        throw 'You are not connected to Microsoft Graph. Please run Connect-MgGraph first.'
    }
    if (-not $Drive) {
        Write-Verbose 'No ID specified, fetching the contents of the /me drive.'
        $Drive = Get-JMgDrive
    }
    if (-not $Drive) {
        Write-Error '-Drive parameter was not supplied and your account does not have a default OneDrive'
        return
    }
    $DriveId = $Drive.Id
    $DrivePath = $Path ? "root:/${Path}:/children" : 'root/children'
    try {
        [MicrosoftGraphDriveItem1[]](Invoke-MgGraphRequest -Method GET "v1.0/drives/$DriveId/$DrivePath" -ErrorAction stop).Value
    } catch {
        $PSItem
        | Set-StatusCodeErrorMessage 'NotFound' "The file or folder '$Path' does not exist in drive $($Drive.Name). Paths should be specified in folder/folder/file.txt format"
        | Write-Error
        return
    }
}

function Save-JmgDriveItem {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #Where to save the file. Defaults to your current Directory
        [String]$Path,
        #Overwrite Files
        [Switch]$Force,
        [Parameter(Mandatory, ValueFromPipeline)]
        [Microsoft.Graph.Powershell.Models.MicrosoftGraphDriveItem1]$DriveItem
    )
    begin {
        if (-not (Get-MgContext)) {
            throw 'You are not connected to Microsoft Graph. Please run Connect-MgGraph first.'
        }
        $jobs = [List[Job2]]@()
    }
    process {
        $dstPath = $Path
        if (-not $dstPath) {
            if (-not $DriveItem.Name) {
                Write-Error 'The drive item supplied does not have a filename. Please specify -Path with the full path to save.'
            }
            $dstPath = Join-Path $PWD $DriveItem.Name
        }
        $existingItem = Get-Item $dstPath -ErrorAction SilentlyContinue
        if ($ExistingItem -is [IO.DirectoryInfo]) {
            $dstPath = Join-Path $dstPath $DriveItem.Name
            $existingItem = Get-Item $dstPath -ErrorAction SilentlyContinue
        }

        if ($ExistingItem -and -not $Force) {
            Write-Error "The file '$dstPath' already exists. Use -Force to overwrite."
            return
        }

        $downloadUriProperty = '@microsoft.graph.downloadUrl'
        if (-not $DriveItem.AdditionalProperties.ContainsKey($downloadUriProperty)) {
            Write-Error "$($DriveItem.Name) is either a folder or cannot be downloaded."
            return
        }
        [uri]$downloadUri = $DriveItem.AdditionalProperties[$downloadUriProperty]
        if ($PSCmdlet.ShouldProcess($dstPath, "Download file $($DriveItem.Name)")) {
            $job = Start-ThreadJob -Name "Download-$($DriveItem.Name)" {
                $ProgressPreference = 'SilentlyContinue'
                Invoke-RestMethod -Uri $USING:downloadUri -OutFile $USING:dstPath
            }
            $jobs.Add($job)
        }
    }
    end {
        if ($jobs.count -gt 0) {
            Receive-Job $jobs -Wait -AutoRemoveJob
        }
    }
}

function Push-JmgDriveItem {
    <#
    .SYNOPSIS
    Upload a file to a drive item
    .EXAMPLE
    Get-Item test.txt | Push-JmgDriveItem (Get-JMgDriveItem)

    Uploads a file to your OneDrive
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #The Path to the file to upload. This only supports individual files at the moment but you can pipe multiple files.
        [Parameter(Mandatory, ValueFromPipeline)][String]$Path,
        #Overwrite Files
        [Switch]$Force,
        #Return the drive items after they are uploaded
        [Switch]$PassThru,
        #The upload destination
        [Microsoft.Graph.Powershell.Models.MicrosoftGraphDriveItem1]$DriveItem
    )
    begin {
        if (-not (Get-MgContext)) {
            throw 'You are not connected to Microsoft Graph. Please run Connect-MgGraph first.'
        }
        $jobs = [List[Job2]]@()
        if (-not $DriveItem) {
            Write-Verbose 'No destination specified, uploading to your OneDrive Documents folder'
            $DriveItem = Get-JMgDriveItem
        }
        #HACK: Because of how the objects get composed there's not a great way to check if the destination is a folder or a file.
        if ($null -ne $DriveItem.Folder.ChildCount) {
            Write-Verbose "Detected that the driveItem $($DriveItem.Name) is a folder, will save to same-named file in the folder"
            $DriveItemIsFolder = $true
        }
    }
    process {
        $err = $null
        $Item = Get-Item $Path -ErrorVariable err
        if ($err) { return }

        if (($Item.Length / 60MB) -ge 1) {
            Write-Error -Exception ([NotImplementedException]'Currently can only process files of 60MB or smaller size')
            return
        }

        if ($Item -is [IO.DirectoryInfo]) {
            Write-Error -Exception ([NotSupportedException]'Folders and recursion are not yet supported, specify the individual files instead. HINT: Get-ChildItem -File.')
            return
        }

        $driveId = $DriveItem.ParentReference.DriveId

        [string]$ItemPath = switch ($true) {
            (Test-IsDriveRoot $DriveItem) {
                'root' + ':/' + $Item.Name + ':'; break
            }
            $driveItemIsFolder {
                'items/' + $driveItem.Id + ':/' + $Item.Name + ':' ; break
            }
            default {
                'items/' + $driveItem.Id
            }
        }
        $createUploadSessionUri = "https://graph.microsoft.com/v1.0/drives/$driveId/$itemPath/createUploadSession"
        $uploadSessionBody = @{
            #BUG: The order of entries in item is apparently important: https://github.com/microsoftgraph/microsoft-graph-docs/issues/17072
            #That's why we use ordered here
            item = [ordered]@{
                '@microsoft.graph.conflictBehavior' = $Force ? 'replace' : 'fail'
                name                                = $Item.Name
            }
        }
        if ($PSCmdlet.ShouldProcess($DriveItem.Name, "Upload File $($Item.Name)")) {
            if ($err) { return }

            $context = @{
                Uri      = $createUploadSessionUri
                Body     = $uploadSessionBody
                Item     = $Item
                PSCmdlet = $PSCmdlet
            }
            $job = Start-ThreadJob -Name "Upload-$($Item.Name)" -ArgumentList $context -ScriptBlock {
                param($context)
                $uploadSession = try {
                    $ProgressPreference = 'SilentlyContinue'
                    Invoke-MgGraphRequest -Method 'POST' -Uri $context.Uri -ContentType 'application/json' -EA stop -SessionVariable session -Body $context.Body
                } catch {
                    if ($psitem.exception.response.statuscode -eq 'Conflict') {
                        $PSItem.ErrorDetails = "The file '$($context.Item.Name)' already exists. Use -Force to overwrite."
                    }
                    $PSItem
                    return
                }
                $uploadParams = @{
                    Method      = 'PUT'
                    Uri         = $UploadSession.uploadUrl
                    ContentType = 'application/octet-stream'

                    Headers     = @{
                        'Content-Range' = 'bytes 0-' + ($context.Item.Length - 1) + '/' + $context.Item.Length
                    }
                    Body        = [IO.File]::ReadAllBytes($context.Item.FullName)
                }
                [Microsoft.Graph.Powershell.Models.MicrosoftGraphDriveItem1](Invoke-MgGraphRequest @UploadParams)
            }
            $jobs.Add($job)
        }
    }
    end {
        if ($jobs.count -gt 0) {
            Receive-Job $jobs -Wait -AutoRemoveJob
            | ForEach-Object {
                if ($PSItem -is [ErrorRecord]) {
                    $PSCmdlet.WriteError($PSItem)
                } elseif ($PassThru) {
                    $PSItem
                }
            }
        }
    }
}


function Test-IsDriveRoot ($DriveItem) {
    $DriveItem.AdditionalProperties.'@odata.context' -match '/root/\$entity$'
}

filter Set-StatusCodeErrorMessage ([Net.HttpStatusCode]$code, [String]$message, [Parameter(Mandatory, ValueFromPipeline)][ErrorRecord]$err) {
    if ($err.Exception.GetType().Name -eq 'HttpResponseException' -and
        $err.Exception.Response.statuscode -eq $code
    ) {
        $err.ErrorDetails = $message
    }
    return $err
}
