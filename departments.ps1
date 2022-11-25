function Get-RAETXMLDepartments {
    param(
        [parameter(Mandatory)]
        $XMLBasePath,

        [parameter(Mandatory)]
        $FileFilter
    )

    $File = Get-ChildItem -Path $XMLBasePath -Filter $FileFilter |
        Sort-Object -Property 'LastWriteTime' -Descending |
        Select-Object -First 1

    if ($File.Count -eq 0) {
        return
    }

    # Read content as XML
    [xml]$xml = Get-Content $File.FullName

    $Elements = $xml.GetElementsByTagName("orgEenheid")

    $Elements | ForEach-Object {

        $Result = [PSCustomObject]@{}

        $_.ChildNodes | ForEach-Object {
            $Result | Add-Member -NotePropertyName $_.LocalName -NotePropertyValue $_.'#text' -Force
        }

        $Result
    }
}

function Get-RAETXMLManagers {
    param(
        [parameter(Mandatory)]
        $XMLBasePath,

        [parameter(Mandatory)]
        $FileFilter,

        [parameter(Mandatory)]
        [string[]]$ManagerRoleCodes
    )

    $File = Get-ChildItem -Path $XMLBasePath -Filter $FileFilter |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($file.Count -eq 0) {
        return
    }

    # Read content as XML
    [xml]$xml = Get-Content $File.FullName

    $Elements = $xml.GetElementsByTagName("roltoewijzing") | Where-Object {
        $_.oeRolCode -in $ManagerRoleCodes -and
        [DateTime]::ParseExact($_.begindatum, "yyyy-MM-dd", $null) -le [DateTime]::Today -and
        (
            [String]::IsNullOrEmpty($_.einddatum) -or
            [DateTime]::ParseExact($_.einddatum, "yyyy-MM-dd", $null) -ge [DateTime]::Today
        )
    }

    # Process all records
    $Elements | ForEach-Object {

        $Result = [PSCustomObject]@{}

        $_.ChildNodes | ForEach-Object {
            $Result | Add-Member -NotePropertyName $_.LocalName -NotePropertyValue $_.'#text' -Force
        }

        $Result
    }
}

Write-Verbose -Verbose "Import started"

# Init variables
$connectionSettings = ConvertFrom-Json $configuration

$XmlPath = $($connectionSettings.xmlPath)

$FileTime = (Get-ChildItem -File (Join-Path $xmlPath -ChildPath "IAM_BA_*.xml") | Sort-Object -Descending -Property CreationTime | Select-Object -First 1).name.split('_')[2]

Write-Verbose -Verbose "FileTime: $FileTime"

# Get the source data
$Departments = Get-RAETXMLDepartments -XMLBasePath $XmlPath -FileFilter "rst_orgeenheid_$($FileTime)_*.xml"
$managers = Get-RAETXMLManagers -XMLBasePath $XmlPath -FileFilter "roltoewijzing_*_$($FileTime)_*.xml" -ManagerRoleCodes "MGR"

# Group managers per OE
$managers = $managers | Group-Object orgEenheidID -AsHashTable

# Extend the departments with required and additional fields
$Departments | Add-Member -MemberType AliasProperty -Name "ExternalId" -Value orgEenheidID
$Departments | Add-Member -MemberType AliasProperty -Name "DisplayName" -Value NaamLang
$Departments | Add-Member -MemberType AliasProperty -Name "Name" -Value NaamLang
$Departments | Add-Member -MemberType AliasProperty -Name "ParentExternalId" -Value hogereOrgEenheid

$departments | Add-Member -NotePropertyMembers @{
    ManagerExternalId = $null
}

Write-Verbose -Verbose "Exporting data to HelloID"
$Departments | ForEach-Object {

    # Add the manager
    $manager = $managers[$_.orgEenheidID]

    if ($manager.Count -eq 1) {
        $_.ManagerExternalId = $manager.persNr
    }
    elseif ($manager.Count -gt 1) {
        $_.ManagerExternalId = (
            $manager <# | Sort-Object begindatum -Descending #> | Select-Object -First 1
        ).persnr

        #Write-Verbose -Verbose "[Departments] Multiple managers found for OE $($_.ExternalId): ($([string]$manager.persnr)). Keeping manager $($_.ManagerExternalId)."
    }
}

Write-Output (
    $Departments | Select-Object -Property @(
        "ExternalId", "DisplayName", "Name", "ParentExternalId", "ManagerExternalId"
    ) | ConvertTo-Json -Depth 5 -Compress
)

Write-Verbose -Verbose "Exported data to HelloID"
