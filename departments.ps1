function Get-RAETXMLDepartments {
    param(
        [parameter(Mandatory = $true)]$XMLBasePath,
        [parameter(Mandatory = $true)]$FileFilter,
        [parameter(Mandatory = $true)][ref]$departments
    )

    $files = Get-ChildItem -Path $XMLBasePath -Filter $FileFilter | Sort-Object LastWriteTime -Descending
    if ($files.Count -eq 0) { return }

    # Read content as XML
    [xml]$xml = Get-Content $files[0].FullName

    # Process all records
    foreach ($afdeling in $xml.GetElementsByTagName("orgEenheid")) {
        $department = [PSCustomObject]@{}

        foreach ($child in $afdeling.ChildNodes) {
            $department | Add-Member -MemberType NoteProperty -Name $child.LocalName -Value $child.'#text' -Force
        }

        [void]$departments.value.Add($department)
    }
}

function Get-RAETXMLManagers {
    param(
        [parameter(Mandatory = $true)]$XMLBasePath,
        [parameter(Mandatory = $true)]$FileFilter,
        [parameter(Mandatory = $true)]$ManagerRoleCodes,
        [parameter(Mandatory = $true)][ref]$managers
    )

    $files = Get-ChildItem -Path $XMLBasePath -Filter $FileFilter | Sort-Object LastWriteTime -Descending
    if ($files.Count -eq 0) { return }

    # Read content as XML
    [xml]$xml = Get-Content $files[0].FullName

    # Process all records
    foreach ($roltoewijzing in $xml.GetElementsByTagName("roltoewijzing")) {
        $manager = [PSCustomObject]@{}

        foreach ($child in $roltoewijzing.ChildNodes) {
            $manager | Add-Member -MemberType NoteProperty -Name $child.LocalName -Value $child.'#text' -Force
        }

        # Make sure the manager meets the criteria
        if ($manager.oeRolCode -notin $ManagerRoleCodes) { continue }

        if ([DateTime]::ParseExact($manager.begindatum, "yyyy-MM-dd", $null) -ge (Get-Date)) { continue }

        if ([string]::IsNullOrEmpty($manager.einddatum) -eq $false) {
            if ([DateTime]::ParseExact($manager.einddatum, "yyyy-MM-dd", $null) -le (Get-Date)) { continue }
        }

        [void]$managers.value.Add($manager)
    }
}

Write-Verbose -Verbose "[Departments] Import started"

# Init variables
$connectionSettings = ConvertFrom-Json $configuration

$xmlPath = $($connectionSettings.xmlPath)

# Get the source data
$departments = New-Object System.Collections.ArrayList
$managers = New-Object System.Collections.ArrayList
Get-RAETXMLDepartments -XMLBasePath $xmlPath -FileFilter "rst_orgeenheid_*.xml" ([ref]$departments)
Get-RAETXMLManagers -XMLBasePath $xmlPath -FileFilter "roltoewijzing_*.xml" -ManagerRoleCodes @("MGR") ([ref]$managers)

# Group managers per OE
$managers = $managers | Group-Object orgEenheidID -AsHashTable

# Extend the departments with required and additional fields
$departments | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$departments | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
$departments | Add-Member -MemberType NoteProperty -Name "Name" -Value $null -Force
$departments | Add-Member -MemberType NoteProperty -Name "ManagerExternalId" -Value $null -Force
$departments | Add-Member -MemberType NoteProperty -Name "ParentExternalId" -Value $null -Force

Write-Verbose -Verbose "[Departments] Exporting data to HelloID"
$departments | ForEach-Object {
    $_.ExternalId = $_.orgEenheidID
    $_.DisplayName = $_.NaamLang
    $_.Name = $_.NaamLang
    $_.ParentExternalId = $_.hogereOrgEenheid

    # Add the manager
    $managerObject = $managers[$_.orgEenheidID]
    if ($null -ne $managerObject) {
        if ($managerObject.persnr -is [system.array] ) {
            $_.ManagerExternalId = $managerObject.persnr[0]
            $string = [string]$managerObject.persnr
            Write-Verbose -Verbose "[Departments] Multiple managers found for OE $($_.ExternalId): ($string). Keeping manager $($_.ManagerExternalId)."
        } else {
            $_.ManagerExternalId = $managerObject.persNr
        }
    }

    $json = $_ | ConvertTo-Json -Depth 3

    Write-Output $json
}

Write-Verbose -Verbose "[Departments] Exported data to HelloID"
