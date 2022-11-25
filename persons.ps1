# Get the source data
function Get-RAETXMLFunctions {
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

Function Get-RAETXMLBAFiles {
    param(
        [parameter(Mandatory)]
        $XMLBasePath,

        [parameter(Mandatory)]
        $FileFilter
    )

    # List all files in the selected folder
    Get-ChildItem -Path $XMLBasePath -Filter $FileFilter
}

Function Get-DisplayName {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PScustomObject]
        $Person
    )

    # Map DisplayName (as seen in Raw data)
    $externalID = "$($Person.persNr_identificatiePS)".trim()
    $firstName = "$($Person.roepnaam_P01003)".trim()
    $prefix = "$($Person.voorvoegsels_P00302)".trim()
    $lastName = "$($Person.geboortenaam_P00301)".trim()
    $partnerPrefix = "$($Person.voorvoegsels_P00391)".trim()
    $partnerLastname = "$($Person.geboortenaam_P00390)".trim()

    $nameConvention = $Person.k_naamgebruik.Code

    switch ($nameConvention) {
        'E' {
            # Birthname
            $displayName = $firstName

            if (-not[String]::IsNullOrEmpty($prefix)) {
                $displayName = $displayName + " " + $prefix
            }

            return $displayName + " " + $lastName + " ($externalID)"
        }
        'B' {
            # Partnername - Birthname
            $displayName = $firstName

            if (-not[String]::IsNullOrEmpty($partnerPrefix)) {
                $displayName = $displayName + " " + $partnerPrefix
            }
            $displayName = $displayName + " " + $partnerLastname
            $displayName = $displayName + " -"

            if (-not[String]::IsNullOrEmpty($prefix)) {
                $displayName = "$displayName $prefix"
            }
            $displayName = "$displayName $lastName"
            return "$displayName ($externalID)"
        }
        'P' {
            # Partnername
            $displayName = $firstName

            if (-not[String]::IsNullOrEmpty($partnerPrefix)) { $displayName = $displayName + " " + $partnerPrefix }
            $displayName = $displayName + " " + $partnerLastname

            return $displayName + " ($externalID)"
        }
        'C' {
            # Birthname - Partnername
            $displayName = $firstName

            if (-not[String]::IsNullOrEmpty($prefix)) {
                $displayName = $displayName + " " + $prefix
            }

            $displayName = $displayName + " " + $lastName

            $displayName = $displayName + " -"

            if (-not[String]::IsNullOrEmpty($partnerPrefix)) {
                $displayName = $displayName + " " + $partnerPrefix
            }

            $displayName = $displayName + " " + $partnerLastname

            return $displayName + " ($externalID)"
        }
        default {
            # Birthname
            $displayName = $firstName

            if (-not[String]::IsNullOrEmpty($prefix)) {
                $displayName = $displayName + " " + $prefix
            }

            $displayName = $displayName + " " + $lastName

            return $displayName + " ($externalID)"
        }
    }
}

Function Format-VismaPerson {
    [CmdletBinding()]
    param (
        [Parameter()]
        $XmlPerson
    )

    $person = [PSCustomObject]@{}

    $XmlPerson.GetElementsByTagName("identificatiePS").ChildNodes | ForEach-Object {
        $person | Add-Member -NotePropertyName "$($_.LocalName)_identificatiePS" -NotePropertyValue $_.'#text' -Force
    }

    $XmlPerson.GetElementsByTagName("rubriekcode") | ForEach-Object {
        $person | Add-Member -NotePropertyName "$($_.ParentNode.LocalName)_$($_.InnerText)" -NotePropertyValue $_.ParentNode.waarde -Force
    }

    $person
}

Function Format-VismaContract {
    [CmdletBinding()]
    param (
        [Parameter()]
        $XmlContract,

        [Parameter()]
        $Functions,

        [Parameter()]
        $Departments,

        [Parameter()]
        $FunctionMemberNames,

        [Parameter()]
        $DepartmentMemberNames,

        [parameter()]
        [bool]
        $usePositions = $false
    )

    $contract = [PSCustomObject]@{}

    foreach ($id in $XmlContract.GetElementsByTagName("identificatieDV").ChildNodes) {
        $contract | Add-Member -NotePropertyName ("dv_" + $id.LocalName + "_identificatieDV") -NotePropertyValue $id.'#text' -Force
    }

    foreach ($node in $XmlContract | Get-Member -MemberType Property) {

        if ($node.name -NE 'inroostering' -And $node.name -NE 'loonVerdeling') {

            foreach ($rubriekcode in $XmlContract.($node.Name).GetElementsByTagName("rubriekcode")) {

                $waarde = $rubriekcode.ParentNode.waarde

                $contract | Add-Member -NotePropertyName ("dv_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText) -NotePropertyValue $waarde -Force

                # Add the value description as a separate property
                if ([string]::IsNullOrEmpty($rubriekcode.ParentNode.omschrijvingWaarde) -eq $false) {
                    $contract | Add-Member -NotePropertyName ("dv_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText + "_desc") -NotePropertyValue $rubriekcode.ParentNode.omschrijvingWaarde -Force
                }
            }
        }
    }

    if ([string]::IsNullOrEmpty($contract.dv_functiePrimair_P01107) -eq $true) {
        foreach ($name in $functionMemberNames) {
            $contract | Add-Member -NotePropertyName ("dv_functiePrimair_P01107_" + $name) -NotePropertyValue "" -Force
        }
    }
    else {
        $contractFunction = $functions[$contract.dv_functiePrimair_P01107]
        foreach ($name in $functionMemberNames) {
            $contract | Add-Member -NotePropertyName ("dv_functiePrimair_P01107_" + $name) -NotePropertyValue $contractFunction."$name" -Force
        }
    }

    <#
	$pattern = '[^a-zA-Z]'
    if ([string]::IsNullOrEmpty($contract.dv_aanvullendeRubriek_E2078 ) -eq $true) {
        foreach ($name in $locationsMemberNames) {
            $contract | Add-Member -NotePropertyName ("dv_aanvullendeRubriek_E2078_" + ($name -replace $pattern, '')) -NotePropertyValue "" -Force
        }
    }
    else
    {
        $contractLocation = $locations[$contract.dv_aanvullendeRubriek_E2078 ]
        foreach ($name in $locationsMemberNames) {
                $contract | Add-Member -NotePropertyName ("dv_aanvullendeRubriek_E2078_" + ($name -replace $pattern, '')) -NotePropertyValue $contractLocation."$name" -Force
        }
    }
    #>

    if ([string]::IsNullOrEmpty($contract.dv_orgEenheid_P01106) -eq $true) {
        foreach ($name in $departmentMemberNames) {
            $contract | Add-Member -NotePropertyName ("dv_dv_orgEenheid_P01106_" + $name) -NotePropertyValue "" -Force
        }
    }
    else {
        $contractDepartment = $departments[$contract.dv_orgEenheid_P01106]
        foreach ($name in $departmentMemberNames) {
            $contract | Add-Member -NotePropertyName ("dv_dv_orgEenheid_P01106_" + $name) -NotePropertyValue $contractDepartment."$name" -Force
        }
    }

    if ($usePositions) {
        # make a contract per inzet
        $inzetRegels = $XmlContract.GetElementsByTagName("inzet")

        foreach ($inzet in $inzetRegels) {

            #clone contract object
            $position = [PSCustomObject]@{}
            foreach ($propery in $contract.psobject.properties) {
                $position | Add-Member -MemberType $propery.MemberType -Name $propery.Name -Value $propery.Value
            }

            foreach ($id in $inzet.GetElementsByTagName("identificatieIZ").ChildNodes) {
                $position | Add-Member -NotePropertyName ("iz_" + $id.LocalName + "_identificatieIZ") -NotePropertyValue $id.'#text' -Force
            }

            # this should be done before the inzet loop & should exclude all nodes in 'inroostering'
            foreach ($rubriekcode in $inzet.GetElementsByTagName("rubriekcode")) {
                $waarde = $rubriekcode.ParentNode.waarde

                $position | Add-Member -NotePropertyName ("iz_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText) -NotePropertyValue $waarde -Force

                # Add the value description as a separate property
                if ([string]::IsNullOrEmpty($rubriekcode.ParentNode.omschrijvingWaarde) -eq $false) {
                    $position | Add-Member -NotePropertyName ("iz_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText + "_desc") -NotePropertyValue $rubriekcode.ParentNode.omschrijvingWaarde -Force
                }
            }

            if ([string]::IsNullOrEmpty($position.iz_operationeleFunctie_P01122) -eq $true) {
                foreach ($name in $FunctionMemberNames) {
                    $contract | Add-Member -NotePropertyName ("iz_operationeleFunctie_P01122_" + $name) -NotePropertyValue "" -Force
                }
            }
            else {
                $contractFunction = $functions[$position.iz_operationeleFunctie_P01122]
                foreach ($name in $FunctionMemberNames) {
                    $position | Add-Member -NotePropertyName ("iz_operationeleFunctie_P01122_" + $name) -NotePropertyValue $contractFunction."$name" -Force
                }
            }

            if ([string]::IsNullOrEmpty($position.iz_operationeleOrgEenheid_P01121 ) -eq $true) {
                foreach ($name in $DepartmentMemberNames) {
                    $contract | Add-Member -NotePropertyName ("iz_operationeleOrgEenheid_P01121_" + $name) -NotePropertyValue "" -Force
                }
            }
            else {
                $contractDepartment = $departments[$position.iz_operationeleOrgEenheid_P01121 ]
                foreach ($name in $DepartmentMemberNames) {
                    $position | Add-Member -NotePropertyName ("iz_operationeleOrgEenheid_P01121_" + $name) -NotePropertyValue $contractDepartment."$name" -Force
                }
            }

            $positionActive = (([DateTime]$position.iz_begindatum_P01125).addDays(0) -le $now -and ([string]::IsNullOrEmpty($position.iz_einddatum_P01126) -or ([DateTime]$position.iz_einddatum_P01126).addDays(0) -gt $now))
            $position | Add-Member -NotePropertyName ("iz_is_active") -NotePropertyValue $positionActive -Force

            $position
        }
    }
    else {
        $contract
    }
}

Write-Verbose -Verbose "Person import started";

# Init variables
$connectionSettings = ConvertFrom-Json $configuration

$xmlPath = $($connectionSettings.xmlPath)
$usePositions = [System.Convert]::ToBoolean($connectionSettings.usePositions)

$FileTime = (Get-ChildItem -File (Join-Path $xmlPath -ChildPath "IAM_BA_*.xml") | Sort-Object -Descending -Property CreationTime | Select-Object -First 1).name.split('_')[2]

Write-Verbose -Verbose "Parsing function file...";
$functions = Get-RAETXMLFunctions -XMLBasePath $xmlPath -FileFilter "rst_functie_$($FileTime)_*.xml"

Write-Verbose -Verbose "Parsing department file...";
$departments = Get-RAETXMLDepartments -XMLBasePath $xmlPath -FileFilter "rst_orgeenheid_$($FileTime)_*.xml"

Write-Verbose -Verbose "Parsing BA files...";
$files = Get-RAETXMLBAFiles -XMLBasePath $xmlPath -FileFilter "IAM_BA_$($FileTime)_*.xml"

$VismaContract = @{
    Functions             = $functions | Group-Object functieCode -AsHashTable
    Departments           = $departments | Group-Object orgEenheidID -AsHashTable
    FunctionMemberNames   = ($functions.GetEnumerator() | Select-Object -First 1 | Get-Member -MemberType NoteProperty).Name
    DepartmentMemberNames = ($departments.GetEnumerator() | Select-Object -First 1 | Get-Member -MemberType NoteProperty).Name
}

$ExternalIDs = [System.Collections.ArrayList]::new()

$Persons = [Collections.Generic.List[PSCustomObject]]::new()

foreach ($file in $files) {
    # Write-verbose -verbose "processing: $($file.FullName)"

    try {
        [xml]$xml = Get-Content $file.FullName

        $employees = $xml.GetElementsByTagName("werknemer")

        foreach ($employee in $employees) {

            if (-Not $employee.HasChildNodes) {
                Write-Verbose -Verbose "Found empty Employee object"
                continue
            }

            $person = Format-VismaPerson -XmlPerson $employee.getElementsByTagName("persoon")

            $XmlContracts = $employee.GetElementsByTagName("dienstverband")

            [array]$Contracts = $XmlContracts | ForEach-Object {
                Format-VismaContract @VismaContract -XmlContract $_
            }

            # controle of er active contracten zijn, tot 180 dagen na uitdienst
            if (-Not ($Contracts | Where-Object {
                        [string]::IsNullOrWhiteSpace($_.dv_einddatum_P00830) -or
                        [datetime]$_.dv_einddatum_P00830 -ge [datetime]::Today.AddDays(-180)
                    })) {
                continue
            }

            if ($ExternalIDs.Contains("$($person.persNr_identificatiePS)")) {
                Write-Verbose -Verbose "Extern nummer met ID '$($person.persNr_identificatiePS)' is dubbel in de export ($(Get-DisplayName -Person $person))"

                continue
            }

            [void]$ExternalIDs.Add("$($person.persNr_identificatiePS)")

            $person | Add-Member -NotePropertyMembers @{
                ExternalId     = $person.persNr_identificatiePS
                EmployeeNumber = ($Contracts | Select-Object -First 1)[0].dv_opdrachtgeverNr_P01103 + '_' + $person.persNr_identificatiePS
                DisplayName    = Get-DisplayName -Person $person
                Contracts      = $Contracts
            }

            ## WRITE OUTPUT HIER
            # Write-Output (
            #     $person | ConvertTo-Json -Depth 10 -Compress
            # )

            $Persons.add($person)
        }
    }
    catch {
        throw "Error gevonden in bestand '$($file.FullName)' ($($_))"
    }
}

Write-Verbose -Verbose "Exporting Data"

Write-Output (
    $Persons | ConvertTo-Json -Depth 10 -Compress
)

try {
    # clean-up old files
    $limit = (Get-Date).AddDays(-8)

    # Delete files older than the $limit.
    Get-ChildItem -Path $xmlPath -Force | Where-Object {
        -Not $_.PSIsContainer -and $_.CreationTime -lt $limit
    } | Remove-Item -Force
}
catch {
    Write-Verbose -Verbose "oude bestanden konden niet verwijderd worden... morgen weer een kans"
}
