# Get the source data
function Get-RAETXMLFunctions {
    param(
        [parameter(Mandatory = $true)]$XMLBasePath,
        [parameter(Mandatory = $true)]$FileFilter,
        [parameter(Mandatory = $true)][ref]$functions
    )

    $files = Get-ChildItem -Path $XMLBasePath -Filter $FileFilter | Sort-Object LastWriteTime -Descending
    if ($files.Count -eq 0) { return }

    # Read content as XML
    [xml]$xml = Get-Content $files[0].FullName

    # Process all records
    foreach ($functie in $xml.GetElementsByTagName("functie")) {
        $function = [PSCustomObject]@{}

        foreach ($child in $functie.ChildNodes) {
            $function | Add-Member -MemberType NoteProperty -Name $child.LocalName -Value $child.'#text' -Force
        }

        [void]$functions.value.Add($function)
    }
}

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

function Get-RAETXMLBAFiles {
    param(
        [parameter(Mandatory = $true)]$XMLBasePath,
        [parameter(Mandatory = $true)]$functions,
        [parameter(Mandatory = $true)]$locations,
        [parameter(Mandatory = $true)]$departments,
        [parameter(Mandatory = $true)][ref]$persons,
        [parameter(Mandatory = $true)][ref]$contracts
    )
    
    # set date
    $now = New-Object "System.DateTime" -ArgumentList (Get-Date).Year, (Get-Date).Month, (Get-Date).Day

    # Get function member names
    $functionsFirstRecord = $functions.GetEnumerator() | Select-Object -first 1
    $functionMemberNames = ($functionsFirstRecord|Get-Member -MemberType NoteProperty).Name

    # Group functions on externalId
    $functions = $functions | Group-Object functieCode -AsHashTable

    # Get department member names
    $departmentsFirstRecord = $departments.GetEnumerator() | Select-Object -first 1
    $departmentMemberNames = ($departmentsFirstRecord|Get-Member -MemberType NoteProperty).Name

    # Group departments on externalId
    $departments = $departments | Group-Object orgEenheidID -AsHashTable
    
    # Get location member names
    $locationsFirstRecord = $locations.GetEnumerator() | Select-Object -first 1
    $locationsMemberNames = ($locationsFirstRecord|Get-Member -MemberType NoteProperty).Name
    
    # Group locations on externalId
    $locations = $locations | Group-Object externalId -AsHashTable

    # List all files in the selected folder
    $files = Get-ChildItem -Path $XMLBasePath -Filter "*.xml"

    # Process all files
    $count = 1
    foreach ($file in $files) {
        [xml]$xml = Get-Content $file.FullName

        foreach ($werknemer in $xml.GetElementsByTagName("persoon")) {
            $person = [PSCustomObject]@{}

            foreach ($id in $werknemer.GetElementsByTagName("identificatiePS").ChildNodes) {
                $person | Add-Member -MemberType NoteProperty -Name ($id.LocalName + "_identificatiePS") -Value $id.'#text' -Force
            }

            foreach ($rubriekcode in $werknemer.GetElementsByTagName("rubriekcode")) {
                $person | Add-Member -MemberType NoteProperty -Name ($rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText) -Value $rubriekcode.ParentNode.waarde -Force
            }

            [void]$persons.value.Add($person)
        }

        foreach ($dienstverband in $xml.GetElementsByTagName("dienstverband")) {
            
            $contract = [PSCustomObject]@{}
			
            foreach ($id in $dienstverband.GetElementsByTagName("identificatieDV").ChildNodes) {
                $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + $id.LocalName + "_identificatieDV") -Value $id.'#text' -Force
            }
   
            foreach ($node in $dienstverband | Get-Member -MemberType Property) {
           
                if ($node.name -NE 'inroostering' -And $node.name -NE 'loonVerdeling') {

                    foreach ($rubriekcode in $dienstverband.($node.Name).GetElementsByTagName("rubriekcode")) {
                    
                        $waarde = $rubriekcode.ParentNode.waarde
                           
                        $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText) -Value $waarde -Force

                        # Add the value description as a separate property
                        if ([string]::IsNullOrEmpty($rubriekcode.ParentNode.omschrijvingWaarde) -eq $false) {
                            $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText + "_desc") -Value $rubriekcode.ParentNode.omschrijvingWaarde -Force
                        }
                    }
                }
            }
             
            if ([string]::IsNullOrEmpty($contract.dv_functiePrimair_P01107) -eq $true) {
                foreach ($name in $functionMemberNames) {
                    $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "functiePrimair_P01107" + "_" + $name) -Value "" -Force
                }
            }
            else
            {
                $contractFunction = $functions[$contract.dv_functiePrimair_P01107]
                foreach ($name in $functionMemberNames) {
                     $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "functiePrimair_P01107" + "_" + $name) -Value $contractFunction."$name" -Force
                }
            }
			
			$pattern = '[^a-zA-Z]'
            if ([string]::IsNullOrEmpty($contract.dv_aanvullendeRubriek_E2078 ) -eq $true) {
                foreach ($name in $locationsMemberNames) {
                    $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "aanvullendeRubriek_E2078" + "_" + ($name -replace $pattern, '')) -Value "" -Force
                }
            }
            else
            {
                $contractLocation = $locations[$contract.dv_aanvullendeRubriek_E2078 ]
                foreach ($name in $locationsMemberNames) {
                     $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "aanvullendeRubriek_E2078" + "_" + ($name -replace $pattern, '')) -Value $contractLocation."$name" -Force
                }
            }

            if ([string]::IsNullOrEmpty($contract.dv_orgEenheid_P01106) -eq $true) {
                foreach ($name in $departmentMemberNames) {
                    $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "dv_orgEenheid_P01106" + "_" + $name) -Value "" -Force
                }
            }
            else
            {
                $contractDepartment = $departments[$contract.dv_orgEenheid_P01106 ]
                foreach ($name in $departmentMemberNames) {
                     $contract | Add-Member -MemberType NoteProperty -Name ("dv_" + "dv_orgEenheid_P01106" + "_" + $name) -Value $contractDepartment."$name" -Force
                }
            }

			if($usePositions)
			{
				# make a contract per inzet
				foreach ($inzet in $dienstverband.GetElementsByTagName("inzet")) {
					
					#clone contract object
					$position = [PSCustomObject]@{}
					foreach ($propery in $contract.psobject.properties) {
						$position | Add-Member -MemberType $propery.MemberType -Name $propery.Name -Value $propery.Value
					}

					foreach ($id in $inzet.GetElementsByTagName("identificatieIZ").ChildNodes) {
						$position | Add-Member -MemberType NoteProperty -Name ("iz_" + $id.LocalName + "_identificatieIZ") -Value $id.'#text' -Force
					}
					
					# this should be done before the inzet loop & should exclude all nodes in 'inroostering'
					foreach ($rubriekcode in $inzet.GetElementsByTagName("rubriekcode")) {
						$waarde = $rubriekcode.ParentNode.waarde

						$position | Add-Member -MemberType NoteProperty -Name ("iz_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText) -Value $waarde -Force

						# Add the value description as a separate property
						if ([string]::IsNullOrEmpty($rubriekcode.ParentNode.omschrijvingWaarde) -eq $false) {
							$position | Add-Member -MemberType NoteProperty -Name ("iz_" + $rubriekcode.ParentNode.LocalName + "_" + $rubriekcode.InnerText + "_desc") -Value $rubriekcode.ParentNode.omschrijvingWaarde -Force
						}
					}
				
					if ([string]::IsNullOrEmpty($position.iz_operationeleFunctie_P01122) -eq $true) {
						foreach ($name in $FunctionMemberNames) {
							$contract | Add-Member -MemberType NoteProperty -Name ("iz_" + "operationeleFunctie_P01122" + "_" + $name) -Value "" -Force
						}
					}
					else
					{
						$contractFunction = $functions[$position.iz_operationeleFunctie_P01122]
						foreach ($name in $FunctionMemberNames) {
							$position | Add-Member -MemberType NoteProperty -Name ("iz_" + "operationeleFunctie_P01122" + "_" + $name) -Value $contractFunction."$name" -Force
						}
					}

					if ([string]::IsNullOrEmpty($position.iz_operationeleOrgEenheid_P01121 ) -eq $true) {
						foreach ($name in $DepartmentMemberNames) {
							$contract | Add-Member -MemberType NoteProperty -Name ("iz_" + "operationeleOrgEenheid_P01121" + "_" + $name) -Value "" -Force
						}
					}
					else
					{
						$contractDepartment = $departments[$position.iz_operationeleOrgEenheid_P01121 ]
						foreach ($name in $DepartmentMemberNames) {
							$position | Add-Member -MemberType NoteProperty -Name ("iz_" + "operationeleOrgEenheid_P01121" + "_" + $name) -Value $contractDepartment."$name" -Force
						}
					}
		  
					$positionActive = (([DateTime]$position.iz_begindatum_P01125).addDays(0) -le $now -and ([string]::IsNullOrEmpty($position.iz_einddatum_P01126) -or ([DateTime]$position.iz_einddatum_P01126).addDays(0) -gt $now))
					$position | Add-Member -MemberType NoteProperty -Name ("iz_" + "is_active") -Value $positionActive -Force
	  
					[void]$contracts.value.Add($position)
				}
			}
			else
			{
				[void]$contracts.value.Add($contract)
			}

        }
        $count += 1
    }
}

Write-Verbose -Verbose "Person import started";

# Init variables
$connectionSettings = ConvertFrom-Json $configuration

$xmlPath = $($connectionSettings.xmlPath)
$usePositions = [System.Convert]::ToBoolean($connectionSettings.usePositions)
$locationCsv = $($connectionSettings.locationCsv)

# Get the source data
$persons = New-Object System.Collections.ArrayList
$contracts = New-Object System.Collections.ArrayList
$functions = New-Object System.Collections.ArrayList
$departments = New-Object System.Collections.ArrayList
$locations = Import-Csv -Path $locationCsv -Delimiter ";"

Write-Verbose -Verbose "Parsing function file...";
Get-RAETXMLFunctions -XMLBasePath $xmlPath -FileFilter "rst_functie_*.xml" ([ref]$functions)

Write-Verbose -Verbose "Parsing department file...";
Get-RAETXMLDepartments -XMLBasePath $xmlPath -FileFilter "rst_orgeenheid_*.xml" ([ref]$departments)

Write-Verbose -Verbose "Parsing person/contracts files...";
Get-RAETXMLBAFiles -XMLBasePath $xmlPath $functions $locations $departments ([ref]$persons) ([ref]$contracts)

# Group contracts on externalId
$contracts = $contracts | Group-Object dv_persNrDV_identificatieDV -AsHashTable

# Augment the persons
Write-Verbose -Verbose "Augmenting persons...";
$persons | Add-Member -MemberType NoteProperty -Name "Contracts" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
if($usePositions)
{
	$persons | Add-Member -MemberType NoteProperty -Name "AllActiveFunctions" -Value $null -Force
	$persons | Add-Member -MemberType NoteProperty -Name "AllActiveDepartments" -Value $null -Force
}
$persons | ForEach-Object {
    # Map required fields
    $_.ExternalId = $_.persNr_identificatiePS
    $_.DisplayName = $_.persNr_identificatiePS

    # Add the contracts
    $personContracts = $contracts[$_.persNr_identificatiePS]
    if ($null -ne $personContracts) {
        $_.Contracts = $personContracts
    }
	
	if($usePositions)
	{
		# add all active functions and departments to a new person attribute
		$activePersonContracts = $personContracts | Where-Object {$_.iz_is_active -eq $True}
		foreach ($activePersonContract in $activePersonContracts) {
			if (![string]::IsNullOrEmpty($_.AllActiveFunctions))
			{
				$_.AllActiveFunctions = $_.AllActiveFunctions + " -- "    
			}
			$_.AllActiveFunctions = $_.AllActiveFunctions + $activePersonContract.iz_operationeleFunctie_P01122_desc 
			
			if (![string]::IsNullOrEmpty($_.AllActiveDepartments))
			{
				$_.AllActiveDepartments = $_.AllActiveDepartments + " -- "    
			}
			$_.AllActiveDepartments = $_.AllActiveDepartments + $activePersonContract.iz_operationeleOrgEenheid_P01121_naamLang 
		}
		if (![string]::IsNullOrEmpty($_.AllActiveFunctions)) { $_.AllActiveFunctions = $_.AllActiveFunctions | Select-Object -Unique }
		if (![string]::IsNullOrEmpty($_.AllActiveDepartments)) { $_.AllActiveDepartments = $_.AllActiveDepartments | Select-Object -Unique }
	}
}

# Make sure persons are unique
$persons = $persons | Sort-Object ExternalId -Unique
Write-Verbose -Verbose "Person import completed";
Write-Verbose -Verbose "Exporting data to HelloID";

# Output the json
foreach ($Person in $persons) {
    $json = $person | ConvertTo-Json -Depth 3
    Write-Output $json
}

Write-Verbose -Verbose "Exported data to HelloID";
