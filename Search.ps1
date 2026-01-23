#SuperSearch Search Travis Webb V1.2026
#This script searches local folders document files for words (and optionally hyperlink paths) and generate a text file with the names of all the files. The folder must be local. To use this with InSight, you must first sync the libraries to windows explore so that they have a local path.  

function Get-SafeFileName {
    param(
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $escapedChars = [Regex]::Escape(-join $invalidChars)
    $pattern = "[{0}]" -f $escapedChars
    $safeName = [Regex]::Replace($Name, $pattern, "_")
    return $safeName.TrimEnd(' ', '.')
}

function Get-SearchResultsSubject {
    param(
        [string[]]$SearchTerms
    )

    $termsText = $SearchTerms -join ", "
    return "Search results - $termsText - $(Get-Date -Format 'yyyy-MM-dd')"
}

function Get-SearchResultsFileName {
    param(
        [string]$Subject,
        [string]$Extension = ".xlsx"
    )

    $safeSubject = Get-SafeFileName -Name $Subject
    if ([string]::IsNullOrWhiteSpace($safeSubject)) {
        $safeSubject = "Search results"
    }

    if ([string]::IsNullOrWhiteSpace($Extension)) {
        $Extension = ".xlsx"
    }
    if ($Extension[0] -ne ".") {
        $Extension = "." + $Extension
    }

    return "$safeSubject$Extension"
}

function Enable-PreventSleep {
    if (-not ("PowerState" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class PowerState {
    [DllImport("kernel32.dll")]
    public static extern uint SetThreadExecutionState(uint esFlags);
    public const uint ES_CONTINUOUS = 0x80000000;
    public const uint ES_SYSTEM_REQUIRED = 0x00000001;
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;
}
"@
    }

    [PowerState]::SetThreadExecutionState([PowerState]::ES_CONTINUOUS -bor [PowerState]::ES_SYSTEM_REQUIRED -bor [PowerState]::ES_DISPLAY_REQUIRED) | Out-Null
}

function Disable-PreventSleep {
    if ("PowerState" -as [type]) {
        [PowerState]::SetThreadExecutionState([PowerState]::ES_CONTINUOUS) | Out-Null
    }
}

function Release-ComObject {
    param(
        [object]$ComObject
    )

    if ($null -eq $ComObject) {
        return $null
    }

    try {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null
    } catch {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        } catch {
            # ignore release failures
        }
    }

    return $null
}

function Invoke-ComCleanup {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Get-CustomPropertyValue {
    param(
        [object]$Properties,
        [string]$Name
    )

    if ($Properties -eq $null) {
        return ""
    }

    try {
        $property = $Properties.Item($Name)
        if ($property -ne $null -and $property.Value -ne $null) {
            return $property.Value.ToString().Trim()
        }
    } catch {
        return ""
    }

    return ""
}

function Convert-TaxonomyValueToNames {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    $names = @()
    $parts = $Value -split ";#"
    foreach ($part in $parts) {
        if ($part -match "\|") {
            $name = ($part -split "\|")[0].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $names += $name
            }
        }
    }

    if ($names.Count -eq 0) {
        return $Value.Trim()
    }

    return ($names | Select-Object -Unique) -join "; "
}

function Get-FirstXmlValue {
    param(
        [xml]$XmlDoc,
        [string]$LocalName
    )

    if ($XmlDoc -eq $null) {
        return ""
    }

    $node = $XmlDoc.SelectSingleNode("//*[local-name()='$LocalName']")
    if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
        return $node.InnerText.Trim()
    }

    return ""
}

function Get-FirstXmlValueByNames {
    param(
        [xml]$XmlDoc,
        [string[]]$LocalNames
    )

    if ($XmlDoc -eq $null -or $LocalNames -eq $null) {
        return ""
    }

    foreach ($name in $LocalNames) {
        $value = Get-FirstXmlValue -XmlDoc $XmlDoc -LocalName $name
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Get-TaxonomyTermsFromXml {
    param(
        [xml]$XmlDoc,
        [string]$FieldLocalName
    )

    if ($XmlDoc -eq $null) {
        return ""
    }

    $terms = @()
    $nodes = $XmlDoc.SelectNodes("//*[local-name()='$FieldLocalName']//*[local-name()='TermName']")
    foreach ($node in $nodes) {
        $term = $node.InnerText.Trim()
        if (-not [string]::IsNullOrWhiteSpace($term)) {
            $terms += $term
        }
    }

    if ($terms.Count -eq 0) {
        return ""
    }

    return ($terms | Select-Object -Unique) -join "; "
}

function Get-DocxMetadataFromPackage {
    param(
        [string]$FilePath
    )

    $metadata = @{
        DocId = ""
        Summary = ""
        Notes = ""
        Tags = ""
        EnterpriseKeywords = ""
        Author = ""
    }

    try {
        if (-not (Test-Path -Path $FilePath)) {
            return $metadata
        }

        $extension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
        if ($extension -notin @(".docx",".docm",".dotx",".dotm")) {
            return $metadata
        }

        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $archive = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            foreach ($entry in $archive.Entries) {
                if (-not $entry.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $isCustomXml = $entry.FullName.StartsWith("customXml/", [System.StringComparison]::OrdinalIgnoreCase)
                $isCustomProps = $entry.FullName.Equals("docProps/custom.xml", [System.StringComparison]::OrdinalIgnoreCase)
                $isCoreProps = $entry.FullName.Equals("docProps/core.xml", [System.StringComparison]::OrdinalIgnoreCase)
                if (-not $isCustomXml -and -not $isCustomProps -and -not $isCoreProps) {
                    continue
                }

                $stream = $entry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $xmlText = $reader.ReadToEnd()
                } finally {
                    if ($reader) { $reader.Dispose() }
                    if ($stream) { $stream.Dispose() }
                }

                if ([string]::IsNullOrWhiteSpace($xmlText)) {
                    continue
                }

                try {
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    $xmlDoc.LoadXml($xmlText)
                } catch {
                    continue
                }

                if ($isCustomXml) {
                    if ([string]::IsNullOrWhiteSpace($metadata.DocId)) {
                        $metadata.DocId = Get-FirstXmlValueByNames -XmlDoc $xmlDoc -LocalNames @(
                            "Doc_x002d_ID",
                            "Document_x0020_Control",
                            "Document_x002d_Control",
                            "DocumentControl"
                        )
                    }
                    if ([string]::IsNullOrWhiteSpace($metadata.Summary)) {
                        $metadata.Summary = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "Summary"
                    }
                    if ([string]::IsNullOrWhiteSpace($metadata.Notes)) {
                        $metadata.Notes = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "Notes"
                    }
                    if ([string]::IsNullOrWhiteSpace($metadata.Tags)) {
                        $metadata.Tags = Get-TaxonomyTermsFromXml -XmlDoc $xmlDoc -FieldLocalName "MediaServiceAutoTags"
                    }
                    if ([string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords)) {
                        $metadata.EnterpriseKeywords = Get-TaxonomyTermsFromXml -XmlDoc $xmlDoc -FieldLocalName "TaxKeywordTaxHTField"
                    }
                } elseif ($isCustomProps) {
                    foreach ($propertyNode in $xmlDoc.SelectNodes("//*[local-name()='property']")) {
                        $propertyName = $propertyNode.GetAttribute("name")
                        if ([string]::IsNullOrWhiteSpace($propertyName)) {
                            continue
                        }

                        $value = ""
                        if ($propertyNode.ChildNodes.Count -gt 0) {
                            $value = ($propertyNode.ChildNodes.Item(0).InnerText).Trim()
                        }

                        switch -Regex ($propertyName) {
                            "^Doc-?ID$|^Doc_x002d_ID$|^Document Control$" {
                                if ([string]::IsNullOrWhiteSpace($metadata.DocId)) {
                                    $metadata.DocId = $value
                                }
                            }
                            "^Summary$" {
                                if ([string]::IsNullOrWhiteSpace($metadata.Summary)) {
                                    $metadata.Summary = $value
                                }
                            }
                            "^Notes$" {
                                if ([string]::IsNullOrWhiteSpace($metadata.Notes)) {
                                    $metadata.Notes = $value
                                }
                            }
                            "^Tags$|^MediaServiceAutoTags$" {
                                if ([string]::IsNullOrWhiteSpace($metadata.Tags)) {
                                    $metadata.Tags = Convert-TaxonomyValueToNames -Value $value
                                }
                            }
                            "^Enterprise Keywords$|^TaxKeyword$|^TaxKeywordTaxHTField$" {
                                if ([string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords)) {
                                    $metadata.EnterpriseKeywords = Convert-TaxonomyValueToNames -Value $value
                                }
                            }
                        }
                    }
                }

                if ($isCoreProps) {
                    $authorFromCore = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "creator"
                    if ([string]::IsNullOrWhiteSpace($metadata.Author) -and -not [string]::IsNullOrWhiteSpace($authorFromCore)) {
                        $metadata.Author = $authorFromCore
                    }
                    if ([string]::IsNullOrWhiteSpace($metadata.Tags)) {
                        $metadata.Tags = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "keywords"
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($metadata.DocId) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Summary) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Notes) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Tags) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Author)) {
                    break
                }
            }
        } finally {
            if ($archive) { $archive.Dispose() }
        }
    } catch {
        Write-Warning "Failed to read metadata from ${FilePath}: $($_.Exception.Message)"
    }

    return $metadata
}

function Get-DocxMetadataFromCustomXmlParts {
    param(
        [object]$Doc
    )

    $metadata = @{
        DocId = ""
        Summary = ""
        Notes = ""
        Tags = ""
        EnterpriseKeywords = ""
        Author = ""
    }

    if ($Doc -eq $null) {
        return $metadata
    }

    $parts = $null
    try {
        $parts = $Doc.CustomXMLParts
        for ($i = 1; $i -le $parts.Count; $i++) {
            $part = $null
            try {
                $part = $parts.Item($i)
                $xmlText = $part.XML
                if ([string]::IsNullOrWhiteSpace($xmlText)) {
                    continue
                }

                $xmlDoc = $null
                try {
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    $xmlDoc.LoadXml($xmlText)
                } catch {
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($metadata.DocId)) {
                    $metadata.DocId = Get-FirstXmlValueByNames -XmlDoc $xmlDoc -LocalNames @(
                        "Doc_x002d_ID",
                        "Document_x0020_Control",
                        "Document_x002d_Control",
                        "DocumentControl"
                    )
                }
                if ([string]::IsNullOrWhiteSpace($metadata.Summary)) {
                    $metadata.Summary = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "Summary"
                }
                if ([string]::IsNullOrWhiteSpace($metadata.Notes)) {
                    $metadata.Notes = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "Notes"
                }
                if ([string]::IsNullOrWhiteSpace($metadata.Tags)) {
                    $metadata.Tags = Get-TaxonomyTermsFromXml -XmlDoc $xmlDoc -FieldLocalName "MediaServiceAutoTags"
                }
                if ([string]::IsNullOrWhiteSpace($metadata.Tags)) {
                    $metadata.Tags = Get-FirstXmlValue -XmlDoc $xmlDoc -LocalName "Tags"
                }
                if ([string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords)) {
                    $metadata.EnterpriseKeywords = Get-TaxonomyTermsFromXml -XmlDoc $xmlDoc -FieldLocalName "TaxKeywordTaxHTField"
                }

                if (-not [string]::IsNullOrWhiteSpace($metadata.DocId) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Summary) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Notes) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Tags) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords) -and
                    -not [string]::IsNullOrWhiteSpace($metadata.Author)) {
                    break
                }
            } finally {
                $part = Release-ComObject -ComObject $part
            }
        }
    } catch {
        # ignore COM errors
    } finally {
        $parts = Release-ComObject -ComObject $parts
    }

    return $metadata
}

function Get-DocumentMetadata {
    param(
        [object]$Doc,
        [string]$FilePath
    )

    $metadata = @{
        DocId = ""
        Summary = ""
        Notes = ""
        Tags = ""
        EnterpriseKeywords = ""
        Author = ""
    }

    try {
        $customProps = $Doc.CustomDocumentProperties
        $metadata.DocId = Get-CustomPropertyValue -Properties $customProps -Name "Doc-ID"
        if ([string]::IsNullOrWhiteSpace($metadata.DocId)) {
            $metadata.DocId = Get-CustomPropertyValue -Properties $customProps -Name "Doc_x002d_ID"
        }
        if ([string]::IsNullOrWhiteSpace($metadata.DocId)) {
            $metadata.DocId = Get-CustomPropertyValue -Properties $customProps -Name "Document Control"
        }
        $metadata.Summary = Get-CustomPropertyValue -Properties $customProps -Name "Summary"
        $metadata.Notes = Get-CustomPropertyValue -Properties $customProps -Name "Notes"
        $metadata.Tags = Get-CustomPropertyValue -Properties $customProps -Name "Tags"
        $metadata.EnterpriseKeywords = Get-CustomPropertyValue -Properties $customProps -Name "Enterprise Keywords"
        if ([string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords)) {
            $metadata.EnterpriseKeywords = Get-CustomPropertyValue -Properties $customProps -Name "TaxKeyword"
        }
        if (-not [string]::IsNullOrWhiteSpace($metadata.EnterpriseKeywords)) {
            $metadata.EnterpriseKeywords = Convert-TaxonomyValueToNames -Value $metadata.EnterpriseKeywords
        }
    } catch {
        # ignore COM errors
    }

    try {
        if ($Doc -ne $null -and [string]::IsNullOrWhiteSpace($metadata.Tags)) {
            $builtInProps = $Doc.BuiltInDocumentProperties
            $keywordsValue = Get-CustomPropertyValue -Properties $builtInProps -Name "Keywords"
            if (-not [string]::IsNullOrWhiteSpace($keywordsValue)) {
                $metadata.Tags = $keywordsValue
            } else {
                $tagsValue = Get-CustomPropertyValue -Properties $builtInProps -Name "Tags"
                if (-not [string]::IsNullOrWhiteSpace($tagsValue)) {
                    $metadata.Tags = $tagsValue
                }
            }
        }
    } catch {
        # ignore COM errors 
    }

    try {
        if ($Doc -ne $null -and [string]::IsNullOrWhiteSpace($metadata.Author)) {
            $builtInProps = $Doc.BuiltInDocumentProperties
            $authorValue = Get-CustomPropertyValue -Properties $builtInProps -Name "Author"
            if (-not [string]::IsNullOrWhiteSpace($authorValue)) {
                $metadata.Author = $authorValue
            } else {
                $creatorValue = Get-CustomPropertyValue -Properties $builtInProps -Name "Creator"
                if (-not [string]::IsNullOrWhiteSpace($creatorValue)) {
                    $metadata.Author = $creatorValue
                }
            }
        }
    } catch {
        # ignore errors
    }

    $needsFallback = $false
    foreach ($key in $metadata.Keys) {
        if ([string]::IsNullOrWhiteSpace($metadata[$key])) {
            $needsFallback = $true
            break
        }
    }

    if ($needsFallback) {
        $xmlPartsMetadata = Get-DocxMetadataFromCustomXmlParts -Doc $Doc
        $metadataKeys = @($metadata.Keys)
        foreach ($key in $metadataKeys) {
            if ([string]::IsNullOrWhiteSpace($metadata[$key]) -and -not [string]::IsNullOrWhiteSpace($xmlPartsMetadata[$key])) {
                $metadata[$key] = $xmlPartsMetadata[$key]
            }
        }
    }

    $needsPackageFallback = $false
    foreach ($key in $metadata.Keys) {
        if ([string]::IsNullOrWhiteSpace($metadata[$key])) {
            $needsPackageFallback = $true
            break
        }
    }

    if ($needsPackageFallback) {
        $packageMetadata = Get-DocxMetadataFromPackage -FilePath $FilePath
        $metadataKeys = @($metadata.Keys)
        foreach ($key in $metadataKeys) {
            if ([string]::IsNullOrWhiteSpace($metadata[$key]) -and -not [string]::IsNullOrWhiteSpace($packageMetadata[$key])) {
                $metadata[$key] = $packageMetadata[$key]
            }
        }
    }

    return $metadata
}

function Invoke-DocumentSearch {
    param(
        [string]$Path = ([Environment]::GetFolderPath('MyDocuments')),
        [string[]]$FindTerms = @("Search Terms, Comma Seperated"),
        [bool]$MatchCase = $false,
        [bool]$MatchWholeWord = $true,
        [bool]$SearchTextContent = $true,
        [bool]$SearchLinkPaths = $true,
        [ValidateSet("AddressOnly","AddressAndSub","All")]
        [string]$LinkSearchMode = "All",
        [bool]$SearchMetadata = $true,
        [bool]$IncludeMetadataColumns = $true,
        [bool]$IncludeSubfolders = $true,
        [bool]$SearchFileName = $false,
        [int]$DocumentTimeoutSeconds = 120,
        [bool]$PreventSleep = $true,
        [scriptblock]$ShouldStop = $null,
        [ValidateSet("Excel","ExcelTable","Csv")]
        [string]$OutputFormat = "Excel",
        [bool]$SendEmailResults = $false,
        [string[]]$EmailTo = @("user@domain.com"),
        [string]$EmailFrom = "",
        [string]$OutputDirectory = ([Environment]::GetFolderPath('MyDocuments')),
        [string[]]$WordExts = @('.docx','.doc','.docm')
    )

$SearchTerms = @($FindTerms | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ -ne "" } | Select-Object -Unique)
if (-not $SearchTerms -or $SearchTerms.Count -eq 0) {
    Write-Error "No search terms provided. Update `$FindTerms."
    return
}
if (-not $SearchTextContent -and -not $SearchLinkPaths -and -not $SearchMetadata -and -not $SearchFileName) {
    Write-Error "Enable at least one of `$SearchTextContent, `$SearchLinkPaths, `$SearchMetadata, or `$SearchFileName."
    return
}
$SearchTermsText = $SearchTerms -join ", "
$EmailSubject = Get-SearchResultsSubject -SearchTerms $SearchTerms
$OutputExtension = if ($OutputFormat -eq "Csv") { ".csv" } else { ".xlsx" }
$OutputFileName = Get-SearchResultsFileName -Subject $EmailSubject -Extension $OutputExtension
$ResolvedOutputDirectory = $OutputDirectory
if ([string]::IsNullOrWhiteSpace($ResolvedOutputDirectory)) {
    $ResolvedOutputDirectory = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
}
try {
    if (-not (Test-Path -Path $ResolvedOutputDirectory)) {
        New-Item -ItemType Directory -Path $ResolvedOutputDirectory -Force | Out-Null
    }
} catch {
    Write-Error "Unable to create output directory: $ResolvedOutputDirectory"
    return
}
$OutputFilePath = Join-Path -Path $ResolvedOutputDirectory -ChildPath $OutputFileName

if (Test-Path -LiteralPath $OutputFilePath) {
    try {
        Remove-Item -LiteralPath $OutputFilePath -Force
    } catch {
        throw "Output file is in use. Close it and re-run: $OutputFilePath"
    }
}

if ($PreventSleep) {
    Enable-PreventSleep
}

$Word = New-Object -ComObject Word.Application #create word object
$Word.Visible = $false #hides the window
$Word.DisplayAlerts = 0 #disable prompts that block automation
$Word.AutomationSecurity = 3 #disable macros while scanning

function Invoke-ComWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,
        [int]$Retries = 5,
        [int]$DelayMs = 200,
        [string]$ActionName = "COM call"
    )

    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            & $Action
            return $true
        } catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111) { # RPC_E_CALL_REJECTED
                throw
            }
            Start-Sleep -Milliseconds ($DelayMs * ($i + 1))
        }
    }

    Write-Warning "$ActionName failed after $Retries retries."
    return $false
}

function Test-TermMatch {
    param(
        [AllowNull()]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [string]$Term,
        [bool]$MatchCase = $false,
        [bool]$MatchWholeWord = $true
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $false
    }

    if ($MatchWholeWord) {
        $pattern = "(?<!\w)" + [Regex]::Escape($Term) + "(?!\w)"
        $options = if ($MatchCase) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
        return [Regex]::IsMatch($Text, $pattern, $options)
    }

    $comparison = if ($MatchCase) { [System.StringComparison]::Ordinal } else { [System.StringComparison]::OrdinalIgnoreCase }
    return ($Text.IndexOf($Term, $comparison) -ge 0)
}

function Get-LinkCandidates {
    param(
        [string]$Address,
        [string]$SubAddress,
        [string]$DisplayText,
        [ValidateSet("AddressOnly","AddressAndSub","All")]
        [string]$Mode = "All"
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    $baseValues = switch ($Mode) {
        "AddressOnly" { @($Address) }
        "AddressAndSub" { @($Address, $SubAddress) }
        default { @($Address, $SubAddress, $DisplayText) }
    }

    $addCandidate = {
        param([string]$Value)
        if (-not [string]::IsNullOrWhiteSpace($Value) -and -not $candidates.Contains($Value)) {
            $candidates.Add($Value)
        }
    }

    foreach ($value in $baseValues) {
        & $addCandidate $value
    }

    foreach ($value in $baseValues) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $decoded = [System.Uri]::UnescapeDataString($value)
        & $addCandidate $decoded

        if ($decoded -match '^file:///') {
            $trimmed = $decoded -replace '^file:///+', ''
            $trimmed = $trimmed -replace '/', '\'
            & $addCandidate $trimmed
        }

        try {
            $uri = [System.Uri]$value
            if ($uri.IsFile -and -not [string]::IsNullOrWhiteSpace($uri.LocalPath)) {
                & $addCandidate $uri.LocalPath
            }
        } catch {
            # ignore invalid URIs
        }
    }

    return $candidates
}

$MatchResults = @()
try {
    $StopSearch = $false
    $CheckStop = {
        if ($ShouldStop -ne $null -and (& $ShouldStop)) {
            $StopSearch = $true
            return $true
        }
        return $false
    }

    $pathItem = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $pathItem) {
        Write-Error "Search path not found: $Path"
        return
    }

    if ($pathItem.PSIsContainer) {
        if ($IncludeSubfolders) {
            $SearchFiles = Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue
        } else {
            $SearchFiles = Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue
        }
    } else {
        $SearchFiles = @($pathItem)
    }
    $MatchResults = @(
        foreach ($File in ($SearchFiles | Where-Object { $_.Extension -in $WordExts })) { #Foreach doc/docx/docm file in the above folder and add -Resurse after $Path to include subfolders
        if (& $CheckStop) {
            break
        }
        $Doc = $null
        $Content = $null

        try {
            if (& $CheckStop) {
                $StopSearch = $true
                break
            }

            $DocumentStopwatch = $null
            if ($DocumentTimeoutSeconds -gt 0) {
                $DocumentStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            }
            $CheckTimeout = {
                if ($DocumentStopwatch -ne $null -and $DocumentStopwatch.Elapsed.TotalSeconds -ge $DocumentTimeoutSeconds) {
                    throw "Document processing timed out after $DocumentTimeoutSeconds seconds: $($File.FullName)"
                }
            }

            $Doc = $Word.Documents.Open($File.FullName, $false, $true) #Open the document read-only
            $Content = $Doc.Content #get the 'content' object from the document
            $TermMatched = @{}
            $FoundLocations = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($Term in $SearchTerms) {
                $TermMatched[$Term] = $false
            }
            $MatchedCount = 0

            if ($SearchTextContent) {
                foreach ($Term in $SearchTerms) {
                    if (& $CheckStop) { break }
                    & $CheckTimeout
                    $Content.Start = 0
                    $Content.End = $Doc.Content.End
                                                #term,case sensitive,whole word,wildcard,soundslike,synonyms,direction,wrappingmode
                    if ($Content.Find.Execute($Term,$MatchCase,   $MatchWholeWord,$false,  $false,    $false,  $true,    1)){ #execute a search
                        $TermMatched[$Term] = $true
                        $MatchedCount++
                        $FoundLocations.Add("Body") | Out-Null
                    }
                }
            }
            if ($StopSearch) { break }

            if ($SearchLinkPaths) {
                $Hyperlinks = $null
                try {
                    $Hyperlinks = $Doc.Hyperlinks
                    for ($i = 1; $i -le $Hyperlinks.Count; $i++) {
                        if (& $CheckStop) { break }
                        & $CheckTimeout
                        $Hyperlink = $Hyperlinks.Item($i)
                        $Address = $Hyperlink.Address
                        $SubAddress = $Hyperlink.SubAddress
                        $DisplayText = $Hyperlink.TextToDisplay
                        $LinkCandidates = Get-LinkCandidates -Address $Address -SubAddress $SubAddress -DisplayText $DisplayText -Mode $LinkSearchMode

                        foreach ($Term in $SearchTerms) {
                            foreach ($candidate in $LinkCandidates) {
                                if (Test-TermMatch -Text $candidate -Term $Term -MatchCase $MatchCase -MatchWholeWord $MatchWholeWord) {
                                    if (-not $TermMatched[$Term]) {
                                        $TermMatched[$Term] = $true
                                        $MatchedCount++
                                    }
                                    $FoundLocations.Add("Link") | Out-Null
                                    break
                                }
                            }
                        }

                        $Hyperlink = Release-ComObject -ComObject $Hyperlink
                        if ($MatchedCount -eq $SearchTerms.Count) {
                            break
                        }
                    }
                    if ($StopSearch) { break }
                } catch {
                    Write-Warning "Failed to scan hyperlinks in $($File.FullName): $($_.Exception.Message)"
                } finally {
                    $Hyperlinks = Release-ComObject -ComObject $Hyperlinks
                }
            }
            if ($StopSearch) { break }

            if ($SearchFileName) {
                $fileCandidates = @(
                    $File.Name,
                    [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

                foreach ($Term in $SearchTerms) {
                    if (& $CheckStop) { break }
                    & $CheckTimeout
                    foreach ($candidate in $fileCandidates) {
                        if (Test-TermMatch -Text $candidate -Term $Term -MatchCase $MatchCase -MatchWholeWord $MatchWholeWord) {
                            if (-not $TermMatched[$Term]) {
                                $TermMatched[$Term] = $true
                                $MatchedCount++
                            }
                            $FoundLocations.Add("Title") | Out-Null
                            break
                        }
                    }
                }
            }
            if ($StopSearch) { break }

            $Metadata = $null
            if ($SearchMetadata -or $IncludeMetadataColumns) {
                $Metadata = Get-DocumentMetadata -Doc $Doc -FilePath $File.FullName
            }
            if ($Metadata -eq $null) {
                $Metadata = @{
                    DocId = ""
                    Summary = ""
                    Notes = ""
                    Tags = ""
                    EnterpriseKeywords = ""
                    Author = ""
                }
            }

            if ($SearchMetadata) {
                $metadataFields = @()
                if ($Metadata -ne $null) {
                    $metadataFields = @(
                        @{ Name = "Doc-ID"; Value = $Metadata.DocId },
                        @{ Name = "Summary"; Value = $Metadata.Summary },
                        @{ Name = "Notes"; Value = $Metadata.Notes },
                        @{ Name = "Tags"; Value = $Metadata.Tags },
                        @{ Name = "Enterprise Keywords"; Value = $Metadata.EnterpriseKeywords },
                        @{ Name = "Author"; Value = $Metadata.Author }
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) }
                }

                if ($metadataFields.Count -gt 0) {
                    foreach ($Term in $SearchTerms) {
                        if (& $CheckStop) { break }
                        & $CheckTimeout
                        foreach ($field in $metadataFields) {
                            if (Test-TermMatch -Text $field.Value -Term $Term -MatchCase $MatchCase -MatchWholeWord $MatchWholeWord) {
                                if (-not $TermMatched[$Term]) {
                                    $TermMatched[$Term] = $true
                                    $MatchedCount++
                                }
                                $FoundLocations.Add($field.Name) | Out-Null
                            }
                        }
                    }
                }
            }
            if ($StopSearch) { break }

            $MatchedTerms = $SearchTerms | Where-Object { $TermMatched[$_] }

            if ($MatchedTerms.Count -gt 0) {
                $MatchedTermsText = ($MatchedTerms -join "; ")
                $FoundText = ($FoundLocations | Sort-Object) -join "; "
                Write-Host "$($File.Name) contains $MatchedTermsText" -ForegroundColor Green
                $result = [PSCustomObject]@{
                    MatchedTerms = $MatchedTermsText
                    Found = $FoundText
                    FileName = $File.Name
                    FullPath = $File.FullName
                }
                if ($IncludeMetadataColumns) {
                    $result | Add-Member -NotePropertyName DocId -NotePropertyValue ($Metadata.DocId)
                    $result | Add-Member -NotePropertyName Summary -NotePropertyValue ($Metadata.Summary)
                    $result | Add-Member -NotePropertyName Notes -NotePropertyValue ($Metadata.Notes)
                    $result | Add-Member -NotePropertyName Tags -NotePropertyValue ($Metadata.Tags)
                    $result | Add-Member -NotePropertyName EnterpriseKeywords -NotePropertyValue ($Metadata.EnterpriseKeywords)
                    $result | Add-Member -NotePropertyName Author -NotePropertyValue ($Metadata.Author)
                }
                $result
            } else {
                Write-Host "$($File.Name) does not contain any terms" -ForegroundColor Red
            }
        } catch {
            Write-Warning "Skipping $($File.FullName): $($_.Exception.Message)"
        } finally {
            $Content = Release-ComObject -ComObject $Content
            if ($Doc -ne $null) {
                try {
                    Invoke-ComWithRetry -Action { $Doc.Close([ref]$false) | Out-Null } -ActionName "Close document" | Out-Null #close the document
                } catch {
                    Write-Warning "Failed to close $($File.FullName): $($_.Exception.Message)"
                } finally {
                    $Doc = Release-ComObject -ComObject $Doc
                }
                $Doc = $null
            }
        }
    }
    )
} finally {
    if ($Word -ne $null) {
        try {
            Invoke-ComWithRetry -Action { $Word.Quit() | Out-Null } -ActionName "Quit Word" | Out-Null #quit the word process
        } catch {
            Write-Warning "Failed to quit Word: $($_.Exception.Message)"
        }
        $Word = Release-ComObject -ComObject $Word
    }
}

$ColumnValueMap = @{
    "MatchedTerms" = { param($row) $row.MatchedTerms }
    "Found" = { param($row) $row.Found }
    "Doc-ID" = { param($row) $row.DocId }
    "FileName" = { param($row) $row.FileName }
    "FullPath" = { param($row) $row.FullPath }
    "Author" = { param($row) $row.Author }
    "Notes" = { param($row) $row.Notes }
    "Summary" = { param($row) $row.Summary }
    "Tags" = { param($row) $row.Tags }
    "Enterprise Keywords" = { param($row) $row.EnterpriseKeywords }
}

if ($IncludeMetadataColumns) {
    $Headers = @(
        "MatchedTerms",
        "Found",
        "Doc-ID",
        "FileName",
        "FullPath",
        "Author",
        "Notes",
        "Summary",
        "Tags",
        "Enterprise Keywords"
    )
} else {
    $Headers = @("MatchedTerms","Found","FileName","FullPath")
}

if ($OutputFormat -eq "Csv") {
    $EscapeCsv = {
        param($value)
        $text = if ($null -eq $value) { "" } else { $value.ToString() }
        $text = $text -replace '"', '""'
        return '"' + $text + '"'
    }

    $lines = @()
    $lines += ($Headers | ForEach-Object { & $EscapeCsv $_ }) -join ","
    foreach ($Result in $MatchResults) {
        $rowValues = foreach ($header in $Headers) {
            $getter = $ColumnValueMap[$header]
            if ($getter -ne $null) {
                & $getter $Result
            } else {
                ""
            }
        }
        $lines += ($rowValues | ForEach-Object { & $EscapeCsv $_ }) -join ","
    }

    $lines | Set-Content -LiteralPath $OutputFilePath -Encoding UTF8
} else {
    $Excel = $null
    $Workbook = $null
    $Worksheet = $null
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        $Workbook = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Worksheets.Item(1)

        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $Worksheet.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
        }

        $Row = 2
        foreach ($Result in $MatchResults) {
            $rowValues = foreach ($header in $Headers) {
                $getter = $ColumnValueMap[$header]
                if ($getter -ne $null) {
                    & $getter $Result
                } else {
                    ""
                }
            }

            for ($i = 0; $i -lt $rowValues.Count; $i++) {
                $Worksheet.Cells.Item($Row, $i + 1).Value2 = $rowValues[$i]
            }
            $Row++
        }

        if ($OutputFormat -eq "ExcelTable") {
            try {
                $lastRow = [Math]::Max(1, $Row - 1)
                $range = $Worksheet.Range(
                    $Worksheet.Cells.Item(1, 1),
                    $Worksheet.Cells.Item($lastRow, $Headers.Count)
                )
                $table = $Worksheet.ListObjects.Add(1, $range, $null, 1)
                $table.Name = "SearchResults"
                $table.TableStyle = "TableStyleMedium2"
            } catch {
                Write-Warning "Failed to create table: $($_.Exception.Message)"
            }
        }

        $Worksheet.Columns.AutoFit() | Out-Null
        $Workbook.SaveAs($OutputFilePath) | Out-Null
    } finally {
        $Worksheet = Release-ComObject -ComObject $Worksheet
        if ($Workbook -ne $null) {
            try {
                Invoke-ComWithRetry -Action { $Workbook.Close($true) | Out-Null } -ActionName "Close workbook" | Out-Null
            } catch {
                Write-Warning "Failed to close workbook: $($_.Exception.Message)"
            }
            $Workbook = Release-ComObject -ComObject $Workbook
        }
        if ($Excel -ne $null) {
            try {
                Invoke-ComWithRetry -Action { $Excel.Quit() | Out-Null } -ActionName "Quit Excel" | Out-Null
            } catch {
                Write-Warning "Failed to quit Excel: $($_.Exception.Message)"
            }
            $Excel = Release-ComObject -ComObject $Excel
        }
    }
}

if ($SendEmailResults) {
    if (-not $EmailTo -or $EmailTo.Count -eq 0) {
        Write-Error "Email is enabled but `$EmailTo is empty."
    } else {
        $Outlook = $null
        $MailItem = $null
        try {
            $MatchedFileCount = $MatchResults.Count
            $EmailBody = @"
Search results
Terms: $SearchTermsText
Matches: $MatchedFileCount
Output file: $OutputFilePath
"@

            $Outlook = New-Object -ComObject Outlook.Application
            $MailItem = $Outlook.CreateItem(0) # 0 = MailItem
            $MailItem.To = ($EmailTo -join ";")
            if (-not [string]::IsNullOrWhiteSpace($EmailFrom)) {
                $MailItem.SentOnBehalfOfName = $EmailFrom
            }
            $MailItem.Subject = $EmailSubject
            $MailItem.Body = $EmailBody
            $null = $MailItem.Attachments.Add($OutputFilePath)
            $MailItem.Send()
            Write-Host "Email sent to $($EmailTo -join ', ')" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to send email via Outlook: $($_.Exception.Message)"
        } finally {
            $MailItem = Release-ComObject -ComObject $MailItem
            $Outlook = Release-ComObject -ComObject $Outlook
        }
    }
}

Invoke-ComCleanup

if ($PreventSleep) {
    Disable-PreventSleep
}

return $MatchResults
}

if ($MyInvocation.InvocationName -ne '.') {
    return Invoke-DocumentSearch
}


