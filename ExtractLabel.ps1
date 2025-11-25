param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath
)

# Resolve the file path and confirm it exists
try {
    $fullPath = Resolve-Path -LiteralPath $FilePath -ErrorAction Stop
} catch {
    Write-Error "File not found: $FilePath"
    exit 1
}

# Load ZipArchive support from .NET
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Try to open the Office file as a zip package
try {
    $zipFile = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
} catch {
    Write-Error "Unable to open the file as a zip archive. Confirm it is a valid Office document."
    exit 1
}

# Locate the custom properties file
$customProps = $zipFile.Entries | Where-Object { $_.FullName -eq "docProps/custom.xml" }

if (-not $customProps) {
    Write-Error "No custom.xml file found. This document does not contain Purview metadata."
    $zipFile.Dispose()
    exit 1
}

# Extract custom.xml to a temporary file
$tempFile = Join-Path $env:TEMP ("custom_{0}.xml" -f (Get-Random))
$stream = $customProps.Open()

try {
    $output = [System.IO.File]::Create($tempFile)
    $stream.CopyTo($output)
    $output.Close()
} finally {
    $stream.Close()
    $zipFile.Dispose()
}

# Load and parse the XML with namespaces
[xml]$xml = Get-Content $tempFile

$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("c", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
$ns.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")

# Find all Purview label properties
$labelNodes = $xml.SelectNodes("//c:property[contains(@name, 'MSIP_Label_')]", $ns)

# Convert XML nodes into objects
$results = foreach ($node in $labelNodes) {
    $propName = $node.GetAttribute("name")
    $valueNode = $node.SelectSingleNode("vt:lpwstr", $ns)

    [PSCustomObject]@{
        Name  = $propName
        Value = $valueNode.InnerText
    }
}

# Clean up temp file
Remove-Item $tempFile -Force

# Output results
$results | Format-Table -AutoSize
