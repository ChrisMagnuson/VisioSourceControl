[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

function expand-VisioDocument {
    param (
        [System.IO.FileInfo]$Document,
        [System.IO.DirectoryInfo]$Destination
    )
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Document.FullName, $Destination.FullName)
}

function Compress-FolderIntoVisioDocument {
    param (
        [System.IO.DirectoryInfo]$DirectoryContainingSourceControlledVisioDocument,
        [System.IO.FileInfo]$Document
    )
    [System.IO.Compression.ZipFile]::CreateFromDirectory($DirectoryContainingSourceControlledVisioDocument.FullName, $Document.FullName)
}

function Edit-SourceControlledVisioDocument {
    param (
        [System.IO.DirectoryInfo]$DirectoryContainingSourceControlledVisioDocument
    )
    try {
        [System.IO.FileInfo]$DocumentToCreate = [System.IO.FileInfo]"$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        Compress-FolderIntoVisioDocument -DirectoryContainingSourceControlledVisioDocument $DirectoryContainingSourceControlledVisioDocument -Document $DocumentToCreate  
        Start-Process "C:\Program Files (x86)\Microsoft Office\Office15\VISIO.EXE" -NoNewWindow -Wait -ArgumentList $DocumentToCreate.FullName
        Remove-Item -Recurse $DirectoryContainingSourceControlledVisioDocument
        expand-VisioDocument -Document $DocumentToCreate -Destination $DirectoryContainingSourceControlledVisioDocument
        Remove-Item "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        
        $XmlFiles = gci -Recurse $DirectoryContainingSourceControlledVisioDocument | 
        where extension -eq ".xml"

        $XmlFiles | %{
            $_;
            [xml](gc $_.FullName) | Format-XML | Set-Content -Encoding UTF8 -path $_.fullname
        }
    } catch {
        if (Test-Path -Path "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx") {
            Remove-Item "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        }

        $_.Exception|format-list -force
    }
}

#http://blogs.msdn.com/b/powershell/archive/2008/01/18/format-xml.aspx
function Format-XML {
    param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )][xml]$xml, 
        $indent=2
    )
    process {
        $StringWriter = New-Object System.IO.StringWriter 
        $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
        $xmlWriter.Formatting = "indented" 
        $xmlWriter.Indentation = $Indent 
        $xml.WriteContentTo($XmlWriter) 
        $XmlWriter.Flush() 
        $StringWriter.Flush() 
        Write-Output $StringWriter.ToString() 
    }
}