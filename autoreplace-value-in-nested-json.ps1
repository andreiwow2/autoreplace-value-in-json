function Get-FolderName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [string]$Message = "Select a directory.",

        [string]$InitialDirectory = [System.Environment+SpecialFolder]::MyComputer,

        [switch]$ShowNewFolderButton
    )

    $browserForFolderOptions = 0x00000041                                  # BIF_RETURNONLYFSDIRS -bor BIF_NEWDIALOGSTYLE
    if (!$ShowNewFolderButton) { $browserForFolderOptions += 0x00000200 }  # BIF_NONEWFOLDERBUTTON

    $browser = New-Object -ComObject Shell.Application
    # To make the dialog topmost, you need to supply the Window handle of the current process
    [intPtr]$handle = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle

    # see: https://msdn.microsoft.com/en-us/library/windows/desktop/bb773205(v=vs.85).aspx
    $folder = $browser.BrowseForFolder($handle, $Message, $browserForFolderOptions, $InitialDirectory)

    $result = $null
    if ($folder) { 
        $result = $folder.Self.Path 
    } 

    # Release and remove the used Com object from memory
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($browser) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()


    return $result
}

$folder = Get-FolderName
if ($folder) { 
    Get-ChildItem -Path $folder -Filter *.json -Recurse -File | ForEach-Object {
        $filename = $_.FullName
        $a = Get-Content $filename -raw | ConvertFrom-Json
		$a.inventurBezeichnung = 11
		
		if ($a.bestandsErfassung -and $a.bestandsErfassung -is [System.Array]) {
			foreach($item in $a.bestandsErfassung) {
				if ($item.PSObject.Properties.Name -contains 'erfasserKurzzeichen') {
					$item.erfasserKurzzeichen = 5
				}
				
				if ($item.zaehlSaetze -and $item.zaehlSaetze -is [System.Array]) {
					foreach ($product in $item.zaehlSaetze) {
						if ($product.PSObject.Properties.Name -contains 'lagerplatzBezeichnung') {
							$product.lagerplatzBezeichnung = $product.lagerplatzBezeichnung -replace '-','/'
						}

      						if ($product.PSObject.Properties.Name -contains 'artikelNr') {
						    if (-not $product.artikelNr.EndsWith('..')) {
			                                $product.artikelNr += '..'
			                            }
						}
					}
				}
			}
		}
		
        $a | ConvertTo-Json -Depth 100 | Out-File $filename -Encoding UTF8
        Write-Host "File $($filename).json was changed."
    }
 }
else { "You did not select a directory." }
