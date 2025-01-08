# Ensure the necessary assembly is loaded
Add-Type -AssemblyName System.Windows.Forms

function Get-FolderName {
    # Create an OpenFileDialog object
    $dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        CheckFileExists = $false
        ValidateNames = $false
        FileName = "Select Folder"
    }

    # Show the dialog
    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Get the folder part of the selected path
        return ([System.IO.Path]::GetDirectoryName($dialog.FileName))
    } else {
        return $null
    }
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
