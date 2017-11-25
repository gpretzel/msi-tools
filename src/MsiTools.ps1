function Get-LastUsedSourcePath {
    param(
        [MsiTools.InstalledProduct]$product
    )

    $msiPath = $product.GetLastUsedSourcePath()
    if ($msiPath -And ( Test-Path $msiPath )) {
        $msiPath
    } else {
        $null
    }
}


function Get-CachedMsiPath {
    param(
        [MsiTools.InstalledProduct]$product
    )

    $msiPath = $product.GetProductInfo([MsiTools.InstallProperty]::INSTALLPROPERTY_LOCALPACKAGE)
    if ($msiPath -And ( Test-Path $msiPath )) {
        $msiPath
    } else {
        $null
    }
}


function Get-PackageCachePath {
    [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonApplicationData) | Join-Path -ChildPath "Package Cache"
}


#
# Returns list of source msi files that are likely to be safe for deletion.
# All condition for such msi files should be met:
#   1. Source msi file should not be located in "C:\ProgramData\Package Cache" or its subdirectory.
#   2. Installed product should have existing cached msi file.
#
function Get-LikelySafeToDeleteMSIs {
    $cacheDir = Get-PackageCachePath

    $products = New-Object -TypeName MsiTools.Products

    $reply = @()
    foreach ($product in $products) {
        $msiPath = Get-LastUsedSourcePath $product
        if ($msiPath -And (Get-CachedMsiPath $product) -And -not $msiPath.StartsWith($cacheDir, "CurrentCultureIgnoreCase")) {
            $reply += $msiPath
        }
    }

    $reply | sort
}
