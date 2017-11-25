# Take list of installed products and filters out those that don't have source msi files.
# If list is not specified, builds new one asking the system through MsiTools.Products.
function FilterProductsWithSources {
    param(
        [array]$products
    )

    if (!$products) {
        $products = New-Object -TypeName MsiTools.Products
    }

    $reply = @()
    foreach ($product in $products) {
        $msiPath = $product.GetLastUsedSourcePath()
        if ($msiPath) {
            if (Test-Path $msiPath) {
                $reply += $product
            }
        }
    }

    return $reply
}
