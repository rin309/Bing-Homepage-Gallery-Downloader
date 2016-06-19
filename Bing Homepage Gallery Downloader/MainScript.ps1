#
# Bing Homepage Gallery Downloader
#

$script:ApplicationTitle = "壁紙の一覧データを扱いやすくする"
$script:LastUpdated = "2016/6/20"
$script:Author = "morokoshidog"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ここから
#Get-Content "$env:UserProfile\Source\Repos\Bing Homepage Gallery Downloader\Bing Homepage Gallery Downloader\Sample\browsedata.json" -Encoding UTF8 -Raw | ConvertFrom-Json
# ここまで

# か

# ここから
$BrowsedataJson = Invoke-WebRequest -Uri http://www.bing.com/gallery/home/browsedata?z=0
$BrowsedataJson = $BrowsedataJson.Content
$BrowsedataJson = $BrowsedataJson.Replace("(function(w,n){var a=w[n]=w[n]||{};a.browseData=","")
$BrowsedataJson = $BrowsedataJson.Replace(";})(window, 'BingGallery');","")
$BrowsedataJson = ConvertFrom-Json [$BrowsedataJson]
# ここまで

0..($BrowsedataJson.imageIds.Count -1) | ForEach {
    $ItemIndex = $_
    $Item = @{}
    $Item["imageIds"] = $BrowsedataJson.imageIds[$ItemIndex]
    $Item["categories"] = $BrowsedataJson.categories[$ItemIndex]
    $Item["tags"] = $BrowsedataJson.tags[$ItemIndex]
    $Item["holidays"] = $BrowsedataJson.holidays[$ItemIndex]
    $Item["regions"] = $BrowsedataJson.regions[$ItemIndex]
    $Item["countries"] = $BrowsedataJson.countries[$ItemIndex]
    $Item["colors"] = $BrowsedataJson.colors[$ItemIndex]
    $Item["shortNames"] = $BrowsedataJson.shortNames[$ItemIndex]
    $Item["imageNames"] = $BrowsedataJson.imageNames[$ItemIndex]
    $Item["dates"] = $BrowsedataJson.dates[$ItemIndex]
    New-Object PSOBject -Property $Item
} | ConvertTo-Json
