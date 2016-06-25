#
# Bing Homepage Gallery Downloader
#

Add-Type -AssemblyName System.Windows.Forms

# URLのメモ
# http://az615200.vo.msecnd.net/site/scripts/app.56d9becb.js
# http://www.bing.com/gallery/home/imagedetails/26843?z=0
# "//az608707.vo.msecnd.net/files/","//az619519.vo.msecnd.net/files/","//az619822.vo.msecnd.net/files/"

function GetContentInfoFromImageIds($imageIds, $SelectedPath){
	$Url = ("http://www.bing.com/gallery/home/imagedetails/%imageIds%?z=0" -replace "%imageIds%", $imageIds)
	$ImageDetailsJson = Invoke-WebRequest -Uri $Url
	$ImageDetailsJson = $ImageDetailsJson.Content
	$ImageDetailsJson = ConvertFrom-Json [$ImageDetailsJson]
	if ($ImageDetailsJson.wallpaper -eq "true"){
		DownalodImageFromImageIds $ImageDetailsJson.wpFullFilename $SelectedPath
	}
}

function DownalodImageFromImageIds($wpFullFilename, $SelectedPath){
	$Url = ("http://az619519.vo.msecnd.net/files/%wpFullFilename%" -replace "%wpFullFilename%", $wpFullFilename)
	$FilePath = ("%SelectedPath%\%wpFullFilename%" -replace "%SelectedPath%", $SelectedPath -replace "%wpFullFilename%", $wpFullFilename)
	$BrowsedataJson = Invoke-WebRequest -Uri $Url -OutFile $FilePath
}

function DownalodThumbnailImageFromImageIds($Item){
	$BrowsedataJson = Invoke-WebRequest -Uri ("http://az619519.vo.msecnd.net/files/%imageNames%_640x360.jpg" -replace "%imageIds%", $imageIds)
}


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

$BasedData = 0..($BrowsedataJson.imageIds.Count -1) | ForEach {
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
} #| ConvertTo-Json

$FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowserDialog.Description = "壁紙の保存先を指定してください。ダウンロードするファイルが {0} となりますので、新しいフォルダーを選択することをおすすめします。" -replace "{0}", $BrowsedataJson.imageIds.Count
$FolderBrowserDialog.SelectedPath = "%UserProfile%\Pictures" -replace "%UserProfile%", $env:USERPROFILE
if($FolderBrowserDialog.ShowDialog() -ne "OK")
{
    exit
}

$BasedData | ForEach {GetContentInfoFromImageIds $_.imageIds $FolderBrowserDialog.SelectedPath}
