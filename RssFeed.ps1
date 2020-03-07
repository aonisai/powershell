# Param([String]$url, [String]$encodeing='UTF-8')
[String]$url = "https://wings.msn.to/contents/rss.php"
[String]$encoding='UTF-8'

$cli = New-Object System.Net.WebClient

$cli.Encoding = [System.Text.Encoding]::UTF8
$rss = [xml]$cli.DownloadString($url)

# $doc = [System.Text.Encoding]::GetEncoding($encoding).GetString($cli.DownloadData($url))

"<html><body><ul>"
foreach($item in $rss.RDF.item){
    "<li><a href='" + $item.link.Trim() + "'>"
    [String]$item.title + "</a></li>"
}
"</ul></body></html>"