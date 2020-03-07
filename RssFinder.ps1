Param(
    [string]$search_word = "indicators*\s*of\s*compromise",
    [string]$search_word_02 = "ioc",
    [int]$search_days_arg = 30, # 何日前まで取得するか指定, デフォルトは30日
    [parameter(mandatory=$true)]$rss_list_arg # 調査rssのリスト(txt), 必須
)

$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8

function Find-Match{
    # 日付の比較
    If($search_days -le $date){
        $wc_all = $wc.DownloadString($item.link.Trim())
        if($wc_all | Select-String -Pattern $search_word){
            Write-Host "Yes!!"
           "url:" + $item.link.Trim()
           "title:" + $item.title
        }
        elseif($wc_all | Select-String -Pattern $search_word_02){
            Write-Host "ioc!!"
            "url:" + $item.link.Trim()
            "title:" + $item.title
        }
    }
}

[String]$encoding='UTF-8'
# 各rssを配列で格納
[string[]]$rss_list = (Get-Content $rss_list_arg) -as [string[]]

# $search_daysがデフォルトかどうか
If([string]::IsNullOrEmpty($search_days_arg)){
    # デフォルトの場合30日前に設定
}else{
    # 指定があった場合, 実行日時から計算して何日になるか計算
    $search_days = (Get-Date).AddDays(-$search_days_arg)
}

# rssリストのループ
foreach($rss in $rss_list){
    $rss = [xml]$wc.DownloadString($rss)
    if($rss.rss -ne $null){ # rss=2.0
        "========================================"
        "blog_title:" + $rss.rss.channel.title + ""
        foreach($item in $rss.rss.channel.item){
            # 日付の比較
            $date = [DateTime]$item.pubDate
            Find-Match
<#            If($search_days -le $date){
                $wc_all = $wc.DownloadString($item.link.Trim())
                if($wc_all | Select-String -Pattern $search_word){
                    Write-Host "Yes!!"
                    "url:" + $item.link.Trim()
                    "title:" + $item.title
                }
                elseif($wc_all | Select-String -Pattern $search_word_02){
                    Write-Host "ioc!!"
                    "url:" + $item.link.Trim()
                    "title:" + $item.title
                }
            } #>
        }
    }
    else{ # rss=1.0
        "blog_title:" + $rss.RDF.channel.title + ""
        # rssのitemのループ
        foreach($item in $rss.RDF.item){
            # 日付の比較
            $date = [DateTime]$item.date
            Find-Match
<#            If($search_days -le $date){
                $wc_all = $wc.DownloadString($item.link.Trim())
                if($wc_all | Select-String -Pattern $search_word){
                    Write-Host "Yes!!"
                    "url:" + $item.link.Trim()
                    "title:" + $item.title
                }
                elseif($wc_all | Select-String -Pattern $search_word_02){
                    Write-Host "ioc!!"
                    "url:" + $item.link.Trim()
                    "title:" + $item.title

                }
            } #>
        }
    }
    "========================================"
}
