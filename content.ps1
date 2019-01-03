#
# htmlファイルを読み込み、単純にテキストとして処理
# 文字抽出してうまいことしようとしている
#

# ファイルの読込み、1行ずつ処理
foreach ($l in Get-Content -Path hoeghoge/Version+8084+Content+Release+Notes.html){
    :	
}

# 正規表現で,h3タグを検索
# Select-String -Pattern "<h3>.*</h3>" -InputObject $l | % {$_.Matches.Value }
$h3 = [RegEx]::Matches($l, "<h3>(\w|\s|-)+\(\d+\)<\/h3>")

# 項目名をリストで取得
$sig_category = @()
$sig_number_in_category = @()
$sig_category = [RegEx]::Matches($h3.Value, "[A-Z]([a-zA-Z]|\s|-)+")
$hoge = [RegEx]::Matches($h3.Value, "\(\d+\)")
$sig_number_in_category = [RegEx]::Matches($hoge.Value, "\d+")

# 検索ヒットした結果のインデックスをリストに取得
$index = @()
$h3 | % {
	$index = $h3.Index
}

# indexからh3タグ別にリスト作成
$strings = @()
for($i=0; $i -lt $index.Count; $i++){
    $number = $index[$i]
    # 最後のカテゴリのシグネチャを抽出
	if($i -eq $index.Count-1){
		# $strings.Add($sig_category[$i].Value, $l.SubString($index[$i], $l.Length-$index[$i]))
		$strings += $l.SubString($number, $l.Length-$number)
	}
    else{　 # 最後以外のカテゴリのシグネチャを抽出
		# $strings.Add($sig_category[$i].Value, $l.SubString($index[$i], $index[$i+1]-$index[$i]))
		$strings += $l.SubString($number, $index[$i+1]-$number)
	}
}

# h3のリストからタグ要素の削除
$value_list = @()
foreach($s in $strings){
    $td = [RegEx]::Matches($s, ">((\w|-)+<br\/>)*(\w|\s|-|\.|\/)*<\/td")
	# タブ要素の排除
	foreach($value in $td.Value){
	    $value = $value.Replace(">", "")
	    $value = $value.Replace("<br/", ",")
	    $value_list += [string[]]$value.Replace("</td", "")
    }
}

# エクセルへの書込
# カテゴリ名でシートをわける
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $True
$book = $excel.Workbooks.Add()

$j = 0

# カテゴリーごとのループ処理
foreach($category_name in $sig_category.value){

    $i = 0
    $sig_number = 0 # シグネチャ数のカウント
    $row = 0
    $column = 1

    $excel.WorkSheets.Add()
    $sheet = $excel.WorkSheets.item(1)
    
    # sheetの文字数制限に対応
    if($category_name.Length -ge 32){
        $sheet.name = $category_name.Substring(0,31)
    }
    else{
        $sheet.name = $category_name
    }

    # int型を指定
    [System.Int32] $tmp = $sig_number_in_category[$j].Value
    $tmp = $tmp+1

    # 各要素を入力
    foreach($value in $value_list){
        # シビリティーが来たら次の行に移動
        if($value -match "(critical|informational|high|medium)"){
		    $row++
		    $column = 1
            $sig_number++
        }
        # カテゴリーに含まれるシグネチャ数まで記入したらループを抜ける
        if($sig_number -eq $tmp){
           break
        }

	    $sheet.Cells.Item($row, $column) = $value
	    $column++
        $i++
    }
    $j++
    $value_list[$i]
    $value_list = $value_list[$i..($value_list.Length-1)]
    Write-Host $value_list[0..6]
}

$book.SaveAs("C:\Users\m-oku\Desktop\powershell_test\hoge.xlsx")
# $book.Save()
$excel.Quit()
$excel = $null
[GC]::Collect()
