$tmppath = ""

# 入力によってテンプレートを指名
# stringの1文字に制限
[validateLength(1,1)]$in = [string](Read-Host "テンプレートを選択して下さい。`r`n
    a(残業申請):`r`n
    b(テレワーク申請):`r`n
    c():`r`n")
    
[string]$file_name
switch($in)
{
    "a" {$file_name = "a.txt"}
    "b" {$file_name = "b.txt"}
    "c" {$file_name = "c.txt"}
    default {Write-Host ("default")}
}

# file読込
cd $tmppath
if(Test-Path $file_name){
    $file = $(Get-Content $file_name)
}else{
    Write-Host "file_name error"
    exit
}

テンプレ―トを連想配列に格納
$mail_list = @{}
foreach($line in $file){
    if($line.contains(":")){
        $tmp_list = $line.split(":")
        $mail_list[$tmp_list[0]] = $tmp_list[1]
        Write-Host $tmp_list[0] "," $mail_list[$tmp_list[0]]
    }else{
        Write-Host "error #1"
    }
}
