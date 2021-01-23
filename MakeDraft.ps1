$tmppath = "hoge"

# 入力によってテンプレートを指名
# stringの1文字に制限
[validateLength(1,1)]$in = [string](Read-Host "テンプレートを選択して下さい。`r`n
    1(テレワーク開始連絡):`r`n
    2(出勤申請):`r`n
    3(テレワーク終了連絡):`r`n
    4(事前残業申請):`r`n"
    )

# 今日の日付取得
$today = Get-Date
# 来週の日付取得
$nextweek = @{mon = ""; fri = ""}
$nextweek["fri"] = $today.AddDays([int][dayofweek]::Friday - $today.DayOfWeek)    
if ($today -ge $nextweek["fri"]) {
    $nextweek["fri"] = $nextweek["fri"].AddDays(7)
}
$nextweek["mon"] = $nextweek["fri"].AddDays(-4)

switch($in)
{
    "1" {$file_name = "a.txt"}
    "2" {$file_name = "b.txt"}
    "3" {$file_name = "c.txt"}
    "4" {$file_name = "d.txt"}
    default {Write-Host "input error"; exit}
}

# file読込
cd $tmppath
if(Test-Path $file_name){
    $file = $(Get-Content $file_name)
}else{
    Write-Host "file_name error"
    exit
}

# 下書き作成
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.createItem(0)
foreach($line in $file){
    if($line -match "(?<mail_key>^\w*):(?<mail_value>.*$)"){
        if($Matches.ContainsValue("subject")){
            $mail.Subject = $Matches["mail_value"] + " " + $today.ToString("MM/dd")
        }elseif($Matches.ContainsValue("body")){
            $mail.body += $Matches["mail_value"]
            $mail.body
            $Matches["mail_value"]
            Write-output $Matches["mail_value"]
            Write-Host $Matches["mail_value"]
        }elseif($Matches.ContainsValue("recipients_to")){
            $mail.Recipients.Add($Matches["mail_value"]).type = 1
        }elseif($Matches.ContainsValue("recipients_cc")){
            $mail.Recipients.Add($Matches["mail_value"]).type = 2
        }else{
            Write-Host "match error"
            $Matches
        }
    }
}

$mail.save()
$inspector = $mail.GetInspector
$inspector.Display()