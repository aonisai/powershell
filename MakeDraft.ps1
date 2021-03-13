$tmppath = "hoge"

# 入力によってテンプレートを指名
# stringの1文字に制限
[validateLength(1,1)]$in = [string](Read-Host "テンプレートを選択して下さい。`r`n
    1:`r`n
    2:`r`n
    3:`r`n
    4:`r`n"
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
    "1" {
        $file_name = "01..txt";
        $mailbody = "01.body.txt"
        $subject_dates = $today.ToString("MM/dd")
        }
    "2" {
        $file_name = "02..txt";
        $mailbody = "02.body.txt"
        $subject_dates = $nextweek["mon"].ToString("MM/dd")+ " - " + $nextweek["fri"].ToString("MM/dd")
        }
    "3" {
        $file_name = "03.txt";
        $mailbody = "03_body.txt"
        }
    "4" {
        $file_name = "04.【.txt";
        $mailbody = "04._body.txt"
        $subject_dates = $nextweek["mon"].ToString("MM/dd") + " - " + $nextweek["fri"].ToString("MM/dd")
        }
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
# body以外の読込
foreach($line in $file){
    if($line -match "(?<mail_key>^\w*):(?<mail_value>.*$)"){
        if($Matches.ContainsValue("subject")){
            # $Matches["mail_value"] = $Matches["mail_value"].replace("today", $subject_dates)
            $Matches["mail_value"] += $subject_dates
            $mail.Subject = $Matches["mail_value"]
        }
        elseif($Matches.ContainsValue("recipients_to")){
            $mail.Recipients.Add($Matches["mail_value"]).type = 1
        }elseif($Matches.ContainsValue("recipients_cc")){
            $mail.Recipients.Add($Matches["mail_value"]).type = 2
        }else{
            # Write-Host "match error"
            # $Matches
        }
    }
}
# bodyの読込
if(Test-Path $mailbody){
    # $mail.body = (Get-Content $mailbody) -join "`r`n"
    $i = 0
    $mailbody_f = @()
    foreach($l in Get-Content $mailbody){
        if($l -match "[A-Z]+"){
            $l = $l -replace "[A-Z]+DAY", $nextweek["mon"].AddDays($i).ToString("MM/dd")
            $i++
        }
        $mailbody_f += $l
    }
     $mail.body = $mailbody_f -join "`r`n"
}else{
    Write-Host "mailbody error"
    exit
}

$mail.save()
$inspector = $mail.GetInspector
$inspector.Display()
