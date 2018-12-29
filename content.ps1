#
# html�t�@�C����ǂݍ��݁A�P���Ƀe�L�X�g�Ƃ��ď���
# �������o���Ă��܂����Ƃ��悤�Ƃ��Ă���
#

# �t�@�C���̓Ǎ��݁A1�s������
foreach ($l in Get-Content -Path C:/Users/m-oku/Downloads/Version+8084+Content+Release+Notes.html){
    :	
}

# ���K�\����,h3�^�O������
# Select-String -Pattern "<h3>.*</h3>" -InputObject $l | % {$_.Matches.Value }
$h3 = [RegEx]::Matches($l, "<h3>(\w|\s|-)+\(\d+\)<\/h3>")

# ���ږ������X�g�Ŏ擾
$sig_category = @()
$sig_number_in_category = @()
$sig_category = [RegEx]::Matches($h3.Value, "[A-Z]([a-zA-Z]|\s|-)+")
$hoge = [RegEx]::Matches($h3.Value, "\(\d+\)")
$sig_number_in_category = [RegEx]::Matches($hoge.Value, "\d+")

# �����q�b�g�������ʂ̃C���f�b�N�X�����X�g�Ɏ擾
$index = @()
$h3 | % {
	$index = $h3.Index
}

# index����h3�^�O�ʂɃ��X�g�쐬
$strings = @()
for($i=0; $i -lt $index.Count; $i++){
    $number = $index[$i]
    # �Ō�̃J�e�S���̃V�O�l�`���𒊏o
	if($i -eq $index.Count-1){
		# $strings.Add($sig_category[$i].Value, $l.SubString($index[$i], $l.Length-$index[$i]))
		$strings += $l.SubString($number, $l.Length-$number)
	}
    else{�@ # �Ō�ȊO�̂̃J�e�S���̃V�O�l�`���𒊏o
		# $strings.Add($sig_category[$i].Value, $l.SubString($index[$i], $index[$i+1]-$index[$i]))
		$strings += $l.SubString($number, $index[$i+1]-$number)
	}
}

# h3�̃��X�g����^�O�v�f�̍폜
$value_list = @()
foreach($s in $strings){
    $td = [RegEx]::Matches($s, ">((\w|-)+<br\/>)*(\w|\s|-|\.|\/)*<\/td")
	# �^�u�v�f�̔r��
	foreach($value in $td.Value){
	    $value = $value.Replace(">", "")
	    $value = $value.Replace("<br/", ",")
	    $value_list += [string[]]$value.Replace("</td", "")
    }
}

# �G�N�Z���ւ̏���
# �J�e�S�����ŃV�[�g���킯��
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $True
$book = $excel.Workbooks.Add()

$j = 0

# �J�e�S���[���Ƃ̃��[�v����
foreach($category_name in $sig_category.value){

    $i = 0
    $sig_number = 0 # �V�O�l�`�����̃J�E���g
    $row = 0
    $column = 1

    $excel.WorkSheets.Add()
    $sheet = $excel.WorkSheets.item(1)
    
    # sheet�̕����������ɑΉ�
    if($category_name.Length -ge 32){
        $sheet.name = $category_name.Substring(0,31)
    }
    else{
        $sheet.name = $category_name
    }

    # int�^���w��
    [System.Int32] $tmp = $sig_number_in_category[$j].Value
    $tmp = $tmp+1

    # �e�v�f�����
    foreach($value in $value_list){
        # �V�r���e�B�[�������玟�̍s�Ɉړ�
        if($value -match "(critical|informational|high|medium)"){
		    $row++
		    $column = 1
            $sig_number++
        }
        # �J�e�S���[�Ɋ܂܂��V�O�l�`�����܂ŋL�������烋�[�v�𔲂���
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
