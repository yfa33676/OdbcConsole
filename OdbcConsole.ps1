#-------------------------------------------------------------------
# �g����
#-------------------------------------------------------------------
#
# 0. ���ݒ�
# �f�[�^�\�[�X��
  $DSN = ""
#
# ���[�U�[��
  $UID = ""
#
# �p�X���[�h
  $PWD = ""
#
# 1. ����ps1�t�@�C����PowerShell�Ŏ��s
#
# 2. [�f�[�^�\�[�X��]: ���o�͂����̂ŃR�}���h�����
$help = "
      select ...        : SELECT�������s
      update ...        : UPDATE�������s
      insert ...        : INSERT�������s
      delete ...        : DELETE�������s
      clear �܂��� cls  : ��ʂ��N���A����
      csv               : ���O�̌��ʂ�CSV�t�@�C��(SJIS)�ɏo��
      clip              : ���O�̌��ʂ��N���b�v�{�[�h�ɃR�s�[(�^�u��؂�)
      tables            : �e�[�u���ꗗ���o��
      columns           : �J�����ꗗ���o��
      database          : DB�����o��
      sql               : SQL�t�@�C��(SJIS)���J���Ď��s
      mode              : ���[�h�ύX(�O���b�h > �R���\�[��(�e�[�u��) > �R���\�[��(���X�g))
      exit �܂��� quit  : �I��
      help              : �R�}���h�ꗗ

  �� F7 �œ��͗����_�C�A���O�AF8 �œ��͗���⊮�AF9 �œ��͗���ԍ��Ăяo��
  �� ESC �œ��̓N���A
"
#-------------------------------------------------------------------
# ��������
#-------------------------------------------------------------------

# �ڑ����J��
try {
  if($DSN -eq ""){
    "�ڑ�����f�[�^�\�[�X��ԍ��œ��͂��Ă�������"
    $DSNList = Get-OdbcDsn
    if ($DSNList.length -eq $null){
      "1:" + $DSNList.Name
      $DSNIndex = Read-Host
      $DSN = $DSNList.Name
    } else {
      (1 .. $DSNList.length) | % {[string]$_ + ":" + ($DSNList | % Name)[$_ - 1]}
      $DSNIndex = Read-Host
      $DSN = (Get-OdbcDsn | % Name)[$DSNIndex - 1]
    }
  }

  if($PWD -eq ""){
    $Credential = Get-Credential -Credential $UID
    $UID = $Credential.UserName
    # �p�X���[�h�𕽕��ɖ߂�
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    $PWD = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
  }
} catch{
  exit
}

try {
  # �ڑ�
  $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=" + $DSN + ";UID=" + $UID + ";PWD=" + $PWD)
  # ���s
  $Con.Open()
} catch{
  if ($Error.Exception.InnerException[0]){
    $Error.Exception.InnerException[0].Message
  }
  pause
  exit
}

# �t�@�C���ۑ��_�C�A���O
Add-Type -AssemblyName System.Windows.Forms
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
$SaveFileDialog.Filter = "CSV�t�@�C��(*.CSV)|*.csv|���ׂẴt�@�C��(*.*)|*.*"
$SaveFileDialog.InitialDirectory = ".\"

# �t�@�C���J���_�C�A���O
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$OpenFileDialog.Filter = "SQL�t�@�C��(*.SQL)|*.sql|���ׂẴt�@�C��(*.*)|*.*"
$OpenFileDialog.InitialDirectory = ".\"

# �G���R�[�f�B���O�iSJIS�j
$OutputEncoding = [console]::OutputEncoding

# �R���\�[�����[�h
$Console = $false

# ���X�g�\�����[�h
$List = $false

#-------------------------------------------------------------------
# �又��
#-------------------------------------------------------------------

while($true){
  # �v�����v�g
  $q = Read-Host [$DSN] | % Trim
  $title = $q
  
  # �I��
  if(($q -eq "exit") -or ($q -eq "quit")){
    break
  }
  
  # �N���A�X�N���[��
  if(($q -eq "clear") -or ($q -eq "cls")){
    Clear-Host
    continue
  }

  # CSV�o��
  if($q -eq "csv"){
    if ($csv -ne $null){
      $SaveFileDialog.Filename = "result.csv"
      if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $csv | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
      }
    }
    continue
  }

  # �N���b�v�{�[�h�ɃR�s�[
  if($q -eq "clip"){
    if ($csv -ne $null){
      $csv | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Set-Clipboard
    }
    continue
  }

  # �e�[�u���ꗗ
  if($q -eq "tables"){
    $Schema = Read-Host "�X�L�[�}��" | % Trim
    try {
      $csv = $Con.GetSchema("Tables", ($Con.Database, $Schema)) | Select-Object TABLE_SCHEM, TABLE_NAME
      if ($Console){
        if ($List){
          $csv | Format-List | Out-Host -Paging
        } else {
          $csv | Format-Table | Out-Host -Paging
        }
      } else{
        $csv | Out-GridView -Title $title
      }
    } catch {
    }
    continue
  }

  # �J�����ꗗ
  if($q -eq "columns"){
    $Schema = Read-Host "�X�L�[�}��" | % Trim
    $Table = Read-Host "�e�[�u����" | % Trim
    try {
      $csv = ($Con.GetSchema("Columns", ($Con.Database, $Schema, $Table)) | Select-Object TABLE_SCHEM, TABLE_NAME, COLUMN_NAME, TYPE_NAME)
      if (!$Console){
        $csv | Out-GridView -Title $title
      } else{
        if (!$List){
          $csv | Format-Table | Out-Host -Paging
        } else {
          $csv | Format-List | Out-Host -Paging
        }
      }
    } catch {
    }
    continue
  }

  # DB��
  if($q -eq "database"){
    $Con.Database
    continue
  }

  # SQL�t�@�C��
  if($q -eq "sql"){
    $OpenFileDialog.Filename = ""
    if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
      $title = $OpenFileDialog.Filename
      $q = Get-Content -Path $OpenFileDialog.Filename
    } else{
      continue
    }
  }

  # ���͂Ȃ�
  if($q -eq ""){
     continue
  }
  
  # ���[�h�`�F���W
  if ($q -eq "mode"){
     if (!$Console -and !$List) {
       $Console = $true
       $List = $false
       "�R���\�[��(�e�[�u��)" | Out-Host
     } elseif ($Console -and !$List) {
       $Console = $true
       $List = $true
       "�R���\�[��(���X�g)" | Out-Host
     } elseif ($Console -and $List) {
       $Console = $false
       $List = $false
       "�O���b�h" | Out-Host
     }
     continue
  }

  # �w���v
  if($q -eq "help"){
     $help
     continue
  }

  # SQL�R�}���h
  $Cmd = New-Object System.Data.Odbc.OdbcCommand
  $Cmd.CommandText = $q
  $Cmd.Connection = $Con

  # SQL���s
  if($q -match "select*"){
    $da  = New-Object System.Data.Odbc.OdbcDataAdapter
    $da.SelectCommand = $Cmd
    $DataSet = New-Object System.Data.DataSet
    try {
      $nRecs = $da.Fill($DataSet)
    } catch{
      $Error.Exception.InnerException[0].Message
      continue
    }
    # �f�[�^�\��
    try {
      $csv = $DataSet.Tables[0]
      if (!$Console){
        $csv | Out-GridView -Title $title
      } else{
        if (!$List){
          $csv | Format-Table | Out-Host -Paging
        } else {
          $csv | Format-List | Out-Host -Paging
        }
      }
    } catch {
    } finally{
      $nRecs | Out-Host
      $nRecs | Out-Null
    }
} else{
    try {
      # ���s
      $Cmd.ExecuteNonQuery()
    } catch{
      $Error.Exception.InnerException[0].Message
      continue
    }
  }
}

#-------------------------------------------------------------------
# �I������
#-------------------------------------------------------------------

# �ڑ������
$Con.Close()
