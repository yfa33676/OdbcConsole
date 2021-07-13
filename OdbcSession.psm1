function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # �f�[�^�\�[�X��
    [string]$UID = "", # ���[�U�[��
    [string]$PWD = ""  # �p�X���[�h
  )

  begin{
    $help = "
      select ...             : SELECT�������s
      update ...             : UPDATE�������s
      insert ...             : INSERT�������s
      delete ...             : DELETE�������s
      clear �܂��� cls       : ��ʂ��N���A����
      csv                    : ���O�̌��ʂ�CSV�t�@�C��(SJIS)�ɏo��
      clip                   : ���O�̌��ʂ��N���b�v�{�[�h�ɃR�s�[(�^�u��؂�)
      tables �܂��� tbl      : �e�[�u���ꗗ���o��
      columns �܂��� col     : �J�����ꗗ���o��
      views                  : �r���[�ꗗ���o��
      database               : DB�����o��
      sql                    : SQL�t�@�C��(SJIS)���J���Ď��s
      text                   : �����s���̓_�C�A���O
      mode                   : ���[�h�ύX(�O���b�h > �R���\�[��(�e�[�u��) > �R���\�[��(���X�g))
      transaction �܂��� trn : �g�����U�N�V�����̊J�n 
      commit                 : �R�~�b�g
      rollback �܂��� rol    : ���[���o�b�N
      exit �܂��� quit       : �I��
      help                   : �R�}���h�ꗗ

  �� F7 �œ��͗����_�C�A���O�AF8 �œ��͗���⊮�AF9 �œ��͗���ԍ��Ăяo��
  �� ESC �œ��̓N���A
"
    # �ڑ����J��
    try {
      if($DSN -eq ""){
        $DSNList = Get-OdbcDsn
        if ($DSNList -eq $null){
          "�ڑ��ł���f�[�^�\�[�X������܂���" | Out-Host      
          pause
          break
        } else {
          "�ڑ�����f�[�^�\�[�X��ԍ��œ��͂��Ă�������" | Out-Host
          if ($DSNList.length -eq $null){
            "1:" + $DSNList.Name
            $DSNIndex = Read-Host
            if ($DSNIndex -eq 1){
              $DSN = $DSNList.Name
            } else{
              break
            }
          } else {
            (1 .. $DSNList.length) | % {[string]$_ + ":" + ($DSNList | % Name)[$_ - 1]}
            $DSNIndex = Read-Host
            $DSN = (Get-OdbcDsn | % Name)[$DSNIndex - 1]
          }
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
      break
    }
   
    # ODBC�ڑ�
    $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN;UID=$UID;PWD=$PWD")
    # �ڑ����J��
    try{
      $Con.Open()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      pause
      break
    }

    # �_�C�A���O
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # �t�@�C����ۑ�����_�C�A���O
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
    $SaveFileDialog.Filter = "CSV�t�@�C��(*.CSV)|*.csv|���ׂẴt�@�C��(*.*)|*.*"
    $SaveFileDialog.InitialDirectory = ".\"

    # �t�@�C�����J���_�C�A���O
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
    $OpenFileDialog.Filter = "SQL�t�@�C��(*.SQL)|*.sql|���ׂẴt�@�C��(*.*)|*.*"
    $OpenFileDialog.InitialDirectory = ".\"

    # �e�L�X�g���̓_�C�A���O
    $TextInputDialog = New-Object System.Windows.Forms.Form
    $TextInputDialog.Size = New-Object System.Drawing.Size(800,600) 

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "OK"
    $OKButton.DialogResult = "OK"
    $OKButton.Dock = "Bottom"
    $TextInputDialog.Controls.Add($OKButton)
    $TextInputDialog.AcceptButton = $OKButton

    $TextBox = New-Object System.Windows.Forms.Textbox
    $TextBox.Multiline = $true
    $TextBox.AcceptsReturn = $true
    $TextBox.WordWrap = $true
    $TextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $TextBox.Multiline = $true
    $TextBox.MaxLength = 0
    $TextBox.Font = New-Object Drawing.Font("�l�r �S�V�b�N",10)
    $TextBox.Dock = "Fill"
    $TextBox.Add_KeyDown({
      if ($_.Control -and $_.KeyCode -eq "A"){
        $TextBox.SelectAll()
      }
    })

    $TextInputDialog.Controls.Add($TextBox)
    $TextInputDialog.add_load({$TextBox.Select()})

    # �G���R�[�f�B���O�iSJIS�j
    $OutputEncoding = [console]::OutputEncoding

    # �R���\�[�����[�h
    $Console = $false

    # ���X�g�\�����[�h
    $List = $false

    # SQL�R�}���h
    $Cmd = New-Object System.Data.Odbc.OdbcCommand
    $Cmd.Connection = $Con

    # SQL���s�֐�
    function Execute-SQL{
      param(
        [string]$CommandText,
        [string]$Title
      )
      $Cmd.CommandText = $CommandText

      # SQL���s
      if($Cmd.CommandText -match "select*"){
        $da  = New-Object System.Data.Odbc.OdbcDataAdapter
        $da.SelectCommand = $Cmd
        $DataSet = New-Object System.Data.DataSet
        try {
          $results = $da.Fill($DataSet)
        } catch{
          $_.Exception.InnerException[0].Message
          continue
        }
        # �f�[�^�\��
        try {
          $csv = $DataSet.Tables[0]
          
          if (!$Console){
            $csv | Out-GridView -Title $Title
          } else{
            if (!$List){
              $csv | Format-Table | Out-Host -Paging
            } else {
              $csv | Format-List | Out-Host -Paging
            }
          }
        } catch {
        }
        $results | Out-Host
      } else{
        try {
          # ���s
          $Cmd.ExecuteNonQuery() | Out-Host
        } catch{
          $_.Exception.InnerException[0].Message
          continue
        }
      }
      return $csv
    }

  }
  process{
    while($true){
      # �v�����v�g
      if($transaction.IsolationLevel -eq $null){
        $q = Read-Host [$DSN] | % Trim
      } else {
        $q = Read-Host [$DSN][T] | % Trim
      }
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
      if(($q -eq "tables") -or ($q -eq "tbl")){
        $Schema = Read-Host "�X�L�[�}��" | % Trim
        try {
          $csv = $Con.GetSchema("Tables", ($Con.Database, $Schema)) | Select-Object TABLE_SCHEM, TABLE_NAME
          if (!$Console){
            $csv | Out-GridView -Title ($title + " " + $Schema)
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

      # �J�����ꗗ
      if(($q -eq "columns") -or ($q -eq "col")){
        $Schema = Read-Host "�X�L�[�}��" | % Trim
        $Table = Read-Host "�e�[�u����" | % Trim
        try {
          $csv = ($Con.GetSchema("Columns", ($Con.Database, $Schema, $Table)) | Select-Object TABLE_SCHEM, TABLE_NAME, COLUMN_NAME, TYPE_NAME, COLUMN_SIZE)
          if (!$Console){
            $csv | Out-GridView -Title ($title + " " + $Schema + " " + $Table)
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

      # �r���[�ꗗ
      if($q -eq "views"){
        $Schema = Read-Host "�X�L�[�}��" | % Trim
        try {
          $csv = ($Con.GetSchema("Views", ($Con.Database, $Schema)) | Select-Object TABLE_SCHEM, TABLE_NAME)
          if (!$Console){
            $csv | Out-GridView -Title ($title + " " + $Schema)
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

      # SQL�t�@�C�����s
      if($q -eq "sql"){
        if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
          $sql = ((Get-Content -Path $OpenFileDialog.Filename) -Replace "--.*$","" -Join " " -Split ";") | % Trim | ? Length -ne 0
          $OpenFileDialog.Filename = $OpenFileDialog.Filename | Split-Path -Leaf
          foreach($CommandText in $sql){
            $csv = Execute-SQL -CommandText $CommandText -Title $OpenFileDialog.Filename
          }
        }
        continue
      }

      # �����s�e�L�X�g����
      if($q -eq "text"){
        if($TextInputDialog.ShowDialog() -eq "OK"){
          $sql = ($TextBox.Lines -Replace "--.*$","" -Join " " -Split ";") | % Trim | ? Length -ne 0
          foreach($CommandText in $sql){
            $csv = Execute-SQL -CommandText $CommandText -Title ($CommandText.SubString(0,30) + "...")
          }
        }
        continue
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
      
      # �g�����U�N�V����
      if($transaction.IsolationLevel -eq $null){
        # �g�����U�N�V�����J�n
        if (($q -eq "transaction") -or ($q -eq "trn")){
          $transaction = $Con.BeginTransaction()
          $Cmd.Transaction  = $transaction
          continue
        }
      } else {
        # �R�~�b�g
        if ($q -eq "commit"){
          $transaction.Commit()
          continue
        }
        # ���[���o�b�N
        if (($q -eq "rollback") -or ($q -eq "rol")){
          $transaction.RollBack()
          continue
        }
      }

      # SQL���s
      if (($q -match "select*") -or ($q -match "update*") -or ($q -match "insert*") -or ($q -match "delete*")){
        $csv = Execute-SQL -CommandText $q -Title $q
        continue
      }
      
      # �F�����Ă��Ȃ��R�}���h
      "$q �̓R�}���h�Ƃ��ĔF������܂���B" | Out-Host
      continue
    }

  }
  end{
    # �ڑ������
    try{
      $Con.Close()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      pause
      break
    }
   }
}
