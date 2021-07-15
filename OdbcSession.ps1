#-------------------------------------------------------------------
# ���ݒ�
#-------------------------------------------------------------------
# �ڑ�������
param(
    [string]$DSN = "", # �f�[�^�\�[�X��
    [string]$UID = "", # ���[�U�[��
    [string]$PWD = ""  # �p�X���[�h
)

# ���[�h
enum Mode{�O���b�h; �e�[�u��; ���X�g; ���[�h��}
$Mode = [Mode]::�O���b�h

# �_�C�A���O
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �t�@�C����ۑ�����_�C�A���O
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
$SaveFileDialog.InitialDirectory = ".\"

# �e�L�X�g���̓_�C�A���O
$TextInputDialog = New-Object System.Windows.Forms.Form
$TextInputDialog.Size = New-Object System.Drawing.Size(800,600) 

# �e�L�X�g�{�b�N�X
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

# OK�{�^��
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Text = "OK"
$OKButton.DialogResult = "OK"
$OKButton.Dock = "Bottom"
$TextInputDialog.Controls.Add($OKButton)
$TextInputDialog.AcceptButton = $OKButton

$TextInputDialog.add_load({
  $TextInputDialog.Activate()
  $TextBox.Select()
})


# �G���R�[�f�B���O�iSJIS�j
$OutputEncoding = [console]::OutputEncoding

# �N�G�����s
function Execute-Query{
  param(
    [string]$CommandText
  )
  
  # ODBC�f�[�^�A�_�v�^
  $DataAdapter = New-Object System.Data.Odbc.OdbcDataAdapter($CommandText, $Con)
  $DataSet = New-Object System.Data.DataSet
  
  try {
    $Records = $DataAdapter.Fill($DataSet)
  } catch{
    $Error.Exception.InnerException[0].Message
    continue
  }
  
  $Result = $DataSet.Tables[0]
  try{
    switch ($Mode){
      �O���b�h {$Result | Out-GridView -Title $InputCommand}
      �e�[�u�� {$Result | Format-Table | Out-Host -Paging}
      ���X�g   {$Result | Format-List  | Out-Host -Paging}
    }
  } catch {
  }
  
  $Records | Out-Null
  return $Result
}

# ��N�G�����s
function Execute-NonQuery{
  param(
    [string]$CommandText
  )
  
  # ODBC�R�}���h
  $Cmd = New-Object System.Data.Odbc.OdbcCommand($CommandText, $Con)
  
  try {
    # ���s
    $Cmd.ExecuteNonQuery()
  } catch{
    $Error.Exception.InnerException[0].Message
    continue
  }
  
}

#-------------------------------------------------------------------
# ��������
#-------------------------------------------------------------------
# �ڑ�
$Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN;UID=$UID;PWD=$PWD")
# �ڑ����J��
try {
  $Con.Open()
} catch{
  if ($Error.Exception.InnerException[0]){
    $Error.Exception.InnerException[0].Message
    Read-Host | Out-Null
    return
  }
}

#-------------------------------------------------------------------
# �又��
#-------------------------------------------------------------------
while($true){
  # �v�����v�g
  $InputCommand = Read-Host "[$DSN]" | % Trim
  
  # �f�[�^�x�[�X��
  if($InputCommand -eq "database" -or $InputCommand -eq "db"){
    $Con.Database
    continue
  }

  # �e�[�u���ꗗ
  if($InputCommand -eq "tables" -or $InputCommand -eq "tbl"){
    $Schema = Read-Host "�X�L�[�}��" | % Trim
    $Result = $Con.GetSchema("Tables", ($Con.Database, $Schema)) | Select-Object TABLE_SCHEM, TABLE_NAME
    $Title = $InputCommand
    try{
      switch ($Mode){
        �O���b�h {$Result | Out-GridView -Title $Title}
        �e�[�u�� {$Result | Format-Table | Out-Host -Paging}
        ���X�g   {$Result | Format-List  | Out-Host -Paging}
      }
    } catch {
    }
    continue
  }

  # CSV�o��
  if($InputCommand -eq "csv"){
    if ($Result -ne $null){
      $SaveFileDialog.Filename = "result.csv"
      $SaveFileDialog.Filter = "CSV�t�@�C��(*.CSV)|*.csv|���ׂẴt�@�C��(*.*)|*.*"
      if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $Result | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
      }
    }
    continue
  }
  
  # �����s�e�L�X�g����
  if($InputCommand -eq "text"){
    if($TextInputDialog.ShowDialog() -eq "OK"){
      $text = ($TextBox.Lines -Replace "--.*$","" -Join " " -Split ";") | % Trim | ? Length -ne 0
      foreach($CommandText in $text){
        $Result = Execute-Query -CommandText $CommandText -Title $InputCommand
      }
    }
    continue
  }
  
  # select��
  if($InputCommand -like "select*"){
    $Result = Execute-Query -CommandText $InputCommand
    continue
  }
  
  # update��
  if($InputCommand -like "update*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }
  
  # insert��
  if($InputCommand -like "insert*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }
  
  # delete��
  if($InputCommand -like "delete*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }

  # ���[�h�ύX
  if($InputCommand -eq "mode"){
    $Mode++
    $Mode%=[Mode]::���[�h��
    $Mode
    continue
  }

  # �N���A�X�N���[��
  if($InputCommand -eq "clear" -or $InputCommand -eq "cls"){
    Clear-Host
    continue
  }

  # �I��
  if ($InputCommand -eq "exit"){
    break
  }

  # ���͂Ȃ�
  if($InputCommand -eq ""){
     continue
  }

  # �F�����Ă��Ȃ��R�}���h
  "$InputCommand �̓R�}���h�Ƃ��ĔF������܂���B"
}
#-------------------------------------------------------------------
# �I������
#-------------------------------------------------------------------
# �ڑ������
try {
  $Con.Close()
} catch{
  if ($Error.Exception.InnerException[0]){
    $Error.Exception.InnerException[0].Message
  }
}
