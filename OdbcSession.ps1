#-------------------------------------------------------------------
# 環境設定
#-------------------------------------------------------------------
# 接続文字列
param(
    [string]$DSN = "", # データソース名
    [string]$UID = "", # ユーザー名
    [string]$PWD = ""  # パスワード
)

# モード
enum Mode{グリッド; テーブル; リスト; モード数}
$Mode = [Mode]::グリッド

# ダイアログ
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ファイルを保存するダイアログ
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
$SaveFileDialog.InitialDirectory = ".\"

# テキスト入力ダイアログ
$TextInputDialog = New-Object System.Windows.Forms.Form
$TextInputDialog.Size = New-Object System.Drawing.Size(800,600) 

# テキストボックス
$TextBox = New-Object System.Windows.Forms.Textbox
$TextBox.Multiline = $true
$TextBox.AcceptsReturn = $true
$TextBox.WordWrap = $true
$TextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$TextBox.Multiline = $true
$TextBox.MaxLength = 0
$TextBox.Font = New-Object Drawing.Font("ＭＳ ゴシック",10)
$TextBox.Dock = "Fill"
$TextBox.Add_KeyDown({
  if ($_.Control -and $_.KeyCode -eq "A"){
    $TextBox.SelectAll()
  }
})
$TextInputDialog.Controls.Add($TextBox)

# OKボタン
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


# エンコーディング（SJIS）
$OutputEncoding = [console]::OutputEncoding

# クエリ実行
function Execute-Query{
  param(
    [string]$CommandText
  )
  
  # ODBCデータアダプタ
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
      グリッド {$Result | Out-GridView -Title $InputCommand}
      テーブル {$Result | Format-Table | Out-Host -Paging}
      リスト   {$Result | Format-List  | Out-Host -Paging}
    }
  } catch {
  }
  
  $Records | Out-Null
  return $Result
}

# 非クエリ実行
function Execute-NonQuery{
  param(
    [string]$CommandText
  )
  
  # ODBCコマンド
  $Cmd = New-Object System.Data.Odbc.OdbcCommand($CommandText, $Con)
  
  try {
    # 実行
    $Cmd.ExecuteNonQuery()
  } catch{
    $Error.Exception.InnerException[0].Message
    continue
  }
  
}

#-------------------------------------------------------------------
# 初期処理
#-------------------------------------------------------------------
# 接続
$Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN;UID=$UID;PWD=$PWD")
# 接続を開く
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
# 主処理
#-------------------------------------------------------------------
while($true){
  # プロンプト
  $InputCommand = Read-Host "[$DSN]" | % Trim
  
  # データベース名
  if($InputCommand -eq "database" -or $InputCommand -eq "db"){
    $Con.Database
    continue
  }

  # テーブル一覧
  if($InputCommand -eq "tables" -or $InputCommand -eq "tbl"){
    $Schema = Read-Host "スキーマ名" | % Trim
    $Result = $Con.GetSchema("Tables", ($Con.Database, $Schema)) | Select-Object TABLE_SCHEM, TABLE_NAME
    $Title = $InputCommand
    try{
      switch ($Mode){
        グリッド {$Result | Out-GridView -Title $Title}
        テーブル {$Result | Format-Table | Out-Host -Paging}
        リスト   {$Result | Format-List  | Out-Host -Paging}
      }
    } catch {
    }
    continue
  }

  # CSV出力
  if($InputCommand -eq "csv"){
    if ($Result -ne $null){
      $SaveFileDialog.Filename = "result.csv"
      $SaveFileDialog.Filter = "CSVファイル(*.CSV)|*.csv|すべてのファイル(*.*)|*.*"
      if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $Result | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
      }
    }
    continue
  }
  
  # 複数行テキスト入力
  if($InputCommand -eq "text"){
    if($TextInputDialog.ShowDialog() -eq "OK"){
      $text = ($TextBox.Lines -Replace "--.*$","" -Join " " -Split ";") | % Trim | ? Length -ne 0
      foreach($CommandText in $text){
        $Result = Execute-Query -CommandText $CommandText -Title $InputCommand
      }
    }
    continue
  }
  
  # select文
  if($InputCommand -like "select*"){
    $Result = Execute-Query -CommandText $InputCommand
    continue
  }
  
  # update文
  if($InputCommand -like "update*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }
  
  # insert文
  if($InputCommand -like "insert*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }
  
  # delete文
  if($InputCommand -like "delete*"){
    Execute-NonQuery -CommandText $InputCommand
    continue
  }

  # モード変更
  if($InputCommand -eq "mode"){
    $Mode++
    $Mode%=[Mode]::モード数
    $Mode
    continue
  }

  # クリアスクリーン
  if($InputCommand -eq "clear" -or $InputCommand -eq "cls"){
    Clear-Host
    continue
  }

  # 終了
  if ($InputCommand -eq "exit"){
    break
  }

  # 入力なし
  if($InputCommand -eq ""){
     continue
  }

  # 認識していないコマンド
  "$InputCommand はコマンドとして認識されません。"
}
#-------------------------------------------------------------------
# 終了処理
#-------------------------------------------------------------------
# 接続を閉じる
try {
  $Con.Close()
} catch{
  if ($Error.Exception.InnerException[0]){
    $Error.Exception.InnerException[0].Message
  }
}
