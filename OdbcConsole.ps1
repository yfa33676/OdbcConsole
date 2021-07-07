#-------------------------------------------------------------------
# 使い方
#-------------------------------------------------------------------
#
# 0. 環境設定
# データソース名
  $DSN = ""
#
# ユーザー名
  $UID = ""
#
# パスワード
  $PWD = ""
#
# 1. このps1ファイルをPowerShellで実行
#
# 2. [データソース名]: が出力されるのでコマンドを入力
$help = "
      select ...        : SELECT文を実行
      update ...        : UPDATE文を実行
      insert ...        : INSERT文を実行
      delete ...        : DELETE文を実行
      clear または cls  : 画面をクリアする
      csv               : 直前の結果をCSVファイル(SJIS)に出力
      clip              : 直前の結果をクリップボードにコピー(タブ区切り)
      tables            : テーブル一覧を出力
      columns           : カラム一覧を出力
      database          : DB名を出力
      sql               : SQLファイル(SJIS)を開いて実行
      mode              : モード変更(グリッド > コンソール(テーブル) > コンソール(リスト))
      exit または quit  : 終了
      help              : コマンド一覧

  ※ F7 で入力履歴ダイアログ、F8 で入力履歴補完、F9 で入力履歴番号呼び出し
  ※ ESC で入力クリア
"
#-------------------------------------------------------------------
# 初期処理
#-------------------------------------------------------------------

# 接続を開く
try {
  if($DSN -eq ""){
    "接続するデータソースを番号で入力してください"
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
    # パスワードを平文に戻す
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    $PWD = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
  }
} catch{
  exit
}

try {
  # 接続
  $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=" + $DSN + ";UID=" + $UID + ";PWD=" + $PWD)
  # 実行
  $Con.Open()
} catch{
  if ($Error.Exception.InnerException[0]){
    $Error.Exception.InnerException[0].Message
  }
  pause
  exit
}

# ファイル保存ダイアログ
Add-Type -AssemblyName System.Windows.Forms
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
$SaveFileDialog.Filter = "CSVファイル(*.CSV)|*.csv|すべてのファイル(*.*)|*.*"
$SaveFileDialog.InitialDirectory = ".\"

# ファイル開くダイアログ
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$OpenFileDialog.Filter = "SQLファイル(*.SQL)|*.sql|すべてのファイル(*.*)|*.*"
$OpenFileDialog.InitialDirectory = ".\"

# エンコーディング（SJIS）
$OutputEncoding = [console]::OutputEncoding

# コンソールモード
$Console = $false

# リスト表示モード
$List = $false

#-------------------------------------------------------------------
# 主処理
#-------------------------------------------------------------------

while($true){
  # プロンプト
  $q = Read-Host [$DSN] | % Trim
  $title = $q
  
  # 終了
  if(($q -eq "exit") -or ($q -eq "quit")){
    break
  }
  
  # クリアスクリーン
  if(($q -eq "clear") -or ($q -eq "cls")){
    Clear-Host
    continue
  }

  # CSV出力
  if($q -eq "csv"){
    if ($csv -ne $null){
      $SaveFileDialog.Filename = "result.csv"
      if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $csv | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
      }
    }
    continue
  }

  # クリップボードにコピー
  if($q -eq "clip"){
    if ($csv -ne $null){
      $csv | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Set-Clipboard
    }
    continue
  }

  # テーブル一覧
  if($q -eq "tables"){
    $Schema = Read-Host "スキーマ名" | % Trim
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

  # カラム一覧
  if($q -eq "columns"){
    $Schema = Read-Host "スキーマ名" | % Trim
    $Table = Read-Host "テーブル名" | % Trim
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

  # DB名
  if($q -eq "database"){
    $Con.Database
    continue
  }

  # SQLファイル
  if($q -eq "sql"){
    $OpenFileDialog.Filename = ""
    if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
      $title = $OpenFileDialog.Filename
      $q = Get-Content -Path $OpenFileDialog.Filename
    } else{
      continue
    }
  }

  # 入力なし
  if($q -eq ""){
     continue
  }
  
  # モードチェンジ
  if ($q -eq "mode"){
     if (!$Console -and !$List) {
       $Console = $true
       $List = $false
       "コンソール(テーブル)" | Out-Host
     } elseif ($Console -and !$List) {
       $Console = $true
       $List = $true
       "コンソール(リスト)" | Out-Host
     } elseif ($Console -and $List) {
       $Console = $false
       $List = $false
       "グリッド" | Out-Host
     }
     continue
  }

  # ヘルプ
  if($q -eq "help"){
     $help
     continue
  }

  # SQLコマンド
  $Cmd = New-Object System.Data.Odbc.OdbcCommand
  $Cmd.CommandText = $q
  $Cmd.Connection = $Con

  # SQL実行
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
    # データ表示
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
      # 実行
      $Cmd.ExecuteNonQuery()
    } catch{
      $Error.Exception.InnerException[0].Message
      continue
    }
  }
}

#-------------------------------------------------------------------
# 終了処理
#-------------------------------------------------------------------

# 接続を閉じる
$Con.Close()
