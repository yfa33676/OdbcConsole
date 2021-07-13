function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # データソース名
    [string]$UID = "", # ユーザー名
    [string]$PWD = ""  # パスワード
  )

  begin{
    $help = "
      select ...             : SELECT文を実行
      update ...             : UPDATE文を実行
      insert ...             : INSERT文を実行
      delete ...             : DELETE文を実行
      clear または cls       : 画面をクリアする
      csv                    : 直前の結果をCSVファイル(SJIS)に出力
      clip                   : 直前の結果をクリップボードにコピー(タブ区切り)
      tables または tbl      : テーブル一覧を出力
      columns または col     : カラム一覧を出力
      database               : DB名を出力
      sql                    : SQLファイル(SJIS)を開いて実行
      mode                   : モード変更(グリッド > コンソール(テーブル) > コンソール(リスト))
      transaction または trn : トランザクションの開始 
      commit                 : コミット
      rollback または rol    : ロールバック
      exit または quit       : 終了
      help                   : コマンド一覧

  ※ F7 で入力履歴ダイアログ、F8 で入力履歴補完、F9 で入力履歴番号呼び出し
  ※ ESC で入力クリア
"
    # 接続を開く
    try {
      if($DSN -eq ""){
        $DSNList = Get-OdbcDsn
        if ($DSNList -eq $null){
          "接続できるデータソースがありません" | Out-Host      
          pause
          break
        } else {
          "接続するデータソースを番号で入力してください" | Out-Host
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
        # パスワードを平文に戻す
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
        $PWD = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
      }
    } catch{
      break
    }
   
    # ODBC接続
    $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN;UID=$UID;PWD=$PWD")
    # 接続を開く
    try{
      $Con.Open()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      pause
      break
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

    # SQLコマンド
    $Cmd = New-Object System.Data.Odbc.OdbcCommand
    $Cmd.Connection = $Con

    # SQL実行関数
    function Execute-SQL{
      param(
        [string]$CommandText,
        [string]$Title
      )
      $Cmd.CommandText = $CommandText

      # SQL実行
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
        # データ表示
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
          # 実行
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
      # プロンプト
      if($transaction.IsolationLevel -eq $null){
        $q = Read-Host [$DSN] | % Trim
      } else {
        $q = Read-Host [$DSN][T] | % Trim
      }
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
      if(($q -eq "tables") -or ($q -eq "tbl")){
        $Schema = Read-Host "スキーマ名" | % Trim
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

      # カラム一覧
      if(($q -eq "columns") -or ($q -eq "col")){
        $Schema = Read-Host "スキーマ名" | % Trim
        $Table = Read-Host "テーブル名" | % Trim
        try {
          $csv = ($Con.GetSchema("Columns", ($Con.Database, $Schema, $Table)) | Select-Object TABLE_SCHEM, TABLE_NAME, COLUMN_NAME, TYPE_NAME)
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

      # DB名
      if($q -eq "database"){
        $Con.Database
        continue
      }

      # SQLファイル実行
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
      
      # トランザクション
      if($transaction.IsolationLevel -eq $null){
        # トランザクション開始
        if (($q -eq "transaction") -or ($q -eq "trn")){
          $transaction = $Con.BeginTransaction()
          $Cmd.Transaction  = $transaction
          continue
        }
      } else {
        # コミット
        if ($q -eq "commit"){
          $transaction.Commit()
          continue
        }
        # ロールバック
        if (($q -eq "rollback") -or ($q -eq "rol")){
          $transaction.RollBack()
          continue
        }
      }

      # SQL実行
      if (($q -match "select*") -or ($q -match "update*") -or ($q -match "insert*") -or ($q -match "delete*")){
        $csv = Execute-SQL -CommandText $q -Title $q
        continue
      }
      
      # 認識していないコマンド
      "$q はコマンドとして認識されません。" | Out-Host
      continue
    }

  }
  end{
    # 接続を閉じる
    try{
      $Con.Close()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      pause
      break
    }
   }
}
