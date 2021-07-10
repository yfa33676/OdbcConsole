function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # データソース名
    [string]$UID = "", # ユーザー名
    [string]$PWD = ""  # パスワード
  )
  begin{
    

    # ODBC接続
    $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN")
    # 接続を開く
    try{
      $Con.Open()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      break
    }
  }
  process{
    while($true){
      # プロンプト
      $InputString = Read-Host "[$DSN]"

      # 終了
      if($InputString -eq "exit"){
        break
      }
    }
  }
  end{
    # 接続を閉じる
    try{
      $Con.Close()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
    }
   }
}
