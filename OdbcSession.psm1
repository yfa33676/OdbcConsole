function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # データソース名
    [string]$UID = "", # ユーザー名
    [string]$PWD = ""  # パスワード
  )
  begin{
  }
  process{
    while($true){
      # プロンプト
      $InputCommand = Read-Host "[$DSN]"
      
      # 終了
      if ($InputCommand -eq "exit"){
        break
      }
      
      # 認識していないコマンド
      "$InputCommand はコマンドとして認識されません。"
    }
  }
  end{
  }
}
