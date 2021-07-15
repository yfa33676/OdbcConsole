#-------------------------------------------------------------------
# 環境設定
#-------------------------------------------------------------------
param(
    [Parameter(ValueFromPipeline=$true)]
    [string]$DSN = "", # データソース名
    [string]$UID = "", # ユーザー名
    [string]$PWD = ""  # パスワード
)

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
