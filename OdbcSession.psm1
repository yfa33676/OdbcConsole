function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # �f�[�^�\�[�X��
    [string]$UID = "", # ���[�U�[��
    [string]$PWD = ""  # �p�X���[�h
  )
  begin{
    

    # ODBC�ڑ�
    $Con = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN")
    # �ڑ����J��
    try{
      $Con.Open()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
      break
    }
  }
  process{
    while($true){
      # �v�����v�g
      $InputString = Read-Host "[$DSN]"

      # �I��
      if($InputString -eq "exit"){
        break
      }
    }
  }
  end{
    # �ڑ������
    try{
      $Con.Close()
    } catch{
      $_.Exception.InnerException[0].Message | Out-Host
    }
   }
}
