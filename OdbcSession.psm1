function Enter-OdbcSession{
  [CmdletBinding()]
  param(
    [string]$DSN = "", # �f�[�^�\�[�X��
    [string]$UID = "", # ���[�U�[��
    [string]$PWD = ""  # �p�X���[�h
  )
  begin{
  }
  process{
    while($true){
      # �v�����v�g
      $InputCommand = Read-Host "[$DSN]"
      
      # �I��
      if ($InputCommand -eq "exit"){
        break
      }
      
      # �F�����Ă��Ȃ��R�}���h
      "$InputCommand �̓R�}���h�Ƃ��ĔF������܂���B"
    }
  }
  end{
  }
}
