#-------------------------------------------------------------------
# ���ݒ�
#-------------------------------------------------------------------
param(
    [Parameter(ValueFromPipeline=$true)]
    [string]$DSN = "", # �f�[�^�\�[�X��
    [string]$UID = "", # ���[�U�[��
    [string]$PWD = ""  # �p�X���[�h
)

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
