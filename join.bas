'
' https://www.cg-method.com/entry/2016/10/25/004021/
'
' ��LURL �ɂ́A���L�Q��ނ̃}�N�����Љ��Ă���
'   �e�t�@�C���̃f�U�C���e���v���[�g��1��ނ��̏ꍇ
'   �e�t�@�C���̃f�U�C���e���v���[�g���X���C�h���Ƃɂ΂�΂�ȏꍇ
' �ł��A���̉��ɂ͂Ȃ�̂��Ƃ��킩���̂ŁA�P��ނ��̏ꍇ���x�[�X�ɂ���B
'
Attribute VB_Name = "Module1"
Sub join()
  '�e�t�@�C���̃f�U�C���e���v���[�g��1��ނ��̏ꍇ
  Dim newPre As Presentation '�V�K�v���[���e�[�V����
  Dim myPre As Presentation '�����v���[���e�[�V����
  Dim i As Long, j As Long
  Dim LstSld As Long, CntSld As Long
  Dim ArrSld() As Long
  Dim fd As FileDialog '�t�@�C���_�C�A���O
  '�C�ӂ�*.ppt�t�@�C���Ăяo��
  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
    .InitialFileName = "C:" '"E:\Office\PowerPoint\VBA�R�[�h"
    .Filters.Add "PowerPoint File", "*.ppt;*.pptx;*.pptm;*.pps", 1
    If .Show <> -1 Then Exit Sub
  End With
  '�V�K�v���[���e�[�V����
  Set newPre = Presentations.Add
  For i = 1 To fd.SelectedItems.Count
    Set myPre = Presentations.Open(fd.SelectedItems.Item(i), _
                            msoTrue, , msoFalse)
    With newPre.Slides
      LstSld = .Count '�V�K�v���[���̍Ō�̃X���C�h�ԍ�
      CntSld = myPre.Slides.Count
      '�����v���[������V�K�v���[���ɑ}��
      .InsertFromFile myPre.FullName, LstSld, 1, CntSld
      ReDim ArrSld(1 To CntSld)
      For j = 1 To CntSld
        ArrSld(j) = LstSld + j
      Next j
      '�����v���[���X���C�h1�̃f�U�C�����܂Ƃ߂ē\��t��
      .Range(ArrSld).Design = myPre.Slides(1).Design
    End With
    myPre.Close
  Next i
End Sub
