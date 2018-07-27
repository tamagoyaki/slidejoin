Attribute VB_Name = "Module1"
'
' �p���|�̃t�@�C������������}�N��
'
'
' https://www.cg-method.com/entry/2016/10/25/004021/
'
' ��LURL �ɂ́A���L�Q��ނ̃}�N�����Љ��Ă���
'   �e�t�@�C���̃f�U�C���e���v���[�g��1��ނ��̏ꍇ
'   �e�t�@�C���̃f�U�C���e���v���[�g���X���C�h���Ƃɂ΂�΂�ȏꍇ
' �ł��A���̉��ɂ͂Ȃ�̂��Ƃ��킩���̂ŁA�P��ނ��̏ꍇ���x�[�X�ɂ���B
'
'
' ��������pptp �t�@�C�����L����slideslist ��ǂݍ���ŁA���̏��Ԃ�slide ����
' ����悤�ɂ����B�e�L�g�[���Bslideslist �̃`�F�b�N�Ƃ����ĂȂ��̂ł�����
' �����ĂˁB
'

Sub slidejoin()
  '�e�t�@�C���̃f�U�C���e���v���[�g��1��ނ��̏ꍇ
  Dim newPre As Presentation '�V�K�v���[���e�[�V����
  Dim myPre As Presentation '�����v���[���e�[�V����
  Dim i As Long, j As Long
  Dim LstSld As Long, CntSld As Long
  Dim ArrSld() As Long
  Dim fd As FileDialog '�t�@�C���_�C�A���O

  ' slides file list
  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
    .AllowMultiSelect = False
    .InitialFileName = "C:" '"E:\Office\PowerPoint\VBA�R�[�h"
    .Filters.Add "Slide list", "*.slidelist", 1
    If .Show <> -1 Then Exit Sub
  End With

  ' read slides file
  '
  '   slides file is a text file which describes pptp file with path.
  '
  '        c:\a.pptp
  '        b.pptp
  '        c.pptp
  '
  '
  Set slideslist = CreateObject("Scripting.FileSystemObject").OpenTextFile(fd.SelectedItems.Item(1), 1)

  ' Set screen size to 4:3
  Set newPre = Presentations.Add
  With newPre
     newPre.PageSetup.SlideSize = ppSlideSizeOnScreen
  End With

  Set regx_comment = CreateObject("vbscript.regexp")
  With regx_comment
     .Global = True
     .Pattern = "^#.*$"
  End With
  
  Set regx_empty = CreateObject("vbscript.regexp")
  With regx_empty
     .Global = True
     .Pattern = "^ *$"
  End With
     
  
  '
  ' join slides
  '
  Do While slideslist.AtEndOfStream <> True
    file = slideslist.ReadLine

    ' ignore the line if it's a comment or emply
    comment_found = regx_comment.test(file)
    empty_found = regx_empty.test(file)

    If comment_found Or empty_found = True Then
       GoTo CONTINUE
    End If
    
    Set myPre = Presentations.Open(file, msoTrue, msoFalse, msoFalse)
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
    
CONTINUE:
Loop
slideslist.Close

End Sub
