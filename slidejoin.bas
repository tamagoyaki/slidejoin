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
  '        hoge.pptp show=1,2 hide=3
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
     
  Set regx_file = CreateObject("vbscript.regexp")
  With regx_file
     .Global = True
     .IgnoreCase = True
     .Pattern = "^[-./_a-z0-9]+"
  End With

  Set regx_show = CreateObject("vbscript.regexp")
  With regx_show
     .Global = True
     .Pattern = " show=[0-9,]+"
  End With

  Set regx_hide = CreateObject("vbscript.regexp")
  With regx_hide
     .Global = True
     .Pattern = " hide=[0-9,]+"
  End With

  Set regx_num = CreateObject("vbscript.regexp")
  With regx_num
     .Global = True
     .Pattern = "[0-9]+"
  End With
  
  '
  ' join slides
  '
  Do While slideslist.AtEndOfStream <> True
    line = slideslist.ReadLine

    ' ignore the line if it's a comment or emply
    comment_found = regx_comment.test(line)
    empty_found = regx_empty.test(line)

    If comment_found Or empty_found = True Then
       GoTo CONTINUE
    End If

    ' extracts the options (filename, show, hide ...)
    set file_matches = regx_file.execute(line)
    file = file_matches(0).value

    set show_matches = regx_show.execute(line)
    if show_matches.count then
       set show = regx_num.execute(show_matches(0).value)
    else
       set show = regx_num.execute("")
    end if

    set hide_matches = regx_hide.execute(line)
    if hide_matches.count then
       set hide = regx_num.execute(hide_matches(0).value)
    else
       set hide = regx_num.execute("")
    end if

    ' join the presentations
    Set myPre = Presentations.Open(file, msoTrue, msoFalse, msoFalse)
    With newPre.Slides
      LstSld = .Count ' as the last slide's number
      CntSld = myPre.Slides.Count 

      ' join
      .InsertFromFile myPre.FullName, LstSld, 1, CntSld
      ReDim ArrSld(1 To CntSld)
      For j = 1 To CntSld
        ArrSld(j) = LstSld + j
      Next j

      ' copy the design
      .Range(ArrSld).Design = myPre.Slides(1).Design

      ' apply show option
      for each num in show
	 newPre.Slides(LstSld + num).SlideShowTransition.Hidden = msoFalse
      next

      ' apply hide option
      for each num in hide
	 newPre.Slides(LstSld + num).SlideShowTransition.Hidden = msoTrue
      next
    End With
    myPre.Close
    
CONTINUE:
Loop
slideslist.Close

End Sub
