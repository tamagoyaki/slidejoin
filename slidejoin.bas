'
' パワポのファイルを結合するマクロ
'
'
' https://www.cg-method.com/entry/2016/10/25/004021/
'
' 上記URL には、下記２種類のマクロが紹介されている
'   各ファイルのデザインテンプレートが1種類ずつの場合
'   各ファイルのデザインテンプレートがスライドごとにばらばらな場合
' でも、今の俺にはなんのことやらわからんので、１種類ずつの場合をベースにする。
'
'
' 結合するpptp ファイルを記したslideslist を読み込んで、その順番でslide 結合
' するようにした。テキトーだ。slideslist のチェックとかしてないのでちゃんと
' 書いてね。
'
Attribute VB_Name = "Module1"

Sub join()
  '各ファイルのデザインテンプレートが1種類ずつの場合
  Dim newPre As Presentation '新規プレゼンテーション
  Dim myPre As Presentation '既存プレゼンテーション
  Dim i As Long, j As Long
  Dim LstSld As Long, CntSld As Long
  Dim ArrSld() As Long
  Dim fd As FileDialog 'ファイルダイアログ

  ' slides file list
  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
    .AllowMultiSelect = False
    .InitialFileName = "C:" '"E:\Office\PowerPoint\VBAコード"
    .Filters.Add "Slides list", "*.txt", 1
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
  Set newPre = Presentations.Add


  '
  ' join slides
  '
  Do While slideslist.AtEndOfStream <> True
    file = slideslist.ReadLine
    
    Set myPre = Presentations.Open(file, msoTrue, msoFalse, msoFalse)
    With newPre.Slides
      LstSld = .Count '新規プレゼンの最後のスライド番号
      CntSld = myPre.Slides.Count
      '既存プレゼンから新規プレゼンに挿入
      .InsertFromFile myPre.FullName, LstSld, 1, CntSld
      ReDim ArrSld(1 To CntSld)
      For j = 1 To CntSld
        ArrSld(j) = LstSld + j
      Next j
      '既存プレゼンスライド1のデザインをまとめて貼り付け
      .Range(ArrSld).Design = myPre.Slides(1).Design
    End With
    myPre.Close
Loop
slideslist.Close

End Sub
