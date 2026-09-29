' iLogic 규칙: 전체조립도의 1단계 하위 조립체(iam)를 체크박스로 선택해 3D DWF로 사본 저장
' 저장 위치: 전체조립도 폴더\dwf 변환(yyyy-MM-dd)  (이미 있으면 _2, _3 ...)
' 사용법: 전체조립도(iam)를 연 상태에서 실행
AddReference "System.Windows.Forms"
AddReference "System.Drawing"
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim topDoc As Inventor.AssemblyDocument = TryCast(app.ActiveDocument, Inventor.AssemblyDocument)
    If topDoc Is Nothing Then
        MessageBox.Show("전체조립도(iam)를 활성화한 상태에서 실행하세요.", "DWF 변환")
        Return
    End If
    If topDoc.FullFileName = "" OrElse Not File.Exists(topDoc.FullFileName) Then
        MessageBox.Show("전체조립도가 저장되지 않았습니다. 먼저 저장하세요.", "DWF 변환")
        Return
    End If

    ' 1) 1단계 하위 조립체 수집 (중복 제거, 억제/가상 구성요소 제외)
    Dim subDocs As New List(Of Inventor.AssemblyDocument)
    Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each occ As Inventor.ComponentOccurrence In topDoc.ComponentDefinition.Occurrences
        If occ.Suppressed Then Continue For
        If occ.DefinitionDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then Continue For
        Dim d As Inventor.AssemblyDocument = TryCast(occ.Definition.Document, Inventor.AssemblyDocument)
        If d Is Nothing OrElse d.FullFileName = "" Then Continue For
        If seen.Add(d.FullFileName) Then subDocs.Add(d)
    Next
    If subDocs.Count = 0 Then
        MessageBox.Show("1단계 하위 조립체가 없습니다.", "DWF 변환")
        Return
    End If

    ' 2) 체크박스 목록창
    Dim selected As New List(Of Inventor.AssemblyDocument)
    Dim frm As New Form()
    frm.Text = "DWF 변환할 조립도 선택"
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.Size = New Size(480, 520)
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False

    Dim clb As New CheckedListBox()
    clb.CheckOnClick = True
    clb.Location = New Point(12, 12)
    clb.Size = New Size(440, 390)
    For Each d As Inventor.AssemblyDocument In subDocs
        clb.Items.Add(Path.GetFileName(d.FullFileName), False)
    Next

    Dim btnAll As New Button()
    btnAll.Text = "전체 선택"
    btnAll.Location = New Point(12, 412)
    btnAll.Size = New Size(90, 30)
    Dim btnNone As New Button()
    btnNone.Text = "전체 해제"
    btnNone.Location = New Point(108, 412)
    btnNone.Size = New Size(90, 30)
    Dim btnOk As New Button()
    btnOk.Text = "변환"
    btnOk.Location = New Point(280, 412)
    btnOk.Size = New Size(80, 30)
    btnOk.DialogResult = DialogResult.OK
    Dim btnCancel As New Button()
    btnCancel.Text = "취소"
    btnCancel.Location = New Point(372, 412)
    btnCancel.Size = New Size(80, 30)
    btnCancel.DialogResult = DialogResult.Cancel

    AddHandler btnAll.Click, Sub(s, e)
                                 For i As Integer = 0 To clb.Items.Count - 1
                                     clb.SetItemChecked(i, True)
                                 Next
                             End Sub
    AddHandler btnNone.Click, Sub(s, e)
                                  For i As Integer = 0 To clb.Items.Count - 1
                                      clb.SetItemChecked(i, False)
                                  Next
                              End Sub

    frm.Controls.AddRange(New Control() {clb, btnAll, btnNone, btnOk, btnCancel})
    frm.AcceptButton = btnOk
    frm.CancelButton = btnCancel

    If frm.ShowDialog() <> DialogResult.OK Then Return
    For i As Integer = 0 To clb.Items.Count - 1
        If clb.GetItemChecked(i) Then selected.Add(subDocs(i))
    Next
    If selected.Count = 0 Then
        MessageBox.Show("선택된 조립도가 없습니다.", "DWF 변환")
        Return
    End If

    ' 3) 출력 폴더: 전체조립도폴더\dwf 변환(yyyy-MM-dd)[_n]
    Dim baseDir As String = Path.GetDirectoryName(topDoc.FullFileName)
    Dim folderName As String = "dwf 변환(" & DateTime.Now.ToString("yyyy-MM-dd") & ")"
    Dim outDir As String = Path.Combine(baseDir, folderName)
    Dim n As Integer = 2
    While Directory.Exists(outDir)
        outDir = Path.Combine(baseDir, folderName & "_" & n)
        n += 1
    End While
    Directory.CreateDirectory(outDir)

    ' 4) DWF 변환기 준비
    Dim dwfAddIn As Inventor.TranslatorAddIn = Nothing
    For Each ai As Inventor.ApplicationAddIn In app.ApplicationAddIns
        If ai.ClassIdString = "{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}" Then
            dwfAddIn = ai
            Exit For
        End If
    Next
    If dwfAddIn Is Nothing Then
        MessageBox.Show("DWF 변환 애드인을 찾을 수 없습니다.", "DWF 변환")
        Return
    End If
    If Not dwfAddIn.Activated Then dwfAddIn.Activate()

    ' 5) 선택한 조립도를 DWF로 사본 저장
    Dim okList As New List(Of String)
    Dim failList As New List(Of String)
    For Each d As Inventor.AssemblyDocument In selected
        Dim fileName As String = Path.GetFileNameWithoutExtension(d.FullFileName)
        Dim outPath As String = Path.Combine(outDir, fileName & ".dwf")
        Dim k As Integer = 2
        While File.Exists(outPath)
            outPath = Path.Combine(outDir, fileName & "_" & k & ".dwf")
            k += 1
        End While
        Try
            Dim ctx As Inventor.TranslationContext = app.TransientObjects.CreateTranslationContext()
            ctx.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism
            Dim opts As Inventor.NameValueMap = app.TransientObjects.CreateNameValueMap()
            Dim data As Inventor.DataMedium = app.TransientObjects.CreateDataMedium()
            If dwfAddIn.HasSaveCopyAsOptions(d, ctx, opts) Then
                opts.Value("Launch_Viewer") = 0
            End If
            data.FileName = outPath
            dwfAddIn.SaveCopyAs(d, ctx, opts, data)
            okList.Add(Path.GetFileName(outPath))
        Catch ex As Exception
            failList.Add(fileName & " : " & ex.Message)
        End Try
    Next

    Dim msg As String = "완료: " & okList.Count & "개 / 실패: " & failList.Count & "개" & vbCrLf & outDir
    If failList.Count > 0 Then msg &= vbCrLf & vbCrLf & String.Join(vbCrLf, failList)
    MessageBox.Show(msg, "DWF 변환")
    Process.Start("explorer.exe", """" & outDir & """")
End Sub
