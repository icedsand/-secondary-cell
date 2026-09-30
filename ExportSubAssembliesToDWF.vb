AddReference "System.Windows.Forms"
AddReference "System.Drawing"
' iLogic 규칙: 전체조립도의 1단계 하위 조립체(iam)를 체크박스로 선택해 3D DWF(게시옵션 전체)로 사본 저장
' 저장 위치: 전체조립도 폴더\dwf 변환(yyyy-MM-dd)  (이미 있으면 _2, _3 ...)
' 사용법: 전체조립도(iam)를 연 상태에서 실행
' Sub Main/Class 없이 일반 코드로 작성 (iLogic이 자동으로 감쌈).
' iLogic이 Inventor 네임스페이스를 자동 Import 하므로 File/Path/Point 등 이름 충돌을 피하려고 전체 이름을 씀.

Dim app As Inventor.Application = ThisApplication
Dim topDoc As Inventor.AssemblyDocument = TryCast(app.ActiveDocument, Inventor.AssemblyDocument)
If topDoc Is Nothing Then
    System.Windows.Forms.MessageBox.Show("전체조립도(iam)를 활성화한 상태에서 실행하세요.", "DWF 변환")
    Return
End If
If topDoc.FullFileName = "" OrElse Not System.IO.File.Exists(topDoc.FullFileName) Then
    System.Windows.Forms.MessageBox.Show("전체조립도가 저장되지 않았습니다. 먼저 저장하세요.", "DWF 변환")
    Return
End If

' 1) 1단계 하위 조립체 수집 (중복 제거, 억제/가상 구성요소 제외)
Dim subDocs As New System.Collections.Generic.List(Of Inventor.AssemblyDocument)
Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
For Each occ As Inventor.ComponentOccurrence In topDoc.ComponentDefinition.Occurrences
    If occ.Suppressed Then Continue For
    If occ.DefinitionDocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then Continue For
    Dim sd As Inventor.AssemblyDocument = TryCast(occ.Definition.Document, Inventor.AssemblyDocument)
    If sd Is Nothing OrElse sd.FullFileName = "" Then Continue For
    If seen.Add(sd.FullFileName) Then subDocs.Add(sd)
Next
If subDocs.Count = 0 Then
    System.Windows.Forms.MessageBox.Show("1단계 하위 조립체가 없습니다.", "DWF 변환")
    Return
End If

' 2) 체크박스 목록창 (전체 선택/해제는 창을 닫았다 다시 여는 방식: 이벤트 핸들러 불필요)
Dim states(subDocs.Count - 1) As Boolean
Dim res As System.Windows.Forms.DialogResult
Do
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "DWF 변환할 조립도 선택"
    frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    frm.Size = New System.Drawing.Size(480, 520)
    frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False

    Dim clb As New System.Windows.Forms.CheckedListBox()
    clb.CheckOnClick = True
    clb.Location = New System.Drawing.Point(12, 12)
    clb.Size = New System.Drawing.Size(440, 390)
    For i As Integer = 0 To subDocs.Count - 1
        clb.Items.Add(System.IO.Path.GetFileName(subDocs(i).FullFileName), states(i))
    Next

    Dim btnAll As New System.Windows.Forms.Button()
    btnAll.Text = "전체 선택"
    btnAll.Location = New System.Drawing.Point(12, 412)
    btnAll.Size = New System.Drawing.Size(90, 30)
    btnAll.DialogResult = System.Windows.Forms.DialogResult.Retry
    Dim btnNone As New System.Windows.Forms.Button()
    btnNone.Text = "전체 해제"
    btnNone.Location = New System.Drawing.Point(108, 412)
    btnNone.Size = New System.Drawing.Size(90, 30)
    btnNone.DialogResult = System.Windows.Forms.DialogResult.Ignore
    Dim btnOk As New System.Windows.Forms.Button()
    btnOk.Text = "변환"
    btnOk.Location = New System.Drawing.Point(280, 412)
    btnOk.Size = New System.Drawing.Size(80, 30)
    btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "취소"
    btnCancel.Location = New System.Drawing.Point(372, 412)
    btnCancel.Size = New System.Drawing.Size(80, 30)
    btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel

    frm.Controls.Add(clb)
    frm.Controls.Add(btnAll)
    frm.Controls.Add(btnNone)
    frm.Controls.Add(btnOk)
    frm.Controls.Add(btnCancel)
    frm.AcceptButton = btnOk
    frm.CancelButton = btnCancel

    res = frm.ShowDialog()
    For i As Integer = 0 To subDocs.Count - 1
        states(i) = clb.GetItemChecked(i)
        If res = System.Windows.Forms.DialogResult.Retry Then states(i) = True
        If res = System.Windows.Forms.DialogResult.Ignore Then states(i) = False
    Next
    frm.Dispose()
Loop While res = System.Windows.Forms.DialogResult.Retry OrElse res = System.Windows.Forms.DialogResult.Ignore

If res <> System.Windows.Forms.DialogResult.OK Then Return
Dim selected As New System.Collections.Generic.List(Of Inventor.AssemblyDocument)
For i As Integer = 0 To subDocs.Count - 1
    If states(i) Then selected.Add(subDocs(i))
Next
If selected.Count = 0 Then
    System.Windows.Forms.MessageBox.Show("선택된 조립도가 없습니다.", "DWF 변환")
    Return
End If

' 3) 출력 폴더: 전체조립도폴더\dwf 변환(yyyy-MM-dd)[_n]
Dim baseDir As String = System.IO.Path.GetDirectoryName(topDoc.FullFileName)
Dim folderName As String = "dwf 변환(" & System.DateTime.Now.ToString("yyyy-MM-dd") & ")"
Dim outDir As String = System.IO.Path.Combine(baseDir, folderName)
Dim n As Integer = 2
While System.IO.Directory.Exists(outDir)
    outDir = System.IO.Path.Combine(baseDir, folderName & "_" & n)
    n += 1
End While
System.IO.Directory.CreateDirectory(outDir)

' 4) DWF 변환기 준비
Dim dwfAddIn As Inventor.TranslatorAddIn = Nothing
For Each ai As Inventor.ApplicationAddIn In app.ApplicationAddIns
    If ai.ClassIdString = "{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}" Then
        dwfAddIn = ai
        Exit For
    End If
Next
If dwfAddIn Is Nothing Then
    System.Windows.Forms.MessageBox.Show("DWF 변환 애드인을 찾을 수 없습니다.", "DWF 변환")
    Return
End If
If Not dwfAddIn.Activated Then dwfAddIn.Activate()

' 5) 선택한 조립도를 DWF로 사본 저장
Dim okCount As Integer = 0
Dim failList As New System.Collections.Generic.List(Of String)
For Each d As Inventor.AssemblyDocument In selected
    Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(d.FullFileName)
    Dim outPath As String = System.IO.Path.Combine(outDir, fileName & ".dwf")
    Dim k As Integer = 2
    While System.IO.File.Exists(outPath)
        outPath = System.IO.Path.Combine(outDir, fileName & "_" & k & ".dwf")
        k += 1
    End While
    Try
        Dim ctx As Inventor.TranslationContext = app.TransientObjects.CreateTranslationContext()
        ctx.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism
        Dim opts As Inventor.NameValueMap = app.TransientObjects.CreateNameValueMap()
        Dim data As Inventor.DataMedium = app.TransientObjects.CreateDataMedium()
        If dwfAddIn.HasSaveCopyAsOptions(d, ctx, opts) Then
            opts.Value("Launch_Viewer") = 0
            ' 게시 옵션: 전체 (급행/전체/사용자 정의 중 전체)
            opts.Value("Publish_Mode") = Inventor.DWFPublishModeEnum.kCompleteDWFPublish
        End If
        data.FileName = outPath
        dwfAddIn.SaveCopyAs(d, ctx, opts, data)
        okCount += 1
    Catch ex As System.Exception
        failList.Add(fileName & " : " & ex.Message)
    End Try
Next

Dim msg As String = "완료: " & okCount & "개 / 실패: " & failList.Count & "개" & vbCrLf & outDir
If failList.Count > 0 Then msg &= vbCrLf & vbCrLf & String.Join(vbCrLf, failList.ToArray())
System.Windows.Forms.MessageBox.Show(msg, "DWF 변환")
System.Diagnostics.Process.Start("explorer.exe", """" & outDir & """")
