AddReference "System.Windows.Forms"
AddReference "System.Drawing"
Imports WF = System.Windows.Forms
Imports SD = System.Drawing
Imports IO = System.IO
' ===== [1/2] 여기부터 =====
' 전체조립도의 1단계 하위 조립체(iam)를 선택해 3D DWF(게시옵션: 전체)로 사본 저장
' 저장: 전체조립도 폴더\dwf 변환(yyyy-MM-dd)  (이미 있으면 _2, _3 ...)
' 마지막 줄은 "' ===== [2/2] 끝 =====" 입니다. 붙여넣은 뒤 마지막 줄이 있는지 확인하세요.

Dim topDoc As AssemblyDocument = TryCast(ThisApplication.ActiveDocument, AssemblyDocument)
If topDoc Is Nothing OrElse Not IO.File.Exists(topDoc.FullFileName) Then
    WF.MessageBox.Show("저장된 전체조립도(iam)를 활성화한 상태에서 실행하세요.", "DWF 변환")
    Return
End If

' 1) 1단계 하위 조립체 수집 (억제 제외, 중복 제거)
Dim subDocs As New List(Of AssemblyDocument)
Dim names As New List(Of String)
For Each occ As ComponentOccurrence In topDoc.ComponentDefinition.Occurrences
    If occ.Suppressed Then Continue For
    Dim sd1 As AssemblyDocument = TryCast(occ.Definition.Document, AssemblyDocument)
    If sd1 Is Nothing OrElse sd1.FullFileName = "" Then Continue For
    If names.Contains(sd1.FullFileName.ToLower()) Then Continue For
    names.Add(sd1.FullFileName.ToLower())
    subDocs.Add(sd1)
Next
If subDocs.Count = 0 Then
    WF.MessageBox.Show("1단계 하위 조립체가 없습니다.", "DWF 변환")
    Return
End If

' 2) 체크박스 창 (전체 선택/해제 누르면 창이 다시 열림)
Dim states(subDocs.Count - 1) As Boolean
Dim res As WF.DialogResult = WF.DialogResult.Retry
While res = WF.DialogResult.Retry OrElse res = WF.DialogResult.Ignore
    Dim frm As New WF.Form()
    frm.Text = "DWF 변환할 조립도 선택"
    frm.StartPosition = WF.FormStartPosition.CenterScreen
    frm.ClientSize = New SD.Size(464, 454)
    frm.FormBorderStyle = WF.FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    Dim clb As New WF.CheckedListBox()
    clb.CheckOnClick = True
    clb.Bounds = New SD.Rectangle(12, 12, 440, 390)
    For i As Integer = 0 To subDocs.Count - 1
        clb.Items.Add(IO.Path.GetFileName(subDocs(i).FullFileName), states(i))
    Next
    frm.Controls.Add(clb)
    Dim captions() As String = {"전체 선택", "전체 해제", "변환", "취소"}
    Dim results() As WF.DialogResult = {WF.DialogResult.Retry, WF.DialogResult.Ignore, WF.DialogResult.OK, WF.DialogResult.Cancel}
    Dim xs() As Integer = {12, 108, 280, 372}
    For b As Integer = 0 To 3
        Dim btn As New WF.Button()
        btn.Text = captions(b)
        btn.DialogResult = results(b)
        btn.Bounds = New SD.Rectangle(xs(b), 412, 85, 30)
        frm.Controls.Add(btn)
        If b = 3 Then frm.CancelButton = btn
    Next
    res = frm.ShowDialog()
    For i As Integer = 0 To subDocs.Count - 1
        states(i) = clb.GetItemChecked(i)
        If res = WF.DialogResult.Retry Then states(i) = True
        If res = WF.DialogResult.Ignore Then states(i) = False
    Next
    frm.Dispose()
End While
If res <> WF.DialogResult.OK Then Return
' ===== [1/2] 끝 =====
' ===== [2/2] 여기부터 =====
Dim selected As New List(Of AssemblyDocument)
For i As Integer = 0 To subDocs.Count - 1
    If states(i) Then selected.Add(subDocs(i))
Next
If selected.Count = 0 Then
    WF.MessageBox.Show("선택된 조립도가 없습니다.", "DWF 변환")
    Return
End If

' 3) 출력 폴더
Dim baseDir As String = IO.Path.GetDirectoryName(topDoc.FullFileName)
Dim folderName As String = "dwf 변환(" & DateTime.Now.ToString("yyyy-MM-dd") & ")"
Dim outDir As String = IO.Path.Combine(baseDir, folderName)
Dim n As Integer = 2
While IO.Directory.Exists(outDir)
    outDir = IO.Path.Combine(baseDir, folderName & "_" & n)
    n += 1
End While
IO.Directory.CreateDirectory(outDir)

' 4) DWF 변환기
Dim dwfAddIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")
If Not dwfAddIn.Activated Then dwfAddIn.Activate()

' 5) DWF 사본 저장 (게시옵션: 전체)
Dim okCount As Integer = 0
Dim fails As String = ""
For Each d As AssemblyDocument In selected
    Dim fName As String = IO.Path.GetFileNameWithoutExtension(d.FullFileName)
    Try
        Dim ctx As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext()
        ctx.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim opts As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap()
        Dim data As DataMedium = ThisApplication.TransientObjects.CreateDataMedium()
        If dwfAddIn.HasSaveCopyAsOptions(d, ctx, opts) Then
            opts.Value("Launch_Viewer") = 0
            opts.Value("Publish_Mode") = DWFPublishModeEnum.kCompleteDWFPublish
        End If
        data.FileName = IO.Path.Combine(outDir, fName & ".dwf")
        dwfAddIn.SaveCopyAs(d, ctx, opts, data)
        okCount += 1
    Catch ex As Exception
        fails &= vbCrLf & fName & " : " & ex.Message
    End Try
Next

WF.MessageBox.Show("완료: " & okCount & "개 / 실패: " & (selected.Count - okCount) & "개" & vbCrLf & outDir & vbCrLf & fails, "DWF 변환")
System.Diagnostics.Process.Start("explorer.exe", """" & outDir & """")
' ===== [2/2] 끝 =====
