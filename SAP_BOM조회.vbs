' ============================================================
' SAP BOM 조회 자동화 스크립트
' 완제품 코드(9코드)를 입력하면 SAP ZSDR9030 실행 후
' 결과를 바탕화면에 TXT로 자동 저장합니다.
' ============================================================

' HTA 입력 폼 실행
Dim fso, wshShell, htaPath, tempPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
htaPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\SAP_BOM입력.hta"
tempPath = wshShell.ExpandEnvironmentStrings("%TEMP%") & "\bom_input.txt"

' 이전 임시파일 삭제
If fso.FileExists(tempPath) Then fso.DeleteFile tempPath

' HTA가 있는지 확인
If Not fso.FileExists(htaPath) Then
    MsgBox "SAP_BOM입력.hta 파일이 없습니다." & vbCrLf & "SAP_BOM조회.vbs와 같은 폴더에 넣어주세요.", vbCritical, "오류"
    WScript.Quit
End If

' HTA를 TEMP 폴더에 복사 후 실행
Dim htaTempPath
htaTempPath = wshShell.ExpandEnvironmentStrings("%TEMP%") & "\SAP_BOM입력.hta"
fso.CopyFile htaPath, htaTempPath, True

' 완료 플래그 파일
Dim flagPath
flagPath = wshShell.ExpandEnvironmentStrings("%TEMP%") & "\bom_done.flag"
If fso.FileExists(flagPath) Then fso.DeleteFile flagPath

wshShell.Run "mshta.exe """ & htaTempPath & """", 1, False

' HTA가 닫힐 때까지 대기 (플래그 파일 감지)
Dim waitCount
waitCount = 0
Do While Not fso.FileExists(flagPath) And Not fso.FileExists(tempPath)
    WScript.Sleep 500
    waitCount = waitCount + 1
    If waitCount > 7200 Then
        MsgBox "입력 대기 시간이 초과되었습니다.", vbExclamation, "시간 초과"
        WScript.Quit
    End If
Loop
WScript.Sleep 300

' 플래그 파일 정리
If fso.FileExists(flagPath) Then fso.DeleteFile flagPath

' 결과 파일 읽기
If Not fso.FileExists(tempPath) Then
    MsgBox "입력이 취소되었습니다.", vbInformation, "SAP BOM 조회"
    WScript.Quit
End If

Dim codes()
ReDim codes(0)
Dim orderQtys()
ReDim orderQtys(0)
Dim prodNames()
ReDim prodNames(0)
Dim codeCount
codeCount = 0
Dim i

Dim ts, line, parts, c
Set ts = fso.OpenTextFile(tempPath, 1)
Do While Not ts.AtEndOfStream
    line = ts.ReadLine
    parts = Split(line, "|")
    If UBound(parts) >= 0 Then
        c = Trim(parts(0))
        If c <> "" Then
            If codeCount > UBound(codes) Then
                ReDim Preserve codes(codeCount)
                ReDim Preserve orderQtys(codeCount)
                ReDim Preserve prodNames(codeCount)
            End If
            codes(codeCount) = c
            If UBound(parts) >= 1 Then
                orderQtys(codeCount) = Trim(parts(1))
            Else
                orderQtys(codeCount) = ""
            End If
            If UBound(parts) >= 2 Then
                prodNames(codeCount) = Trim(parts(2))
            Else
                prodNames(codeCount) = ""
            End If
            codeCount = codeCount + 1
        End If
    End If
Loop
ts.Close
Set ts = Nothing
fso.DeleteFile tempPath

If codeCount = 0 Then
    MsgBox "입력된 코드가 없습니다.", vbExclamation, "오류"
    WScript.Quit
End If

' 수량 콤마 포맷 함수
Function FormatQty(val)
    If val = "" Then
        FormatQty = ""
        Exit Function
    End If
    Dim num
    num = Replace(val, ",", "")
    If IsNumeric(num) Then
        FormatQty = FormatNumber(CDbl(num), 0, -1, 0, -1)
    Else
        FormatQty = val
    End If
End Function

' 입력 확인
Dim confirmMsg
confirmMsg = "아래 내용으로 BOM 조회를 실행합니다." & vbCrLf & vbCrLf
For i = 0 To codeCount - 1
    confirmMsg = confirmMsg & "  " & (i + 1) & ". " & codes(i)
    If prodNames(i) <> "" Then
        confirmMsg = confirmMsg & "  " & prodNames(i)
    End If
    If orderQtys(i) <> "" Then
        confirmMsg = confirmMsg & "  (발주: " & FormatQty(orderQtys(i)) & ")"
    End If
    confirmMsg = confirmMsg & vbCrLf
Next
confirmMsg = confirmMsg & vbCrLf & "총 " & codeCount & "개 / 실행하시겠습니까?"

If MsgBox(confirmMsg, vbYesNo + vbQuestion, "SAP BOM 조회 확인") = vbNo Then
    WScript.Quit
End If

' SAP GUI 연결
Dim SapGuiAuto, application, connection, session
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    MsgBox "SAP GUI가 실행되어 있지 않습니다." & vbCrLf & "SAP에 먼저 로그인해 주세요.", vbCritical, "오류"
    WScript.Quit
End If
Set application = SapGuiAuto.GetScriptingEngine
On Error GoTo 0

' ERP connection 찾기 (APO·BW 등 다른 시스템 동시 로그인 대응)
' SAP Logon에서 로그인 순서에 따라 Children(0)이 달라지므로 단계적으로 매칭
Dim conn, descUp
Set connection = Nothing

' 1차: Description에 "ERP" + "PRD" 둘 다 포함 (운영 시스템 우선)
For Each conn In application.Children
    descUp = UCase(conn.Description)
    If InStr(descUp, "ERP") > 0 And InStr(descUp, "PRD") > 0 Then
        Set connection = conn
        Exit For
    End If
Next

' 2차: PRD가 들어간 connection (이름에 ERP 키워드가 없는 환경 대응)
If connection Is Nothing Then
    For Each conn In application.Children
        descUp = UCase(conn.Description)
        If InStr(descUp, "PRD") > 0 _
            And InStr(descUp, "APO") = 0 And InStr(descUp, "BW") = 0 _
            And InStr(descUp, "SCM") = 0 And InStr(descUp, "CRM") = 0 Then
            Set connection = conn
            Exit For
        End If
    Next
End If

' 3차: Description에 "ERP" 포함 (DEV/QAS 등도 허용)
If connection Is Nothing Then
    For Each conn In application.Children
        descUp = UCase(conn.Description)
        If InStr(descUp, "ERP") > 0 Then
            Set connection = conn
            Exit For
        End If
    Next
End If

' 4차: APO·BW·SCM·CRM이 아닌 connection
If connection Is Nothing Then
    For Each conn In application.Children
        descUp = UCase(conn.Description)
        If InStr(descUp, "APO") = 0 And InStr(descUp, "BW") = 0 _
            And InStr(descUp, "SCM") = 0 And InStr(descUp, "CRM") = 0 Then
            Set connection = conn
            Exit For
        End If
    Next
End If

' 5차: connection이 1개뿐이면 그것 사용 (단일 시스템 환경)
If connection Is Nothing And application.Children.Count = 1 Then
    Set connection = application.Children(0)
End If

If connection Is Nothing Then
    Dim debugList
    debugList = ""
    For Each conn In application.Children
        debugList = debugList & "  - " & conn.Description & vbCrLf
    Next
    MsgBox "SAP ERP를 찾지 못했습니다." & vbCrLf & vbCrLf & _
        "현재 로그인된 SAP 시스템 목록:" & vbCrLf & debugList & vbCrLf & _
        "ERP 시스템에 로그인 후 다시 실행해 주세요." & vbCrLf & _
        "(이 메시지를 캡처해서 관리자에게 보내면 매칭 규칙을 보강할 수 있습니다.)", _
        vbCritical, "ERP 미연결"
    WScript.Quit
End If
Set session = connection.Children(0)

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

' SAP 트랜잭션 실행
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSDR9030"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1000

' 변형 선택
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 500
session.findById("wnd[1]/usr/lbl[10,5]").setFocus
session.findById("wnd[1]/usr/lbl[10,5]").caretPosition = 6
session.findById("wnd[1]").sendVKey 2
WScript.Sleep 500

' 플랜트 선택
session.findById("wnd[0]/usr/ctxtP_WERKS").setFocus
session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 3
session.findById("wnd[0]").sendVKey 4
WScript.Sleep 500
session.findById("wnd[1]/usr/lbl[22,5]").setFocus
session.findById("wnd[1]/usr/lbl[22,5]").caretPosition = 12
session.findById("wnd[1]").sendVKey 2
WScript.Sleep 500

' 자재코드 다중 입력 창 열기
session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
WScript.Sleep 500

' 코드 입력 (다중입력 테이블 - 화면에 보이는 행 수만큼만 한 번에 입력 가능 → 페이지 스크롤)
Dim tblId, tbl, visibleRows, pageStart, vIdx
tblId = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
Set tbl = session.findById(tblId)
visibleRows = tbl.VisibleRowCount
If visibleRows < 1 Then visibleRows = 12

For i = 0 To codeCount - 1
    pageStart = (i \ visibleRows) * visibleRows
    vIdx = i Mod visibleRows
    If vIdx = 0 And pageStart > 0 Then
        ' 다음 페이지로 스크롤
        Set tbl = session.findById(tblId)
        tbl.verticalScrollbar.position = pageStart
        WScript.Sleep 150
    End If
    session.findById(tblId & "/ctxtRSCSEL_255-SLOW_I[1," & vIdx & "]").text = codes(i)
Next

' 실행
session.findById("wnd[1]/tbar[0]/btn[8]").press
WScript.Sleep 500
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2000

' 결과 ALV 그리드에서 전체 행 선택
Dim grid
Set grid = session.findById("wnd[0]/shellcont[1]/shell")
Dim rowCount
rowCount = grid.RowCount

If rowCount = 0 Then
    MsgBox "조회 결과가 없습니다.", vbExclamation, "결과 없음"
    WScript.Quit
End If

grid.selectedRows = "0-" & (rowCount - 1)

' 엑스포트
grid.pressToolbarContextButton "&MB_EXPORT"
WScript.Sleep 300
grid.selectContextMenuItem "&PC"
WScript.Sleep 1000

On Error Resume Next
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
WScript.Sleep 200
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 500
On Error GoTo 0

' 저장 경로 지정
Dim savePath, desktopPath, fileName
desktopPath = wshShell.SpecialFolders("Desktop")
fileName = "BOM조회_" & Replace(Replace(Replace(FormatDateTime(Now, 0), "/", ""), " ", "_"), ":", "") & ".txt"
savePath = desktopPath & "\" & fileName

On Error Resume Next
session.findById("wnd[1]/usr/ctxtDY_PATH").text = desktopPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 1000
On Error GoTo 0

' 발주 수량 정보를 TXT 파일 끝에 추가
WScript.Sleep 500
If fso.FileExists(savePath) Then
    Dim hasQty
    hasQty = False
    For i = 0 To codeCount - 1
        If orderQtys(i) <> "" Then hasQty = True
    Next
    If hasQty Then
        Set ts = fso.OpenTextFile(savePath, 8, False)
        ts.WriteLine ""
        ts.WriteLine "##ORDER_QTY##"
        For i = 0 To codeCount - 1
            If orderQtys(i) <> "" Then
                ts.WriteLine codes(i) & "|" & Replace(orderQtys(i), ",", "") & "|" & prodNames(i)
            End If
        Next
        ts.Close
        Set ts = Nothing
    End If
End If

' 완료 메시지
Dim inputList
inputList = ""
For i = 0 To codeCount - 1
    inputList = inputList & codes(i)
    If i < codeCount - 1 Then inputList = inputList & ", "
Next

If fso.FileExists(savePath) Then
    MsgBox "BOM 조회 완료!" & vbCrLf & vbCrLf & _
        "저장 위치: " & savePath & vbCrLf & _
        "조회 건수: " & rowCount & "건" & vbCrLf & _
        "입력 코드: " & inputList & vbCrLf & vbCrLf & _
        "이 파일을 웹 사이트에 업로드하세요.", _
        vbInformation, "SAP BOM 조회 완료"
Else
    MsgBox "BOM 조회가 실행되었습니다." & vbCrLf & vbCrLf & _
        "조회 건수: " & rowCount & "건" & vbCrLf & _
        "파일 저장 다이얼로그가 나타나면 바탕화면에 저장해 주세요.", _
        vbInformation, "SAP BOM 조회 완료"
End If

Set grid = Nothing
Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set SapGuiAuto = Nothing
Set fso = Nothing
Set wshShell = Nothing
