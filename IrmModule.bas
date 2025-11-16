Attribute VB_Name = "mdlGeneral"
Public GlobalCustomer, GlobalSchedule, GlobalDeposit, GlobalToken As Long
Public cnDatabase As New ADODB.Connection
Public conStr, mnuTransValue, usr As String
Public Function NUMERIC(ByVal value As Integer) As Integer
    Dim LENG
    If value <> 8 Then
        LENG = Chr(value)
        LENG = InStr(1, "1234567890", LENG, vbTextCompare)
        NUMERIC = LENG
    Else
         NUMERIC = value
    End If
End Function
Public Function DBConn()
    Dim str As String
    If cnDatabase.State = adStateOpen Then
        cnDatabase.Close
    End If
    'conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\Dyeing.mdb;Persist Security Info=False"
    conStr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=dyeing;Initial Catalog=Dyeing;Data Source=."
    cnDatabase.Open conStr
End Function
Public Function ValAutoNumber(tblName As String, colName As String) As String
    Dim RsMax As New ADODB.Recordset
    Dim sql As String
    Dim rtnValue As String
    rtnValue = "0"
    If RsMax.State = adStateOpen Then
        RsMax.Close
    End If
    sql = "Select max(" & colName & ") as maxVal from " & tblName
    RsMax.Open sql, cnDatabase, adOpenForwardOnly, adLockReadOnly
    If Not RsMax.EOF Then
        rtnValue = IIf(IsNull(RsMax.Fields(0)), "1", RsMax.Fields(0))
    End If
    RsMax.Close
    Set RsMax = Nothing
    ValAutoNumber = CStr(CDbl(rtnValue) + 1)
End Function
Public Function GetID(colName As String) As String
    Dim RsMax As New ADODB.Recordset
    Dim sql As String
    Dim rtnValue As String
    rtnValue = "0"
    If RsMax.State = adStateOpen Then
        RsMax.Close
    End If
    sql = "Select " & colName & ".nextVal as maxVal from Dual"
    RsMax.Open sql, cnDatabase, adOpenForwardOnly, adLockReadOnly
    If Not RsMax.EOF Then
        rtnValue = IIf(IsNull(RsMax.Fields(0)), "1", RsMax.Fields(0))
    End If
    RsMax.Close
    Set RsMax = Nothing
    GetID = rtnValue
End Function
Public Function MonthDiff(StDt As Date, EnDt As Date) As Long
    Dim fyr, fmn, syr, smn, rstyr, diffMonth, diffYear, RST As Long
    fyr = Year(StDt)
    fmn = Month(StDt)
    syr = Year(EnDt)
    smn = Month(EnDt)
    rstyr = syr - fyr
    If rstyr > 0 Then
        mnSt = (fmn - 12) + 1
        mnEnd = smn
        diffMonth = mnSt + mnEnd
    diffYear = ((syr - fyr) - 1) * 12
    RST = diffYear + diffMonth
Else
    RST = (smn - fmn) + 1
End If
    MonthDiff = RST
End Function
Public Function FillRecordSet(strQry As String) As ADODB.Recordset
    Dim RST As New ADODB.Recordset
    
    If RST.State = adStateOpen Then
        RST.Close
    End If
    RST.CursorLocation = adUseClient
    Debug.Print strQry
    RST.Open strQry, cnDatabase, adOpenDynamic, adLockOptimistic
    
    Set FillRecordSet = RST
    Set RST = Nothing
End Function
Public Sub FillColorCombo(dbQuery As String, targetCombo As ComboBox, comboTextField As String, comboIndexField As String)
    Dim targetRST As New ADODB.Recordset
    Set targetRST = FillRecordSet(dbQuery)
    
    targetCombo.Clear
    targetCombo.AddItem "-- Select --"
    targetCombo.ItemData(0) = 0
    While Not targetRST.EOF
        targetCombo.AddItem targetRST(comboTextField)
        targetCombo.ItemData(targetCombo.NewIndex) = targetRST(comboIndexField)
        targetRST.MoveNext
    Wend
    
    targetRST.Close
    Set targetRST = Nothing
End Sub
Public Sub FillCombo(dbQuery As String, targetCombo As ComboBox, comboTextField As String, comboIndexField As String)
    Dim targetRST As New ADODB.Recordset
    Set targetRST = FillRecordSet(dbQuery)
    
    targetCombo.Clear
    
    While Not targetRST.EOF
        targetCombo.AddItem targetRST(comboTextField)
        targetCombo.ItemData(targetCombo.NewIndex) = targetRST(comboIndexField)
        targetRST.MoveNext
    Wend
    
    targetRST.Close
    Set targetRST = Nothing
End Sub
Public Sub FillCombo2(NameofValues As String, targetCombo As ComboBox)
    If LCase(NameofValues) = "monthnames" Then
        For month_no = 1 To 12
            targetCombo.AddItem MonthName(month_no, False)
            targetCombo.ItemData(targetCombo.NewIndex) = month_no
        Next
    End If
    
    If LCase(NameofValues) = "daynames" Then
        For day_no = 1 To 7
            targetCombo.AddItem WeekdayName(day_no, , vbSaturday)
            targetCombo.ItemData(targetCombo.NewIndex) = day_no
        Next
    End If
End Sub
Public Function getItemTypeName(ItemTypeCode As String) As String
Dim str As String
    Set rstGetVal = FillRecordSet("select ItemTypeName from ItemType where ItemTypeCode =" & ItemTypeCode)
        If Not (rstGetVal.EOF) Then
            str = rstGetVal("ItemTypeName")
        End If
        rstGetVal.Close
        Set rstGetVal = Nothing
        getItemTypeName = str
End Function
Public Function getItemName(ItemCode As String) As String
Dim str As String
    Set rstGetVal = FillRecordSet("select ItemName from Item where ItemCode =" & ItemCode)
        If Not (rstGetVal.EOF) Then
            str = rstGetVal("ItemName")
        End If
        rstGetVal.Close
        Set rstGetVal = Nothing
        getItemName = str
End Function
Public Function getPartyName(PartyCode As String) As String
Dim str As String
    Set rstGetVal = FillRecordSet("select PartyName from Party where PartyCode =" & PartyCode)
        If Not (rstGetVal.EOF) Then
            str = rstGetVal("PartyName")
        End If
        rstGetVal.Close
        Set rstGetVal = Nothing
        getPartyName = str
End Function
Function selectValueInCombo(cboObj As ComboBox, value As String)
    Dim i As Integer
    Dim found As Boolean
    found = False
    For i = 0 To cboObj.ListCount - 1
        If (cboObj.ItemData(i) = value) Then
            found = True
            cboObj.ListIndex = i
        End If
    Next
    If found = False And cboObj.ListCount > 0 Then
        cboObj.ListIndex = -1
    End If
    selectValueInCombo = found
End Function
Public Function getFieldValue(Code As String, TableName As String, FieldName As String, WhereFieldName As String) As String
Dim str As String
    Set rstGetVal = FillRecordSet("select " & FieldName & " from " & TableName & " where " & WhereFieldName & " =" & Code)
    
        If Not (rstGetVal.EOF) Then
            str = rstGetVal(" & FieldName & ")
        End If
        rstGetVal.Close
        Set rstGetVal = Nothing
        getFieldValue = str
End Function
Public Function getAvgCost(Code As Integer) As Double
Dim Curr_Avg As Double, a As Integer
    Set rs = FillRecordSet("select AvgCost from vwAvailableQty where ItemCode = " & Code)
    If Not rs.EOF Then
        Curr_Avg = rs("avgCost")
    End If
    rs.Close
    Set rs = Nothing
    getAvgCost = Curr_Avg
End Function
