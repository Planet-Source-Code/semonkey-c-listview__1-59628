Attribute VB_Name = "m_Main"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpdirectory As String, ByVal nShowCmd As Long) As Long
Public cnx As ADODB.Connection
Public rs As ADODB.Recordset
Public c_lv As c_Listview
Public NombreBD As String
Public UbicacionBD As String
Public Function OpenFile(UbiFileOrWebSite As String, Form As Object)
    Dim lret    As Long
    lret = ShellExecute(Form.hwnd, "Open", UbiFileOrWebSite, _
        "", "", 1)
End Function
Public Sub main()
    Screen.MousePointer = vbHourglass
    
    Set cnx = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set c_lv = New c_Listview
    
    NombreBD = "bd1.mdb"
    UbicacionBD = App.Path & "\" & NombreBD
    cnx.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
        "Dbq=" & UbicacionBD & ";" & _
        "Uid=;" & _
        "Pwd="
    frm_Main.Show
    Screen.MousePointer = vbDefault
End Sub
