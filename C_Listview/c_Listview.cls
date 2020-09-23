VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Listview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module       : c_Listview
' Author       : SEPONCE, From Chile
' Proposit     : Load a listview with a recordset and export the Records to TXT, CSV and HTML format
'/*************************************************************************************/
'/ I did it this class because  i needed  to reduce the time  in load a listview with
'/ a recordset and reduce the code lines in my aplications for exporting the records to
'/ this format.
'/ Enjoy this class and if you like it please vote.
'/*************************************************************************************/
'---------------------------------------------------------------------------------------

Private Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
Private Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)
Private Header() As Integer
Private i As Integer
Private j As Integer
Dim l As Integer
Private Progress As Object
Private bar As Boolean
Private SubChar As String * 1
Private Enclose_Field As String
Private Field_Separator As String
Private MaxChar As Integer

Public Function Load_ListView(ByVal ListView As ListView, ByVal Recordset As ADODB.Recordset, Optional GridLines = True, Optional Ajustar = True) As Boolean
    
    Dim sResult As Boolean
    Dim i As Integer
    Dim tLi As ListItem
    Dim Contador As Long
    Dim Contador1 As Long
    If Ajustar = True Then
        ReDim Header(1, Val(Recordset.Fields.Count - 1))
        For i = 0 To UBound(Header, 2)
            Header(1, i) = Len(Recordset.Fields(i).Name)
        Next i
        If Not Recordset.EOF Then
            While Not Recordset.EOF
                For i = 0 To UBound(Header, 2)
                    If IsNull(Recordset.Fields(i)) = False Then
                        Obtener_Ancho Recordset.Fields(i), i, MaxChar
                    End If
                Next i
                Recordset.MoveNext
                Contador = Contador + 1
            Wend
            Recordset.MoveFirst
        End If
    End If
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = 3
        .Appearance = ccFlat
        .GridLines = GridLines
        .FullRowSelect = True
        .Font = "Courier New" 'Se puede reemplazar la Fuente pero siempre usar una con paso fijo
        For i = 1 To Recordset.Fields.Count
            .ColumnHeaders.Add , , Recordset.Fields(i - 1).Name
            If Ajustar = True Then
                .ColumnHeaders(i).Width = Header(1, i - 1) * 140.5
            End If
        Next i
    End With
    If Recordset.EOF Then Exit Function
    With ListView.ListItems
        .Clear
        Recordset.MoveFirst
        While Not Recordset.EOF
            Set tLi = ListView.ListItems.Add(, , Recordset.Fields(0))
            For i = 1 To Recordset.Fields.Count - 1
                If IsNull(Recordset.Fields(i)) Then
                    tLi.SubItems(i) = "(Null)"
                Else
                    tLi.SubItems(i) = Left(Recordset.Fields(i), MaxChar)
                End If
            Next i
            ValidateRect ListView.hwnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect ListView.hwnd, 0&, 0&
            Recordset.MoveNext
            Contador1 = Contador1 + 1
            If bar = True Then
                Progress.Value = (Contador1 / Contador) * 100
            End If
        Wend
        sResult = True
    End With
    InvalidateRect ListView.hwnd, 0&, 0&
    Load_ListView = sResult
    Set Recordset = Nothing
    If bar = True Then
        Progress.Value = 0.001
    End If
End Function


Public Function Export_ListviewToTXT(TempFile, ListView, Optional Encabezado = False) As Boolean
    If ListView.ListItems.Count = 0 Then Exit Function
    With ListView
        ReDim Header(1, Val(ListView.ColumnHeaders.Count))
        For i = 1 To UBound(Header, 2)
            Header(1, i) = Len(ListView.ColumnHeaders(i).Text)
        Next i
        For i = 1 To ListView.ListItems.Count
            Obtener_Ancho ListView.ListItems(i), 1
            For j = 1 To ListView.ColumnHeaders.Count - 1
                Obtener_Ancho ListView.ListItems(i).SubItems(j), j + 1, MaxChar
            Next
        Next
        Open TempFile For Output As #1
        If Encabezado = True Then
            For i = 1 To .ColumnHeaders.Count
                Print #1, .ColumnHeaders(i).Text;
                Print #1, Space(Header(1, i) - Len(.ColumnHeaders(i).Text)) & Space(2);
            Next
            Print #1,
            For l = 1 To .ColumnHeaders.Count
                Print #1, CrearCadena(SubChar, Header(1, l)) & Space(2);
            Next
            Print #1,
        End If
        For i = 1 To .ListItems.Count
            Print #1, .ListItems(i) & Space(Header(1, 1) - Len(.ListItems(i))) & Space(2);
            For j = 1 To .ColumnHeaders.Count - 1
                If Len(.ListItems(i).SubItems(j)) = 0 Then
                    Print #1, Space(Header(1, j + 1)) & Space(2);
                ElseIf Len(.ListItems(i).SubItems(j)) >= MaxChar Then
                    Print #1, Left(.ListItems(i).SubItems(j), MaxChar) & Space(2);
                Else
                    Print #1, .ListItems(i).SubItems(j) & Space(Header(1, j + 1) - Len(.ListItems(i).SubItems(j))) & Space(2);
                End If
            Next j
            Print #1,
            If bar = True Then
                Progress.Value = (i / ListView.ListItems.Count) * 100
            End If
        Next i
        Export_ListviewToTXT = True
        Close #1
        If bar = True Then
            Progress.Value = 0.001
        End If
    End With
End Function

Public Function Export_ListviewToCSV(ListView As ListView, sFile As String, Optional Encabezado = True) As Boolean
    On Error GoTo ContErr
    If ListView.ListItems.Count = 0 Then Exit Function
    
    Open sFile For Output As #1
    If Encabezado = True Then
        For i = 1 To ListView.ColumnHeaders.Count
            Print #1, Enclose_Field & ListView.ColumnHeaders(i).Text & Enclose_Field & Field_Separator;
        Next i
        Print #1,
    End If
    For i = 1 To ListView.ListItems.Count
        Print #1, Enclose_Field & ListView.ListItems(i) & Enclose_Field & Field_Separator;
        For j = 1 To ListView.ColumnHeaders.Count - 1
            Print #1, Enclose_Field & ListView.ListItems(i).SubItems(j) & Enclose_Field & Field_Separator;
        Next
        Print #1,
        If bar = True Then
            Progress.Value = (i / ListView.ListItems.Count) * 100
        End If
    Next i
    Print #1,
    Export_ListviewToCSV = True
    Close #1
    If bar = True Then
        Progress.Value = 0.001
    End If
    Exit Function
ContErr:
    MsgBox Err.Description & ", Imposible escribir en el archivo por que se encuentra abierto por otra aplicación", vbCritical
End Function


Public Function Export_LisviewToHTML(ListView As ListView, TempFile As String, Optional Title As String, Optional HojaCSS As String) As Boolean
    
    Dim i As Integer
    If ListView.ListItems.Count = 0 Then Exit Function
    Open TempFile For Output As #1
    Print #1, "<html>" & vbNewLine & _
        "<head>" & vbNewLine & _
        "<title>" & Title & "</title>" & vbNewLine & _
        "<link href='" & HojaCSS & "' rel='stylesheet' type='Text/css'>" & vbNewLine & _
        "</head>" & vbNewLine & _
        "<body>" & vbNewLine & _
        "<table width='75%' border='1' align='center'>" & vbNewLine
    Print #1, "<tr>"
    ReDim Header(1, Val(ListView.ColumnHeaders.Count))
    For i = 1 To UBound(Header, 2)
        Header(1, i) = Len(ListView.ColumnHeaders(i).Text)
    Next i
    For i = 1 To ListView.ListItems.Count
        Obtener_Ancho ListView.ListItems(i), 1
        For j = 1 To ListView.ColumnHeaders.Count - 1
            Obtener_Ancho ListView.ListItems(i).SubItems(j), j
        Next
    Next
    Print #1, "<th>" & ListView.ColumnHeaders(1) & "</th>"
    For i = 1 To ListView.ColumnHeaders.Count - 1
        Print #1, "<th>" & ListView.ColumnHeaders(i + 1) & Crear_Espacios(Header(1, i) - Len(ListView.ColumnHeaders(i))) & "</th>"
    Next i
    Print #1, "</tr>"
    
    For i = 1 To ListView.ListItems.Count
    Print #1, "<tr>"
        Print #1, "<td>" & ListView.ListItems(i) & "</td>"
        For j = 1 To ListView.ColumnHeaders.Count - 1
            If Len(ListView.ListItems(i).SubItems(j)) = 0 Then
                Print #1, "<td>&nbsp;</td>"
            Else
                Print #1, "<td>" & ListView.ListItems(i).SubItems(j) & "</td>"
            End If
        Next
        Print #1, "</tr>"
        If bar = True Then
            Progress.Value = (i / ListView.ListItems.Count) * 100
        End If
    Next
    Print #1, "</table>"
    Print #1, "</body>" & vbNewLine & "</html>"
    Export_LisviewToHTML = True
    Close #1
    If bar = True Then
        Progress.Value = 0.001
    End If
End Function

Public Property Let MaxCharForFields(ByVal vnewval As Integer)
    MaxChar = vnewval
End Property

Public Property Let SubCharTXT(ByVal vnewval As String)
    SubChar = Left(vnewval, 1)
End Property

Public Property Let FieldEncloseCSV(ByVal vnewval As String)
    
    If vnewval = Empty Then
        Enclose_Field = ""
    Else
        Enclose_Field = Left(vnewval, 1)
    End If
End Property

Public Property Let FieldSeparatorCSV(ByVal vnewval As String)
    Field_Separator = Left(vnewval, 1)
End Property

Public Property Let Progresbar(ByVal ProgressBar As Object)
    If IsProgressbar(ProgressBar) = True Then
        bar = True
        Set Progress = ProgressBar
    Else
        bar = False
    End If
End Property

Private Sub Class_Initialize()
    SubChar = "*"
    Field_Separator = ";"
    MaxChar = 35
End Sub

Private Function Obtener_Ancho(Campo As String, indice As Integer, Optional Max = 35) As Long
    Dim largo As Integer
    largo = Len(Campo)
    If largo >= Max Then
        Header(1, indice) = Max
        Obtener_Ancho = Max
    ElseIf largo > Header(1, indice) And largo < Max Then
        Header(1, indice) = largo
        Obtener_Ancho = largo
    Else
        Obtener_Ancho = Header(1, indice)
        Exit Function
    End If
End Function

Private Function Crear_Espacios(Espacios As Integer) As String
    For j = 1 To Espacios
        Crear_Espacios = Crear_Espacios & "&nbsp;"
    Next j
End Function

Private Function CrearCadena(Caracter As String, largo As Integer) As String
    Dim bResult As String
    bResult = Caracter
    For i = 1 To largo - 1
        bResult = bResult & Caracter
    Next i
    CrearCadena = bResult
End Function

Private Function IsProgressbar(ByVal Progresbar As Object) As Boolean
    If TypeName(Progresbar) = "ProgressBar" Then
        IsProgressbar = True
    End If
End Function

