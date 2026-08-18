Attribute VB_Name = "Módulo1"
Sub PadronizarAbasParaBD()

    Dim pasta As String
    Dim arquivo As String
    Dim wb As Workbook
    Dim contador As Long
    Dim erros As Long
    
    pasta = "Caminho da pasta"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    arquivo = Dir(pasta & "*.xls*")
    
    Do While arquivo <> ""
        
        ' Ignora arquivos temporários do Excel
        If Left(arquivo, 2) <> "~$" Then
            
            On Error Resume Next
            
            Set wb = Nothing
            
            Set wb = Workbooks.Open( _
                Filename:=pasta & arquivo, _
                UpdateLinks:=False, _
                ReadOnly:=False)
            
            If Err.Number <> 0 Or wb Is Nothing Then
                
                erros = erros + 1
                Err.Clear
                
            Else
                
                On Error GoTo 0
                
                ' Renomeia a única aba
                If wb.Worksheets.Count = 1 Then
                    
                    wb.Worksheets(1).Name = "BD"
                    
                    wb.Save
                    contador = contador + 1
                    
                Else
                    
                    erros = erros + 1
                    
                End If
                
                wb.Close SaveChanges:=False
                
            End If
            
            On Error GoTo 0
            
        End If
        
        arquivo = Dir
        
    Loop
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Processo concluído!" & vbCrLf & vbCrLf & _
           "Arquivos processados: " & contador & vbCrLf & _
           "Arquivos com erro/ignorados: " & erros, _
           vbInformation, "Padronização das abas"

End Sub
