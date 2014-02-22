Attribute VB_Name = "obelix"
' Copyright (c) 2010 Nohros Systems Inc. (www.nohros.com)
' Neylor Ohmaly Rodrigues e Silva (neylor.silva@nohros.com)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the "Software"), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
' to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
'
Option Explicit
Option Base 0

Private sqlConn As Object
Public connection_string_ As String

Public Type token
    TokenBegin As Long
    TokenEnd As Long
    TokenWidth As Long
End Type

Function OpenSQLConnection()

    On Error GoTo Catch:
    
    If sqlConn Is Nothing Then
        Set sqlConn = CreateObject("ADODB.Connection")
    End If
       
    If connection_string_ = Empty Then
        frmConnection.Show
    End If
    
    If sqlConn.State = adStateClosed And connection_string_ <> Empty Then
        sqlConn.Provider = "SQLOLEDB"
        sqlConn.Open connection_string_
    End If
    
    GoTo Finally
    
Catch:
    connection_string_ = Empty
    
Finally:
End Function

Function CloseSQLConnection() As Boolean
    If sqlConn Is Nothing Then
        CloseSQLConnection = True
    Else
        sqlConn.Close
    End If
End Function

Function GetFromSQL(ByVal sql As String, Optional closeConn As Boolean = True) As String
    
    On Error GoTo Catch:
    
    OpenSQLConnection
    
    Dim cmd As Object
    Dim rs As Object
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandType = adCmdText
    
    cmd.ActiveConnection = sqlConn
    cmd.CommandText = sql
    
    Set rs = cmd.Execute
    
    GetFromSQL = ReadRecordSet(rs)
    
    DoEvents
    
    If closeConn Then
        CloseSQLConnection
    End If
    
    GoTo Finally

Catch:
    GetFromSQL = Err.Description
    
Finally:
    Set cmd = Nothing
    Set rs = Nothing
    
End Function

Private Function ReadRecordSet(ByRef rs As Object) As String
    ' Read the recordset to a string
    Dim s As String
    Dim f As Object
    
    If rs Is Nothing Then
        s = "NODATA"
    ElseIf rs.State = adStateClosed Then
        s = "NODATA"
    ElseIf rs.EOF Then
        s = "NODATA"
    Else
        While Not rs.EOF
            For Each f In rs.Fields
                If f.Attributes And adFldIsNullable And False Then
                    s = s + ";" + "NULL"
                Else
                    s = s + ";" + CStr(f.Value)
                End If
            Next
            
            rs.MoveNext ' next record
        Wend
    End If
    ReadRecordSet = s
End Function

Private Function ClearConnectionString()
    connection_string_ = Empty
    ClearConnectionString = True
End Function

' Remove caracteres indesejaveis
' uNome : Nome co carecteres a remover
' Retorno : Nome sem os caracteres
Function CGLRFull(ByVal uNome As String) As String
    
    uNome = StrConv(uNome, vbUpperCase)
    
    uNome = Replace(uNome, " E ", " ")
    uNome = Replace(uNome, " DE ", " ")
    uNome = Replace(uNome, " DA ", " ")
    uNome = Replace(uNome, " DO ", " ")
    uNome = Replace(uNome, " DES ", " ")
    uNome = Replace(uNome, " DOS ", " ")
    uNome = Replace(uNome, " DAS ", " ")
    uNome = Replace(uNome, " DAS ", " ")
    
    uNome = Replace(uNome, "Á", "A")
    uNome = Replace(uNome, "À", "A")
    uNome = Replace(uNome, "Â", "A")
    uNome = Replace(uNome, "Ã", "A")
    
    uNome = Replace(uNome, "É", "E")
    uNome = Replace(uNome, "È", "E")
    uNome = Replace(uNome, "Ê", "E")
            
    uNome = Replace(uNome, "Í", "I")
    uNome = Replace(uNome, "Ì", "I")
    uNome = Replace(uNome, "Î", "I")
    
    uNome = Replace(uNome, "Ó", "O")
    uNome = Replace(uNome, "Ò", "O")
    uNome = Replace(uNome, "Õ", "O")
    uNome = Replace(uNome, "Ô", "O")
    
    uNome = Replace(uNome, "Ú", "U")
    uNome = Replace(uNome, "Ù", "U")
    uNome = Replace(uNome, "Û", "U")
    uNome = Replace(uNome, "Ü", "U")
    
    uNome = Replace(uNome, Chr(13), " ")
  
    CGLRFull = uNome
End Function

' Remove os acentos contidos em um nome
' uNome : Nome com acentos
' Retorno : Nome maiusculo sem os acentos
Function RemoverAcento(uNome As String) As String
    
    uNome = StrConv(uNome, vbUpperCase)
    
    uNome = Replace(uNome, "Á", "A")
    uNome = Replace(uNome, "À", "A")
    uNome = Replace(uNome, "Â", "A")
    uNome = Replace(uNome, "Ã", "A")
    
    uNome = Replace(uNome, "É", "E")
    uNome = Replace(uNome, "È", "E")
    uNome = Replace(uNome, "Ê", "E")
            
    uNome = Replace(uNome, "Í", "I")
    uNome = Replace(uNome, "Ì", "I")
    uNome = Replace(uNome, "Î", "I")
    
    uNome = Replace(uNome, "Ó", "O")
    uNome = Replace(uNome, "Ò", "O")
    uNome = Replace(uNome, "Õ", "O")
    uNome = Replace(uNome, "Ô", "O")
    
    uNome = Replace(uNome, "Ú", "U")
    uNome = Replace(uNome, "Ù", "U")
    uNome = Replace(uNome, "Û", "U")
    uNome = Replace(uNome, "Ü", "U")
    
    RemoverAcento = uNome
End Function

' Remove caracteres indesejados do CPF
' ex. ),(,(,.,-
' Realiza a validacao do CPF
' uCpf : Cpf com caracteres indesejaveis
' Retorno : Cpf formatado.
'           Indicador de validade do CPF
Function CpfValido(ByVal uCpf As String) As String

    uCpf = RemoverEspeciais(uCpf)
   
    If (CpfValido_(uCpf)) Then
        CpfValido = uCpf
    Else
        CpfValido = "***CPF Inválido***"
    End If
    
End Function


' Remove caracteres indesejados de uma string
' ex. CPF, RG, telefone
' uStr : String contendo caracteres indesejados
' Retorno : String formatado
Function RemoverEspeciais(ByVal uStr As String) As String
    uStr = Replace(uStr, "-", "")
    uStr = Replace(uStr, ".", "")
    uStr = Replace(uStr, ")", "")
    uStr = Replace(uStr, "(", "")
    uStr = Replace(uStr, "@", "")
    uStr = Replace(uStr, ",", "")
    uStr = Replace(uStr, " ", "")
    
    RemoverEspeciais = uStr
End Function

' Remove caracteres indesejados de uma string
' ex. CPF, RG, telefone
' uStr : String contendo caracteres indesejados
' Retorno : String formatado
Function SomenteLetras(ByVal uStr As String) As String
    uStr = Replace(uStr, "-", "")
    uStr = Replace(uStr, ".", "")
    uStr = Replace(uStr, ")", "")
    uStr = Replace(uStr, "(", "")
    uStr = Replace(uStr, "@", "")
    uStr = Replace(uStr, ",", "")
    
    uStr = Replace(uStr, "1", "")
    uStr = Replace(uStr, "2", "")
    uStr = Replace(uStr, "3", "")
    uStr = Replace(uStr, "4", "")
    uStr = Replace(uStr, "5", "")
    uStr = Replace(uStr, "6", "")
    uStr = Replace(uStr, "7", "")
    uStr = Replace(uStr, "8", "")
    uStr = Replace(uStr, "9", "")
    uStr = Replace(uStr, "0", "")
    
    SomenteLetras = uStr
End Function

'**
'* Verifica se a string especificada e um CPF valido.
'*
'* @param uCpf Texto a ser validado como CPF.
'
'* @return Verdadeiro caso o texto expecificado seja um CPF valido; Falso em caso contrario.
'*/
Private Function CpfValido_(ByVal uCpf As String) As Boolean
    
    Dim soma As Integer
    Dim resultado As Integer
    Dim digitos As String
    Dim revCpf As String
    Dim i As Integer
    
    ' Valida tamanho do CPF
    If Len(uCpf) <> 11 Then
        CpfValido_ = False
        Exit Function
    End If
    
    If Not IsNumeric(uCpf) Then
        CpfValido_ = False
        Exit Function
    End If
        
    ' Multiplicacao da soma dos nove primeiros digitos por
    ' seg.: 10,9,8...2
    revCpf = StrReverse(Mid(uCpf, 1, 9))
    For i = 2 To 10
        soma = soma + i * CInt(Mid(revCpf, i - 1, 1))
    Next i
    
    resultado = soma Mod 11
        
    ' Primeiro digito verificador
    If resultado = 0 Or resultado = 1 Then
        digitos = "0"
    Else
        digitos = CStr(11 - resultado)
    End If
    
    ' Multiplicacao da soma dos nove primeiros digitos por
    ' seg.: 11,10,9...3
    soma = 0
    For i = 3 To 11
        soma = soma + i * CInt(Mid(revCpf, i - 2, 1))
    Next i
    
    soma = soma + CInt(digitos) * 2
    
    resultado = soma Mod 11
    
    ' Segundo digito verificador
    If resultado = 0 Or resultado = 1 Then
        digitos = digitos & "0"
    Else
        digitos = digitos & CStr(11 - resultado)
    End If
    
    CpfValido_ = (digitos = Mid(uCpf, 10, 2))
    
End Function

Function FuzzyMatchByWord(ByVal lsPhrase1 As String, ByVal lsPhrase2 As String, Optional lbStripVowels As Boolean = False, Optional lbDiscardExtra As Boolean = False) As Double

'
' Compare two phrases and return a similarity value (between 0 and 100).
'
' Arguments:
'
' 1. Phrase1        String; any text string
' 2. Phrase2        String; any text string
' 3. StripVowels    Optional to strip all vowels from the phrases
' 4. DiscardExtra   Optional to discard any unmatched words
'
   
    'local variables
    Dim lsWord1() As String
    Dim lsWord2() As String
    Dim ldMatch() As Double
    Dim ldCur As Double
    Dim ldMax As Double
    Dim liCnt1 As Integer
    Dim liCnt2 As Integer
    Dim liCnt3 As Integer
    Dim lbMatched() As Boolean
    Dim lsNew As String
    Dim lsChr As String
    Dim lsKeep As String
   
    'set default value as failure
    FuzzyMatchByWord = 0
   
    'create list of characters to keep
    lsKeep = "BCDFGHJKLMNPQRSTVWXYZ0123456789 "
    If Not lbStripVowels Then
        lsKeep = lsKeep & "AEIOU"
    End If
   
    'clean up phrases by stripping undesired characters
    'phrase1
    lsPhrase1 = Trim$(UCase$(lsPhrase1))
    lsNew = ""
    For liCnt1 = 1 To Len(lsPhrase1)
        lsChr = Mid$(lsPhrase1, liCnt1, 1)
        If InStr(lsKeep, lsChr) <> 0 Then
            lsNew = lsNew & lsChr
        End If
    Next
    lsPhrase1 = lsNew
    lsPhrase1 = Replace(lsPhrase1, "  ", " ")
    lsWord1 = Split(lsPhrase1, " ")
    If UBound(lsWord1) = -1 Then
        Exit Function
    End If
    ReDim ldMatch(UBound(lsWord1))
    'phrase2
    lsPhrase2 = Trim$(UCase$(lsPhrase2))
    lsNew = ""
    For liCnt1 = 1 To Len(lsPhrase2)
        lsChr = Mid$(lsPhrase2, liCnt1, 1)
        If InStr(lsKeep, lsChr) <> 0 Then
            lsNew = lsNew & lsChr
        End If
    Next
    lsPhrase2 = lsNew
    lsPhrase2 = Replace(lsPhrase2, "  ", " ")
    lsWord2 = Split(lsPhrase2, " ")
    If UBound(lsWord2) = -1 Then
        Exit Function
    End If
    ReDim lbMatched(UBound(lsWord2))
   
    'exit if empty
    If Trim$(lsPhrase1) = "" Or Trim$(lsPhrase2) = "" Then
        Exit Function
    End If
   
    'compare words in each phrase
    For liCnt1 = 0 To UBound(lsWord1)
        ldMax = 0
        For liCnt2 = 0 To UBound(lsWord2)
            If Not lbMatched(liCnt2) Then
                ldCur = FuzzyMatch(lsWord1(liCnt1), lsWord2(liCnt2))
                If ldCur > ldMax Then
                    liCnt3 = liCnt2
                    ldMax = ldCur
                End If
            End If
        Next
        lbMatched(liCnt3) = True
        ldMatch(liCnt1) = ldMax
    Next
   
    'discard extra words
    ldMax = 0
    For liCnt1 = 0 To UBound(ldMatch)
        ldMax = ldMax + ldMatch(liCnt1)
    Next
    If lbDiscardExtra Then
        liCnt2 = 0
        For liCnt1 = 0 To UBound(lbMatched)
            If lbMatched(liCnt1) Then
                liCnt2 = liCnt2 + 1
            End If
        Next
    Else
        liCnt2 = UBound(lsWord2) + 1
    End If
   
    'return overall similarity
    FuzzyMatchByWord = 100 * (ldMax / liCnt2)
   
End Function

Function FuzzyMatch(Fstr As String, Sstr As String) As Double
    
    Dim L, L1, L2, M, SC, t, R As Integer
   
    L = 0
    M = 0
    SC = 1
   
    L1 = Len(Fstr)
    L2 = Len(Sstr)
   
    Do While L < L1
        L = L + 1
        For t = SC To L1
            If Mid$(Sstr, L, 1) = Mid$(Fstr, t, 1) Then
                M = M + 1
                SC = t
                t = L1 + 1
            End If
        Next t
    Loop
   
    If L1 = 0 Then
        FuzzyMatch = 0
    Else
        FuzzyMatch = M / L1
    End If

End Function

Function HsbcCode(ByVal doc As String, ByVal cedente As String, ByVal maturityDate As Date)
    Dim chrDate As String
    Dim weight(7) As Integer
    Dim lenght As Integer
    Dim digit As Integer
    Dim digit1 As Integer
    Dim digit2 As Integer
    Dim sum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sacno As Double
    Dim rev As String
    
    weight(0) = 9
    weight(1) = 8
    weight(2) = 7
    weight(3) = 6
    weight(4) = 5
    weight(5) = 4
    weight(6) = 3
    weight(7) = 2
    
    chrDate = Format(maturityDate, "ddmmyy")
    
    lenght = Len(doc)
    j = 0
    rev = StrReverse(doc)
    
    If lenght > 13 Then
        HsbcCode = "*** NOSSO NÚMERO INVÁLIDO ***"
        Exit Function
    End If
    
    ' Calculo do primeiro digito verificador
    For i = 1 To lenght
        digit = CInt(Mid(rev, i, 1))
        sum = sum + digit * weight(j)
        
        j = j + 1
        If j = 8 Then j = 0
    Next i
    
    digit1 = sum Mod 11
    If digit1 = 10 Then _
        digit1 = 0
        
    sacno = CDbl(doc & digit1 & "4") + CDbl(cedente) + CDbl(chrDate)
    
    lenght = Len(CStr(sacno))
    j = 0
    sum = 0
    rev = StrReverse(sacno)
    
    ' Calculo do segundo digito verificador
    For i = 1 To lenght
        digit = CInt(Mid(rev, i, 1))
        sum = sum + digit * weight(j)
        
        j = j + 1
        If j = 8 Then j = 0
    Next i
    
    digit2 = sum Mod 11
    If digit2 = 10 Then _
        digit2 = 0
    
    HsbcCode = CStr(digit1) & "4" & CStr(digit2)
End Function

' Remove caracteres indesejados de uma string
' ex. CPF, RG, telefone
' uStr : String contendo caracteres indesejados
' Retorno : String formatado
Public Function SplitByValue(ByVal row As Integer, ByVal col As Integer)
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim xlCell As Range
    Dim xbCell As Range
    Dim sheetName As String
    
    If col <= 0 Then
        Return
    End If
    
    Set xlCell = ActiveSheet.Cells(row, col)
    ' Loop for each value cells
    Do While xlCell.Value <> Empty
        Set wb = Workbooks.Add
        Set sh = wb.Sheets(1)
        Set xbCell = sh.Range("A2")
        sheetName = xlCell.Text
        ' Loop for cells with same value
        Do While xlCell.Value = sheetName
            xlCell.EntireRow.Copy xbCell.EntireRow
            Set xlCell = xlCell.offset(1, 0)
            Set xbCell = xbCell.offset(1, 0)
        Loop
        wb.SaveAs sheetName
        wb.Close False, sheetName
        ' Goto next cell
        Set xlCell = xlCell.offset(1, 0)
    Loop
End Function


'**
'* Alinha os caracteres do texto especificado adicionando o caracter character_a_preencher
'* a esquerda do texto a ser alinhado.
'* <p> Se o tamanho do texto informado for menor do que o tamanho especificado e o campo
'* remover_caracteres for verdadeio os caracteres mais a direita serão eliminados para que o
'* texto atinja o tamanho especificado.
'*
'* @param texto Texto a ser alinhado.
'* @param tamanho_final Tamanho final que o texto devera conter apos o alinhamento.
'* @param caracter_a_preencher Caracter que sera adicionado a esquerda do texto.
'* @param truncar Verdadeiro se o texto devera ser truncado caso o tamanho do texto seja maior do que o tamanho especificado;
'*        falso em caso contrario.
'*
Function PreencherEsquerda(ByVal texto As String, ByVal caracter_a_preencher As String, ByVal tamanho_final As Integer, Optional truncar As Boolean)
    Dim novo_texto As String
    Dim quantidade_a_adicionar As Integer
    
    novo_texto = texto
    
    quantidade_a_adicionar = tamanho_final - Len(texto)
    If quantidade_a_adicionar < 0 Then
        If truncar Then
            novo_texto = Mid(texto, 1, tamanho_final)
        End If
        GoTo Finally
    End If

    novo_texto = Replace(Space(quantidade_a_adicionar), " ", caracter_a_preencher) + novo_texto
    
Catch:
Finally:
    PreencherEsquerda = novo_texto
End Function

'**
'* Alinha os caracteres do texto especificado adicionando o caracter character_a_preencher
'* a direita do texto a ser alinhado.
'* <p> Se o tamanho do texto informado for menor do que o tamanho especificado e o campo
'* remover_caracteres for verdadeio os caracteres mais a direita serão eliminados para que o
'* texto atinja o tamanho especificado.
'*
'* @param texto Texto a ser alinhado.
'* @param tamanho_final Tamanho final que o texto devera conter apos o alinhamento.
'* @param caracter_a_preencher Caracter que sera adicionado a direita do texto.
'* @param truncar Verdadeiro se o texto devera ser truncado caso o tamanho do texto seja maior do que o tamanho especificado;
'*        falso em caso contrario.
'*
Function PreencherDireita(ByVal texto As String, ByVal caracter_a_preencher As String, ByVal tamanho_final As Integer, Optional truncar As Boolean)
    Dim novo_texto As String
    Dim quantidade_a_adicionar As Integer
    
    novo_texto = texto
    
    quantidade_a_adicionar = tamanho_final - Len(texto)
    If quantidade_a_adicionar < 0 Then
        If truncar Then
            novo_texto = Mid(texto, 1, tamanho_final)
        End If
        GoTo Finally
    End If

    novo_texto = novo_texto + Replace(Space(quantidade_a_adicionar), " ", caracter_a_preencher)
    
Catch:
Finally:
    PreencherDireita = novo_texto
End Function

'**
'* Formata uma sequencia de caracteres substituindo cada elemento de formato
'* no texto especificado pelo texto equivalente especificado.
'* <p>
'* Um formato e composto por um sinal de dolar mais um numero que identifica
'* a ordem em que o texto correspondente aparece na sequencia de formatos.
'* <p>
'* Caso o numero que acompanha a definicao do formato nao exista na lista de
'* textos correspondentes o formato sera substituido por um texto vazio. Por
'* exemplo:
'* <p>
'*     =FORMATARTEXTO("Este texto possui o formato $1.", B2)
'* <p>
'* No exemplo acima o texto $1 sera substituido pelo texto existente na celula B2.
'* <p>
'* A lista de textos pode conter celulas ou textos explicitos.
'*
'* @param texto Texto composto a ser formatado
'* @param formato Objetos contendo os formatos a serem aplicados.
'* @return Copia do formato onde cada item de formato e substituido por seu texto correspondente.
'*/
Function FormatarTexto(ByVal texto As String, ParamArray formatos() As Variant)
    Dim i As Long
    Dim formats_size As Integer
    Dim formated_text As String
    Dim text_size As Integer
    Dim tokens() As String
    Dim ch As String
    Dim format_token_position_string As String
    Dim format_token_position As Integer
    Dim is_digit As Boolean
    Dim lower_bound As Integer
    Dim formato() As Variant
    Dim sub_formato() As Variant
    Dim last_dimenssion As Integer
        
    formato = formatos
    
    ' checks if the format array is multidimensional
    On Error GoTo Singl
    For i = 0 To 60000
        sub_formato = formato(0)
        formato = sub_formato
    Next i
    
Multi:
    formato = formatos(0)
    
Singl:
    On Error GoTo 0
    
    i = 1
    formats_size = UBound(formato)
    text_size = Len(texto)
    
    If formats_size = 0 Then
        formated_text = texto
        GoTo Finally
    End If
    
    Do While i <= text_size
        ch = Mid(texto, i, 1)
        If ch = "$" Then
            format_token_position_string = ""
            i = i + 1
            Do While i <= text_size
                ch = Mid(texto, i, 1)
                
                is_digit = IsDigit(ch)
                If is_digit Then
                    format_token_position_string = ch & format_token_position_string
                End If
                
                If Not is_digit Or i = text_size Then
                    format_token_position = CInt(format_token_position_string) - 1
                    If format_token_position <= formats_size Then
                        formated_text = formated_text & CStr(formato(format_token_position)) + ch 'IIf(i = text_size, "", ch)
                    Else
                        formated_text = formated_text & "$" & CStr(format_token_position - formats_size) + ch
                    End If
                    
                    ' if i is equals text size the last character must be appended to
                    ' the formated text because after exit from here, [i] will be
                    ' greater than the [text_size] and the code that appends
                    ' non-digit characters to the formated text will not be called.
                    'If i = text_size And Not is_digit Then
                        'formated_text = formated_text & ch
                    'End If
                    
                    Exit Do
                End If
                i = i + 1
            Loop
        Else
            formated_text = formated_text & ch
        End If
        i = i + 1
    Loop
    
    GoTo Finally
    
Catch:
    formated_text = Err.Description
    
Finally:
    FormatarTexto = formated_text
End Function

Function MUDARCOR(ByVal cell As Range, ByVal color As Long, ParamArray parts() As Variant)
    Dim words() As String
    Dim tokens() As token
    Dim token As token
    Dim last_token_position As Long
    Dim last_token_width As Long
    Dim i As Long
    
    words = Split(cell.Value, " ")
    
    ReDim tokens(UBound(words))
    
    last_token_position = 0
    last_token_width = 0
    
    For i = 0 To UBound(words)
        With tokens(i)
            .TokenWidth = Len(words(i))
            .TokenBegin = last_token_position + 1
            .TokenEnd = last_token_position + .TokenWidth
            
            last_token_position = .TokenEnd
        End With
    Next i
    
    For i = 0 To UBound(parts)
        token = tokens(parts(i) - 1)
        Range("d7").Characters(token.TokenBegin, token.TokenWidth).Font.color = color
    Next i
    
    MUDARCOR = cell.Value2
End Function

Function RGB(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Long
    RGB = VBA.Information.RGB(red, green, blue)
End Function
'**
'* Concatena todos os elementos especificados, utilizando o separador entre cada elemento.
'*
'* @param separador Texto a ser utilizado como separador.
'* @param elementos Elementos que serao concatenados
Function Juntar(ByVal separador As String, ParamArray elementos() As Variant)
    Juntar = Join(elementos, separador)
End Function


Function FirstSpaceIndex(name) As Integer
    FirstSpaceIndex = InStr(1, name, " ")
End Function

Public Function PrimeiroNome(name) As String
    Dim index As Integer
    Dim trimmed As String
    
    trimmed = Trim(name)
    index = FirstSpaceIndex(trimmed)
    PrimeiroNome = Mid$(trimmed, 1, index - 1)
End Function

Public Function NomeDoMeio(name) As String
    Dim index As Integer
    Dim index2 As Integer
    Dim trimmed As String
    
    trimmed = Trim(name)
    index = FirstSpaceIndex(trimmed)
    trimmed = StrReverse(Mid$(trimmed, index + 1, Len(trimmed) - index))
    
    index = FirstSpaceIndex(trimmed)
    NomeDoMeio = StrReverse(Mid$(trimmed, index + 1, Len(trimmed) - index))
End Function

Public Function UltimoNome(name) As String
    UltimoNome = StrReverse(PrimeiroNome(StrReverse(name)))
End Function

Public Function FixSize(name, size, expr, reverse) As String
    Dim trimmed
    Dim length
    Dim fixed
    
    trimmed = Trim(name)
    fixed = Mid(trimmed, 1, size)
    length = Len(fixed)
    If length < size Then
        If reverse Then
            fixed = Replace(Space(size - length), " ", expr) + fixed
        Else
            fixed = fixed + Replace(Space(size - length), " ", expr)
        End If
    End If
    FixSize = fixed
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    Dim asc_code As Integer
    asc_code = Asc(ch)
    IsDigit = (asc_code > 48 And asc_code < 58)
End Function
