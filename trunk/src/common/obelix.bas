Attribute VB_Name = "obelix"
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
Function FORMATARTEXTO(ByVal texto As String, ParamArray formato() As Variant)
    Dim i As Integer
    Dim formats_size As Integer
    Dim formated_text As String
    Dim text_size As Integer
    Dim tokens() As String
    Dim ch As String
    Dim format_token_position_string As String
    Dim format_token_position As Integer
    Dim is_digit As Boolean
    
    i = 1
    formats_size = UBound(formato)
    text_size = Len(texto)
    
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
                        formated_text = formated_text & CStr(formato(format_token_position)) + IIf(i = text_size, "", ch)
                    End If
                    
                    ' if i is equals text size the last character must be appended to
                    ' the formated text because after exit from here, [i] will be
                    ' greater than the [text_size] and the code that appends
                    ' non-digit characters to the formated text will not be called.
                    If i = text_size And Not is_digit Then
                        formated_text = formated_text & ch
                    End If
                    
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
    FORMATARTEXTO = formated_text
End Function

Function IsDigit(ByVal ch As String) As Boolean
    Dim asc_code As Integer
    asc_code = Asc(ch)
    IsDigit = (asc_code > 48 And asc_code < 58)
End Function

'**
'* Returns a specified date with the specified number interval added to a specified datepart of that date
'*
'* @param datepart Is the part of date to which an integer number is added. The following table list
'* all valid datepart arguments.
'* <p>  year yy, yyyy
'* <p>  quarter qq, q
'* <p>  month mm, m
'* <p>  dayofyear dy, y
'* <p>  day dd, d
'* <p>  week wk, ww
'* <p>  weekday dw, w
'* <p>  hour hh
'* <p>  minute mi, n
'* <p>  decond ss, s
'* <p>  millisecond ms
'* <p>  microsecond mcs
'* <p>  nanosecond ns
'* @param number The number that is added to a datepart of date. If you specify a value with a decimal fraction,
'*               the fraction is truncated and not rounded.
'* @param base_date The date to add the datepart
Public Function SOMARDATA(ByVal datepart As String, ByVal number As Double, ByVal base_date As Date) As Date
    SOMARDATA = DateAdd("m", number, base_date)
End Function
