Attribute VB_Name = "obelix_base"

'**
'* Formata uma sequencia de caracteres substituindo cada elemento de formato
'* no texto especificado pelo texto equivalente especificado.
'*
'* Um formato e composto por um sinal de dolar mais um numero que identifica
'* a ordem em que o texto correspondente aparece na sequencia de formatos.
'*
'* Caso o numero que acompanha a definicao do formato nao exista na lista de
'* textos correspondentes o formato sera substituido por um texto vazio. Por
'* exemplo:
'*
'*   =FORMATARTEXTO("Este texto possui o formato $1.", B2)
'*
'* No exemplo acima o texto $1 sera substituido pelo texto existente na celula B2.
'*
'* A lista de textos pode conter celulas ou textos explicitos.
'*
'* @param texto Texto composto a ser formatado
'* @param formato Objetos contendo os formatos a serem aplicados.
'* @return Copia do formato onde cada item de formato e substituido por seu texto correspondente.
'*/
Public Function FormatarTexto(ByVal texto As String, ParamArray formatos() As Variant) As String
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
Public Function PreencherEsquerda( _
  ByVal texto As String, _
  ByVal caracter_a_preencher As String, _
  ByVal tamanho_final As Integer, _
  Optional truncar As Boolean _
) As String

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
'*
'* Se o tamanho do texto informado for menor do que o tamanho especificado e o campo
'* remover_caracteres for verdadeio os caracteres mais a direita serão eliminados para que o
'* texto atinja o tamanho especificado.
'*
'* @param texto Texto a ser alinhado.
'* @param tamanho_final Tamanho final que o texto devera conter apos o alinhamento.
'* @param caracter_a_preencher Caracter que sera adicionado a direita do texto.
'* @param truncar Verdadeiro se o texto devera ser truncado caso o tamanho do texto seja maior do que o tamanho especificado;
'*        falso em caso contrario.
'*
Function PreencherDireita( _
  ByVal texto As String, _
  ByVal caracter_a_preencher As String, _
  ByVal tamanho_final As Integer, _
  Optional truncar As Boolean _
)
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
'* Retorna um valor inteiro representando um valor de cor RGB a partir dos componentes
'* vermelho, verde e azul da cor.
'*
'* @param red The intensity of the red color
'* @param green The intensity of the green color
'* @param blue The intensity of the blue color
Function RGB( _
  ByVal red As Integer, _
  ByVal green As Integer, _
  ByVal blue As Integer _
) As Long
  RGB = VBA.Information.RGB(red, green, blue)
End Function

'**
'* Concatena todos os elementos especificados, utilizando o separador entre cada elemento.
'*
'* @param separador Texto a ser utilizado como separador.
'* @param elementos Elementos que serao concatenados
Public Function Juntar( _
  ByVal separador As String, _
  ParamArray elementos() As Variant _
) As String
  Juntar = Join(elementos, separador)
End Function

'**
'* Verifica se determinado caracter rerpesenta um digito, ou seja um numero inteiro
'* compreendido entre o 0 e 9
'*
'* @param red The intensity of the red color
'* @param green The intensity of the green color
'* @param blue The intensity of the blue color
Function IsDigit(ByVal ch As String) As Boolean
  Dim asc_code As Integer
  
  asc_code = Asc(ch)
  IsDigit = (asc_code > 48 And asc_code < 58)
End Function
