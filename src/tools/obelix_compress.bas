Attribute VB_Name = "obelix_compress"
' Copyright (c) 2010 Nohros Systems Inc.
' Copyright (c) 2003 Dermot Balson.
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
'
' This module contains compression code written by Dermot Balson. It uses well known and modern
' compression algorithms notably the Burrows-Wheeler transform, assisted by Bring to Front,
' Run Length Encoding, and Arithmetic coding.
'
Option Explicit
Option Base 1

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public SortList() As Long
Public GroupList() As Long, GroupSize() As Long
Public NewGroupList() As Long, NewGroupSize() As Long
Public OrigSort() As Long, GroupOrder() As Long
Public nInBytes As Long, nGroups As Long, nNewGroups As Long
Public Debugging As Boolean

Public sList() As Long
Public nBytes As Long

Public Const StartProbAdjust = 1
Public Const MaxProbAdjust = 1000
Public Const AdjustFactor = 0.5
Public Const ProbAdd = 30
Public Const StartP = 4
Public Const BWT_Chunk = 5000000
Public Const B32 = 4294967296#
Public Const B24 = 16777216
Public Const B16 = 65536

Sub CompressData(B1() As Byte, B2() As Byte)
    Dim m&
    
    m = RLE(B1(), B2())
    BWT B2(), B1()
    BTF B1(), B2()
    m = RLE(B2(), B1())
    Ari B1(), B2()
End Sub

Sub DecompressData(B1() As Byte, B2() As Byte)
    Dim m&
    
    Undo_Ari B1(), B2()
    m = Undo_RLE(B2(), B1())
    Undo_BTF B1(), B2()
    Undo_BWT B2(), B1()
    m = Undo_RLE(B1(), B2())
End Sub

'Arithmetic coding
Sub Ari(InBytes() As Byte, OutBytes() As Byte, Optional Debugging As Boolean)
    Dim P(0 To 255) As Long, cumP(-1 To 255) As Long
    Dim cumStep&, cumDiv&
    Dim i As Long, j As Long, k As Long, bNo As Long
    Dim L As Double, R As Double
    Dim rr As Double, w As Long
    Dim u As Long, V As Single, y As Long
    Dim N&
    Dim nonFF As Byte, nFF As Long
    Dim LastProbUpdate As Long
    Dim BeenHereBefore As Boolean, BB As Byte
    
    N = UBound(InBytes)
    ReDim OutBytes(N * 2 + 3)
    
    P(0) = StartP
    cumP(0) = P(0)
    For i = 1 To 255
      P(i) = StartP
      cumP(i) = cumP(i - 1) + P(i)
    Next i
    
    cumStep = StartProbAdjust
    cumDiv = MaxProbAdjust
    LastProbUpdate = StartProbAdjust
    
    w = 3 'number of characters output, leave the first 3 to count the total bytes
    L = 0 'left hand value
    R = B32
    
    For i = 1 To N
      bNo = InBytes(i)
      rr = Int(R / cumP(255))
      L = L + rr * cumP(bNo - 1)
      
      R = rr * (cumP(bNo) - cumP(bNo - 1))
        
      If L >= B32 Then
        L = L - B32
        nonFF = nonFF + 1
        For u = 1 To nFF
          w = w + 1
          OutBytes(w) = nonFF
          nonFF = 0
        Next u
        nFF = 0
      End If
      
      Do While R <= B24
        BB = Int(L / B24)
      
        If Not BeenHereBefore Then
          nonFF = BB
          nFF = 0
          BeenHereBefore = True
        ElseIf BB = 255 Then
          nFF = nFF + 1
        Else
          w = w + 1
          OutBytes(w) = nonFF
          For u = 1 To nFF
            w = w + 1
            OutBytes(w) = 255
          Next u
          nFF = 0
          nonFF = BB
        End If
        
        L = (L - CDbl(BB) * B24) * 256
        R = R * 256
        
      Loop
    
      'update frequency count
      P(bNo) = P(bNo) + ProbAdd
      'update cumulative stats if we reach the next step
      If i = cumStep Then
    
        'every 1000 or so steps, divide all the values by 2 to "age" them and give more weight to
        'subsequent items
        If cumStep > cumDiv Then 'age the stats
          V = 1 / AdjustFactor
            If P(0) < V Then P(0) = 1 Else P(0) = P(0) / V
            cumP(0) = P(0)
          For j = 1 To 255
            If P(j) < V Then P(j) = 1 Else P(j) = P(j) / V
            cumP(j) = cumP(j - 1) + P(j)
          Next j
          cumDiv = cumDiv + MaxProbAdjust
        Else 'don't age the stats
          cumP(0) = P(0)
          For j = 1 To 255
            cumP(j) = cumP(j - 1) + P(j)
          Next j
        End If
        
        'increase the size of the next step just a little
        cumStep = cumStep + LastProbUpdate
        If LastProbUpdate < 1000 Then LastProbUpdate = LastProbUpdate + 5
        
      End If
    
    Next i
    
    Do While L > 0
      BB = Int(L / B24)
      If BB = 255 Then
        nFF = nFF + 1
      Else
        w = w + 1
        OutBytes(w) = nonFF
        For u = 1 To nFF
          w = w + 1
          OutBytes(w) = 255
        Next u
        nFF = 0
        nonFF = BB
      End If
      L = (L - CDbl(BB) * B24) * 256
    Loop
    
    If nonFF > 0 Then
      w = w + 1
      OutBytes(w) = nonFF
    End If
    For u = 1 To nFF
      w = w + 1
      OutBytes(w) = 255
    Next u
    
    If w < 6 Then w = 6
    ReDim Preserve OutBytes(w)
    
    u = Int(N / 256)
    OutBytes(3) = N - u * 256
    N = u
    u = Int(N / 256)
    OutBytes(2) = N - u * 256
    OutBytes(1) = u

End Sub

Sub Undo_Ari(InBytes() As Byte, OutBytes() As Byte, Optional Debugging As Boolean)
    Dim P(0 To 255) As Long, cumP(-1 To 255) As Long
    Dim cumStep&, cumDiv&
    Dim i As Long, j As Long, k As Long, bNo As Long, nBytes As Long
    Dim R As Double, w As Long
    Dim u As Long, V As Single, y As Long, nW As Long
    Dim N&, D As Double, rr As Double
    Dim LastProbUpdate As Long
    Dim x As Single

    LastProbUpdate = StartProbAdjust
    
    nBytes = (CLng(InBytes(1)) * 256 + InBytes(2)) * 256 + InBytes(3)
    
    N = UBound(InBytes)
    nW = N * 2
    ReDim OutBytes(nBytes)
    
    'initialise cumulative probability array
    P(0) = StartP
    cumP(0) = P(0)
    For i = 1 To 255
      P(i) = StartP
      cumP(i) = cumP(i - 1) + P(i)
    Next i
    cumStep = StartProbAdjust
    cumDiv = MaxProbAdjust
    
    'read in the first 3 bytes
    D = ((CDbl(InBytes(4)) * 256 + InBytes(5)) * 256 + InBytes(6)) * 256 + InBytes(7)
    i = 7
    R = B32
    
    For w = 1 To nBytes
      
      rr = Int(R / cumP(255))
      u = Int(D / rr)
      
      For bNo = 0 To 255
        If u < cumP(bNo) Then Exit For
      Next bNo
      If bNo > 255 Then If u = cumP(255) Then bNo = 255 Else bNo = 255: Stop
        
      D = D - Int(rr * cumP(bNo - 1))
      
      R = rr * (cumP(bNo) - cumP(bNo - 1))
        
      Do While R <= B24
        R = R * 256
        i = i + 1
        If i <= N Then
          D = D * 256 + InBytes(i)
        Else
          D = D * 256
        End If
      Loop
      
      OutBytes(w) = bNo
      
      'update frequency count
      P(bNo) = P(bNo) + ProbAdd
      
      'update cumulative stats if we reach the next step
      If w = cumStep Then
    
        'every 1000 or so steps, divide all the values by 2 to "age" them and give more weight to
        'subsequent items
        If cumStep > cumDiv Then 'age the stats
          V = 1 / AdjustFactor
            If P(0) < V Then P(0) = 1 Else P(0) = P(0) / V
            cumP(0) = P(0)
          For j = 1 To 255
            If P(j) < V Then P(j) = 1 Else P(j) = P(j) / V
            cumP(j) = cumP(j - 1) + P(j)
          Next j
          cumDiv = cumDiv + MaxProbAdjust
        Else 'don't age the stats
          cumP(0) = P(0)
          For j = 1 To 255
            cumP(j) = cumP(j - 1) + P(j)
          Next j
        End If
        
        'increase the size of the next step just a little
        cumStep = cumStep + LastProbUpdate
        If LastProbUpdate < 1000 Then LastProbUpdate = LastProbUpdate + 5
        
      End If
  
    Next w

    ReDim Preserve OutBytes(nBytes)
End Sub

'Bring to front
Sub BTF(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim i As Long, j As Long
    Dim tPosString As String '* 256
    Dim tChr As String * 1
    Dim tPos(0 To 255) As Long
    Dim N As Long

    N = UBound(InBytes)
    ReDim OutBytes(N)
    
    For i = 0 To 255
      tPosString = tPosString & Chr$(i)
    Next i
    
    For i = 1 To N
      tChr = Chr$(InBytes(i))
      OutBytes(i) = InStr(tPosString, tChr) - 1
      CopyMemory ByVal StrPtr(tPosString) + 2, ByVal StrPtr(tPosString), OutBytes(i) * 2
      CopyMemory ByVal StrPtr(tPosString), ByVal StrPtr(tChr), 2
    Next i
    
    Erase tPos
    For i = 1 To N
      j = OutBytes(i)
      tPos(j) = tPos(j) + 1
    Next i
End Sub

Sub Undo_BTF(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim i As Long, j As Long
    Dim tPosString As String '* 256
    Dim tChr As String * 1
    Dim tPos(0 To 255) As Integer
    Dim N As Long
    
    N = UBound(InBytes)
    ReDim OutBytes(N)
    
    For i = 0 To 255
      tPosString = tPosString & Chr$(i)
    Next i
    
    For i = 1 To N
      tChr = Mid$(tPosString, InBytes(i) + 1, 1)
      OutBytes(i) = Asc(tChr)
      CopyMemory ByVal StrPtr(tPosString) + 2, ByVal StrPtr(tPosString), InBytes(i) * 2
      CopyMemory ByVal StrPtr(tPosString), ByVal StrPtr(tChr), 2
    Next i
End Sub

Sub BWT(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim N&, u1&, u2&, u3&, u4&
    N = UBound(InBytes)
    
    If N > BWT_Chunk - 4 Then
      Dim B1() As Byte, B2() As Byte
      Do While u2 < N
        u2 = u1 + BWT_Chunk - 4: If u2 > N Then u2 = N
        ReDim B1(u2 - u1)
        CopyMemory ByVal VarPtr(B1(1)), ByVal VarPtr(InBytes(u1 + 1)), (u2 - u1)
        SortByteArray B1(), B2()
        u4 = UBound(B2)
        If u3 > 0 Then
          ReDim Preserve OutBytes(UBound(OutBytes) + u4)
        Else
          ReDim OutBytes(u4)
        End If
        CopyMemory OutBytes(u3 + 1), B2(1), u4
        u1 = u2
        u3 = u3 + u4
      Loop
    Else
      SortByteArray InBytes(), OutBytes()
    End If
End Sub

Sub Undo_BWT(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim N&, u1&, u2&, u3&, u4&
    N = UBound(InBytes)
    If N > BWT_Chunk Then
      Dim B1() As Byte, B2() As Byte
      Do While u2 < N
        u2 = u1 + BWT_Chunk: If u2 > N Then u2 = N
        ReDim B1(u2 - u1)
        CopyMemory ByVal VarPtr(B1(1)), ByVal VarPtr(InBytes(u1 + 1)), (u2 - u1)
        Decode_BWT B1(), B2()
        u4 = UBound(B2)
        If u3 > 0 Then
          ReDim Preserve OutBytes(UBound(OutBytes) + u4)
        Else
          ReDim OutBytes(u4)
        End If
        CopyMemory OutBytes(u3 + 1), B2(1), u4
        u1 = u2
        u3 = u3 + u4
      Loop
    Else
      Decode_BWT InBytes(), OutBytes()
    End If
End Sub

Sub Decode_BWT(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim i As Long, j As Long, T As Single
    Dim tStartByte As Long
    Dim N As Long
    
    N = UBound(InBytes)
    ReDim OutBytes(N - 4)
    ReDim sList(N)
    
    'get starting item, stored in first 4 bytes
    tStartByte = 0
    j = 1
    For i = 1 To 3
      tStartByte = tStartByte + InBytes(i) * j
      j = j * 256
    Next i
    tStartByte = tStartByte + InBytes(i) * j
    
    Dim C1(0 To 255) As Long
    
    For i = 5 To N
      j = InBytes(i)
      C1(j) = C1(j) + 1
    Next i
    
    For i = 1 To 255
      C1(i) = C1(i) + C1(i - 1)
    Next i
    
    For i = N To 5 Step -1
      j = InBytes(i)
      sList(C1(j)) = i - 4
      C1(j) = C1(j) - 1
    Next i
    
    j = tStartByte
    For i = 1 To N - 4
      OutBytes(i) = InBytes(sList(j) + 4)
      j = sList(j)
    Next i
End Sub

Function ReadFile(tFile As String, OutBytes() As Byte) As Long
    Dim ihwndFile As Integer
    
    'On Error GoTo ErrFailed
    'Open file
    ihwndFile = FreeFile
    
    Open tFile For Binary Access Read As #ihwndFile
    'Size the array to hold the file contents
    ReDim OutBytes(1 To LOF(ihwndFile))

    Get #ihwndFile, , OutBytes
    Close #ihwndFile
       
    ReadFile = UBound(OutBytes)
End Function

Function WriteFile(tFile As String, InBytes() As Byte) As Long
    Dim ihwndFile As Integer
    
    'On Error GoTo ErrFailed
    'Open file
    ihwndFile = FreeFile
    
    Open tFile For Output As #ihwndFile
    Close
    
    Open tFile For Binary Access Write As #ihwndFile
    'Size the array to hold the file contents
    ReDim OutBytes(1 To UBound(InBytes))
    
    Put #ihwndFile, , InBytes
    Close #ihwndFile
       
    WriteFile = UBound(InBytes)
       
End Function

Function RLE(ByRef InBytes() As Byte, OutBytes() As Byte) As Long
    Dim i&, j&, m&
    Dim c&, T&, u&
    Dim N As Long
    
    N = UBound(InBytes)
    
    ReDim OutBytes(N * 1.5) As Byte
    
    j = 0
    
    For i = 1 To N
    
      j = j + 1
      OutBytes(j) = InBytes(i)
      
      If InBytes(i) = c Then
        
        m = -1
        
        Do
          m = m + 1
          If i = N Then
            Exit Do
          End If
          i = i + 1
          If m = 255 Then
            Exit Do
          End If
        Loop While InBytes(i) = c
        
        If m >= 0 Then
          j = j + 1
          OutBytes(j) = m
        End If
        
        If InBytes(i) <> c Or m = 255 Then
          j = j + 1
          OutBytes(j) = InBytes(i)
        End If
        
      End If
      
      c = InBytes(i)
    
    Next i
    
    RLE = j
    ReDim Preserve OutBytes(j)
End Function

Function Undo_RLE(ByRef InBytes() As Byte, OutBytes() As Byte) As Long
    Dim i&, j&, m&, k&
    Dim c&, T&, u&
    Dim N As Long
    
    N = UBound(InBytes)
    k = N * 3
    ReDim tmp(k) As Byte
    
    j = 0
    
    For i = 1 To N
    
      j = j + 1
      If j > k Then
        k = k + N
        ReDim Preserve tmp(k)
      End If
      tmp(j) = InBytes(i)
      
      If InBytes(i) = c Then
        'If i > 55 Then Stop
        i = i + 1
        If i > N Then Exit For
        u = InBytes(i)
        For m = 1 To u
          j = j + 1
          If j > k Then
            k = k + N
            ReDim Preserve tmp(k)
          End If
          tmp(j) = c
        Next m
        c = -1
      Else
        c = InBytes(i)
      End If
    
    Next i
    
    Undo_RLE = j
    ReDim OutBytes(j)
    
    For i = 1 To j
      OutBytes(i) = tmp(i)
    Next i
End Function

Sub SortByteArray(ByRef InBytes() As Byte, ByRef OutBytes() As Byte)
    Dim i As Long, j As Long, k As Long, StartGroup As Long, nCurrGroups As Long
    Dim N As Long, m As Long, u As Long, V As Long, w As Long, z As Long, x As Long
    Dim aSize As Long, StartByte As Long
    
    nInBytes = UBound(InBytes)
    
    ReDim OrigSort(0 To nInBytes)
    ReDim GroupOrder(nInBytes) As Long
    Dim C1(0 To 255, 0 To 255) As Long
    ReDim SortList(nInBytes)
    u = nInBytes '* 5 + 10000
    ReDim GroupList(0 To u), GroupSize(0 To u)
    
    Erase C1
    For i = 1 To nInBytes - 1
      u = InBytes(i)
      V = InBytes(i + 1)
      C1(u, V) = C1(u, V) + 1
    Next i
    C1(InBytes(nInBytes), InBytes(1)) = C1(InBytes(nInBytes), InBytes(1)) + 1
    
    m = 1: nGroups = 0: GroupList(0) = 0
    
    For i = 0 To 255
      For j = 0 To 255
        If C1(i, j) > 0 Then
          nGroups = nGroups + 1
          u = m + C1(i, j) - 1
          For k = m To u
            GroupOrder(k) = m
          Next k
          GroupList(nGroups) = m
          GroupSize(nGroups - 1) = m - GroupList(nGroups - 1)
          m = m + C1(i, j)
          C1(i, j) = m - 1
        End If
      Next j
    Next i
    If m > GroupList(nGroups) Then
      GroupSize(nGroups) = m - GroupList(nGroups)
    End If
    
    SortList(C1(InBytes(nInBytes), InBytes(1))) = nInBytes
    OrigSort(nInBytes) = GroupOrder(C1(InBytes(nInBytes), InBytes(1)))
    C1(InBytes(nInBytes), InBytes(1)) = C1(InBytes(nInBytes), InBytes(1)) - 1
    
    For i = nInBytes - 1 To 1 Step -1
      u = InBytes(i)
      V = InBytes(i + 1)
      w = C1(u, V)
      SortList(w) = i
      OrigSort(i) = GroupOrder(w)
      w = w - 1
      C1(u, V) = w
    Next i
    
    'make copy of original sort group list so we can update it
    aSize = UBound(OrigSort) - LBound(OrigSort) + 1
    ReDim GroupOrder(0 To nInBytes)
    CopyMemory ByVal VarPtr(GroupOrder(0)), ByVal VarPtr(OrigSort(0)), LenB(OrigSort(0)) * aSize
    
    w = 1
    Do
      w = w * 2
       
      nNewGroups = 0
      ReDim NewGroupList(0 To nInBytes), NewGroupSize(0 To nInBytes)
    
      StartGroup = 1 'nCurrGroups + 1
      nCurrGroups = nNewGroups
      
      For i = 1 To nGroups
        If GroupSize(i) > 1 Then
          SortGroup i, w
        End If
      Next i
      
      'replace sort group list with new version
      CopyMemory ByVal VarPtr(OrigSort(0)), ByVal VarPtr(GroupOrder(0)), LenB(GroupOrder(0)) * aSize
      
      If nNewGroups = 0 Then Exit Do
      
      CopyMemory ByVal VarPtr(GroupList(0)), ByVal VarPtr(NewGroupList(0)), LenB(NewGroupList(0)) * UBound(NewGroupList)
      CopyMemory ByVal VarPtr(GroupSize(0)), ByVal VarPtr(NewGroupSize(0)), LenB(NewGroupSize(0)) * UBound(NewGroupSize)
      nGroups = nNewGroups
      
    Loop
    
    ReDim OutBytes(nInBytes + 4)
    
    For i = 1 To nInBytes
      j = SortList(i) - 1
      If j = 0 Then
        j = nInBytes
        StartByte = i
      End If
      OutBytes(i + 4) = InBytes(j)
    Next i
    
    For i = 1 To 3
      j = Int(StartByte / 256)
      OutBytes(i) = StartByte - j * 256
      StartByte = j
    Next i
    OutBytes(4) = StartByte
End Sub


Function SortGroup(GroupNo As Long, Depth As Long) As Long
    Dim i As Long, j As Long, k As Long, m As Long, g As Long, u As Long, V As Long, w As Long, z As Long
    Dim index As Long, Index2 As Long, FirstItem As Long, Distance As Long, value As Long, NumEls As Long, OrigIndex As Long
    
    FirstItem = GroupList(GroupNo)
    NumEls = GroupSize(GroupNo) - 1
    ReDim tGroup(NumEls)
    
    Distance = 0
    Do
      Distance = Distance * 3 + 1
    Loop Until Distance > NumEls + 1
    
    Do
      Distance = Distance \ 3
      For index = FirstItem + Distance To FirstItem + NumEls
        value = SortList(index)
        u = value + Depth: Do While u > nInBytes: u = u - nInBytes: Loop
        z = OrigSort(u)
        Index2 = index
        Do
          w = SortList(Index2 - Distance) + Depth: Do While w > nInBytes: w = w - nInBytes: Loop
          
          If OrigSort(w) <= z Then Exit Do
          SortList(Index2) = SortList(Index2 - Distance)
          Index2 = Index2 - Distance
          If Index2 < FirstItem + Distance Then Exit Do
        Loop
        SortList(Index2) = value
      Next
    Loop Until Distance <= 1
    
    w = GroupList(GroupNo) + GroupSize(GroupNo) - 1
    u = SortList(GroupList(GroupNo)) + Depth: Do While u > nInBytes: u = u - nInBytes: Loop
    u = OrigSort(u)
    z = 1
    For i = GroupList(GroupNo) + 1 To w
      V = SortList(i) + Depth: Do While V > nInBytes: V = V - nInBytes: Loop
      V = OrigSort(V)
      If u = V Then
        z = z + 1
      Else
        If z > 1 Then
          nNewGroups = nNewGroups + 1
          g = i - z
          NewGroupList(nNewGroups) = g
          NewGroupSize(nNewGroups) = z
          For j = g To i - 1
            GroupOrder(SortList(j)) = g: If g = 0 Then Stop
          Next j
        Else
          GroupOrder(SortList(i - 1)) = i - 1: If i - 1 = 0 Then Stop
        End If
        z = 1
      End If
      u = V
    Next i
    If z > 1 Then
      nNewGroups = nNewGroups + 1
      g = i - z
      NewGroupList(nNewGroups) = g
      NewGroupSize(nNewGroups) = z
      For j = g To i - 1
        GroupOrder(SortList(j)) = g: If g = 0 Then Stop
      Next j
    Else
      GroupOrder(SortList(i - 1)) = i - 1: If i - 1 = 0 Then Stop
    End If
    
    GroupSize(GroupNo) = 0
End Function
