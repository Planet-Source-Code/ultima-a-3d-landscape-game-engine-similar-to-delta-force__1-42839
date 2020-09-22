Attribute VB_Name = "HelperModule_Functions"
Public Function InitHelp()
    Dim V As VERTEX
    VertexSizeInBytes = Len(V)
End Function

Public Function GetCol(Col As Long) As RGBValue
    With GetCol
        .R = Col And 255
        .G = (Col And 65280) / 256&
        .B = (Col And 16711680) / 65535
    End With
End Function

Public Function MakeVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single) As VERTEX
    With MakeVertex
        .Position.X = X
        .Position.Y = Y
        .Position.Z = Z
        .tu = U
        .tv = V
    End With
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Public Function Min(A As Single, B As Single) As Single
    If A > B Then
        Min = B
    Else
        Min = A
    End If
End Function

Public Function Max(A As Single, B As Single) As Single
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function

Public Function Normalise(V As D3DVECTOR) As D3DVECTOR
    Dim Leng As Single
    With Normalise
        Leng = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)
        If Leng > 0 Then
            .X = V.X / Leng
            .Y = V.Y / Leng
            .Z = V.Z / Leng
        End If
    End With
End Function

Public Function VectorSubtract(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With VectorSubtract
        .X = V1.X - V2.X
        .Y = V1.Y - V2.Y
        .Z = V1.Z - V2.Z
    End With
End Function

Public Function CrossProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With CrossProduct
        .X = V1.Y * V2.Z - V1.Z * V2.Y
        .Y = V1.Z * V2.X - V1.X * V2.Z
        .Z = V1.X * V2.Y - V1.Y * V2.X
    End With
End Function

Public Function DotProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function

Public Function Dist(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    Dist = Sqr((V2.X - V1.X) * (V2.X - V1.X) + (V2.Y - V1.Y) * (V2.Y - V1.Y) + (V2.Z - V1.Z) * (V2.Z - V1.Z))
End Function

Public Function Acos(XX As Single) As Single
    If Abs(XX) < 1 Then
        Acos = Atn(-XX / Sqr(-XX * XX + 1)) + 1.5707963267949
    ElseIf XX = 1 Then
        Acos = 0
    ElseIf XX = -1 Then
        Acos = Pi
    End If
End Function

Public Function Asin(XX As Single) As Single
    If Abs(XX) < 1 Then
        Asin = Atn(XX / Sqr(-XX * XX + 1))
    ElseIf XX = 1 Then
        Asin = 1.5707963267948
    ElseIf XX = -1 Then
        Asin = 4.7123889803846
    End If
End Function
