Attribute VB_Name = "Module1"
Function adicionalNoturno(hrEntrada As Double, hrTermino As Double)

Dim minEntrada As Integer
Dim minTermino As Integer
Dim minCount As Integer
Dim minNoturno As Integer

minEntrada = hrEntrada * 24 * 60
minTermino = hrTermino * 24 * 60

If hrEntrada <> Empty And hrTermino <> Empty Then

    For minCount = minEntrada To minTermino - 1
    
        If minCount < 300 Or minCount > 1319 And minCount < 1740 Then
        
            minNoturno = minNoturno + 1
        
        End If
    
    Next minCount
    
    adicionalNoturno = minNoturno / 1440

Else
    
    adicionalNoturno = 0

End If

End Function

Function adicionalNoturnoCompleto(hrEntrada As Double, hrSaida As Double, hrRetorno As Double, hrTermino As Double)

Dim minEntrada As Integer
Dim minTermino As Integer
Dim minCount As Integer
Dim minNoturnoProg As Integer
Dim hrNoturnoProg As Double

If hrEntrada <> Empty And hrSaida <> Empty And hrRetorno <> Empty And hrTermino <> Empty Then

    minEntrada = hrEntrada * 24 * 60
    minTermino = hrTermino * 24 * 60

    For minCount = minEntrada To minTermino - 1
    
        If minEntrada < 1320 And minCount > 1739 Then
        
            minNoturnoProg = minNoturnoProg + 1
        
        End If
    
    Next minCount
        
    hrNoturnoProg = minNoturnoProg / 1440
        
    adicionalNoturnoCompleto = hrNoturnoProg + adicionalNoturno(hrEntrada, hrSaida) + adicionalNoturno(hrRetorno, hrTermino)

Else
    adicionalNoturnoCompleto = 0
End If


End Function

Function calcExtras(hrEntrada As Double, hrSaida As Double, hrRetorno As Double, hrTermino As Double)


End Function
