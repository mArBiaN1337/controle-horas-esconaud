Attribute VB_Name = "Module1"
Function adicionalNoturno_1Intervalo(hrEntrada As Double, hrTermino As Double)

Dim minEntrada As Integer
Dim minTermino As Integer
Dim minCount As Integer
Dim minNoturno As Integer

minEntrada = hrEntrada * 1440
minTermino = hrTermino * 1440

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

Function adicionalNoturno_2Intervalos(hrEntrada As Double, hrSaida As Double, hrRetorno As Double, hrTermino As Double)

Dim minEntrada As Integer
Dim minTermino As Integer
Dim minCount As Integer
Dim minNoturnoProg As Integer
Dim hrNoturnoProg As Double

minEntrada = hrEntrada * 1440
minTermino = hrTermino * 1440

    For minCount = minEntrada To minTermino - 1
    
        If minEntrada < 1320 And minCount > 1739 Then
        
            minNoturnoProg = minNoturnoProg + 1
        
        End If
    
    Next minCount
    
    hrNoturnoProg = minNoturnoProg / 1440
    
    adicionalNoturno4 = hrNoturnoProg + adicionalNoturno_1Intervalo(hrEntrada, hrSaida) + adicionalNoturno_1Intervalo(hrRetorno, hrTermino)

End Function


