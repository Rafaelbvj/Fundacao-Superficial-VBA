Type Sapata
    Rem Dimensão em X em metros
    A               As Double
    Rem Dimensão em Y em metros
    B               As Double
    Rem Altura em metros
    H               As Double
    Rem Zona de pressão (1 a 4)
    Zona            As Integer
    Rem Tensão atuante em MPa
    qmax            As Double
    Rem Excentricidades em metros
    ex              As Double
    ey              As Double
End Type

Type Pilar
    Rem Coordenadas Do centroide em metros - Entrada
    X               As Double
    Y               As Double
    Rem Dimensões Do pilar em metros - Entrada
    PX              As Double
    PY              As Double
    Rem Carregamento em tf e tf.m - Entrada
    P               As Double
    Mx              As Double
    My              As Double
    Rem Informações da sapata calculada - Saída
    Sapata          As Sapata
End Type

Rem proporçao -> verdadeiro para usar proporção da dimensão com 'fator' (fator = A / B) e falso para usar a equivalencia com os pilares (A - PX) = (B - PY)
Rem sigma_adm é a tensão admissível do solo em MPa
Rem A funcao retorna a tensão atuante com as dimensões de sapata definidas no objeto do tipo Pilar
Function dimensionar_sapata_isolada(sigma_adm As Double, ByRef pi As Pilar, proporcao As Boolean, fator As Double) As Double
    Dim A#, B#, ex#, ey#, sigma_a#, area#
    Dim com_exc     As Boolean
    Dim duplo_momento As Boolean
    Dim z_5         As Boolean
    Dim first As Boolean
    first = False
    Rem inicializa tensão atuante com um valor maior que a tensão admissivel somente para entrar no While
    sigma_a = 2 * sigma_adm
    Rem Chute inicial
    area = 1.05 * pi.P / (sigma_adm * 100)
    Rem Formula de Bhaskara
    B = ((pi.PY - pi.PX) + Sqr((pi.PX - pi.PY) ^ 2 + 4 * area)) / 2
    A = B - pi.PY + pi.PX
    Rem Calcular excentricidades
    ex = pi.My / pi.P
    ey = pi.Mx / pi.P
    pi.Sapata.ex = ex
    pi.Sapata.ey = ey
    com_exc = True
    duplo_momento = True
    Rem Verifica se existe excentricidades
    If ex = 0 And ey = 0 Then
        com_exc = False
    End If
    Do While sigma_a > sigma_adm
    
        If first Then
            Rem Incrementar dimensões
            A = A + A / 1000
            Rem **Proporção ou Equivalencia de comprimento lateral da sapata**
            If proporcao Then
            B = A / fator
            Else
                B = A - pi.PX + pi.PY
            End If
        End If
        first = True
        
        Rem Calcular a tensão atuante máxima
        If com_exc Then
            
            Rem Com excentricidade
            If ex = 0 Then
                theta = WorksheetFunction.pi() / 2
            Else
                theta = Atn(Abs(ey / ex))
            End If
            
            Rem Região 2
            raio_max = Sqr(1 / ((Cos(theta) / (A / 3)) ^ 2 + (Sin(theta) / (B / 3)) ^ 2))
            raio = Sqr(ey ^ 2 + ex ^ 2)
            If raio > raio_max Then
                Debug.Print "Excentricidade localizada em região inaceitavel - Incrementa-se as dimensões."
            Else
                
                Rem Momento em uma direção
                If ex = 0 And Abs(ey) > (B / 6) Then
                    sigma_a = (4 / 3) * pi.P / (A * (B - 2 * Abs(ey))) / 100
                    duplo_momento = False
                End If
                If ey = 0 And Abs(ex) > (A / 6) Then
                    sigma_a = (4 / 3) * pi.P / (B * (A - 2 * Abs(ex))) / 100
                    duplo_momento = False
                End If
                
                Rem Momento nas duas direções
                If duplo_momento Then
                    z_5 = True
                    
                    Rem Região 1
                    If (Abs(ex / A) + Abs(ey / B)) <= (1 / 6) Then
                        sigma_a = (pi.P / (A * B)) * (1 + 6 * Abs(ex) / A + 6 * Abs(ey) / B) / 100
                        z_5 = False
                        pi.Sapata.Zona = 1
                    End If
                    
                    Rem Região 3
                    pos = (Abs(ex) - A / 6) * ((12 * B) / (A * 4))
                    If Abs(ex) > A / 6 And Abs(ey) < pos Then
                        s = (B / 12) * ((B / Abs(ey)) + Sqr((B / Abs(ey)) ^ 2 - 12))
                        t_alpha = (3 / 2) * (A - 2 * Abs(ex)) / (s + Abs(ey))
                        sigma_a = (12 * pi.P / (B * t_alpha)) * (B + 2 * s) / (B ^ 2 + 12 * (s ^ 2)) / 100
                        pi.Sapata.Zona = 3
                        z_5 = False
                    End If
                    
                    Rem Região 4
                    pos = (Abs(ey) - B / 6) * ((12 * A) / (B * 4))
                    If Abs(ey) > B / 6 And Abs(ex) < pos Then
                        s = (A / 12) * ((A / Abs(ex)) + Sqr((A / Abs(ex)) ^ 2 - 12))
                        t_alpha = (3 / 2) * (B - 2 * Abs(ey)) / (s + Abs(ex))
                        sigma_a = (12 * pi.P / (A * t_alpha)) * (A + 2 * s) / (A ^ 2 + 12 * (s ^ 2)) / 100
                        pi.Sapata.Zona = 4
                        z_5 = False
                    End If
                    
                    Rem Região 5
                    If z_5 Then
                        alph = Abs(ex) / A + Abs(ey) / B
                        sigma_a = (pi.P * alph / (A * B)) * (12 - 3.9 * (6 * alph - 1) * (1 - 2 * alph) * (2.3 - 2 * alph)) / 100
                        pi.Sapata.Zona = 5
                    End If
                    
                End If
                duplo_momento = True
            End If
        Else
            Rem Sem excentricidade
            sigma_a = (pi.P / (A * B)) / 100
        End If

    Loop
    pi.Sapata.qmax = sigma_a
    pi.Sapata.A = A
    pi.Sapata.B = B
    Rem H >= (A - PX)/3 -> Sapata Rígida
    Rem H < (A - PX)/3 -> Sapata Flexível
    pi.Sapata.H = (A - pi.PX) / 3
    dimensionar_sapata_isolada = sigma_a
End Function

Rem sigma_adm em MPa
Function dimensionar_sapata_associada(ByRef P1 As Pilar, ByRef P2 As Pilar, sigma_adm As Double) As Double
    Dim X_rel       As Double
    Dim Y_rel       As Double
    Dim A#, B#
    Dim theta       As Double
    Dim D           As Pilar
    
    X_rel = P2.X - P1.X
    Y_rel = P2.Y - P1.X
    Rem Angulo em radianos da reta que liga os centroides
    If (X_rel = 0) Then
        theta = WorksheetFunction.pi() / 2
    Else
        theta = Atn(Y_rel / X_rel)
    End If
    
    Rem centro de carga
    Dim Xc          As Double
    Dim Yc          As Double
    Xc = X_rel * P2.P / (P2.P + P1.P)
    Yc = Y_rel * P2.P / (P2.P + P1.P)
    Rem Esforços resultantes
    Dim Pc#, Mxc#, Myc#, Myt#, Mxt#
    Mxt = P1.Mx + P2.Mx
    Myt = P1.My + P2.My
    
    Dim pilar_ficticio As Pilar
    pilar_ficticio.P = P1.P + P2.P
    pilar_ficticio.Mx = Mxt * Cos(theta) + Myt * Sin(theta)
    pilar_ficticio.My = -Mxt * Sin(theta) + Myt * Cos(theta)
    
    Rem Geometria pilar fictício
    Dim area#: area = P1.PX * P1.PY + P2.PX * P2.PY
    pilar_ficticio.PY = pilar_ficticio.PX = Sqr(area)
    
    Err = dimensionar_sapata_isolada(sigma_adm, pilar_ficticio, True, 2.2)
    
    P1.Sapata = pilar_ficticio.Sapata
    P2.Sapata = pilar_ficticio.Sapata
    Rem Tem que subtrair a altura da sapata pela largura da Viga de rigidez/3
    P1.Sapata.H = P1.Sapata.A / 3
    P2.Sapata.H = P2.Sapata.A / 3
    dimensionar_sapata_associada = Err
    
End Function

    Rem K é a excentricidade do pilar na sapata de divisa (distância entre o centroide do pilar e o centroide da sapata)
Function dimensionar_sapata_divisa(K As Double, ByRef P1_divisa As Pilar, ByRef P2 As Pilar, sigma_adm As Double)
    Dim X_rel#, Y_rel#, U#
    Dim theta       As Double
    
    X_rel = P2.X - P1_divisa.X
    Y_rel = P2.Y - P1_divisa.X
    Rem Angulo em radianos da reta que liga os centroides dos pilares
    If (X_rel = 0) Then
        theta = WorksheetFunction.pi() / 2
    Else
        theta = Atn(Y_rel / X_rel)
    End If
    
    U = Sqr(X_rel ^ 2 + Y_rel ^ 2)
    
    Dim M_perp#, Pc#
    
    Rem Calcula momento perpendicular ao eixo da viga de equilíbrio
    M_perp = P1_divisa.Mx * Sin(theta) + P1_divisa.My * Cos(theta)
    
    Rem Corrige o P (carga vertical) com a excentricidade K
    Pc = P1_divisa.P * U / (U - K) + M_perp / (U - K)
    P1_divisa.P = Pc
    
    Rem Dimensiona a sapata da divisa com o P corrigido
    sigma_a = dimensionar_sapata_isolada(sigma_adm, P1_divisa, True, 2.5)
    
    Pc = P1_divisa.P * (1 - U / (U - K)) - M_perp / (U - K)
    
    If Pc > 0 Then
        Rem Sobrecarga
        P2.P = P2.P + Pc
    Else
        Rem Só pode considerar 50% Do alívio - norma, se existir
        P2.P = P2.P + 0.5 * Pc
    End If
    
    sigma_a = dimensionar_sapata_isolada(sigma_adm, P2, True, 1)
    
    dimensionar_sapata_divisa = sigma_a
End Function

Rem **************** Autor: Rafael Costa - Aluno universitário da UNB ********************************
