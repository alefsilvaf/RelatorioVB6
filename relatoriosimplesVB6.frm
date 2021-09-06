Option Explicit

Private Sub CmdOk_Click()



   Dim vlsCriterio As String
   Dim vlsTabela As String
   Dim filtro As String
   'Dim statuscartao As String
   
   vlsTabela = "minhaview"
   filtro = ""
   
   'Filtrando pela faixa etária
   If IdadeInicial.Text <> Empty Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
      vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Idade} >= " & IdadeInicial.Text & " "
      filtro = "Idade maior que: " & IdadeInicial.Text & ". "
      
   End If
   
   If IdadeFinal.Text <> Empty Then
      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
      vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Idade} <= " & IdadeFinal.Text & " "
      filtro = "Idade menor que: " & IdadeInicial.Text & ". "
      
   End If
   
   'Filtrando por data de uso do cartão
   
     If DataInicial.Text <> Empty Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
      vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Data_Compra} >=  DateTime(" & Format(DataInicial.Text & " 00:00:00", "YYYY,MM,DD,HH,mm,SS") & ")"
      filtro = "Data Inicial: " & DataInicial.Text & " "

   End If
   
   
   If DataFinal.Text <> Empty Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
      vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Data_Compra} <=  DateTime(" & Format(DataFinal.Text & " 23:59:59", "YYYY,MM,DD,HH,mm,SS") & ")"
      filtro = "Data Final: " & DataFinal.Text & " "

   End If
   
    'Filtrando por número de faturas
   If txtNFaturas.Text <> Empty Then
   
      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
        
   vlsCriterio = vlsCriterio & "{" & vlsTabela & ".N_Faturas} >= " & txtNFaturas.Text
   
   filtro = filtro & "Número de faturas maior que: " & txtNFaturas.Text & " "
   End If
   

   If chkAtivo.value = 1 Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " AND "
      End If

   vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Bloqueado} = 0 "
        filtro = filtro & "Não Bloqueados / Aguardando Desbloqueio "
   End If

   If Outros.value = 1 Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " AND "
      End If

   vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Bloqueado} = 1 "
        filtro = filtro & "Bloqueados "
   End If
   
   'Filtrando pela porcentagem disponível
   If txtPorcentagem.Text <> Empty Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
      vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Valor_Disponivel} >= ({" & vlsTabela & ".Vr_Limite} * 0." & txtPorcentagem.Text & ")"
      
      filtro = "Porcentagem: " & txtPorcentagem.Text & " "
      
   End If
   
   'Filtrando pelas opções de exclusão de casos
   If optSaque.value = 1 Then

      If Trim(vlsCriterio) <> Empty Then
         vlsCriterio = vlsCriterio & " And "
      End If
      
   vlsCriterio = vlsCriterio & "{" & vlsTabela & ".Ctrl_TipoTransacao} <> 29 and not isnull({" & vlsTabela & ".Ctrl_Caixa}) "
        filtro = filtro & "Sem saque em andamento. "
   End If
   
   FrmRelatorios.vlsSQLFILTRO = vlsCriterio
   
   FrmRelatorios.vlbCancRelatorio = False
   
   Unload Me
   Exit Sub
    
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

