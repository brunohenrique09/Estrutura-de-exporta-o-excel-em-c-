'''Aqui foi feita uma função que sera utilizada no proximo método.

'O Swith Case  da continuidade ao processo de montagem da planilha.
'Os valores dos cases fazem referencia ao dicionário criado anteriormente.
'Em cada case, pegamos a letra que o usuário escreveu atraves da seleção do combo.
'CmbGrupo.Text = Letra
'Atribuimos o valor da letra que o usuário selecionou no combo, e o case faz referência a qual campo o usuário selecionou no passo anterior (parte 2)
'Assim sabemos qual é o combo, e definimos a letra que será a posição dentro do excel.





private string RetornaValorLayout(int? nomecoluna, string letra)
    {
      try
      {
        switch (nomecoluna)
        {
          case 2:
            cmbGrupo.Text = letra;
            break;
          case 3:
            cmbItem.Text = letra;
            break;
          case 4:
            cmbDescricao.Text = letra;
            break;
          case 7:
            cmbMarcaModelo.Text = letra;
            break;
          case 11:
            cmbProcedencia.Text = letra;
            break;
          case 12:
            cmbFabricante.Text = letra;
            break;
          case 13:
            cmbTipo.Text = letra;
            break;
          case 14:
            cmbEmbalagem.Text = letra;
            break;
          case 15:
            cmbRegistroProduto.Text = letra;
            break;
          case 16:
            cmbQuantidade.Text = letra;
            break;
          case 17:
            cmbValorUnitario.Text = letra;
            break;
          case 18:
            cmbTotal.Text = letra;
            break;
          case 19:
            cmbUnidade.Text = letra;
            break;
        }
      }
      catch (Exception ex)
      {
        MessageBoxDetail.Show("Ocorreu um erro no método RetornaValorLayout.", ex.ToString());
      }
      fechacombo();
      return string.Empty;
    }
