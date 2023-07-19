'Essa é a função que é utilizada na parte anterior.  Ela é responsável por retornadas os dados da entidade
'PropostaEntradaExportacaoMovimento. Após passar todos os dados, observe na parte parte final do projeto,
'nossa lista adiciona todos os valores referenciados na entidade colunasExcel, e estará pronta pra ser exportada.
'>> _lstColunasExcel.Add(colunasExcel);  <<
  
  public string RetornValueColum(PropostaEntradaExportacaoMovimento dados, string nomeCombo)
    {
      try
      {
        switch (nomeCombo)
        {
          case "cmbGrupo":
            return dados.LoteAxeno?.ToString();

          case "cmbItem":
            return dados.Ordem.ToString();

          case "cmbDescricao":
            return dados.DescricaoProduto.ToString();

          case "cmbMarcaModelo":
            return dados.Marca?.ToString();

          case "cmbProcedencia":
            return dados.Procedencia.ToString();

          case "cmbFabricante":
            return dados.Fabricante?.ToString();

          case "cmbTipo":
            return dados.TipoProduto?.ToString();

          case "cmbEmbalagem":
            return dados.Unidade.ToString() + ' ' + dados.Conversao.ToString();

          case "cmbRegistroProduto":
            return dados.ValidadeRegistro?.ToString();

          case "cmbQuantidade":
            return dados.Quantidade.ToString();

          case "cmbValorUnitario":
            return dados.ValorUnitario.ToString();

          case "cmbTotal":
            return dados.TotalItem.ToString();

          case "cmbUnidade":
            return dados.UnidadeUsada?.ToString();
        }
      }
      catch (Exception ex)
      {
        MessageBoxDetail.Show("Ocorreu um erro no método RetornValueColum.", ex.ToString());
      }
      return string.Empty;
    }
