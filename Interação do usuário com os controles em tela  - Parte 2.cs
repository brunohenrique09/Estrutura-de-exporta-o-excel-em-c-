'Aqui é verificado todos os controles que são do tipo LookUpEdit, 
'e se o usuário realizou alguma edição.
'Caso tenha sido feita alguma edição nos controles,
'é  criada uma nova lista, pegando uma entidade criada LayoutImportacaoExportacaoColunas,
'realizando um laço para mapear a posição que cada número irá representar na exportação para o excel,
'seguindo o padrão da planilha, iniciando na letra A e terminando na letra Z.
'Foi passado o valor "" como zero, permitindo que o usuário possa deixar colunas vazias no excel caso queira.



private void cmbLayout_EditValueChanged(object sender, EventArgs e)
    {
      try
      {
        var control = sender as LookUpEdit;

        if (control == null) return;

        if (control.EditValue == null) return;

        var layout = _lstLayoutImportacaoExportacao.FirstOrDefault(c => c.Sigla == control.EditValue.ToString());

        if (layout == null) return;

        tblSelecionarColunasExport.Controls.OfType<LookUpEdit>().ToList().ForEach(c => c.EditValue = null);

        _lstExportacaoColunaExcel = layout.LayoutImportacaoExportacaoColunas;
        _lstExportacaoColunaExcel.ForEach(a =>
        {
          {
            {
              switch (a.Posicao)
              {
                case 0:
                  RetornaValorLayout(a.NomeColuna, "");
                  break;
                case 1:
                  RetornaValorLayout(a.NomeColuna, "A");
                  break;
                case 2:
                  RetornaValorLayout(a.NomeColuna, "B");
                  break;
                case 3:
                  RetornaValorLayout(a.NomeColuna, "C");
                  break;
                case 4:
                  RetornaValorLayout(a.NomeColuna, "D");
                  break;
                case 5:
                  RetornaValorLayout(a.NomeColuna, "E");
                  break;
                case 6:
                  RetornaValorLayout(a.NomeColuna, "F");
                  break;
                case 7:
                  RetornaValorLayout(a.NomeColuna, "G");
                  break;
                case 8:
                  RetornaValorLayout(a.NomeColuna, "H");
                  break;
                case 9:
                  RetornaValorLayout(a.NomeColuna, "I");
                  break;
                case 10:
                  RetornaValorLayout(a.NomeColuna, "J");
                  break;
                case 11:
                  RetornaValorLayout(a.NomeColuna, "K");
                  break;
                case 12:
                  RetornaValorLayout(a.NomeColuna, "L");
                  break;
                case 13:
                  RetornaValorLayout(a.NomeColuna, "M");
                  break;
                case 14:
                  RetornaValorLayout(a.NomeColuna, "N");
                  break;
                case 15:
                  RetornaValorLayout(a.NomeColuna, "O");
                  break;
                case 16:
                  RetornaValorLayout(a.NomeColuna, "P");
                  break;
                case 17:
                  RetornaValorLayout(a.NomeColuna, "Q");
                  break;
                case 18:
                  RetornaValorLayout(a.NomeColuna, "R");
                  break;
                case 19:
                  RetornaValorLayout(a.NomeColuna, "S");
                  break;
                case 20:
                  RetornaValorLayout(a.NomeColuna, "T");
                  break;
                case 21:
                  RetornaValorLayout(a.NomeColuna, "U");
                  break;
                case 22:
                  RetornaValorLayout(a.NomeColuna, "V");
                  break;
                case 23:
                  RetornaValorLayout(a.NomeColuna, "W");
                  break;
                case 24:
                  RetornaValorLayout(a.NomeColuna, "X");
                  break;
                case 25:
                  RetornaValorLayout(a.NomeColuna, "Y");
                  break;
                case 26:
                  RetornaValorLayout(a.NomeColuna, "Z");
                  break;
              }
            }
          }
          return;
        });
      }
      catch (Exception ex)
      {
        MessageBoxDetail.Show("Ocorreu um erro no método cmbLayout_EditValueChanged.", ex.ToString());
      }
    }
