     'Aqui inicia a montagem dos dados que serão exportados pras colunas que foram mapeadas pela seleção do usuário anteriormente.
     'É utilizada a entidade  colunasExcel, que é mapeada especificamente pra ser utilizada pro método de exportação
     'CreateExcel();		 
   
        'pra exemplificar a entidade, coloquei uma propriedade como exemplo.
		
		 /// <summary>
        /// Representa o campo A.
        /// </summary>
        [DataMember]
        [DataField("A")]
        [DbParameter(Ignore = true)]
        [JsonProperty("A")]
        public string A { get; set; }
		
		
		'A função RetornValueColum, vai nos auxiliar nos dados que serão exportados pro excel posteriormente.
		
		
		
 'OBS: Antes de iniciar o trecho a baixo, tem toda a parte de validação  dos metodos anteriores que não coloquei pra
 'não ficar ainda mais extenso o código.
 
 _lstExportacaoColunaMovimento.ForEach(d =>
        {
          var colunasExcel = new ExcelColunas();
          lstExportacao.ForEach(c =>
          {
            switch (c.Letra)
            {
              case "A":
                colunasExcel.A = RetornValueColum(d, c.Valor);
                break;
              case "B":
                colunasExcel.B = RetornValueColum(d, c.Valor);
                break;
              case "C":
                colunasExcel.C = RetornValueColum(d, c.Valor);
                break;
              case "D":
                colunasExcel.D = RetornValueColum(d, c.Valor);
                break;
              case "E":
                colunasExcel.E = RetornValueColum(d, c.Valor);
                break;
              case "F":
                colunasExcel.F = RetornValueColum(d, c.Valor);
                break;
              case "G":
                colunasExcel.G = RetornValueColum(d, c.Valor);
                break;
              case "H":
                colunasExcel.H = RetornValueColum(d, c.Valor);
                break;
              case "I":
                colunasExcel.I = RetornValueColum(d, c.Valor);
                break;
              case "J":
                colunasExcel.J = RetornValueColum(d, c.Valor);
                break;
              case "K":
                colunasExcel.K = RetornValueColum(d, c.Valor);
                break;
              case "L":
                colunasExcel.L = RetornValueColum(d, c.Valor);
                break;
              case "M":
                colunasExcel.M = RetornValueColum(d, c.Valor);
                break;
              case "N":
                colunasExcel.N = RetornValueColum(d, c.Valor);
                break;
              case "O":
                colunasExcel.O = RetornValueColum(d, c.Valor);
                break;
              case "P":
                colunasExcel.P = RetornValueColum(d, c.Valor);
                break;
              case "Q":
                colunasExcel.Q = RetornValueColum(d, c.Valor);
                break;
              case "R":
                colunasExcel.R = RetornValueColum(d, c.Valor);
                break;
              case "S":
                colunasExcel.S = RetornValueColum(d, c.Valor);
                break;
              case "T":
                colunasExcel.T = RetornValueColum(d, c.Valor);
                break;
              case "U":
                colunasExcel.U = RetornValueColum(d, c.Valor);
                break;
              case "V":
                colunasExcel.V = RetornValueColum(d, c.Valor);
                break;
              case "W":
                colunasExcel.W = RetornValueColum(d, c.Valor);
                break;
              case "X":
                colunasExcel.X = RetornValueColum(d, c.Valor);
                break;
              case "Y":
                colunasExcel.Y = RetornValueColum(d, c.Valor);
                break;
              case "Z":
                colunasExcel.Z = RetornValueColum(d, c.Valor);
                break;
            }
          });
          _lstColunasExcel.Add(colunasExcel);
        });
