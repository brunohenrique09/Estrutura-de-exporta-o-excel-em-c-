'''Dicionários criados para exportação.
'''Aqui foram criados os dicionários com o nome de cada combo referente aos campos que serão utilizados
'''para retornar os dados conforme seleção do usuário.
'''O dicionario de posição, faz referencia as colunas do excel, iniciando em A e terminando em Z.

 

 /// <summary>
        /// Dicionário para o combo NomeColuna do Formulario FormCadastroLayoutImportacaoExportacao.
        /// </summary>
        public static readonly Dictionary<int, string> DicNomeColunaLayoutImportacaoExportacao = new Dictionary<int, string>
        {
            {1,  "Lote" },
            {2,  "Grupo" },
            {3,  "Item" },
            {4,  "Descrição" },
            {5,  "Marca" },
            {6,  "Modelo" },
            {7,  "Marca / Modelo" },
            {8,  "Modelo / Versão" },
            {9,  "Marca / Fabricante" },
            {10, "Estimado" },
            {11, "Procedência" },
            {12, "Fabricante" },
            {13, "Tipo" },
            {14, "Embalagem" },
            {15, "Registro de Produto" },
            {16, "Quantidade" },
            {17, "Unitário" },
            {18, "Total" },
            {19, "Unidade" },
        };

        /// <summary>
        /// Dicionário para o combo Posicao do Formulario FormCadastroLayoutImportacaoExportacao.
        /// </summary>
        public static readonly Dictionary<int, string> DicPosicaoLayoutImportacaoExportacao = new Dictionary<int, string>
        {
            {0,  string.Empty },
            {1,  "A" },
            {2,  "B" },
            {3,  "C" },
            {4,  "D" },
            {5,  "E" },
            {6,  "F" },
            {7,  "G" },
            {8,  "H" },
            {9,  "I" },
            {10, "J" },
            {11, "K" },
            {12, "L" },
            {13, "M" },
            {14, "N" },
            {15, "O" },
            {16, "P" },
            {17, "Q" },
            {18, "R" },
            {19, "S" },
            {20, "T" },
            {21, "U" },
            {22, "V" },
            {23, "W" },
            {24, "X" },
            {25, "Y" },
            {26, "Z" },
        };
