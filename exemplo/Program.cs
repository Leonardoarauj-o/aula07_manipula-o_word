using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo {
    class Program {
        static void Main (string[] args) {
            Document exemploDoc = new Document ();
            #region criacao de documento
            //cria um documento com o nome exemploDoc//
            Section secaoCapa = exemploDoc.AddSection ();

            #endregion

            #region criar paragrafo
            //criar um paragrafo com o nome titulo e adiciona a secao secaoCap//
            //os paragrafos sao necessarios para inserçao de texto, imagens, tabelas, etc//
            Paragraph titulo = secaoCapa.AddParagraph ();
            #endregion

            #region adicionar um texto ao paragrafo
            //Adicionar o texto exemplos de titulos ao paragrafo titulo
            titulo.AppendText ("Exemolo de Titulo\n\n");

            #endregion

            #region formatar paragrafo
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            ParagraphStyle estilo01 = new ParagraphStyle (exemploDoc);

            //Adiciona um nome ao estilo01//
            estilo01.Name = "cor do Titulo";

            //definir cor do titulo
            estilo01.CharacterFormat.TextColor = Color.DarkRed;

            //define que o texto será em negrito
            estilo01.CharacterFormat.Bold = true;

            //adiciona o estilo01 ao documento exemploDoc
            exemploDoc.Styles.Add (estilo01);

            //aplica o estilo01 ao paragrafo titulo
            titulo.ApplyStyle (estilo01.Name);
            #endregion

            #region Trabalhar com Tabulação

            //adiciona um paragrafo textoCapa a seção secaoCapa
            Paragraph TextoCapa = secaoCapa.AddParagraph ();

            //adiciona um testo ao paragrafo com tabulação
            TextoCapa.AppendText ("\tEate é um exemplo de texto com t.");

            //adicionar um novo paragrafo a mesma seção (secaoCapa)
            Paragraph textoCapa2 = secaoCapa.AddParagraph ();

            //adicionar um texto ao paragrafo textocapa2 com concatenação
            textoCapa2.AppendText ("\tBasicamente, então, uma seção representa uma pagina do documento e os paragrafos dentro de uma seção " + " obviamente, aparecem na mesma pagina");

            #endregion 

            #region inserir imagens
            Paragraph imagemCapa = secaoCapa.AddParagraph ();
            //adiciona um texto ao paragrafo imagemCapa
            imagemCapa.AppendText ("\n\n\tAgora vamos enserir uma imagem ao documento\n\n");

            imagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

            DocPicture imagemExemplo = imagemCapa.AppendPicture(Image.FromFile (@"saida\imglogo_csharp.png"));
                
            //define uma largura e uma altura para a imagem 
            imagemExemplo.Width = 300;
            imagemExemplo.Height = 300;
            #endregion

            #region Adicionar nova seção
            //Adicionar uma nova seção
            Section secaoCorpo = exemploDoc.AddSection ();

            Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph ();

            paragrafoCorpo1.AppendText ("\tEste é um exemplo de paragrafo criado em uma seção, " + " \tComo foio criada uma nova seção, perceba que este texto aparece em uma nova pagina,");

            #endregion

            #region Adicionar uma tabela
            //Adicionar o cabeçalho da tabela
            Table tabela = secaoCorpo.AddTable (true);

            // cria o cabeçalho da tabela
            string[] cabeçalho = { "Item", "Descrição", "Qtd", "Preço Unit.", "Preço" };

            String[][] dados = {
                new String[] { "Cenoura", "Vegetal muito nutritivo", "1", "R$ 4,00", "R$ 4,00" },
                new String[] { "Batata", "Vegetal muito nutritivo", "2", "R$ 5,00", "R$ 10,00" },
                new String[] { "Alface", "Vegetal muito nutritivo", "1", "R$ 1,50", "R$ 1,50" },
                new String[] { "Tomate", "Vegetal muito nutritivo", "1", "R$ 12,00", "R$ 12,00" }
            };

            //Adiciona as celulas na tabela
            tabela.ResetCells (dados.Length + 1, cabeçalho.Length);

            //adiciona uma linha na posição [0] do vetor de linhas
            //e define que esta linha é o cabeçalho
            TableRow Linha1 = tabela.Rows[0];
            Linha1.IsHeader = true;

            //define a altura da linha 
            Linha1.Height = 23;

            //Formatação do cabeçalho
            Linha1.RowFormat.BackColor = Color.AliceBlue;

            //percorre as colunas do cabeçalho
            for (int i = 0; i < cabeçalho.Length; i++) {
                Paragraph p = Linha1.Cells[i].AddParagraph ();
                Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //Formatação dos dados do cabeçalho
                TextRange TR = p.AppendText (cabeçalho[i]);
                TR.CharacterFormat.FontName = (cabeçalho[i]);
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.Teal;
                TR.CharacterFormat.Bold = true;

                //adiciona as linhas do corpo da tabela
                for (int r = 0; r < dados.Length; r++) {
                    TableRow LinhaDados = tabela.Rows[r + 1];

                    //define a altura da linha 
                    LinhaDados.Height = 20;

                    for (int c = 0; c < dados[r].Length; c++) {
                        //alinha as celulas
                        LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                        //preenche os dados nas linhas
                        Paragraph p2 = LinhaDados.Cells[c].AddParagraph ();

                        TextRange TR2 = p2.AppendText (dados[r][c]);

                        //formata as celulas 
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR2.CharacterFormat.FontName = "calibri";
                        TR2.CharacterFormat.FontSize = 12;
                        TR2.CharacterFormat.TextColor = Color.Brown;
                    }
                }

                #endregion

                #region salvar arquivo
                
                //salvar um arquivo em .Docx
                //assim como np word, caso ja exista um arquivo com este nome, é substituida
                //ultilozar o metodo SaveToFile para salvar o arquivo no formato desejado 
                exemploDoc.SaveToFile (@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);

            }
        }
    }