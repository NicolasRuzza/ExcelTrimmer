using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExtratorDeConteudo.Class;
using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ExtratorDeConteudo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTable excelData;
        private List<RegexField> regras = new List<RegexField>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnCarregar_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Filter = "Arquivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
            };

            if (openFile.ShowDialog() == true)
            {
                try
                {
                    excelData = LerExcel(openFile.FileName);
                    dataGridOriginal.ItemsSource = excelData.DefaultView;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao carregar Excel: " + ex.Message);
                }
            }
        }

        private DataTable LerExcel(string path)
        {
            var dt = new DataTable();

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                bool primeiraLinha = true;

                // Otimizar busca e definição de sufixo para os nomes de cabeçalhos
                Dictionary<string, int> controleSufixo = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> headersUsados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Regex para descobrir a "Raiz" do nome
                Regex regexRaiz = new Regex(@"^(.*)_(\d+)$");

                foreach (var row in ws.RowsUsed())
                {
                    if (primeiraLinha)
                    {
                        foreach (var cell in row.Cells())
                        {
                            string header = cell.Value.ToString();
                            // Escapa caracteres que quebram o WPF
                            header = header.Trim()
                                .Replace("/", " ou ")
                                .Replace(".", " ");

                            if (string.IsNullOrEmpty(header))
                                header = "ColunaSemNome";

                            string headerFinal = header;
                            string baseHeader = header;

                            if (headersUsados.Contains(baseHeader))
                            {
                                Match match = regexRaiz.Match(header);
                                if (match.Success)
                                {
                                    baseHeader = match.Groups[1].Value; // Pega "quarto" de "quarto_101"
                                }

                                if (!controleSufixo.ContainsKey(baseHeader))
                                    controleSufixo.Add(baseHeader, 1);

                                // Caso exista uma ou mais colunas sufixadas pelo autor do excel, o código
                                // irá percorrer o loop até encontrar uma combinação ainda não utilizada
                                do
                                {
                                    headerFinal = $"{baseHeader}_{controleSufixo[baseHeader]}";
                                    controleSufixo[baseHeader]++;
                                } while (headersUsados.Contains(headerFinal));
                            }

                            headersUsados.Add(headerFinal);
                            dt.Columns.Add(headerFinal);
                        }

                        primeiraLinha = false;
                    }
                    else
                    {
                        // O datatable é construído com base na quantidade de headers. Se, por ventura, o
                        // usuário tiver preenchido uma célula perdida além das colunas de cabeçalho, irá
                        // disparar erro.

                        // O foreach foi descontinuado, pois ignorava células vazias dentro do escopo do cabeçalho.

                        // Em vez de iterar nas células da linha (que pula vazios), agora itera-se pelas colunas
                        // esperadas do DataTable.
                        dt.Rows.Add();
                        DataRow dr = dt.Rows[dt.Rows.Count - 1];
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            // O ClosedXML usa índice base 1, enquanto o DataTable usa base 0.
                            // row.Cell(i + 1) pega a célula naquela posição exata, mesmo se estiver vazia.
                            var cell = row.Cell(i + 1);

                            dr[i] = cell.Value.ToString().Trim();
                        }
                    }
                }
            }

            return dt;
        }

        private void BtnAdicionarRegra_Click(object sender, RoutedEventArgs e)
        {
            StackPanel linha = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 5) };

            var colunaBox = new TextBox { Width = 60, Margin = new Thickness(0, 0, 10, 0) };
            var filtroBox = new TextBox { Width = 180, Margin = new Thickness(0, 0, 10, 0) };
            var replaceBox = new TextBox { Width = 180, Margin = new Thickness(0, 0, 10, 0) };
            var substitutoBox = new TextBox { Width = 120, Margin = new Thickness(0, 0, 10, 0) };
            var btnRemover = new Button { Content = "❌", Width = 30 };

            // Eventos automáticos
            colunaBox.KeyUp += (s, e) => AplicarRegras();
            filtroBox.KeyUp += (s, e) => AplicarRegras();
            replaceBox.KeyUp += (s, e) => AplicarRegras();
            substitutoBox.KeyUp += (s, e) => AplicarRegras();

            var regra = new RegexField
            (
                colunaBox,
                filtroBox,
                replaceBox,
                substitutoBox,
                linha
            );

            btnRemover.Click += (s, e) =>
            {
                stackRegraCampos.Children.Remove(linha);
                regras.Remove(regra);
                AplicarRegras();
            };

            linha.Children.Add(colunaBox);
            linha.Children.Add(filtroBox);
            linha.Children.Add(replaceBox);
            linha.Children.Add(substitutoBox);
            linha.Children.Add(btnRemover);

            stackRegraCampos.Children.Add(linha);
            regras.Add(regra);
        }

        private void AplicarRegras()
        {
            if (excelData == null || excelData.Rows.Count == 0)
                return;

            try
            {
                // Copia os dados do excel para não sobreescrever o original
                var tabelaTemp = excelData.Copy();

                var linhas = tabelaTemp.AsEnumerable().ToList();

                var colunasSelecionadas = new HashSet<int>();

                // Aplicar regras sequencialmente
                foreach (var regra in regras)
                {
                    if (!int.TryParse(regra.TxtColumn.Text, out int col))
                        continue;

                    var columnCount = excelData.Columns.Count;
                    if (col < 0 || col > columnCount - 1)
                        throw new Exception($"Índice de coluna inválido! A planilha tem {columnCount} colunas. Use de 0 a {columnCount - 1}.");

                    colunasSelecionadas.Add(col);

                    string filtro = regra.TxtRegexFilter.Text?.Trim();
                    string replace = regra.TxtRegexReplace.Text?.Trim();
                    string substituto = regra.TxtOverride.Text ?? "";

                    if (!string.IsNullOrEmpty(filtro))
                    {
                        var rFiltro = new Regex(filtro);
                        linhas = linhas.Where(row =>
                            rFiltro.IsMatch(row[col]?.ToString() ?? "")
                        ).ToList();
                    }

                    if (!string.IsNullOrEmpty(replace))
                    {
                        var rReplace = new Regex(replace);
                        foreach (var row in linhas)
                        {
                            string valor = row[col]?.ToString() ?? "";
                            row[col] = rReplace.Replace(valor, substituto).Trim();
                        }
                    }
                }

                var resultado = new DataTable();
                foreach (var i in colunasSelecionadas)
                    resultado.Columns.Add(excelData.Columns[i].ColumnName);

                foreach (var linha in linhas)
                {
                    var nova = resultado.NewRow();
                    int j = 0;
                    foreach (var i in colunasSelecionadas)
                        nova[j++] = linha[i];

                    resultado.Rows.Add(nova);
                }

                dataGridResultado.ItemsSource = resultado.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao aplicar regras: " + ex.Message);
            }
        }

        private void BtnExportarTxt_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridResultado.ItemsSource == null)
            {
                MessageBox.Show("Nada a exportar.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Arquivo de Texto|*.txt",
                FileName = "resultado_formatado.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    var view = dataGridResultado.ItemsSource as DataView;
                    if (view == null || view.Count == 0)
                    {
                        MessageBox.Show("Nada a exportar.");
                        return;
                    }

                    var dt = view.ToTable();
                    var linhas = new List<string>();

                    foreach (DataRow row in dt.Rows)
                    {
                        var valores = row.ItemArray.Select(cell => $"'{cell.ToString().Replace("'", "''")}'");
                        string linhaFormatada = $"({string.Join(", ", valores)}),";
                        linhas.Add(linhaFormatada);
                    }

                    File.WriteAllLines(saveFileDialog.FileName, linhas);
                    MessageBox.Show("Arquivo exportado com êxito!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao exportar: " + ex.Message);
                }
            }
        }

        private void BtnExportarCSV_Click(object sender, RoutedEventArgs e)
        {
            var salvar = new SaveFileDialog
            {
                Filter = "Arquivo CSV (*.csv)|*.csv",
                FileName = "resultado.csv"
            };

            if (salvar.ShowDialog() == true)
            {
                var view = dataGridResultado.ItemsSource as DataView;
                if (view == null || view.Count == 0)
                {
                    MessageBox.Show("Nada a exportar.");
                    return;
                }

                var dt = view.ToTable();
                var linhas = new List<string>();
                ExportarComoCSV(dt, salvar.FileName);
                MessageBox.Show("Exportado com sucesso!");
            }
        }

        private void ExportarComoCSV(DataTable tabela, string caminhoArquivo)
        {
            var linhas = new List<string>();

            // Cabeçalho
            var colunas = tabela.Columns.Cast<DataColumn>().Select(c => c.ColumnName);
            linhas.Add(string.Join(",", colunas));

            // Linhas
            foreach (DataRow row in tabela.Rows)
            {
                var valores = row.ItemArray.Select(cell =>
                    cell == null ? "" : "\"" + cell.ToString().Replace("\"", "\"\"") + "\""
                );
                linhas.Add(string.Join(",", valores));
            }

            File.WriteAllLines(caminhoArquivo, linhas, Encoding.UTF8);
        }

        private void BtnExportarExcel_Click(object sender, RoutedEventArgs e)
        {
            var salvar = new SaveFileDialog
            {
                Filter = "Arquivo Excel (*.xlsx)|*.xlsx",
                FileName = "resultado.xlsx"
            };

            if (salvar.ShowDialog() == true)
            {
                var view = dataGridResultado.ItemsSource as DataView;
                if (view == null || view.Count == 0)
                {
                    MessageBox.Show("Nada a exportar.");
                    return;
                }

                var dt = view.ToTable();
                var linhas = new List<string>();
                ExportarComoExcel(dt, salvar.FileName);
                MessageBox.Show("Exportado como Excel com êxito!");
            }
        }

        private void ExportarComoExcel(DataTable tabela, string caminhoArquivo)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Resultado");
                ws.Cell(1, 1).InsertTable(tabela);
                workbook.SaveAs(caminhoArquivo);
            }
        }

        private void BtnLimpar_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Tem certeza que deseja limpar todos os dados da tela?",
                                         "Limpar Aplicação",
                                         MessageBoxButton.YesNo,
                                         MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                // Limpa as grids
                dataGridOriginal.ItemsSource = null;
                dataGridResultado.ItemsSource = null;

                // Limpa stack de regras
                stackRegraCampos.Children.Clear();

                // Limpa as variáveis de maior escopo
                if (excelData != null) excelData.Clear();
                regras.Clear();

                // C# garbage collector
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}