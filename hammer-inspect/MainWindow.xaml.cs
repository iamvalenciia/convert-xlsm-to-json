using ExcelDataReader;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Documents; // Añadido para el soporte de Paragraph, Run y texto enriquecido
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;

namespace hammer_inspect
{
    // ── Data model ────────────────────────────────────────────────────────────
    public class ResultRow : INotifyPropertyChanged
    {
        public string SheetName { get; set; } = string.Empty;
        public string RowLabel { get; set; } = string.Empty;
        public string MatchSummary { get; set; } = string.Empty;

        // Raw row data — JSON is built lazily only when copying
        public Dictionary<string, object?> RowData { get; set; } = new();

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    // ── Main window ───────────────────────────────────────────────────────────
    public sealed partial class MainWindow : Window
    {
        private string _selectedPath = string.Empty;

        // Full parsed workbook kept in memory
        private Dictionary<string, List<Dictionary<string, object?>>> _workbook = new();

        // Active sheet filter — null means "all sheets"
        private string? _activeSheetFilter = null;

        // Guard flag to prevent pill Checked/Unchecked loops
        private bool _updatingPills = false;

        // Observable collection bound to the ListView
        private readonly ObservableCollection<ResultRow> _results = new();

        // Pills keyed by sheet name (__ALL__ for the "All sheets" pill)
        private readonly Dictionary<string, ToggleButton> _pills = new();

        // Debounce for search text changes
        private CancellationTokenSource? _searchCts;

        // Sheets where header is NOT on row 0
        private static readonly Dictionary<string, int> HeaderSkipRows = new()
        {
            { "Group Definitions", 3 }
        };

        // ── Init ─────────────────────────────────────────────────────────────
        public MainWindow()
        {
            this.InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            LvResults.ItemsSource = _results;
        }

        // ── Select file ───────────────────────────────────────────────────────
        private async void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            WinRT.Interop.InitializeWithWindow.Initialize(
                picker, WinRT.Interop.WindowNative.GetWindowHandle(this));

            picker.ViewMode = PickerViewMode.List;
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            picker.FileTypeFilter.Add(".xlsm");

            var file = await picker.PickSingleFileAsync();
            if (file is null) return;

            _selectedPath = file.Path ?? string.Empty;
            TxtFilePath.Text = file.Name ?? "Unnamed file";
            BtnConvert.IsEnabled = true;
            TxtStatus.Text = "File loaded — ready to convert.";
        }

        // ── Convert ───────────────────────────────────────────────────────────
        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedPath)) return;

            BtnConvert.IsEnabled = false;
            TxtStatus.Text = "Converting…";

            try
            {
                string jsonPath = Path.ChangeExtension(_selectedPath, ".json");

                // Parse on background thread
                _workbook = await Task.Run(() => ParseWorkbook(_selectedPath));

                // Save JSON on background thread
                await Task.Run(() =>
                {
                    var opts = new JsonSerializerOptions { WriteIndented = true };
                    File.WriteAllText(jsonPath, JsonSerializer.Serialize(_workbook, opts), Encoding.UTF8);
                });

                TxtStatus.Text = $"Saved: {Path.GetFileName(jsonPath)}";

                BuildPills();
                PanelViewer.Visibility = Visibility.Visible;
                RunSearch(string.Empty);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Error: {ex.Message}";
            }
            finally
            {
                BtnConvert.IsEnabled = true;
            }
        }

        // ── Sheet pills ───────────────────────────────────────────────────────
        private void BuildPills()
        {
            PanelSheetPills.Children.Clear();
            _pills.Clear();

            AddPill("__ALL__", "All sheets", isChecked: true);
            foreach (var name in _workbook.Keys)
                AddPill(name, name, isChecked: false);
        }

        private void AddPill(string key, string label, bool isChecked)
        {
            var pill = new ToggleButton
            {
                Content = label,
                IsChecked = isChecked,
                Margin = new Thickness(0),
                Padding = new Thickness(10, 4, 10, 4)
            };

            pill.Click += (s, e) =>
            {
                if (_updatingPills) return;
                SelectPill(key);
            };

            PanelSheetPills.Children.Add(pill);
            _pills[key] = pill;
        }

        private void SelectPill(string key)
        {
            _updatingPills = true;
            try
            {
                _activeSheetFilter = key == "__ALL__" ? null : key;
                foreach (var kv in _pills)
                    kv.Value.IsChecked = kv.Key == key;
            }
            finally
            {
                _updatingPills = false;
            }

            RunSearch(TxtSearch.Text);
        }

        // ── Search — Optimizado asíncronamente para 30,000 filas ─────────────────
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            var text = TxtSearch.Text;

            DispatcherQueue.TryEnqueue(async () =>
            {
                try
                {
                    await Task.Delay(150, token);
                    if (!token.IsCancellationRequested)
                    {
                        await Task.Run(() => RunSearchAsync(text, token), token);
                    }
                }
                catch (OperationCanceledException)
                {
                    // Ignorar silenciosamente: es el comportamiento esperado al cancelar búsquedas previas [1]
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error inesperado en búsqueda: {ex.Message}");
                }
            });
        }

        private void RunSearch(string query)
        {
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            RunSearchAsync(query, token);
        }

        private void RunSearchAsync(string query, CancellationToken token)
        {
            try
            {
                string q = query.Trim();
                bool hasQuery = q.Length > 0;
                string queryLower = q.ToLowerInvariant();

                // Filtrado de libros a escanear
                var sheetsToScan = _activeSheetFilter is null
                 ? _workbook
                    : _workbook.Where(kv => kv.Key == _activeSheetFilter)
                            .ToDictionary(kv => kv.Key, kv => kv.Value);

                var tempResults = new List<ResultRow>();

                foreach (var (sheetName, rows) in sheetsToScan)
                {
                    if (token.IsCancellationRequested) return;

                    // Paralelización óptima con PLINQ controlada por Token
                    var matchedRows = rows
                     .Select((row, index) => new { row, index })
                     .AsParallel()
                     .WithCancellation(token)
                     .Select(x =>
                     {
                         if (hasQuery)
                         {
                             var matched = x.row
                              .Where(kv => kv.Value != null &&
                                              kv.Value.ToString()!.Contains(queryLower, StringComparison.OrdinalIgnoreCase))
                              .ToList();

                             if (matched.Count > 0)
                             {
                                 return new ResultRow
                                 {
                                     SheetName = sheetName,
                                     RowLabel = $"Row {x.index + 1}  ·  {matched.Count} match{(matched.Count > 1 ? "es" : "")}",
                                     MatchSummary = string.Join("   |   ", matched.Select(kv => $"{kv.Key}: {kv.Value}")),
                                     RowData = x.row
                                 };
                             }
                         }
                         else
                         {
                             // OPTIMIZACIÓN CRÍTICA: Se limita el preview a las primeras 100 filas por hoja
                             if (x.index < 100)
                             {
                                 var preview = x.row
                                  .Where(kv => kv.Value != null && !string.IsNullOrWhiteSpace(kv.Value.ToString()))
                                  .Take(4)
                                  .Select(kv => $"{kv.Key}: {kv.Value}");

                                 return new ResultRow
                                 {
                                     SheetName = sheetName,
                                     RowLabel = $"Row {x.index + 1}",
                                     MatchSummary = string.Join("   |   ", preview),
                                     RowData = x.row
                                 };
                             }
                         }
                         return null;
                     })
                     .Where(x => x != null)
                     .ToList();

                    tempResults.AddRange(matchedRows!);
                }

                var finalUiBatch = tempResults.Take(1000).ToList();

                DispatcherQueue.TryEnqueue(() =>
                {
                    if (token.IsCancellationRequested) return;

                    LvResults.ItemsSource = null;
                    _results.Clear();

                    foreach (var item in finalUiBatch)
                    {
                        _results.Add(item);
                    }

                    LvResults.ItemsSource = _results;

                    int total = tempResults.Count;
                    TxtResultCount.Text = hasQuery
                     ? $"{total} result{(total != 1 ? "s" : "")} (showing top 1,000)"
                        : $"{total} row{(total != 1 ? "s" : "")} (preview)";

                    BtnCopyResults.IsEnabled = total > 0;
                });
            }
            catch (OperationCanceledException)
            {
                // Cancelación limpia silenciosa [1]
            }
        }

        // ── Evento Clic para Pop-up de Visualización de Fila ──────────────────────────
        private async void BtnRowDetails_Click(object sender, RoutedEventArgs e)
        {
            if (sender is HyperlinkButton btn && btn.DataContext is ResultRow row)
            {
                await ShowRowDetailsDialogAsync(row);
            }
        }

        private async Task ShowRowDetailsDialogAsync(ResultRow row)
        {
            // Contenedor principal scrollable
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 550,
                MinWidth = 480
            };

            // RESOLUCIÓN DE RECURSOS SEGURA DESDE EL RESOURCE DICTIONARY
            Brush? accentBrush = null;
            if (Application.Current.Resources.TryGetValue("AccentFillColorDefaultBrush", out object ab) && ab is Brush)
            {
                accentBrush = (Brush)ab;
            }
            accentBrush ??= new SolidColorBrush(Microsoft.UI.Colors.DodgerBlue); // Respaldo si no se encuentra

            Brush? defaultBrush = null;
            if (Application.Current.Resources.TryGetValue("TextFillColorPrimaryBrush", out object db) && db is Brush)
            {
                defaultBrush = (Brush)db;
            }
            defaultBrush ??= new SolidColorBrush(Microsoft.UI.Colors.Black); // Respaldo oscuro por defecto

            Brush? mutedBrush = null;
            if (Application.Current.Resources.TryGetValue("TextFillColorSecondaryBrush", out object mb) && mb is Brush)
            {
                mutedBrush = (Brush)mb;
            }
            mutedBrush ??= new SolidColorBrush(Microsoft.UI.Colors.Gray); // Respaldo gris para nulos

            // CREACIÓN DE UN ÚNICO RICHTEXTBLOCK PARA SOPORTAR SELECCIÓN Y COPIADO SEAMLESS
            var richTextBlock = new RichTextBlock
            {
                IsTextSelectionEnabled = true, // Permite marcar el texto con el ratón
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 13.5,
                Padding = new Thickness(0, 8, 16, 8)
            };

            // Estructura de cabecera decorativa simulando apertura de llaves JSON
            var openBracePara = new Paragraph();
            openBracePara.Inlines.Add(new Run { Text = "{", Foreground = defaultBrush });
            richTextBlock.Blocks.Add(openBracePara);

            // Iterar y estilizar dinámicamente cada campo del JSON de la fila
            int index = 0;
            int totalKeys = row.RowData.Count;
            foreach (var kvp in row.RowData)
            {
                var itemPara = new Paragraph
                {
                    Margin = new Thickness(16, 0, 0, 0) // Sangría de indentación JSON
                };

                // Campo / Clave
                var keyRun = new Run
                {
                    Text = $"\"{kvp.Key}\": ",
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
                };

                // Valor
                var valueRun = new Run();

                if (kvp.Value != null && !string.IsNullOrWhiteSpace(kvp.Value.ToString()))
                {
                    // Campo con Valor: Clave coloreada en el azul Accent corporativo de WinUI 3
                    keyRun.Foreground = accentBrush;

                    // Formatear texto del valor (comillas si es cadena de texto)
                    string formattedVal = kvp.Value is string ? $"\"{kvp.Value}\"" : kvp.Value.ToString()!;
                    valueRun.Text = formattedVal;
                    valueRun.Foreground = accentBrush;
                }
                else
                {
                    // Campo nulo: Clave en color neutro (gris/negro) y valor como "null" en cursiva tenue
                    keyRun.Foreground = defaultBrush;
                    valueRun.Text = "null";
                    valueRun.Foreground = mutedBrush;
                    valueRun.FontStyle = Windows.UI.Text.FontStyle.Italic;
                }

                itemPara.Inlines.Add(keyRun);
                itemPara.Inlines.Add(valueRun);

                // Insertar coma de separación JSON al final de cada línea (excepto la última)
                index++;
                if (index < totalKeys)
                {
                    itemPara.Inlines.Add(new Run { Text = ",", Foreground = defaultBrush });
                }

                richTextBlock.Blocks.Add(itemPara);
            }

            // Cierre de llaves JSON
            var closeBracePara = new Paragraph();
            closeBracePara.Inlines.Add(new Run { Text = "}", Foreground = defaultBrush });
            richTextBlock.Blocks.Add(closeBracePara);

            scrollViewer.Content = richTextBlock;

            // Instanciar ContentDialog de WinUI 3
            var dialog = new ContentDialog
            {
                Title = $"{row.SheetName}   ·   {row.RowLabel}",
                Content = scrollViewer,
                CloseButtonText = "Close",
                DefaultButton = ContentDialogButton.Close,
                XamlRoot = this.Content.XamlRoot // Obligatorio definir XamlRoot en WinUI 3 Desktop para evitar excepciones [2]
            };

            try
            {
                await dialog.ShowAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al abrir el cuadro de diálogo: {ex.Message}");
            }
        }

        // ── Copy results — grouped by sheet ───────────────────────────────────
        private void BtnCopyResults_Click(object sender, RoutedEventArgs e)
        {
            if (_results.Count == 0) return;

            var grouped = _results
                .GroupBy(r => r.SheetName)
                .ToList();

            var sb = new StringBuilder();
            var opts = new JsonSerializerOptions { WriteIndented = true };

            sb.AppendLine("{");
            for (int g = 0; g < grouped.Count; g++)
            {
                var group = grouped[g];
                var rowList = group.ToList();

                sb.AppendLine($"  \"{EscapeJson(group.Key)}\": [");

                for (int i = 0; i < rowList.Count; i++)
                {
                    string rowJson = JsonSerializer.Serialize(rowList[i].RowData, opts);
                    string indented = string.Join("\n", rowJson.Split('\n').Select(l => "    " + l));

                    sb.AppendLine(indented);

                    if (i < rowList.Count - 1)
                    {
                        sb.AppendLine("    ,");
                    }
                }

                sb.Append("  ]");
                if (g < grouped.Count - 1)
                    sb.AppendLine(",");
                else
                    sb.AppendLine();
            }

            sb.AppendLine("}");

            var pkg = new DataPackage();
            pkg.SetText(sb.ToString());
            Clipboard.SetContent(pkg);

            // Retroalimentación visual temporal
            BtnCopyResults.Content = "Copied ✓";
            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            timer.Tick += (s, _) =>
            {
                BtnCopyResults.Content = "Copy Results";
                timer.Stop();
            };
            timer.Start();
        }

        private static string EscapeJson(string s) =>
            s.Replace("\\", "\\\\").Replace("\"", "\\\"");

        // ── Excel parser ──────────────────────────────────────────────────────
        private static Dictionary<string, List<Dictionary<string, object?>>> ParseWorkbook(string path)
        {
            var workbook = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);

            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var cfg = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            };

            var ds = reader.AsDataSet(cfg);

            foreach (DataTable table in ds.Tables)
            {
                string sheet = table.TableName;
                int hdr = HeaderSkipRows.TryGetValue(sheet, out int skip) ? skip : 0;

                if (table.Rows.Count <= hdr)
                {
                    workbook[sheet] = new();
                    continue;
                }

                var headers = ExtractHeaders(table.Rows[hdr], table.Columns.Count);
                if (headers.Count == 0)
                {
                    workbook[sheet] = new();
                    continue;
                }

                var rows = new List<Dictionary<string, object?>>(table.Rows.Count);
                for (int r = hdr + 1; r < table.Rows.Count; r++)
                {
                    var dr = table.Rows[r];
                    if (IsRowEmpty(dr, headers.Count)) continue;

                    var obj = new Dictionary<string, object?>(headers.Count);
                    foreach (var (ci, cn) in headers)
                    {
                        object raw = dr[ci];
                        obj[cn] = raw == DBNull.Value ? null : Normalize(raw);
                    }

                    rows.Add(obj);
                }

                workbook[sheet] = rows;
            }

            return workbook;
        }

        private static List<(int Idx, string Name)> ExtractHeaders(DataRow row, int cols)
        {
            var list = new List<(int, string)>(cols);
            var seen = new Dictionary<string, int>(cols);

            for (int i = 0; i < cols; i++)
            {
                object raw = row[i];
                if (raw == DBNull.Value || raw is null) continue;

                string name = raw.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(name)) continue;

                if (seen.TryGetValue(name, out int cnt))
                {
                    seen[name] = cnt + 1;
                    name = $"{name}_{cnt + 1}";
                }
                else
                {
                    seen[name] = 1;
                }

                list.Add((i, name));
            }

            return list;
        }

        private static bool IsRowEmpty(DataRow row, int maxCols)
        {
            int limit = Math.Min(maxCols, row.Table.Columns.Count);
            for (int i = 0; i < limit; i++)
            {
                object v = row[i];
                if (v != DBNull.Value && v is not null && !string.IsNullOrWhiteSpace(v.ToString()))
                    return false;
            }

            return true;
        }

        private static object Normalize(object v) => v switch
        {
            double d when d == Math.Floor(d) && d is >= int.MinValue and <= int.MaxValue => (int)d,
            double d => d,
            DateTime dt => dt.ToString("yyyy-MM-ddTHH:mm:ss"),
            bool b => b,
            _ => v.ToString() ?? string.Empty
        };
    }
}
