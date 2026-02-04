using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;

namespace NetworkDiagramApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private string _excelPath = string.Empty;
        private string _outputName = string.Empty;
        private string _statusText = "待機中...";
        private string _logText = string.Empty;
        private bool _isProcessing = false;
        private List<string> _availableSheets = new List<string>();
        private string _selectedSheet = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            
            AddLog("アプリケーションを起動しました。");
            AddLog("Excelファイルを選択して、「図を生成」ボタンをクリックしてください。");
            
            // 起動時に更新チェック（非同期・非ブロッキング）
            _ = CheckForUpdatesAsync();
        }

        #region Properties

        public string ExcelPath
        {
            get => _excelPath;
            set
            {
                _excelPath = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanGenerate));
            }
        }

        public string OutputName
        {
            get => _outputName;
            set
            {
                _outputName = value;
                OnPropertyChanged();
            }
        }

        public string StatusText
        {
            get => _statusText;
            set
            {
                _statusText = value;
                OnPropertyChanged();
            }
        }

        public string LogText
        {
            get => _logText;
            set
            {
                _logText = value;
                OnPropertyChanged();
            }
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set
            {
                _isProcessing = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanGenerate));
            }
        }

        public bool CanGenerate => !string.IsNullOrEmpty(ExcelPath) && !IsProcessing;

        public List<string> AvailableSheets
        {
            get => _availableSheets;
            set
            {
                _availableSheets = value;
                OnPropertyChanged();
            }
        }

        public string SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                _selectedSheet = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers

        private void BtnCreateTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Title = "Excelテンプレートを保存",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "ネットワーク構成テンプレート.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ExcelTemplateGenerator.Generate(dialog.FileName);
                    AddLog($"✓ テンプレートを作成: {Path.GetFileName(dialog.FileName)}");
                    
                    var result = MessageBox.Show(
                        $"テンプレートを作成しました！\n\n{dialog.FileName}\n\nフォルダを開きますか？",
                        "完了",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);

                    if (result == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{dialog.FileName}\"");
                    }
                }
                catch (Exception ex)
                {
                    AddLog($"✗ テンプレート作成エラー: {ex.Message}");
                    MessageBox.Show($"テンプレート作成に失敗しました:\n\n{ex.Message}",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Excelファイルを選択してください",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                ExcelPath = dialog.FileName;
                AddLog($"✓ ファイルを選択: {Path.GetFileName(ExcelPath)}");
                
                // シート一覧を取得
                try
                {
                    var sheets = ExcelReader.GetSheetNames(ExcelPath);
                    AvailableSheets = sheets;
                    
                    // デフォルトで「構成図作成」または「構成」を選択
                    if (sheets.Contains("構成図作成"))
                    {
                        SelectedSheet = "構成図作成";
                    }
                    else if (sheets.Contains("構成"))
                    {
                        SelectedSheet = "構成";
                    }
                    else if (sheets.Count > 0)
                    {
                        SelectedSheet = sheets[0];
                    }
                    
                    AddLog($"   利用可能なシート: {string.Join(", ", sheets)}");
                    AddLog($"   選択中のシート: {SelectedSheet}");
                }
                catch (Exception ex)
                {
                    AddLog($"⚠ シート一覧の取得に失敗: {ex.Message}");
                }
            }
        }

        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ExcelPath))
            {
                MessageBox.Show("Excelファイルを選択してください。",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!File.Exists(ExcelPath))
            {
                MessageBox.Show("選択されたファイルが見つかりません。",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            IsProcessing = true;
            StatusText = "処理中...";

            try
            {
                string outputDir = Path.GetDirectoryName(ExcelPath) ?? "";
                string outputName = string.IsNullOrWhiteSpace(OutputName)
                    ? "network.drawio"
                    : OutputName;

                if (!outputName.EndsWith(".drawio", StringComparison.OrdinalIgnoreCase))
                {
                    outputName += ".drawio";
                }

                string outputPath = Path.Combine(outputDir, outputName);

                AddLog("");
                AddLog("========================================");
                AddLog("図の生成を開始します...");
                AddLog("========================================");

                // Task.Runの中でログを呼び出さず、UIスレッドで直接実行
                await Task.Run(() =>
                {
                    try
                    {
                        GenerateDrawIO(ExcelPath, outputPath, SelectedSheet);
                    }
                    catch (Exception ex)
                    {
                        // UIスレッドで例外を再スロー
                        Dispatcher.Invoke(() => throw ex);
                    }
                });

                AddLog("========================================");
                AddLog($"✓ 出力完了: {outputPath}");
                AddLog("========================================");
                StatusText = "完了！";

                var result = MessageBox.Show(
                    $"図の生成が完了しました！\n\n{outputPath}\n\nフォルダを開きますか？",
                    "完了",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{outputPath}\"");
                }
            }
            catch (Exception ex)
            {
                AddLog("");
                AddLog($"✗ エラーが発生しました: {ex.Message}");
                StatusText = "エラー";
                
                MessageBox.Show($"エラーが発生しました:\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                if (StatusText == "処理中...")
                {
                    StatusText = "待機中...";
                }
            }
        }

        #endregion

        #region Core Logic

        private void GenerateDrawIO(string excelPath, string outputPath, string sheetName)
        {
            Dispatcher.Invoke(() => AddLog("📖 Excelファイルを読み込んでいます..."));
            var nodes = ExcelReader.LoadNodes(excelPath, (msg) => Dispatcher.Invoke(() => AddLog(msg)), sheetName);
            Dispatcher.Invoke(() => AddLog($"   → {nodes.Count} 件のノードを読み込みました"));

            Dispatcher.Invoke(() => AddLog("🌳 ツリー構造を構築しています..."));
            var tree = new TreeBuilder(nodes);
            Dispatcher.Invoke(() => AddLog($"   → ルートノード: {tree.Roots.Count} 件"));

            Dispatcher.Invoke(() => AddLog("📐 レイアウトを計算しています..."));
            var calculator = new TreeLayoutCalculator(tree);
            var positions = calculator.Calculate();

            // VLANグループを抽出
            var vlanGroups = new System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<string>>();
            foreach (var kvp in nodes)
            {
                if (kvp.Value.VLAN.HasValue)
                {
                    if (!vlanGroups.ContainsKey(kvp.Value.VLAN.Value))
                    {
                        vlanGroups[kvp.Value.VLAN.Value] = new System.Collections.Generic.List<string>();
                    }
                    vlanGroups[kvp.Value.VLAN.Value].Add(kvp.Key);
                }
            }

            if (vlanGroups.Count > 0)
            {
                Dispatcher.Invoke(() => AddLog($"🔧 VLAN枠の配置を調整しています... ({vlanGroups.Count} 個のVLAN)"));
                var vlanManager = new VLANFrameManager();
                vlanManager.ResolveCollisions(vlanGroups, positions);
            }

            Dispatcher.Invoke(() => AddLog("📝 draw.ioファイルを生成しています..."));
            var generator = new DrawIOGenerator(nodes, positions);
            generator.Generate(outputPath);
        }

        #endregion

        #region Helper Methods

        private void AddLog(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => AddLog(message));
                return;
            }

            if (!string.IsNullOrEmpty(LogText))
            {
                LogText += "\n";
            }
            LogText += $"[{DateTime.Now:HH:mm:ss}] {message}";
        }

        private System.Windows.Controls.ScrollViewer? FindScrollViewer(DependencyObject obj)
        {
            for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(obj, i);
                if (child is System.Windows.Controls.ScrollViewer scrollViewer)
                {
                    return scrollViewer;
                }
                var result = FindScrollViewer(child);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        #endregion

        #region Update Check

        private async Task CheckForUpdatesAsync()
        {
            try
            {
                await Task.Delay(1000); // UI表示後に実行

                AddLog($"現在のバージョン: v{UpdateChecker.GetCurrentVersion()}");
                AddLog("更新を確認中...");

                var updateInfo = await UpdateChecker.CheckForUpdatesAsync();

                if (!string.IsNullOrEmpty(updateInfo.ErrorMessage))
                {
                    // エラーは無視（ログのみ）
                    AddLog($"  → {updateInfo.ErrorMessage}");
                    return;
                }

                if (updateInfo.HasUpdate)
                {
                    AddLog($"✨ 新しいバージョン v{updateInfo.LatestVersion} が利用可能です！");
                    
                    // 更新通知ダイアログ
                    ShowUpdateNotification(updateInfo);
                }
                else
                {
                    AddLog("  → 最新バージョンを使用中");
                }
            }
            catch (Exception ex)
            {
                // 更新チェック失敗は無視
                AddLog($"更新チェックエラー: {ex.Message}");
            }
        }

        private void ShowUpdateNotification(UpdateChecker.UpdateInfo updateInfo)
        {
            var message = $"新しいバージョンが利用可能です！\n\n" +
                          $"現在のバージョン: v{updateInfo.CurrentVersion}\n" +
                          $"最新バージョン: v{updateInfo.LatestVersion}\n\n";

            if (!string.IsNullOrEmpty(updateInfo.ReleaseNotes))
            {
                var notes = updateInfo.ReleaseNotes;
                if (notes.Length > 200)
                {
                    notes = notes.Substring(0, 200) + "...";
                }
                message += $"更新内容:\n{notes}\n\n";
            }

            message += "ダウンロードページを開きますか？";

            var result = MessageBox.Show(
                message,
                "更新のお知らせ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = updateInfo.DownloadUrl,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ブラウザを開けませんでした: {ex.Message}",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        #endregion

        #region Drag and Drop

        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドロップゾーンの背景色を変更
                if (sender is Border border)
                {
                    border.Background = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(230, 244, 255));
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Border_DragLeave(object sender, DragEventArgs e)
        {
            // 元の背景色に戻す
            if (sender is Border border)
            {
                border.Background = System.Windows.Media.Brushes.White;
            }
        }

        private void Border_Drop(object sender, DragEventArgs e)
        {
            // 元の背景色に戻す
            if (sender is Border border)
            {
                border.Background = System.Windows.Media.Brushes.White;
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    
                    // .xlsx または .xls ファイルのみ受け付ける
                    if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        ExcelPath = filePath;
                        AddLog($"✓ ファイルをドロップ: {Path.GetFileName(filePath)}");
                        
                        // シート一覧を取得
                        try
                        {
                            var sheets = ExcelReader.GetSheetNames(filePath);
                            AvailableSheets = sheets;
                            
                            // デフォルトで「構成図作成」または「構成」を選択
                            if (sheets.Contains("構成図作成"))
                            {
                                SelectedSheet = "構成図作成";
                            }
                            else if (sheets.Contains("構成"))
                            {
                                SelectedSheet = "構成";
                            }
                            else if (sheets.Count > 0)
                            {
                                SelectedSheet = sheets[0];
                            }
                            
                            AddLog($"   利用可能なシート: {string.Join(", ", sheets)}");
                            AddLog($"   選択中のシート: {SelectedSheet}");
                        }
                        catch (Exception ex)
                        {
                            AddLog($"⚠ シート一覧の取得に失敗: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Excelファイル (.xlsx または .xls) を選択してください。",
                            "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
