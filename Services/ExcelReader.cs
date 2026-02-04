using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NetworkDiagramApp
{
    public class ExcelReader
    {
        public static List<string> GetSheetNames(string excelPath)
        {
            var sheetNames = new List<string>();
            
            try
            {
                using var document = SpreadsheetDocument.Open(excelPath, false);
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null) return sheetNames;

                var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
                foreach (var sheet in sheets)
                {
                    if (sheet.Name != null)
                    {
                        sheetNames.Add(sheet.Name.Value);
                    }
                }
            }
            catch
            {
                // エラーは無視
            }
            
            return sheetNames;
        }

        private static void SafeLog(Action<string> logger, string message)
        {
            try
            {
                logger?.Invoke(message);
            }
            catch
            {
                // ログ失敗は無視
            }
        }

        public static Dictionary<string, NetworkNode> LoadNodes(string excelPath, Action<string> logger, string? sheetName = null)
        {
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException($"Excelファイルが見つかりません: {excelPath}");
            }

            var nodes = new Dictionary<string, NetworkNode>();

            try
            {
                using var document = SpreadsheetDocument.Open(excelPath, false);
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    throw new Exception("ワークブックを読み込めませんでした。");
                }

                // シートを選択
                var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
                DocumentFormat.OpenXml.Spreadsheet.Sheet? targetSheet = null;
                
                if (!string.IsNullOrEmpty(sheetName))
                {
                    // 指定されたシート名で検索
                    targetSheet = sheets.FirstOrDefault(s => s.Name == sheetName);
                    if (targetSheet != null)
                    {
                        SafeLog(logger, $"✓ シート「{sheetName}」を読み込みます");
                    }
                    else
                    {
                        SafeLog(logger, $"⚠ シート「{sheetName}」が見つかりません。");
                    }
                }
                
                if (targetSheet == null)
                {
                    // 「構成図作成」または「構成」を探す
                    targetSheet = sheets.FirstOrDefault(s => s.Name == "構成図作成") ?? 
                                  sheets.FirstOrDefault(s => s.Name == "構成");
                    
                    if (targetSheet != null)
                    {
                        SafeLog(logger, $"✓ シート「{targetSheet.Name}」を読み込みます");
                    }
                }
                
                WorksheetPart? worksheetPart = null;
                
                if (targetSheet != null)
                {
                    string relationshipId = targetSheet.Id?.Value ?? "";
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                }
                else
                {
                    // 見つからない場合は最初のシート
                    worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    SafeLog(logger, "⚠ 対象シートが見つかりません。最初のシートを使用します。");
                }

                if (worksheetPart == null)
                {
                    throw new Exception("ワークシートが見つかりません。");
                }

                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Elements<Row>().ToList();

                SafeLog(logger, $"📊 総行数: {rows.Count}");

                bool dataStarted = false;
                int rowIndex = 0;

                foreach (var row in rows)
                {
                    rowIndex++;
                    var cells = row.Elements<Cell>().ToList();

                    // B列（ID）を取得
                    string? idValue = GetCellValue(cells, "B", workbookPart);
                    
                    // デバッグ: 最初の10行を表示
                    if (rowIndex <= 10)
                    {
                        string dValue = GetCellValue(cells, "D", workbookPart) ?? "";
                        SafeLog(logger, $"   行{rowIndex}: B='{idValue ?? "(空)"}'  D='{dValue}'");
                    }
                    
                    if (string.IsNullOrWhiteSpace(idValue))
                    {
                        if (dataStarted)
                        {
                            SafeLog(logger, $"   → データ終了 (行{rowIndex}で空白を検出)");
                            break;
                        }
                        continue;
                    }

                    if (idValue == "ここまで") 
                    {
                        SafeLog(logger, $"   → 「ここまで」マーカーを検出 (行{rowIndex})");
                        break;
                    }

                    // ヘッダー行をスキップ
                    string? parentsValue = GetCellValue(cells, "A", workbookPart);
                    
                    bool isHeader = IsHeaderRow(idValue, parentsValue);
                    
                    if (rowIndex <= 3)
                    {
                        SafeLog(logger, $"   A列='{parentsValue ?? "(空)"}'  ヘッダー判定={isHeader}");
                    }
                    
                    if (isHeader)
                    {
                        SafeLog(logger, $"   → ヘッダー行をスキップ (行{rowIndex})");
                        dataStarted = true;
                        continue;
                    }

                    dataStarted = true;

                    // ID検証
                    if (nodes.ContainsKey(idValue))
                    {
                        throw new Exception($"行 {rowIndex}: ID重複「{idValue}」");
                    }

                    // ノードデータを作成
                    var node = new NetworkNode
                    {
                        ID = idValue,
                        Name = GetCellValue(cells, "D", workbookPart) ?? idValue,
                        IP = GetCellValue(cells, "E", workbookPart) ?? string.Empty,
                        VLAN = ExtractVLANNumber(GetCellValue(cells, "F", workbookPart)),
                        Note = GetCellValue(cells, "G", workbookPart) ?? string.Empty,
                        Parents = SplitParents(parentsValue),
                        ExcelRow = rowIndex
                    };

                    nodes[idValue] = node;
                }

                SafeLog(logger, $"✓ 読み込み完了: {nodes.Count}件のノード");

                if (nodes.Count == 0)
                {
                    throw new Exception(
                        "❌ データが見つかりません。以下を確認してください:\n\n" +
                        "1. B列にIDが入力されているか\n" +
                        "2. ヘッダー行に「ID」という文字が含まれているか\n" +
                        "3. データ開始行より上に空行がないか\n" +
                        "4. シートが正しく選択されているか\n\n" +
                        $"上記ログで最初の10行のB列とC列の内容を確認してください。");
                }

                // 接続元IDの検証
                ValidateParentReferences(nodes, logger);

                return nodes;
            }
            catch (IOException ex)
            {
                throw new IOException(
                    "Excelファイルを開けません。\n" +
                    "Excelで開いている場合は閉じてから再度実行してください。\n" +
                    $"ファイル: {excelPath}", ex);
            }
        }

        private static string? GetCellValue(List<Cell> cells, string columnName, WorkbookPart workbookPart)
        {
            var cell = cells.FirstOrDefault(c => GetColumnName(c.CellReference?.Value) == columnName);
            if (cell == null) return null;

            // 数式セルの場合、CachedValueを取得
            if (cell.CellFormula != null && cell.CellValue != null)
            {
                string cachedValue = cell.CellValue.Text;
                
                // SharedStringの場合
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                    if (stringTable != null && int.TryParse(cachedValue, out int index))
                    {
                        // ふりがな（ルビ）を除外してテキストのみ取得
                        var sharedString = stringTable.ElementAt(index) as DocumentFormat.OpenXml.Spreadsheet.SharedStringItem;
                        if (sharedString != null)
                        {
                            return GetTextWithoutPhonetic(sharedString);
                        }
                    }
                }
                
                return cachedValue?.Trim();
            }

            // 通常のセル値を取得
            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (stringTable != null && int.TryParse(value, out int index))
                {
                    // ふりがな（ルビ）を除外してテキストのみ取得
                    var sharedString = stringTable.ElementAt(index) as DocumentFormat.OpenXml.Spreadsheet.SharedStringItem;
                    if (sharedString != null)
                    {
                        return GetTextWithoutPhonetic(sharedString);
                    }
                }
            }

            return value?.Trim();
        }

        private static string GetTextWithoutPhonetic(DocumentFormat.OpenXml.Spreadsheet.SharedStringItem sharedString)
        {
            // Text要素のみを取得（PhoneticRun要素は無視）
            var textElements = sharedString.Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>();
            var textParts = new List<string>();
            
            foreach (var textElement in textElements)
            {
                // PhoneticRun内のTextは除外
                if (textElement.Parent?.LocalName != "rPh") // rPh = PhoneticRun
                {
                    textParts.Add(textElement.Text);
                }
            }
            
            return string.Join("", textParts).Trim();
        }

        private static string? GetColumnName(string? cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return null;
            return Regex.Match(cellReference, "[A-Z]+").Value;
        }

        private static bool IsHeaderRow(string idValue, string? parentsValue)
        {
            // B列の値をチェック
            if (idValue.Contains("ID") || idValue.Contains("id"))
            {
                return true;
            }
            
            // 「機器」「IPアドレス」「VLAN」などの列名っぽい文字
            if (idValue.Contains("機器") || idValue.Contains("名") || 
                idValue.Contains("アドレス") || idValue.Contains("VLAN") ||
                idValue.Contains("備考") || idValue.Contains("接続"))
            {
                return true;
            }

            // G列（接続元ID）のヘッダーパターン
            if (!string.IsNullOrEmpty(parentsValue))
            {
                if (parentsValue.Contains("接続") || parentsValue.Contains("ID"))
                {
                    return true;
                }
            }

            return false;
        }

        private static int? ExtractVLANNumber(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            var match = Regex.Match(value, @"(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : null;
        }

        private static List<string> SplitParents(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return new List<string>();

            return value.Split(',')
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrEmpty(p))
                .ToList();
        }

        private static void ValidateParentReferences(Dictionary<string, NetworkNode> nodes, Action<string> logger)
        {
            foreach (var kvp in nodes)
            {
                var node = kvp.Value;
                for (int i = 0; i < node.Parents.Count; i++)
                {
                    string parentID = node.Parents[i];

                    if (!nodes.ContainsKey(parentID))
                    {
                        // 類似IDを検索
                        string? similarID = FindSimilarID(parentID, nodes.Keys);
                        if (similarID != null)
                        {
                            SafeLog(logger, $"⚠ 警告: 接続元ID「{parentID}」→「{similarID}」として解釈します。");
                            node.Parents[i] = similarID;
                        }
                        else
                        {
                            var availableIDs = string.Join(", ", nodes.Keys.Take(10));
                            throw new Exception(
                                $"未定義の接続元ID: {parentID} → {node.ID}\n" +
                                $"B列に「{parentID}」というIDが存在しません。\n" +
                                $"利用可能なID: {availableIDs}...");
                        }
                    }
                }
            }
        }

        private static string? FindSimilarID(string target, IEnumerable<string> availableIDs)
        {
            string targetNormalized = target.Replace("_", "").Replace("-", "");

            foreach (string id in availableIDs)
            {
                string idNormalized = id.Replace("_", "").Replace("-", "");
                if (targetNormalized.Equals(idNormalized, StringComparison.OrdinalIgnoreCase))
                {
                    return id;
                }
            }

            // 数字部分のマッチング (例: "UTM2" → "UTM_2")
            var targetMatch = Regex.Match(target, @"^([A-Za-z\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]+)(\d+)$");
            if (targetMatch.Success)
            {
                string targetBase = targetMatch.Groups[1].Value;
                string targetNum = targetMatch.Groups[2].Value;

                foreach (string id in availableIDs)
                {
                    if (id == $"{targetBase}_{targetNum}" || id == $"{targetBase}{targetNum}")
                    {
                        return id;
                    }
                }
            }

            return null;
        }
    }
}
