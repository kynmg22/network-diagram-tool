using System;
using System.IO;
using System.Reflection;

namespace NetworkDiagramApp
{
    public class ExcelTemplateGenerator
    {
        // 埋め込みリソースからテンプレートを出力
        public static void Generate(string outputPath)
        {
            try
            {
                // テンプレートファイルのパスを取得
                string templatePath = GetTemplatePath();
                
                if (File.Exists(templatePath))
                {
                    // テンプレートファイルをコピー
                    File.Copy(templatePath, outputPath, true);
                }
                else
                {
                    // テンプレートファイルが見つからない場合は簡易版を生成
                    GenerateSimpleTemplate(outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"テンプレート生成エラー: {ex.Message}", ex);
            }
        }

        private static string GetTemplatePath()
        {
            // 実行ファイルと同じフォルダのTemplateフォルダを確認
            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string templatePath = Path.Combine(exePath, "Template", "構成図作成_template.xlsx");
            
            return templatePath;
        }

        // テンプレートファイルがない場合の簡易版生成
        private static void GenerateSimpleTemplate(string outputPath)
        {
            using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(outputPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                // 構成シート
                var worksheetPart1 = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                worksheetPart1.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                var sheet1 = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart1),
                    SheetId = 1,
                    Name = "構成"
                };
                sheets.Append(sheet1);

                var sheetData1 = worksheetPart1.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();

                // ヘッダー行
                var headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = 1 };
                headerRow.Append(
                    CreateCell("A1", "接続元ID(自動: カンマ区切り)"),
                    CreateCell("B1", "ID（自動）"),
                    CreateCell("C1", "ID（選択）"),
                    CreateCell("D1", "機器名"),
                    CreateCell("E1", "IPアドレス"),
                    CreateCell("F1", "VLANID(数字のみ:例 1,10)複数非対応"),
                    CreateCell("G1", "備考"),
                    CreateCell("H1", "接続元1(選択)"),
                    CreateCell("I1", "接続元2(選択)"),
                    CreateCell("J1", "接続元3(選択)")
                );
                sheetData1.Append(headerRow);

                // IDリストシート
                var worksheetPart2 = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                worksheetPart2.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                var sheet2 = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart2),
                    SheetId = 2,
                    Name = "IDリスト"
                };
                sheets.Append(sheet2);

                var sheetData2 = worksheetPart2.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                var idListHeader = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = 1 };
                idListHeader.Append(
                    CreateCell("A1", "ID"),
                    CreateCell("B1", "機器名")
                );
                sheetData2.Append(idListHeader);

                workbookPart.Workbook.Save();
            }
        }

        private static DocumentFormat.OpenXml.Spreadsheet.Cell CreateCell(string cellReference, string value)
        {
            return new DocumentFormat.OpenXml.Spreadsheet.Cell()
            {
                CellReference = cellReference,
                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString,
                InlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value))
            };
        }
    }
}
