//#define DEBUG  //デバッグログが必要な場合、この行を有効にしてください
using Azure.AI.FormRecognizer.DocumentAnalysis;
using Azure;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
//using System.Diagnostics.Metrics;
using System.Linq;
using System.Threading.Tasks;
using UiPath.DocumentProcessing.Contracts;
using UiPath.DocumentProcessing.Contracts.DataExtraction;
using UiPath.DocumentProcessing.Contracts.Dom;
using UiPath.DocumentProcessing.Contracts.Results;
using UiPath.DocumentProcessing.Contracts.Taxonomy;
using System.IO;
using System.Drawing;
using Newtonsoft.Json.Linq;
using UiPath.OCR.Contracts.DataContracts;
using System.Windows.Markup;

namespace SampleActivities.Basic.DataExtraction
{
    // DUが期待するデータとして登録するための2次元配列クラス
    public class Cell
    {
        public string Content { get; set; }                 // 実際の値
        public int RowIndex { get; set; }                   // 行の位置
        public int ColumnIndex { get; set; }                // 列の位置
        public List<PointF> BoundingPolygon { get; set; }   // 座標情報
        public int TableRows { get; private set; }          // テーブルの総行数
        public int TableColumns { get; private set; }       // テーブルの総列数
        public int PageNumber { get; private set; }         // テーブルの総ページ数

        public Cell(string content, int rowIndex, int columnIndex, List<PointF> boundingPolygon, int tableRows, int tableColumns, int pageNumber)
        {
            Content = content;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            BoundingPolygon = boundingPolygon;
            TableRows = tableRows;
            TableColumns = tableColumns;
            PageNumber = pageNumber;
        }
    }

    [DisplayName("Sample Azure-Layout Extractor")]
    public class AzureLayout : ExtractorAsyncCodeActivity
    {
        [Category("Server")]
        [RequiredArgument]
        [Description("MLモデルサービス endpoint 情報")]
        public InArgument<string> Endpoint { get; set; }

        [Category("Server")]
        [RequiredArgument]
        [Description("MLモデルサービス endpoint Api Key 情報")]
        public InArgument<string> ApiKey { get; set; }

        [Category("Server")]                // DU 抽出結果をオプションとして指定可能にする
        [RequiredArgument]
        [Description("DU 抽出結果")]
        public InArgument<ExtractionResult> DuExtractionResult { get; set; }

        [Category("Server")]                // 読取対象テーブルインデックスをオプションとして指定可能にする
        [RequiredArgument]
        [Description("読取対象テーブルインデックス（1つ目のテーブルは 0 を指定します）")]
        public InArgument<int> TableIndex { get; set; }

        Object lockObj = new Object();

        ExtractorResult result;
        ExtractionResult du_result;
        List<PageLayout> pages;
        int tableIndex;
        string[] itemFields;

        public override Task<ExtractorDocumentTypeCapabilities[]> GetCapabilities()
        {
#if DEBUG
            Console.WriteLine("GetCapabilities called");
#endif
            //Azure Form Recognizer invoice fields definition 
            List<ExtractorFieldCapability> fields = InitializeFields();

            return Task.FromResult(new[] {
                new ExtractorDocumentTypeCapabilities{
                    DocumentTypeId = "azure.layout.demo",
                    Fields = fields.ToArray()
                }
            });
        }

        private List<ExtractorFieldCapability> InitializeFields()
        {
            var fields = new List<ExtractorFieldCapability>();

            // Azure Invoice Custom Activity をベースにしています（必要に応じて変更が必要です）
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PurchaseOrder", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorName", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorTaxId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "VendorAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerTaxId", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CustomerAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "BillingAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "BillingAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ShippingAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ShippingAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PaymentTerm", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "SubTotal", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "TotalTax", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "InvoiceTotal", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "AmountDue", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "RemittanceAddress", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "RemittanceAddressRecipient", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceStartDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "ServiceEndDate", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PreviousUnpaidBalance", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "CurrencyCode", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "PaymentOptions", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "TotalDiscount", Components = new ExtractorFieldCapability[0], SetValues = new string[0] });
            fields.Add(new ExtractorFieldCapability { FieldId = "Items", Components = new[] {
                new ExtractorFieldCapability {FieldId = "Description", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "UnitPrice", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Quantity", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Amount", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "ProductCode", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Unit", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Date", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "Tax", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                new ExtractorFieldCapability {FieldId = "TaxRate", Components = new ExtractorFieldCapability[0], SetValues = new string[0]},
                }, SetValues = new string[0] });
            return fields;
        }

        public override Boolean ProvidesCapabilities()
        {
            return true;
        }

        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, object state)
        {
            //get arguments passed to DataExtractionScope 
            ExtractorDocumentType documentType = ExtractorDocumentType.Get(context);
            ResultsDocumentBounds documentBounds = DocumentBounds.Get(context);
            string text = DocumentText.Get(context);
            Document document = DocumentObjectModel.Get(context);
            string documentPath = DocumentPath.Get(context);
            string endpoint = Endpoint.Get(context);
            string apiKey = ApiKey.Get(context);

            this.du_result = DuExtractionResult.Get(context);   // DU 抽出結果をオプションとして指定可能にする
            this.tableIndex = TableIndex.Get(context);          // 読取対象テーブルインデックスをオプションとして指定可能にする
            this.pages = new List<PageLayout>();

            var task = new Task( _ => Execute(documentType, documentBounds, text, document, documentPath, endpoint, apiKey), state);
            task.Start();
            if (callback != null)
            {
                task.ContinueWith(s => callback(s));
                task.Wait();
            }
            return task;
        }

        protected override async void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            var task = (Task)result;
            ExtractorResult.Set(context, this.result);
            await task;
        }

        // これ以降に Custom Activity 処理が実装されています。
        protected void Execute(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                                    string text, Document document, string documentPath,
                                    string endPoint, string apiKey)                 // DU から必要な情報が引き渡されています
        {
            this.result = ComputeResult(documentType, documentBounds, text, document, documentPath, endPoint, apiKey);
        }

        private ExtractorResult ComputeResult(ExtractorDocumentType documentType, ResultsDocumentBounds documentBounds,
                                string text, Document dom, string documentPath, string endpoint, string apiKey)
        {
#if DEBUG
            Console.WriteLine("ComputeResult called");
#endif
            var credential = new AzureKeyCredential(apiKey);                        // Activity で設定したパラメータが使われています
            var client = new DocumentAnalysisClient(new Uri(endpoint), credential); // Activity で設定したパラメータが使われています
            var extractorResult = new ExtractorResult();                            // DU へ戻す結果領域
            var resultsDataPoints = new List<ResultsDataPoint>();                   // 結果に含まれる内部的な領域（位置情報などを含む）

            //Azure Form Recognizer invoice fields definition 
            List<ExtractorFieldCapability> fields = InitializeFields();

            // 明細データの列名リストを取得します
            var itemsField = fields.FirstOrDefault(field => field.FieldId == "Items");
            if (itemsField != null && itemsField.Components.Length > 0)
            {
                // "Items" コンポーネントの FieldId を String[] へ格納
                itemFields = itemsField.Components.Select(component => component.FieldId).ToArray();
            }
            else
            {
                itemFields = new string[0]; // 列がない場合は空の配列を返します
            }
#if DEBUG
            foreach (var item in itemFields)
            {
                Console.WriteLine($"item of ItemFields = {item}");
            }
#endif

            // Azure Layout API の呼び出し
            AnalyzeDocumentOperation operation = client.AnalyzeDocument(WaitUntil.Completed, "prebuilt-layout", File.OpenRead(documentPath));
            AnalyzeResult result = operation.Value;

            // Azure Layout API 結果から各ページの座標情報を設定 
            foreach (var x in result.Pages)
            {
                this.pages.Add(new PageLayout((double)x.Width, (double)x.Height, x.Unit.ToString()));
            }

            // ExtractorDocumentType の各フィールドの型を確認し、型ごとの処理を実行します。
            // Azure Invoice Custom Activity をベースにしたため、このような実装になっています。
            foreach (var du_field in documentType.Fields)
            {
                if (du_field.Type == FieldType.Text)
                {
                    // テキストの場合
                    var dp = CreateTextFieldDataPoint(du_field, this.du_result, dom, pages.ToArray());
                    if (dp != null)
                    {
                        resultsDataPoints.Add(dp);
                    }
                }
                else if (du_field.Type == FieldType.Date)
                {
                    // 標準フィールドの日付型の項目へ対応するロジックが必要
                }
                else if (du_field.Type == FieldType.Number)
                {
                    // 標準フィールドの数値型の項目へ対応するロジックが必要
                }
                else if (du_field.Type == FieldType.Table)
                {
                    // テーブルの場合
                    resultsDataPoints.Add(CreateTableFieldDataPoint(du_field, result, dom, pages.ToArray(), this.tableIndex, this.itemFields));
                }
            }
            extractorResult.DataPoints = resultsDataPoints.ToArray();
            return extractorResult;
        }

        // DU の OCR 結果をマージしています
        private static ResultsDataPoint CreateTextFieldDataPoint(Field du_field, ExtractionResult du_result, Document dom, PageLayout[] pages)
        {
#if DEBUG
            Console.WriteLine("CreateTextFieldDataPoint called");
            Console.WriteLine($"Field Name : {du_field.FieldName}");
#endif
            // 標準フィールドに値が設定されている場合は ResultsDataPoint として戻します
            if (du_result.ResultsDocument.Fields.FirstOrDefault(f => f.FieldName == du_field.FieldName).Values.Length > 0 &&
                du_result.ResultsDocument.Fields.FirstOrDefault(f => f.FieldName == du_field.FieldName).Values[0].Value != null)
            {
                // 指定されたフィールドの情報を取得しています
                var fieldValue = du_result.ResultsDocument.Fields.FirstOrDefault(f => f.FieldName == du_field.FieldName);

                String du_value = fieldValue.Values[0].Value;                   // 値
                var reference = fieldValue.Values[0].Reference;                 // 位置情報
                float confidence = fieldValue.Values[0].Confidence;             // 信頼度
                float ocrconfidence = fieldValue.Values[0].OcrConfidence;       // OCR 信頼度

                ResultsValue resultsValue = new ResultsValue(du_value, reference, confidence, ocrconfidence);

                return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                du_field.Type,
                new[] { resultsValue });
            }
            else
            {
                return null;
            }
        }

        private static ResultsDataPoint CreateTableFieldDataPoint(Field du_field, AnalyzeResult az_result, Document dom, PageLayout[] pages, int tableIndex, string[] itemFields)
        {
#if DEBUG
            Console.WriteLine("CreateTableFieldDataPoint called");
            foreach (var item in itemFields)
            {
                Console.WriteLine($"item of ItemFields = {item}");
            }
#endif
            int i = 0;
            String fieldName;
            List<ResultsDataPoint> dataPoints = new List<ResultsDataPoint>();
            List<IEnumerable<ResultsDataPoint>> rows = new List<IEnumerable<ResultsDataPoint>>();

            // Azure Layout API 結果を2次元配列へ変換
            Cell[,] cells = ConvertTableTo2DArrayWithPosition(az_result, tableIndex);

            // テーブルに含まれる行数分のループ（0行目のヘッダー部分は読取対象から除いています）
            for (int rowIndex = 1; rowIndex < cells[0,0].TableRows; rowIndex++)
            {
                List<ResultsDataPoint> row = new List<ResultsDataPoint>();

                // テーブルに含まれる列数分のループ
                for (int columnIndex = 0; columnIndex < cells[0,0].TableColumns; columnIndex++)
                {
                    Cell cell = cells[rowIndex, columnIndex];
                    if (cell != null && !cell.Content.Equals("")) // セルが存在する、かつ値を含む場合のみ処理する
                    {
                        // セルの内容を基に ResultsValue を作成 （右2つの信頼度に関するパラメータは暫定的に 1.0 を指定）
                        ResultsValue cellValue = new ResultsValue(cell.Content, new ResultsContentReference(cell.PageNumber, 0, null), 1.0f, 1.0f);

                        // 明細データの列名リストから columnIndex の位置に対応した列名を設定します
                        if (itemFields != null && itemFields.Length >= columnIndex && itemFields[columnIndex] != "")
                        {
                            fieldName = itemFields[columnIndex];
                        }
                        else
                        {
                            fieldName = "";
                        }
#if DEBUG
                        Console.WriteLine($"fieldName = {fieldName}");
#endif

                        // ResultsDataPoint を作成
                        row.Add(new ResultsDataPoint(fieldName, fieldName, FieldType.Text,
                                    new [] { CreateResultsValue(i++, dom, cell, pages) }));
                    }
                }

                if (row.Count > 0) // 対象となる行に列データが存在する場合のみ追加
                {
                    rows.Add(row);
                }
            }

            // 将来的に非推奨になると警告が出ているが、サンプルが元々このメソッドを利用していた （右2つの信頼度に関するパラメータは暫定的に 1.0 を指定）
            var tableValue = ResultsValue.CreateTableValue(du_field, dataPoints, rows.ToArray(), 1.0f, 1.0f);

            // テーブルデータの ResultsDataPoint を作成
            return new ResultsDataPoint(
                du_field.FieldId,
                du_field.FieldName,
                FieldType.Table,
                new[] { tableValue }
            );
        }

        // 急ぎで用意したため実装内容がかなり怪しいですが期待通り動いているようです
        private static ResultsValue CreateResultsValue(int wordIndex, Document dom, Cell cell, PageLayout[] pages)
        {
            var words = new UiPath.DocumentProcessing.Contracts.Dom.Word[0];
            float ocr_confidence = 1.0f;
            Rectangle rect;
            // Azure Layout API 結果のセルに含まれる座標情報を DU が期待する領域クラスへマッピングしています。
            // 基準の支点となる最初の2つのパラメータへ -2 することで少し大きい領域を指定しています。
            // 幅、高さを指定する後の2つのパラメータでは、1.1 を掛けて少し大きな幅や高さになるように調整しています。
            rect = new Rectangle((Int32)cell.BoundingPolygon[0].X - 2,
                                 (Int32)cell.BoundingPolygon[0].Y - 2,
                                 (Int32)(Math.Abs(cell.BoundingPolygon[1].X - cell.BoundingPolygon[0].X) * 1.1),
                                 (Int32)(Math.Abs(cell.BoundingPolygon[2].Y - cell.BoundingPolygon[0].Y) * 1.1));

            // DOM 情報とセル情報で領域をマッチングし、該当する領域分の配列を生成しています。
            words = dom.Pages[cell.PageNumber - 1].Sections.SelectMany(s => s.WordGroups)
                .SelectMany(w => w.Words).Where(t => rect.Contains(new Rectangle((Int32)t.Box.Left, (Int32)t.Box.Top, (Int32)t.Box.Width, (Int32)t.Box.Height))).ToArray();

#if DEBUG
    Console.WriteLine($"CreateResultsValue : {words.Length} words found");
#endif
            List<Box> boxes = new List<Box>();
            List<ResultsValueTokens> tokens = new List<ResultsValueTokens>();
            foreach (var w in words)
            {
                boxes.Add(w.Box);
#if DEBUG
    Console.WriteLine($"CreateResultsValue : Box : {w.Box.Left}, {w.Box.Top}, {w.Box.Width}, {w.Box.Height}");
#endif
                ocr_confidence = Math.Min( ocr_confidence, w.OcrConfidence);
            }

#if DEBUG
    Console.WriteLine($"CreateResultsValue : {boxes.Count} boxes is found");
#endif
            if (boxes.Count == 0)
            {
                boxes.Add(Box.CreateChecked(0, 0, 0, 0));
            }
            // 抽出した値の領域への参照を設定しているようです。
            tokens.Add(new ResultsValueTokens(0, 0, cell.PageNumber - 1,
                                (float)dom.Pages[cell.PageNumber - 1].Size.Width,
                                (float)dom.Pages[cell.PageNumber - 1].Size.Height, boxes.ToArray()));
            var reference = new ResultsContentReference(cell.PageNumber, 0, tokens.ToArray());
            // 元々 DOM で設定されていた信頼度を Azure Layout API 結果として設定しています。
            return new ResultsValue(cell.Content, reference, (float)ocr_confidence, (float)ocr_confidence);
        }

        // 利用されていない（PDF 対応する際に活用できるかも知れません）
        /*
                private static double ConvertSize(double curX, double curWidth, double baseWidth)
                {
                    return curX / curWidth * baseWidth; 
                }
        */

        // Azure Layout API のレスポンスデータ（Cell配列）を2次元配列へ変換します
        public static Cell[,] ConvertTableTo2DArrayWithPosition(AnalyzeResult analyzeResult, int tableIndex)
        {
            var table = analyzeResult.Tables[tableIndex];
            int rows = table.RowCount;
            int columns = table.ColumnCount;

            Cell[,] tableArray = new Cell[rows, columns];

            // Azure Layout API 結果ではセルに位置情報や座標情報が含まれるため、後程、テーブル形式として制御がやり易くなるように2次元配列へ変更
            foreach (var cell in table.Cells)
            {
                    // セルに含まれる座標情報を取得（セルの中に複数含まれる座標情報を配列化しています）
                    var boundingPolygon = cell.BoundingRegions.SelectMany(br => br.BoundingPolygon).ToList();

                    // セルの各情報を対応する2次元配列へ格納します
                    tableArray[cell.RowIndex, cell.ColumnIndex] = new Cell(
                    cell.Content,
                    cell.RowIndex,
                    cell.ColumnIndex,
                    boundingPolygon,
                    rows,
                    columns,
                    cell.BoundingRegions[0].PageNumber  // PageNumber は後続処理で利用しているため、2次元配列へ含めています
                    );
            }
#if DEBUG
            PrintConvertedTable(tableArray);
#endif
            return tableArray;
        }

#if DEBUG
        // 2次元配列のCellデータのログを出力するデバッグ用メソッド
        private static void PrintConvertedTable(Cell[,] cell)
        {
            int rows = cell[0, 0].TableRows;
            int columns = cell[0, 0].TableColumns;

            Console.WriteLine($"PrintConvertedTable : Rows {rows}  Columns {columns}");

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (cell[i, j] != null && !cell[i, j].Content.Equals(""))
                    {
                        Console.WriteLine($"PrintConvertedTable Row: {i} Col: {j} Content: {cell[i, j].Content} X={cell[i, j].BoundingPolygon[0].X} Y={cell[i, j].BoundingPolygon[0].Y} page={cell[i, j].PageNumber}");
                    }
                    else
                    {
                        Console.WriteLine("[Empty] ");
                    }
                }
                Console.WriteLine(); // 次の行へ
            }
        }
#endif
    }
}
