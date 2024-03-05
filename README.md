# Document-Processing-Code-Samples
Code samples for document processing activities.

Activies are located in Samples/SampleActivities/Basic/

Visual Studio 2022 インストール(.NET SDK 6.0 が必須)

---

新たに表抽出を行うカスタムアクティビティとして [Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs) を追加しました。
このアクティビティは Azure AI Document Intelligence が提供する Layout API を利用しています。

現在は `請求書_1-0.PNG` を利用する前提の実装となっています。
そのため、帳票の表部分は5行である前提、かつ列としては Description、UnitPrice、Quantity、Amount のみ対応しております。

また、AzureLayout.cs を呼び出す際に事前に DU の OCR で読み取ったデータを引き渡すと Layout API 結果（列フィールド）と DU OCR 結果（標準フィールド）をマージする実装となっています。
なお、標準フィールドのマージ処理は暫定的に VendorName、CustomerName、InvoiceId、InvoiceTotal の4項目のテキストデータのみ対応しております。（座標データ、信頼度情報は対応しておりません）


A new custom activity, [Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs), has been added to perform table extraction.
This activity uses the Layout API provided by Azure AI Document Intelligence.

Currently, the implementation assumes that `請求書_1-0.PNG` is used.
Therefore, the table part of the form is assumed to have 5 rows, and only Description, UnitPrice, Quantity, and Amount are supported as columns.

In addition, when calling AzureLayout.cs, if the data read by DU OCR is handed over in advance, the Layout API results (column fields) and DU OCR results (standard fields) are merged.
Note that the standard field merging process tentatively supports only the four text data items of VendorName, CustomerName, InvoiceId, and InvoiceTotal. (Bounding Polygon and Confidence information are not supported.)
