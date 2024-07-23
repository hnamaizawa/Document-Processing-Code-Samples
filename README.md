# Document-Processing-Code-Samples
Code samples for document processing activities.

Activies are located in Samples/SampleActivities/Basic/

Visual Studio 2022 インストール(.NET SDK 6.0 が必須)

---

## 謝辞

このプロジェクトは、多くの人々の支援と励ましによって実現しました。特に、以下の方々に深い感謝を表します。

- **Kim**さんには、参考にさせていただいた[プロジェクト](https://github.com/javaos74/Document-Processing-Code-Samples)に関する様々な情報をご提供いただきました。ありがとうございました。

また、このプロジェクトに貢献してくださったすべての協力者に心から感謝いたします。

This project was made possible by the support and encouragement of many people. In particular, we would like to express our deepest gratitude to the following people

- **Kim**-san for providing us with various information about the [project](https://github.com/javaos74/Document-Processing-Code-Samples) we are referencing. Thank you very much.

We would also like to extend our sincere thanks to all the contributors who have worked on this project.

---

## 最新の更新情報

- **バージョン 1.0.8 のリリース (2024-07-23)**: AI inside DX Suite OCR Engine アクティビティを追加しました。DX Suite の全文OCR機能を用いたアクティビティが利用可能となりました。ドキュメントをデジタル化する際にご利用ください詳細は[変更ログ](https://github.com/hnamaizawa/Document-Processing-Code-Samples/commit/4c66bce8704612991fa87ada1a17996bd3e09bda)をご覧ください。

- **バージョン 1.0.7 のリリース (2024-05-28)**: Cogent Labs SmartRead OCR Engine アクティビティを追加しました。SmartRead の全文OCR機能を用いたアクティビティが利用可能となりました。ドキュメントをデジタル化する際にご利用ください詳細は[変更ログ](https://github.com/hnamaizawa/Document-Processing-Code-Samples/commit/53f60a7ad9c5e5569a3ee523f5070efaed3324eb)をご覧ください。

- **バージョン 1.0.6 のリリース (2024-03-26)**: AzureLayout.cs を全面的に見直し、DU の抽出結果から位置情報や信頼度のマージへ対応し、Azure Layout API 抽出結果のマッピング処理改善、対象となる明細テーブル指定機能の追加などを実施しました。詳細は[変更ログ](https://github.com/hnamaizawa/Document-Processing-Code-Samples/commit/6b701e6b89fdc2233dcf4496f6c6d0fff63ebbcb)をご覧ください。

- **バージョン 1.0.5 のリリース (2024-03-19)**: Clova OCR Engine アクティビティのデフォルトの言語設定を **ja** に変更しました。詳細は[変更ログ](https://github.com/hnamaizawa/Document-Processing-Code-Samples/commit/64f4d4211d1bc2e177a36957bb17d33c9735e81d)をご覧ください。

- **表抽出アクティビティの追加 (2024-03-04)**:  Azure Layout API を用いて表抽出を行うアクティビティを追加しました。詳細は[説明](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/README.md#%E8%A1%A8%E6%8A%BD%E5%87%BA%E3%82%A2%E3%82%AF%E3%83%86%E3%82%A3%E3%83%93%E3%83%86%E3%82%A3%E3%81%AE%E8%AA%AC%E6%98%8E)をご覧ください。

---

## 表抽出アクティビティの説明

新たに表抽出を行うカスタムアクティビティとして [Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs) を追加しました。
このアクティビティは Azure AI Document Intelligence が提供する Layout API を利用しています。

テスト用データとして `請求書_1-0.PNG` を含んでおります。  
実際に読み取りたい帳票の明細データに合わせて[こちら](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs#L132-L140)のロジックで Items に含まれる項目の順番を調整する必要があります。  
（Azure Layout API 結果の表データの列の順番に従い Items に定義した項目をマッピングする実装となっています）

例としては、 `請求書_1-0.PNG` の明細データの列は、商品名（Description）、単価（UnitPrice）、数量（Quantity）、金額（Amount）の順番となっていますが、上記の Items 内の項目も同じ順番で定義されております。

また、AzureLayout.cs を呼び出す際に事前に DU の抽出器で読み取ったデータを引き渡すと Layout API 結果（列フィールド）と DU の抽出結果（標準フィールド）をマージする実装となっています。  


A new custom activity, [Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs), has been added to perform table extraction.
This activity uses the Layout API provided by Azure AI Document Intelligence.  
  
`請求書_1-0.PNG` is included as test data. You will need to adjust the order of the items in Items using the logic [here](https://github.com/hnamaizawa/Document-Processing-Code-Samples/blob/master/Samples/SampleActivities/Basic/DataExtraction/AzureLayout.cs#L132-L140) according to the detail data of the form you actually want to read.  
(The implementation is to map the items defined in Items according to the order of the columns of the table data in the Azure Layout API results.)

For example, the columns of detail data in `請求書_1-0.PNG` are in the order Product Name (Description), Unit Price (UnitPrice), Quantity (Quantity), and Amount (Amount), and the columns in the Items above are defined in the same order.

In addition, when calling AzureLayout.cs, if the data read by DU Extrator is handed over in advance, the Layout API results (column fields) and DU Extractor results (standard fields) are merged.
