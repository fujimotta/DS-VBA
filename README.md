# DSWorksheet / DSLookup

> **「列番号を書くのをやめた日から、VBAは壊れなくなった。」**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel VBA](https://img.shields.io/badge/Excel-VBA-green.svg)]()

---

## これは何か

**DSWorksheet** と **DSLookup** は、Excel VBAでのデータ操作を根本から変える2つのクラスモジュールです。

| モジュール | 役割 |
|---|---|
| **DSWorksheet** | データを「安全に・壊さず」読み書きする土台 |
| **DSLookup** | その土台から「素早く・正確に」データを取り出す検索エンジン |

この2つをセットで使うことで、**「誰が書いても壊れない・誰が読んでも分かる」** VBAコードが実現します。

---

## なぜ作ったか

あるとき、前任者から引き継いだVBAコードを触った瞬間、私はその現実を目の当たりにしました。

変数名は謎、ループは深さ不明、Excelの標準機能で済むソートをわざわざバブルソートで再実装——その結果、データの順番がぐちゃぐちゃになっていました。善意で書かれたコードが、誰も保守できない魔窟になっていたのです。

**私はVBAに嫌気がさしました。**

しかしその後、ふと気づきました。「嫌いなのはVBAではなく、保守できないコードだ」と。

かつて業務で触れたデータベースフレームワークの設計思想——CRUDを安全かつシンプルに扱う仕組み——をVBAで再現できないか。そこから生まれたのがこの2つのモジュールです。

---

## できること

### DSWorksheet

**1. 列を「番号」ではなく「項目名」で操作する**

```vba
' 従来の書き方（列を1本追加したら全部崩れる）
Cells(i, 3).Value = "example"

' DSWorksheetの書き方（何列追加しても崩れない）
ds.Value("メールアドレス") = "example"
```

**2. CSVを「全列テキスト」として安全に読み込む**

Excelの自動変換（0落ち・日付化・指数表記）を物理レベルで遮断します。

```vba
' 「00123」が「123」に化ける問題が起きない
Dim ds As New DSWorksheet
Call ds.Init("C:\data\input.csv")
```

**3. CRUDがシンプルに書ける**

```vba
' 読み取り（イテレーション）
Call ds.MoveReset
Do While ds.MoveNext
    Debug.Print ds.Value("商品コード"), ds.Value("数量")
Loop

' 書き込み
ds.Value("ステータス") = "完了"

' 追加
ds.Add("商品コード") = "A001"
ds.Add("数量") = 10
Call ds.AddNextRow

' 削除（遅延一括削除でインデックスずれなし）
Call ds.Delete
Call ds.DeleteCommit
```

---

### DSLookup

**1. 条件を重ねて高速フィルタリング**

```vba
Dim lk As New DSLookup
Call lk.Init(ds)

' AND条件を積み重ねる
Call lk.AddCondition("ステータス", "未処理")
Call lk.AddCondition("担当者", "山田")

Do While lk.MoveNext
    Debug.Print ds.Value("商品コード")
Loop
```

**2. Matchを使った高速単一検索**

```vba
' 商品コードからメールアドレスを取得
Dim email As String
email = lk.Lookup("商品コード", "A001", "メールアドレス")
```

---

## クイックスタート

最小構成のサンプルです。5分で動かせます。

```vba
Sub Sample()
    ' 1. 初期化
    Dim ds As New DSWorksheet
    Call ds.Init("C:\data\顧客リスト.xlsx", "Sheet1")

    ' 2. 全件ループして表示
    Call ds.MoveReset
    Do While ds.MoveNext
        Debug.Print ds.Value("氏名") & " / " & ds.Value("メールアドレス")
    Loop

    ' 3. 後片付け
    Call ds.Cleanup
End Sub
```

---

## インストール方法

1. このリポジトリから `DSWorksheet.cls` と `DSLookup.cls` をダウンロード
2. VBAエディタ（Alt + F11）を開く
3. メニューから「ファイル」→「ファイルのインポート」を選択
4. 2つのファイルを順番にインポート
5. 完了

**動作環境：** Excel 2016以降 / Windows（Mac環境では一部制限あり）

---

## 導入事例

某製造業における業務改善プロジェクトにて導入。
「安定性・保守性の高さ」が現場から評価され、導入後に**業務改善専任部門の新設**につながった。
その後、契約満了まで継続稼働。

---

## 作者

**ふじもった**
VBAの現場で「保守できないコードの地獄」を経験し、このモジュールを開発。
「簡単であることは、究極の安全である」をコンセプトに、データを壊さないVBA設計を追求しています。

- 🎥 YouTube：[ふじもったの教室](#) ← チャンネルURL（開設後に更新）
- 📩 お問い合わせ・導入支援：概要欄のリンクからご確認ください

---

## ライセンス

[MIT License](https://opensource.org/licenses/MIT)

自由に使用・改変・再配布できます。商用利用も可能です。
