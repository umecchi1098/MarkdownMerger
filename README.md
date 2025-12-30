# Markdown Merger

Markdown / ZIP / フォルダ内のMarkdownをまとめて1ファイルに結合する小さなCLIツールです。ドラッグ＆ドロップでも実行できます。

## 主な機能
- Markdown / ZIP / フォルダを混在入力可能（ZIP内は並べ替えのオン/オフ指定可）
- 元ファイル名コメントを付与して結合（`--no-source` で無効化）
- UTF-8優先＋失敗時はShift_JISで再読込、改行はLFへ正規化
- 出力先自動決定（重複回避付き）または `-o` で明示指定
- 進捗スピナー表示、Explorer経由実行時は終了前にEnter待ち

## 使い方
```
MarkdownMerger.exe <入力...> [-o <出力.md>] [--no-source] [--no-sort-zip]
MarkdownMerger.exe -i <入力...> [-o <出力.md>] [--no-source] [--no-sort-zip]
```
- 入力: Markdown / ZIP / フォルダを複数指定できます。ドラッグ＆ドロップでもOKです。
- 出力: `-o` 省略時は最初の入力と同じフォルダに `merged.md` などの自動パスで保存します（重複時は連番付与）。
- `--no-source`: `<!-- source: ... -->` コメントを出力しません。
- `--no-sort-zip`: ZIP内の順序をエントリ順のままにします（既定は名前順ソート）。

### 実行例
```
dotnet run --project MarkdownMerger/MarkdownMerger.csproj -- docs/notes README.md archive.zip -o combined.md
```

## ビルド
- 前提: .NET 10 SDK
- 開発: `dotnet build` または `dotnet run --project MarkdownMerger/MarkdownMerger.csproj -- <args>`
- 発行（単一ファイル例）: `dotnet publish -c Release -r win-x64 /p:PublishSingleFile=true`

## メモ
- エンコードはUTF-8を優先し、失敗時はShift_JISで再試行します。
- 改行はLFに正規化し、ファイル間には区切り行 `---` を挿入します。
- ZIPは `__MACOSX/` を除外し、Markdown拡張子のみ抽出します。
- Explorerから起動した場合は処理後にEnter入力待ちになります。

## 名称
- 正式名称: Markdown Merger
- 実行ファイル名: `MarkdownMerger.exe`
