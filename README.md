# excel2csv

Excel ブックを 1 つまたは複数ドラッグ＆ドロップして、同じ列名のシートを縦結合し、単一の CSV または CSV.GZ に変換するための Docker 封じ込め型ツールです。

## できること

- 複数ブック、複数シートを 1 本のファイルへ縦結合します。
- すべての読み取り対象シートの列名と列順が同じであることを検証します。
- `.csv` は BOM 付き UTF-8、`.csv.gz` は BOM なし UTF-8 で出力します。
- Office パスワード付きブックは `msoffcrypto-tool` で復号して読み込みます。
- Python 依存関係は Docker イメージ内に閉じ込めます。

## Windows でドラッグ＆ドロップ

1. Docker Desktop を起動します。Docker Desktop が使えない場合、`wsl docker version` が通る WSL 側 Docker に自動 fallback します。
2. CSV を作る場合は `excel2csv.cmd` に Excel ファイルまたはフォルダをドロップします。
3. CSV.GZ を作る場合は `excel2csv-gzip.cmd` にドロップします。
4. パスワード付きブックが見つかった場合は、表示されるプロンプトにパスワードを入力します。

出力先を指定しない場合、最初にドロップしたファイルと同じフォルダに `merged_YYYYmmdd_HHMMSS.csv` または `merged_YYYYmmdd_HHMMSS.csv.gz` を作ります。

## Windows の「送る」メニューから実行

`install-sendto.cmd` を実行すると、現在のユーザーの「送る」メニューに次の 2 つを登録します。管理者権限は不要です。

- `excel2csv CSV`
- `excel2csv CSV.GZ`

Excel ファイルまたはフォルダを選択して右クリックし、`送る` からどちらかを選ぶと変換できます。登録を削除する場合は `uninstall-sendto.cmd` を実行します。

Windows の「共有先」は通常の `.cmd` や PowerShell だけでは受け口として登録できません。対応する場合は、Share Target として登録する MSIX/UWP/WinUI の小さなアプリを別途作る必要があります。

## WSL / Linux から実行

```bash
scripts/excel2csv.sh -o out/merged.csv data/a.xlsx data/b.xlsx
scripts/excel2csv.sh --gzip -o out/merged.csv.gz data/
scripts/excel2csv.sh --ask-password -o out/merged.csv data/protected.xlsx
EXCEL2CSV_PASSWORD='secret' scripts/excel2csv.sh -o out/merged.csv data/protected.xlsx
```

初回実行時に `excel2csv:local` Docker イメージを build します。

## Docker CLI を直接使う

```bash
docker build -t excel2csv:local .
docker run --rm \
  -v "$PWD/data:/data:ro" \
  -v "$PWD/out:/out" \
  excel2csv:local \
  /data/a.xlsx /data/b.xlsx \
  -o /out/merged.csv
```

gzip 出力:

```bash
docker run --rm \
  -v "$PWD/data:/data:ro" \
  -v "$PWD/out:/out" \
  excel2csv:local \
  /data \
  -o /out/merged.csv.gz
```

## CLI オプション

```bash
excel2csv INPUT... -o OUTPUT [--gzip] [--ask-password] [--password PASSWORD] [--password-file FILE] [--sheet NAME] [--recursive]
```

- `INPUT` は Excel ファイルまたはフォルダです。
- フォルダ指定時は既定で直下のみ読みます。配下も読む場合は `--recursive` を付けます。
- `--sheet` は複数回指定できます。省略時は全シートを読みます。
- `--ask-password` は暗号化ブックを検出した時点で対話的にパスワードを要求します。
- `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb` を対象にします。
- 暗号化された複数ファイルは、現時点では同じパスワードを共有している前提です。

## テスト

```bash
docker build -t excel2csv:local .
docker run --rm --entrypoint pytest excel2csv:local
```

## Git remote

この作業ディレクトリを Git リポジトリ化する場合の remote:

```bash
git remote add origin https://github.com/htanaka1234/excel2csv.git
```
