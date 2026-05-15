---
name: generate-ddl
description: xing-project 専用のDDL生成・PR作成ワークフロー。develop 最新化 → feature/{issue-id}-generate-ddl ブランチ作成 → ddl_generator.py で指定テーブルのCREATE文を生成 → docs/db へ配置 → コミット&プッシュ → develop 宛 PR 作成までを一括実行する。「DDL生成」「テーブル定義書からDDLを作って」「generate-ddl」のような指示で起動する。
allowed-tools: Bash, Read, Glob, Edit, Write
argument-hint: "<issue-id> <table-id> [<table-id> ...]"
---

# DDL 生成 & PR 作成スキル

xing-project で「テーブル定義書 → DDL（MySQL）→ docs/db 配置 → develop への PR」までを一気通貫で実行する。

# 前提

- **必ず xing-project 上で実行する**。それ以外のディレクトリでは起動を拒否する。
- バンドルされたスクリプト: `{SKILL_DIR}/scripts/ddl_generator.py`
- DDL 出力先（スクリプト側）: `./output/CREATE文（テーブル単位）/*.sql`（カレントが xing-project ルートになる前提）
- 最終配置先: `docs/db/`
- 依存: Python 3 + openpyxl（プロジェクト共通環境）、`gh` CLI（PR 作成）

# context

引数={{args}}  
形式: `<issue-id> <table-id> [<table-id> ...]`  
例: `123 mclia mbshki rbshki`

# step

## step0: 引数チェック

1. `{{args}}` を空白で分割する。
2. 第1引数を `issue-id`、残りを `table-id-list`（半角スペース区切り）として扱う。
3. `issue-id` または `table-id-list` が空なら、使い方を提示して終了する:

   ```
   ❌ 引数不足です。
   使い方: generate-ddl <issue-id> <table-id> [<table-id> ...]
   例: generate-ddl 123 mclia mbshki rbshki
   ```

## step1: 実行場所の検証（xing-project 限定ガード）

1. カレントディレクトリ名を取得する:

   ```bash
   basename "$(pwd)"
   ```

2. `xing-project` でなければ即時終了する:

   ```
   ❌ このスキルは xing-project ディレクトリでのみ実行できます。
      現在のディレクトリ: <現在のパス>
      cd D:\workspace01\xing-project を行ってから再実行してください。
   ```

3. git リポジトリかどうか確認する:

   ```bash
   git rev-parse --is-inside-work-tree
   ```

   `true` でなければエラーで停止する。

## step2: 現在ブランチの記録

```bash
ORIG_BRANCH=$(git rev-parse --abbrev-ref HEAD)
echo "現在のブランチ: $ORIG_BRANCH"
```

`develop` だった場合は、後のリベース更新でなく `git pull --ff-only origin develop` のフローに分岐する。

## step3: develop ブランチを最新化（カレントブランチは維持）

「**現在のブランチを切り替えない**」が絶対条件。以下の順で処理する。

### 3-1: リモートから fetch

```bash
git fetch origin develop
```

ネットワーク/権限エラーなら、その内容をそのまま表示して終了する。

### 3-2: ローカル develop の有無で分岐

```bash
git show-ref --verify --quiet refs/heads/develop
```

- **終了コード 0（ローカル develop が存在）**:
  - 現在ブランチが `develop` の場合:

    ```bash
    git pull --ff-only origin develop
    ```

  - 現在ブランチが `develop` 以外の場合（チェックアウト不要で fast-forward 更新）:

    ```bash
    git fetch origin develop:develop
    ```

    > `refs/heads/develop:refs/heads/develop` 形式の fetch は **fast-forward のみ許可**。非 FF（ローカル develop が独自にコミットを持っている等）の場合は失敗する。

- **終了コード非 0（ローカル develop が存在しない）**:

  ```bash
  git fetch origin develop:develop
  ```

  これでリモートを基にローカル develop が新規作成される。現在のブランチはそのまま維持される。

### 3-3: 失敗時の取り扱い

`git fetch origin develop:develop` または `git pull --ff-only` が失敗した場合、典型的には以下のいずれか:

- 非 fast-forward（ローカル develop が枝分かれしている）
- コンフリクト
- 認証エラー
- リモート develop が存在しない

エラーメッセージをそのまま表示し、**処理を中止する**（ブランチ作成も DDL 生成も行わない）。ユーザーに復旧手順（手動で develop を整理する等）を案内する:

```
❌ develop ブランチの更新に失敗しました。
  原因: <git のエラー出力>
  対応: ローカル develop の状態を確認し、コンフリクト解消後に再実行してください。
        例) git switch develop && git pull --rebase origin develop && git switch -
```

## step4: feature ブランチを作成

ブランチ名: `feature/<issue-id>-generate-ddl`

```bash
NEW_BRANCH="feature/<issue-id>-generate-ddl"
```

### 4-1: 既存ブランチ衝突チェック（推奨方針: 既存ならエラー停止）

```bash
git show-ref --verify --quiet "refs/heads/$NEW_BRANCH"
```

- **終了コード 0（既に存在）** → エラー停止:

  ```
  ❌ ブランチ '$NEW_BRANCH' は既に存在します。
     既存ブランチを削除/リネームしてから再実行してください:
     例) git branch -D $NEW_BRANCH
  ```

- **終了コード非 0** → 作成へ進む。

### 4-2: develop を起点にブランチ作成 → 切替

```bash
git switch -c "$NEW_BRANCH" develop
```

> `git checkout -b ... develop` 相当。失敗した場合（develop が見つからない等）はエラー停止。

## step5: DDL 生成スクリプトを実行

バンドルされた Python スクリプトをカレント（xing-project ルート）から呼び出す。

```bash
python "<SKILL_DIR>/scripts/ddl_generator.py" -n <table-id-1> [<table-id-2> ...] -o "./output"
```

> **注意**:
> - `<SKILL_DIR>` は `$CLAUDE_PROJECT_DIR` ではなく、本 SKILL.md が置かれているディレクトリ。実行時には絶対パスに解決して指定する（例: `D:/Users/XRMTUser13/.claude/skills/generate-ddl/scripts/ddl_generator.py`）。
> - ユーザー指定の起動コマンド書式 `python ddl_generator.py -n {table id list} -o ".\output"` を厳守する（コマンド名は `python`）。
> - `-y` を付けないと上書き確認プロンプトが出る場合がある。自動化のため `-y` を末尾に追加してよい（スクリプトの仕様に依存。エラーなら外して再実行）。
> - 標準出力・標準エラーの内容はユーザーに提示する。

実行後、`./output/CREATE文（テーブル単位）/` 配下に `<テーブルID>_*.sql` が出力されることを期待する。

### 5-1: 生成失敗時

スクリプトが非 0 終了、または `output/CREATE文（テーブル単位）/` が空、もしくは指定テーブル数より少ない場合は、エラーメッセージとともに **その時点で停止する**（コミット・PR は作らない）。

```
❌ DDL 生成に失敗しました。
   <python stderr>
   生成済みファイル: <あれば列挙>
```

## step6: 生成 SQL を docs/db へ配置

```bash
SRC_DIR="output/CREATE文（テーブル単位）"
DST_DIR="docs/db"

mkdir -p "$DST_DIR"
cp -f "$SRC_DIR"/*.sql "$DST_DIR/"
```

- 同名ファイルは **上書き**（`cp -f`）。
- 配置後、`docs/db/` 配下に追加/変更された SQL の一覧を表示する。

## step7: コミット & プッシュ

### 7-1: ステージング

`docs/db/` 配下の SQL のみを対象にする（`output/` は対象外）。

```bash
git add docs/db/*.sql
```

### 7-2: コミットメッセージ

```
feat: テーブル定義書からDDLを生成 (テーブル: <table-id-1> <table-id-2> ...) #<issue-id>
```

HEREDOC で渡す:

```bash
git commit -m "$(cat <<'EOF'
feat: テーブル定義書からDDLを生成 (テーブル: <tables>) #<issue-id>

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

`<tables>` `<issue-id>` は実値に展開してから渡すこと。

### 7-3: プッシュ

```bash
git push -u origin "$NEW_BRANCH"
```

認証エラー等はそのまま表示し、その時点で停止する（PR 作成は行わない）。

## step8: develop 宛 PR を作成

`gh` CLI を使用する。

```bash
gh pr create \
  --base develop \
  --head "$NEW_BRANCH" \
  --title "feat: テーブル定義書からDDLを生成 (テーブル: <tables>) #<issue-id>" \
  --body "$(cat <<'EOF'
## Summary
- Issue #<issue-id> 対応として、テーブル定義書から DDL（CREATE 文）を自動生成
- 対象テーブル: <tables>
- 生成された SQL を `docs/db/` に配置

## 生成ファイル
<ここに docs/db 配下に追加/更新されたファイルを箇条書き>

## Test plan
- [ ] `docs/db/` に対象テーブルの SQL が揃っていることを確認
- [ ] `generate-base-code` で Entity / BaseMapper / Mapper XML が生成できることを確認
- [ ] `cd backend && mvn clean compile` が通ること

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

`<issue-id>` `<tables>` は実値に展開する。`gh pr create` が返した URL は最後にユーザーへ提示する。

## step9: 完了レポート

以下のサマリを 1 通だけ出力する:

```
✅ DDL 生成 & PR 作成が完了しました。
  - ブランチ: feature/<issue-id>-generate-ddl
  - 生成テーブル: <tables>
  - 配置先: docs/db/
  - PR: <URL>
  - 現在ブランチ: feature/<issue-id>-generate-ddl
```

# rules

## ガード（絶対遵守）

- **カレントディレクトリ名が `xing-project` でなければ即時中止**。`cd` 等で自動移動しない。
- **`git fetch origin develop:develop` / `git pull --ff-only` が失敗したら以降の処理を一切実行しない**（コンフリクト時は手動解決を促す）。
- **`develop` ブランチを破壊するような操作（`git reset --hard` / `git push --force` 等）は絶対に行わない**。
- **feature ブランチが既に存在する場合はエラー停止**。ユーザーの明示指示がない限り `-D` で消さない。
- **`output/` ディレクトリ自体は git に含めない**（コミット対象は `docs/db/*.sql` のみ）。

## 安全策

- DDL 生成スクリプトの呼び出し前に、テーブル ID リスト・出力先・実行コマンドをユーザーに 1 度提示すること。
- スクリプト失敗、SQL 0 件、`gh pr create` 失敗のいずれも **段階で停止** し、後続を実行しない。
- コミット時に `git add -A` / `git add .` を使わない（Box 同期ファイル等を巻き込まないため）。`docs/db/*.sql` のみを明示する。
- PowerShell / bash どちらから呼ばれてもよいように、`<SKILL_DIR>` のパス区切りはフォワードスラッシュ（`/`）で扱う。

## エラー時のメッセージ整形

すべてのエラー出力は以下のフォーマットで統一する:

```
❌ <短い概要>
   原因: <git/python/gh の生のエラーメッセージ>
   対応: <ユーザーが取るべきアクション>
```

# 補足

- `ddl_generator.py` の詳細は `<SKILL_DIR>/scripts/ddl_generator.py` を参照（Excel テーブル定義書を読み込み、MySQL CREATE 文を出力する）。
- 本スキルは「全件削除→再投入」のような破壊的 SQL 実行は行わない（DDL ファイル生成のみ）。実際の `CREATE/DROP` をDBへ反映する場合は別途 `generate-base-code` スキル等を経由すること。
