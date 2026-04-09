# 案件実績CSV自動ダウンロード

社内採用管理システムから複数条件で案件データをCSV取得し、指定フォルダに自動保存するスクリプトです。
タスクスケジューラと組み合わせることで完全無人化が可能です。

## 機能

- Seleniumによるブラウザ自動操作（ログイン〜CSVダウンロードまで）
- 15種類の条件（雇用形態・ステータス・属性）で案件データを自動取得
- ダウンロード完了を検知してファイルを自動リネーム・保存
- 実行ログの自動出力

## 効果

| 項目 | 改善前 | 改善後 |
|------|--------|--------|
| 作業時間 | 約3時間（手動） | 約30分（自動） |
| 作業方式 | 毎回手動操作 | タスクスケジューラで無人実行 |

## 必要環境

- Python 3.x
- Google Chrome
- 以下のパッケージ

```
pip install selenium webdriver-manager python-dotenv
```

## セットアップ

1. `.env.example` をコピーして `.env` を作成
2. `.env` に認証情報・パスを設定

```
LOGIN_ID=your_login_id
LOGIN_PASSWORD=your_password
DOWNLOAD_FOLDER=\\your_server\path\to\folder
LOG_FILE=C:\Users\yourname\Desktop\実行ログ.txt
```

3. スクリプトを実行

```
python anken_jisseki.py
```

## ダウンロードされるCSV一覧

雇用形態（dispatch / referral / scheduled_temp / daily_referral）とステータス（recruiting / completed / lost_order など）の組み合わせで15パターンのCSVを取得します。
ファイル名・条件は `DOWNLOAD_TARGETS` リストで自由にカスタマイズ可能です。
