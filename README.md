# sme-automation

**中小企業向け業務自動化スクリプト集**  
エンジニアでない経営者が [Claude Code](https://claude.ai/code) と対話しながら開発した業務自動化ツール群です。

> 「コードが書けない社長が、Claude Codeで会社の朝を変えた実録」

---

## 概要

社員約30名のITサービス会社（SES業）で、以下のSaaSを日常業務で使用しています：

- **サイボウズ Office**（スケジュール・掲示板・ワークフロー・DB）
- **MFクラウド勤怠**（勤怠管理）
- **Chatwork**（社内チャット）
- **Outlook**（メール）
- **NAS（TeraStation）**（社内ファイルサーバー）

これらをAPIやSeleniumで連携し、毎朝の情報収集・レポーティングを自動化しています。

---

## スクリプト一覧

### 🌅 朝の自動レポート

| スクリプト | 説明 |
|-----------|------|
| `morning_report.py` | 毎朝7時にメール・Chatworkの動向をChatworkマイチャットへ自動投稿 |

### 📅 サイボウズ連携

| スクリプト | 説明 |
|-----------|------|
| `cybozu_schedule.py` | スケジュール・ワークフロー・掲示板・ファイル管理を一括チェック |
| `cybozu_trouble_check.py` | トラブル管理DB（特定のdid）のオープン件数・状況を確認 |
| `cybozu_leave_report.py` | 有給取得状況のレポート生成 |

### 📧 メール・チャット

| スクリプト | 説明 |
|-----------|------|
| `outlook_check.py` | Outlookの未読メール・要注意メールを自動チェック |
| `outlook_sales_analysis.py` | 直近1ヶ月のメール傾向分析 |
| `chatwork_check.py` | Chatwork全ルームの未読・要注意メッセージ分析 |
| `chatwork_weekly_report.py` | 週次レポートをChatworkへ投稿 |
| `chatwork_notify.py` | 汎用Chatwork通知ユーティリティ |

### ✈️ 出張・旅程管理

| スクリプト | 説明 |
|-----------|------|
| `check_travel.py` | ANA・Hilton予約を取得しサイボウズスケジュールと照合・同期 |
| `fix_cybozu_travel.py` | 既知の出張予約データをサイボウズに手動登録 |

### 💰 勤怠・給与

| スクリプト | 説明 |
|-----------|------|
| `mf_attendance.py` | MFクラウド勤怠データをエクスポートしChatworkへ投稿 |
| `mf_payroll.py` | MFクラウド給与データの取得・集計 |

### 📊 財務分析

| スクリプト | 説明 |
|-----------|------|
| `read_cf.py` / `read_cf2.py` | NASのキャッシュフローExcelを読み込み分析 |
| `update_cf.py` | キャッシュフローファイルの更新 |
| `read_urikake.py` | 売掛金データの読み込み |

### 🔧 ユーティリティ

| スクリプト | 説明 |
|-----------|------|
| `save_credential.py` | Windows Credential Managerへの認証情報登録 |

---

## 動作環境

- **OS**: Windows 11
- **Python**: 3.10以上
- **ブラウザ**: Microsoft Edge（Selenium headlessモード使用）

### 主な依存ライブラリ

```
selenium
openpyxl
win32com.client  # pywin32（Outlook連携）
```

---

## セットアップ

### 1. 認証情報の登録

本スクリプト群は機密情報を**Windows Credential Manager**で管理します。コード内にパスワードを直接書きません。

```powershell
# 例：Chatwork APIトークンを登録
python save_credential.py
```

使用するCredential Manager キー：

| キー名 | 用途 |
|--------|------|
| `Chatwork_API_Token` | Chatwork API |
| `Cybozu_nakkou` | サイボウズOfficeログイン |
| `ANA_MB` | ANAマイレージクラブ |
| `Hilton_HH` | Hilton Honorsログイン |
| `MF_nakkou` | MFクラウドログイン |

### 2. 設定値の変更

各スクリプトの先頭に設定値があります。自社環境に合わせて変更してください。

```python
# 例：morning_report.py
SCRIPTS_DIR = r'C:\Users\yourname\Scripts'
CW_ROOM_ID  = 'YOUR_CHATWORK_ROOM_ID'   # マイチャットのルームID
```

### 3. 朝の自動レポート設定（Task Scheduler）

```powershell
# 毎朝7時に自動実行するタスクを登録
$action  = New-ScheduledTaskAction -Execute "python" -Argument "-X utf8 C:\Users\yourname\Scripts\morning_report.py"
$trigger = New-ScheduledTaskTrigger -Daily -At "07:00"
Register-ScheduledTask -TaskName "MorningReport" -Action $action -Trigger $trigger -RunLevel Highest
```

---

## セキュリティについて

- パスワード・APIトークン等はすべてWindows Credential Manager管理
- NASのIPアドレス・ルームID等の固有情報はサンプル値に変更して公開
- ブラウザ操作はすべてheadlessモード（画面非表示）

---

## 背景・開発経緯

このリポジトリは、エンジニアではない経営者が [Claude Code](https://claude.ai/code) との対話のみでゼロから開発した業務自動化スクリプト集です。

「AIに仕様を日本語で伝える → コードが生成される → 動作確認する」というサイクルを繰り返し、約2週間で10本以上のスクリプトを整備しました。

導入の実録記事を Zenn・note で公開予定です。

---

## ライセンス

MIT License

---

## 作者

**中村 光（Kou Nakamura）**  
株式会社ブレインディレクション 代表取締役  
GitHub: [@KouNakamura](https://github.com/KouNakamura)
