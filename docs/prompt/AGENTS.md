# Majin_Prompt_v3 Agent Guide

## プロジェクト概要
- Google Apps Script を `clasp` で管理するリポジトリです。
- 主なソースは `コード.js` と `appsscript.json` にあります。

## 開発環境
- Node.js と `clasp` CLI を利用して Google Apps Script にデプロイします。
- `.clasp.json` はリポジトリ外または個別に管理される前提です。
- コード編集はこのリポジトリ内で行い、`clasp push/pull` で Apps Script と同期します。

## 基本フロー
1. 作業前に `clasp pull` で最新状態を取得します。
2. ローカルで `コード.js` や関連ファイルを編集します。
3. 単体テストなどローカルで可能な確認を行います。
4. 変更内容を `clasp push` で Apps Script 側へ反映します。
5. 必要に応じて Git でコミットし、リモートリポジトリへ反映します。

## 制約条件
- コーディング規約やフォーマットに従い、一貫性を保ちます。
- Google Apps Script 特有の制限（実行時間、クォータ）を意識します。
- 共有アカウントや認証情報はハードコードしないでください。
- 作業を完了したら Apps Script 側へ `clasp push` で同期すること。
- 開発が完了した場合でも、Git への push を行う前に必ずユーザーに「push してよいか」確認します。

## ロール
- **Developer Agent**: コーディング、バグ修正、テスト。
- **Reviewer Agent**: 変更内容の確認と品質保証。
- **Ops Agent**: `clasp` を使ったデプロイや環境設定の支援。

## コミュニケーション
- 進捗や課題はチケットまたはチャットで共有します。
- 仕様不明点がある場合は必ず確認を行い、仮定で進めないようにします。
