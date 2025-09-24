# GEMINI.md

## Table of Contents

*   [Project Overview](#project-overview)
*   [Building and Running](#building-and-running)
*   [Getting Started](#getting-started)
*   [Architecture](#architecture)
*   [Development Conventions](#development-conventions)
*   [Contributing](#contributing)

## Project Overview

This is a Google Apps Script project for automatically generating Google Slides presentations. The project is managed using `clasp`, a command-line tool for developing Google Apps Script projects locally.

The main logic is contained in the `コード.js` file. This script defines a `slideData` object, which is an array of JavaScript objects, each representing a slide. The script then iterates over this array and generates a Google Slides presentation based on the data.

The project also includes a `system prompt.md` file, which contains a detailed prompt for a Gemini agent. This agent is designed to generate the `slideData` object from unstructured text, making it easy to create presentations from notes, articles, or other documents.

The `AGENTS.md` file outlines the roles and responsibilities of different AI agents involved in the project, including a Developer Agent, a Reviewer Agent, and an Ops Agent.

The `scripts` directory contains helper scripts for image analysis and inspection, which can be used to process images for the presentations.

### プロジェクト概要 (Japanese)

これは、Googleスライドのプレゼンテーションを自動生成するためのGoogle Apps Scriptプロジェクトです。このプロジェクトは、Google Apps Scriptプロジェクトをローカルで開発するためのコマンドラインツールである`clasp`を使用して管理されます。

主なロジックは`コード.js`ファイルに含まれています。このスクリプトは、各スライドを表すJavaScriptオブジェクトの配列である`slideData`オブジェクトを定義します。次に、スクリプトはこの配列を繰り返し処理し、データに基づいてGoogleスライドのプレゼンテーションを生成します。

このプロジェクトには、Geminiエージェントの詳細なプロンプトを含む`system prompt.md`ファイルも含まれています。このエージェントは、非構造化テキストから`slideData`オブジェクトを生成するように設計されており、メモ、記事、その他のドキュメントからプレゼンテーションを簡単に作成できます。

`AGENTS.md`ファイルには、開発者エージェント、レビュー担当者エージェント、運用担当者エージェントなど、プロジェクトに関与するさまざまなAIエージェントの役割と責任の概要が記載されています。

`scripts`ディレクトリには、プレゼンテーション用の画像を処理するために使用できる画像分析および検査用のヘルパースクリプトが含まれています。

## Building and Running

To use this project, you need to have Node.js and the `clasp` CLI installed.

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/Apple_Majin.git
    ```
2.  **Install dependencies:**
    ```bash
    npm install
    ```
3.  **Log in to `clasp`:**
    ```bash
    clasp login
    ```
4.  **Create a new Google Apps Script project or use an existing one.**
5.  **Update the `.clasp.json` file with your script ID.**
6.  **Push the code to your Google Apps Script project:**
    ```bash
    clasp push
    ```
7.  **Open the Google Slides presentation and run the `generatePresentation` function from the "カスタム設定" menu.**

### ビルドと実行 (Japanese)

このプロジェクトを使用するには、Node.jsと`clasp` CLIがインストールされている必要があります。

1.  **リポジトリをクローンします:**
    ```bash
    git clone https://github.com/your-username/Apple_Majin.git
    ```
2.  **依存関係をインストールします:**
    ```bash
    npm install
    ```
3.  **`clasp`にログインします:**
    ```bash
    clasp login
    ```
4.  **新しいGoogle Apps Scriptプロジェクトを作成するか、既存のプロジェクトを使用します。**
5.  **`.clasp.json`ファイルをスクリプトIDで更新します。**
6.  **Google Apps Scriptプロジェクトにコードをプッシュします:**
    ```bash
    clasp push
    ```
7.  **Googleスライドのプレゼンテーションを開き、[カスタム設定]メニューから`generatePresentation`関数を実行します。**

## Getting Started

To get started with this project, you will need to have a Google account and be familiar with Google Apps Script. You will also need to have Node.js and the `clasp` CLI installed on your local machine.

Once you have met these prerequisites, you can clone the repository and follow the instructions in the "Building and Running" section to set up the project.

### はじめに (Japanese)

このプロジェクトを開始するには、Googleアカウントを持ち、Google Apps Scriptに精通している必要があります。また、ローカルマシンにNode.jsと`clasp` CLIがインストールされている必要があります。

これらの前提条件を満たしたら、リポジトリをクローンし、「ビルドと実行」セクションの手順に従ってプロジェクトをセットアップできます。

## Architecture

The project follows a simple architecture:

*   **`コード.js`**: This is the main script that contains the core logic for generating the Google Slides presentation. It defines the `slideData` object and includes functions for creating different types of slides.
*   **`appsscript.json`**: This is the manifest file for the Google Apps Script project. It defines the project's dependencies and other settings.
*   **`system prompt.md`**: This file contains the prompt for the Gemini agent that generates the `slideData` object.
*   **`AGENTS.md`**: This file defines the roles and responsibilities of the AI agents working on the project.
*   **`scripts/`**: This directory contains helper scripts for image analysis and inspection.

### アーキテクチャ (Japanese)

このプロジェクトは、シンプルなアーキテクチャに従っています。

*   **`コード.js`**: これは、Googleスライドプレゼンテーションを生成するためのコアロジックを含むメインスクリプトです。`slideData`オブジェクトを定義し、さまざまな種類のスライドを作成するための関数が含まれています。
*   **`appsscript.json`**: これは、Google Apps Scriptプロジェクトのマニフェストファイルです。プロジェクトの依存関係やその他の設定を定義します。
*   **`system prompt.md`**: このファイルには、`slideData`オブジェクトを生成するGeminiエージェントのプロンプトが含まれています。
*   **`AGENTS.md`**: このファイルは、プロジェクトに取り組んでいるAIエージェントの役割と責任を定義します。
*   **`scripts/`**: このディレクトリには、画像分析と検査のためのヘルパースクリプトが含まれています。

## Development Conventions

*   The main source code is in `コード.js`.
*   The `slideData` object is the primary data structure for generating presentations.
*   The `system prompt.md` file is used to generate the `slideData` object using a Gemini agent. See [`system prompt.md`](./system%20prompt.md) for more details.
*   The `AGENTS.md` file defines the roles and responsibilities of the AI agents. See [`AGENTS.md`](./AGENTS.md) for more details.
*   The `scripts` directory contains helper scripts for image processing.
*   Use `clasp pull` to fetch the latest code from the Google Apps Script project.
*   Use `clasp push` to push local changes to the Google Apps Script project.

### 開発規約 (Japanese)

*   メインのソースコードは`コード.js`にあります。
*   `slideData`オブジェクトは、プレゼンテーションを生成するための主要なデータ構造です。
*   `system prompt.md`ファイルは、Geminiエージェントを使用して`slideData`オブジェクトを生成するために使用されます。詳細については、[`system prompt.md`](./system%20prompt.md)を参照してください。
*   `AGENTS.md`ファイルは、AIエージェントの役割と責任を定義します。詳細については、[`AGENTS.md`](./AGENTS.md)を参照してください。
*   `scripts`ディレクトリには、画像処理用のヘルパースクリプトが含まれています。
*   Google Apps Scriptプロジェクトから最新のコードを取得するには、`clasp pull`を使用します。
*   ローカルの変更をGoogle Apps Scriptプロジェクトにプッシュするには、`clasp push`を使用します。

## Contributing

Contributions to this project are welcome. Please follow these guidelines when contributing:

*   Fork the repository and create a new branch for your changes.
*   Make your changes and test them thoroughly.
*   Submit a pull request with a clear description of your changes.

### 貢献 (Japanese)

このプロジェクトへの貢献を歓迎します。貢献する際は、以下のガイドラインに従ってください。

*   リポジトリをフォークし、変更用の新しいブランチを作成します。
*   変更を加えて、十分にテストしてください。
*   変更内容を明確に説明したプルリクエストを送信してください。