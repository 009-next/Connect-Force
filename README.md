# Connect-Force

Connect-Forceは、現場の記録、会話、画面上の情報を根拠とともに整理し、人の承認を通して次の業務へつなぐAIエージェントです。設計の中心は「AIが選び、人が確定する」です。

![Connect-Force](./hackathon-kit/assets/connect-force-hero.png)

## 体験できること

- **物エージェント**: QR・NFCで物や場所の担当・記録・経緯を開き、マスキング済みの情報からカード下書きを作成
- **統合分析**: 会話と共有画面をつなぎ、対象物の確認、表・説明・下書きの作成を支援
- **Mission Room**: 設備IDから履歴、注意点、条件付き対応案、承認待ちを一画面で確認
- **音声分析**: 明示送信した短いマイク音声をORCA ROUTERで分析し、要約は利用者の確認後に利用

## 審査員向けの起動方法

### Python版 物エージェント

Python 3.11以上を用意します。

```bash
pip install -r requirements.txt
python -m pytest app/tests -q
MIRUCON_ENV=dev PORT=8000 python -m app.web owner@example.test
```

ブラウザで `http://127.0.0.1:8000/login` を開きます。開発用のワンタイムコードは起動したターミナルに表示されます。

### 統合分析で、自分の Gmail アドレスを使う（任意）

トップ画面の「自分の Gmail アドレス」に、自分の gmail.com のアドレスを設定すると、統合分析の「送信の準備」で開く Gmail の作成画面の宛先（To）と開くアカウントに入ります（統合分析がオンのとき表示）。未設定なら宛先なしで開きます。アドレスは、本人だけが設定でき、gmail.com / googlemail.com 以外は受け付けません。送信は、Gmail で人が行います。

「送信の準備」のリンクは、画面には自前の短い URL（`/f/<id>/gmail`）だけを載せ、実際の長い Gmail の作成画面の URL は、そのリンクを踏んだときにサーバーが作ってリダイレクトします。下書きの本文は、この仕組みのおかげで、URL の長さのために削られません。

### 統合分析の共有先フォルダ（オーナー向けの注意）

「共有」で表・文書をコピーする先のフォルダは、**アプリ（Python プロセス）の実行アカウントが書き込める場所**を指定してください。設定の時点で、実際に試し書きをして確認します（書けなければ、その場で断ります）。設定した後にフォルダの権限や空き容量が変わった場合に備え、「共有」を押した時点でも、書き込みに失敗すれば安全に断ります（内部のパスは画面に表示しません）。

### 検証スクリプト（実 API を使う任意の検証。`app/tests` の自動テストとは別）

`verification/` に、実 API（`ORCA_API_KEY` または `ANTHROPIC_API_KEY`）を使って本番の経路を通す検証スクリプトがあります。実行には課金が伴うため、`--yes` を付けない限り、費用の見込みだけを表示して終了します（すべて `--max-cost` で上限を持てます）。

```bash
pip install -r requirements-verification.txt
```

素材（画像・音声）は `verification/assets/README.md` の手順で用意してください。

| スクリプト | 内容 |
|---|---|
| `run_fusion_probe.py` | 統合分析（音声を文字にする→確認→対象物の赤丸・表・文書・メール下書き）を、本番の経路で通す |
| `run_repair_probe.py` | 表の形が壊れた場合に、AI が候補を作り直し、Excel の複数シートに分ける経路を、本番の経路で通す |
| `extract_frame.py` | 画像を 1 コマにした動画から、コマを取り出せるかを確認する（画面共有が使えないときの代わりの入力） |
| `record_demo.py` | ブラウザでの一連の操作を録画する |

### Node.js版 Mission Room

Node.js 24以上とpnpmを用意します。

```powershell
pnpm install
.\scripts\run.ps1 pc
```

PowerShellに表示される一回限りのログインURLを開き、「合成デモを試す」を選びます。根拠表示、対応案の比較、内容確認、承認保存までを体験できます。

## データ配置

公開リポジトリでは、利用者ごとのデータ用フォルダを `your_folder` と表記します。実際のVaultを使う場合は、Markdownを `your_folder/vault/` に置くか、起動前に絶対パスを環境変数で設定します。

```powershell
$env:CONNECT_FORCE_VAULT_ROOT = "D:\\work\\your_folder\\vault"
```

実データ、生成物、環境変数の値はGitへ追加しません。

## 実装・検証

- セキュリティ: 入力検査、端末側マスキング、承認、CSRF、限定した音声送信経路をコードで制御
- コスト: Mission Roomの通常分析はモデル呼出し0回。音声分析は回数と予約予算を制限
- 信頼性: 43件の自動・HTTP統合テスト、型検査、Eve構成検査、本番ビルド、秘密情報スキャンを実施
- ORCA ROUTER E2E: 27.684秒の音声を3回連続で処理し、すべてHTTP 200で完了。合計見積は$0.0011394、デモ上限は$2

APIキーは環境変数で設定し、ソース、Git、チャット、画面へ貼り付けません。

## ハッカソン資料

- [紹介動画リンク先](https://youtu.be/-M7Yo9JgqXE)
- [紹介パンフレット](https://drive.google.com/file/d/18QcyHRg3RN7bBtcFcSCe-44j_hj5mifu/view?usp=sharing)

## ライセンス

ハッカソン提出用のソースです。第三者サービスの利用条件と組織の情報管理ルールに従って利用してください。
