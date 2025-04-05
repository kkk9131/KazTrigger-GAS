# Day5｜積込リストにチェックボックスでLine通知＋自動通知実装

## 概要

- 積込リストの **A113セルにチェック**を入れることで、
  - 現在のシートを **日付＋連番付きで複製**
  - 複製先シートの入力エリアを **初期化**
  - 各行に **チェックボックスを自動付与**
- B113にチェックが入ると、
  - 担当者名と時刻を取得して **LINEに通知**
- 実行中に重複処理が起きないよう `IS_RUNNING` フラグで制御

## 使用技術

- Google Apps Script（GAS）
  - `onEdit(e)` トリガー
  - `PropertiesService` で状態管理
  - `UrlFetchApp` を使った LINE API 通信
- Google スプレッドシート
  - チェックボックス機能
  - シートのコピー & 名前変更

## 気づき・学び

- チェックボックス1つで処理の起点を作ると、ユーザーが迷わず使いやすい
- `PropertiesService` による排他制御は、GASにおける多重トリガー対策の定番！
- LINE通知は `muteHttpExceptions: true` をつけておくと、失敗時もログが拾いやすく安心
- 通知トリガー・シート複製・初期化が1つの流れになることで、「自動で動く感」が強くなる

## 参考リンク（あれば）

- [LINE Messaging API リファレンス](https://developers.line.biz/ja/reference/messaging-api/)
- [GAS公式：onEditイベント](https://developers.google.com/apps-script/guides/triggers#onedite)

---

「勝手に動くしくみを、自分の手で。」