# Day5｜積込リストにチェックボックスでLine通知＋自動通知実装

## 🔥 今日やったこと
- 積込リストにチェックボックスを追加し、ユーザーが完了を記録できるようにした
- A113セルのチェックでシート自動複製 → 日付＋連番で新規シート生成
- 新しいシートに対して入力欄（数量など）を初期化＋チェックボックスも追加
- 完了チェックが入るとLINE通知を送信（担当者・時刻付き）
- 多重トリガー防止のため、`IS_RUNNING` フラグで処理を制御

## 🧠 ポイント
- `PropertiesService` を活用して「今処理中かどうか」をフラグ管理
- チェック1つで自動処理 → シート生成、初期化、通知まで一気通貫！
- LINE通知の仕組みを外部API経由で実装済み（テスト送信関数付き）

## 📦 使用コード（一部抜粋）

```js
function onEdit(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isRunning = scriptProperties.getProperty("IS_RUNNING");
  if (isRunning === "true") return;

  if (e.range.getRow() === 113 && e.range.getColumn() === 1 && e.range.getValue() === true) {
    try {
      scriptProperties.setProperty("IS_RUNNING", "true");
      // シート複製ロジック...
    } finally {
      scriptProperties.deleteProperty("IS_RUNNING");
    }
  }
}