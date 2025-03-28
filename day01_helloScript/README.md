# Day1｜GASでB3セルにメッセージを書いてみた

## 概要
Google Apps Scriptで、スプレッドシートのB3セルに「お疲れ様です!」と書き込む簡単なスクリプト。

## 使用技術
- Google Apps Script

## 気づき・学び
- `.getRange("B3").setValue()` の使い方を覚えた。
- Logger.log() で確認できるのが便利。