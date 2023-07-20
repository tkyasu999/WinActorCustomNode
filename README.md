# name
- WinActorCustomNode

## Overview
- WinActorカスタムノードを作成するためのVBSプログラム.

## Requirement
- WinActor 7.0以上
- VBScript 5.812

## Usage
- WinActorでカスタムノードを作成し、スクリプトに機能が該当するVBSプログラムを貼り付ける
- カスタムノードは、「スクリプト実行」ノードを利用する

## Features
- RangeCelJoin
    - Excelにおいて、指定した範囲セルを結合する
    - プロパティ
        - ファイル名：Excelファイルのフルパス
        - シート名：対象のシート名
        - 開始セル：結合したい開始セル
        - 終了セル：結合したい終了セル
- ExcelSearch
    - Excelにおいて、指定された範囲で一致したセルの行番号と列番号を取得する. その際に, 検索一致度に完全一致と部分一致を含む.
    - プロパティ
        - 検索タイプ：文字列 or 日付
        - 一致度：完全一致 or 部分一致
        - ファイル名：対象Excelファイルのフルパス
        - シート名：アクティブシート名
        - 検索単語：検索したいキーワード
        - 開始セル：検索範囲の開始セル
        - 終了セル：検索範囲の終了セル
        - 結果（行）：検索ヒットした行
        - 結果（列）：検索ヒットした列
- ExcelShapeCopy
    - Excelにおいて、指定した範囲で図形コピーする.
    - プロパティ
        - ファイル名：対象Excelファイルのフルパス
        - シート名：アクティブシート名
        - 開始セル：コピー範囲の開始セル
        - 終了セル：コピー範囲の終了セル
- SendingMultipleOutlookAttachments
    - Outlookメールで, 複数ファイルを添付して送信する.
    - プロパティ
        - 宛先（To）：メール送信先 To
        - 宛先（Cc）：メール送信先 Cc
        - 件名：メール件名
        - 本文：メール本文
        - 添付用のフォルダ指定：添付ファイルを保存しているフォルダのフルパス
- PowerPointSlidePaste
    - PowerPointにおいて、指定した番号のスライドを複製する
    - プロパティ
        - スライド番号：複製したい元スライド番号（1～）
- PowerPointTextPaste
    - PowerPointにおいて、コピーしたテキストを指定したスライド、位置へ貼り付ける.
    - プロパティ
        - スライド番号：スライド番号（1～）
        - Index：貼り付けるテキストオブジェクトのインデックス番号（1～）
        - Top：貼り付けオブジェクトの左上Y座標（左上0, px）
        - Left：貼り付けオブジェクトの左上X座標（左上0, px）
- PowerPointSlideSelect
    - PowerPointにおいて、指定した番号のスライドを選択する
    - プロパティ
        - スライド番号：選択したいスライド番号（1～）

## Reference
- WinActor

## Author
- tkyasu999

## Licence
[MIT](https://github.com/tkyasu999/DebugNodeVbs/blob/main/LICENSE)