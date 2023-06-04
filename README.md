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
- PowerPointSlidePaste
    - PowerPointにおいて、指定した番号のスライドを複製する
    - プロパティ
        - スライド番号：複製したい元スライド番号（1～）
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

## Reference
- WinActor

## Author
- superbunnbun

## Licence
[MIT](https://github.com/superbunnbun/DebugNodeVbs/blob/main/LICENSE)