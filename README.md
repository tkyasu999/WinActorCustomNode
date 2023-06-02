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

## Reference
- WinActor

## Author
- superbunnbun

## Licence
[MIT](https://github.com/superbunnbun/DebugNodeVbs/blob/main/LICENSE)