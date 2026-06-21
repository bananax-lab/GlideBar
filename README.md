# GlideBar

GlideBar is a Windows AutoHotkey v2 script for fast vertical and horizontal scrolling with `F15` + mouse movement.

`F15` キーを押しながらマウスを動かすことで、Excel / JMP / Notepad などを縦横にスクロールする Windows 用 AutoHotkey v2 スクリプトです。

---


## Overview

GlideBar is designed to feel like grabbing and moving a scrollbar, making it easier to move quickly through large tables and long lists.

It was made for my own workflow because existing swipe-scroll tools did not fully match what I wanted, especially when working with large Excel sheets.

For Excel, GlideBar uses a dedicated scrolling method instead of relying on normal wheel scrolling. This helps reduce lag and allows faster movement through large tables.

Since it is intended for tables with thousands or tens of thousands of rows, the default scrolling speed is set fairly high. The speed, sensitivity, and hotkey can be adjusted through parameters.

## Features

- `F15` + mouse movement scrolling
- Vertical and horizontal scrolling
- Axis lock based on initial mouse movement
- Dedicated ScrollRow / ScrollColumn control for Excel
- Wheel-based scrolling for non-Excel applications
- Fast default settings for large tables
- Customizable hotkey and parameters
- Click blocking while GlideBar is active
- Built with AutoHotkey v2

## Status

Current stable version: `v1.242`

`v1.3x` is experimental and may contain bugs.

Acceleration control based on mouse movement amount is planned for the v1.3 series.

## Requirements

- Windows
- AutoHotkey v2

## Usage

1. Install AutoHotkey v2
2. Download `GlideBar.ahk`
3. Run the script
4. Hold `F15` and move the mouse

## Notes

This tool is mainly tested in my own Windows environment.

Behavior may vary depending on applications, mouse drivers, and Windows settings.

## Disclaimer

This is a personal productivity tool.

It is provided as-is, without warranty.
I cannot guarantee that it works in every environment.

---


## 概要

GlideBar は、`F15` キーを押しながらマウスを動かすことで、Excel / JMP / Notepad などを縦横にスクロールできる Windows 用 AutoHotkey v2 スクリプトです。

スクロールバーをつかんで動かすような感覚で、表や一覧を素早く移動することを目的にしています。

既存の swipe-scroll 系ツールも便利ですが、自分の用途では、Excel の大きな表を扱うときのラグや、縦横スクロールの軸ロックが気になったため、自作しました。

Excel では通常のホイール入力ではなく、専用のスクロール処理を使っています。これにより、大きな表でもできるだけ軽快にスクロールできるようにしています。

数千〜数万行規模の表を扱うことを想定しているため、初期設定ではスクロール速度をかなり速めにしています。速度、感度、ホットキーはパラメータで調整できます。

## 主な特徴

- `F15` + マウス移動でスクロール
- 縦方向・横方向の両方に対応
- 初動方向による軸ロック
- Excel 向けの専用 ScrollRow / ScrollColumn 制御
- Excel 以外のアプリではホイール入力ベースでスクロール
- 大きな表を想定した高速寄りの初期設定
- ホットキーや各種パラメータを変更可能
- GlideBar 動作中の誤クリック防止
- AutoHotkey v2 製

## 開発状況

現在の安定版: `v1.242`

`v1.3x` は実験版で、バグを含む可能性があります。

v1.3 系では、マウス移動量に応じた加速制御を実装予定です。

## 必要環境

- Windows
- AutoHotkey v2

## 使い方

1. AutoHotkey v2 をインストールする
2. `GlideBar.ahk` をダウンロードする
3. スクリプトを実行する
4. `F15` を押しながらマウスを動かす

## 注意点

このツールは、主に自分の Windows 環境で動作確認しています。

アプリケーション、マウスドライバ、Windows の設定によって挙動が変わる可能性があります。

## 免責

これは個人用に作成した作業効率化ツールです。

現状のまま提供しており、動作保証はありません。
すべての環境で正常に動作することは保証できません。