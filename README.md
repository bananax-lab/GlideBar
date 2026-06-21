# GlideBar

GlideBar is an AutoHotkey v2 script for fast scrolling with F15 + mouse movement.

F15キーを押しながらマウスを動かすことで、Excel / JMP / Notepad などを縦横にスクロールする Windows 用 AutoHotkey v2 スクリプトです。

## GlideBar とは

GlideBar は、`F15` キーを押しながらマウスを動かすことで、Excel / JMP / Notepad などを縦横にスクロールできる Windows 用 AutoHotkey v2 スクリプトです。

スクロールバーをつかんで動かすような感覚で、表や一覧を素早く移動することを目的にしています。

## 背景

既存の swipe-scroll 系ツールも便利ですが、自分の用途では Excel の大きな表を扱うときにラグが出たり、縦横の軸ロックが欲しくなったりしました。

GlideBar では、Excel だけ専用のスクロール処理にするなど、実務で使いやすいように細かく調整しています。

特に、数千〜数万行規模の表を扱うことを想定しているため、初期設定ではスクロール速度をかなり速めにしています。速度や感度はパラメータで調整できます。

## 主な特徴

- `F15` + マウス移動でスクロール
- 縦方向・横方向の両方に対応
- 初動方向による軸ロック
- Excel 向けの専用スクロール処理
- JMP / Notepad など一般アプリにも対応
- 高速スクロール向けの初期設定
- ホットキー変更可能
- AutoHotkey v2 製

## 今後の予定

v1.3 系では、スクロール量に応じた加速制御を実装予定です。







## Status

Current stable version: v1.242

v1.3x is experimental and may contain bugs.

## Features

- F15 + mouse movement scrolling
- Axis lock based on initial mouse movement
- Direct ScrollRow / ScrollColumn control for Excel
- Wheel-based scrolling for non-Excel applications
- Click blocking while GlideBar is active

## Requirements

- Windows
- AutoHotkey v2

## Usage

1. Install AutoHotkey v2
2. Download `GlideBar.ahk`
3. Run the script
4. Hold F15 and move the mouse

## Notes

This tool is mainly tested in my own Windows environment.

Behavior may vary depending on applications, mouse drivers, and Windows settings.

## Disclaimer

This is a personal productivity tool.
It is provided as-is, without warranty.
I cannot guarantee that it works in every environment.