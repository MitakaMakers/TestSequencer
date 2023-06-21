[English](README.md) | [日本語](README.ja.md)

# エクセルコマンドとは
エクセルコマンドは測定器を簡単に制御するためのオープンソースソフトウェアです。
エクセルファイルに通信コマンドを記入し、各種通信インタフェースで送受信することができます。
将来的にフルークの MET/CAL のような誰でも使える校正自動化ソフトウェアを目標にしています。

# 特徴
- Excel VBA（表計算マクロ言語）と TMCTL.DLL で動作します。インストール作業は不要です。
- 測定器と GP-IB、RS232、USB、LAN で通信できます。
- 最大8台の測定器の制御や計測ができます。
- 日本語表記と英語表記を切り替えられます。

# 対象機器
- IEEE488.2-1987 に対応する測定器

# 動作環境
- OS：Windows 2000, XP, Vista, 7, 8, 10, 11
- Excel：2010, 2013, 2016, 2019, 2021
  - Office for Mac や Microsoft 365 は非対応です。

# 通信インタフェース
- GP-IB：ナショナルインスツルメンツ製 GP-IB インタフェースが動作する環境。
  - 別途 [NI-488.2](https://www.ni.com/ja-jp/support/downloads/drivers/download.ni-488-2.html) をインストールしてください。
- RS232C：シリアルポート又は仮想COMポートが動作する環境。
- LAN：ソケット通信、VXI-11 または HiSLIP が動作する環境。
- USB：ナショナルインスツルメンツ製 NI-VISA または　横河計測製 USB ドライバが動作する環境。
  - 別途 [NI-VISA](https://www.ni.com/ja-jp/support/downloads/drivers/download.ni-visa.html) または[横河計測製 USB ドライバ](https://tmi.yokogawa.com/jp/library/documents-downloads/software/usb-drivers/)をインストールしてください。

# 使い方
## ダウンロードと展開
ZIP ファイルをダウンロード後にファイルを展開し、ExcelCommand.xlsm, tmctl.dll, tmctl64.dll, YKMUSB.dll, YKMUSB64.dll の5個のファイルを同一ディレクトリに置いてください。

## マクロを含むブックを開く
Excel の初期の設定ではマクロを含むブックを開こうとすると、セキュリティの警告を表示してマクロを無効にします。マクロを有効に設定する方法は、次の通りです。

- 「ファイル」タブをクリックし、「オプション」をクリック
- 左側の一覧から「セキュリティセンター」をクリックし、 「セキュリティセンターの設定」をクリック
- 左側の一覧から「マクロの設定」をクリック
- 「マクロの設定」の一覧から「VBAマクロを有効にする」をクリック

## アドレス文字列
![<img src="docs/101j.png">](docs\101j.png)

## 命令
![<img src="docs/102j.png"](docs\102j.png)

## 記入例
![<img src="docs/103j.png"](docs\103j.png)

# 著作権表記
Excel Commmand: An excel macro file to communicate some measurement insturuments.

Copyright (C) 2023 Takatoshi Yamaoka

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.