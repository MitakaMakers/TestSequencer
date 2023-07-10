# 試験環境

## ハードウェア
|No.|分類|製品名|備考|
|---|---|---|---|
|1|PC|Panasonic Let'snote CF-NX3||
|2|PC|DELL Latitude 5300||
|3|GPIB|NI GPIB-USB-HS||
|4|GPIB|Prologix GPIB-USB|
|5|RS-232|SUNTAC VS-60R||
|6|USB デバイス|Raspberry Pi Pico||
|7|USB デバイス|Raspberry Pi 3 MODEL B||
|8|ネットワークハブ|NETGEAR GS105E||
<!--||USB,LAN デバイス|Raspberry Pi Zero W|| -->
<!--||ネットワークケーブル|ELECOM LD-GPYT シリーズ|| -->
<!--||USBハブ|ELECOM U3H-A408S シリーズ|| -->
<!--||USBケーブル|ELECOM U2C-ACNBK シリーズ|| -->

## ソフトウェア
|No.|分類|製品名|備考|
|---|---|---|---|
|1|OS|Windows 11||
|2|OS|Windows 10||
|3|OS|Windows 8.1||
|4|表計算ソフト|Office 2021||
|5|表計算ソフト|Office 2007|| <!-- メルカリ -->
<!--||表計算ソフト|Office 2019|| amazon 16800-->
<!--||表計算ソフト|Office 2016|| yahoo 3690 -->
<!--||表計算ソフト|Office 2013|| yahoo 8000 -->
<!--||表計算ソフト|Office 2010|| yahoo オークション 520 -->
<!--||表計算ソフト|Office 2010|| amazon 6980 -->
<!--||表計算ソフト|Office 2003|| amazon 3800 -->

# 作業メモ
## 表計算ソフトが対応する OS のバージョン表
||VBA|Win11|Win10|Win8|Win7|Vista|XP|2000|
|---|---|---|---|---|---|---|---|---|
|Office 2021|VBA 7.1|(Yes)|Yes|-|-|-|-|-|
|Office 2019|VBA 7.1|(Yes)|Yes|-|-|-|-|-|
|Office 2016|VBA 7.1|(Yes)|Yes|Yes|-|-|-|-|
|Office 2013|VBA 7.1|-|(Yes)|Yes|Yes|-|-|-|
|Office 2010|VBA 7.0|-|-|(Yes)|Yes|Yes|Yes|-|
|Office 2007|VBA 6.5|-|-|(Yes)|(Yes)|(Yes)|Yes|-|
|Office 2003|VBA 6.4|-|-|-|(Yes)|(Yes)|Yes|Yes||
|Office 2002|VBA 6.3|-|-|-|-|(Yes)|Yes|Yes|

## 表計算ソフトと OS のサポート期限
||リリース日|メインストリームサポートの終了日|延長サポートの終了日|備考|
|---|---|---|---|---|
|Office 2021|2021年10月5日|2026年10月13日|2026年10月13日||
|Office 2019|2018年9月24日|2023年10月10日|2025年10月14日||
|Office 2016|2015年9月22日|2020年10月13日|2025年10月14日|パッケージ版が廃止された最初のバージョン|
|Office 2013|2013年1月19日|2018年4月10日|2023年4月11日||
|Office 2010|2010年7月5日|2015年10月13日|2020年10月13日|最初の 64ビット版 Office |
|Office 2007|2007年1月27日|2012年10月9日|2017年10月11日|「xlsx」などの新しいファイル形式が導入された最初のバージョン|
|Office 2003|2003年11月17日|2009年4月14日|2014年4月8日|デジタル署名が導入された最初のバージョン|
|Office 2002|2001年3月31日|2006年7月11日|2011年7月12日||
|Windows 11|2021年11月4日|未定|未定||
|Windows 10|2015年7月29日|2025年10月14日|2025年10月14日||
|Windows 8.1|2012年10月30日|2018年1月9日|2023年1月10日||
|Windows 7|2009年10月22日|2015年1月13日|2020年1月14日|SHA-256が使える一番古いバージョン|
|Windows Vista|	2007年1月25日|2012年4月10日|2017年4月11日||
|Windows XP|2001年12月31日|2009年4月14日|2014年4月8日||
|Windows 2000|2000年3月31日|2005年6月30日|2010年7月13日||
|Visual Studio 2022|2021年11月8日|2027年1月12日|2032年1月13日||
|Visual Studio 2019|2019年4月2日|2024年4月9日|2029年4月10日||
|Visual Studio 2017|2017年3月7日|2022年4月12日|2027年4月13日|Windows XP 用のコード作成をサポートする最後のバージョン|
|Visual Studio 2015|2015年7月20日|2020年10月13日|2025年10月14日||
|Visual Studio 2013|2014年1月15日|2019年4月9日|2024年4月9日||
|Visual Studio 2012|2012年10月31日|2018年1月9日|2023年1月10日||
|Visual Studio 2010|2010年6月29日|2015年7月14日|2020年7月14日||
|Visual Studio 2008|2008年2月19日|2013年4月9日|2018年4月10日||
|Visual Studio 2005|2006年1月27日|2011年4月12日|2016年4月12日|Unicode 対応 した最初のバージョン|
|Visual Studio 2003|2003年7月10日|2008年10月14日|2013年10月8日|Java 訴訟の和解条件により配布終了|
|Visual Studio 2002|2002年4月15日|2007年7月10日|2009年7月14日|Java 訴訟の和解条件により配布終了|
|Visual Basic 6.0|1998年9月5日|2005年3月31日|2008年4月8日||
|Visual C++ 6.0|1998年9月25日|2004年9月30日|2005年9月30日||

## C++, VB 開発環境が対応する OS のバージョン表
||Win11|Win10|Win8|Win7|Vista|XP|2000|98|
|---|---|---|---|---|---|---|---|---|
|Visual C++ 2022 再頒布パッケージ|Yes|Yes|Yes|Yes|Yes|-|-|-|
|Visual C++ 2019 再頒布パッケージ|Yes|Yes|Yes|Yes|Yes|(Yes)|-|-|
|Visual C++ 2017 再頒布パッケージ|Yes|Yes|Yes|Yes|Yes|Yes|-|-|
|Visual C++ 2015 再頒布パッケージ|(Yes)|Yes|Yes|Yes|Yes|Yes|-|-|
|Visual C++ 2013 再頒布パッケージ|-|-|Yes|Yes|Yes|Yes|-|-|
|Visual C++ 2012 再頒布パッケージ|-|-|Yes|Yes|Yes|Yes|-|-|
|Visual C++ 2010 再頒布パッケージ|-|-|-|Yes|Yes|Yes|-|-|
|Visual C++ 2008 再頒布パッケージ|-|-|-|Yes|Yes|Yes|-|-|
|Visual C++ 2005 再頒布パッケージ|-|-|-|Yes|Yes|Yes|-|-|
|Visual Basic 6.0ランタイム|Yes|Yes|Yes|Yes|Yes|Yes|Yes|Yes|
|Visual Studio 2022 開発環境|Yes|Yes|-|-|-|-|-|-|
|Visual Studio 2019 開発環境|Yes|Yes|Yes|Yes|-|-|-|-|
|Visual Studio 2017 開発環境|-|Yes|Yes|Yes|-|-|-|-|
|Visual Studio 2015 開発環境|-|Yes|Yes|Yes|-|-|-|-|
|Visual Studio 2013 開発環境|-|-|Yes|Yes|-|-|-|-|
|Visual Studio 2012 開発環境|-|-|Yes|Yes|-|-|-|-|
|Visual Studio 2010 開発環境|-|-|(Yes)|Yes|Yes|Yes|-|-|
|Visual Studio 2008 開発環境|-|-|(Yes)|(Yes)|(Yes)|Yes|Yes|-|
|Visual Studio 2005 開発環境|-|-|-|(Yes)|(Yes)|Yes|Yes|-|
|Visual Studio 2003 開発環境|-|-|-|-|-|Yes|Yes|-|
|Visual Studio 2002 開発環境|-|-|-|-|-|Yes|Yes|-|
|Visual Basic 6.0 開発環境|-|-|-|-|-|(Yes)|(Yes)|Yes|
