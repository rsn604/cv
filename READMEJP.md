　CVは、Linuxのターミナル、あるいはWindowsのコマンドプロンプト上で動作する簡易スプレッドシートです。
Excelのようなオフィスソフトを立ち上げるまでもない計算処理を手軽に行うことが目的です。四則演算はもちろんのこと、よく使う関数を内蔵し、16進演算もサポートします。

---
## 起動
実行ファイルは、**cv**（あるいはcv.exe）1本だけなので、適当なディレクトリに入れて起動します。

```
$ cv
```

下記のように使用するシートを指定することも可能です。

```
cv sample.cv2
```

終了は下記のメニューから**Quit**選択、あるいは **CTRL+QQ** (CTRLを押したままQを2回)を押します。

以下、添付のサンプルシート（sample.cv2）を例に説明します。

---
## データ入力
特に記述するまでもないでしょう。文字を入力すれば左寄せ、数字入力は右寄せに表示されます。小数点は **.** を入力します。

/ (or PF05)を入力すると、メニューが表示されます。

![MENU](/images/CV_02Menu.png)

セルの移動などの操作については、上記 **H Help** で下記の画面が出てきますので、参照してください。。

![HELP](/images/CV_01Help.png)

---
### (1)　計算式

+(or PF06) を入力すると、計算式入力モードになります。各種四則演算(**+ - * /**)を指定します。

![Formula](/images/CV_03Formula.png)

**@INT**、 **@TRUNC**、 **@ROUND** 関数による処理も可能です。 

等号（**= >= <= <>**）を含む式はその内容を評価、**True**の場合 **1**、**False**の場合 **0** となります。このような式を * で連結すれば、**AND**条件、+ で連結すれば、**OR**条件として評価できます。これは、次項で使用している@IF関数とともに使うことを想定しています。

---
### (2)　関数
計算式入力モードでは、一般的な関数が使用できます。

![FUNCTION](/images/CV_04Function.png)

パラメータは、**@SUM(A1..C3)** のように **Range** を指定します。上記の例では、**@MAX**、**@MIN**、**@AVG**、**@SUM**を集計に、**@VLOOKUP**を他のセルからの読み込み、**@IF**を条件判断に使用しています。

---
### (3)　16進数
**CV**では、**0x**を文字列の最初に入力すると16進数と解釈します。

![HEX](/images/CV_05Hex.png)

16進数は演算にそのまま使用できます。結果は整数となるので、16進表示する場合は、**@HEX**関数を使用します。

---
### (4)　Range
Rangeは、各種設定や関数のパラメータとして事前に指定する必要があります。
始点セルで**CTRL+B**、終点セルで**CTRL->KK**(CTRLを押したままKを2回) とすると設定され、**A1..C2**のように画面右上に表示されます。

![RANGE](/images/CV_06Range.png)

この範囲が次に設定する処理の対象になります。

---
### (5)　その他
メニューから各種処理を選択することができます。

```
File   ->  Load  Save  Clear

Trans  ->  Text Write  CSV Write  CSV Read   

Cell   ->  Width  Decimal  Center  Right  Left  Color

Sort   ->  Ascending  Desending
```

---
## 日本語処理
LinuxではUTF-8、WindowsではShift-JISの漢字コードをサポートしています。

![KANJI](/images/CV_07Kanji.png)

---
## ソースコードとコンパイル
**CV**は、<a href="https://www.freepascal.org/" target="_blank">Free Pascal</a>でコンパイルされています。ダウンロード、インストールしてください。

各環境に複数のファイルがあるようですが、下記のファイルを使用しました。

```
Linux x86_64    fpc-3.2.2.x86_64-linux.tar
Linux ARM       fpc-3.2.2.arm-linux.tar
Windows         fpc-3.2.2.win32.and.win64.exe
```

コンパイル用のシェルを用意していますので、これを利用してください。

```
Linux    compunix.sh
Windows  compwin.bat
```

### Binary
Binaryイメージ (**cv-bin.tar.gz**) を、アップロードしています。右のResouseタグをクリックしてください。

```
bin - linux_x86_64 - cv
    - linux_arm    - cv
    - windows      - cv.exe
sample.cv2
```
上記の構造になっていますので、環境に合わせてファイルを使用してください。



