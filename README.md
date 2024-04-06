CV is a simple spreadsheet that runs on the Linux terminal or Windows command prompt.

The purpose is to easily perform calculation processing that does not require launching office software such as Excel.
In addition to the four arithmetic operations, CV has built-in frequently used functions and also supports hexadecimal operations.

---
## Boot
There is only one executable file, cv (or cv.exe), so put it in an appropriate directory and start it. 

```
$ cv
```

You can also specify the sheet to use as shown below.

```
cv sample.cv2
```

To exit, from the menu below  select Quit , or press CTRL+QQ (hold down CTRL and press Q twice).

The following explanation uses the attached sample sheet (sample.cv2) as an example.

---
## data entry
There is no need to specifically describe it.
If you enter characters, they will be aligned to the left, and if you enter numbers, aligned to the right.  Enter '.' for decimal point .

Enter '/ '(or PF05) to display a menu. 

![MENU](/images/CV_02Menu.png)

For operations such as moving cells, refer to 'H Help' from 'Menu' screen .

![HELP](/images/CV_01Help.png)

---
### (1)　Calculation formula
Enter '+'(or PF06) to enter calculation formula-input mode. Specify various arithmetic operations like ( + - * / ). 

![Formula](/images/CV_03Formula.png)

Processing with @INT , @TRUNC , and @ROUND functions is also possible. 

Expressions containing equal signs ( = >= <= <> ) is evaluated . If result is true, return value is 1 . If false return 0 .
By concatenating Expressions with '*' can be evaluated as AND condition, also '+' for OR condition .
This is intended to be used with the @IF function in next section.

---
### (2)　Function
General functions can be used in formula-input mode. 

![FUNCTION](/images/CV_04Function.png)

Specify Range as parameter, like @SUM(A1..C3) .
In the above example, @MAX , @MIN , @AVG , and @SUM are used for aggregation, @VLOOKUP is used for reading from other cells, and @IF is used for conditional judgment.

---
### (3)　Hexadecimal number 
In CV , 0x at the beginning of string, it will be interpreted as a hexadecimal number.

![HEX](/images/CV_05Hex.png)

Hexadecimal numbers can be used as is for calculations. The result will be an integer, so if you want to display it in hexadecimal, use @HEX function. 

---
### (4)　Range
Range must be specified in advance as a parameter for various settings and functions.

On the start cell, press CTRL+B. And on the end cell, press CTRL->KK (press K twice while holding down CTRL). Range will be displayed at the top right of the screen like A1..C2 . 

![RANGE](/images/CV_06Range.png)

This range will be the target of the next processing. 

---
### (5)　Others 
You can select various processing from 'Menu'. 

```
File   ->  Load  Save  Clear

Trans  ->  Text Write  CSV Write  CSV Read   

Cell   ->  Width  Decimal  Center  Right  Left  Color

Sort   ->  Ascending  Desending
```

---
## Japanese processing
Supports Kanji codes, UTF-8 on Linux and Shift-JIS on Windows. 

![KANJI](/images/CV_07Kanji.png)

---
## Source code and compilation 
CV is compiled with by <a href="https://www.freepascal.org/" target="_blank">Free Pascal</a> . Please download and install it. 

It seems that each environment has multiple files, but I used the files below.

```
Linux x86_64    fpc-3.2.2.x86_64-linux.tar
Linux ARM       fpc-3.2.2.arm-linux.tar
Windows         fpc-3.2.2.win32.and.win64.exe
```

Shell for compilationis prepared, please use it.

```
Linux    compunix.sh
Windows  compwin.bat
```

### Binary
the Binary image ( cv-bin.tar.gz ) is uploaded. Click on the Resouse tag on the right. 

```
bin - linux_x86_64 - cv
    - linux_arm    - cv
    - windows      - cv.exe
sample.cv2
```
The structure is as shown above, use the file according to your environment. 



