{$MODE TP}

Unit cvcell ;

Interface

Uses cvdef, cvcrt, cvscrn, cvinpt, cveval, cvdata, cvconv, cvundo ;

Function GetNumericWidth(x,y : Integer; fData:Real):  Integer;
Function AdjustWidth(x:Integer):  Integer ;
Procedure AdjustAllWidth ;
Procedure ChangeWidthByValue(hWidth:Integer) ;
Procedure ChangeWidth(hTopCol,hEndCol:integer) ;
Procedure ChangeDecPoint(hTopCol, hTopRow, hEndCol, hEndRow:integer);
Procedure EraseCellData(hTopCol, hTopRow, hEndCol, hEndRow: Integer);
Procedure TransCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer; cFlag:Byte)
;
Procedure MoveCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer);
Procedure CopyCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer);
Procedure ChangeCellJustify(hTopCol, hTopRow, hEndCol, hEndRow:integer; cFlag:Byte) ;
Procedure CellMenu ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Numeric Width                                 }
{ ---------------------------------------------- }
Function GetNumericWidth(x,y : Integer; fData:Real):  Integer;

Var 
   i :  Integer ;
Begin
   i := GetCellDecPoint(x, y);
   If i<>0 Then
      i := i+2
   Else
      i := i+1;
   If fData<0 Then
      i := i+1;
   GetNumericWidth := round(ln(abs(fData))/ln(10))+i ;
End ;

{ ---------------------------------------------- }
{  Adjust Width                                  }
{ ---------------------------------------------- }
Function AdjustWidth(x:Integer):  Integer ;

Var 
   hMaxCellWidth:  Integer ;
   j:  Integer ;

   hWidth, hMaxWidth:  Integer ;
   sData     :  String;
   hErrPos    :  Integer ;
   hStatus    :  Integer ;
   fData     :  Real;
   sRet:  String ;

Begin
   hMaxCellWidth := MaxCol-ghLeftSide-1 ;
   hMaxWidth := 0 ;
   For j:= 1 To ghMaxY Do
      Begin
         If Sheet[x,j].tpMain <> Nil Then
            Begin
               sData := Sheet[x,j].tpMain^ ;
               If IsCellFormula(Sheet[x,j]) Then
                  Begin
                     Evaluate(hStatus, sData, fData, hErrPos, sRet, x, j) ;
                     If hErrPos=0 Then
                        Begin
                           If hStatus = h_FORMULA_NUMERIC Then
                              Begin
                                 hWidth := GetNumericWidth(x,j,fData) ;
                                 If hWidth > hMaxWidth Then
                                    hMaxWidth := hWidth ;
                              End
                           Else If hStatus = h_FORMULA_STRING Then
                                   Begin
                                      hWidth := length(sRet) ;
                                      If hWidth > hMaxWidth Then
                                         hMaxWidth := hWidth ;
                                   End;
                        End;
                  End
               Else
                  Begin
                     hWidth := length(sData) ;
                     If hWidth > hMaxWidth Then
                        hMaxWidth := hWidth ;
                  End ;
            End ;
      End;

   If hMaxWidth > 0 Then
      Begin
         If hMaxWidth < hMaxCellWidth Then
            gcCellWidth[x] := hMaxWidth+1
         Else
            gcCellWidth[x] := hMaxCellWidth ;
      End ;
   AdjustWidth := hMaxWidth ;
End ;

Procedure AdjustAllWidth ;

Var 
   i:  Integer ;
Begin
   For i:= 1 To ghMaxX Do
      AdjustWidth(i) ;

   SetScreen;
End ;

{ ---------------------------------------------- }
{  Change Width                                  }
{ ---------------------------------------------- }
Procedure ChangeWidthByValue(hWidth:Integer) ;

Var 
   hMaxCellWidth:  Integer ;
Begin
   hMaxCellWidth := MaxCol-ghLeftSide-1 ;

   If (hWidth>=1) And (hWidth<= hMaxCellWidth) Then
      Begin
         gcCellWidth[ghX] := hWidth;
         SetMaxX(ghX) ;                 { CVDATA }

         SetScreen;
      End;
End ;

Procedure ChangeWidth(hTopCol,hEndCol:integer) ;

Var 
   hWidth, hResult, hMaxCellWidth:  Integer ;
   sWidth:  String ;
   cEdit :  Byte ;
   cAscii :  Char;

Begin

  { ******************************** }
  {   Get Width                      }
  { ******************************** }

   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   hMaxCellWidth := MaxCol-ghLeftSide-1 ;


//write('Specify width (Current:',gcCellWidth[ghX],' [ 1..', hMaxCellWidth, ']) or "A":adjust current, "AA":adjust all.');
   write('Specify width [ 1..', hMaxCellWidth, ']) or "A":Adjust current, "AA":Adjust all.');
   cEdit := c_Edit ;

   GetInputData(sWidth, cAscii, cEdit);
   If (sWidth = '') Or (cAscii = ESCKEY) Then
      exit ;

   If (UpcaseString(sWidth) = 'AA') Then
      Begin
         AdjustAllWidth ;
         exit ;
      End ;

   If (UpcaseString(sWidth) = 'A') Then
      Begin
         If AdjustWidth(ghX) > 0 Then
            SetScreen;
         exit ;
      End ;

   val(sWidth,hWidth,hResult);
   If hResult <> 0 Then
      exit ;

   If hWidth <= hMaxCellWidth Then
      ChangeWidthByValue(hWidth) ;

End;

{ ---------------------------------------------- }
{  Change Decimal Point                          }
{ ---------------------------------------------- }
Procedure ChangeDecPoint(hTopCol, hTopRow, hEndCol, hEndRow:integer);

Var 
   hWidth, hResult,i, j, hSeq:  Integer ;
   sWidth:  String ;
   cEdit :  Byte ;
   cAscii :  Char;

Begin

  { ******************************** }
  {   Get Decimal Point              }
  { ******************************** }
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);

   //   write('Specify decimal point (Current:',GetCellDecPoint(hEndCol, hEndRow) ,' [ 0..7 ] )');
   write('Specify decimal point (Current:',GetCellDecPoint(hEndCol, hEndRow) ,' [ 0..',h_MAXDECPOINT
   ,'] )');
   cEdit := c_Edit ;
   GetInputData(sWidth, cAscii, cEdit);
   If (sWidth = '') Or (cAscii = ESCKEY) Then
      exit ;

   val(sWidth,hWidth,hResult);
   If hResult <> 0 Then
      exit ;

   hSeq := 0 ;
   If (hWidth>=0) And (hWidth<=h_MAXDECPOINT) Then
      Begin
         For j:=hTopRow To hEndRow Do
            Begin
               For i:=hTopCol To hEndCol Do
                  Begin
                     UndoLog(hSeq,i, j) ;   { CVUNDO }
                     hSeq := hSeq + 1 ;
                     SetCellDecPoint(i, j, hWidth);
                     //                     SetScreen ;        { CVSCRN }
                  End ;
            End ;
      End ;
   SetMaxX(hEndCol) ;                     { CVDATA }
   SetScreen ;        { CVSCRN }
End;

{ ---------------------------------------------- }
{  Erase Cell Data                               }
{ ---------------------------------------------- }
Procedure EraseCellData(hTopCol, hTopRow, hEndCol, hEndRow: Integer);

Var 
   i, j, hSeq:  Integer;

Begin
   hSeq := 0 ;
   For i:=hTopCol To hEndCol Do
      For j:=hTopRow To hEndRow Do
         If Sheet[i,j].tpMain<>Nil Then
            Begin
               UndoLog(hSeq,i, j) ;   { CVUNDO }
               hSeq := hSeq + 1 ;
               FreeCellArea(Sheet[i,j]) ;      { CVEVAL }
            End;
  { ******************************** }
  {   Check Max Cell Data            }
  { ******************************** }
   If (hEndCol>=ghMaxX) And (hEndROW>=ghMaxY) Then
      Begin
         If (hTopCol<ghMaxX) Then
            ghMaxX := hTopCol ;
         If (hTopRow<ghMaxY) Then
            ghMaxY := hTopRow ;
      End ;

   SetScreen;     { CVSCRN }
End;

{ ---------------------------------------------- }
{  Trans  Cell Data                              }
{ ---------------------------------------------- }
Procedure TransCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer; cFlag:Byte)
;

Var 
   hOrgx,hOrgy:  Integer;
   hDestx,hDesty:  Integer;
   i, j, hSeq:  Integer ;
Begin

   If (hTopCol = hDestCol) And (hTopRow = hDestRow) Then
      exit ;

   hSeq := 0 ;
   For i:=hTopCol To hEndCol Do
      Begin
         If hTopCol>hDestCol Then
            hOrgx := i
         Else
            hOrgx := hEndCol-(i-hTopCol);
         hDestx := hOrgx+(hDestCol-hTopCol);
         If hDestx > h_MAXCOL Then
            exit ;

   { ******************************** }
   {   Trans Cell Data                }
   { ******************************** }
         For j:=hTopRow To hEndRow Do
            Begin
               If hTopRow>hDestRow Then
                  hOrgy := j
               Else
                  hOrgy := hEndRow-(j-hTopRow);
               hDesty := hOrgy+(hDestRow-hTopRow);
               If hDesty > h_MAXROW Then
                  exit ;

       { ******************************** }
              {   Copy Cell Data                 }
              { ******************************** }
               If cFlag = c_COPYCELL Then
                  Begin
                     UndoLog(hSeq,hDestx,hDesty) ;   { CVUNDO }
                     hSeq := hSeq + 1 ;

                     Sheet[hDestx,hDesty].cCellType := Sheet[hOrgx,hOrgy].cCellType ;
                     Sheet[hDestx,hDesty].cCellColor := Sheet[hOrgx,hOrgy].cCellColor ;
                     If Sheet[hOrgx,hOrgy].tpMain<>Nil Then
                        Begin

                           If Sheet[hDestx,hDesty].tpMain = Nil Then
                              GetCellArea(Sheet[hDestx,hDesty]) ;
                           Sheet[hDestx,hDesty].tpMain^ := Sheet[hOrgx,hOrgy].tpMain^ ;
                        End ;
                  End

              { ******************************** }
              {   Move Cell Data                 }
              { ******************************** }
               Else
                  Begin
                     UndoLog(hSeq,hOrgx,hOrgy) ;   { CVUNDO }
                     hSeq := hSeq + 1 ;
                     UndoLog(hSeq,hDestx,hDesty) ;   { CVUNDO }
                     hSeq := hSeq + 1 ;
                     Sheet[hDestx,hDesty] := Sheet[hOrgx,hOrgy];
                     Sheet[hOrgx,hOrgy].tpMain := Nil;
                     SetCellNotFormula(Sheet[hOrgx,hOrgy]) ;
                     Sheet[hOrgx,hOrgy].cCellColor 
                     := Sheet[hOrgx,hOrgy].cCellColor;
                  End ;

              { ******************************** }
              {   Convert Formula                }
              { ******************************** }
               If IsCellFormula(Sheet[hDestx,hDesty]) Then
                  Convert(Sheet[hDestx,hDesty].tpMain^,
                          hOrgx,
                          hOrgy,
                          hDestx,
                          hDesty,
                          c_CHANGE_XY) ;
            End;
      End;

 { ******************************** }
 {   Check Max Data                 }
 { ******************************** }
   If ghMaxX < hDestCol+(hEndCol-hTopCol) Then
      ghMaxX := hDestCol+(hEndCol-hTopCol) ;

   If ghMaxY < hDestRow+(hEndRow-hTopRow) Then
      ghMaxY := hDestRow+(hEndRow-hTopRow) ;

End;

{ ---------------------------------------------- }
{  Move  Cell Data                               }
{ ---------------------------------------------- }
Procedure MoveCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer);
Begin
   TransCellData(hTopCol, hTopRow,
                 hEndCol, hEndRow,
                 hDestCol, hDestRow,
                 c_MOVECELL);
   SetScreen;        { CVSCRN }
End;

{ ---------------------------------------------- }
{  Copy  Cell Data                               }
{ ---------------------------------------------- }
Procedure CopyCellData(hTopCol, hTopRow, hEndCol, hEndRow, hDestCol, hDestRow: Integer);
Begin
   TransCellData(hTopCol, hTopRow,
                 hEndCol, hEndRow,
                 hDestCol, hDestRow,
                 c_COPYCELL);
   SetScreen;        { CVSCRN }
End;

{ ---------------------------------------------- }
{  Change Cell Justify                           }
{ ---------------------------------------------- }
Procedure ChangeCellJustify(hTopCol, hTopRow, hEndCol, hEndRow:integer; cFlag:Byte) ;

Var 
   i, j, hSeq:  Integer;
Begin
   hSeq := 0 ;
   For j:=hTopRow To hEndRow Do
      Begin
         For i:=hTopCol To hEndCol Do
            Begin
               UndoLog(hSeq,i, j) ;   { CVUNDO }
               hSeq := hSeq + 1 ;

               If Sheet[i,j].tpMain <> Nil Then
                  Begin
                     ChangeFirstChar(Sheet[i,j].tpMain^) ;
                     If  (cFlag = c_CELL_CENTER) Then
                        Sheet[i,j].tpMain^ := s_CELL_CENTER +Sheet[i,j].tpMain^;

                     If  (cFlag = c_CELL_RIGHT) Then
                        Sheet[i,j].tpMain^ := s_CELL_RIGHT+Sheet[i,j].tpMain^;

                     If  (cFlag = c_CELL_LEFT) Then
                        Sheet[i,j].tpMain^ := s_CELL_LEFT+Sheet[i,j].tpMain^;
                  End ;
            End ;
      End ;

   SetScreen;     { CVSCRN }
End ;

{ ---------------------------------------------- }
{  Change Cell Color                             }
{ ---------------------------------------------- }
Procedure ChangeCellColor(hTopCol, hTopRow, hEndCol, hEndRow:integer) ;

Type 
   ColorFunction =  array [0..15] Of Byte ;


Const 

   ColorAttributes:  ColorFunction =  (Black, Blue, Green, Cyan, Red, Magenta, Brown, LightGray,
                                       DarkGray, LightBlue, LigtnGreen, LightCyan, LightRed,
                                       LightMagenta, Yellow, White) ;

Var 
   i, j, hSeq      :  Integer;
   hColor, hResult :  Integer ;
   sColor     :  String ;
   cEdit     :  Byte ;
   cAscii     :  Char;
   cAttr     :  Byte ;
Begin
  { ******************************** }
  {   Set Color List                 }
  { ******************************** }
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   cAttr := GetAttr ;
   For i:=0 To Length(ColorAttributes)-1 Do
      Begin
         SetAttr(ColorAttributes[i]) ;
         //SetHighVideo ;
         write(ColorAttributes[i], ' ') ;
      End ;

{
   SetAttr(White) ;
   SetHighVideo ;
   write('H', ' ' ) ;
   SetLowVideo ;
   write('L', ' ' ) ;
}
   SetAttr(White*Reverse And Not Blink) ;
   write('R', ' ' ) ;
   SetAttr(cAttr) ;

  { ******************************** }
  {   Get Color                      }
  { ******************************** }
   cEdit := c_Edit ;
   GetInputData(sColor, cAscii, cEdit);
   If (sColor = '') Or (cAscii = ESCKEY) Then
      exit ;

   val(sColor, hColor, hResult);
   If hResult <> 0 Then
      hColor := -1 ;

   hSeq := 0 ;
   For j:=hTopRow To hEndRow Do
      Begin
         For i:=hTopCol To hEndCol Do
            Begin
               UndoLog(hSeq,i, j) ;   { CVUNDO }
               hSeq := hSeq + 1 ;
               If (hColor >= 0) And (hColor <= 15) Then
                  Sheet[i,j].cCellColor := hColor
               Else If UpcaseString(sColor) = 'R' Then
                       Sheet[i,j].cCellColor := (Sheet[i,j].cCellColor shl 4) And $70  ;
            End ;
      End ;

   SetScreen;     { CVSCRN }
End ;

{ ---------------------------------------------- }
{  Cell Menu                                     }
{ ---------------------------------------------- }
Procedure CellMenu ;

Var 
   hMenu:  Integer ;
Begin
   //WriteRange;      { Write target Range CV#SCRN }

   ClearMenuLine ;

{$IFDEF JPN}
   MenuList[1].Menu := 'W セル幅 ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'D 小数点 ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C 中揃え ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'R 右揃え ';
   MenuList[4].Num  := 4;
   MenuList[5].Menu := 'L 左揃え ';
   MenuList[5].Num  := 5;
   MenuList[6].Menu := 'O 色     ';
   MenuList[6].Num  := 5;
   MenuList[7].Menu := 'Q 終了   ';
   MenuList[7].Num  := 0;

{$ELSE NOJPN}
   MenuList[1].Menu := 'W Width   ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'D Decimal ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C Center  ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'R Right   ';
   MenuList[4].Num  := 4;
   MenuList[5].Menu := 'L Left    ';
   MenuList[5].Num  := 5;
   MenuList[6].Menu := 'O Color   ';
   MenuList[6].Num  := 6;
   MenuList[7].Menu := 'Q Quit    ';
   MenuList[7].Num  := 0;
{$ENDIF}
   hMenu :=  SelectVertMenu(1, h_MENULINE, MenuList, 7);
   CursorOn ;                  { CVCRT  }
   If hMenu = 1 Then
      ChangeWidth(ghTopCol, ghX) ;

   If hMenu = 2 Then
      ChangeDecPoint(ghTopCol, ghTopRow, ghEndCol, ghEndRow) ;

   If hMenu = 3 Then
      ChangeCellJustify(ghTopCol, ghTopRow, ghEndCol, ghEndRow, c_CELL_CENTER) ;

   If hMenu = 4 Then
      ChangeCellJustify(ghTopCol, ghTopRow, ghEndCol, ghEndRow, c_CELL_RIGHT) ;

   If hMenu = 5 Then
      ChangeCellJustify(ghTopCol, ghTopRow, ghEndCol, ghEndRow, c_CELL_LEFT) ;
   If hMenu = 6 Then
      ChangeCellColor(ghTopCol, ghTopRow, ghEndCol, ghEndRow) ;
   CursorOff ;                 { CVCRT  }
End ;
End .
