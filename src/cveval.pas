{$MODE TP}

Unit cveval ;

Interface

Uses cvdef, cvcrt ;
Function GetXString(xint: Integer):  String ;
Function GetXInteger(x: String):  Integer ;
Procedure ChangeFirstChar(Var sData:String) ;
Procedure RTrim(Var sData:String) ;
Procedure LTrim(Var sData:String) ;
Procedure Trim(Var sData:String) ;
Function UpcaseString(sData:String):  String ;
Function IsCellNumeric(Var tpPtr:tpCell):  Boolean ;
Function IsCellFormula(Var tpPtr:tpCell):  Boolean ;
Function IsCellRecalc(Var tpPtr:tpCell):  Boolean ;
Procedure SetCellFormula(Var tpPtr:tpCell) ;
Procedure SetCellNumeric(Var tpPtr:tpCell) ;
Procedure SetCellRecalc(Var tpPtr:tpCell) ;
Procedure SetCellNotFormula(Var tpPtr:tpCell) ;
Procedure SetCellNotNumeric(Var tpPtr:tpCell) ;
Procedure SetCellNotRecalc(Var tpPtr:tpCell) ;
Procedure GetCellArea(Var tpPtr: tpCell);
Procedure FreeCellArea(Var tpPtr: tpCell);
Function Hex(v: LongInt; w: Integer):  String;
Function H2I(s: String; Var hData:LongInt):  Integer;
Procedure NextCh(Var Position : Integer;
                 Var NextChar : Char;
                 Formula  : String);
Function GetCellXY(Var hPos: Integer;
                   Var Ch: Char;
                   Formula : String;
                   Var EFX, EFY :Integer ):  Integer;
Function GetRange(Var hPos: Integer;
                  Var Ch: Char;
                  Formula : String;
                  Var OldEFX, OldEFY, EFX, EFY :Integer ):  Integer;
Function GetErrMsg(hErrCode :Integer):  String;

Procedure Evaluate(Var hStatus  : Integer;
                   Var Formula  : String ;
                   Var fData   : Real;
                   Var hErrPos  : Integer;
                   Var sRet    : String;
                   curX, curY : Integer );

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Get X as String                               }
{ ---------------------------------------------- }
Function GetXString(xint: Integer):  String ;

Var 
   num :  Integer ;
   xstr :  String ;
Begin
   num := xint ;
   xstr := '' ;
   While (num > 26) Do
      Begin
         xstr := xstr + 'A' ;
         num := num - 26 ;
      End ;
   xstr := xstr + chr($40+num) ;
   GetXString := xstr ;
End ;

{ ---------------------------------------------- }
{  Get X as Integer                              }
{ ---------------------------------------------- }
Function GetXInteger(x: String):  Integer ;

Var 
   xint :  Integer ;
   i, len :  Integer ;
Begin
   xint := 0 ;
   len := length(x) ;
   For i:=1 To len-1  Do
      Begin
         If x[i] = 'A' Then
            xint := xint + 26
         Else
            Begin
               GetXInteger := 0 ;
               exit ;
            End ;
      End ;

   If Not (x[len] In ['A'..'Z']) Then
      Begin
         GetXInteger := 0 ;
         exit ;
      End ;

   i := 0;
   Repeat
      i := i+1 ;
   Until chr(i+$40) = x[len] ;
   xint := xint+i ;
   GetXInteger := xint ;
End ;

{ ---------------------------------------------- }
{  Delete Flag                                   }
{ ---------------------------------------------- }
Procedure ChangeFirstChar(Var sData:String) ;
Begin
   If copy(sData,1,Length(s_CELL_CENTER)) = s_CELL_CENTER Then
      delete(sData, 1, Length(s_CELL_CENTER))
   Else If copy(sData,1,Length(s_CELL_LEFT)) = s_CELL_LEFT Then
           delete(sData, 1, Length(s_CELL_LEFT))
   Else If copy(sData,1,Length(s_CELL_RIGHT)) = s_CELL_RIGHT Then
           delete(sData, 1, Length(s_CELL_RIGHT))
   Else If copy(sData,1,Length(s_CELL_FILLCHAR)) = s_CELL_FILLCHAR Then
           delete(sData, 1, Length(s_CELL_FILLCHAR)) ;

End ;

{ ---------------------------------------------- }
{  Triming Space                                 }
{ ---------------------------------------------- }
Procedure RTrim(Var sData:String) ;

Var 
   i:  Integer ;
Begin

   For i:=Length(sData) Downto 1 Do
      Begin
         If (Copy(sData, i, 1)<>' ') Then
            Begin
               sData := Copy(sData,1,i) ;
               exit ;
            End ;
      End ;
   sData := '' ;
End ;

Procedure LTrim(Var sData:String) ;

Var 
   i:  Integer ;
Begin
   For i:=1 To Length(sData) Do
      Begin
         If (Copy(sData, i, 1)<>' ') Then
            Begin
               sData := Copy(sData, i, Length(sData)-i+1) ;
               exit ;
            End ;
      End ;
   sData := '' ;
End ;

Procedure Trim(Var sData:String) ;
Begin
   RTrim(sData) ;
   LTrim(sData) ;
End ;

{ ---------------------------------------------- }
{  Upcase String                                 }
{ ---------------------------------------------- }
Function UpcaseString(sData:String):  String ;

Var 
   sStr:  String ;
   i:  Integer ;
   cCh :  char ;
Begin
   sStr := '' ;
   For i:=1 To length(sData) Do
      Begin
         cCh := sData[i] ;
         sStr := sStr+Upcase(cCh) ;
      End ;
   UpcaseString := sStr ;
End;

{ ---------------------------------------------- }
{  Check Cell Type                               }
{ ---------------------------------------------- }
Function IsCellNumeric(Var tpPtr:tpCell):  Boolean ;
Begin

   If tpPtr.tpMain = Nil Then
      Begin
         IsCellNumeric := False ;
         exit ;
      End ;

   If (tpPtr.cCellType And h_CELL_NUMERIC) <> 0 Then
      IsCellNumeric := True
   Else
      IsCellNumeric := False ;
End ;

Function IsCellFormula(Var tpPtr:tpCell):  Boolean ;
Begin
   If tpPtr.tpMain = Nil Then
      Begin
         IsCellFormula := False ;
         exit ;
      End ;

   If (tpPtr.cCellType And h_CELL_FORMULA) <> 0 Then
      IsCellFormula := True
   Else
      IsCellFormula := False ;
End ;

Function IsCellRecalc(Var tpPtr:tpCell):  Boolean ;
Begin
   If (tpPtr.cCellType And h_CELL_RECALC) <> 0 Then
      IsCellRecalc := True
   Else
      IsCellRecalc := False ;
End ;

{ ---------------------------------------------- }
{  Set Cell Type                                 }
{ ---------------------------------------------- }
Procedure SetCellFormula(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType Or h_CELL_FORMULA) ;
End ;

Procedure SetCellNumeric(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType Or h_CELL_NUMERIC) ;
End ;

Procedure SetCellRecalc(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType Or h_CELL_RECALC) ;
End ;

Procedure SetCellNotFormula(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType And $3f) ;
   //tpPtr.cCellType := (tpPtr.cCellType and $bf) ;
End ;

Procedure SetCellNotNumeric(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType And $7f) ;
   //tpPtr.cCellType := (tpPtr.cCellType and $bf) ;
End ;

Procedure SetCellNotRecalc(Var tpPtr:tpCell) ;
Begin
   tpPtr.cCellType := (tpPtr.cCellType And $df) ;
End ;

{ ---------------------------------------------- }
{  Get Cell Area                                 }
{ ---------------------------------------------- }
Procedure GetCellArea(Var tpPtr: tpCell);
Begin
   If tpPtr.tpMain = Nil Then
      New(tpPtr.tpMain)  ;
   tpPtr.tpMain^ := '' ;
End ;

{ ---------------------------------------------- }
{  Free Cell Area                                }
{ ---------------------------------------------- }
Procedure FreeCellArea(Var tpPtr: tpCell);
Begin
   If tpPtr.tpMain <> Nil Then
      Dispose(tpPtr.tpMain) ;
   tpPtr.tpMain := Nil ;
   SetCellNotFormula(tpPtr) ;
End ;

{ ---------------------------------------------- }
{  Integer to HEX string                         }
{ ---------------------------------------------- }
Function Hex(v: LongInt; w: Integer):  String;

Var 
   s               :  String;
   i               :  Integer;

Const 
   hexc            :  array [0 .. 15] Of Char =  '0123456789abcdef';
Begin
   s[0] := Chr(w);
   For i := w Downto 1 Do
      Begin
         s[i] := hexc[v And $F];
         v := v shr 4
      End;
   i := 0;
   Repeat
      i := i+1 ;
   Until s[i] <> '0' ;
   If i >0 Then
      delete(s,1,i-1);
   Hex := '0x'+s;
End;

{ ---------------------------------------------- }
{  HEX to Integer                                }
{ ---------------------------------------------- }
Function H2I(s: String; Var hData:LongInt):  Integer;

Var 
   hErrPos :  Integer ;
   s2    :  String ;
Begin
   If copy(s,1,2) = '0x' Then
      s2 := copy(s,3,length(s)-2)
   Else
      s2 := s ;

   val('$'+s2, hData, hErrPos);

   //	 GotoXy(10,9) ;
   //writeln('S2:',s2,' hData:',hData,' hErrPos:',hErrPos) ;
   H2I := hErrPos ;
End ;

{ ---------------------------------------------- }
{  Get Char                                      }
{ ---------------------------------------------- }
Procedure NextCh(Var Position : Integer;
                 Var NextChar : Char;
                 Formula  : String);
Begin
   Repeat
      Position := Position + 1;
      If Position <= Length(Formula) Then
         NextChar := Formula[Position]
      Else
         NextChar := Eofline;
   Until NextChar <> ' ';
End;

{ ---------------------------------------------- }
{  Get Cell XY                                   }
{ ---------------------------------------------- }
Function GetCellXY(Var hPos: Integer;
                   Var Ch: Char;
                   Formula : String;
                   Var EFX, EFY :Integer ):  Integer;

Var 
   xstr :  String ;
   hErr :  Integer ;
Begin
   If Ch='$' Then
      NextCh(hPos,Ch,Formula) ;

   xstr := Ch ;
   NextCh(hPos,Ch,Formula) ;
   While (Ch In ['A'..'Z','a'..'z']) Do
      Begin
         xstr := xstr + Ch ;
         NextCh(hPos,Ch,Formula) ;
      End ;
   EFX := GetXInteger(UpcaseString(xstr)) ;

  { Validate COL 1}
   If EFX = 0 Then
      Begin
         GetCellXY := -1 ;
         exit ;
      End ;

   If Ch='$' Then
      NextCh(hPos,Ch,Formula)  ;

  { Validate ROW 1}
   If Not (Ch In Numbers) Then
      Begin
         GetCellXY := -1 ;
         exit ;
      End ;

   xstr := '' ;
   While (Ch In Numbers) And (Ch <> Eofline) Do
      Begin
         xstr := xstr+Ch;
         NextCh(hPos,Ch,Formula) ;
      End ;
   Val(xstr,EFY,hErr);
   GetCellXY := hErr ;
End ;

{ ---------------------------------------------- }
{  Get Char                                      }
{ ---------------------------------------------- }
Function GetRange(Var hPos: Integer;
                  Var Ch: Char;
                  Formula : String;
                  Var OldEFX, OldEFY, EFX, EFY :Integer ):  Integer;
Begin
   If GetCellXY(hPos,Ch,Formula,OldEFX,OldEFY)<> 0 Then
      Begin
         GetRange := -1 ;
         exit ;
      End ;

   If Ch<>'.' Then
      Begin
         GetRange := 1 ;
         exit ;
      End ;

   NextCh(hPos,Ch,Formula)  ;
   // Check '..'
   If Ch<>'.' Then
      Begin
         GetRange := -1 ;
         exit ;
      End ;

   NextCh(hPos,Ch,Formula)  ;

   If GetCellXY(hPos,Ch,Formula,EFX,EFY)<> 0 Then
      Begin
         GetRange := -1 ;
         exit ;
      End ;
   GetRange := 0 ;
End ;

{ ---------------------------------------------- }
{  Get Error Message                             }
{ ---------------------------------------------- }
Function GetErrMsg(hErrCode :Integer):  String;

Type 
   ErrCodeList =  array[1..25] Of integer ;
   ErrMsgList  =  array[1..25] Of String ;

Const 
   ErrCode:  ErrCodeList =  (
                             h_ERR_DEFINE_FUNC,
                             h_ERR_NOT_NUMERIC,
                             h_ERR_NOT_EXPRESSION,
                             h_ERR_RANGE,
                             h_ERR_FIRST_CELL_OVERLAPED,
                             h_ERR_SECOND_CELL_OVERLAPED,
                             h_ERR_CELL_IVALID_ORDER,
                             h_ERR_H2I_CONVERSION,
                             h_ERR_ZERO_DIVIDE,
                             h_ERR_VLOOKUP_STRING,
                             h_ERR_VLOOKUP_KEY,
                             h_ERR_VLOOKUP_KEYVALUE,
                             h_ERR_VLOOKUP_KEYCELL,
                             h_ERR_VLOOKUP_KEYBLANK,
                             h_ERR_VLOOKUP_KEY_NOTNUMERIC,
                             h_ERR_VLOOKUP_NOTSEP,
                             h_ERR_VLOOKUP_RANGE,
                             h_ERR_VLOOKUP_COLNUM,
                             h_ERR_VLOOKUP_EVAL_RANGE,
                             h_ERR_VLOOKUP_NOTARGET,
                             h_ERR_ATTR_NOTMACH,
                             h_ERR_VIF_STRING,
                             h_ERR_VIF_NOTSEP,
                             h_ERR_VIF_LOGICAL_ERR,
                             h_ERR_NOT_COMPLETED) ;

   ErrMsg:  ErrMsgList =  (
                           'INVALID FUNC',
                           'NOT_NUMERIC',
                           'NOT_EXPRESSION',
                           'INVALID RANGE',
                           'FIRST_CELL_OVERLAPED',
                           'SECOND_CELL_OVERLAPED',
                           'IVALID_CELL_ORDER',
                           'HEX2INT_CONVERSION_ERROR',
                           'ZERO_DIVIDE',
                           'VLOOKUP:INVALID STRING',
                           'VLOOKUP:INVALID KEY',
                           'VLOOKUP:INVALID KEYVALUE',
                           'VLOOKUP:INVALID KEYCELL',
                           'VLOOKUP:KEYBLANK',
                           'VLOOKUP:KEY_NOTNUMERIC',
                           'VLOOKUP:NOTSEP',
                           'VLOOKUP:INVALID RANGE',
                           'VLOOKUP:INVALID COLMUN',
                           'VLOOKUP:INVALID EVAL_RANGE',
                           'VLOOKUP:NO_TARGET',
                           'ATTR_NOTMACH',
                           'IF:INVALID STRING',
                           'IF:NOTSEP',
                           'IF:LOGICAL_ERR',
                           'FORMULA_NOT_COMPLETED') ;

Var 
   i    :  Integer;
   sErrMsg :  String ;

   Label BPOINT ;
Begin
   sErrMsg := '' ;
   For i:=1 To length(ErrCode) Do
      Begin
         If hErrCode = ErrCode[i] Then
            Begin
               sErrMsg := ErrMsg[i]  ;
               goto BPOINT ;
            End ;
      End ;

   BPOINT:
            //writeln('ErrCode:', hErrCode, ' ErrMsg:', sErrMsg) ;
            GetErrMsg := sErrMsg ;
End;

{ ---------------------------------------------- }
{  Evaluate Formula                              }
{ ---------------------------------------------- }
Procedure Evaluate2(Var hStatus  : Integer;
                    Var Formula  : String ;
                    Var fData   : Real;
                    Var hErrPos  : Integer;
                    Var sRet    : String;
                    curX, curY : Integer);

Var 
   hPos:  Integer;    { Current position in formula                     }
   Ch:  Char;        { Current character being scanned                 }
   fCellSum, fCellMax, fCellMin, fCellCnt, fCellCountA:  Real;
   sData:  String ;
   hStartPos :  Integer ;

Function LogicalExpression:  Real;

Var 
   L  :  Real;
   L2  :  Real;
   sPrevRet  :  String ;
   hPrevStatus  :  Integer ;
   xstr :  String;

   Label BPOINT;

Function Expression:  Real;

Var 
   E:  Real;
   E2:  Real;
   Opr:  Char;
   sPrevRet  :  String ;
   hPrevStatus  :  Integer ;

Function SimpleExpression:  Real;

Var 
   S  :  Real;
   Opr :  Char;
   S2  :  real ;
Function Term:  Real;

Var 
   T:  Real;
   T2:  Real;

Function SignedFactor:  Real;

Function Factor:  Real;

Var 
   F        :  Real;
   OldEFX, OldEFY :  Integer;
   EFX, EFY    :  Integer;
   SumFX, SumFY  :  Integer;
   Start,Err    :  Integer;

   hRet      :  Integer ;
   hResult      :  LongInt ;

   Label BPOINT;

Function DefinedFunctions:  Real;

Type 
   StandardFunction =  (fabs,fsqrt,fsqr,fsin,fcos,
                        farctan,fln,flog,fexp,ffact,
                        fint,fpi,ftrunc,fround,fmax,
                        fmin,favg,fcount,fcounta,fsum,
                        fhex,fstr);
   StandardFunctionList =  array[StandardFunction] Of string[8];

Const 
   StandardFunctionNames:  StandardFunctionList 
                           =  ('@ABS','@SQRT','@SQR','@SIN','@COS',
                               '@ARCTAN','@LN','@LOG','@EXP','@FACT','@INT','@PI',
                               '@TRUNC','@ROUND',
                               '@MAX','@MIN','@AVG','@NCOUNT','@COUNT','@SUM', '@HEX', '@STR');

Var 
   U:  Real;
   L:  Integer;       { intermidiate variables }
   Found:  Boolean;
   Sf:  StandardFunction;
   sResult  :  String ;

Function Fact(I: Integer):  Real;
Begin
   If I > 0 Then
      Begin
         Fact := I*Fact(I-1);
      End
   Else
      Fact := 1;
End  { Fact };

Begin { DefinedFunctions }
   found := false;
   For sf:=fabs To fstr Do
      If Not found Then
         Begin
            l := Length(StandardFunctionNames[sf]);
            If UpcaseString(copy(Formula,hPos,l))=StandardFunctionNames[sf] Then
               Begin
                  hPos := hPos+l-1;
                  NextCh(hPos,Ch,Formula) ;

                  fCellMax := 0;
                  fCellMin := 0;
                  fCellCnt := 0;
                  fCellCountA := 0;
                  fCellSum := 0;

                  u := Factor;

                  sRet := '' ;
                  Case sf Of 
                     fabs :  u := abs(u);
                     fsqrt :
                              Begin
                                 If u > 0 Then
                                    u := sqrt(u)
                                 Else
                                    u := -1;
                              End ;
                     fsqr :
                             Begin
                                If Abs(u) < Sqrt(exp(38) * ln(1)) Then
                                   u := sqr(u)
                                Else
                                   u := -1;
                             End;

                     fsin :  u := sin(u);
                     fcos :  u := cos(u);
                     farctan :  u := arctan(u);
                     fln  :  u := ln(u);
                     flog :  u := ln(u)/ln(10);
                     fexp :  u := exp(u);
                     ffact :  u := fact(trunc(u));
                     fint :  u := int(u);
                     fpi  :  u := pi ;

                     fround :  u := round(u);
                     ftrunc :  u := trunc(u);
                     fmax :  If fCellMax > 0 Then
                                u := fCellMax;
                     fmin :  If fCellMin > 0 Then
                                u := fCellMin;
                     favg :  If fCellCnt > 0 Then
                                u := fCellSum/fCellCnt ;
                     fcount :
                               u := fCellCnt ;
                     fcounta :
                                u := fCellCountA ;
{
															fsum :
                             u:=u ;
}
                     fhex :
                             Begin
                                sRet := sRet+Hex(round(u), 8) ;
                                //u:= 0;
                                hStatus := h_FORMULA_STRING;
                             End ;
                     fstr :
                             Begin
                                //															str(round(u), sResult) ;
                                str(u:gcCellWidth[ghX]:GetCellDecPoint(ghX, ghY),sResult);
                                Trim(sResult) ;
                                sRet := sRet+sResult ;
                                //u:= 0;
                                hStatus := h_FORMULA_STRING;
                             End ;

                  End; { case }
                  Found := true;
               End;
         End;
   If Not Found Then
      Begin

         //	 GotoXy(10,12) ;
         // writeln('Not Found Formula hPos:',hPos,' Ch:',Ch,' hErrPos:',hErrPos) ;
         hStatus := h_ERR_DEFINE_FUNC;
         hErrPos := hPos;
         exit ;
      End ;
   DefinedFunctions := U;
End; { function DefinedFunctions}

Begin { Function Factor }

  { Constant value like "abc"}
   If Ch='"' Then
      Begin
         Start := hPos;
         NextCh(hPos,Ch,Formula);
         While (Ch <> '"') And (Ch <> Eofline) Do
            Begin
               NextCh(hPos,Ch,Formula);
            End ;

         sRet := Copy(Formula,Start+1,hPos-Start-1) ;
         hStatus := h_FORMULA_STRING ;
         NextCh(hPos,Ch,Formula) ;
         goto BPOINT ;
      End ;

  { Numeric value }
   If Ch In Numbers Then
      Begin
         Start := hPos;
         Repeat
            NextCh(hPos,Ch,Formula)
         Until Not (Ch In Numbers);

         If Upcase(Ch)='X' Then
            Begin
               Start := hPos;
               Repeat
                  NextCh(hPos,Ch,Formula)
               Until Not (Ch In HexDigit);
               hErrPos := H2I(Copy(Formula,Start+1,hPos-Start-1),hResult) ;
               //	 GotoXy(10,10) ;
               //	 writeln('hPos:',hPos,' Start:',Start,' hRet:',hRet,' hErrPos:',hErrPos) ;
               If hErrPos <> 0 Then
                  Begin
                     hStatus := h_ERR_NOT_NUMERIC;
                     hErrPos := hPos;
                     exit ;
                  End ;

               F := hResult ;
               hStatus := h_FORMULA_NUMERIC ;
               goto BPOINT ;
            End ;

         If Ch='.' Then
            Repeat
               NextCh(hPos,Ch,Formula)
            Until Not (Ch In Numbers);

         If Upcase(Ch)='E' Then
            Begin
               NextCh(hPos,Ch,Formula) ;
               Repeat
                  NextCh(hPos,Ch,Formula)
               Until Not (Ch In Numbers);
            End;

         Val(Copy(Formula,Start,hPos-Start),F,hErrPos);
         If hErrPos <> 0 Then
            Begin
               hStatus := h_ERR_NOT_NUMERIC;
               hErrPos := hPos;
               exit ;
            End ;

         hStatus := h_FORMULA_NUMERIC ;
         goto BPOINT ;
      End ;

  { (expr) }
   If Ch='(' Then
      Begin
         NextCh(hPos,Ch,Formula) ;
         //         F := Expression;
         F := LogicalExpression;
         If Ch=')' Then NextCh(hPos,Ch,Formula)
         Else
            Begin
               hStatus := h_ERR_NOT_EXPRESSION;
               hErrPos := hPos;
               exit ;
            End ;
         goto BPOINT ;
      End ;

  { Cell reference }
   If (Ch In ['A'..'Z','a'..'z']) Or (Ch='$') Then
      Begin
         F := 0;
         hRet := GetRange(hPos,Ch,Formula,OldEFX,OldEFY,EFX,EFY);

         If hRet = -1 Then
            Begin
               hStatus := h_ERR_RANGE;
               hErrPos := hPos;
               //writeln('Formula:',Formula,' hRet:',hRet,' hErrPos:',hErrPos) ;

               exit ;
            End ;

         { @@@@ Validate Range }
         If (curX = OldEFX) And (curY = OldEFY) Then
            Begin
               hStatus := h_ERR_FIRST_CELL_OVERLAPED;
               hErrPos := hPos;
               exit ;
            End ;

         If sheet[OldEFX,OldEFY].tpMain <> Nil Then
            Begin
               sData := sheet[OldEFX,OldEFY].tpMain^ ;
               ChangeFirstChar(sData) ;

               If IsCellFormula(sheet[OldEFX,OldEFY])  Then
                  Begin
                     Evaluate2(hStatus,sData,f,hErrPos,sRet,curX,curY);
                     SetCellRecalc(sheet[OldEFX,OldEFY]) ;
                  End
               Else If IsCellNumeric(sheet[OldEFX,OldEFY]) Then
                       Begin
                          val(sData,F,Err) ;
                          hStatus := h_FORMULA_NUMERIC;
                          SetCellRecalc(sheet[OldEFX,OldEFY]) ;
                       End
               Else If UpcaseString(copy(sData,1,2)) = '0X' Then
                       Begin
                          hErrPos := H2I(Copy(sData,3,length(sData)-2),hResult) ;
                          If hErrPos <> 0 Then
                             Begin
                                hStatus := h_ERR_NOT_NUMERIC;
                                hErrPos := hPos;
                                exit ;
                             End ;

                          F := hResult ;
                          SetCellRecalc(sheet[OldEFX,OldEFY]) ;
                          hStatus := h_FORMULA_NUMERIC ;
                       End

               Else
                  Begin
                     If sData <> '' Then
                        Begin
                           sRet := sData ;
                           SetCellRecalc(sheet[OldEFX,OldEFY]) ;
                           hStatus := h_FORMULA_STRING ;
                        End ;
                  End ;
            End
         Else
            Begin
               hStatus := h_FORMULA_NUMERIC ;
            End;

         If hRet = 1 Then
            goto BPOINT ;

         { @@@@ Validate Range }
         If (curX >= oldEFX) And (curX <= EFX) And (curY >= oldEFY) And (curY <= EFY) Then
            Begin
               hStatus := h_ERR_SECOND_CELL_OVERLAPED;
               hErrPos := hPos;
               exit ;
            End ;

         If (EFX < oldEFX) Or (EFY < oldEFY) Then
            Begin
               hStatus := h_ERR_CELL_IVALID_ORDER;
               hErrPos := hPos;
               exit ;
            End ;
         { ^^^^^^^^^^^^^^^^^^}

         fCellSum := 0;
         For SumFY:=OldEFY To EFY Do
            Begin
               For SumFX:=OldEFX To EFX Do
                  Begin
                     F := 0;
                     If sheet[SumFX,SumFY].tpMain <> Nil Then
                        Begin
                           sData := sheet[SumFX,SumFY].tpMain^ ;
                           ChangeFirstChar(sData) ;
                           If IsCellFormula(sheet[SumFX,SumFY]) Then
                              Begin
                                 Evaluate(hStatus,sData,f,hErrPos,sRet,curX,curY);
                                 SetCellRecalc(sheet[SumFX,SumFY]) ;
                              End
                           Else If IsCellNumeric(sheet[SumFX,SumFY]) Then
                                   Begin
                                      val(sData,F,Err) ;
                                      //hStatus := h_FORMULA_NUMERIC; 
                                      SetCellRecalc(sheet[SumFX,SumFY]) ;
                                   End
                           Else If UpcaseString(copy(sData,1,2)) = '0X' Then
                                   Begin
                                      hErrPos := H2I(Copy(sData,3,length(sData)-2),hResult) ;
                                      If hErrPos <> 0 Then
                                         Begin
                                            hStatus := h_ERR_NOT_NUMERIC;
                                            hErrPos := hPos;
                                            exit ;
                                         End ;

                                      F := hResult ;
                                      SetCellRecalc(sheet[SumFX,SumFY]) ;
                                      hStatus := h_FORMULA_NUMERIC ;
                                   End ;

                           If IsCellNumeric(sheet[SumFX,SumFY]) Then
                              fCellCnt := fCellCnt+1;
                           fCellCountA := fCellCountA+1 ;
                        End ;

{
                     If sheet[SumFX,SumFY].tpMain <>Nil Then
                        Begin
                           If IsCellNumeric(sheet[SumFX,SumFY]) Then
                              fCellCnt := fCellCnt+1;
                           fCellCountA := fCellCountA+1 ;
                        End ;
}
                     // Cumulative Formula
                     If (IsCellNumeric(sheet[SumFX,SumFY])) Or (copy(sData,1,2) = '0x') Then
                        Begin
                           If fCellSum=0 Then
                              Begin
                                 fCellMax := f ;
                                 fCellMin := f ;
                              End
                           Else
                              Begin
                                 If f>fCellMax Then
                                    fCellMax := f ;
                                 If f<fCellMin Then
                                    fCellMin := f ;
                              End ;

                           fCellSum := fCellSum+f;
                           f := fCellSum;
                           hStatus := h_FORMULA_NUMERIC;
                        End
                     Else
                        f := fCellSum;
                  End ;
            End;
         //hStatus := h_FORMULA_NUMERIC; 
         goto BPOINT ;
      End;

   f := DefinedFunctions ;

   BPOINT:
            Factor := F
End ; { function Factor}

Begin { SignedFactor }
   If Ch='-' Then
      Begin
         NextCh(hPos,Ch,Formula) ;
         SignedFactor := -Factor;
      End
   Else
      SignedFactor := Factor;
End { SignedFactor };

Begin { Term }
   T := SignedFactor;
   While Ch='^' Do
      Begin
         NextCh(hPos,Ch,Formula) ;
         t2 := SignedFactor;
         If hStatus <> h_FORMULA_NUMERIC Then
            Begin
               hStatus := h_ERR_NOT_NUMERIC;
               hErrPos := hPos ;
               exit;
            End ;
         //t := exp(ln(t)*SignedFactor);
         t := exp(ln(t)*t2);
      End;
   Term := t;
End { Term };


Begin { SimpleExpression }
   s := term;
   While Ch In ['*','/'] Do
      Begin
         Opr := Ch;

         If hStatus = h_FORMULA_STRING Then
            Begin
               hStatus := h_ERR_NOT_NUMERIC;
               hErrPos := hPos ;
               exit;
            End ;
         NextCh(hPos,Ch,Formula) ;
         Case Opr Of 
            '*':
                  Begin
                     s2 := term ;
                     If hStatus = h_FORMULA_STRING Then
                        Begin
                           hStatus := h_ERR_NOT_NUMERIC;
                           hErrPos := hPos ;
                           exit;
                        End ;
                     s := s*s2;
                  End ;
            '/':
                  Begin
                     s2 := term ;
                     If hStatus = h_FORMULA_STRING Then
                        Begin
                           hStatus := h_ERR_NOT_NUMERIC;
                           hErrPos := hPos ;
                           exit;
                        End
                     Else If s2 = 0 Then
                             Begin
                                hStatus := h_ERR_ZERO_DIVIDE;
                                hErrPos := hPos ;
                                exit;
                             End
                     Else
                        s := s/s2;
                  End ;
         End;
      End;
   SimpleExpression := s;
End { SimpleExpression };

Begin { Expression }
   E := SimpleExpression;
   While Ch In ['+','-'] Do
      Begin
         Opr := Ch;
         hPrevStatus := hStatus ;
         sPrevRet := sRet ;

         If (hStatus = h_FORMULA_STRING) And (Opr = '-') Then
            Begin
               //	 GotoXY(10,10) ;								 

           //	 writeln('Expression ',hStatus,' ', Formula,' fData:',E,' hPos:',hPos,' sRet:',sRet) ;

               hStatus := h_ERR_NOT_NUMERIC;
               hErrPos := hPos ;
               exit;
            End ;

         NextCh(hPos,Ch,Formula) ;
         Case Opr Of 
            '+':
                  Begin
                     e2 := SimpleExpression;
                     If hStatus <> hPrevStatus Then
                        //if hStatus = h_FORMULA_STRING then
                        Begin
                           //hStatus := h_ERR_NOT_NUMERIC;
                           hStatus := h_ERR_ATTR_NOTMACH ;
                           hErrPos := hPos ;
                           exit;
                        End ;
                     If hStatus = h_FORMULA_STRING Then
                        sRet := sPrevRet+sRet
                     Else
                        e := e+e2;
                  End ;
            '-':
                  Begin
                     e2 := SimpleExpression;

                     If hStatus = h_FORMULA_STRING Then
                        Begin
                           hStatus := h_ERR_NOT_NUMERIC;
                           hErrPos := hPos ;
                           exit;
                        End;

                     e := e-e2 ;
                  End ;
         End;
      End;
   Expression := E;
End { Expression };

Begin { LogicalExpression }
   L := Expression;
   hPrevStatus := hStatus ;
   sPrevRet := sRet ;
   If Ch In ['>','<','='] Then
      Begin
         xstr := Ch ;
         NextCh(hPos,Ch,Formula) ;
         If Ch In ['>','='] Then
            Begin
               xstr := xstr+Ch;
               NextCh(hPos,Ch,Formula) ;
            End ;

         L2 := Expression;
         If hStatus <> hPrevStatus Then
            Begin
               hStatus := h_ERR_ATTR_NOTMACH;
               hErrPos := hPos ;
               exit;
            End;

         If xstr = '>'Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L > L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet > sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End ;

         If xstr = '<'Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L < L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet < sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End ;

         If xstr = '='Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L = L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet = sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End ;

         If xstr = '>='Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L >= L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet >= sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End ;

         If xstr = '<='Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L <= L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet <= sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End ;

         If xstr = '<>'Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  Begin
                     If L <> L2 Then
                        L := 1
                     Else
                        L := 0 ;
                  End
               Else
                  Begin
                     If sPrevRet <> sRet Then
                        L := 1
                     Else
                        L := 0 ;
                  End ;
               hStatus := h_FORMULA_NUMERIC;
               goto BPOINT ;
            End;
         BPOINT:
      End;
   LogicalExpression := L;
End { Logical Expression };


Begin { procedure Evaluate2 }

   ChangeFirstChar(Formula) ;
   If Formula[1]='.' Then
      Formula := '0'+Formula;
   If Formula[1]='+' Then
      delete(Formula,1,1);

   hStatus := 0;
   hErrPos := 0;
   hPos := 0;

   NextCh(hPos,Ch,Formula) ;
   fData := LogicalExpression;

   // Not completed
   If Ch <> Eofline Then
      Begin
         If hStatus < h_ERR_DEFINE_FUNC Then
            hStatus := h_ERR_NOT_COMPLETED;
         hErrPos := hPos ;
         exit;
      End ;

End { Evaluate2 };

{ ---------------------------------------------- }
{  IF                                            }
{ ---------------------------------------------- }
Procedure VIF(Var hStatus: Integer; Var Formula: String; Var fData: Real; Var hErrPos: Integer; Var
              sRet: String; curX, curY : Integer );

Var 
   Ch    :  Char;
   sLogical :  String ;
   sTrue  :  String ;
   sFalse  :  String ;
   hPos   :  Integer ;
   hLen   :  Integer ;
   //	 fAnswer	: Real ;
Begin
   hPos := 1 ;
   hLen := Length('@IF') ;
   If UpcaseString(copy(Formula,hPos,hLen)) <> '@IF' Then
      Begin
         hStatus := h_ERR_VIF_STRING ;
         hErrPos := hPos ;
         exit
      End ;
   hPos := hPos+hLen-1 ;
   NextCh(hPos,Ch,Formula) ;
   If Ch <> '(' Then
      Begin
         hStatus := h_ERR_VIF_NOTSEP ;
         hErrPos := hPos ;
         exit
      End ;

   //@@@@ Logical
   NextCh(hPos,Ch,Formula) ;
   sLogical := '' ;
   While (Ch <> ',') And (Ch <> Eofline) Do
      Begin
         sLogical := sLogical+Ch;
         NextCh(hPos,Ch,Formula) ;
      End ;

   If Ch <> ',' Then
      Begin
         hStatus := h_ERR_VIF_NOTSEP ;
         hErrPos := hPos ;
         exit;
      End;

   // @@@TRUE
   NextCh(hPos,Ch,Formula) ;
   sTrue := '' ;
   While (Ch <> ',') And (Ch <> Eofline) Do
      //While (Ch <> ',') Do
      Begin
         sTrue := sTrue+Ch;
         NextCh(hPos,Ch,Formula) ;
      End ;

   If Ch <> ',' Then
      Begin
         hStatus := h_ERR_VIF_NOTSEP ;
         hErrPos := hPos ;
         exit;
      End;

   // @@@FALSE
   NextCh(hPos,Ch,Formula) ;
   sFalse := '' ;
   While (Ch <> ')') And (Ch <> Eofline) Do
      Begin
         sFalse := sFalse+Ch;
         NextCh(hPos,Ch,Formula) ;
      End ;
   If Ch <> ')' Then
      Begin
         hStatus := h_ERR_VIF_NOTSEP ;
         hErrPos := hPos ;
         exit;
      End;

   Evaluate2(hStatus, sLogical, fData, hErrPos, sRet, curX, curY) ;
   If hErrPos <> 0 Then
      Begin
         hStatus := h_ERR_VIF_LOGICAL_ERR ;
         hErrPos := hPos ;
         exit;
      End;

   If fData >= h_LOGICAL_TRUE Then
      Begin
         Evaluate2(hStatus, sTrue, fData, hErrPos, sRet, curX, curY) ;
         If hErrPos <> 0 Then
            Begin
               hStatus := h_ERR_VIF_TRUE_FIELD_ERR ;
               hErrPos := hPos ;
               exit;
            End;
      End
   Else
      Begin
         Evaluate2(hStatus, sFalse, fData, hErrPos, sRet, curX, curY) ;
         If hErrPos <> 0 Then
            Begin
               hStatus := h_ERR_VIF_FALSE_FIELD_ERR ;
               hErrPos := hPos ;
               exit;
            End;
      End ;
   //NextCh(hPos,Ch,Formula) ;
   hErrPos := 0 ;

End ;

{ ---------------------------------------------- }
{  VLOOKUP                                       }
{ ---------------------------------------------- }
Function VLookup(Var hStatus :Integer; Var hPos :Integer; Formula: String; Var hErrPos: Integer;curX
                 , curY : Integer):  String ;

Var 
   hLen, hRet,hCol, hTarget           :  Integer;
   Ch                      :  Char;
   KeyEFX, KeyEFY, OldEFX, OldEFY, EFX, EFY,X,Y :  Integer;
   xstr                     :  String ;
   fData                    :  Real;
   sRet                     :  String ;
   fKey                     :  Real ;
   sKey                     :  String ;

Begin
   hLen := Length('@VLOOKUP') ;
   If UpcaseString(copy(Formula,hPos,hLen)) <> '@VLOOKUP' Then
      Begin
         hStatus := h_ERR_VLOOKUP_STRING ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit
      End ;
   hPos := hPos+hLen-1 ;
   NextCh(hPos,Ch,Formula) ;
   If Ch <> '(' Then
      Begin
         hStatus := h_ERR_VLOOKUP_KEY ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit
      End ;

   //@@@@ Key
   NextCh(hPos,Ch,Formula) ;
   fKey := 0 ;
   sKey := '' ;

   If Ch In Numbers Then
      Begin
         xstr := '' ;
         While (Ch <> ',') And (Ch <> Eofline) Do
            Begin
               xstr := xstr+Ch;
               NextCh(hPos,Ch,Formula) ;
            End ;
         If Ch <> ',' Then
            Begin
               hStatus := h_ERR_VLOOKUP_NOTSEP ;
               hErrPos := hPos ;
               VLookup := 'ERR' ;
               exit;
            End;

         Evaluate2(hStatus, xstr, fData, hErrPos, sRet, curX, curY) ;
         If hErrPos <> 0 Then
            Begin
               hStatus := h_ERR_VLOOKUP_KEYVALUE ;
               hErrPos := hPos ;
               VLookup := 'ERR' ;
               exit;
            End;
         fKey := fData ;
      End
   Else
      Begin
         hRet := GetCellXY(hPos,Ch,Formula,KeyEFX,KeyEFY) ;
         If hRet <> 0 Then
            Begin
               hStatus := h_ERR_VLOOKUP_KEYCELL ;
               hErrPos := hPos ;
               VLookup := 'ERR' ;
               exit
            End;

         If Sheet[KeyEFX, KeyEFY].tpMain = Nil Then
            Begin
               hStatus := h_ERR_VLOOKUP_KEYBLANK ;
               hErrPos := hPos ;
               VLookup := 'ERR' ;
               exit
            End;

         If IsCellNumeric(Sheet[KeyEFX, KeyEFY]) Then
            Begin
               Evaluate2(hStatus, Sheet[KeyEFX, KeyEFY].tpMain^, fData, hErrPos, sRet, curX, curY) ;
               If hErrPos <> 0 Then
                  Begin
                     hStatus := h_ERR_VLOOKUP_KEY_NOTNUMERIC ;
                     hErrPos := hPos ;
                     VLookup := 'ERR' ;
                     exit ;
                  End;
               fKey := fData ;
            End
         Else
            sKey := Sheet[KeyEFX, KeyEFY].tpMain^ ;
      End ;

   If Ch <> ',' Then
      Begin
         hStatus := h_ERR_VLOOKUP_NOTSEP ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit;
      End;

   // @@@Range
   NextCh(hPos,Ch,Formula) ;
   hRet := GetRange(hPos,Ch,Formula,OldEFX,OldEFY,EFX,EFY);
   If Ch <> ',' Then
      Begin
         hStatus := h_ERR_VLOOKUP_RANGE ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit;
      End;

   // @@@ColNumber 
   NextCh(hPos,Ch,Formula) ;
   xstr := '' ;
   While (Ch <> ')') And (Ch <> Eofline) Do
      //	 While (Ch <> ')')  Do
      Begin
         xstr := xstr+Ch;
         NextCh(hPos,Ch,Formula) ;
      End ;

   If Ch <> ')' Then
      Begin
         hStatus := h_ERR_VLOOKUP_COLNUM ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit;
      End;

   Evaluate2(hStatus, xstr, fData, hErrPos, sRet, curX, curY) ;
   If hErrPos <> 0 Then
      Begin
         hStatus := h_ERR_VLOOKUP_EVAL_RANGE ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
         exit;
      End;

   hCol := round(fData) ;

   //	 GotoXY(20,11) ;

//	 writeln('KeyEFX:',KeyEFX,' KeyEFY:',KeyEFY,' fKey:',round(fKey),' sKey:',sKey,' OldEFX:',OldEFX,' OldEFY:',OldEFY,' EFX:',EFX,' EFY:',EFY,' hCol:',hCol) ;
   hTarget := 0 ;
   For Y:=OldEFY To EFY Do
      Begin
         If (Sheet[OldEFX, Y].tpMain <> Nil) Then
            Begin
               If (fKey <> 0) Then
                  Begin
                     Evaluate2(hStatus, Sheet[OldEFX, Y].tpMain^, fData, hErrPos, sRet, curX, curY)
                     ;

                     If (hErrPos = 0) And (fKey = fData) Then
                        hTarget := Y ;

                  End
               Else
                  Begin
                     If sKey = Sheet[OldEFX, Y].tpMain^ Then
                        hTarget := Y ;
                  End ;

               For X:=OldEFX To EFX Do
                  Begin
                     SetCellRecalc(sheet[X,Y]) ;
                  End;
            End ;
      End ;

   // writeln('Target:',hTarget,' hCol:',hCol) ;

   If hTarget <> 0 Then
      Begin
         SetCellRecalc(Sheet[OldEFX+hCol, hTarget]) ;

         VLookup := Sheet[OldEFX+hCol, hTarget].tpMain^ ;
         hStatus := h_FORMULA_STRING ;
      End
   Else
      Begin
         hStatus := h_ERR_VLOOKUP_NOTARGET ;
         hErrPos := hPos ;
         VLookup := 'ERR' ;
      End ;
End ;

{ ---------------------------------------------- }
{  Evaluate                                      }
{ ---------------------------------------------- }
Procedure Evaluate(Var hStatus  : Integer;
                   Var Formula  : String ;
                   Var fData   : Real;
                   Var hErrPos  : Integer;
                   Var sRet    : String;
                   curX, curY : Integer );

Var 
   hPos     :  Integer;
   Ch      :  Char;
   //   Form					: Boolean;
   sData    :  String ;
   sSaveFormula :  String ;
   hStartPos  :  Integer ;
   bPrevFormula    :  Boolean ;
Begin
   sSaveFormula := Formula ;
   ChangeFirstChar(Formula) ;
   If Formula[1]='.' Then
      Formula := '0'+Formula;
   If Formula[1]='+' Then
      delete(Formula,1,1);

   //	 sSaveFormula := Formula ;
   bPrevFormula := false;
   hStatus := 0 ;
   hPos := 0;
   hErrPos := 0;
   sRet := '' ;

  { @IF }
   If (UpcaseString(copy(Formula,1,3))='@IF') Then
      Begin
         VIF(hStatus, Formula, fData, hErrPos, sRet, curX, curY);
         exit ;
      End ;

  { @VLOOKUP }
   hStartPos := pos('@VLOOKUP',UpcaseString(Formula));
   If hStartPos >0 Then
      Begin
         Repeat
            hPos := hStartPos ;
            sData := VLookup(hStatus, hPos, Formula, hErrPos, curX, curY) ;
            If sData = 'ERR' Then
               exit ;

            ChangeFirstChar(sData) ;
            delete(Formula, hStartPos,hPos-hStartPos+1);
            insert(sData, Formula, hStartPos) ;
            hPos := hStartPos + Length(sData) ;

            bPrevFormula := true;
            hStartPos := pos('@VLOOKUP',UpcaseString(Formula));

         Until hStartPos = 0 ;
      End ;
   If bPrevFormula Then
      sRet := Formula ;

   Evaluate2(hStatus,Formula,fData,hErrPos,sRet,curX, curY);

   If (bPrevFormula) And (hErrPos <> 0) And (sRet <> '') Then
      Begin
         hStatus := h_FORMULA_STRING;
         hErrPos := 0 ;
      End ;

   If (bPrevFormula) And (hStatus = h_FORMULA_NUMERIC) And (sRet <> '') Then
      sRet := '' ;

   Formula := sSaveFormula ;
End ;

End.
