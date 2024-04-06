{$MODE TP}

Unit cvscrn ;

Interface

Uses cvdef, cvcrt, cveval ;
Function SelectVertMenu(X1, Y1: byte; ListMenu : Menus; Count : byte):  byte;
Procedure WriteRange;
Procedure Centerlize(Var sWork: String; hLen: Integer);
Procedure RightJustify(Var sWork: String; hLen: Integer);
Procedure LeftJustify(Var sWork: String; hLen: Integer);
Function SetData(Var tpPtr:tpCell):  String ;
Procedure ClearMenuLine;
Procedure WriteData(Var tpPtr:tpCell; X,Y,Width:Integer; cAttr: Byte);
Procedure SetCurrentDataOnCell(Var tpPtr: tpCell; cAttr: Byte);
Procedure SetCurrentDataOnRow;
Procedure SetCurrentDataOnLine(Var tpPtr: tpCell);
Procedure GetColCount;
Function GetPreviousFirstX(hFirstX:Integer):  Integer;
//procedure SetXAxis ;
//procedure SetYAxis ;
Procedure SetScreenDetail;
Procedure SetScreen;
Procedure ScrollUP;
Procedure ScrollDown;
Procedure ScrollRight;
Procedure ScrollLeft;
Procedure CursorUP;
Procedure CursorDown;
Procedure CursorRight;
Procedure CursorLeft;
Procedure PageDown;
Procedure PageUp;
Procedure PageRight;
Procedure PageLeft;
Procedure Screen_Initialize;

{ ---------------------------------------------- }

Implementation



(* ----------------------------------------------------------------
 Menu selection
  --------------------------------------------------------------- *)
Function SelectVertMenu(X1, Y1: byte; ListMenu : Menus; Count : byte):  byte;

Var 
   i, Xpos :  integer;
   Aschii, Code :  char;
Begin
   CursorOff ;
   Xpos := X1 ;
   For i := 1 To Count Do
      Begin
         WriteVramStr(Xpos, Y1, ListMenu[i].Menu, LightGray);
         Xpos := Xpos + CellLength(ListMenu[i].Menu);
      End;

   Xpos := X1 ;
   i := 1;
   WriteVramStr(Xpos, Y1, ListMenu[i].Menu, (Reverse*gcCellAttr And Not Blink));

   //GotoXy(Xpos+CellLength(ListMenu[i].Menu) div 2, Y1) ;
   GetKey( Aschii, Code );
   While True Do
      Begin
         If Aschii = #0 Then
            Aschii := Code;

         Case Aschii Of 
            LEFTKEY :
                       Begin
                          WriteVramStr(Xpos, Y1, ListMenu[i].Menu, LightGray);
                          i := i-1;
                          If i < 1 Then i := Count;
                          Xpos := CellLength(ListMenu[i].Menu)*(i-1)+X1 ;
                          WriteVramStr(Xpos, Y1, ListMenu[i].Menu, (Reverse*gcCellAttr And Not Blink
                          ));
                       End;

            RIGHTKEY :
                        Begin
                           WriteVramStr(Xpos, Y1, ListMenu[i].Menu, LightGray);
                           i := i+1;
                           If i > Count Then i := 1;
                           Xpos := CellLength(ListMenu[i].Menu)*(i-1)+X1;
                           WriteVramStr(Xpos, Y1, ListMenu[i].Menu, (Reverse*gcCellAttr And Not
                                        Blink));
                        End;

            CRKEY :
                     Begin
                        WriteVramStr(Xpos, Y1, ListMenu[i].Menu, LightGray);
                        SelectVertMenu := ListMenu[i].Num;
                        exit;
                     End;

            ESCKEY :
                      Begin
                         SelectVertMenu := 0;
                         exit;
                      End;
            Else
               For i:=1 To Count Do
                  Begin
                     //if Upcase(Chr(Aschii)) = Copy(ListMenu[i].Menu,1,1) then
                     If Upcase(Aschii) = Copy(ListMenu[i].Menu,1,1) Then
                        Begin
                           WriteVramStr(Xpos, Y1, ListMenu[i].Menu, LightGray);
                           SelectVertMenu := ListMenu[i].Num ;
                           exit ;
                        End ;
                  End ;
            Write(^G);
         End;
         GetKey( Aschii, Code );
      End;
End;

{ ---------------------------------------------- }
{  Write Range                                   }
{ ---------------------------------------------- }
Procedure WriteRange;
Begin
   If (ghTopCol <> ghEndCol) Or (ghTopRow <> ghEndRow) Then
      Begin
         GotoXY(MaxCol-18, h_GUIDELINE);
         Write('Range['+GetXString(ghTopCol), ghTopRow, '..', GetXString(ghEndCol), ghEndRow, ']') ;
      End ;
End;

{ ---------------------------------------------- }
{  Max Writable Length on Cell                   }
{ ---------------------------------------------- }
Function MaxCellLength(sData : String; x, y:Integer):  Integer;

Var 
   dataLen, cellLen, i :  integer ;

   Label BPOINT;

Begin
   cellLen := gcCellWidth[x] ;
   dataLen := CellLength(sData) ;


   i := x+1 ;

   While (i < ghFirstX+ghColCount) Do
      Begin
         //         If (dataLen <= cellLen) Or (Sheet[i,y].tpMain <> Nil) or (i=x+2) Then
         If (dataLen <= cellLen) Or (Sheet[i,y].tpMain <> Nil) Then
            goto BPOINT ;
         cellLen := cellLen + gcCellWidth[i] ;
         i := i + 1 ;
      End ;
   BPOINT:
            If cellLen > dataLen Then
               cellLen := dataLen ;
   If cellLen < gcCellWidth[x] Then
      cellLen := gcCellWidth[x] ;

   MaxCellLength := cellLen ;
End ;

{ ---------------------------------------------- }
{  Justify cell                                  }
{ ---------------------------------------------- }
Procedure Centerlize(Var sWork: String; hLen: Integer);

Var 
   i, hLen2 :  Integer;
Begin
   hLen2 := hLen Div 2 ;
   For i:=1 To hLen2 Do
      sWork := ' '+sWork ;
   For i:=1 To hLen2 Do
      sWork := sWork+' ' ;
End ;

Procedure RightJustify(Var sWork: String; hLen: Integer);

Var 
   i :  Integer;
Begin
   i := hLen ;
   While i>0 Do
      Begin
         sWork := ' '+sWork ;
         i := i-1 ;
      End ;
End ;

Procedure LeftJustify(Var sWork: String; hLen: Integer);

Var 
   i :  Integer;
Begin
   i := hLen ;
   While i>0 Do
      Begin
         sWork := sWork+' ' ;
         i := i-1 ;
      End ;
End ;

{ ---------------------------------------------- }
{  Set Data                                      }
{ ---------------------------------------------- }
Function SetData(Var tpPtr:tpCell):  String ;

Const 
   c_NUMERIC =  0 ;
   c_STRING =  1 ;

Var 
   i     :  Integer;
   maxCellLen  :  Byte ;
   sWork     :  String;
   sData     :  String;
   hErrPos    :  Integer ;
   hStatus    :  Integer ;
   hResult    :  Integer;
   fData     :  Real;
   sRet:  String ;

   Label BPOINT;

(**********************)
(* Start of SetData2  *)
(**********************)
Procedure SetData02(cFlag: Byte) ;

Var 
   i, hLen :  Integer;
Begin
  { ------------------------------------------------ }
  {  Centerize                                       }
  { ------------------------------------------------ }
   hLen := gcCellWidth[ghX] - CellLength(sWork) ;
   If (Copy(tpPtr.tpMain^,1,Length(s_CELL_CENTER)) = s_CELL_CENTER) And (Copy(tpPtr.tpMain^,Length(
      s_CELL_CENTER)+1,Length(s_CELL_CENTER)) <> s_CELL_CENTER) Then
      Begin
         Centerlize(sWork, hLen) ;
      End
  { ------------------------------------------------ }
  {  Right                                           }
  { ------------------------------------------------ }
   Else If (Copy(tpPtr.tpMain^,1,Length(s_CELL_RIGHT)) = s_CELL_RIGHT) And (Copy(tpPtr.tpMain^,
           Length(s_CELL_RIGHT)+1,Length(s_CELL_RIGHT)) <> s_CELL_RIGHT) Then
           Begin
              RightJustify(sWork, hLen) ;
           End

  { ------------------------------------------------ }
  {  Left                                            }
  { ------------------------------------------------ }
   Else If (Copy(tpPtr.tpMain^,1,Length(s_CELL_LEFT)) = s_CELL_LEFT) And (Copy(tpPtr.tpMain^,Length(
           s_CELL_LEFT)+1,Length(s_CELL_LEFT)) <> s_CELL_LEFT) Then
           Begin
              LeftJustify(sWork, hLen) ;
           End

  { ------------------------------------------------ }
  {  Pading                                          }
  { ------------------------------------------------ }
   Else If (Copy(tpPtr.tpMain^,1,Length(s_CELL_FILLCHAR)) = s_CELL_FILLCHAR) And (Copy(tpPtr.tpMain^
           ,Length(s_CELL_FILLCHAR)+1,Length(s_CELL_FILLCHAR)) <> s_CELL_FILLCHAR) Then
           Begin
              sWork := '' ;
              While gcCellWidth[ghX]>CellLength(sWork) Do
                 sWork := sWork+Copy(tpPtr.tpMain^,2,1) ;
           End

  { ------------------------------------------------ }
  {  Default                                         }
  { ------------------------------------------------ }
   Else
      Begin
         If cFlag = c_STRING Then
            //Centerlize(sWork, hLen)
            LeftJustify(sWork, hLen)
         Else
            RightJustify(sWork, hLen) ;
      End ;
End ;

Procedure SetNumericData;

Var 
   j :  Integer ;
Begin
   i := GetCellDecPoint(ghX, ghY);
   If i<>0 Then
      i := i+2
   Else
      i := i+1;
   If fData<0 Then
      i := i+1;
   If int(ln(abs(fData))/ln(10))+i>gcCellWidth[ghX] Then
      Begin
         sWork := '';
         For j:=1 To gcCellWidth[ghX] Do
            sWork := sWork+'*';
      End
   Else
      Begin
         str(fData:gcCellWidth[ghX]:GetCellDecPoint(ghX, ghY),sWork);
         Trim(sWork) ;
         SetData02(c_NUMERIC) ;
      End ;
End ;

(**********************)
(* Start of SetData   *)
(**********************)
Begin
   sWork := '' ;

   If (ghX > h_MAXCOL) Or (ghY > h_MAXROW) Then
      Begin
         SetData := sWork ;
         exit ;
      End ;

   If tpPtr.tpMain<>Nil Then
      Begin
         sData := tpPtr.tpMain^;
         ChangeFirstChar(sData) ;
      End ;
   If (tpPtr.tpMain<>Nil) And (sData<>'')  Then
      Begin
     { Numeric }
         If IsCellNumeric(tpPtr) Then
            Begin

               hErrPos := 0 ;
               If IsCellFormula(tpPtr) Then
                  Begin

                   { ------------------------------------------------ }
                   {  Evaluate sFormula                               }
                   { ------------------------------------------------ }
                     Evaluate(hStatus, sData, fData, hErrPos, sRet, ghX, ghY) ;

                     If hErrPos<>0 Then
                        Begin
                           //	 GotoXY(10,10) ;								 

//	 writeln('CellNumeric ',hStatus,' ', sData,' fData:',round(fData),' hErrPos:',hErrPos,' sRet:',sRet) ;

                           // @@@@@
                           //                           SetCellNotFormula(tpPtr) ;
                           SetCellNotNumeric(tpPtr) ;
                        End ;
                  End
               Else
                  val(sData, fData, hResult);

               If (fData<>0) And (hErrPos=0) Then
                  Begin
                     SetNumericData ;
                     goto BPOINT;
                  End;
               If fData=0 Then
                  Begin
                     str(fData:gcCellWidth[ghX]:GetCellDecPoint(ghX, ghY),sWork);
                     Trim(sWork) ;
                     SetData02(c_NUMERIC) ;
                     goto BPOINT;
                  End;
               goto BPOINT;
            End ;

     { String }
         If IsCellFormula(tpPtr) Then
            Begin

        { ------------------------------------------------ }
        {  Evaluate sFormula(String)                       }
        { ------------------------------------------------ }
               Evaluate(hStatus, sData, fData, hErrPos, sRet, ghX, ghY) ;
               If (hErrPos=0) Then
                  If (hStatus = h_FORMULA_STRING) Then
                     sData := sRet
               Else If (hStatus = h_FORMULA_Numeric) Then
                       Begin
                          SetCellNumeric(tpPtr) ;
                          SetNumericData ;
                          goto BPOINT;
                       End
               Else
                  SetCellNotFormula(tpPtr) ;
            End ;

         maxCellLen := MaxCellLength(sData, ghX, ghY);
         sWork := GetCellData(sData, 1, maxCellLen);
         SetData02(c_STRING) ;
         goto BPOINT;
      End ;

   sWork := '';

   BPOINT:
            SetData := sWork ;

End ;

{ ---------------------------------------------- }
{  Clear Menu                                    }
{ ---------------------------------------------- }
Procedure ClearMenuLine;
Begin
   WinClrScr( 1, h_MENULINE, MaxCol, h_MENULINE );
End;

{ ---------------------------------------------- }
{  Write Data                                    }
{ ---------------------------------------------- }
Procedure WriteData(Var tpPtr:tpCell; X,Y,Width:Integer; cAttr: Byte);

Var 
   sWork   :  String ;
   cellLen :  Integer ;
Begin
   If tpPtr.tpMain<>Nil Then
      sWork := SetData(tpPtr)
   Else
      sWork := '' ;


   cellLen := CellLength(sWork) ;
   If cellLen < Width Then
      cellLen := Width ;

   WriteVramStrWithLength(X,Y,cellLen,sWork,cAttr) ;

End ;


{ ---------------------------------------------- }
{  Set Current Data on Cell                      }
{ ---------------------------------------------- }
Procedure SetCurrentDataOnCell(Var tpPtr: tpCell; cAttr: Byte);
{
var
   i :  Integer ;
}
Begin
   WriteData(tpPtr,
             gcCellPos[ghScrX],
             ghScrY+ghColNameLine,
             gcCellWidth[ghX],
             cAttr) ;
   // @@@@Keep cAttr 
   // @@@@Insert logic to erase previous cell data .
   //SetScreenLine ;

End ;

{ ---------------------------------------------- }
{  Set Current Data on Rowl                      }
{ ---------------------------------------------- }

Procedure SetCurrentDataOnRow;

Var 
   i :  Integer ;
   hSaveX, hSaveY:  Integer ;
Begin
   hSaveX := ghX ;
   hSaveY := ghY ;
   For i:=1 To ghColCount Do
      Begin
         ghX := ghFirstX+i-1 ;
         If Sheet[ghX,ghY].tpMain <> Nil Then
            WriteData(Sheet[ghX,ghY],
                      gcCellPos[i],
                      ghColNameLine+ghY-ghFirstY+1,
                      gcCellWidth[ghX],
                      Sheet[ghX,ghY].cCellColor) ;
      End ;
   ghX := hSaveX ;
   ghY := hSaveY ;
End ;


{ ---------------------------------------------- }
{  Set Current Data OnLine                       }
{ ---------------------------------------------- }
Procedure SetCurrentDataOnLine(Var tpPtr: tpCell);

Var 
   sCellType :  String ;
   //sCellColor : String ;
Begin
   SetAttr(gcLineAttr) ;
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   sCellType := HEX(tpPtr.cCellType,2) ; { CVEVAL }
   //sCellColor := HEX(tpPtr.cCellColor,2) ; { CVEVAL }

   delete(sCellType,1,2) ;
   If sCellType = '' Then
      sCellType := '00' ;

   Write(GetXString(ghX), ghY, '[',gcCellWidth[ghX],' 0x',sCellType ,']=>');
   If (tpPtr.tpMain <>Nil) Then
      WriteString(GetCellData(tpPtr.tpMain^, 1, MaxCol-34), gcLineAttr) ;

   WriteRange;
End;

{ ---------------------------------------------- }
{  Calcurate Column Count                        }
{ ---------------------------------------------- }
Procedure GetColCount;

Var 
   i, hLen:  Integer ;
Begin
   ghColCount := 0 ;
   hLen := gcCellPos[1] ;

   i := 1 ;

   While (hLen < MaxCol) And (i<= h_MAXCOL-ghFirstX+1) Do
      Begin
         hLen := hLen + gcCellWidth[i+ghFirstX-1] ;
         If hLen >= MaxCol Then
            exit ;
         ghColCount := ghColCount + 1 ;
         If ghColCount > h_MAXCOL - ghFirstX+1 Then
            Begin
               ghColCount := h_MAXCOL - ghFirstX+1 ;
               exit ;
            End ;
         gcCellPos[i+1] := gcCellPos[i] + gcCellWidth[i+ghFirstX-1] ;
         i := i + 1 ;
      End;

End ;

{ ---------------------------------------------- }
{  Calcurate Previous Firstx                     }
{ ---------------------------------------------- }
Function GetPreviousFirstX(hFirstX:Integer):  Integer;

Var 
   i, hLen:  Integer ;
Begin
   i := hFirstX-1 ;

   hLen := gcCellPos[1] + gcCellWidth[i] ;
   While (i > 1) And (hLen < MaxCol) Do
      Begin
         hLen := hLen + gcCellWidth[i] ;
         If hLen < MaXCol Then
            i := i - 1 ;
      End;

   If i < 1 Then
      i := 1 ;
   GetPreviousFirstX := i ;

End ;
{ ---------------------------------------------- }
{  Set X-Axis                                    }
{ ---------------------------------------------- }
Procedure SetXAxis ;

Var 
   i, j     :  Integer ;
   w     :  integer ;
   currentAttr :  Byte ;
   xstr     :  String ;
Begin
   currentAttr := GetAttr ;
   GotoXY(1,ghColNameLine);
   WinClrScr( 1, ghColNameLine, MaxCol, ghColNameLine );
   TextColor(Black) ;
   TextBackground(Cyan) ;

   For i:=1 To ghColCount Do
      Begin
         w := gcCellWidth[ghFirstX+i-1] ;
         xstr := GetXString(ghFirstX+i-1) ;   {CVEVAL}

         For j:=1 To w Do
            Begin
               GotoXY(gcCellPos[i]+j-1,ghColNameLine) ;
               If j = ((gcCellWidth[ghFirstX+i-1]-length(xstr)) Div 2)+((gcCellWidth[ghFirstX+i-1]-
                  length(xstr)) Mod 2) Then
                  Begin
                     write(xstr) ;
                     j := j+length(xstr)-1 ;
                  End
               Else
                  write(' ') ;
            End ;
      End;
   SetAttr(currentAttr) ;
End ;

{ ---------------------------------------------- }
{  Set Y-Axis                                    }
{ ---------------------------------------------- }
Procedure SetYAxis ;

Var 
   i :  Integer ;
   currentAttr :  Byte ;
Begin
   currentAttr := GetAttr ;
   WinClrScr( 1, ghColNameLine+1, ghLeftSide, MaxRow );
   gotoxy(1,ghColNameLine) ;
   TextColor(AXIS_TEXTCOLOR) ;
   TextBackground(AXIS_TEXTBACKGROUND) ;

   For i:=1 To ghLeftSide-1 Do
      Begin
         gotoxy(i, ghColNameLine) ;
         write(' ') ;
      End ;

   For i:=1 To MaxRow-ghColNameLine  Do
      Begin
         gotoxy(1,i+ghColNameLine) ;
         write(ghFirstY+i-1:ghLeftSide-1) ;
      End ;
   SetAttr(currentAttr) ;
End ;

{ ---------------------------------------------- }
{  Set Screen                                    }
{ ---------------------------------------------- }
Procedure SetScreenDetail;

Var 
   i,j :  Integer;
   hSaveX, hSaveY:  Integer ;
Begin
   hSaveX := ghX ;
   hSaveY := ghY ;
   For j:=1 To MaxRow-ghColNameLine Do
      Begin
         For i:=1 To ghColCount Do
            Begin
               ghX := ghFirstX+i-1 ;
               ghY := ghFirstY+j-1 ;

               If (Sheet[ghX,ghY].tpMain <> Nil) Or (Sheet[ghX,ghY].cCellType <> 0) Then
                  WriteData(Sheet[ghX,ghY],
                            gcCellPos[i],
                            j+ghColNameLine,
                            gcCellWidth[ghX],
                            Sheet[ghX,ghY].cCellColor) ;
            End ;
      End ;
   ghX := hSaveX ;
   ghY := hSaveY ;

End;

{ ---------------------------------------------- }
{  Set Screen                                    }
{ ---------------------------------------------- }
Procedure SetScreen;
Begin
   //SetLowVideo ;

   If gcForm = c_NORMAL Then
      Begin
         ghColNameLine := h_COLNAMELINE;
         ghLeftSide := h_LEFTSIDE;
      End
   Else
      Begin
         ghColNameLine := h_COLNAMELINE-1;
         ghLeftSide := 1;
      End ;

   gcCellPos[1] := ghLeftSide;
   //WinClrScr( 1, 1, MaxCol, MaxRow );
   SetAttr(Black) ;
   ClrScr ;
   GetColCount ;

   If gcForm = c_NORMAL Then
      Begin
         SetXAxis ;
         SetYAxis ;
      End ;

   SetScreenDetail ;

End;


{ ---------------------------------------------- }
{  Write YAxis                                   }
{ ---------------------------------------------- }
{$IFDEF WIN}
Procedure WriteYAxis(number, x, y : Integer) ;

Var 
   currentAttr :  Byte ;
Begin
   currentAttr := GetAttr ;
   TextColor(AXIS_TEXTCOLOR) ;
   TextBackground(AXIS_TEXTBACKGROUND) ;
   gotoxy(x,y) ;
   write(number:ghLeftSide-1) ;
   SetAttr(currentAttr) ;
End ;
{$ENDIF WIN}

{ ---------------------------------------------- }
{  Scroll Up                                     }
{ ---------------------------------------------- }
Procedure ScrollUP;

{$IFDEF WIN}

Var 
   hSaveX :  Integer ;
   i   :  Integer ;
Begin
   GotoXY(1, ghColNameLine+1) ;
   DelLine ;
   ghFirstY := ghFirstY + 1 ;

   If gcForm = c_NORMAL Then
      WriteYAxis(ghY, 1, MaxRow) ;

   hSaveX := ghX ;
   For i:=1 To ghColCount Do
      Begin
         ghX := ghFirstX+i-1 ;
         WriteData(Sheet[ghX,ghY],
                   gcCellPos[i],
                   MaxRow,
                   gcCellWidth[ghX],
                   //cAttr)
                   Sheet[ghX,ghY].cCellColor) ;
      End;
   ghX := hSaveX;

{$ELSE NOWIN}
   Begin
      If ghFirstY < h_MaxRow Then
         Begin
            //SetLowVideo ;
            ghFirstY := ghFirstY + 1 ;
            SetScreen ;
         End ;
{$ENDIF WIN}
   End ;

{ ---------------------------------------------- }
{  Scroll Down                                   }
{ ---------------------------------------------- }
   Procedure ScrollDown;

{$IFDEF WIN}

   Var 
      i:  Integer ;
      hSaveX:  Integer ;

   Begin
      //WinScrollDown(1, ghColNameLine+1, MaxCol, MaxRow, 1) ;
      GotoXY(1, ghColNameLine+1) ;
      InsLine ;
      ghFirstY := ghFirstY - 1 ;

      If gcForm = c_NORMAL Then
         WriteYAxis(ghY, 1, ghColNameLine+1) ;

      hSaveX := ghX ;
      For i:=1 To ghColCount Do
         Begin
            ghX := ghFirstX+i-1 ;

            WriteData(Sheet[ghX,ghY],
                      gcCellPos[i],
                      ghColNameLine+1,
                      gcCellWidth[ghX],
                      //cAttr)
                      Sheet[ghX,ghY].cCellColor) ;
         End;
      ghX := hSaveX;

{$ELSE NOWIN}
      Begin
         If ghFirstY > 1 Then
            Begin
               //SetLowVideo ;
               ghFirstY := ghFirstY - 1 ;
               SetScreen ;
            End ;
{$ENDIF WIN}

      End ;

{ ---------------------------------------------- }
{  Scroll Right                                  }
{ ---------------------------------------------- }
      Procedure ScrollRight;
      Begin
         If ghFirstX > 1 Then
            Begin
               ghFirstX := ghFirstX - 1 ;
               SetScreen ;
            End ;
      End ;

{ ---------------------------------------------- }
{  Scroll Left                                   }
{ ---------------------------------------------- }
      Procedure ScrollLeft;
      Begin
         If ghFirstX < h_MaxCol Then
            Begin
               ghFirstX := ghFirstX + 1 ;
               SetScreen ;

               If ghX > ghFirstX+ghColCount-1 Then
                  Begin
                     ghX := ghFirstX+ghColCount-1 ;
                     ghScrX := ghColCount ;
                  End ;
            End ;
      End ;

{ ---------------------------------------------- }
{  Cursor Up                                     }
{ ---------------------------------------------- }
      Procedure CursorUP;
      Begin
         SetCurrentDataOnCell(Sheet[ghX,ghY], Sheet[ghX,ghY].cCellColor);
         SetCurrentDataOnRow;

         If ghY > 1 Then
            Begin
               ghY := ghY - 1 ;
               If ghScrY > 1 Then
                  ghScrY := ghScrY - 1
               Else
                  ScrollDown ;
            End ;
      End;

{ ---------------------------------------------- }
{  Cursor Down                                   }
{ ---------------------------------------------- }
      Procedure CursorDown;
      Begin
         SetCurrentDataOnCell(Sheet[ghX,ghY], Sheet[ghX,ghY].cCellColor);
         SetCurrentDataOnRow;

         If ghY < h_MAXROW Then
            Begin
               ghY := ghY + 1 ;
               If ghScrY < MaxRow-ghColNameLine  Then
                  ghScrY := ghScrY + 1
               Else
                  ScrollUp ;
            End ;
      End;

{ ---------------------------------------------- }
{  Cursor Right                                  }
{ ---------------------------------------------- }
      Procedure CursorRight;
      Begin
         SetCurrentDataOnCell(Sheet[ghX,ghY], Sheet[ghX,ghY].cCellColor);
         SetCurrentDataOnRow;

         If ghX < h_MAXCOL Then
            Begin
               ghX := ghX+1;
               If ghScrX < ghColCount Then
                  ghScrX := ghScrX + 1
               Else
                  SCrollLeft ;
            End ;
      End;

{ ---------------------------------------------- }
{  Cursor Left                                   }
{ ---------------------------------------------- }
      Procedure CursorLeft;
      Begin
         SetCurrentDataOnCell(Sheet[ghX,ghY], Sheet[ghX,ghY].cCellColor);
         SetCurrentDataOnRow;

         If ghX > 1 Then
            Begin
               ghX := ghX - 1;
               If ghScrX > 1 Then
                  ghScrX := ghScrX - 1
               Else
                  ScrollRight ;
            End ;
      End;

{ ---------------------------------------------- }
{  Page Down                                     }
{ ---------------------------------------------- }
      Procedure PageDown;
      Begin
         ghY := ghFirstY + MaxRow-ghColNameLine ;
         If ghY > h_MAXROW Then
            ghY := h_MAXROW ;

         ghX := ghFirstX ;
         ghFirstY := ghY ;
         ghScrY := 1 ;
         SetScreen;
         ghX := ghX + ghScrX-1 ;
      End;

{ ---------------------------------------------- }
{  Page Up                                       }
{ ---------------------------------------------- }
      Procedure PageUp;
      Begin
         ghY := ghFirstY - (MaxRow-ghColNameLine) ;
         If ghY < 1 Then
            ghY := 1 ;

         ghX := ghFirstX ;
         ghFirstY := ghY ;
         ghScrY := 1 ;
         SetScreen;
         ghX := ghX + ghScrX-1 ;
      End;

{ ---------------------------------------------- }
{  Page Right                                    }
{ ---------------------------------------------- }
      Procedure PageRight;
      Begin
         ghFirstX := ghFirstX + ghColCount ;
         If ghFirstX > h_MAXCOL Then
            ghFirstX := h_MAXCOL ;
         SetScreen;

         If ghScrx > ghColCount Then
            ghScrX := 1 ;
         ghX := ghFirstX + ghScrX-1
      End;

{ ---------------------------------------------- }
{  Page Left                                     }
{ ---------------------------------------------- }
      Procedure PageLeft;
      Begin

         If ghFirstX = 1 Then
            ghScrX := 1
         Else
            ghFirstX := GetPreviousFirstX(ghFirstX);

         SetScreen;
         ghX := ghFirstX + ghScrX-1
      End;

{ ---------------------------------------------- }
{  Screen Initalize                              }
{ ---------------------------------------------- }
      Procedure Screen_Initialize;
      Begin
         ghX := 1;
         ghY := 1;
         ghScrX := 1 ;
         Ghscry := 1 ;
         ghFirstX := 1 ;
         ghFirstY := 1 ;

         SetScreen ;
      End;
   End.
