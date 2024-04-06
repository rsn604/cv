{$MODE TP}

Unit cvcrt ;

Interface

Uses cvdef ;
Function IsMultiByte( C:Char):  Boolean;
{$IFDEF WIN}
Function IsMultiByte2( C:Char):  Boolean;
{$ENDIF WIN}
Procedure CheckRowCol( Var Row, Col : integer);
Procedure ClrScr;
Procedure ClrEol;
Procedure GotoXY( Xpos, Ypos : integer);
Procedure CursorOn ;
Procedure CursorOff ;
Function GetAttr:  byte;
Procedure SetAttr(Attr : byte);
Procedure SetHighVideo;
Procedure SetLowVideo;

Procedure TextColor(Color: Byte);
Procedure TextBackground(Color: Byte);
Procedure GetCur( Var Xpos, Ypos : integer; Var Cur : byte);
//procedure SetCur( Xpos, Ypos : byte; Cur : byte);
Procedure InsLine;
Procedure DelLine;
Function CharLength(c : Char; Var w2: Integer):  Integer;
Function CellLength(Str : String):  Integer;
Procedure WriteChar( Code : char; Attr :byte);
Procedure WriteString( S : String; Attr :byte);
Procedure GetKey( Var Asciicode, Scancode : Char);
//procedure SetVideoMode(cMode:Byte) ;
Function GVideoMode:  byte;
//procedure WinScrollUp( X1, Y1, X2, Y2, cLine : byte );
//procedure WinScrollDown( X1, Y1, X2, Y2, cLine : byte );
Procedure WinClrScr( X1, Y1, X2, Y2 : integer );
Function GetCellData(sStr : String; hStartDataPos, hLen:Integer):  String ;
Function GetCellDataWithEndPos(sStr : String; hStartDataPos, hLen:Integer; Var hEndDataPos,
                               hEndPos:Integer):  String ;
Procedure WriteVramStrWithLength(Xp, Yp, hLen : integer; sStr : String; cAttr: byte);
Procedure WriteVramStr( Xp, Yp : integer ;  Str : String; Attr : byte);

{ ---------------------------------------------- }

Implementation

Uses crt ;



(* ----------------------------------------------------------------
 Check MultiByte char
   --------------------------------------------------------------- *)
Function IsMultiByte( C:Char):  Boolean;
Begin
{$IFDEF WIN}
   If (((Byte(c)>=$81) And (Byte(c)<=$9f)) Or ((Byte(c)>=$e0) And (Byte(c)<= $fc))) Then
      IsMultiByte := True
   Else
      IsMultiByte := False ;
{$ELSE NOWIN}
   If c >= #128 Then
      IsMultiByte := True
   Else
      IsMultiByte := False ;
{$ENDIF WIN}
End ;

{$IFDEF WIN}
Function IsMultiByte2( C:Char):  Boolean;
Begin
   If (((Byte(c)>=$40) And (Byte(c)<=$7e)) Or ((Byte(c)>=$80) And (Byte(c)<= $fc))) Then
      IsMultiByte2 := True
   Else
      IsMultiByte2 := False ;
End ;
{$ENDIF WIN}



(* ----------------------------------------------------------------
 Current  Row, Col
   --------------------------------------------------------------- *)
Procedure CheckRowCol( Var Row, Col : integer);
Begin
   Col := crt.WindMaxX ;
   Row := crt.WindMaxY;
{$IFDEF WIN}
   If Row > 100 Then
      Row := 25 ;
{$ENDIF WIN}

End;



(* ----------------------------------------------------------------
   Clear screen
   --------------------------------------------------------------- *)
Procedure ClrScr;
Begin
   crt.ClrScr ;
End;



(* ----------------------------------------------------------------
    Clear aafter cursor position
   --------------------------------------------------------------- *)
Procedure ClrEol;
Begin
   crt.ClrEol ;
End;



(* ----------------------------------------------------------------
   Set cursor
   --------------------------------------------------------------- *)
Procedure GotoXY( Xpos, Ypos : integer);
Begin
   crt.GotoXY(Xpos, Ypos) ;
End;

{ ---------------------------------------------- }
{  Cursor On                                     }
{ ---------------------------------------------- }
Procedure CursorOn ;
Begin
   crt.cursoron ;
End ;

{ ---------------------------------------------- }
{  Cursor Off                                     }
{ ---------------------------------------------- }
Procedure CursorOff ;
Begin
   crt.cursoroff ;
End ;

{ ---------------------------------------------- }
{  Set Text Attrubute                            }
{ ---------------------------------------------- }
Procedure TextColor(Color: Byte);
Begin
   crt.TextColor(Color) ;
End ;

Procedure TextBackground(Color: Byte);
Begin
   crt.TextBackGround(Color) ;
End ;

{ ---------------------------------------------- }
{  Get Current Attrubute                         }
{ ---------------------------------------------- }
Function GetAttr:  byte;
Begin
   GetAttr := crt.TextAttr ;
End;

{ ---------------------------------------------- }
{  Set Attrubute                                 }
{ ---------------------------------------------- }
Procedure SetAttr(Attr : byte);
Begin
   crt.TextAttr := Attr ;

End;

Procedure SetHighVideo;
Begin
   crt.TextAttr := crt.TextAttr Or $08 ;
End;

Procedure SetLowVideo;
Begin
   crt.TextAttr := crt.TextAttr And $77 ;
End;



(* ----------------------------------------------------------------
   Get cursor pos and attribute
   --------------------------------------------------------------- *)
//procedure GetCur( var Xpos, Ypos : byte; var Cur : integer);
Procedure GetCur( Var Xpos, Ypos : integer; Var Cur : byte);
Begin

   Xpos := crt.WhereX ;
   Ypos := crt.WhereY ;
   Cur := GetAttr ;
End;



(* ----------------------------------------------------------------
      Set cursor pos and attribute
   --------------------------------------------------------------- *)


{
//procedure SetCur( Xpos, Ypos : byte; Cur : integer);
procedure SetCur( Xpos, Ypos : byte; Cur : byte);
begin
   GotoXY(Xpos, Ypos) ;
   SetAttr(Cur) ;
end;
}



(* ----------------------------------------------------------------
    Insert line at cursor position
   --------------------------------------------------------------- *)

Procedure InsLine;
Begin
   crt.InsLine ;
End;



(* ----------------------------------------------------------------
    Delete line at cursor position
   --------------------------------------------------------------- *)

Procedure DelLine;
Begin
   crt.DelLine ;
End;



(* ----------------------------------------------------------------
    Get width of char .
   --------------------------------------------------------------- *)
Function CharLength(c : Char; Var w2: Integer):  Integer;

Var w :  byte ;
Begin
{$IFDEF WIN}
   w := 1 ;

   If IsMultiByte(c) Then
      Begin
         w := 2 ;
         w2 := 2 ;
      End
   Else
      Begin
         w := 1 ;
         w2 := 1 ;
      End ;
   CharLength := w ;

{$ELSE NOWIN}
   w := 1 ;
   If (Byte(c) And $f0) = c_2BYTE Then
      Begin
         w := 2 ;
         w2 := 2 ;
      End
   Else If (Byte(c) And $f0) = c_3BYTE Then
           Begin
              w := 3 ;
              w2 := 2 ;
           End
   Else If (Byte(c) And $f0) = c_4BYTE Then
           Begin
              w := 4 ;
              w2 := 2 ;
           End
   Else
      Begin
         w := 1 ;
         w2 := 1 ;
      End ;
   CharLength := w ;
{$ENDIF WIN}
End;



(* ----------------------------------------------------------------
 Get max length within CELL .
   --------------------------------------------------------------- *)
Function CellLength(Str : String):  Integer;

Var w, w2, cellLen :  Integer ;
   i  :  integer ;
Begin
   cellLen := 0 ;
   i := 1;
   //i := 0;

   If Length(Str) > 0 Then
      Repeat
         w := CharLength(Str[i], w2) ;
         i := i + w ;
         cellLen := cellLen + w2 ;
      Until i > Length(Str);
   CellLength := cellLen ;
End;



(* ----------------------------------------------------------------
 Write char on cursor current position .
   --------------------------------------------------------------- *)
Procedure WriteChar( Code : char; Attr :byte);
Begin
   crt.TextAttr := Attr ;
   write(code) ;
End;



(* ----------------------------------------------------------------
 Write string and attribute .
   --------------------------------------------------------------- *)
{$IFDEF WIN}
Procedure WriteString( S : String; Attr :byte);

Var 
   X, Y :  integer;
   C     :  byte;
Begin
   GetCur(X, Y, C);
   GotoXY(X, Y);
   crt.TextAttr := Attr ;

   write(S) ;
End ;

{$ELSE NOWIN}
Procedure WriteString( S : String; Attr :byte);

Var 
   i, j, k, X, Y :  integer;
   C    :  byte;
   w, w2   :  integer ;
Begin
   GetCur( X, Y, C );
   j := X;
   i := 1 ;
   While (i <= Length(S)) Do
      Begin
         GotoXY(j, Y );
         w := CharLength(S[i], w2) ;
         For k:=1 To w Do
            //GotoXY(j+k-1, Y );
            WriteChar(S[i+k-1], Attr);
         i := i + w ;
         j := j + w2;

         // @@@@@@@@@	  
         // これを入れないと、漢字入力後のカーソル移動がおかしい。
         GotoXy(1,1) ;
      End;
End;
{$ENDIF WIN}


(* ----------------------------------------------------------------
 Get key code .
   --------------------------------------------------------------- *)
Procedure GetKey( Var Asciicode, Scancode : Char);
Begin
   Asciicode := crt.ReadKey ;
   If Asciicode = #0 Then
      Scancode := crt.ReadKey;
End;



(* ----------------------------------------------------------------
 Set display size.
 (video mode is ignored.) 
   --------------------------------------------------------------- *)
{
procedure SetVideoMode(cMode:Byte) ;
begin
end ;
}
Function GVideoMode:  byte;
Begin
   CheckRowCol(MaxRow, MaxCol);
   GVideoMode := 0 ;
End;



(* ----------------------------------------------------------------
   Scroll Up
   --------------------------------------------------------------- *)


{
procedure WinScrollUp( X1, Y1, X2, Y2, cLine : byte );
begin
   GotoXY(X1,Y1) ;
   crt.DelLine

end;
}


(* ----------------------------------------------------------------
   Scroll Down
   --------------------------------------------------------------- *)


{
procedure WinScrollDown( X1, Y1, X2, Y2, cLine : byte );
begin
   GotoXY(X1,Y1) ;
   crt.InsLine
end;
}


(* ----------------------------------------------------------------
   Clera range
   --------------------------------------------------------------- *)
Procedure WinClrScr( X1, Y1, X2, Y2 : integer );

Var x, y :  integer ;
Begin

   For y:=Y1 To Y2 Do
      Begin
         //				 if (X1 = 1) and (X2 = MaxCol) then
         If X2 = MaxCol Then
            Begin
               GotoXY(X1, y) ;
               ClrEol ;
            End
         Else
            Begin
               For x:=X1 To X2 Do
                  Begin
                     GotoXY(x, y) ;
                     WriteChar(' ', GetAttr) ;
                  End;
            End ;
      End;
End;


(* ----------------------------------------------------------------
  Get available data within cell
   --------------------------------------------------------------- *)
Function GetCellData(sStr : String; hStartDataPos, hLen:Integer):  String ;

Var 
   dummy1, dummy2 :  Integer ;
Begin
   dummy1 := 0 ;
   dummy2 := 0 ;
   GetCellData := GetCellDataWithEndPos(sStr, hStartDataPos, hLen, dummy1, dummy2) ;
End ;



(* ----------------------------------------------------------------
  Get available data within cell and end position
   --------------------------------------------------------------- *)
Function GetCellDataWithEndPos(sStr : String; hStartDataPos, hLen:Integer; Var hEndDataPos,
                               hEndPos:Integer):  String ;

Var 
   Buffer :  String;
   i, j :  Integer ;
   w, w2, hPos:  integer ;

   Label BPOINT;

Begin
   //gotoxy(20,15) ;
   //writeln('GetCellData:', sStr, ' hLen:', hLen) ;
   Buffer := '' ;
   i := hStartDataPos ;
   hPos := 1 ;
   Repeat
      If i > Length(sStr) Then
         Begin
            //hEndPos := CellLength(sStr)+1 ;
            hEndPos := CellLength(copy(sStr, hStartDataPos, length(sStr)-hStartDataPos+1))+1 ;
            hEndDataPos := Length(sStr)+1 ;
            goto BPOINT ;
         End ;

      w := CharLength(sStr[i], w2) ;

      If (w >= 2) And (hPos > hLen - 1) Then
         goto BPOINT ;

      For j := 1 To w Do
         Buffer := Buffer+sStr[i+j-1] ;

      hEndPos := hPos ;
      hPos := hPos + w2 ;

      hEndDataPos := i ;
      i := i + w ;
      //Until hPos >= hLen;
   Until hPos > hLen;

   BPOINT:
            //hEndDataPos := hStartDataPos + length(Buffer)  ;
            //hEndPos := cellLen ;
            GetCellDataWithEndPos := Buffer ;
   //gotoxy(20,16) ;
   //writeln('GetCellData2:', Buffer, ':Len:', Length(Buffer)) ;

End ;



(* ----------------------------------------------------------------
   Write string at position
   --------------------------------------------------------------- *)
Procedure WriteVramStrWithLength( Xp, Yp, hLen : integer; sStr : String; cAttr : byte);

Var 
   j, cellLen    :  Integer ;
   sWork     :  String;
Begin

   cellLen := CellLength(sStr) ;

   CursorOff ;
   GotoXY(Xp, Yp) ;
   sWork := sStr ;

   For j := 1 To hLen - cellLen Do
      sWork := sWork + ' ' ;
   WriteString(sWork, cAttr);
End ;



(* ----------------------------------------------------------------
   Write string at position
   --------------------------------------------------------------- *)
Procedure WriteVramStr( Xp, Yp : integer ;  Str : String; Attr : byte);
Begin

   WriteVramStrWithLength( Xp, Yp, CellLength(Str), Str, Attr);
End;

End.
