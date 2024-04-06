{$MODE TP}

Unit cvfile ;

Interface

Uses cvdef, cvcrt, cvinpt, cvscrn, cveval ;
Procedure PutMsg(sStr : String);
Procedure GetFileName(Var sFileName: String; Var cAscii:Char);
Function CheckYNC(sStr : String):  byte;
Procedure ClearAllCells ;
Function LoadTableMain(Var sFileName:String; bClear:Boolean):  Boolean ;

Procedure FileMenu ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Put Msg                                       }
{ ---------------------------------------------- }
Procedure PutMsg(sStr : String);

Var 
   cAsii, cCode :  Char;
Begin
   WinClrScr( 1, h_GUIDELINE, MAxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   WriteString(sStr , gcLineAttr) ;
   GetKey( cAsii, cCode );
End ;

{ ---------------------------------------------- }
{  Get File Name                                 }
{ ---------------------------------------------- }
Procedure GetFileName(Var sFileName: String; Var cAscii:Char);

Var 
   cEdit :  Byte ;

Begin
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);

   WriteString('Specify File Name [ESC]:CANCEL', gcLineAttr);

   cEdit := c_Edit ;
   //GetInputData(sFileName, h_MAXCELLSIZE, cAscii, cEdit);
   GetInputData(sFileName, cAscii, cEdit);
End ;

{ ---------------------------------------------- }
{  Check Yes No Cancel                           }
{ ---------------------------------------------- }
Function CheckYNC(sStr : String):  byte;

Var 
   cAscii, cCode :  Char;
   Ch:  Char ;
Begin
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   WriteString(sStr+' Y)es/N)o  ' , gcLineAttr) ;

   Repeat
      GetKey( cAscii, cCode );
      Ch := UpCase(cAscii);
      Case Ch Of 
         'Y':  CheckYNC := YES_No;
         'N':  CheckYNC := NO_No;
         Else
            Write(^G);
      End;
   Until Ch In['Y','N'];

End;

{ ---------------------------------------------- }
{  Clear All Cell                                }
{ ---------------------------------------------- }
Procedure ClearAllCells ;

Var 
   i, j:  Integer ;
Begin
   For j:=1 To h_MAXROW Do
      For i:=1 To h_MAXCOL Do
         Begin
            If j = 1 Then
               gcCellWidth[i] := h_INIT_CELLSIZE;
            If Sheet[i,j].tpMain <> Nil Then
               FreeCellArea(Sheet[i,j]) ;
            Sheet[i,j].cCellType := 0 ;
            Sheet[i,j].cCellColor := gcCellAttr ;
         End ;
End ;

{ ---------------------------------------------- }
{  Open File                                     }
{ ---------------------------------------------- }
Procedure OpenFile(Var tpFile:Text; Var sFileName:String) ;

Var 
   hPos    :  integer;
Begin
   hPos := Pos('.',sFileName);
   If hPos = 0 Then
      sFileName := sFileName+'.cv2';
   Assign(tpFile, sFileName);
 {$I-}
   Reset(tpFile);
 {$I+}

End ;

{ ---------------------------------------------- }
{  Save File  Main                               }
{ ---------------------------------------------- }
Procedure SaveTableDataMain(Var sFileName:String ; hTopCol,hTopRow: integer; hEndCol,hEndRow:
                            integer);

Var 
   //  i, j, hPos : integer;
   i, j:  integer;
   tpFile :  Text;
   hNumber:  Integer ;
   tpPtr:  tpCell ;
   sData :  String;

Begin
   OpenFile(tpFile, sFileName) ;
   If IOresult=0 Then   { File:already exists }
      Begin
         hNumber := CheckYNC('Overwrite file ? ');
         If hNumber <> 1 Then
            exit ;
      End;

  {$I-}
   rewrite(tpFile);
  {$I+};
   If IOresult<>0 Then   { I/O Error }
      Begin
         PutMsg('Fail to save file.');
         exit ;
      End ;

   sData := '$$CVDATA##2';
   writeln(tpFile,sData);


   writeln(tpFile,hTopRow);
   writeln(tpFile,hEndRow);
   writeln(tpFile,hTopCol);
   writeln(tpFile,hEndCol);



{
   for i:=hTopCol to hEndCol do
      writeln(tpFile,gcCellDecPoint[i]);
	  //writeln(tpFile, GetCellDecPoint(i, ghY));
}

   For i:=hTopCol To hEndCol Do
      writeln(tpFile,gcCellWidth[i]);

   For j:=hTopRow To hEndRow Do
      Begin
         For i:=hTopCol To hEndCol Do
            Begin
               sData := 'D';
               tpPtr := Sheet[i,j] ;
               If tpPtr.tpMain<>Nil Then
                  sData := sData+tpPtr.tpMain^;
               writeln(tpFile,sData);
               writeln(tpFile,tpPtr.cCellType) ;
               writeln(tpFile,tpPtr.cCellColor) ;
            End;
      End;
   writeln(tpFile,gcForm);
   Close(tpFile);

End ;

{ ---------------------------------------------- }
{  Save File                                     }
{ ---------------------------------------------- }
Procedure SaveTableData(hTopCol,hTopRow: integer; hEndCol,hEndRow: integer);

Var 
   cAscii :  Char;

Begin

  { ******************************** }
  {   Get File Name                  }
  { ******************************** }
   GetFileName(gsFileName, cAscii) ;
   If (gsFileName = '') Or (cAscii = ESCKEY) Then
      exit ;

   SaveTableDataMain(gsFileName, hTopCol,hTopRow,hEndCol,hEndRow);
End ;

Procedure SaveTable;
Begin
   SaveTableData(1, 1, ghMaxX, ghMaxY);
End;

{ ---------------------------------------------- }
{  Load File Main                                }
{ ---------------------------------------------- }
Function LoadTableMain(Var sFileName:String; bClear:Boolean):  Boolean ;

Var 
   i, j      :  integer;
   hTopRow,hTopCol :  integer;
   hEndRow,hEndCol :  integer;
   sData     :  String ;
   cData   :  byte;
   tpFile     :  Text;
   flag      :  Integer ;
   cAscii     :  Char ;
   version     :  Char ;
Begin
   OpenFile(tpFile, sFileName) ;
   While (IOresult <> 0) Do     { File not found }
      Begin
         PutMsg('File not found.');
         If bClear = False Then
            Begin
               LoadTableMain := False ;
               exit ;
            End ;

         GetFileName(sFileName, cAscii) ;
         If (sFileName = '') Or (cAscii = ESCKEY) Then
            Begin
               LoadTableMain := False ;
               exit ;
            End ;
         OpenFile(tpFile, sFileName) ;
      End ;


  { ******************************** }
  {   Clear Cells                    }
  { ******************************** }
   If bClear Then
      ClearAllCells ;

   version := '1' ;
   readln(tpFile,sData);

   { Check CV file }
   If (length(sData) < 10) Or (copy(sData, 1, 10) <> '$$CVDATA##') Then
      Begin
         PutMsg('Illegal format .');
         LoadTableMain := False ;
         exit ;
      End ;
   If length(sData) = 11 Then
      version := sData[11] ;

   readln(tpFile,hTopRow);
   readln(tpFile,hEndRow);
   readln(tpFile,hTopCol);
   readln(tpFile,hEndCol);
   ghMaxX := hEndCol ;
   ghMaxY := hEndRow ;

   If version <> '2' Then
      For i:=hTopCol To hEndCol Do
         readln(tpFile,gcCellDecPoint[i]);

   For i:=hTopCol To hEndCol Do
      readln(tpFile,gcCellWidth[i]);

   For j:=hTopRow To hEndRow Do
      Begin
         For i:=hTopCol To hEndCol Do
            Begin
               readln(tpFile,sData);
               readln(tpFile,cData);
               If (length(sData)>1) Or (cData <> $00)    Then
                  Begin
                     If Sheet[i,j].tpMain = Nil Then
                        GetCellArea(Sheet[i,j]);
                     If (Length(sData) > 1) And (Length(sData) <= h_MAXCELLSIZE)  Then
                        Sheet[i,j].tpMain^ := copy(sData, 2, Length(sData));
                     // @@@@@
                     If version <> '2' Then
                        cData := (cData And $f0) + gcCellDecPoint[i] ;
                     //@@@@

                     Sheet[i,j].cCellType := cData;
                  End ;
               If version = '2' Then
                  Begin
                     readln(tpFile,cData);
                     Sheet[i,j].cCellColor := cData ;
                  End ;
            End;
      End ;

   readln(tpFile,gcForm);
   Close(tpFile);
   LoadTableMain := True ;

End ;

{ ---------------------------------------------- }
{  Load File                                     }
{ ---------------------------------------------- }
Procedure LoadTable(bClear: Boolean);

Var 
   cAscii  :  Char ;
Begin

  { ******************************** }
  {   Get File Name                  }
  { ******************************** }
   gsFileName := '' ;
   GetFileName(gsFileName, cAscii) ;
   If (gsFileName = '') Or (cAscii = ESCKEY) Then
      exit ;

   If LoadTableMain(gsFileName, bClear) Then
      Screen_Initialize;
End;

{ ---------------------------------------------- }
{  File Menu                                     }
{ ---------------------------------------------- }
Procedure FileMenu ;

Var 
   hMenu:  Integer ;
Begin
   ClearMenuLine ;
{$IFDEF JPN}

   MenuList[1].Menu := 'L ロード ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'S セーブ ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C クリア ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'Q 終了   ';
   MenuList[4].Num  := 0;
{$ELSE NOJPN}

   MenuList[1].Menu := 'L Load   ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'S Save   ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C Clear  ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'Q Quit   ';
   MenuList[4].Num  := 0;
{$ENDIF}

   hMenu :=  SelectVertMenu(1, h_MENULINE, MenuList, 4);
   CursorOn ;
   If hMenu = 1 Then
      LoadTable(True) ;

   If hMenu = 2 Then
      SaveTable ;

   If hMenu = 3 Then
      Begin
         ClearAllCells ;
         Screen_Initialize;
      End;
   CursorOff ;
End ;
End .
