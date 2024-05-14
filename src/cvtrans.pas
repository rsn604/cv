{$MODE TP}

Unit cvtrans ;

Interface

Uses cvdef, cvcrt, cvcell, cvscrn, cvdata, cveval, cvfile, cvundo ;
Procedure TransMenu ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Trans TextData  cFlag: c_TEXTFILE,c_CSVFILE   }
{ ---------------------------------------------- }
Function GetCellData(x, y: integer):  String ;

Var 
   sData:  String ;
   hErrPos:  Integer ;
   sRet:  String ;
   fData:  Real ;
   hStatus:  Integer ;

Begin
   sData := Sheet[x,y].tpMain^ ;
   If IsCellFormula(Sheet[x,y]) Then
      Begin
         Evaluate(hStatus, sData, fData, hErrPos, sRet, x, y) ;
         If hErrPos = 0 Then
            Begin
               If hStatus = h_FORMULA_NUMERIC Then
                  str(fData:gcCellWidth[x]:GetCellDecPoint(x, y),sData)
               Else
                  sData := sRet ;
               Trim(sData) ;
            End;
      End ;
   GetCellData := sData ;
End ;

{ ---------------------------------------------- }
{  Trans TextData  cFlag: c_TEXTFILE,c_CSVFILE   }
{ ---------------------------------------------- }
Procedure TransTextData(hTopCol,hTopRow: integer; hEndCol,hEndRow: integer;cFlag:Byte);

Var 
   tpFile        :  Text;
   sFileName      :  String ;
   sData        :  String ;
   i, j, hNumber, hPos :  Integer ;
   cAscii        :  Char ;
   saveCellWidth    :  array [1..h_MAXCOL] Of Byte ;
   hDataLen       :  Integer ;
Begin
  { ******************************** }
  {   Get File Name                  }
  { ******************************** }
   sFileName := '' ;
   GetFileName(sFileName, cAscii) ;               { CVFILE }
   If (sFileName = '') Or (cAscii = ESCKEY) Then
      exit ;

   hPos := Pos('.',sFileName);
   If hPos = 0 Then
      If cFlag = c_TEXTFILE Then
         sFileName := sFileName+'.txt'
   Else
      sFileName := sFileName+'.csv' ;

   Assign(tpFile, sFileName);

  {$I-}
   Reset(tpFile);
  {$I+}
   If IOresult=0 Then   { File:already exists }
      Begin
         hNumber := CheckYNC('Overwrite ? ');
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

   If cFlag = c_TEXTFILE Then
      Begin
         For i:=hTopCol To hEndCol Do
            saveCellWidth[i] := gcCellWidth[i] ;
         AdjustAllWidth ;
      End ;

   For j:=hTopRow To hEndRow Do
      Begin
         For i:=hTopCol To hEndCol Do
            Begin
               If Sheet[i,j].tpMain <> Nil Then
                  //                  sData := Sheet[i,j].tpMain^
                  sData := GetCellData(i,j)
               Else
                  sData := '' ;

               Trim(sData) ;      { CVEVAL }
               If cFlag = c_CSVFILE Then
                  Begin
                     //Trim(sData) ;      { CVEVAL }
                     If pos(',', sData) > 0 Then
                        sData := '"'+sData+'"' ;

                     If i<>hEndCol Then
                        Begin
                           write(tpFile,sData) ;
                           write(tpFile,',') ;
                        End
                     Else
                        writeln(tpFile,sData) ;
                  End
               Else
                  Begin
                     If length(sData) > gcCellwidth[i] Then
                        hDataLen := gcCellwidth[i]
                     Else
                        hDataLen := gcCellWidth[i]-length(sData) ;
                     If IsCellNumeric(Sheet[i,j]) Then
                        RightJustify(sData, hDataLen)
                     Else
                        LeftJustify(sData, hDataLen);

                     If i<>hEndCol Then
                        write(tpFile,sData)
                     Else
                        writeln(tpFile,sData) ;
                  End ;
            End;
      End;

   If cFlag = c_TEXTFILE Then
      Begin
         For i:=hTopCol To hEndCol Do
            gcCellWidth[i] := saveCellWidth[i] ;

         SetScreen ;
      End ;

   Close(tpFile);
End;

Procedure TransText;
Begin
   TransTextData(1, 1, ghMaxX, ghMaxY, c_TEXTFILE);
End;

Procedure TransCSV;
Begin
   TransTextData(1, 1, ghMaxX, ghMaxY, c_CSVFILE);
End;

{ ---------------------------------------------- }
{  Get CSV Main                                  }
{ ---------------------------------------------- }
{$B-}
Procedure GetCSVMain(sFileName:String; bClear:Boolean);

Const 
   c_INIT =  0 ;
   c_BREAK =  9 ;

Var 
   i, j, hPos :  integer;
   hSize   :  Integer ;
   sData   :  String ;
   cSw    :  byte;
   cSep    :  Char ;
   tpFile   :  File Of char;
   Ch     :  Char ;

Begin
   hPos := Pos('.',sFileName);
   If hPos = 0 Then
      sFileName := sFileName+'.csv';
   Assign(tpFile, sFileName);

 {$I-}
   Reset(tpFile);
 {$I+}
   If IOresult <> 0 Then     { File not found }
      Begin
         PutMsg('File not found.');
         exit ;
      End ;

  { ******************************** }
  {   Clear Cells                    }
  { ******************************** }
   If bClear Then
      ClearAllCells ;

   // Max Cell No.	  
   ghMaxX := 1 ;
   ghMaxY := 1 ;

   j := 0 ;
   While (Not Eof(tpFile)) And (j<h_MAXROW) Do
      Begin
         Read (tpFile,Ch);
         j := j+1 ;
         i := 0 ;

         While (i<h_MAXCOL) And (Ch<>#10) And (Not Eof(tpFile)) Do
            Begin
               i := i+1 ;
               sData := '' ;
               hSize := 0 ;
               cSw := c_INIT ;
               If (Ch='"') Then
                  Begin
                     cSep := '"' ;
                     Read (tpFile,Ch);
                  End
               Else
                  cSep := ',' ;

               While (cSw<>c_BREAK)  And (Not Eof(tpFile)) Do
                  Begin
                     If Ch = cSep Then
                        Begin
                           If (Ch='"') Then
                              Read (tpFile,Ch);

                           cSw := c_BREAK ;
                           // if loop will broke by MAXCOL, Unread next char.
                           If (i<h_MAXCOL) Then
                              Read (tpFile,Ch);
                        End;

                     If Ch=#13 Then
                        Read (tpFile,Ch);

                     If Ch=#10 Then
                        cSw := c_BREAK ;

                     If cSw<>c_BREAK Then
                        Begin
                           If hSize<h_MAXCELLSIZE Then
                              Begin
                                 sData := sData+Ch;
                                 hSize := hSize+1 ;
                              End ;
                           Read (tpFile,Ch);
                        End ;
                     //GotoXy(20,22) ;
                     // Write('Ch;', Ch, ':sData:', sData, ':length:', length(sData)) ;
                  End ;

               ghX := i ;
               ghY := j ;
               CheckData(sData,Sheet[i,j] );  {CVDATA}

               // update max cell no.
               If ghX > ghMaxX Then
                  ghMaxX := ghX ;
               If ghY > ghMaxY Then
                  ghMaxY := ghY ;
            End ;
      End;
   Close(tpFile);
End ;

{ ---------------------------------------------- }
{  Get Trans File                                }
{ ---------------------------------------------- }
Procedure GetTransFile(bClear: Boolean; cFlag:Byte);

Var 
   sFileName :  String ;
   cAscii  :  Char ;
Begin

  { ******************************** }
  {   Get File Name                  }
  { ******************************** }
   sFileName := '' ;
   GetFileName(sFileName, cAscii) ;
   If (sFileName = '') Or (cAscii = ESCKEY) Then
      exit ;
   GetCSVMain(sFileName, bClear);
   ClearStack ;         { CVUNDO }
   Screen_Initialize;
End;

{ ---------------------------------------------- }
{  Trans Menu                                    }
{ ---------------------------------------------- }
Procedure TransMenu ;

Var 
   hMenu:  Integer ;

Begin
   ClearMenuLine ;
{$IFDEF JPN}
   MenuList[1].Menu := 'T Text書出 ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'C CSV 書出 ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'G CSV 読込 ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'Q 終了     ';
   MenuList[4].Num  := 0;

{$ELSE NOJPN}
   MenuList[1].Menu := 'T Text Write ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'C CSV Write  ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'G CSV Read   ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'Q Quit       ';
   MenuList[4].Num  := 0;
{$ENDIF JPN}
   hMenu :=  SelectVertMenu(1, h_MENULINE, MenuList, 4);  {CVSCRN}
   CursorOn ;                  { CVCRT  }
   If hMenu = 1 Then
      TransText ;

   If hMenu = 2 Then
      TransCSV ;

   If hMenu = 3 Then
      GetTransFile(True, c_CSVFILE) ;

   CursorOff ;                 { CVCRT  }
End ;
End .
