{$MODE TP}

Unit cvinpt ;

Interface

Uses cvdef, cvcrt, cvscrn, cveval ;
Procedure GetInputData(Var sData : String; Var cExKey:Char; Var cEdit:Byte);

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Set CalcMode Message                          }
{ ---------------------------------------------- }
Procedure SetCalcModeMessage(Var cEdit:Byte);
Begin
   cEdit := cEdit Or c_CALC ;
   cEdit := cEdit Or c_EDIT ;
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   WriteString('Specify formula ', gcLineAttr)
End ;

{ ---------------------------------------------- }
{  Input Data on h_INPUTLINE                     }
{ ---------------------------------------------- }
Procedure GetInputData(Var sData : String; Var cExKey:Char; Var cEdit:Byte);

Var 
   sInpString           :  String ;
   Ch1, Ch2            :  Char;
   hPos, hDataPos, cSw, i, w, w2 :  Integer ;
   bBreak             :  Boolean ;
   hErrPos            :  Integer ;
   hStartDataPos         :  Integer ;
   sChar             :  String[4];
   cAscii, cCode         :  Char;
   sRet              :  String ;
   fData             :  Real ;
   hStatus :  Integer ;

Const 
   cSymbol :  set Of char =  [#32..#255];
   cEXITKEY1 :  set Of char =  [STOPKEY, CRKEY, ESCKEY];
   cEXITKEY2 :  Set Of char =  [UPKEY, DOWNKEY, RIGHTKEY, LEFTKEY,
                               PGDNKEY,PGUPKEY,HOMEKEY, DELKEY,
                               CTRL_R, CTRL_C,              { Paging }
                               CTRL_RIGHTKEY, CTRL_LEFTKEY, { Paging }
                               CTRL_A,                { AdjustAllWidth}
                               CTRL_B, CTRL_K,        { Mark Cells}
                               CTRL_X,                { Undo }
                               CTRL_N, CTRL_Y,        { Ins/Del Row }
                               ALT_NKEY, ALT_YKEY,    { Ins/Del Col }
                               ALT_LKEY, ALT_SKEY,    { Col Width  }
                               CTRL_Q,
                               CTRL_F,                { Change Form }
                               F1KEY,                 { Help        }
                               F3KEY, F4KEY,          { Col Width   }
                               F5KEY];                { Call Menu  }


  { ---------------------------------------------- }
  {  LEFT KEY                                      }
  { ---------------------------------------------- }
Function ShiftLT(sInpString : String ; hStartDataPos: Integer):  Integer ;

Var 
   hDataPosX  :  Integer ;

   Label BPOINT;

Begin
   hDataPosX := hStartDataPos-1 ;
   If hDataPosX = 0 Then
      Begin
         ShiftLT := hStartDataPos ;
         exit ;
      End ;

{$IFDEF WIN}
   If (hDataPos > 2) And (IsMultiByte(sInpString[hDataPosX-1])) Then
      hDataPosX := hDataPosX - 2
   Else
      hDataPosX := hDataPosX - 1 ;

   If hDataPosX < 1 Then
      hDataPosX := 1 ;
{$ELSE NOWIN}

   // while UTF-8(2-4byte) code 
   While (Byte(sInpString[hDataPosX]) And $80 = $80) Do
      Begin
         // if first byte of UTF-8 code
         If (Byte(sInpString[hDataPosX]) And $40 = $40) Then
            goto BPOINT ;
         hDataPosX := hDataPosX - 1;
      End ;
   BPOINT:

{$ENDIF WIN}

            ShiftLT := hDataPosX
End ;

  { ---------------------------------------------- }
Procedure InputLine_LT ;

Var 
   i, hPosx :  Integer ;
Begin

   If hPos = 1 Then
      Begin
         hStartDataPos := ShiftLT(sInpString, hStartDataPos) ;
         hDataPos := hStartDataPos ;
         exit ;
      End ;

{$IFDEF WIN}
   If (hDataPos > 2) And (IsMultiByte(sInpString[hDataPos-2])) And (IsMultiByte2(sInpString[hDataPos
      -1])) Then
      Begin
         hDataPos := hDataPos - 2 ;
         hPos := hPos - 2 ;
      End
   Else
      Begin
         hDataPos := hDataPos - 1 ;
         hPos := hPos - 1 ;
      End ;

{$ELSE NOWIN}

   If Not (IsMultiByte(sInpString[hDataPos-1])) Then
      Begin
         hDataPos := hDataPos - 1 ;
         hPos := hPos - 1 ;
         exit ;
      End ;

   i := 1 ;
   hPosx := hDataPos ;

   // Scan first byte(11xx xxxx) of UTF-8(2-4byte) code 
   While ((Byte(sInpString[hPosx-i]) And $c0) <> $c0) Do
      Begin

         hDataPos := hDataPos - 1 ;
         i := i + 1 ;
      End ;
   hDataPos := hDataPos - 1 ;
   hPos := hPos - 2 ;

{$ENDIF WIN}

   //GotoXy(20,21) ;
   //Write('hDataPos2:', hDataPos, ' hPos2:', hPos, '   ') ;

End;

  { ---------------------------------------------- }
  {  RIGHT KEY                                     }
  { ---------------------------------------------- }
Function ShiftRT(sInpString : String ; hStartDataPos, hEndDataPos: Integer):  Integer ;

Var 
   v, v2, w, w2 :  integer ;
   j, j2, k, k2 :  integer ;
   hDataPosX  :  Integer ;
Begin
   hDataPosX := hStartDataPos ;
   v := CharLength(sInpString[hEndDataPos], v2) ;
   w := CharLength(sInpString[hEndDataPos+v], w2) ;

   j := CharLength(sInpString[hStartDataPos], j2) ;
   hDataPosX := hDataPosX + j ;
   If w > j Then
      Begin
         k := CharLength(sInpString[hStartDataPos+j], k2) ;
         hDataPosX := hDataPosX + k ;
      End ;
   //GotoXy(20,21) ;
   //Write('hDataPosX:', hDataPosX, ' hDataPos:', hDataPos, '   ') ;

   ShiftRT := hDataPosX ;
End ;

  { ---------------------------------------------- }
Procedure InputLine_RT ;

Var 
   w, w2       :  Integer ;
   hEndDataPos, hEndPos :  Integer ;
Begin
   If hDataPos > length(sInpString) Then
      exit ;

   GetCellDataWithEndPos(sInpString, hStartDataPos, MaxCol, hEndDataPos, hEndPos) ;

   If hPos = hEndPos Then
      Begin
         hStartDataPos := ShiftRT(sInpString, hStartDataPos, hDataPos) ;
         GetCellDataWithEndPos(sInpString, hStartDataPos, MaxCol, hDataPos, hPos) ;
         exit ;
      End ;

   If hDataPos <= length(sInpString) Then
      Begin
         w := CharLength(sInpString[hDataPos], w2) ;
         hDataPos := hDataPos+w;
         hPos := hPos+w2
      End ;
End;

  { ---------------------------------------------- }
  {  BS KEY                                        }
  { ---------------------------------------------- }
Procedure InputLine_BS ;

Var 
   hPosx, i :  Integer ;
Begin
{$IFDEF WIN}
   If (Length(sInpString)>0) And (hDataPos>1) Then
      Begin
         //if not (IsMultiByte(copy(sInpString,hDataPos-1,1)[1])) then


{
	    if (hDataPos = 2) or (not (IsMultiByte(sInpString[hDataPos-2]))) then
	       begin
		  delete(sInpString,hDataPos-1,1);
		  hPos:=hPos-1;
		  hDataPos:=hDataPos-1;
		  exit ;
	       end ;
	    if (IsMultiByte(sInpString[hDataPos-2])) then
}
         If (hDataPos > 2) And (IsMultiByte(sInpString[hDataPos-2])) And (IsMultiByte2(sInpString[
            hDataPos-1])) Then
            Begin
               delete(sInpString,hDataPos-2,2);
               hPos := hPos-2;
               hDataPos := hDataPos-2;
            End
         Else
            Begin
               delete(sInpString,hDataPos-1,1);
               hPos := hPos-1;
               hDataPos := hDataPos-1;
            End ;
      End;

{$ELSE NOWIN}

   If (Length(sInpString)>0) And (hDataPos>1) Then
      Begin
         //if not (IsMultiByte(copy(sInpString,hDataPos-1,1)[1])) then
         If Not (IsMultiByte(sInpString[hDataPos-1])) Then
            Begin
               delete(sInpString,hDataPos-1,1);
               hPos := hPos-1;
               hDataPos := hDataPos-1;
               exit ;
            End ;

         i := 1 ;
         hPosx := hDataPos ;
         // Scan first byte(11xx xxxx) of UTF-8(2-4byte) code 
         //while ((Byte(copy(sInpString,hPosx-i,1)[1]) and $c0) <> $c0) do
         While ((Byte(sInpString[hPosx-i]) And $c0) <> $c0) Do
            Begin
               delete(sInpString,hPosx-i,1);
               hDataPos := hDataPos-1;
               i := i + 1 ;
            End ;
         delete(sInpString,hPosx-i,1);
         hDataPos := hDataPos - 1 ;
         hPos := hPos - 2 ;
      End;
{$ENDIF WIN}

End ;

  { ---------------------------------------------- }
  {  DEL KEY                                       }
  { ---------------------------------------------- }
Procedure InputLine_DEL ;

Var 
   w, w2 :  Integer ;
Begin
   If (Length(sInpString)>0) Then
      Begin
         w := CharLength(copy(sInpString,hDataPos,1)[1], w2) ;
         delete(sInpString, hDataPos, w) ;
      End ;
End;

  { ---------------------------------------------- }
  {  Edit mode                                     }
  { ---------------------------------------------- }
Procedure InputLine_Edit  ;
Begin
   sInpString := sData;
   //cEdit:= Chr(Byte(cEdit) or c_EDIT);
   cEdit := cEdit Or c_EDIT;
   GetCellDataWithEndPos(sInpString, hStartDataPos, MaxCol, hDataPos, hPos);  { CVCRT }
End;

  { ---------------------------------------------- }
  {  Set CellMark                                  }
  { ---------------------------------------------- }
Procedure SetCellMark;

Var 
   sWork:  String ;
Begin
   //str(ghTopRow:3,sWork) ;
   str(ghTopRow:ghLeftSide-1,sWork) ;
   Trim(sWork);
   sInpString := sInpString + GetXString(ghTopCol) + sWork ;
   If (ghTopCol <> ghEndCOl) Or (ghTopRow <> ghEndRow) Then
      Begin
         //str(ghEndRow:3,sWork) ;
         str(ghEndRow:ghLeftSide-1,sWork) ;
         Trim(sWork) ;
         sInpString := sInpString + '..' + GetXString(ghEndCol) + sWork;
      End ;

   hPos := CellLength(sInpString)+1;
   hDataPos := Length(sInpString)+1;
End ;

{ ---------------------------------------------- }
{  Start here                                    }
{ ---------------------------------------------- }
Begin
   If cEdit And c_EDIT <> 0 Then
      sInpString := sData
   Else
      sInpString := '' ;

   hPos := 1;
   hDataPos := 1;
   hStartDataPos := 1;
   Repeat
      ClearMenuLine ;       { CVSCRN }
      Gotoxy(1,h_INPUTLINE);
      WriteString(GetCellData(sInpString, hStartDataPos, MaxCol), gcLineAttr);  { CVCRT }

      //GotoXy(20,19) ;
      //Write('cEdit:', ord(cEdit)) ;
      //Write(' hStartDataPos:', hStartDataPos, ' hDataPos:', hDataPos, ' hPos:', hPos, '   ') ;

      GotoXY(hPos,h_INPUTLINE);

      GetKey(Ch1, cCode);  { CVCRT }

     { ---------------------------------------------- }
     {  Check if Command or Formula                   }
     { ---------------------------------------------- }
      If (Ch1='/') And (sInpString='') And (cEdit And c_EDIT = 0) Then
         Begin
            Ch1 := #0 ;
            cCode := F5KEY;
         End ;
      If (Ch1='+') And (sInpString='') And (cEdit And c_EDIT = 0) Then
         Begin
            Ch1 := #0 ;
            cCode := F6KEY;
         End ;

     { ---------------------------------------------- }
     {  Get Utf-8 Strings (1-4 bytes)                 }
     { ---------------------------------------------- }
      sChar := '' ;
      sChar := Ch1;
      w := CharLength(Ch1, w2) ;
      For i := 1 To w-1 Do
         Begin
            GetKey(Ch2, cCode);
            sChar := sChar + Ch2;
         End ;

      cSw := 0 ;
      bBreak := False ;
      If Ch1 = #0 Then
         Begin
            cSw := 1 ;
            Ch1 := cCode ;

            If (Ch1 In cEXITKEY1)
               Or ((Ch1 In cEXITKEY2) And (cEdit And c_EDIT = 0)) Then
               bBreak := True ;
         End ;

     { ---------------------------------------------- }
     {  Go to CALC Mode                               }
     { ---------------------------------------------- }
      If cCode = F6KEY Then
         SetCalcModeMessage(cEdit) ;

     { ---------------------------------------------- }
     {  Ascii Code > $20                              }
     { ---------------------------------------------- }
      If (Ch1 In cSymbol) And (cSw = 0) Then
         Begin
            If Length(sInpString) <= h_MAXCELLSIZE-Length(sChar) Then
               Begin
                  If hDataPos<=Length(sInpString) Then
                     Begin
                        insert(sChar,sInpString,hDataPos);
                        hPos := hPos+CellLength(sChar);
                        hDataPos := hDataPos+length(sChar);
                     End
                  Else
                     Begin
                        sInpString := sInpString+sChar;
                        hPos := hPos+CellLength(sChar);
                        hDataPos := hDataPos+length(sChar);
                     End;
               End ;
         End

     { ---------------------------------------------- }
     {  Control key                                   }
     { ---------------------------------------------- }
      Else
         If (Ch1 In cEXITKEY1)
            Or ((Ch1 In cEXITKEY2) And (cEdit And c_EDIT = 0)) Then
            Begin
               cExKey := Ch1;
{
							 If length(sInpString)<>0 Then
                  sData := sInpString;
}
                { ************************************** }
                {  Calc Mode                           * }
                { ************************************** }
               If (cEdit And c_CALC <> 0) And (Ch1 = CRKEY) Then
                  Begin
                     //sInpString := UpcaseString(sInpString);

                     Evaluate(hStatus, sInpString, fData, hErrPos, sRet, ghX, ghY) ;

                     If hErrPos = 0 Then
                        Begin
                           SetCellFormula(Sheet[ghX,ghY]) ;
                           //                           If sRet = '' Then
                           If hStatus = h_FORMULA_NUMERIC Then
                              SetCellNumeric(Sheet[ghX,ghY])
                           Else
                              SetCellNotNumeric(Sheet[ghX,ghY]) ;
                           bBreak := True ;
                        End
                     Else
                        Begin
                           WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
                           GotoXY(1,h_GUIDELINE);
                           //Write('Formula Error. POS:', hErrPos-1, ' ',hStatus);
                           Write('Formula Error. POS:', hErrPos-1, ' (',hStatus, ')', GetErrMsg(
                                 hStatus));
                           //GotoXY(hErrPos,h_INPUTLINE) ;
                        End ;
                  End
               Else
                  Begin
                     // ESC->clear input.
                     //If (Ch1 = ESCKEY) And (cEdit And c_EDIT = 0)Then
                     If Ch1 = ESCKEY Then
                        Begin
                           //													 if (cEdit And c_EDIT = 0) or (cEdit And c_CALC <> 0) then
                           If (cEdit And c_EDIT = 0)  Then
                              sInpString := '' ;
                        End ;

                     bBreak := True ;
                  End ;
            End
      Else
         Begin
            Case Ch1 Of 
               LEFTKEY, F3KEY  :  InputLine_LT ;
               RIGHTKEY, F4KEY  :  InputLine_RT ;
               BSKEY     :  InputLine_BS ;
               DELKEY    :  InputLine_DEL ;

               F2KEY, ALT_IKEY    :
                                     Begin
                                        If length(sData)<>0 Then
                                           Begin
                                              InputLine_Edit ;
                                              If IsCellFormula(Sheet[ghX,ghY]) Then
                                                 SetCalcModeMessage(cEdit) ;
                                           End ;
                                     End ;

               CTRL_K    :
                            Begin
                               If (cEdit And c_CALC <> 0) Then
                                  Begin
                                     GetKey(cAscii, cCode) ;

                                     If cAscii = CTRL_C Then
                                        SetCellMark ;
                                  End ;
                            End ;

               TABKEY    :
                            Begin
                               If (cEdit And c_CALC <> 0) Then
                                  SetCellMark ;

                            End ;
            End;
         End ;

   Until (bBreak = True);
   If length(sInpString)<>0 Then
      sData := sInpString
   Else If cEdit And c_EDIT <> 0 Then
           sData := '' ;
End;
End .
