{$MODE TP}

Unit cvsort ;

Interface

Uses cvdef, cvcrt, cvscrn, cvinpt, cveval, cvconv, cvundo ;
Procedure SortCellData;

{ ---------------------------------------------- }

Implementation
{ ---------------------------------------------- }
{  Sort Cell Data                                }
{ ---------------------------------------------- }

Procedure ResetCell(Var hStatus : Integer;
                    Var sData   : String ;
                    Var fData   : Real;
                    Var hErrPos  : Integer;
                    Var sRet   : String;
                    x, y   :Integer ) ;
Begin
   If hErrPos <> 0 Then
      Begin
         SetCellNotFormula(Sheet[x,y]) ;
         fData := 0 ;
      End
   Else If sRet <> '' Then
           Begin
              sData := sRet;
           End;


{
					else
					   str(fData:0:GetCellDecPoint(x,y),Sheet[x,y].tpMain^)
					SetCellNotFormula(Sheet[x,y]);
}

End ;

Procedure SortCellData;

Var 
   cBors:  Byte;
   sOrgData,sDesData :  String;
   fOrgData,fDesData :  Real;
   hErrPos:  Integer ;
   bBrk:  Boolean  ;
   hStatus:  Integer ;

   hMenu:  Integer ;
   hStartX, hEndX:  Integer ;
   hStartY, hEndY:  Integer ;
   sRet :  String;

Const 
   h_ASC =  1;
   h_DEC =  2;

  { ---------------------------------------------- }
  {  Quick Sort Routine                            }
  { ---------------------------------------------- }
Procedure QS(L,R,K : integer) ;

Var 
   tpPtr:  tpCell ;
   s, t, i, j    :  integer ;
Begin
   s := L;
   t := R;
   j := ((L+R) Div 2) ;
   If Sheet[K,j].tpMain<>Nil Then
      Begin
         sOrgData := Sheet[K,j].tpMain^ ;
         If IsCellFormula(Sheet[K,j]) Then
            Begin
               Evaluate(hStatus,sOrgData,fOrgData,hErrPos,sRet,K,j);
               ResetCell(hStatus,sOrgData,fOrgData,hErrPos,sRet,K,j);
            End
         Else If IsCellNumeric(Sheet[K,j]) Then
                 val(sOrgData,fOrgData,hErrPos);
      End
   Else
      Begin
         sOrgData := '' ;
         fOrgData := 0 ;
      End;

   Repeat
      bBrk := true ;
      While (bBrk) Do
         Begin
            If Sheet[K,s].tpMain<>Nil Then
               Begin
                  sDesData := Sheet[K,s].tpMain^  ;
                  If IsCellFormula(Sheet[K,s]) Then
                     Begin
                        Evaluate(hStatus,sDesData,fDesData,hErrPos,sRet,K,s);
                        ResetCell(hStatus,sDesData,fDesData,hErrPos,sRet,K,s);
                     End
                  Else If IsCellNumeric(Sheet[K,s]) Then
                          val(sDesData,fDesData,hErrPos);
               End
            Else
               Begin
                  sDesData := '' ;
                  fDesData := 0;
               End ;

            If IsCellNumeric(Sheet[K,s]) Then
               Begin
                  If ((fDesData>fOrgData)
                     And (cBors=h_DEC))
                     Or ((fDesData<fOrgData)
                     And (cBors=h_ASC)) Then
                     s := s+1
                  Else
                     bBrk := false;
               End
            Else
               Begin
                  If ((sDesData>sOrgData)
                     And (cBors=h_DEC))
                     Or ((sDesData<sOrgData)
                     And (cBors=h_ASC)) Then
                     s := s+1
                  Else
                     bBrk := false;
               End ;
         End ;

      bBrk := true ;
      While (bBrk) Do
         Begin
            If Sheet[K,t].tpMain<>Nil Then
               Begin
                  sDesData := Sheet[K,t].tpMain^  ;
                  If IsCellFormula(Sheet[K,t]) Then
                     Begin
                        Evaluate(hStatus,sDesData,fDesData,hErrPos,sRet,K,t);
                        ResetCell(hStatus,sDesData,fDesData,hErrPos,sRet,K,t);
                     End
                  Else If IsCellNumeric(Sheet[K,t]) Then
                          val(sDesData,fDesData,hErrPos);
               End
            Else
               Begin
                  sDesData := '' ;
                  fDesData := 0;
               End ;

            If IsCellNumeric(Sheet[K,t]) Then
               Begin
                  If ((fOrgData>fDesData)
                     And (cBors=h_DEC))
                     Or ((fOrgData<fDesData)
                     And (cBors=h_ASC)) Then
                     t := t-1
                  Else
                     bBrk := false ;
               End
            Else
               Begin
                  If ((sOrgData>sDesData)
                     And (cBors=h_DEC))
                     Or ((sOrgData<sDesData)
                     And (cBors=h_ASC)) Then
                     t := t-1
                  Else
                     bBrk := false ;
               End ;
         End;

      If s <= t Then
         Begin

            For i:=hStartX To hEndX Do
               Begin
                  tpPtr := Sheet[i,s] ;
                  Sheet[i,s] := Sheet[i,t] ;

                  If  (Sheet[i,s].tpMain<>Nil)
                     And (IsCellFormula(Sheet[i,s])) Then
                     Begin
                        sDesData := Sheet[i,s].tpMain^ ;
                        //GotoXy(19,21) ;
                        //Write('Evaluate 04:', ' sDesData:', sDesData, ' x', i, ' y', s) ;

                        Evaluate(hStatus,sDesData,fDesData,hErrPos,sRet,i,s);
                        ResetCell(hStatus,sDesData,fDesData,hErrPos,sRet,i,s);
                     End ;
                  //Convert(Sheet[i,s].tpMain^,i,t,i,s,c_CHANGE_XY) ;

                  Sheet[i,t] := tpPtr ;

                  If  (Sheet[i,t].tpMain<>Nil)
                     And (IsCellFormula(Sheet[i,t])) Then
                     Begin
                        sDesData := Sheet[i,t].tpMain^ ;
                        //GotoXy(19,21) ;
                        //Write('Evaluate 05:', ' sDesData:', sDesData, ' x', i, ' y', t) ;

                        Evaluate(hStatus,sDesData,fDesData,hErrPos,sRet,i,t);
                        ResetCell(hStatus,sDesData,fDesData,hErrPos,sRet,i,t);
                     End ;
                  //                          Convert(Sheet[i,t].tpMain^,i,s,i,t,c_CHANGE_XY) ;

               End ;

            s := s+1 ;
            t := t-1 ;
         End ;
   Until s>t ;

   If L<t Then QS(L,t,K) ;
   If L<R Then QS(s,R,K) ;
End ;

Var 
   hEdit  :  Byte ;
   cAscii  :  Char ;
   sCell  :  String ;
   hCell  :  Integer ;
   i, j   :  Integer;
   sErrMsg :  String ;
   cScan  :  Char;
   hAttr :  Byte ;

   Label BPOINT;

Begin
   WriteRange;      { Write target Range CVSCRN }
   ClearMenuLine ;

   If (ghTopCol = ghEndCol) And (ghTopRow = ghEndRow) Then
      Begin
         hStartX := 1 ;
         hEndX := ghMaxX ;
         hStartY := 1 ;
         hEndY := ghMaxY ;
      End
   Else
      Begin
         hStartX := ghTopCol ;
         hEndX := ghEndCol ;
         hStartY := ghTopRow ;
         hEndY := ghEndRow ;
      End ;

   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   Write('Sort range['+GetXString(hStartX), hStartY, '..', GetXString(hEndX), hEndY, ']') ;

   ClearMenuLine ;
{$IFDEF JPN}
   MenuList[1].Menu := 'A 昇順  ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'D 降順  ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'Q 終了  ';
   MenuList[3].Num  := 0;
{$ELSE NOJPN}

   MenuList[1].Menu := 'A Ascending  ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'D Descending ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'Q Quit       ';
   MenuList[3].Num  := 0;
{$ENDIF JPN}
   cBors :=  SelectVertMenu(1, h_MENULINE, MenuList, 3);
   If cBors = 0 Then
      exit ;

  { ******************************** }
  {   Get Column                     }
  { ******************************** }
   CursorOn ;
   WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
   GotoXY(1,h_GUIDELINE);
   WriteString('Specify col  [ESC]:CANCEL', gcLineAttr);

   hEdit := c_Edit ;
   GetInputData(sCell, cAscii, hEdit);
   If (sCell = '') Or (cAscii = ESCKEY) Then
      exit ;

   // Error check
   hCell := GetXInteger(UpcaseString(sCell)) ;
   If hCell = 0 Then
      Begin
         WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
         GotoXY(1,h_GUIDELINE);
         WriteString('Sorted col error.. Hit any key', gcLineAttr);
         GetKey(cAscii, cScan);
         exit ;
      End ;

  { Check formula }
   sErrMsg := '' ;
   For j:=hStartY To hEndY Do
      Begin
         For i:=hStartX To hEndX Do
            Begin
               If IsCellFormula(sheet[i,j]) Then
                  Begin
                     sErrMsg := 'ERROR:"Formula" exists within Range.. Hit any key' ;
                     goto BPOINT ;
                  End;
            End ;
      End;

  { Check Col }
   If (hCell < hStartX) Or (hCell > hEndX) Then
      Begin
         sErrMsg := 'ERROR:Col is out of "Range".. Hit any key' ;
         goto BPOINT ;
      End;

  { Check Attribute }
   hAttr := 0 ;
   For j:=hStartY To hEndY Do
      Begin
         If IsCellNumeric(sheet[hCell,j]) Then
            Begin
               If (hAttr <> 0) And (hAttr <> 1) Then
                  Begin
                     sErrMsg := 'ERROR:"Numeric" and "String" are mixed.. Hit any key' ;
                     goto BPOINT ;
                  End;
               hAttr := 1 ;
            End
         Else If sheet[hCell,j].tpMain <> Nil Then
                 Begin
                    If (hAttr <> 0) And (hAttr <> 2) Then
                       Begin
                          sErrMsg := 'ERROR:"Numeric" and "String" are mixed.. Hit any key' ;
                          goto BPOINT ;
                       End;
                    hAttr := 2 ;
                 End ;
      End ;

   BPOINT:
            If sErrMsg <> '' Then
               Begin
                  WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
                  GotoXY(1,h_GUIDELINE);
                  WriteString(sErrMsg, gcLineAttr);
                  GetKey(cAscii, cScan);
                  exit;
               End ;

   QS(hStartY, hEndY, hCell);
   //      End ;

   ClearStack ;         { CVUNDO }
   SetScreen ;
   CursorOff ;
End;
End .
