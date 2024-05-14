{$MODE TP}

Unit cvdata ;

Interface

Uses cvdef, cvcrt, cvscrn, cveval, cvundo ;
Procedure CheckData(sData:String;Var tpPtr:tpCell);
Procedure SetMaxX(hX:Integer) ;
Procedure SetMaxY(hY:Integer) ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Set MaxX                                      }
{ ---------------------------------------------- }
Procedure SetMaxX(hX:Integer) ;
Begin
   If ghMaxX < hX Then
      ghMaxX := hX;
End ;

{ ---------------------------------------------- }
{  Set MaxY                                      }
{ ---------------------------------------------- }
Procedure SetMaxY(hY:Integer) ;
Begin
   If ghMaxY < hY Then
      ghMaxY := hY;
End ;

{ ---------------------------------------------- }
{  Check Input Data                              }
{ ---------------------------------------------- }
Procedure CheckData(sData:String;Var tpPtr:tpCell);

Var 
   fData  :  real;
   hResult :  integer;
   decPoint  :  Byte ;
Begin

 { ************************************** }
 {  Data is inputed                     * }
 { ************************************** }
   If sData <> '' Then
      Begin
         //				 UndoLog(0,ghX,ghY) ;   { CVUNDO }

       { ************************************** }
       {  New Data                            * }
       { ************************************** }
         If tpPtr.tpMain = Nil Then
            Begin
               GetCellArea(tpPtr) ;
               Sheet[ghX, ghY].tpMain := tpPtr.tpMain ;
            End ;

     { ************************************** }
     {  Data is changed                     * }
     { ************************************** }
         If sData <> tpPtr.tpMain^ Then
            Begin
               UndoLog(0,ghX,ghY) ;   { CVUNDO }
               tpPtr.tpMain^ := sData;
               ChangeFirstChar(sData) ;

               If Not IsCellFormula(tpPtr) Then
                  Begin
                     val(sData,fData,hResult);
                     //if hResult <> 0 then
                     If (hResult <> 0) Or (sData[1] = 'E') Or (sData[1] = 'e') Then
                        //SetCellNotFormula(tpPtr)
                        SetCellNotNumeric(tpPtr)
                     Else
                        Begin
                           SetCellNumeric(tpPtr) ;
                           decPoint := pos('.', sData) ;
                           If decPoint <> 0 Then
                              Begin
                                 decPoint := length(sData) - decPoint ;
                                 If (decPoint >=1) And (decPoint <=h_MAXDECPOINT) Then
                                    SetCellDecPoint(ghx, ghy, decPoint)
                                 Else
                                    Begin
                                       FreeCellArea(tpPtr) ;
                                       SetCellNotNumeric(tpPtr);
                                    End ;
                              End ;
                        End ;
                  End ;


{
              if IsCellRecalc(tpPtr) then
                 begin
                    WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
                    GotoXY(1,h_GUIDELINE);
                    SetScreenDetail(gcCellAttr) ;
                 end ;
}
            End;

         SetMaxX(ghX) ;
         SetMaxY(ghY) ;

      End

  { ************************************** }
  {  Data is cleared                     * }
  { ************************************** }
   Else
      Begin
         If tpPtr.tpMain<>Nil Then
            Begin
               UndoLog(0,ghX,ghY) ;   { CVUNDO }
               FreeCellArea(tpPtr);
               Sheet[ghX, ghY].tpMain := Nil ;
            End ;
      End ;

   If IsCellRecalc(tpPtr) Then
      Begin
         WinClrScr( 1, h_GUIDELINE, MaxCol, h_GUIDELINE );
         GotoXY(1,h_GUIDELINE);

         SetScreenDetail ;
      End ;
End;
End .
