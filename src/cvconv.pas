{$MODE TP}

Unit cvconv ;

Interface

Uses cvdef, cvcrt, cveval ;
Procedure Convert(Var Formula: String ;
                  hOrgx,hOrgy: Integer;
                  hDestx,hDesty: Integer;
                  cSw:byte) ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Convert Formula                               }
{ ---------------------------------------------- }
Procedure Convert(Var Formula: String ;
                  hOrgx,hOrgy: Integer;
                  hDestx,hDesty: Integer;
                  cSw:byte) ;

Var 
   sWork       :  String ;
   Pos        :  Integer;
   Ch         :  Char;
   sEXY        :  string;
   hEFX,hEFY,Err   :  integer ;
   hEFX2,hEFY2   :  integer ;
   cFlagCol, cFlagRow :  Byte ;
   xstr        :  string ;
   bConstant     :  Boolean ;

Procedure NextCh;
Begin
   //   Repeat
   Pos := Pos+1;
   If Pos<=Length(Formula) Then
      Ch := Formula[Pos]
   Else
      Ch := eofline;
   //   Until Ch<>' ';
End ;

Function PreviousCh:  Char;
Begin
   PreviousCh := Formula[Pos-1]
End;

Begin
   Pos := 0;
   NextCh;
   sWork := '' ;
   { Skip first 2 chars }
   If Ch = '\' Then
      Begin
         sWork := sWork+Ch ;
         NextCh ;
         sWork := sWork+Ch ;
         NextCh ;
      End ;

   bConstant := False ;
   Repeat
      If ch = '"' Then
         If bConstant = False Then
            bConstant := True
      Else
         bConstant := False ;
      If  ((Ch In ['A'..'Z','a'..'z']) Or (Ch = '$')) And (PreviousCh <> '0') And (bConstant = False
         ) Then
         Begin
            cFlagCol := 0 ;
            If Ch = '$' Then
               Begin
                  cFlagCol := 1 ;
                  NextCh ;
               End ;
            //@@@@
            xstr := Ch ;
            NextCh;
            While (Ch In ['A'..'Z','a'..'z']) Do
               Begin
                  xstr := xstr + Ch ;
                  NextCh;
               End ;
            hEFX := GetXInteger(UpcaseString(xstr)) ;   {CVEVAL}

            cFlagRow := 0 ;
            If Ch = '$' Then
               Begin
                  cFlagRow := 1 ;
                  NextCh ;
               End ;

            If  (Ch In Numbers) Then
               Begin
                  sEXY := '' ;
                  While (Ch In Numbers) And (Ch <> Eofline) Do
                     Begin
                        sEXY := sEXY+Ch;
                        NextCh;
                     End ;
                  val(sEXY,hEFY,Err);

       { ********************************************* }
       {    Change ColName  'A'-'Z'                    }
       { ********************************************* }
                  // @@@ Save Cell XY
                  hEFX2 := hEFX ;
                  hEFY2 := hEFY ;
                  If ((cSw And c_CHANGE_X)<> 0) And (cFlagCol = 0) Then
                     Begin
                        If ((cSw And c_CHANGE_Y) = 0)  And (ghX<=hEFX) Then
                           hEFX := hEFX+hDestx-hOrgx  ;
                        If (cSw And c_CHANGE_Y) <> 0 Then
                           hEFX := hEFX+hDestx-hOrgx  ;
                     End;
                  //@@@@ Check X
                  If (hEFX < 1) Or (hEFX > h_MAXCOL) Then
                     hEFX := hEFX2 ;

                  If cFlagCol = 1 Then
                     sWork := sWork+'$' ;
                  sWork := sWork+GetXString(hEFX);  {CVEVAL}

       { ********************************************* }
       {    Change RowNumber                           }
       { ********************************************* }
                  If ((cSw And c_CHANGE_Y)<>0)  And (cFlagRow = 0)Then
                     Begin
                        //if ((cSw and c_CHANGE_X)=0) and (ghY<=hEFY) then
                        If ((cSw And c_CHANGE_X)=0) And (ghY<=hEFY) Then
                           hEFY := hEFY+hDesty-hOrgy ;
                        If (cSw And c_CHANGE_X)<>0  Then
                           hEFY := hEFY+hDesty-hOrgy ;
                     End ;

                  //@@@@ Check Y
                  If (hEFY < 1)  Or (hEFY > h_MAXROW) Then
                     hEFY := hEFY2 ;

                  str(hEFY,sEXY);

                  If Sheet[hEFX,hEFY].tpMain<>Nil Then
                     SetCellRecalc(Sheet[hEFX,hEFY]) ;
                  If cFlagRow = 1 Then
                     sWork := sWork+'$' ;
                  sWork := sWork+sEXY ;
               End
            Else
               Begin
                  sWork := sWork+xstr ;
               End;
         End
      Else
         Begin
            sWork := sWork+Ch ;
            NextCh;
         End ;
   Until Ch=EofLine ;
   Formula := sWork ;
End ;
End .
