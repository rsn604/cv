{$MODE TP}

Unit cvnewcl ;

Interface

Uses cvdef, cvscrn, cveval, cvconv, cvcrt, cvundo ;
Procedure InsertNewRow(hRow:integer);
Procedure DeleteCurrentRow(hRow:integer);
Procedure InsertNewCol(hCol:integer);
Procedure DeleteCurrentCol(hCol:integer);
Procedure GotoCell(hX, hY:Integer) ;
Procedure ChangeFormMode;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Insert New Row                                }
{ ---------------------------------------------- }
Procedure InsertNewRow(hRow:integer);

Var 
   i, j:  Integer ;

Begin
   If ghMaxY < h_MAXROW Then
      Begin

        { ************************************** }
        {  Slide Down Row                      * }
        { ************************************** }
         For j:=ghMaxY+1 Downto hRow+1 Do
            For i:=1 To ghMaxX Do
               Begin
                  Sheet[i,j] := Sheet[i,j-1] ;
                  //Sheet[i,j].cCellColor := gcCellAttr;

                  If IsCellFormula(Sheet[i,j]) Then
                     Convert(Sheet[i,j].tpMain^,i,j-1,i,j,c_CHANGE_Y) ;
               End;

        { ************************************** }
        {  Change  Row                         * }
        { ************************************** }
         For j:=1 To hRow-1 Do
            For i:=1 To ghMaxX Do
               Begin
                  If IsCellFormula(Sheet[i,j]) Then
                     Convert(Sheet[i,j].tpMain^,i,j,i,j+1,c_CHANGE_Y) ;
               End;

        { ************************************** }
        {  New Row                             * }
        { ************************************** }
         For i:=1 To ghMaxX Do
            Begin
               Sheet[i,hRow].tpMain := Nil;
               Sheet[i,hRow].cCellColor := gcCellAttr;

               If gcForm=c_NOGUIDE Then
                  Sheet[i,hRow].cCellType := Sheet[i,hRow+1].cCellType And $0F
               Else
                  Sheet[i,hRow].cCellType := 0;
            End ;

         ghMaxY := ghMaxY+1 ;

      End;
   UndoLog(h_INSROW, -1, hRow) ;   { CVUNDO }

   SetScreen ;
End ;

{ ---------------------------------------------- }
{  Delete Current Row                            }
{ ---------------------------------------------- }
Procedure DeleteCurrentRow(hRow:integer);

Var 
   i, j, hSeq:  Integer ;

Begin
   hSeq := 0 ;

   For j:=hRow To ghMaxY Do
      For i:=1 To ghMaxX Do
         Begin

           { ************************************** }
           {  Delete Current Row                  * }
           { ************************************** }
            If (Sheet[i,j].tpMain<>Nil) And (hRow = j) Then
               Begin
                  UndoLog(hSeq,i, j) ;   { CVUNDO }
                  hSeq := hSeq + 1 ;
                  FreeCellArea(Sheet[i,j]);
               End;
           { ************************************** }
           {  Slide Up Row                        * }
           { ************************************** }
            Sheet[i,j] := Sheet[i,j+1];
            If IsCellFormula(Sheet[i,j]) Then
               Convert(Sheet[i,j].tpMain^,i,j+1,i,j,c_CHANGE_Y) ;
         End ;

  { ************************************** }
  {  Change  Row                         * }
  { ************************************** }
   For j:=1 To hRow-1 Do
      For i:=1 To ghMaxX Do
         Begin
            If IsCellFormula(Sheet[i,j]) Then
               //Convert(Sheet[i,j] .tpMain^,i,j,i,j+1,c_CHANGE_Y) ;
               Convert(Sheet[i,j] .tpMain^,i,j,i,j-1,c_CHANGE_Y) ;
         End;


  { ************************************** }
  {  Delete Last Row                     * }
  { ************************************** }
   If hRow < ghMaxY Then
      Begin
         For i:=1 To ghMaxX Do
            Begin
               Sheet[i,ghMaxY].tpMain := Nil;
               Sheet[i,ghMaxY].cCellType := 0;
               Sheet[i,ghMaxY].cCellColor := gcCellAttr;
            End ;

         If ghMaxY > 1 Then
            ghMaxY := ghMaxY - 1 ;
      End ;
   UndoLog(h_DELROW, -1, hRow) ;   { CVUNDO }

   SetScreen ;
End ;

{ ---------------------------------------------- }
{  Insert New Col                                }
{ ---------------------------------------------- }
Procedure InsertNewCol(hCol:integer);

Var 
   i, j:  Integer ;

Begin

   If ghMaxX < h_MAXCOL Then
      Begin

        { ************************************** }
        {  Slide Col to left                   * }
        { ************************************** }

         For i:=ghMaxX+1 Downto hCol+1 Do
            Begin
               For j:=1 To ghMaxY Do
                  Begin
                     Sheet[i,j] := Sheet[i-1,j];
                     //Sheet[i,j].cCellColor := gcCellAttr;

                     If IsCellFormula(Sheet[i,j]) Then
                        Convert(Sheet[i,j].tpMain^,i-1,j,i,j,c_CHANGE_X) ;
                  End ;
               gcCellWidth[i] := gcCellWidth[i-1];
            End ;

      { ************************************** }
        {  Change  Col                         * }
        { ************************************** }
         For i:=1 To hCol-1 Do
            For j:=1 To ghMaxY Do
               Begin
                  If IsCellFormula(Sheet[i,j]) Then
                     Convert(Sheet[i,j].tpMain^,i,j,i+1,j,c_CHANGE_X) ;
               End;


        { ************************************** }
        {  New Col                             * }
        { ************************************** }
         gcCellWidth[hCol] := h_INIT_CELLSIZE;
         For j:=1 To ghMaxY Do
            Begin
               Sheet[hCol,j].tpMain := Nil;
               Sheet[hCol,j].cCellColor := gcCellAttr;

               If gcForm=c_NOGUIDE Then
                  Sheet[hCol,j].cCellType := Sheet[hCol+1,j].cCellType And $0F
               Else
                  Sheet[hCol,j].cCellType := 0;
            End ;

         ghMaxX := ghMaxX+1 ;
      End;
   UndoLog(h_INSCOL, hCol, -1) ;   { CVUNDO }
   SetScreen ;
End ;

{ ---------------------------------------------- }
{  Delete Current Col                            }
{ ---------------------------------------------- }
Procedure DeleteCurrentCol(hCol:integer);

Var 
   i, j, hSeq:  Integer ;

Begin
   hSeq := 0 ;

  { ************************************** }
  {  Slide Col to right                  * }
  { ************************************** }
   For i:=hCol To ghMaxX Do
      Begin
         For j:=1 To ghMaxY Do
            Begin
               If (Sheet[i,j].tpMain<>Nil) And ( hCol = i) Then
                  Begin
                     UndoLog(hSeq,i, j) ;   { CVUNDO }
                     hSeq := hSeq + 1 ;
                     FreeCellArea(Sheet[i,j]);
                  End ;
               Sheet[i,j] := Sheet[i+1,j] ;
               If IsCellFormula(Sheet[i,j]) Then
                  Convert(Sheet[i,j].tpMain^,i+1,j,i,j,c_CHANGE_X) ;
            End ;

         gcCellWidth[i] := gcCellWidth[i+1];
      End ;

  { ************************************** }
  {  Change  Col                         * }
  { ************************************** }
   For i:=1 To hCol-1 Do
      For j:=1 To ghMaxY Do
         Begin
            If IsCellFormula(Sheet[i,j]) Then
               Convert(Sheet[i,j].tpMain^,i,j,i-1,j,c_CHANGE_X) ;
         End;

  { ************************************** }
  {  Last Col                            * }
  { ************************************** }
   gcCellWidth[ghMaxX] := h_INIT_CELLSIZE;
   For j:=1 To ghMaxY Do
      Begin
         Sheet[ghMaxX,j].tpMain := Nil;
         Sheet[ghMaxX,j].cCellType := 0;
         Sheet[ghMaxX,j].cCellColor := gcCellAttr;
      End ;

   If ghMaxX > 1 Then
      ghMaxX := ghMaxX-1 ;

   UndoLog(h_DELCOL, hCol, -1) ;   { CVUNDO }
   SetScreen ;
End ;

{ ---------------------------------------------- }
{  Goto Cell                                     }
{ ---------------------------------------------- }
Procedure GotoCell(hX, hY:Integer) ;
Begin
   ghX := hX;
   ghY := hY;
   If (hX<ghFirstX) Or (hX>ghFirstX+ghColCount) Then
      Begin
         ghFirstX := ghX ;
         ghScrX := 1 ;
      End
   Else
      ghScrX := ghX - ghFirstX + 1 ;
   If (hY<ghFirstY) Or (hY>ghFirstY+MaxRow-ghColNameLine) Then
      Begin
         ghFirstY := ghY ;
         ghScrY := 1 ;
      End
   Else
      ghScrY := ghY - ghFirstY + 1 ;

   SetScreen ;
End ;

{ ---------------------------------------------- }
{  Change Form Mode                              }
{ ---------------------------------------------- }
Procedure ChangeFormMode;
Begin
   If gcForm <> c_NORMAL Then
      gcForm := c_NORMAL
   Else
      gcForm := c_NOGUIDE ;

   SetScreen ;
End;
End .
