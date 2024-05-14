{$MODE TP}

Program cv ;

Uses cvdef, cvcrt, cvscrn, cvinpt, cvdata, cvfile, cvtrans, cvcell, cvnewcl, cvsort, cvundo, cvhelp;

{ ---------------------------------------------- }
{  Get Menu                                      }
{ ---------------------------------------------- }
Function GetMenuNum:  Char;

Var 
   hMenu:  Integer ;
Begin
   CursorOff ;
   ClearMenuLine ;

{$IFDEF JPN}
   MenuList[1].Menu := 'F ファイル ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'T 変換     ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C セル     ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'S ソート   ';
   MenuList[4].Num  := 4;
   MenuList[5].Menu := 'H ヘルプ   ';
   MenuList[5].Num  := 5;
   MenuList[6].Menu := 'Q 終了     ';
   MenuList[6].Num  := 6;

{$ELSE NOJPN}
   MenuList[1].Menu := 'F File    ';
   MenuList[1].Num  := 1;
   MenuList[2].Menu := 'T Trans   ';
   MenuList[2].Num  := 2;
   MenuList[3].Menu := 'C Cell    ';
   MenuList[3].Num  := 3;
   MenuList[4].Menu := 'S Sort    ';
   MenuList[4].Num  := 4;
   MenuList[5].Menu := 'H Help    ';
   MenuList[5].Num  := 5;
   MenuList[6].Menu := 'Q Quit    ';
   MenuList[6].Num  := 6;
{$ENDIF}

   hMenu :=  SelectVertMenu(1, h_MENULINE, MenuList, 6);   { CVSCRN }
   If hMenu = 1 Then
      FileMenu ;            { CVFILE  }

   If hMenu = 2 Then
      TransMenu ;           { CVTRANS }

   If hMenu = 3 Then
      CellMenu ;            { CVCELL  }

   If hMenu = 4 Then
      SortCellData;         { CVSORT  }

   If hMenu = 5 Then
      DisplayHelp;          { CVHELP  }

   ClearMenuLine ;
   If hMenu = 6 Then
      GetMenuNum := STOPKEY
   Else
      GetMenuNum := #0 ;
End ;

{ ---------------------------------------------- }
{  Initalize                                     }
{ ---------------------------------------------- }
Procedure Initialize;

Var 
   i,j:  Integer ;

Begin
   ClrScr;
   GVideoMode;                         { CVCRT  }
   CursorOff ;

   //gcLineAttr := White ;
   gcLineAttr := Yellow ;
   gcCellAttr := White And $77 ;

   For i:=1 To h_MAXCOL Do
      Begin
         gcCellWidth[i] := h_INIT_CELLSIZE;
         gcCellDecPoint[i] := 0;

         For j:=1 To h_MAXROW Do
            Begin
               Sheet[i,j].tpMain := Nil;
               Sheet[i,j].cCellType := 0;
               Sheet[i,j].cCellColor := gcCellAttr;
            End ;
      End;

   ghMaxX := 1 ;
   ghMaxY := 1 ;
   ghTopCol := 1 ;
   ghTopRow := 1 ;
   ghEndCol := 1 ;
   ghEndRow := 1 ;
   gcMode := c_NORMAL ;
   ghColNameLine := h_COLNAMELINE;
   ghLeftSide := h_LEFTSIDE;
   gcForm := c_NORMAL ;
   gsFileName := '' ;

   ClearStack ;         { CVUNDO }

   //   Screen_Initialize; 
End;

(***********************************)
(* M a i n   R o u t i n e         *)
(***********************************)

Var 
   sData     :  String ;
   cEdit     :  Byte ;
   cAscii, cCode :  Char;
   sFileName   :  String ;
Begin
   Initialize;
   If Paramcount >= 1 Then
      Begin
         sFileName := ParamStr(1) ;
         If (LoadTableMain(sFileName, False) = False) Then
            Begin
               ClrScr ;
               exit ;
            End ;
      End ;

   Screen_Initialize;          { CVSCRN }
   Repeat
      If gcMode = c_NORMAL Then
         Begin
            ghTopCol := ghX ;
            ghTopRow := ghY ;
            ghEndCol := ghX ;
            ghEndRow := ghY ;
         End ;

      SetCurrentDataOnCell(Sheet[ghX,ghY], (Reverse*gcCellAttr And Not Blink));    { CVSCRN }
      SetCurrentDataOnLine(Sheet[ghX,ghY]);                   { CVSCRN }

   (***********************************)
   (* Get Data form Screen   CVINPT   *)
   (***********************************)
      If Sheet[ghX,ghY].tpMain<>Nil Then
         sData := Sheet[ghX,ghY].tpMain^
      Else
         sData := '';

      cEdit := c_INPUT ;
      CursorOn ;

      GetInputData(sData, cAscii, cEdit);    { CVINPT }
      CursorOff ;

   (***********************************)
   (* Get Data form Screen   CVDATA   *)
   (***********************************)
      CheckData(sData, Sheet[ghX,ghY]);    { CVDATA }

   (***********************************)
   (* Check Keycode          CVSCRN   *)
   (***********************************)
      Case cAscii Of 

   (***********************************)
   (* Mark Block                      *)
   (***********************************)
         ESCKEY     : { Cancel Block }
                       gcMode := c_NORMAL ;

         CTRL_B     : { Start Block }
                       Begin
                          gcMode := c_MARK ;
                          ghTopCol := ghX ;
                          ghTopRow := ghY ;
                       End ;

         CTRL_K   :
                     Begin
                        GetKey(cAscii, cCode) ;
                        If cAscii = #0 Then
                           cAscii := cCode ;

                        If cAscii = CTRL_K Then      { End Block }
                           Begin
                              gcMode := c_MARK_END ;
                              ghEndCol := ghX ;
                              ghEndRow := ghY ;
                           End ;

                        If cAscii = CTRL_V Then      { Move Block  CVCELL }
                           Begin
                              MoveCellData(ghTopCol, ghTopRow,
                                           ghEndCol, ghEndRow,
                                           ghX, ghY);
                           End ;

                        If cAscii = CTRL_C Then      { Copy Block  CVCELL }
                           Begin
                              CopyCellData(ghTopCol, ghTopRow,
                                           ghEndCol, ghEndRow,
                                           ghX, ghY);
                           End ;
                     End ;

   (***********************************)
   (* Undo                    CVUNDO  *)
   (***********************************)
         CTRL_X   :
                     Begin
                        GetKey(cAscii, cCode) ;
                        If cAscii = #0 Then
                           cAscii := cCode ;

                        If (cAscii = CTRL_U) Or (cAscii = #117) Then
                           Begin
                              Undo ;
                              SetScreen;
                           End ;
                     End;

   (***********************************)
   (* Move Cell Pointer       CVSCRN  *)
   (***********************************)
         UPKEY:  CursorUp;
         DOWNKEY:  CursorDown;
         RIGHTKEY:  CursorRight;
         LEFTKEY:  CursorLeft;

         CTRL_C:  PageDown;
         CTRL_R:  PageUP;

         CTRL_RIGHTKEY:  PageLeft;
         CTRL_LEFTKEY:  PageRight;

         HOMEKEY:  Screen_Initialize;

         F5KEY:  cAscii := GetMenuNum ;

   (***********************************)
   (* Insert, Delete Row       CVNEWCL*)
   (***********************************)
         CTRL_N:  InsertNewRow(ghY);

         CTRL_Y  :
                    Begin
                       If gcMode = c_NORMAL Then          { CVNEWCL }
                          DeleteCurrentRow(ghY);

                       If gcMode = c_MARK Then            { CVCELL }
                          EraseCellData(ghTopCol, ghTopRow, ghX, ghY);

                       gcMode := c_NORMAL ;
                    End ;

         DELKEY:
        {EraseCellData(ghX, ghY, ghX, ghY);}
                  EraseCellData(ghTopCol, ghTopRow, ghEndCol, ghEndRow);
    { CVCELL }

   (***********************************)
   (* Insert, Delete Col       CVNEWCL*)
   (***********************************)
         ALT_NKEY:  InsertNewCol(ghX);
         ALT_YKEY:  DeleteCurrentCol(ghX);

   (***********************************)
   (* ColWidth                 CVCELL *)
   (***********************************)
         ALT_LKEY, F4KEY  :
                             ChangeWidthByValue(gcCellWidth[ghX]+1) ;

         ALT_SKEY, F3KEY  :
                             ChangeWidthByValue(gcCellWidth[ghX]-1) ;

         CTRL_A :
                   Begin
                      AdjustWidth(ghX) ;
                      SetScreen;
                      GetKey(cAscii, cCode) ;
                      If cAscii = #0 Then
                         cAscii := cCode ;

                      If cAscii = CTRL_A Then
                         AdjustAllWidth ;
                   End ;

   (***********************************)
   (* GotoCell                 CVNEWCL*)
   (***********************************)
         CTRL_Q   :
                     Begin
                        GetKey(cAscii, cCode) ;
                        If cAscii = #0 Then
                           cAscii := cCode ;

                        If cAscii = CTRL_C Then
                           GotoCell(ghX, ghMaxY) ;
                        If cAscii = CTRL_R Then
                           GotoCell(ghX, 1) ;
                     End ;

   (***********************************)
   (* Change Form Mode         CVNEWCL*)
   (***********************************)
         CTRL_F:  ChangeFormMode;

   (***********************************)
   (* Display Help             CV#HELP*)
   (***********************************)
         F1KEY:  DisplayHelp;

      End;
   Until cAscii = STOPKEY;
   ClearAllCells ;    { CVFILE }
   ClearStack ;         { CVUNDO }

   CursorOn ;
   ClrScr ;
End.
