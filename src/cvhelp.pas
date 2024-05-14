{$MODE TP}

Unit cvhelp ;

Interface

Uses cvdef, cvcrt, cvscrn ;
Procedure DisplayHelp ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Help Screen                                   }
{ ---------------------------------------------- }
Procedure DisplayHelp ;

Var 
   cAscii, cCode  :  Char;

Const 
   h_HELPLINES =  18 ;
   HelpListLines:  Array[1..h_HELPLINES] Of String[255] 
                   =  (
                       '1) Move ',
                       '  CELL         [Arrow key]  GoTop [CTRL+QC]  GoEnd [CTRL+QR]  Go"A1" [HOME]'
                       ,

                   '  Screen       Vertical [CTRL+C][CTRL+R]  Horizontal [CTRL+ArrowR][CTRL+ArrowL]'
                       ,
                       '2) ROW,COL ',
                       '  Ins/del ROW  Insert [CTRL+N]  Delete [CTRL+Y]',
                       '  Ins/del COL  Insert [ALT+N]  Delete [ALT+Y] ',

                   '  Width        Shrink [PF03]or[ALT+S]  Expand [PF04]or[ALT+L]  or ["/"->"Cell"]'
                       ,
                       '               Adjust current col [CTRL+A] ->[CTRL+A] All cols',
                       '3) Range ',
                       '  Set "Range"  [CTRL+B]->[CTRL+KK]  Reset[ESC] ',
                       '4) Edit ',
                       '  Go edit mode [PF02]or[ALT+I]',
                       '  Special key  Menu [/]or[PF05]  Formula [+]or[PF06]  Axis on/off [CTRL+F]',

                 '  First 2 chars   [\'']Left justify [\^]Centerize [\"]Right justify [\x]Fill "x" '
                       ,
                       '  Copy, Move   "Range" -> [CTRL+KC] Copy  [CTRL+KV] Move',
                       '  Attribute    "Range" -> ["/"->"Cell"]',
                       '  Undo         [CTRL+X]->[CTRL+U] or "u"',
                       ''
                      ) ;

Var 
   i  :  Integer ;
   Attr :  byte ;
Begin
   ClrScr;                             { CVCRT  }
   Attr := GetAttr ;
   //GotoXY(1,1) ;                       { CVCRT  }
   For i:=1 To h_HELPLINES Do
      Begin
         GotoXY(1,i) ;                       { CVCRT  }
         WriteString(HelpListLines[i], Attr) ;
      End ;
   GetKey(cAscii, cCode);
   If cAscii = #0 Then
      cAscii := cCode ;

   SetScreen ;              { CVSCRN }
End ;
End .
