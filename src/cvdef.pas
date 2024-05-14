{$MODE TP}

Unit cvdef ;

Interface
{* =========================================================== *}
{*     Constant    Data                                        *}
{* =========================================================== *}

Const 
   YES_No =  1;
   NO_No =  2;
   CANCEL_No =  3;

   EofLine  =  ^M;
   Numbers:  set Of Char =  ['0'..'9'];
   HexDigit:  set Of Char =  ['0'..'9','a'..'f','A'..'F'];

(* Ascii code and scan code *)
   //ALT_QKEY = #16;
   //ALT_WKEY = #17;
   //ALT_EKEY = #18;
   //ALT_RKEY = #19;
   //ALT_TKEY = #20;
   ALT_YKEY =  #21;
   //ALT_UKEY = #22;
   ALT_IKEY =  #23;
   //ALT_OKEY = #24;
   //ALT_PKEY = #25;
   //ALT_AKEY = #30;
   ALT_SKEY =  #31;
   //ALT_DKEY = #32;
   //ALT_FKEY = #33;
   //ALT_GKEY = #34;
   //ALT_HKEY = #35;
   //ALT_JKEY = #36;
   //ALT_KKEY = #37;
   ALT_LKEY =  #38;
   //ALT_ZKEY = #44;
   //ALT_XKEY = #45;
   //ALT_CKEY = #46;
   //ALT_VKEY = #47;
   //ALT_BKEY = #48;
   ALT_NKEY =  #49;
   //ALT_MKEY = #50;

   ALT_LEFTKEY =  #157;
   ALT_RIGHTKEY =  #155;

   F1KEY =  #59;
   F2KEY =  #60;
   F3KEY =  #61;
   F4KEY =  #62;
   F5KEY =  #63;
   F6KEY =  #64;
   //F7KEY = #65;
   //F8KEY = #66;
   //F9KEY = #67;
   //F10KEY = #68;
   //F11KEY = #133;
   //F12KEY = #134;

   HOMEKEY =  #71;
   UPKEY =  #72;
   PGUPKEY =  #73;
   LEFTKEY =  #75;
   RIGHTKEY =  #77;
   ENDKEY =  #79;
   DOWNKEY =  #80;
   PGDNKEY =  #81;
   INSKEY =  #82;
   DELKEY =  #83;
   BSKEY =  #8;
   TABKEY =  #9;
   SPACEKEY =  #32;
   CRKEY =  #13;
   ESCKEY =  #27;

   STOPKEY =  #17 ;
   PERIOD =  #46 ;
   PLUSKEY =  #43 ;
   SLASHKEY =  #47 ;

   CTRL_A    =  #01;
   CTRL_B    =  #02;
   CTRL_C    =  #03;
   //CTRL_E    = #05;
   CTRL_F    =  #06;
   CTRL_K    =  #11;
   //CTRL_L    = #12 ;
   CTRL_N    =  #14;
   //CTRL_P    = #16 ;
   CTRL_Q    =  #17 ;
   CTRL_R    =  #18 ;
   //CTRL_T    = #20 ;
   CTRL_U    =  #21 ;
   CTRL_V    =  #22 ;
   CTRL_X    =  #24 ;
   CTRL_Y    =  #25 ;
{$IFDEF WIN}
   CTRL_RIGHTKEY =  #115 ;
   CTRL_LEFTKEY =  #116 ;
{$ELSE NOWIN}
   CTRL_RIGHTKEY =  #68 ;
   CTRL_LEFTKEY =  #67 ;
{$ENDIF WIN}

   MaxMenu =  20;

(* Color *)
   Black =  0;
   Blue =  1;
   Green =  2;
   Cyan =  3;
   Red =  4;
   Magenta =  5;
   Brown =  6;
   LightGray =  7;
   DarkGray =  8;
   LightBlue =  9;
   LigtnGreen =  10;
   LightCyan =  11;
   LightRed =  12;
   LightMagenta =  13;
   Yellow =  14;
   White =  15;

   Reverse =  16;
   Blink =  128 ;

{ =============================================== }
{    Constant Data                                }
{ =============================================== }
   h_MAXROW =  4096;
   h_MAXCOL =  52;
   h_LEFTSIDE =  5;

   h_INIT_CELLSIZE =  10;
   h_MAXSTRING =  255;
   h_MAXCELLSIZE =  255;
   h_MAXDECPOINT =  7 ;

   h_INPUTLINE =  1;
   h_MENULINE =  1;
   h_GUIDELINE =  2;

   h_COLNAMELINE =  3;

   h_CELL_NUMERIC =  $80 ;
   h_CELL_FORMULA =  $40 ;
   h_CELL_RECALC  =  $20 ;

   c_INPUT =  0 ;
   c_EDIT =  1 ;
   c_CALC =  2 ;

   c_NORMAL =  0 ;
   c_MARK =  1 ;
   c_MARK_END =  2 ;

   c_NOGUIDE =  1 ;

   c_MOVECELL =  0 ;
   c_COPYCELL =  1 ;
   c_CHANGE_X =  1 ;
   c_CHANGE_Y =  2 ;
   c_CHANGE_XY =  3 ;
   c_TEXTFILE =  0 ;
   c_CSVFILE =  1 ;


{
c_VER_KEISEN = $08 ;
c_HOR_KEISEN = $04 ;
c_BOX_TYPE = 1;
c_HOR_TYPE = 2;
c_VER_TYPE = 4;
h_MAX_ENTRY = 100 ;
}
   c_CELL_CENTER =  1;
   c_CELL_RIGHT =  2;
   c_CELL_LEFT =  3;

   s_CELL_LEFT =  '\''' ;
   s_CELL_RIGHT =  '\"' ;
   s_CELL_CENTER =  '\^' ;
   s_CELL_FILLCHAR =  '\' ;

{* =========================================================== *}
{*     for UTF-8, Color                                        *}
{* =========================================================== *}
   c_2BYTE =  $C0 ;
   c_3BYTE =  $E0 ;
   c_4BYTE =  $F0 ;
   AXIS_TEXTCOLOR =  Black ;
   AXIS_TEXTBACKGROUND =  Cyan ;

{* =========================================================== *}
{*     for Evaluation                                          *}
{* =========================================================== *}
   h_FORMULA_NUMERIC =  1 ;
   h_FORMULA_STRING =  2 ;

   h_ERR_DEFINE_FUNC =  20 ;
   h_ERR_NOT_NUMERIC =  21 ;
   h_ERR_NOT_EXPRESSION =  22 ;
   h_ERR_RANGE =  23 ;
   h_ERR_FIRST_CELL_OVERLAPED =  24 ;
   h_ERR_SECOND_CELL_OVERLAPED =  25 ;
   h_ERR_CELL_IVALID_ORDER =  26 ;
   h_ERR_H2I_CONVERSION =  27 ;
   h_ERR_ZERO_DIVIDE =  28 ;

   h_ERR_VLOOKUP_STRING =  40 ;
   h_ERR_VLOOKUP_KEY =  41;
   h_ERR_VLOOKUP_KEYVALUE =  42;
   h_ERR_VLOOKUP_KEYCELL =  43;
   h_ERR_VLOOKUP_KEYBLANK =  44;
   h_ERR_VLOOKUP_KEY_NOTNUMERIC =  45;
   h_ERR_VLOOKUP_NOTSEP =  46;
   h_ERR_VLOOKUP_RANGE =  47;
   h_ERR_VLOOKUP_COLNUM =  48 ;
   h_ERR_VLOOKUP_EVAL_RANGE =  49 ;
   h_ERR_VLOOKUP_NOTARGET =  50 ;

   h_ERR_ATTR_NOTMACH =  60 ;

   h_LOGICAL_TRUE =  1 ;
   h_LOGICAL_FALSE =  0 ;
   h_ERR_VIF_STRING =  70 ;
   h_ERR_VIF_NOTSEP =  71;
   h_ERR_VIF_LOGICAL_ERR =  72;
   h_ERR_VIF_TRUE_FIELD_ERR =  73;
   h_ERR_VIF_FALSE_FIELD_ERR =  74;

   h_ERR_NOT_COMPLETED =  99 ;

{* =========================================================== *}
{*     for Undo                                                *}
{* =========================================================== *}

   //h_MAXSTACK =  100 ;
   h_DELROW =  -1 ;
   h_DELCOL =  -2 ;
   h_INSROW =  -3 ;
   h_INSCOL =  -4 ;
   //h_UNDO_MODE = -99 ;

{* =========================================================== *}
{*     Type definitions                                        *}
{* =========================================================== *}

Type 

(* Definnition for Menu *)
   TypeMenu =  Record
      Menu :  String;
      Num  :  byte;
   End;

   Menus =  array[1..MaxMenu] Of TypeMenu;

(* Define Cell Area *)
   tpData  =  ^String ;
   tpCell   =  Record
      cCellType :  byte ;
      cCellColor :  byte ;
      tpMain :  tpData ;
   End;


{ =============================================== }
{    Variable Data                                }
{ =============================================== }

Var 
  { ----------------------------------------------- }
  {  Col and Row                                    }
  { ----------------------------------------------- }
   MaxRow, MaxCol :  integer;
   ghTopCol, ghEndCol, ghTopRow, ghEndRow:  Integer ;
   ghX, ghY :  Integer ;                             { Current Cell    }
   ghScrX, ghScrY :  Integer ;             { Cell Cursor Position }
   ghFirstX, ghFirstY :  Integer ;         { First Cell No. on Screen}
   ghMaxX, ghMaxY :  Integer ;             { Max Cell No. on Screen}
   ghLeftSide:  Integer;                  { X-Axis start line }
   ghColNameLine:  Integer;               { Y-Axis start Line }

  { ----------------------------------------------- }
  {  Cell Data                                      }
  { ----------------------------------------------- }
   Sheet:  array [1..h_MAXCOL,1..h_MAXROW] Of tpCell;  { Cell Data }
   gcCellWidth:  array [1..h_MAXCOL] Of Byte ;
   gcCellDecPoint:  array [1..h_MAXCOL] Of Byte ;
   gcCellPos:  array [1..h_MAXCOL] Of Byte ;
   ghColCount :  Integer ;                { Col Count on Screen  }
   gcCellAttr, gcLineAttr:  Byte ;    { Cell and line attribute }

   MenuList :  Menus;
   cMenuNum  :  Byte;

   gcMode:  Byte ;
   gcForm:  Byte ;

   gsFileName:  String ;

Function GetCellDecPoint(hx, hy : Integer):  Byte ;
Procedure SetCellDecPoint(hx, hy : Integer; decPoint: Byte) ;

{ ---------------------------------------------- }

Implementation

{ ---------------------------------------------- }
{  Setup decimal point                           }
{ ---------------------------------------------- }
Function GetCellDecPoint(hx, hy : Integer):  Byte ;
Begin
   GetCellDecPoint := Sheet[hx, hy].cCellType And $07 ;
End ;

Procedure SetCellDecPoint(hx, hy : Integer; decPoint: Byte) ;
Begin
   Sheet[hx, hy].cCellType := (Sheet[hx, hy].cCellType And $f8)+decPoint ;
End ;

End.
