{$MODE TP}

Unit cvundo ;

Interface

Uses cvdef, cvcrt, cveval ;
Procedure ClearStack ;
Procedure UndoLog(hSeq, hx, hy : Integer);
Procedure Undo ;

{ ---------------------------------------------- }

Implementation

Uses cvnewcl ;

Const 
   h_MAXSTACK =  100 ;
   h_UNDO_MODE =  -99 ;

Type 
   tpStack =  ^tpLogData ;
   tpLogData =  Record
      hSeq    :  Integer ;
      hx     :  Integer ;
      hy     :  Integer ;
      cCellType :  byte ;
      cCellColor :  byte ;
      sCellData :  String ;
      tpNext   :  tpStack ;
   End;

Var 
   ghStackLevel :  Integer;
   gtpTop:  tpStack;
   ghUndo:  Integer ;

Procedure PrintStack(tpTop : tpStack; hStackLevel:Integer) ;
Begin
   Clreol ;
   writeln('Stacklevel:', hStackLevel, ' hSeq:',tpTop^.hSeq, ' hX:',tpTop^.hX, ' hY:',tpTop^.hY,
           ' cCellType:',tpTop^.cCellType,' cCellColor:' ,tpTop^.cCellColor,' sCellData:',tpTop^.
           sCellData) ;
End ;

Procedure PrintAllStack ;

Var 
   tpTop    :  tpStack;
   hStackLevel :  Integer ;
Begin
   tpTop := gtpTop ;
   hStackLevel := ghStackLevel;
   writeln('-- PrintAllStack ------------------------------------') ;
   While (tpTop <> Nil) Do
      Begin
         PrintStack(tpTop, hStackLevel) ;
         tpTop := tpTop^.tpNext;
         hStackLevel := hStackLevel -1 ;
      End ;
End ;

{ ---------------------------------------------- }
{  Delete bottom entry in stack                  }
{ ---------------------------------------------- }
Procedure DeleteLastStack ;

Var 
   temp:  tpStack ;
   temp2:  tpStack ;
Begin
   temp := Nil ;
   While (gtpTop <> Nil) Do
      Begin
         temp2 := temp;
         temp := gtpTop;
         gtpTop := gtpTop^.tpNext;
      End ;
   If temp <> Nil Then
      Begin
         dispose(temp) ;
         temp2^.tpNext := Nil ;
      End;
   ghStackLevel := ghStackLevel-1 ;
End ;

{ ---------------------------------------------- }
{  Push data into stack                          }
{ ---------------------------------------------- }
Procedure Push(tpLog: tpLogData) ;

Var 
   temp:  tpStack ;
Begin
   new(temp) ;
   temp^.hSeq := tpLog.hSeq ;
   temp^.hx := tpLog.hx ;
   temp^.hy := tpLog.hy ;
   temp^.cCellType := tpLog.cCellType;
   temp^.cCellColor := tpLog.cCellColor;
   temp^.sCellData := tpLog.sCellData;
   temp^.tpNext := gtpTop;
   gtpTop := temp ;
   ghStackLevel := ghStackLevel+1 ;
   If ghStackLevel > h_MAXSTACK Then
      DeleteLastStack ;
End ;

{ ---------------------------------------------- }
{  Pop data from                                 }
{ ---------------------------------------------- }
Procedure Pop(Var tpLog: tpLogData) ;

Var 
   temp:  tpStack ;
Begin

   tpLog.hSeq := gtpTop^.hSeq;
   tpLog.hx := gtpTop^.hx;
   tpLog.hy := gtpTop^.hy;
   tpLog.cCellType := gtpTop^.cCellType;
   tpLog.cCellColor := gtpTop^.cCellColor;
   tpLog.sCellData := gtpTop^.sCellData;
   temp := gtpTop;
   gtpTop := gtpTop^.tpNext ;
   dispose(temp) ;
   ghStackLevel := ghStackLevel-1 ;
End ;

{ ---------------------------------------------- }
{  Clear stack                                   }
{ ---------------------------------------------- }
Procedure ClearStack ;

Var 
   temp:  tpStack ;
Begin
   temp := gtpTop ;
   While (temp <> Nil) Do
      Begin
         gtpTop := gtpTop^.tpNext;
         dispose(temp);
         temp := gtpTop;
      End ;
   ghStackLevel := 0 ;
   gtpTop := Nil ;
   ghUndo := 0 ;
End ;

{ ---------------------------------------------- }
{  Write undo log                                }
{ ---------------------------------------------- }
Procedure UndoLog(hSeq, hx, hy : Integer);

Var 
   tpLog:  tpLogData ;

Begin
   If ghUndo = h_UNDO_MODE Then
      exit ;

   tpLog.hSeq := hSeq ;
   tpLog.hx := hx ;
   tpLog.hy := hy ;

   If hSeq >=0 Then
      Begin
         tpLog.cCellType := Sheet[hx,hy].cCellType ;
         tpLog.cCellColor := Sheet[hx,hy].cCellColor ;
         If Sheet[hx,hy].tpMain = Nil Then
            tpLog.sCellData := ''
         Else
            tpLog.sCellData := Sheet[hx,hy].tpMain^ ;
      End
   Else
      Begin
         tpLog.cCellType := 0 ;
         tpLog.cCellColor := 0 ;
         tpLog.sCellData := ''
      End ;

   Push(tpLog) ;
   //gotoxy(4,15) ;
   //PrintAllStack ;
End ;

{ ---------------------------------------------- }
{  Backout Line                                  }
{ ---------------------------------------------- }
Procedure BackoutLine(Var tpLog: tpLogData) ;
Begin
   Case tpLog.hSeq Of 
      h_DELROW :   InsertNewRow(tpLog.hy);
      h_DELCOL :   InsertNewCol(tpLog.hx) ;
      h_INSROW :
                  Begin
                     DeleteCurrentRow(tpLog.hy);
                     tpLog.hSeq := 0 ;
                  End;
      h_INSCOL :
                  Begin
                     DeleteCurrentCol(tpLog.hx) ;
                     tpLog.hSeq := 0 ;
                  End ;
   End;
End ;

{ ---------------------------------------------- }
{  Backout cell data                             }
{ ---------------------------------------------- }
Procedure BackoutCell(tpLog : tpLogData) ;
Begin
   //writeln('tpLog.sCellData:', tpLog.sCellData) ;
   Sheet[tpLog.hx,tpLog.hy].cCellType := tpLog.cCellType ;
   Sheet[tpLog.hx,tpLog.hy].cCellColor := tpLog.cCellColor ;
   If Sheet[tpLog.hx,tpLog.hy].tpMain <> Nil Then
      Begin
         If tpLog.sCellData <> '' Then
            Sheet[tpLog.hx,tpLog.hy].tpMain^ := tpLog.sCellData
         Else
            FreeCellArea(Sheet[tpLog.hx,tpLog.hy]) ;   { CVEVAL }
      End
   Else
      Begin
         If tpLog.sCellData <> '' Then
            Begin
               GetCellArea(Sheet[tpLog.hx,tpLog.hy]);
               Sheet[tpLog.hx,tpLog.hy].tpMain^ := tpLog.sCellData ;
            End;
      End;
End ;

{ ---------------------------------------------- }
{  Undo                                          }
{ ---------------------------------------------- }
Procedure Undo ;

Var 
   tpLog :  tpLogData ;
Begin
   //gotoxy(4,15) ;
   //PrintAllStack ;
   ghUndo := h_UNDO_MODE ;
   While (gtpTop <> Nil) Do
      Begin
         Pop(tpLog) ;
         If tpLog.hSeq >= 0 Then
            BackoutCell(tpLog)
         Else
            BackoutLine(tpLog) ;

         If tpLog.hSeq = 0 Then
            Begin
               ghUndo := 0 ;
               exit ;
            End;
      End;
   ghUndo := 0 ;
End ;

End .
