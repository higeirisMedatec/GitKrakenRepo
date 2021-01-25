unit ExtractDatasMain;
{ ---------------------------------------------------------------------- }
{ This unit contains tools to extract some datas from a trend }
{ form type, TFormExtract. }
{ 0.00 VL 071220 initial coding }
{ ---------------------------------------------------------------------- }

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, inifiles, Vcl.StdCtrls, StrUtils,
  RegularExpressions;

{ ----------------------------------------------------------------------- }
{ preliminary declarations }
{ CONSTANTS }
// CONST
Const
  FILENAME_APNEA_EVENT = 'ApneaEvents.csv';
  FILENAME_SLEEP_STAGES = 'SleepStages.csv';

Const
  LONG_200HZ = 2000; { 10 " … 200 Hz }
  LONG_400HZ = 4000; { 10 " … 400 Hz }
  LONG_500HZ = 5000; { 10 " … 500 Hz }
  MAGIC100_12B = 204.8; { rapport entre Volts et points ad }
  MAGIC100_16B = 3267.8; { rapport entre Volts et points ad }
  N_CAN_ANAL = 128;

CONST
  PN_EVENT_TO_FIND: smallint = 11;

CONST
  xlWBatWorksheet = -4167;

CONST
  MAX_CAL_TAB = 128;
  FCALNAME = 'SCAL.DON';
  MaxFs = 20;

CONST
  use_comma: boolean = false;

  { Global Type }
  // TYPE
Type
  P_200HZ = ^P11;
  P11 = Array [0 .. LONG_200HZ - 1] Of smallint;
  P_400HZ = ^P411;
  P411 = Array [0 .. LONG_400HZ - 1] Of smallint;
  P_500HZ = ^P511;
  P511 = Array [0 .. LONG_500HZ - 1] Of smallint;
  PPTXT = String[8];

  { ----------------------------------------------------------------------- }
  { Global Variables }
Var
  P_CONTACTS: ARRAY [1 .. N_CAN_ANAL] OF PPTXT;
  SampleRate: smallint;

var
  XL: Variant;
  Sheet: Variant;

Type
  Cal_Record = Record
    Name: String[6];
    Volt0: Single;
    Volt1: Single;
    Uni: String[6];
    Min: Single;
    Max: Single;
    AC_DC: String[2];
    Zeroing: String[30];
    NrFS: Integer;
    FScales: Array [1 .. MaxFs] of Single;
  End;

  Grotabu = Array [0 .. 120000] of Single;

Var
  CAL_TAB: Array [1 .. MAX_CAL_TAB] Of Cal_Record;
  NB_ANALOG_CHANNELS: Byte;
  Tatabu: Grotabu;
  EnormTabu: Array [1 .. 64] Of Grotabu;
  IsDream: boolean;
  Nrofheaderblocks: Integer;

VAR
  RRfile: file of smallint;
  RRtime: file of Integer;
  HF_BTB: ARRAY [1 .. 1000] OF Byte;
  QRS_BTB: ARRAY [1 .. 1000] OF LONGINT;
  Tot_Qrs: LONGINT;

Var
  FILE_200HZ: File Of P11;

Var
  FILE_400HZ: File Of P411;

Var
  FILE_500HZ: File Of P511;

type
  TFormExtract = class(TForm)
    EditPath: TEdit;
    ButtonPath: TButton;
    FileOpenDialogPath: TFileOpenDialog;
    ButtonExtractSleep: TButton;
    ButtonExtractApnea: TButton;
    ListBoxLog: TListBox;
    LabelName: TLabel;
    EditName: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure ButtonPathClick(Sender: TObject);
    procedure ButtonExtractApneaClick(Sender: TObject);
    procedure ButtonExtractSleepClick(Sender: TObject);
    procedure EditPathKeyPress(Sender: TObject; var Key: Char);
    procedure EditPathChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormExtract: TFormExtract;

  { ----------------------------------------------------------------------- }
  { Public Functions }

implementation

{ ----------------------------------------------------------------------- }
{ units used }
uses EEGGLOB, sleepcst, cevtunit, evtunit;

{$R *.dfm}
procedure LoadPatient (var sPathPat: string);

type
  TrHeaderBlock = array [0 .. 3999] of Byte;
  TTrHeaderBlock = ^TrHeaderBlock;

var
  TndFile: File of TrHeaderBlock;
  Header: TTrHeaderBlock;
  i, j, jj: Integer;
  s, S1: string;
  F, DefFile: TextFile;
  StMulti, st1, st2: string;
  JJJJ: Integer;
  Add_Another_Item, Found: Boolean;

  {VLAG}
  tFileTrend : TStringlist;
  tChannelsName : TStringlist;

begin { loadpatient }

  Itemctr := 0;
  TagCtr := 0;
  LeftColCtr := 0;

  { load the available trends from file header }
  { store them into the AvailableTrends list }
  { and prepare the 'trend change' popup menu for if the user clicks on it }
  tChannelsName := Tstringlist.Create;   // VLAG : tString that contains channels name

  {$I-}
  If Force_Read_Only Then
    filemode := 0;
  ASSIGNFILE(TndFile, sPathPat + 'BTREND.DAT');
  RESET(TndFile);
  If Force_Read_Only Then
    filemode := 2;
  if ioresult = 0 then
    begin
      New(Header);
      Read(TndFile, Header^);
      CLOSEFILE(TndFile);
      NrBFicChannels := Header^[2];
      j := 3;
      For i := 1 to NrBFicChannels do
        begin
          s := '';
          while Chr(Header^[j]) <> ',' do
            begin
              s := s + upcase(Chr(Header^[j]));
              inc(j);
            end;
          inc(j);

          { add trend to list }
          tFileTrend.Add(s);
          { create menu item with trend name }

          If s = 'X' Then
            inc(TagCtr)
          Else
            Begin
              Add_Another_Item := TRue;
              For JJJJ := 0 To EXT_NB_EEG_CONTACTS_WITH_A - 1 Do
                If (Pos(EEG_CONTACTS[JJJJ], s) = 1) Then
                  BEGIN
                    { bingo, trend is something like 'FP1ALPHA' }
                    Found := false;
                    For jj := 0 To Itemctr - 1 DO
                      If tndnamelistbox.Items[jj].Caption = EEG_CONTACTS[JJJJ] Then
                        BEGIN
                          Found := TRue;
                          with tndnamelistbox.Items.insert(jj + 1) do
                            BEGIN
                              Caption := s; // Copy(S,Length(EEG_CONTACTS[JJJJ])+1,Length(S));
                              tag := TagCtr;
                              level := 2;
                              Indent := 10;
                            END;
                          inc(Itemctr);
                          inc(TagCtr);
                        END;
                    If (Found = false) Then
                      BEGIN
                        with tndnamelistbox.Items.Add do
                          BEGIN
                            Caption := EEG_CONTACTS[JJJJ];
                            tag := MAX_NR_TRENDS;
                            level := 1;
                            captionfont.Style := [fsbold];
                          END;
                        inc(Itemctr);
                        with tndnamelistbox.Items.insert(Itemctr) do
                          BEGIN
                            Caption := s; // Copy(S,Length(EEG_CONTACTS[JJJJ])+1,Length(S));
                            tag := TagCtr;
                            level := 2;
                            Indent := 10;
                          END;
                        inc(Itemctr);
                        inc(TagCtr);
                      END;
                    Add_Another_Item := false;
                  END;

              If Add_Another_Item Then
                BEGIN
                  with tndnamelistbox.Items.Add do
                    BEGIN
                      Caption := s;
                      tag := TagCtr;
                      level := 1;
                    END;
                  inc(TagCtr);
                  inc(Itemctr);
                END;
            end;
        end;
      Dispose(Header);
    end
  else
    begin
      { no trends for this patient, but maybe there is an hypno, or mark }
      { or some event trend - take eeg recording length as standard }
      NrBFicChannels := 0;
    end;
  {$I-}
  ASSIGNFILE(F, SIGPATH + 'Multiscore.don');
  RESET(F);
  TagCtr := Last_Spec_Tnd + 1;
  if (ioresult = 0) then
    begin
      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := 'AUTOMATIC';
          onclick := Nil;
          tag := MAX_NR_TRENDS;
          level := 1;
          enabled := currentuser.Rights[AUTH_SAVE_CONFIG];
        END;

      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := 'AUTOMATIC HYPNO';
          tag := FindSpecTndNr('AHYPNO');
          level := 2;
          Indent := 10;
          enabled := currentuser.Rights[AUTH_SAVE_CONFIG];
        END;
      inc(Itemctr);
      inc(TagCtr);

      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := 'AUTOMATIC EVENT';
          tag := FindSpecTndNr('AEVENT');
          level := 2;
          Indent := 10;
          enabled := currentuser.Rights[AUTH_SAVE_CONFIG];
        END;
      inc(Itemctr);
      inc(TagCtr);

      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := 'AUTOMATIC NEUR';
          tag := FindSpecTndNr('ANEUR');
          level := 2;
          Indent := 10;
          enabled := currentuser.Rights[AUTH_SAVE_CONFIG];
        END;
      inc(Itemctr);
      inc(TagCtr);

      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := 'AUTOMATIC CARD';
          tag := FindSpecTndNr('ACARD');
          level := 2;
          Indent := 10;
          enabled := currentuser.Rights[AUTH_SAVE_CONFIG];
        END;
      inc(Itemctr);
      inc(TagCtr);

      for i := 1 To 10 DO
        begin
          Readln(F, S1);
          if (S1 <> '') And (i <> Formmultiscore.ScorerNumber) then
            Begin
              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')';
                  onclick := Nil;
                  tag := MAX_NR_TRENDS;
                  level := 1;
                END;

              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')' + ' HYPNO';
                  tag := FindSpecTndNr('HYPN' + IntTostr(i));
                  level := 2;
                  Indent := 10;
                END;

              inc(Itemctr);
              inc(TagCtr);

              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')' + ' EVENT';
                  tag := FindSpecTndNr('EVENT' + IntTostr(i));
                  level := 2;
                  Indent := 10;
                END;
              inc(Itemctr);
              inc(TagCtr);

              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')' + ' NEUR';
                  tag := FindSpecTndNr('NEUR' + IntTostr(i));
                  level := 2;
                  Indent := 10;
                END;
              inc(Itemctr);
              inc(TagCtr);

              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')' + ' CARD';
                  tag := FindSpecTndNr('CARD' + IntTostr(i));
                  level := 2;
                  Indent := 10;
                END;
              inc(Itemctr);
              inc(TagCtr);

              { 3.9181 }
              with tndnamelistbox.Items.Add do
                BEGIN
                  Caption := S1 + '(' + IntTostr(i) + ')' + ' AA';
                  tag := FindSpecTndNr('AA' + IntTostr(i));
                  level := 2;
                  Indent := 10;
                END;
              inc(Itemctr);
              inc(TagCtr);
            end;
        end;
      CLOSEFILE(F);
    end;
  {$I+}
  ioresult;

  For i := Last_Spec_Tnd to -1 do
    begin
      { add the special trends in the same way }
      with tndnamelistbox.Items.Add do
        BEGIN
          Caption := SpecTndNames[i];
          tag := i;
          level := 1;
        END;
    end;

  tndnamelistbox.Items.Sort;
  tndnamelistbox.Items.Endupdate;

  { whatever happens, always have an Exit item in the left column }
  { of the 'change trend' menu }
  // TndItems[ItemCtr] := TMenuItem.Create(FormTrends) ;
  // TndItems[ItemCtr].Caption := 'E&xit' ; ;
  // TndItems[ItemCtr].OnClick := TtndItemClick ;
  // TndItems[ItemCtr].tag := TagCtr ;
  // TndListPopUpMenu.Items.Add(TndItems[ItemCtr]);

  { the menu structure of TndListPopUpMenu is as follows : }
  { N = nr of trends <=24 : tndlistpopupmenu contains TndItems 0..N, element N='Exit' }
  { N=25 : TndListPopUpMenu contains : TndItems 0..24; TndItems 100 (MORE); TndItems 25 (Exit) }
  { N>25 : TndListPopupMenu contains TndItems 0..24; TndItems 100; TndItems N (Exit) }
  { TndItems100 contains TndItems 25..N-1 (second part of list) }
  {$I-}
  ASSIGNFILE(DefFile, TrendPath + 'HLONLOFF.TND');
  RESET(DefFile);
  Readln(DefFile, H_Loff);
  Readln(DefFile, H_Lon);
  CLOSEFILE(DefFile);
  {$I+}
  if ioresult <> 0 then
    begin
      H_Loff := '00:00:00';
      TimeToString(Recording_Length + 5, 10, '00:00', st1, st2, false);
      { +5 : round to next min }
      H_Lon := st2;
    end;

  StringToTab(H_Loff, 10, '00:00', PT_H_LOFF);
  StringToTab(H_Lon, 10, '00:00', Pt_H_Lon);
  Ind1 := 10 * PT_H_LOFF;
  Ind2 := 10 * Pt_H_Lon;
  { nieuwe release }
  // If recording_Length>(6*60*24) Then
  // BEGIN
  // PT_H_LON:=PT_H_LON+(Recording_Length div 6 div 60 div 24)*360*24;
  // Ind2:=Ind2+(Recording_Length div 6 div 60 div 24)*3600*24;
  // END;

  StMulti := '';
  if Formmultiscore <> Nil then
    if Formmultiscore.EnableMultiscore then
      StMulti := ' - Current Scorer : ' + Uppercase(Formmultiscore.ScorerNames[Formmultiscore.ScorerNumber]) + ' (' + IntTostr
        (Formmultiscore.ScorerNumber) + ')';

  { titre au dessus de la form des trends }
  {$IFDEF F}
  Caption := 'Trends de ' + GlobPatient.Name + ',' + GlobPatient.firstname + ' - ' + GlobPatient.date + StMulti;
  {$ENDIF}
  {$IFDEF N}
  Caption := 'Trends van ' + GlobPatient.Name + ',' + GlobPatient.firstname + ' - ' + GlobPatient.date + StMulti;
  {$ENDIF}
  {$IFDEF E}
  Caption := 'Trends of ' + GlobPatient.Name + ',' + GlobPatient.firstname + ' - ' + GlobPatient.date + StMulti;
  {$ENDIF}
  {$IFDEF G}
  Caption := 'Trends von ' + GlobPatient.Name + ',' + GlobPatient.firstname + ' - ' + GlobPatient.date + StMulti;
  {$ENDIF}
  Capbase := Caption;
end; { loadpatient }

procedure PointComma(var s: string);

begin
  if not use_comma then
    exit;
  s := stringreplace(s, '.', ',', []);
end;

{ ----------------------------------------------------------------------- }
{ Functions of the unit }

FUNCTION DisplayPatientInfo(sPathPatient: string): boolean;
BEGIN
  // Add '\' at the end of the string if is not
  if AnsiRightStr(sPathPatient, 1) <> '\' then
  BEGIN
    sPathPatient := sPathPatient + '\';
  END;

  // Get the patient infos
  Retrieve_The_Patient_Data(sPathPatient + '\');
  FormExtract.EditName.Text := Patient.Firstname + ' ' + Patient.Name;
END;

Function CountBradies(p1, m1, p2, m2: Integer): Integer;
var
  i, cur: Integer;
  start, stop: Integer;
begin
  result := 0;
  start := p1 * 60 * 25 + m1; // 25Hz index
  stop := p2 * 60 * 25 + m2;
  for i := 1 to no_event_card - 1 do
  begin
    cur := 60 * 25 * Etable_Card^[i].Point + Etable_Card^[i].Mem div 8;
    // cardio : mem is at 200Hz, pneumo at 25
    if (cur >= start) and (cur <= stop) then
    begin
      FormExtract.ListBoxLog.items.Add('brady ' + inttostr(i));
      result := result + 1;
    end;
  end;

end;

Function CountAcen(p1, m1, p2, m2: Integer): Integer;
var
  i, cur: Integer;
  start, stop: Integer;

begin
  result := 0;
  start := p1 * 60 * 25 + m1; // 25Hz index
  stop := p2 * 60 * 25 + m2;
  for i := 1 to no_event_pneumo - 1 do
    if Etable^[i].Apn = E_Acen then
    begin
      cur := 60 * 25 * Etable^[i].Point1 + Etable^[i].Mem1; // pneumo at 25 Hz
      if (cur >= start) and (cur <= stop) then
      begin
        FormExtract.ListBoxLog.items.Add('Acen ' + inttostr(i));
        result := result + 1;
      end
      else
      begin
        cur := 60 * 25 * Etable^[i].Point2 + Etable^[i].Mem2; // pneumo at 25 Hz
        if (cur >= start) and (cur <= stop) then
        begin
          FormExtract.ListBoxLog.items.Add('Acen ' + inttostr(i));
          result := result + 1;
        end
      end;
    end;

end;

Function CountAobs(p1, m1, p2, m2: Integer): Integer;
var
  i, cur: Integer;
  start, stop: Integer;

begin
  result := 0;
  start := p1 * 60 * 25 + m1; // 25Hz index
  stop := p2 * 60 * 25 + m2;
  for i := 1 to no_event_pneumo - 1 do
    if Etable^[i].Apn = E_Aobs then
    begin
      cur := 60 * 25 * Etable^[i].Point1 + Etable^[i].Mem1; // pneumo at 25 Hz
      if (cur >= start) and (cur <= stop) then
      begin
        FormExtract.ListBoxLog.items.Add('Aobs ' + inttostr(i));
        result := result + 1;
      end
      else
      begin
        cur := 60 * 25 * Etable^[i].Point2 + Etable^[i].Mem2; // pneumo at 25 Hz
        if (cur >= start) and (cur <= stop) then
        begin
          FormExtract.ListBoxLog.items.Add('Aobs ' + inttostr(i));
          result := result + 1;
        end
      end;
    end;

end;

Function CountHpop(p1, m1, p2, m2: Integer): Integer;
var
  i, cur: Integer;
  start, stop: Integer;

begin
  result := 0;
  start := p1 * 60 * 25 + m1; // 25Hz index
  stop := p2 * 60 * 25 + m2;
  for i := 1 to no_event_pneumo - 1 do
    if (Etable^[i].Apn = E_Hypop) or (Etable^[i].Apn = E_Hobs) or
      (Etable^[i].Apn = E_Hcen) then
    begin
      cur := 60 * 25 * Etable^[i].Point1 + Etable^[i].Mem1; // pneumo at 25 Hz
      if (cur >= start) and (cur <= stop) then
      begin
        FormExtract.ListBoxLog.items.Add('hpop ' + inttostr(i));
        result := result + 1;
      end
      else
      begin
        cur := 60 * 25 * Etable^[i].Point2 + Etable^[i].Mem2; // pneumo at 25 Hz
        if (cur >= start) and (cur <= stop) then
        begin
          FormExtract.ListBoxLog.items.Add('hpop ' + inttostr(i));
          result := result + 1;
        end
      end;
    end;

end;

Function CountArsl(p1, m1, p2, m2: Integer): Integer;
var
  i, cur: Integer;
  start, stop: Integer;

begin
  result := 0;
  start := p1 * 60 * 25 + m1; // 25Hz index
  stop := p2 * 60 * 25 + m2;
  for i := 1 to no_event_neuro - 1 do
    if (Etable_neuro^[i].Typ = NE_Arsl) or (Etable_neuro^[i].Typ = NE_ArApn) or
      (Etable_neuro^[i].Typ = NE_ArMov) then
    begin
      cur := 60 * 25 * Etable_neuro^[i].Point1 + Etable_neuro^[i].Mem1;
      // pneumo at 25 Hz
      if (cur >= start) and (cur <= stop) then
      begin
        FormExtract.ListBoxLog.items.Add('arsl ' + inttostr(i));
        result := result + 1;
      end
      else
      begin
        cur := 60 * 25 * Etable_neuro^[i].Point2 + Etable_neuro^[i].Mem2;
        // pneumo at 25 Hz
        if (cur >= start) and (cur <= stop) then
        begin
          FormExtract.ListBoxLog.items.Add('arsl ' + inttostr(i));
          result := result + 1;
        end
      end;
    end;

end;

Function FindStageBefore(p1, m1: Integer): Integer;

var
  i: Integer;

begin
  result := N_U;
  i := p1 * 12 + m1 div 125 - 3; // 3 x 5s = 15s before
  if i < 0 then
    i := 0;
  result := hypnogram^[i];
end;

Function Get_Mean_Value(Ten_SEC_Page: LONGINT; Sec1, Sec2: LONGINT;
  Name: String; Var Min, Mean: Single): boolean;

Var
  Val: Single;
  i, Physcanal: Integer;
  Tab: P11;
  Tab4: P411;
  Tab5: P511;
  Sum, Sumabs: Double;
  Count: LONGINT;
  IndexCan: LONGINT;
  Size: LONGINT;
  Countabu: LONGINT;
  sPathPat: string;
Begin
  // Get the folder of the patient
  sPathPat := FormExtract.EditPath.Text;
  result := true;
  Countabu := 0;
  case SampleRate of
    2:
      begin
        AssignFile(FILE_200HZ, sPathPat + 'FIC1.DAT');
{$I-}
        Reset(FILE_200HZ);
{$I+}
      end;
    4:
      begin
        AssignFile(FILE_400HZ, sPathPat + 'FIC1.DAT');
{$I-}
        Reset(FILE_400HZ);
{$I+}
      end;
    5:
      begin
        AssignFile(FILE_500HZ, sPathPat + 'FIC1.DAT');
{$I-}
        Reset(FILE_500HZ);
{$I+}
      end;
  end;
  If Ioresult <> 0 Then
  BEGIN
    MessageDlg('File ' + sPathPat + 'FIC1.DAT not found', mtwarning, [mbok], 0);
    result := false;
    exit;
  END;

  Physcanal := Get_Index_Cal(Name);

  IndexCan := -1;

  For i := 1 To NB_ANALOG_CHANNELS DO
    IF (P_CONTACTS[i] = Name) Then
      IndexCan := i;

  if IndexCan = -1 then
  begin
    MessageDlg('Channel ' + name + ' not present. Not computed ?', mterror,
      [mbok], 0);
    result := false;
    exit;
  end;

  IndexCan := IndexCan + (NB_ANALOG_CHANNELS * Ten_SEC_Page);
  if Nrofheaderblocks = 2 then
    inc(IndexCan);

{$I-}
  case SampleRate of
    2:
      begin
        Size := Filesize(FILE_200HZ);
        Seek(FILE_200HZ, IndexCan);
        Read(FILE_200HZ, Tab);
        if Ioresult = 0 then
        begin
          If Sec2 < 10 Then
          BEGIN
            For i := Sec1 * 200 to Sec2 * 200 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab[i], Physcanal);
              inc(Countabu);
            END;
          END
          ELSE
          BEGIN
            For i := Sec1 * 200 to 1999 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab[i], Physcanal);
              inc(Countabu);
            END;
          END;
        end;

        if Sec2 >= 10 then
          Repeat
            Sec2 := Sec2 - 10;
            IndexCan := IndexCan + NB_ANALOG_CHANNELS;
            Seek(FILE_200HZ, IndexCan);
            Read(FILE_200HZ, Tab);
            if Ioresult = 0 then
            begin
              If Sec2 > 10 Then
                For i := 1 To LONG_200HZ - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab[i], Physcanal);
                  inc(Countabu);
                END
              Else
                For i := 1 to Sec2 * 200 - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab[i], Physcanal);
                  inc(Countabu);
                END;
            end; { ioresult }
          Until Sec2 <= 10;
        CloseFile(FILE_200HZ);
      end;

    4:
      begin
        Size := Filesize(FILE_400HZ);
        Seek(FILE_400HZ, IndexCan);
        Read(FILE_400HZ, Tab4);
        if Ioresult = 0 then
        begin
          If Sec2 < 10 Then
          BEGIN
            For i := Sec1 * 400 to Sec2 * 400 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab4[i], Physcanal);
              inc(Countabu);
            END;
          END
          ELSE
          BEGIN
            For i := Sec1 * 400 to 3999 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab4[i], Physcanal);
              inc(Countabu);
            END;
          END;
        end;

        if Sec2 >= 10 then
          Repeat
            Sec2 := Sec2 - 10;
            IndexCan := IndexCan + NB_ANALOG_CHANNELS;
            Seek(FILE_400HZ, IndexCan);
            Read(FILE_400HZ, Tab4);
            if Ioresult = 0 then
            begin
              If Sec2 > 10 Then
                For i := 1 To LONG_400HZ - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab4[i], Physcanal);
                  inc(Countabu);
                END
              Else
                For i := 1 to Sec2 * 400 - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab4[i], Physcanal);
                  inc(Countabu);
                END;
            end; { ioresult }
          Until Sec2 <= 10;
        CloseFile(FILE_400HZ);
      end;

    5:
      begin
        Size := Filesize(FILE_500HZ);
        Seek(FILE_500HZ, IndexCan);
        Read(FILE_500HZ, Tab5);
        if Ioresult = 0 then
        begin
          If Sec2 < 10 Then
          BEGIN
            For i := Sec1 * 500 to Sec2 * 500 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab5[i], Physcanal);
              inc(Countabu);
            END;
          END
          ELSE
          BEGIN
            For i := Sec1 * 500 to 4999 Do
            BEGIN
              Tatabu[Countabu] := ADTOUNITS(Tab5[i], Physcanal);
              inc(Countabu);
            END;
          END;
        end;

        if Sec2 >= 10 then
          Repeat
            Sec2 := Sec2 - 10;
            IndexCan := IndexCan + NB_ANALOG_CHANNELS;
            Seek(FILE_500HZ, IndexCan);
            Read(FILE_500HZ, Tab5);
            if Ioresult = 0 then
            begin
              If Sec2 > 10 Then
                For i := 1 To LONG_500HZ - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab5[i], Physcanal);
                  inc(Countabu);
                END
              Else
                For i := 1 to Sec2 * 500 - 1 do
                BEGIN
                  Tatabu[Countabu] := ADTOUNITS(Tab5[i], Physcanal);
                  inc(Countabu);
                END;
            end; { ioresult }
          Until Sec2 <= 10;
        CloseFile(FILE_500HZ);
      end;
  end; { case }

  Min := 1E6;
  Mean := 1E6;
  if Countabu < 100 then
    exit;

  Sum := 0;
  for i := 0 to Countabu - 1 do
  begin
    if Tatabu[i] < Min then
      Min := Tatabu[i];
    Sum := Sum + Tatabu[i];
  end;
  Mean := Sum / Countabu;

End;

{ ----------------------------------------------------------------------- }
/// <summary> Extract the apnea event of the patient </summary>
/// <remarks> Created by VLagarrigue </remarks>
procedure TFormExtract.ButtonExtractApneaClick(Sender: TObject);
var
  sPathPat: string;
  sPathTrend: string;
  FileApnea: TStringlist;
  iFileLoop: smallint;

begin
  // Get the folder of the patient
  sPathPat := FormExtract.EditPath.Text;
  // Get the trend path
  sPathTrend := sPathPat + 'Btrend.dat';

  // Add information in the listbox
  ListBoxLog.items.Add('Apnea events Extraction');
  ListBoxLog.items.Add('Load Events');

  {LoadPatient (trends.pas)}


  { Load event table (EEGGLOB.pas)}
  LoadEventTable;
  // Etable
  ListBoxLog.items.Add('DBG:no_event_pneumo : ' + IntToStr(no_event_pneumo));

  FileApnea := TStringlist.create;
  try
    for iFileLoop := 1 to no_event_pneumo do
    begin
      FileApnea.Add(inttostr(Etable^[iFileLoop].Apn));
    end;
    FileApnea.SaveToFile(sPathPat + FILENAME_APNEA_EVENT);
  finally
    FileApnea.Free
  end;
  { END Load envent table }
  ListBoxLog.items.Add('Apnea events extracted in ' + sPathPat + FILENAME_APNEA_EVENT);
end;

{ ----------------------------------------------------------------------- }
/// <summary> Extract the sleep stage of the patient </summary>
/// <remarks> Created by VLagarrigue
/// Code from export_excell_toux_drmarchand_uzbrussel_2015
/// </remarks>
procedure TFormExtract.ButtonExtractSleepClick(Sender: TObject);
var
  sPathPat: string;
  FileSleep: TStringlist;

begin
  // Get the folder of the patient
  sPathPat := FormExtract.EditPath.Text;

  // Add something in the listbox
  ListBoxLog.items.Add('Sleep Extraction');
  ListBoxLog.items.Add('Load trend patient');

  LoadPatient (sPathPat);
  ListBoxLog.items.Add('File loaded');
  FileSleep.Create;
  try
    FileSleep.Add('FirstLine');
//    for iFileLoop := 1 to no_event_pneumo do
//    begin
//      FileSleep.Add(inttostr(Etable^[iFileLoop].Apn));
//    end;
    FileSleep.SaveToFile(sPathPat + FILENAME_SLEEP_STAGES);
  finally
    FileSleep.Free
  end;


end;

{ ----------------------------------------------------------------------- }
/// <summary> Select the path of the patient </summary>
/// <remarks> Created by VLagarrigue </remarks>
procedure TFormExtract.ButtonPathClick(Sender: TObject);
var
  sPathPatient: string; // Path of the patient datas
  sPathAllPatients: string; // Path of all the patients datas
  tIniConfig: TIniFile; // Ini config file
  iLoopPat: ShortInt;

begin
  // Open a dialog box to select the new path
  // Get the default path of the patients in the ini file
  tIniConfig := TIniFile.create(ChangeFileExt(Application.ExeName, '.ini'));
  tIniConfig.ReadString('Config', 'Path', sPathAllPatients);
  if LENGTH(sPathAllPatients) > 0 then
    FileOpenDialogPath.DefaultFolder := sPathAllPatients;
  tIniConfig.Free;
  // Open the dialog box to choose the patient folder
  FileOpenDialogPath.create(nil);
  FileOpenDialogPath.Options := [fdoPickFolders];
  FileOpenDialogPath.Execute;
  sPathPatient := FileOpenDialogPath.FileName;
  if LENGTH(sPathPatient) > 0 then
  begin
    // Check if the path contains PATxxxx
    if ContainsText(sPathPatient, 'PAT') then // (use StrUtils)
    BEGIN
      // Delete the Patient folder to get the folder of all patients
      // Replace 'PATx' by 'PAT 3 time to delete 3 numbers after PAT (use RegularExpressions)
      for iLoopPat := 1 to 3 do
      BEGIN
        sPathAllPatients := TRegEx.Replace(sPathPatient, 'PAT\d{0,9}', 'PAT');
      END;
      sPathAllPatients := TRegEx.Replace(sPathAllPatients, 'PAT', '');

      // Save the "AllPatients" folder into the ini file
      tIniConfig := TIniFile.create(ChangeFileExt(Application.ExeName, '.ini'));
      tIniConfig.WriteString('Config', 'Path', sPathAllPatients);
      tIniConfig.Free;

      // Set the tEdit path on the interface
      EditPath.Text := sPathPatient + '\';

      // Enable data extract buttons
      FormExtract.ButtonExtractSleep.Enabled := true;
      FormExtract.ButtonExtractApnea.Enabled := true;

      // Display Patient Info on the GUI
      DisplayPatientInfo(sPathPatient);
    END
    else
    BEGIN
      // Error message and set the tEdit to the default directory
      ShowMessage('Please select a patient folder.' + AnsiString(#13#10) +
        'Must be "PATxxx"');
      FormExtract.EditPath.Text := sPathAllPatients;
      FormExtract.ButtonExtractSleep.Enabled := false;
      FormExtract.ButtonExtractApnea.Enabled := false;
    END;
  end; // End if folder chosen
end;

procedure TFormExtract.EditPathChange(Sender: TObject);
begin
  FormExtract.ButtonExtractSleep.Enabled := false;
  FormExtract.ButtonExtractApnea.Enabled := false;
end;

procedure TFormExtract.EditPathKeyPress(Sender: TObject; var Key: Char);
var
  sPathPatient: string; // Path of the patient datas
  sPathAllPatients: string; // Path of all the patients datas
  sPathAllPatients2: string;
  tIniConfig: TIniFile; // Ini config file
  iLoopPat: ShortInt;

begin
  // If the key pressed is enter
  if ord(Key) = VK_RETURN then
  begin
    // Get the default path of the patients in the ini file
    tIniConfig := TIniFile.create(ChangeFileExt(Application.ExeName, '.ini'));

    sPathPatient := EditPath.Text;
    if (LENGTH(sPathPatient) > 0) and directoryexists(sPathPatient) then
    begin
      // Check if the path contains PATxxxx
      if ContainsText(sPathPatient, 'PAT') then // (use StrUtils)
      BEGIN
        // Delete the Patient folder to get the folder of all patients
        // Replace 'PATx' by 'PAT 3 time to delete 3 numbers after PAT (use RegularExpressions)
        for iLoopPat := 1 to 3 do
        BEGIN
          sPathAllPatients := TRegEx.Replace(sPathPatient, 'PAT\d{0,9}', 'PAT');
        END;
        sPathAllPatients := TRegEx.Replace(sPathAllPatients, 'PAT', '');

        // If the last char is '\\', delete the second
        if AnsiRightStr(sPathAllPatients, 2) = '\\' then
        BEGIN
          sPathAllPatients := copy(sPathAllPatients, 0,
            (LENGTH(sPathAllPatients) - 1));
        END;

        // Save the "AllPatients" folder into the ini file
        tIniConfig := TIniFile.create
          (ChangeFileExt(Application.ExeName, '.ini'));
        tIniConfig.WriteString('Config', 'Path', sPathAllPatients);
        tIniConfig.Free;

        // Add '\' at the end of the string if is not
        if AnsiRightStr(sPathPatient, 1) <> '\' then
        BEGIN
          sPathPatient := sPathPatient + '\';
        END;
        // Set the tEdit path on the interface
        EditPath.Text := sPathPatient;

        // Enalble data extract buttons
        FormExtract.ButtonExtractSleep.Enabled := true;
        FormExtract.ButtonExtractApnea.Enabled := true;

        // Display Patient Info on the GUI
        DisplayPatientInfo(sPathPatient);
      END
      else
      BEGIN
        // Error message
        ShowMessage('Please select a patient folder.' + AnsiString(#13#10) +
          'Must be "PATxxx"');
        FormExtract.ButtonExtractSleep.Enabled := false;
        FormExtract.ButtonExtractApnea.Enabled := false;
      END;
    end
    else
    BEGIN
      ShowMessage('Folder doesn''t exist.');
      FormExtract.ButtonExtractSleep.Enabled := false;
      FormExtract.ButtonExtractApnea.Enabled := false;
    END;
  end;
end;

procedure TFormExtract.FormCreate(Sender: TObject);
var
  tIniConfig: TIniFile;
  sPathAllPatients: string;

begin
  // Check if ini config file exists.
  if (FileExists(ChangeFileExt(Application.ExeName, '.ini'))) then
  begin
    // If it exists get the path of the datas folder
    tIniConfig := TIniFile.create(ChangeFileExt(Application.ExeName, '.ini'));
    // Get the string 'Path' in the section 'Config'
    sPathAllPatients := tIniConfig.ReadString('Config', 'Path', '');
    // And set the tEdit path on the interface
    EditPath.Text := sPathAllPatients;
    // Set the default folder dialog path
    FileOpenDialogPath.DefaultFolder := sPathAllPatients;
    // ShowMessage (sPathAllPatients);
    tIniConfig.Free;
  end
  else
  begin
    // If it doesn't exist, create it
    tIniConfig := TIniFile.create(ChangeFileExt(Application.ExeName, '.ini'));
    // Get the string 'Path' in the section 'Config'
    tIniConfig.WriteString('Config', 'Path', 'C:\SHREDC\SIGEEG');
    tIniConfig.Free;
  end;

end;

end.
