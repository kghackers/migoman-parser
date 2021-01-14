   unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, IdHTTP, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, WinInet, Spin, ComCtrls, ExtCtrls, ShellAPI, ComObj,
  ActiveX, ClipBrd, JPEG;

type
  TForm1 = class(TForm)
    Open: TOpenDialog;
    Okno: TMemo;
    Id: TIdHTTP;
    Win: TMemo;
    POP: TEdit;
    POP0: TUpDown;
    Label1: TLabel;
    Label3: TLabel;
    PPO: TEdit;
    PPO0: TUpDown;
    Progress: TProgressBar;
    ZR: TCheckBox;
    Ima: TImage;
    Timer: TTimer;
    IU: TEdit;
    Label7: TLabel;
    KOS: TEdit;
    START: TButton;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    Label9: TLabel;
    GONIK: TCheckBox;
    GOP0: TUpDown;
    GOP: TEdit;
    GPO: TEdit;
    GPO0: TUpDown;
    GroupBox3: TGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    MANIK: TCheckBox;
    MOP0: TUpDown;
    MOP: TEdit;
    MPO: TEdit;
    MPO0: TUpDown;
    GroupBox4: TGroupBox;
    Label12: TLabel;
    Label13: TLabel;
    SUPER: TCheckBox;
    SOP0: TUpDown;
    SOP: TEdit;
    SPO: TEdit;
    SPO0: TUpDown;
    GroupBox5: TGroupBox;
    Label4: TLabel;
    KOC: TEdit;
    Label6: TLabel;
    KOK: TEdit;
    KOB: TEdit;
    KOO: TEdit;
    Label8: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    KOM: TEdit;
    PROFI: TCheckBox;
    Panel: TPanel;
    Label16: TLabel;
    PRODOL: TButton;
    Label2: TLabel;
    Panel1: TPanel;
    Save: TButton;
    Scroll: TScrollBox;
    procedure StartClick(Sender: TObject);
    procedure Server;
    procedure Cvet;
    procedure Rabota;
    procedure Tablica;
    procedure Bezosh;
    procedure Nezach;
    procedure TimerTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure POPKeyPress(Sender: TObject; var Key: Char);
    procedure KOSKeyPress(Sender: TObject; var Key: Char);
    procedure PRODOLClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SaveClick(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;
//********** Массив  dan **********\\
// dan[igrok][1][1] - id игрока
// dan[igrok][1][2] - ник игрока
// dan[igrok][1][3] - ранг игрока
// dan[igrok][1][4] - общий пробег
// dan[igrok][1][5] - пробег в обычном
// dan[igrok][1][6] - пробег в режиме
// dan[igrok][1][7] - пробег в турнире
// dan[igrok][1][8] -
// dan[igrok][1][9] -
// dan[igrok][1][10] - индикатор попадания в зачет
// dan[igrok][1][11] -
// dan[igrok][1][12] - максимальная скорость в безошибке
// dan[igrok][1][13] - переменная для сортировки безошибки
// dan[igrok][1][14] -
// dan[igrok][1][15] -
// dan[igrok][1][16] -
//*********************************\\
var
  Form1: TForm1;
  igrok,stran,cvetop : integer;
  raz, papka, cr : string;   //Разрешение файла
  dan : array[1..500,1..9,1..500] of string;
  lsk : array[1..500,1..11] of real;
  kol : array[2..5] of integer;
  dwConnectionTypes : DWORD;
  rang : char;
  cv1,cv2,cv3,ct1,ct2,ct3 : byte;
  t : tstringlist;
  EX : variant;
  box : array [1..300] of Tcheckbox;
  y : array[2..5] of integer;
  bo : array[1..500,1..4] of integer;
  bot : array[4..7] of integer;
  rekbez : array[1..500] of string;


implementation

{$R *.dfm}

function IsOLEObjectInstalled(Name: String): boolean;
var
  ClassID: TCLSID;
begin
  Result := CLSIDFromProgID(PWideChar(WideString(Name)), ClassID) = S_OK;
end;


function IsConnectedToInternet: Boolean;
begin
  dwConnectionTypes:= INTERNET_CONNECTION_MODEM + INTERNET_CONNECTION_LAN + INTERNET_CONNECTION_PROXY;
  Result := InternetGetConnectedState (@dwConnectionTypes, 0);
end;


Function PosEx(Const SubStr, S: String; Offset: Cardinal = 1): Integer;
var
I,X: Integer;
Len, LenSubStr: Integer;
begin
If Offset = 1 Then
   Result := Pos(SubStr, S)
Else
begin
   I := Offset;
   LenSubStr := Length(SubStr);
   Len := Length(S) - LenSubStr + 1;
   While I <= Len do
   begin
     If S[I] = SubStr[1] Then
     begin
       X := 1;
       While (X < LenSubStr) And (S[I + X] = SubStr[X + 1]) Do
         Inc(X);
       If (X = LenSubStr) Then
       begin
         Result := I;
         Exit;
       End;
     End;
     Inc(I);
   End;
   Result := 0;
End;
End;



//**********Импорт данных**********\\
procedure TForm1.StartClick(Sender: TObject);
var
 prov : boolean;
 imp : array [1..9] of string;
 f,f1,f2,f3,poisk,zaz : integer;
 iskl : array [1..500] of string;
begin
   For f1:=1 to 500 do iskl[f1]:='';
   f:=1;
   For f1:=1 to length(IU.Text) do
     Case IU.Text[f1] of
       '0','1','2','3','4','5','6','7','8','9' : iskl[f]:=iskl[f]+IU.Text[f1];
       ',' : f:=f+1;
     end;
  progress.Position:=0;
  For f1:=1 to 500 do For f2:=1 to 9 do For f3:=1 to 999 do dan[f1][f2][f3]:='';
  //////Проверка файла\\\\\\
  Open.FileName:='';
  If Open.Execute then
    Case Open.FileName[length(Open.FileName)] of
      'l' : raz:='.html';
      'm' : raz:='.htm'
      else exit;
    end;
  If Open.FileName='' then exit;
  //////Импорт данных из файлов\\\\\\
  progress.Position:=10;
  zaz:=1;
  Repeat
    poisk:=1;
    //загрузка файла\\
    If not(fileexists(inttostr(zaz)+raz)) then
      begin  showmessage('Файл '+inttostr(zaz)+raz+' не найден!'); exit; end;
    Okno.Clear;
    Okno.Lines.LoadFromFile(inttostr(zaz)+raz);
    //загрузка данных\\
    While posEx('</ins><ins id="rating_gained', Okno.Lines.text, poisk)>0 do begin
      For f1:=1 to 7 do imp[f1]:='';
      //рекорд\\
      poisk:=posEx('</ins><ins id="rating_gained', Okno.Lines.text, poisk);
      poisk:=poisk-950;
      poisk:=posEx('">    <a ng:show="PlayersList.records[', Okno.Lines.text, poisk);
      If not(Okno.Lines.Text[poisk-1]='e') then imp[7]:=inttostr(zaz);
      poisk:=posEx('</a>    <a ng:hide="PlayersList.records[', Okno.Lines.text, poisk);
      If not(Okno.Lines.Text[poisk-16]='e') then imp[7]:=inttostr(zaz);
      //cкорость\\
      poisk:=posEx('class="bitmore"><span class="bitmore">', Okno.Lines.text, poisk)+5;
      poisk:=posEx('class="bitmore"><span class="bitmore">', Okno.Lines.text, poisk)+5;
      For f1:=poisk+33 to poisk+38 do
        If not(Okno.Lines.Text[f1]='<') then imp[4]:=imp[4]+Okno.Lines.Text[f1] else break;
      //количество ошибок\\
      poisk:=posEx('class="bitmore"><span class="bitmore">', Okno.Lines.text, poisk)+5;
      For f1:=poisk+33 to poisk+38 do
        If not(Okno.Lines.Text[f1]='<') then imp[6]:=imp[6]+Okno.Lines.Text[f1] else break;
      //точность\\
      poisk:=posEx('class="bitmore"><span class="bitmore">', Okno.Lines.text, poisk)+5;
      For f1:=poisk+33 to poisk+39 do
        If not(Okno.Lines.Text[f1]='<') then imp[5]:=imp[5]+Okno.Lines.Text[f1] else break;
      //Id\\
      poisk:=posEx('/" class="rang', Okno.Lines.text, poisk);
      For f1:=poisk-1 downto poisk-30 do
        If not(Okno.Lines.Text[f1]='/') then imp[1]:=Okno.Lines.Text[f1]+imp[1] else break;
      //ранг\\
      poisk:=posEx('/" class="rang', Okno.Lines.text, poisk);
      imp[3]:= Okno.Lines.Text[poisk+14];
      //ник\\
      poisk:=posEx('</a></div></td></tr></tbody></table></div></td><td><div class="car ng-isolate-scope"', Okno.Lines.text, poisk);
      For f1:=poisk-1 downto poisk-30 do
        If not(Okno.Lines.Text[f1]='>') then imp[2]:=Okno.Lines.Text[f1]+imp[2] else break;

      //записываем значения заезда в массив данных (dan)\\
      prov:=true;
      For f1:=1 to f do If iskl[f1]=imp[1] then prov:=false;
      For f1:=1 to igrok do
        If (prov=true)and(dan[f1][1][1]=imp[1]) then begin
          dan[f1][1][3]:=imp[3];                 // ранг
          dan[f1][2][zaz]:=imp[4];               // скорость
          dan[f1][3][zaz]:=imp[5];               // точность
          dan[f1][4][zaz]:=imp[6];               // количество ошибок
          dan[f1][8][zaz]:=imp[7];               // рекорды
          dan[f1][9][zaz]:=imp[3];               // ранг
          prov:=false;
        end;
      If prov=true then begin
        igrok:=igrok+1;
        dan[igrok][1][1]:=imp[1];                // id
        dan[igrok][1][2]:=utf8toansi(imp[2]);    // ник
        dan[igrok][1][3]:=imp[3];                // ранг
        dan[igrok][2][zaz]:=imp[4];              // скорость
        dan[igrok][3][zaz]:=imp[5];              // точность
        dan[igrok][4][zaz]:=imp[6];              // количество ошибок
        dan[igrok][8][zaz]:=imp[7];              // рекорды
        dan[igrok][9][zaz]:=imp[3];              // ранг
        Server;
      end;
    end;
    progress.Position:=progress.Position+10;
    zaz:=zaz+1;
  Until zaz>20;
  Nezach;
  Rabota;
end;



procedure TForm1.Server;
var
  s,f,f3 : integer;
begin
    //общий пробег\\
    Win.Clear;
    Win.Lines.Text:=Id.Get('http://klavogonki.ru/ajax/profile-popup?user_id='+dan[igrok][1][1]+'&gametype=normal');
    Win.Lines.Text:=utf8toansi(Win.Lines.Text);
    s:=pos(' текст',win.Lines.Text);
    For f:=s-1 downto s-11 do If not(Win.Lines.Text[f]='>')
      then dan[igrok][1][4]:=Win.Lines.Text[f]+dan[igrok][1][4] else break;
    //пробег в обычном режиме\\
    s:=pos('<th>Пробег:</th>',win.Lines.Text);
    For f3:=s+22 to s+32 do If not(Win.Lines.Text[f3]=' ')
      then dan[igrok][1][5]:=dan[igrok][1][5]+Win.Lines.Text[f3] else break;
    //рекорд в безошибке\\
    Win.Clear;
    Win.Lines.Text:=Id.Get('http://klavogonki.ru/ajax/profile-popup?user_id='+dan[igrok][1][1]+'&gametype=noerror');
    Win.Lines.Text:=utf8toansi(Win.Lines.Text);
    s:=pos('<th>Лучшая скорость:</th>',win.Lines.Text);
    For f:=s+31 to s+35 do If not(Win.Lines.Text[f]=' ')
      then rekbez[igrok]:=rekbez[igrok]+Win.Lines.Text[f] else break;
end;



procedure TForm1.Rabota;
var
  f,f1,f2,f3,f4,f5: integer;
begin
  For f1:=1 to 500 do For f2:=1 to 4 do bo[f1][f2]:=0;

  //Вычисляем пробег в турнире\\
  For f1:=1 to igrok do begin dan[f1][1][7]:='0';
    For f2:=1 to 17 do If dan[f1][2][f2]<>'' then
      dan[f1][1][7]:=inttostr(strtoint(dan[f1][1][7])+1);
  end;

  //Меняем ранг, отталкиваясь от Безошибки\\
  For f1:=1 to igrok do
    If strtoint(rekbez[f1][1])>=strtoint(dan[f1][1][3]) then dan[f1][1][3]:=inttostr(strtoint(rekbez[f1][1])+1);

  //Вычисляем наилучшие скорости\\
  For f1:=1 to igrok do For f2:=1 to 10 do lsk[f1][f2]:=0;
  For f1:=1 to igrok do For f2:=1 to 20 do If dan[f1][2][f2]='' then dan[f1][2][f2]:='0';
  For f1:=2 to 5 do kol[f1]:=0;

  For f1:=1 to igrok do begin

      //Вычисляем наилучшую скорость в безошибке\\
      f:=18; dan[f1][1][11]:='0';
      For f2:=19 to 20 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      dan[f1][1][12]:=dan[f1][2][f];
      For f2:=18 to 20 do
        dan[f1][1][11]:=inttostr(strtoint(dan[f1][1][11])+strtoint(dan[f1][2][f2]));

      //Вычисляем наилучшую скорость в остальных режимах\\
      f4:=1; f:=1;
      For f2:=2 to 3 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

      f4:=2; f:=4;
      For f2:=5 to 6 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

      f4:=3; f:=7;
      For f2:=8 to 9 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

      f4:=4; f:=10;
      For f2:=11 to 12 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

      f4:=5; f:=13;
      For f2:=14 to 15 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

      f4:=6; f:=16;
      For f2:=17 to 17 do If strtoint(dan[f1][2][f])<strtoint(dan[f1][2][f2]) then f:=f2;
      lsk[f1][f4]:=strtoint(dan[f1][2][f]);
      dan[f1][5][f]:='1';

  end;

  //Вычисляем сумму скоростей и зачетный результат\\
  For f1:=1 to igrok do begin
    For f2:=1 to 6 do begin
      If lsk[f1][f2]>0 then lsk[f1][7]:=lsk[f1][7]+1;
      lsk[f1][8]:=lsk[f1][8]+lsk[f1][f2];
      Case f2 of
        1 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOS.Text));
        2 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOC.Text));
        3 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOK.Text));
        4 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOO.Text));
        5 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOB.Text));
        6 : lsk[f1][9]:=lsk[f1][9]+(lsk[f1][f2]*strtofloat(KOM.Text));
      end
    end;
  end;

  //Количистео ошибок в безошибке\\
  For f1:=1 to igrok do begin
    For f2:=18 to 20 do If (dan[f1][4][f2]<>'') then begin
      bo[f1][1]:=bo[f1][1]+1;
      bo[f1][2]:=bo[f1][2]+strtoint(dan[f1][4][f2]);
    end;
    If (bo[f1][1]=3)and(bo[f1][2]=0) then bo[f1][4]:=1000
      else begin
        bo[f1][4]:=(bo[f1][1]*100)+((bo[f1][1]*50)-(bo[f1][2]*50));
        If bo[f1][1]=3 then bo[f1][4]:=bo[f1][4]+300;
      end;
  end;

  //Включение рангов в таблицу\\
  For f1:=1 to igrok do
    Case dan[f1][1][3][1] of
      '4' : If PROFI.Checked = true then dan[f1][1][10]:='2';
      '5' : If GONIK.Checked = true then dan[f1][1][10]:='2';
      '6' : If MANIK.Checked = true then dan[f1][1][10]:='2';
      '7' : If SUPER.Checked = true then dan[f1][1][10]:='2';
    end;



  For f1:=1 to igrok do If dan[f1][1][10]='2' then
    Case dan[f1][1][3][1] of
      '7' : If (strtoint(dan[f1][1][4])<strtoint(SOP.Text))or(strtoint(dan[f1][1][5])<strtoint(SPO.Text)) then dan[f1][1][10]:='3';
      '6' : If (strtoint(dan[f1][1][4])<strtoint(MOP.Text))or(strtoint(dan[f1][1][5])<strtoint(MPO.Text)) then dan[f1][1][10]:='3';
      '5' : If (strtoint(dan[f1][1][4])<strtoint(GOP.Text))or(strtoint(dan[f1][1][5])<strtoint(GPO.Text)) then dan[f1][1][10]:='3';
      '4' : If (strtoint(dan[f1][1][4])<strtoint(POP.Text))or(strtoint(dan[f1][1][5])<strtoint(PPO.Text)) then dan[f1][1][10]:='3';
    end;
  f2:=10;
  For f1:=1 to igrok do If dan[f1][1][10]='3' then begin
    box[f1]:=Tcheckbox.Create(Scroll);
    With box[f1] do begin
      Parent:=Scroll;
      Height:=20;
      Width:=150;
      Left:=20;
      Top:=f2;
      Name:='box'+ inttostr(f1);
      Caption:=dan[f1][1][2];
    end;
    f2:=f2+21;
  end;

  Panel.Visible:=true;
  PRODOL.SetFocus;
end;



procedure TForm1.Tablica;
var
  naz,g,m,d,sl : string;
  f,f1,f2,f3,f4,f5,fm,w : integer;
  n : array[1..11] of integer;
  ocki : array[2..5] of integer;
  lid : array[2..5,1..3] of string;
  tera : tstringlist;
  troj : array[8..11,1..9] of string;
  top : array[1..5000,1..8] of string;
  mesto : array[8..11] of integer;
  mt : array[8..11] of integer;
  prov : boolean;
  god, mes, den: word;
  zara : array[4..7,1..5] of integer; // зачет рангов
  rek : array[4..7] of string;
  ochrek : integer;
begin
For f1:=4 to 7 do bot[f1]:=0;
DecodeDate(date,god,mes,den);
g:=inttostr(god);
if mes<10 then m:='0'+inttostr(mes) else m:=inttostr(mes);
if den<10 then d:='0'+inttostr(den) else d:=inttostr(den);
  t:=tstringlist.Create;
  If fileexists(papka+'\top.txt') then begin
    t.LoadFromFile(papka+'\top.txt');
    Win.Clear;
    For f1:=0 to t.Count-1 do begin
      t.Strings[f1]:=trim(t.Strings[f1])+'	';
      f3:=1;
      For f2:=1 to length(t.Strings[f1]) do
        If t.Strings[f1][f2]<>'	'
          then Win.SelText:=t.Strings[f1][f2]
          else begin
            If Win.Lines.Text<>'' then top[f1+1][f3]:=Win.Lines.Text;
            Win.Clear; f3:=f3+1;
          end;
    end;
    Win.Clear;
    w:=t.Count;
  end else w:=0;


  tera := tstringlist.Create;
  //Запуск экселя\\
  DeleteFile(GetCurrentDir+'/Migoman.xlsx');
  If fileexists('Migoman.xlsx') then begin
    For f1:=1 to 1000 do begin
      DeleteFile(GetCurrentDir+'/'+'Migoman'+inttostr(f1)+'.xlsx');
      If not(fileexists('Migoman'+inttostr(f1)+'.xlsx')) then
        begin naz:= 'Migoman'+inttostr(f1)+'.xlsx'; break; end;
    end;
  end
  else naz:='Migoman.xlsx';

  CopyFile(pchar(papka+'/1.xlsx'),pchar(naz),true);
  EX := CreateOleObject('Excel.Application');
  EX.DisplayAlerts := false;
  EX.WorkBooks.Open(GetCurrentDir+'/'+naz);



  //Зачет по рангам\\
  For f1:=1 to igrok do
    If dan[f1][1][10]='2' then If lsk[f1][7]>5 then begin
       zara[strtoint(dan[f1][1][3])][1]:=zara[strtoint(dan[f1][1][3])][1]+1;
       if bo[f1][1]>0 then
         zara[strtoint(dan[f1][1][3])][2]:=zara[strtoint(dan[f1][1][3])][2]+1;
    end;
  For f1:=5 to 7 do If zara[f1][1]>0 then
    zara[f1][2]:=Round(zara[f1][2] * 100 / zara[f1][1]);

  //место в зачетах\\
  For f1:=5 to 7 do begin
    f5:=1; f3:=1;
    For f2:=5 to 7 do begin
      If zara[f1][1]<zara[f2][1] then f5:=f5+1;
      If zara[f1][2]<zara[f2][2] then if zara[f2][1]>3 then f3:=f3+1;
    end;
    If zara[f1][1]>3
      then begin zara[f1][3]:=f5; zara[f1][4]:=f3 end
      else begin zara[f1][3]:=3; zara[f1][4]:=3 end;
  end;
  //сортировки\\
  For f1:=1 to 3 do begin
    f:=5;
    For f2:=6 to 7 do If zara[f][1] < zara[f2][1] then f:=f2 else
      If zara[f][1] = zara[f2][1] then
        If zara[f][2] < zara[f2][2] then f:=f2;

    Case f of 5 : sl:='Гонщики'; 6 : sl:='Маньяки'; 7 : sl:='Супермены'; end;
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,2]:=sl;
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,3]:=zara[f][1];
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,4]:=zara[f][2];
    Case zara[f][3] of 1 : zara[f][3]:=300; 2 : zara[f][3]:=100; 3 : zara[f][3]:=0; end;
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,5]:=zara[f][3];
    Case zara[f][4] of 1 : zara[f][4]:=200; 2 : zara[f][4]:=100; 3 : zara[f][4]:=0; end;
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,6]:=zara[f][4];
    zara[f][5]:=zara[f][3]+zara[f][4];
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,7]:=zara[f][5];
    rang:=inttostr(f)[1]; Cvet;
    EX.WorkBooks[1].WorkSheets[13].Cells[f1+2,2].interior.color:=rgb(cv1,cv2,cv3);
    zara[f][1]:=-5;
  end;

  //Главный зачет\\
  For f1:=1 to 5 do n[f1]:=4;
  For f1:=1 to igrok do If ZR.Checked=true
    then begin lsk[f1][10]:=lsk[f1][9]; lsk[f1][11]:=lsk[f1][8] end
    else begin lsk[f1][10]:=lsk[f1][8]; lsk[f1][11]:=lsk[f1][9] end;
  For f1:=1 to igrok do begin
    f:=1;
    For f2:=2 to igrok do
      If lsk[f][10]<lsk[f2][10] then f:=f2 else
        If lsk[f][10]=lsk[f2][10] then
          If lsk[f][11]<lsk[f2][11] then f:=f2;

    If (dan[f][1][10]='1')or(dan[f][1][10]='2') then
      If (lsk[f][10]>0)and(strtoint(dan[f][1][7])>=5) then begin
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],2]:=dan[f][1][1];
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],3]:=n[1]-3;
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],4]:=dan[f][1][2];
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],5]:=inttostr(Round(lsk[f][11]));
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6]:=inttostr(Round(lsk[f][10]));
      EX.WorkBooks[1].WorkSheets[1].Range['b'+inttostr(n[1])+':y'+inttostr(n[1])].Borders.Weight := 2;
      For f3:=1 to 17 do begin
        If dan[f][2][f3]<>'0' then EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3]:=dan[f][2][f3];
        If dan[f][5][f3]='1' then EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3].Font.Bold := true;
        Case f3 of
          13,14,15 : If dan[f][8][f3]<>'' then begin
            EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3].Borders.Weight := 4;
            If strtoint(dan[f][2][f3])<strtoint(dan[f][9][f3])*100
              then EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3].Borders.color := clred
              else EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3].Borders.color := clblue;
          end;
        end;
      end;
      f3:=18;
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],6+f3] := dan[f][1][5];
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],7+f3] := dan[f][1][4];
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],8+f3] := 'Профиль';
      EX.WorkBooks[1].WorkSheets[1].Cells[n[1],8+f3].Hyperlinks.add(EX.WorkBooks[1].WorkSheets[1].Cells[n[1],8+f3],Address:='http://klavogonki.ru/u/#/'+dan[f][1][1]);
      rang:=dan[f][1][3][1]; Cvet;
      EX.WorkBooks[1].WorkSheets[1].Range['b'+inttostr(n[1])+':y'+inttostr(n[1])].interior.color:=rgb(cv1,cv2,cv3);
      n[1]:=n[1]+1;
    end;
    lsk[f][10]:=0;
  end;

  Progress.Position:=Progress.Position+10;

  //Копирование ячеек\\
  EX.WorkBooks[1].WorkSheets[1].Range['b2:b3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['c2:c3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['d2:d3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['e2:e3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['f2:f3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['g2:w2'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['g3:i3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['j3:l3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['m3:o3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['p3:r3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['s3:u3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['v3:w3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['x2:x3'].Merge;
  EX.WorkBooks[1].WorkSheets[1].Range['y2:y3'].Merge;
  //Заливка таблицы\\
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y3'].interior.color:=rgb(191,191,191);
  //ширина ячеек\\
  EX.WorkBooks[1].WorkSheets[1].Rows[3].RowHeight := 30;
  EX.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth := 3;
  EX.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth := 7;
  EX.WorkBooks[1].WorkSheets[1].Columns[3].ColumnWidth := 6;
  EX.WorkBooks[1].WorkSheets[1].Columns[4].ColumnWidth := 18;
  EX.WorkBooks[1].WorkSheets[1].Columns[5].ColumnWidth := 9;
  EX.WorkBooks[1].WorkSheets[1].Columns[6].ColumnWidth := 9;
  EX.WorkBooks[1].WorkSheets[1].Columns[24].ColumnWidth := 7.5;
  EX.WorkBooks[1].WorkSheets[1].Columns[25].ColumnWidth := 7.5;
  For f1:=1 to 17 do EX.WorkBooks[1].WorkSheets[1].Columns[6+f1].ColumnWidth := 5;
  //выравнивание по центру\\
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y'+inttostr(n[1])].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y'+inttostr(n[1])].VerticalAlignment :=2;
  EX.WorkBooks[1].WorkSheets[1].Columns[4].HorizontalAlignment :=2;
  EX.WorkBooks[1].WorkSheets[1].Cells[2,4].HorizontalAlignment :=3;
  //жирный шрифт\\
  EX.WorkBooks[1].WorkSheets[1].Columns[3].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[1].Columns[6].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[1].Columns[26].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y3'].Font.Bold := true;
  //заполнение ячеек в шапке таблицы\\
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y3'].Wraptext:=true;
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y3'].Font.Size:=9;
  EX.WorkBooks[1].WorkSheets[1].Cells[2,2] := 'ID';
  EX.WorkBooks[1].WorkSheets[1].Cells[2,3] := 'Место';
  EX.WorkBooks[1].WorkSheets[1].Cells[2,4] := 'Ник';
  If ZR.Checked=true then EX.WorkBooks[1].WorkSheets[1].Cells[2,5] := 'Сумма лучших скоростей'
                     else EX.WorkBooks[1].WorkSheets[1].Cells[2,5] := 'Скорректи-рованный рез-тат';
  If ZR.Checked=true then EX.WorkBooks[1].WorkSheets[1].Cells[2,6] := 'Зачётный результат'
                     else EX.WorkBooks[1].WorkSheets[1].Cells[2,6] := 'Сумма лучших скоростей';
  EX.WorkBooks[1].WorkSheets[1].Cells[2,7] := 'Скорость в заездах';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,7] := 'Абракадабра';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,10] := 'Короткие';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,13] := 'Соточка';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,16] := 'Частотный';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,19] := 'Обычный';
  EX.WorkBooks[1].WorkSheets[1].Cells[3,22] := 'Мини-марафон';
  EX.WorkBooks[1].WorkSheets[1].Cells[2,24] := 'Пробег в обычном';
  EX.WorkBooks[1].WorkSheets[1].Cells[2,25] := 'Пробег общий';
  //границы таблицы\\
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y3'].Borders.Weight := 2;
  EX.WorkBooks[1].WorkSheets[1].Range['b2:b'+inttostr(n[1]-1)].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['y2:y'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['b2:y2'].Borders[3].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['b'+inttostr(n[1]-1)+':y'+inttostr(n[1]-1)].Borders[4].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['d2:d'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['e2:e'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['f2:f'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['i2:i'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['l2:l'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['o2:o'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['r2:r'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['u2:u'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['w2:w'+inttostr(n[1]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[1].Range['g2:y2'].Borders[4].Weight := 3;

  //Ранговые зачеты\\
  //Проверка на пробег\\
  For f1:=1 to igrok do
    If dan[f1][1][10]='2' then
      Case dan[f1][1][3][1] of
        '4' : If lsk[f1][7]>5 then begin kol[5]:=kol[5]+1; if bo[f1][1]>0 then bot[4]:=bot[4]+1 end;
        '5' : If lsk[f1][7]>5 then begin kol[4]:=kol[4]+1; if bo[f1][1]>0 then bot[5]:=bot[5]+1 end;
        '6' : If lsk[f1][7]>5 then begin kol[3]:=kol[3]+1; if bo[f1][1]>0 then bot[6]:=bot[6]+1 end;
        '7' : If lsk[f1][7]>5 then begin kol[2]:=kol[2]+1; if bo[f1][1]>0 then bot[7]:=bot[7]+1 end;
      end;

  y[2]:=12;
  y[3]:=12+kol[2];
  y[4]:=12+kol[2]+kol[3];
  y[5]:=12+kol[2]+kol[3]+kol[4];
  For f1:=2 to 5 do If kol[f1]>=4 then ocki[f1]:=5000 else ocki[f1]:=1000+(kol[f1]*1000);

  //Сортировка\\
  For f1:=1 to igrok do If ZR.Checked=true
    then begin lsk[f1][10]:=lsk[f1][9]; lsk[f1][11]:=lsk[f1][8] end
    else begin lsk[f1][10]:=lsk[f1][8]; lsk[f1][11]:=lsk[f1][9] end;
  For f1:=1 to igrok do begin
    f:=1;
    For f2:=2 to igrok do
      If lsk[f][10]<lsk[f2][10] then f:=f2 else
        If lsk[f][10]=lsk[f2][10] then
          If lsk[f][11]<lsk[f2][11] then f:=f2;

    Case dan[f][1][3][1] of
      '7' : stran:=2;
      '6' : stran:=3;
      '5' : stran:=4;
      '4' : stran:=5;
    end;

    If (dan[f][1][10]='2')and(lsk[f][10]>0)and(lsk[f][7]=6) then begin
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],2]:=dan[f][1][1];
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],3]:=n[stran]-3;
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],4]:=dan[f][1][2];
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],5]:=inttostr(Round(lsk[f][11]));
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],6]:=inttostr(Round(lsk[f][10]));
      For f3:=1 to 6 do If lsk[f][f3]>0 then EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],6+f3]:=lsk[f][f3];
      EX.WorkBooks[1].WorkSheets[stran].Range['g'+inttostr(n[stran])+':l'+inttostr(n[stran])].Borders.Weight := 2;

      If (dan[f][8][13]<>'')or(dan[f][8][14]<>'')or(dan[f][8][15]<>'') then begin
        EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],11].Borders.Weight := 4;
        EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],11].Borders.Color := clred;
      end;


      If lsk[f][7]=6 then begin

       ochrek:=0;
      //Обработка топа\\
        prov:=true;
        For f4:=1 to w do If (top[f4][1]=dan[f][1][1])and(top[f4][8]=dan[f][1][3])and(prov=true) then begin
          top[f4][2]:=dan[f][1][2];
          If Round(lsk[f][10])>strtoint(top[f4][3]) then top[f4][3]:=floattostr(Round(lsk[f][10]));
          If n[stran]-3=1 then top[f4][4]:=inttostr(strtoint(top[f4][4])+1);
          If n[stran]-3=2 then top[f4][5]:=inttostr(strtoint(top[f4][5])+1);
          If n[stran]-3=3 then top[f4][6]:=inttostr(strtoint(top[f4][6])+1);
          top[f4][7]:=inttostr(strtoint(top[f4][4])+strtoint(top[f4][5])+strtoint(top[f4][6]));
          prov:= false;
          break;
        end;
        If (prov=true) then begin
          w:=w+1;
          top[w][1]:=dan[f][1][1];
          top[w][2]:=dan[f][1][2];
          top[w][3]:=floattostr(Round(lsk[f][10]));
          If n[stran]-3=1 then top[w][4]:='1' else top[w][4]:='0';
          If n[stran]-3=2 then top[w][5]:='1' else top[w][5]:='0';
          If n[stran]-3=3 then top[w][6]:='1' else top[w][6]:='0';
          top[w][7]:=inttostr(strtoint(top[w][4])+strtoint(top[w][5])+strtoint(top[w][6]));
          top[w][8]:=dan[f][1][3];
        end;


        If kol[stran]>3 then
          Case n[stran]-3 of
            1 : begin troj[stran+6][1]:=dan[f][1][2]; troj[stran+6][2]:=inttostr(Round(lsk[f][10])); troj[stran+6][7]:=dan[f][1][1]; end;
            2 : begin troj[stran+6][3]:=dan[f][1][2]; troj[stran+6][4]:=inttostr(Round(lsk[f][10])); troj[stran+6][8]:=dan[f][1][1]; end;
            3 : begin troj[stran+6][5]:=dan[f][1][2]; troj[stran+6][6]:=inttostr(Round(lsk[f][10])); troj[stran+6][9]:=dan[f][1][1]; end;
          end;

        If ocki[stran]>1000 then ocki[stran]:=ocki[stran]-1000 else
          If (ocki[stran]>(kol[stran]-13+1)*100)and(ocki[stran]>100) then ocki[stran]:=ocki[stran]-100 else
            If ocki[stran]>100 then ocki[stran]:=ocki[stran]-50;
        EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],6+f3] := ocki[stran];
        rang:=dan[f][1][3][1]; Cvet;
        If (dan[f][8][13]<>'')or(dan[f][8][14]<>'')or(dan[f][8][15]<>'') then begin
          If rek[strtoint(dan[f][1][3])]<>'' then rek[strtoint(dan[f][1][3])]:=rek[strtoint(dan[f][1][3])]+', ';
          rek[strtoint(dan[f][1][3])]:=rek[strtoint(dan[f][1][3])]+'[url="http://klavogonki.ru/u/#/'+dan[f][1][1]+'/"][color="'+cr+'"][b][u]'+dan[f][1][2]+'[/u][/b][/color][/url] – 1000 очков';
          ochrek:=1000;
        end;
        //Награда\\
        EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],2] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],3] := ocki[stran]+ochrek+zara[strtoint(dan[f][1][3])][5];

        If kol[stran]>3 then begin
          Case stran of
            2 : Case n[stran]-3 of
                  1 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![1-е место: Поздравляю!](https://a.radikal.ru/a34/2005/ce/d776729880da.gif)"';
                  2 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![2-е место: Поздравляю!](https://a.radikal.ru/a25/2005/3d/ceca433f262e.gif)"';
                  3 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![3-е место: Поздравляю!](https://b.radikal.ru/b22/2005/8e/c1e4e77d3538.gif)"';
                end;
            3 : Case n[stran]-3 of
                  1 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![1-е место: Поздравляю!](https://b.radikal.ru/b09/2005/8e/256075539b57.gif)"';
                  2 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![2-е место: Поздравляю!](https://c.radikal.ru/c07/2005/81/705d4ba90ed0.gif)"';
                  3 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![3-е место: Поздравляю!](https://b.radikal.ru/b17/2006/fa/cfcd65f2e1fc.gif)"';
                end;
            4 : Case n[stran]-3 of
                  1 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![1-е место: Поздравляю!](https://c.radikal.ru/c06/2005/eb/b43003f878dc.gif)"';
                  2 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![2-е место: Поздравляю!](https://c.radikal.ru/c01/2005/d8/cd4cda40fe61.gif)"';
                  3 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![3-е место: Поздравляю!](https://c.radikal.ru/c00/2005/7f/6a93a6829107.gif)"';
                end;
            5 : Case n[stran]-3 of
                  1 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![1-е место: Поздравляю!](https://c.radikal.ru/c23/2005/a9/291572c1a43b.gif)"';
                  2 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![2-е место: Поздравляю!](https://d.radikal.ru/d08/2005/13/5c0a0f85df3f.gif)"';
                  3 : EX.WorkBooks[1].WorkSheets[7].Cells[y[stran],5] := '="**Участнику турнира [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&")**![3-е место: Поздравляю!](https://d.radikal.ru/d15/2005/4d/f2e0bef33762.gif)"';
                end;
          end;

          rang:=dan[f][1][3][1]; Cvet;
          Case n[stran]-3 of
            1 : lid[stran][n[stran]-3]:='[img]http://klavogonki.ru/img/smilies/first.gif[/img] – [url="http://klavogonki.ru/u/#/'+dan[f][1][1]+'/"][color="'+cr+'"][b][u]'+dan[f][1][2]+'[/u][/b][/color][/url] –  '+inttostr(ocki[stran])+' очков';
            2 : lid[stran][n[stran]-3]:='[img]http://klavogonki.ru/img/smilies/second.gif[/img] – [url="http://klavogonki.ru/u/#/'+dan[f][1][1]+'/"][color="'+cr+'"][b][u]'+dan[f][1][2]+'[/u][/b][/color][/url] –  '+inttostr(ocki[stran])+' очков';
            3 : lid[stran][n[stran]-3]:='[img]http://klavogonki.ru/img/smilies/third.gif[/img] – [url="http://klavogonki.ru/u/#/'+dan[f][1][1]+'/"][color="'+cr+'"][b][u]'+dan[f][1][2]+'[/u][/b][/color][/url] –  '+inttostr(ocki[stran])+' очков';
          end;
        end;
        y[stran]:=y[stran]+1;
      end else EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],6+f3] := '.';

      n[stran]:=n[stran]+1;
    end;
    lsk[f][10]:=0;
  end;

  Progress.Position:=Progress.Position+10;

  For stran:=2 to 5 do begin
    Case stran of
      2 : rang:='7';
      3 : rang:='6';
      4 : rang:='5';
      5 : rang:='4';
    end;
    Cvet; EX.WorkBooks[1].WorkSheets[stran].tab.color:=rgb(cv1,cv2,cv3);
    //Объединение ячеек\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:b3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['c2:c3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['d2:d3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['e2:e3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['f2:f3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['g2:l2'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['m2:m3'].Merge;
    //Заливка таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].interior.color:=rgb(191,191,191);
    //ширина ячеек\\
    EX.WorkBooks[1].WorkSheets[stran].Rows[3].RowHeight := 30;
    EX.WorkBooks[1].WorkSheets[stran].Columns[1].ColumnWidth := 3;
    EX.WorkBooks[1].WorkSheets[stran].Columns[2].ColumnWidth := 7;
    EX.WorkBooks[1].WorkSheets[stran].Columns[3].ColumnWidth := 6;
    EX.WorkBooks[1].WorkSheets[stran].Columns[4].ColumnWidth := 18;
    EX.WorkBooks[1].WorkSheets[stran].Columns[5].ColumnWidth := 9;
    EX.WorkBooks[1].WorkSheets[stran].Columns[6].ColumnWidth := 9;
    For f1:=1 to 6 do EX.WorkBooks[1].WorkSheets[stran].Columns[6+f1].ColumnWidth := 8.20;
    EX.WorkBooks[1].WorkSheets[stran].Columns[13].ColumnWidth := 6;
    //выравнивание по центру\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m'+inttostr(n[stran])].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m'+inttostr(n[stran])].VerticalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Columns[4].HorizontalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,4].HorizontalAlignment :=3;
    //жирный шрифт\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].Font.Bold := true;
    EX.WorkBooks[1].WorkSheets[stran].Columns[3].Font.Bold := true;
    EX.WorkBooks[1].WorkSheets[stran].Columns[4].Font.Bold := true;
    For f1:=6 to 12 do EX.WorkBooks[1].WorkSheets[stran].Columns[f1].Font.Bold := true;
    //заполнение ячеек в шапке таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].Wraptext:=true;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].Font.Size:=9;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,2] := 'ID';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,3] := 'Место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,4] := 'Ник';
    If ZR.Checked=true then EX.WorkBooks[1].WorkSheets[stran].Cells[2,5] := 'Сумма лучших скоростей'
                       else EX.WorkBooks[1].WorkSheets[stran].Cells[2,5] := 'Скорректи-рованный рез-тат';
    If ZR.Checked=true then EX.WorkBooks[1].WorkSheets[stran].Cells[2,6] := 'Зачётный результат'
                       else EX.WorkBooks[1].WorkSheets[stran].Cells[2,6] := 'Сумма лучших скоростей';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,7] := 'Лучшие скорости';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,7] := 'Абра-кадабра';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,8] := 'Короткие';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,9] := 'Соточка';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,10] := 'Частотный';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,11] := 'Обычный';
    EX.WorkBooks[1].WorkSheets[stran].Cells[3,12] := 'Мини-марафон';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,13] := 'Сумма очков';
    //границы таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['b3:b'+inttostr(n[stran]-1)].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['d3:e'+inttostr(n[stran]-1)].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:b'+inttostr(n[stran]-1)].Borders[1].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['m2:m'+inttostr(n[stran]-1)].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m2'].Borders[3].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['b'+inttostr(n[stran]-1)+':m'+inttostr(n[stran]-1)].Borders[4].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['d2:d'+inttostr(n[stran]-1)].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['e2:e'+inttostr(n[stran]-1)].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['f2:f'+inttostr(n[stran]-1)].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[stran].Range['l2:l'+inttostr(n[stran]-1)].Borders[2].Weight := 3;
  end;



  //Форум\\
  fm:=15;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[b][url="""&C4&"""]МиГоМан № "&C3&"[/url][/b]"'; fm:=fm+2;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Все заезды всех участников:[/b]'; fm:=fm+1;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[hide][img]"&C5&"[/img]"'; fm:=fm+1;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C6&"[/img][/hide]"'; fm:=fm+1;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Командный зачёт:[/b]'; fm:=fm+1;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C7&"[/img]"'; fm:=fm+2;
  If SUPER.Checked=true then begin
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Итоговая таблица Супермены:[/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C8&"[/img]"'; fm:=fm+2;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[list][b]НАГРАДЫ:[/b]'; fm:=fm+1;
    For f1:=1 to 3 do If lid[2][f1]<>'' then begin
      EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := lid[2][f1]; fm:=fm+1;
    end;
    fm:=fm+1;
    If rek[7]<>'' then begin EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b]Личный рекорд в Обычном:[/b] '+rek[7]; fm:=fm+2; end;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b][color="#034BAF"]Бонус - Безошибочный:[/color][/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='="[img]"&D8&"[/img]"'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[/list]'; fm:=fm+2;
  end else EX.WorkBooks[1].WorkSheets[6].Rows[8].EntireRow.Hidden := True;

  If MANIK.Checked=true then begin
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Итоговая таблица Маньяки:[/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C9&"[/img]"'; fm:=fm+2;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[list][b]НАГРАДЫ:[/b]'; fm:=fm+1;
    For f1:=1 to 3 do If lid[3][f1]<>'' then begin
      EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := lid[3][f1]; fm:=fm+1;
    end;
    fm:=fm+1;
    If rek[6]<>'' then begin EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b]Личный рекорд в Обычном:[/b] '+rek[6]; fm:=fm+2; end;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b][color="#034BAF"]Бонус - Безошибочный:[/color][/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='="[img]"&D9&"[/img]"'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[/list]'; fm:=fm+2;
  end else EX.WorkBooks[1].WorkSheets[6].Rows[9].EntireRow.Hidden := True;

  If GONIK.Checked=true then begin
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Итоговая таблица Гонщики:[/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C10&"[/img]"'; fm:=fm+2;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[list][b]НАГРАДЫ:[/b]'; fm:=fm+1;
    For f1:=1 to 3 do If lid[4][f1]<>'' then begin
      EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := lid[4][f1]; fm:=fm+1;
    end;
    fm:=fm+1;
    If rek[5]<>'' then begin EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b]Личный рекорд в Обычном:[/b] '+rek[5]; fm:=fm+2; end;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b][color="#034BAF"]Бонус - Безошибочный:[/color][/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='="[img]"&D10&"[/img]"'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[/list]'; fm:=fm+2;
  end else EX.WorkBooks[1].WorkSheets[6].Rows[10].EntireRow.Hidden := True;

  If PROFI.Checked=true then begin
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b]Итоговая таблица Профи:[/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '="[img]"&C11&"[/img]"'; fm:=fm+2;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[list][b]НАГРАДЫ:[/b]'; fm:=fm+1;
    For f1:=1 to 3 do If lid[5][f1]<>'' then begin
      EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := lid[5][f1]; fm:=fm+1;
    end;
    fm:=fm+1;
    If rek[4]<>'' then begin EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b]Личный рекорд в Обычном:[/b] '+rek[4]; fm:=fm+2; end;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='[b][color="#034BAF"]Бонус - Безошибочный:[/color][/b]'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] :='="[img]"&D11&"[/img]"'; fm:=fm+1;
    EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[/list]'; fm:=fm+2;
  end else EX.WorkBooks[1].WorkSheets[6].Rows[11].EntireRow.Hidden := True;

  fm:=fm+1;
  EX.WorkBooks[1].WorkSheets[6].Cells[fm,2] := '[b][color="#5000C9"]Спасибо всем за азартную игру! В следующее воскресенье, на том же месте, в тот же час![/color][/b] [img]http://klavogonki.ru/img/smilies/formula1.gif[/img]';
  //объединение ячеек\\
  EX.WorkBooks[1].WorkSheets[6].Range['b2:c2'].Merge;
  //Заливка таблицы\\
  EX.WorkBooks[1].WorkSheets[6].Cells[2,2].interior.color:=rgb(217,217,217);
  EX.WorkBooks[1].WorkSheets[6].Cells[2,4].interior.color:=rgb(217,217,217);
  EX.WorkBooks[1].WorkSheets[6].Cells[14,2].interior.color:=rgb(217,217,217);
  //ширина ячеек\\
  EX.WorkBooks[1].WorkSheets[6].Columns[1].ColumnWidth := 5;
  EX.WorkBooks[1].WorkSheets[6].Columns[2].ColumnWidth := 22;
  EX.WorkBooks[1].WorkSheets[6].Columns[3].ColumnWidth := 20;
  EX.WorkBooks[1].WorkSheets[6].Columns[4].ColumnWidth := 20;
  EX.WorkBooks[1].WorkSheets[6].Rows[2].RowHeight := 20;
  EX.WorkBooks[1].WorkSheets[6].Rows[13].RowHeight := 20;
  //выравнивание по центру\\
  EX.WorkBooks[1].WorkSheets[6].Cells[3,3].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[6].Range['d3:d7'].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[6].Rows[2].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[6].Rows[13].HorizontalAlignment :=3;
  //заполнение ячеек в шапке таблицы\\
  For f1:=3 to 7 do EX.WorkBooks[1].WorkSheets[6].Cells[f1,4] := '———';
  EX.WorkBooks[1].WorkSheets[6].Cells[2,2] := 'Ключевые ячейки турнира';
  EX.WorkBooks[1].WorkSheets[6].Cells[2,4] := 'Доп. зачёт';
  EX.WorkBooks[1].WorkSheets[6].Cells[3,2] := 'Номер турнира';
  EX.WorkBooks[1].WorkSheets[6].Cells[4,2] := 'Ссылка на гугл таб';
  EX.WorkBooks[1].WorkSheets[6].Cells[5,2] := 'Общая таблица';
  EX.WorkBooks[1].WorkSheets[6].Cells[6,2] := 'Общая таблица_1';
  EX.WorkBooks[1].WorkSheets[6].Cells[7,2] := 'Командный зачёт';
  EX.WorkBooks[1].WorkSheets[6].Cells[8,2] := 'Супермены';
  EX.WorkBooks[1].WorkSheets[6].Cells[9,2] := 'Маньяки';
  EX.WorkBooks[1].WorkSheets[6].Cells[10,2] := 'Гонщики';
  EX.WorkBooks[1].WorkSheets[6].Cells[11,2] := 'Профи';
  EX.WorkBooks[1].WorkSheets[6].Cells[14,2] := 'Текст для форума';
  //жирный шрифт\\
  EX.WorkBooks[1].WorkSheets[6].Rows[2].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[6].Cells[14,2].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[6].Rows[2].Font.size := 14;
  EX.WorkBooks[1].WorkSheets[6].Cells[14,2].Font.size := 14;
  //границы таблицы\\
  EX.WorkBooks[1].WorkSheets[6].Range['b2:b11'].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['c2:c11'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['d2:d11'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b2:d2'].Borders.Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b12:d12'].Borders[3].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b14:b'+inttostr(fm)].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b14:b'+inttostr(fm)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b14:b14'].Borders.Weight := 3;
  EX.WorkBooks[1].WorkSheets[6].Range['b'+inttostr(fm)+':b'+inttostr(fm)].Borders[4].Weight := 3;



  //Награждения\\
  //объединение ячеек\\
  EX.WorkBooks[1].WorkSheets[7].Range['b2:c2'].Merge;
  EX.WorkBooks[1].WorkSheets[7].Range['b11:e11'].Merge;
  //Заливка таблицы\\
  EX.WorkBooks[1].WorkSheets[7].Cells[2,2].interior.color:=rgb(217,217,217);
  EX.WorkBooks[1].WorkSheets[7].Cells[11,2].interior.color:=rgb(217,217,217);
  //ширина ячеек\\
  EX.WorkBooks[1].WorkSheets[7].Columns[1].ColumnWidth := 5;
  EX.WorkBooks[1].WorkSheets[7].Columns[2].ColumnWidth := 18;
  EX.WorkBooks[1].WorkSheets[7].Columns[3].ColumnWidth := 8.5;
  EX.WorkBooks[1].WorkSheets[7].Columns[4].ColumnWidth := 2;
  EX.WorkBooks[1].WorkSheets[7].Rows[2].RowHeight := 20;
  EX.WorkBooks[1].WorkSheets[7].Rows[11].RowHeight := 20;
  //выравнивание по центру\\
  EX.WorkBooks[1].WorkSheets[7].Columns[3].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[7].Rows[2].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[7].Rows[11].HorizontalAlignment :=3;
  EX.WorkBooks[1].WorkSheets[7].Cells[2,2].VerticalAlignment :=2;
  EX.WorkBooks[1].WorkSheets[7].Cells[11,2].VerticalAlignment :=2;
  //заполнение ячеек в шапке таблицы\\
  EX.WorkBooks[1].WorkSheets[7].Cells[2,2] := 'Турнир';
  EX.WorkBooks[1].WorkSheets[7].Cells[3,2] := 'Номер турнира';
  EX.WorkBooks[1].WorkSheets[7].Cells[4,2] := 'ID турнира';
  EX.WorkBooks[1].WorkSheets[7].Cells[5,2] := 'Номер поста';
  EX.WorkBooks[1].WorkSheets[7].Cells[8,2] := '="За участие в [МиГоМан №"&C3&"](http://klavogonki.ru/forum/events/"&C4&"/page1/#post"&C5&"). Поздравляю!"';
  EX.WorkBooks[1].WorkSheets[7].Cells[11,2] := 'Награждение';
  EX.WorkBooks[1].WorkSheets[7].Cells[5,3] := '2';
  //жирный шрифт\\
  EX.WorkBooks[1].WorkSheets[7].Cells[2,2].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[7].Cells[11,2].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[7].Cells[2,2].Font.size := 14;
  EX.WorkBooks[1].WorkSheets[7].Cells[11,2].Font.size := 14;
  //границы таблицы\\
  EX.WorkBooks[1].WorkSheets[7].Range['b8:b8'].Borders.Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['b2:b5'].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['c2:c5'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['b2:c2'].Borders.Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['b5:c5'].Borders[4].Weight := 3;

  Progress.Position:=Progress.Position+10;

  //Сортировка топа\\
  Okno.Clear; Win.Clear;
  For f1:=8 to 11 do begin n[f1]:=4; mesto[f1]:=0; mt[f1]:=10000 end;
  For f1:=1 to w do begin
    f:=1;
    For f2:=2 to w do
      If strtoint(top[f][3])<strtoint(top[f2][3]) then f:=f2 else If strtoint(top[f][3])=strtoint(top[f2][3]) then
        If strtoint(top[f][7])<strtoint(top[f2][7]) then f:=f2 else If strtoint(top[f][7])=strtoint(top[f2][7]) then
          If strtoint(top[f][4])<strtoint(top[f2][4]) then f:=f2 else If strtoint(top[f][4])=strtoint(top[f2][4]) then
            If strtoint(top[f][5])<strtoint(top[f2][5]) then f:=f2 else If strtoint(top[f][5])=strtoint(top[f2][5]) then
              If strtoint(top[f][6])<strtoint(top[f2][6]) then f:=f2;

    Case top[f][8][1] of
      '7' : stran:=8;
      '6' : stran:=9;
      '5' : stran:=10;
      '4' : stran:=11;
    end;

    If top[f][3]<>'0' then begin
      If mt[stran]>strtoint(top[f][3]) then mesto[stran]:=mesto[stran]+1;
      mt[stran]:=strtoint(top[f][3]);
      EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],15]:= mesto[stran];
      If mesto[stran]<=10 then
        EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],15].font.color:=clWhite;
      For f2:=16 to 21 do If top[f][f2-14]<>'0' then begin
        EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],f2]:= top[f][f2-14];
        If f2=16 then EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],f2].Hyperlinks.add(EX.WorkBooks[1].WorkSheets[stran].Cells[n[stran],f2],Address:='http://klavogonki.ru/u/#/'+top[f][1]);
      end;
      Case stran of
        8  : cvetop:=2900;
        9  : cvetop:=2700;
        10 : cvetop:=2200;
        11 : cvetop:=1500;
      end;
      For f3:=1 to 5 do If cvetop+(f3*100)>strtoint(top[f][3]) then begin cvetop:=f3; break end;
      Cvet;
      EX.WorkBooks[1].WorkSheets[stran].Range['p'+inttostr(n[stran])+':q'+inttostr(n[stran])].interior.color:=rgb(ct1,ct2,ct3);
      If cvetop>5 then EX.WorkBooks[1].WorkSheets[stran].Range['p'+inttostr(n[stran])+':q'+inttostr(n[stran])].font.color:=clwhite;
      For f4:=1 to 8 do Okno.SelText:=top[f][f4]+'	';

      n[stran]:=n[stran]+1;
    end;
    top[f][3]:='0';
    Okno.Lines.Add('');
  end;

  t.LoadFromFile(papka+'\save.txt');
  t.Strings[0]:=inttostr(strtoint(t.Strings[0])+1);
  EX.WorkBooks[1].WorkSheets[6].Cells[3,3] := t.Strings[0];
  EX.WorkBooks[1].WorkSheets[7].Cells[3,3] := t.Strings[0];


  For f1:=8 to 11 do begin
    EX.WorkBooks[1].WorkSheets[f1].Cells[4,3]:=troj[f1][1];
    Case f1 of
      8  : cvetop:=2900;
      9  : cvetop:=2700;
      10 : cvetop:=2200;
      11 : cvetop:=1500;
    end;
    If troj[f1][2]<>'' then
      For f3:=1 to 5 do If cvetop+(f3*100)>strtoint(troj[f1][2]) then begin cvetop:=f3; break end;
    If cvetop>5 then cvetop:=6;
    If troj[f1][2]<>'' then begin
      Case f1 of
        8  : t.Strings[cvetop]:=inttostr(strtoint(t.Strings[cvetop])+1);
        9  : t.Strings[6+cvetop]:=inttostr(strtoint(t.Strings[6+cvetop])+1);
        10 : t.Strings[12+cvetop]:=inttostr(strtoint(t.Strings[12+cvetop])+1);
        11 : t.Strings[18+cvetop]:=inttostr(strtoint(t.Strings[18+cvetop])+1);
      end;
      If cvetop>5 then EX.WorkBooks[1].WorkSheets[f1].Cells[4,3].font.color:=clwhite;
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,3+cvetop]:=troj[f1][2];
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,10]:=troj[f1][3];
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,11]:=troj[f1][4];
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,12]:=troj[f1][5];
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,13]:=troj[f1][6];
      Cvet;
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,2].font.color:=clblue;
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,3].interior.color:=rgb(ct1,ct2,ct3);
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,3+cvetop].interior.color:=rgb(ct1,ct2,ct3);
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,3].Hyperlinks.add(EX.WorkBooks[1].WorkSheets[f1].Cells[4,3],Address:='http://klavogonki.ru/u/#/'+troj[f1][7]);
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,10].Hyperlinks.add(EX.WorkBooks[1].WorkSheets[f1].Cells[4,10],Address:='http://klavogonki.ru/u/#/'+troj[f1][8]);
      EX.WorkBooks[1].WorkSheets[f1].Cells[4,12].Hyperlinks.add(EX.WorkBooks[1].WorkSheets[f1].Cells[4,12],Address:='http://klavogonki.ru/u/#/'+troj[f1][9]);
    end;
  end;

  Progress.Position:=Progress.Position+10;

  //Таблица топа\\
  F3:=1;
  For stran:=8 to 11 do begin
   //Окрашиваем ячейки с результатами\\
   Case stran of
      8  : begin f4:=2800; cv1:=217; cv2:=183; cv3:=255 end;
      9  : begin f4:=2600; cv1:=251; cv2:=197; cv3:=247 end;
      10 : begin f4:=2100; cv1:=255; cv2:=215; cv3:=87 end;
      11 : begin f4:=1400; cv1:=255; cv2:=255; cv3:=201 end;
    end;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:m3'].interior.color:=rgb(cv1,cv2,cv3);
    For f1:=1 to 6 do begin
      cvetop:=f1; Cvet;
      EX.WorkBooks[1].WorkSheets[stran].Cells[2,f1+3] := f4+(f1*100);
      EX.WorkBooks[1].WorkSheets[stran].Cells[2,f1+3].interior.color:=rgb(ct1,ct2,ct3);
    end;
    //Объединение ячеек\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:b3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['c2:c3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['j2:k3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['l2:m3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['o2:o3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['p2:p3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['q2:q3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['r2:r3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['s2:s3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['t2:t3'].Merge;
    EX.WorkBooks[1].WorkSheets[stran].Range['u2:u3'].Merge;
    //Заливка таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['o2:u3'].interior.color:=rgb(242,242,242);
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,18].interior.color:=rgb(252,248,166);
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,19].interior.color:=rgb(218,219,222);
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,20].interior.color:=rgb(255,222,189);
    //ширина ячеек\\
    EX.WorkBooks[1].WorkSheets[stran].Columns[1].ColumnWidth := 4;
    EX.WorkBooks[1].WorkSheets[stran].Columns[2].ColumnWidth :=23;
    EX.WorkBooks[1].WorkSheets[stran].Columns[3].ColumnWidth := 16;
    For f1:=4 to 9 do EX.WorkBooks[1].WorkSheets[stran].Columns[f1].ColumnWidth := 7;
    EX.WorkBooks[1].WorkSheets[stran].Columns[10].ColumnWidth := 16;
    EX.WorkBooks[1].WorkSheets[stran].Columns[11].ColumnWidth := 7;
    EX.WorkBooks[1].WorkSheets[stran].Columns[12].ColumnWidth := 16;
    EX.WorkBooks[1].WorkSheets[stran].Columns[13].ColumnWidth := 7;
    EX.WorkBooks[1].WorkSheets[stran].Columns[14].ColumnWidth := 4;
    EX.WorkBooks[1].WorkSheets[stran].Columns[15].ColumnWidth := 6;
    EX.WorkBooks[1].WorkSheets[stran].Columns[16].ColumnWidth := 16;
    EX.WorkBooks[1].WorkSheets[stran].Columns[17].ColumnWidth := 10;
    For f1:=18 to 20 do EX.WorkBooks[1].WorkSheets[stran].Columns[f1].ColumnWidth := 7;
    EX.WorkBooks[1].WorkSheets[stran].Columns[21].ColumnWidth := 10;
    //выравнивание по центру\\
    EX.WorkBooks[1].WorkSheets[stran].Range['c2:u'+inttostr(n[stran])].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,2].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:u'+inttostr(n[stran])].VerticalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Columns[3].HorizontalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,3].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Columns[10].HorizontalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,10].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Columns[12].HorizontalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,12].HorizontalAlignment :=3;
    EX.WorkBooks[1].WorkSheets[stran].Columns[16].HorizontalAlignment :=2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,16].HorizontalAlignment :=3;
    //жирный шрифт\\
    EX.WorkBooks[1].WorkSheets[stran].Columns[15].Font.Bold := true;
    EX.WorkBooks[1].WorkSheets[stran].Columns[17].Font.Bold := true;
    EX.WorkBooks[1].WorkSheets[stran].Range['r2:u3'].Font.Bold := true;
    //заполнение ячеек в шапке таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:u3'].Wraptext:=true;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,2].font.size := 9;
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,2] := 'Турнир';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,3] := 'Победитель';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,10] := '2 место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,12] := '3 место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,16] := 'Игрок';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,17] := 'Лучший результат';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,18] := '1  место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,19] := '2  место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,20] := '3  место';
    EX.WorkBooks[1].WorkSheets[stran].Cells[2,21] := 'Всего медалей';
    If troj [stran][1]<>'' then
      EX.WorkBooks[1].WorkSheets[stran].Cells[1,1] := '="МиГоМан "&Лист7!C3&" ['+d+'.'+m+'.'+g+']"';
    //границы таблицы\\
    EX.WorkBooks[1].WorkSheets[stran].Range['b2:c3'].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['j2:m3'].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['d2:i2'].Borders[3].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['d2:i2'].Borders[4].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['d3:i3'].Borders[4].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,2].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,2].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,3].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,10].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,12].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,13].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,21].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,21].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,10].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Cells[4,12].Borders[2].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['o2:q3'].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['o2:u2'].Borders[3].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['r3:u3'].Borders[4].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['o2:o'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['p2:p'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['q2:q'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['r2:r'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['u2:u'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['v2:v'+inttostr(n[stran]-1)].Borders[1].Weight := 2;
    EX.WorkBooks[1].WorkSheets[stran].Range['o'+inttostr(n[stran])+':u'+inttostr(n[stran])].Borders[3].Weight := 2;
    //Вывод файла save.txt в эксель\\
    For f2:=4 to 9 do begin If t.Strings[f3]<>'0' then
      EX.WorkBooks[1].WorkSheets[stran].Cells[3,f2]:=t.Strings[f3];
      f3:=f3+1
    end;
  end;
  Win.Lines.Text:=t.Text;


  //создаем вкладки\\
  EX.WorkBooks[1].WorkSheets[1].Name := 'Общая таблица';
  EX.WorkBooks[1].WorkSheets[2].Name := 'Супермены';
  EX.WorkBooks[1].WorkSheets[3].Name := 'Маньяки';
  EX.WorkBooks[1].WorkSheets[4].Name := 'Гонщики';
  EX.WorkBooks[1].WorkSheets[5].Name := 'Профи';
  EX.WorkBooks[1].WorkSheets[6].Name := 'Форум';
  EX.WorkBooks[1].WorkSheets[7].Name := 'Награждение';
  EX.WorkBooks[1].WorkSheets[8].Name := 'Топ_суп';
  EX.WorkBooks[1].WorkSheets[9].Name := 'Топ_ман';
  EX.WorkBooks[1].WorkSheets[10].Name := 'Топ_гон';
  EX.WorkBooks[1].WorkSheets[11].Name := 'Топ_про';
  EX.WorkBooks[1].WorkSheets[12].Name := 'БО-допзачёт';
  EX.WorkBooks[1].WorkSheets[13].Name := 'Командный';
  Bezosh; f1:=12;

  If PROFI.Checked = false then begin EX.WorkBooks[1].WorkSheets[11].Delete; f1:=f1-2 end;
  If GONIK.Checked = false then begin EX.WorkBooks[1].WorkSheets[10].Delete; f1:=f1-2 end;
  If MANIK.Checked = false then begin EX.WorkBooks[1].WorkSheets[9].Delete; f1:=f1-2 end;
  If SUPER.Checked = false then begin EX.WorkBooks[1].WorkSheets[8].Delete; f1:=f1-2 end;
  If PROFI.Checked = false then EX.WorkBooks[1].WorkSheets[5].Delete;
  If GONIK.Checked = false then EX.WorkBooks[1].WorkSheets[4].Delete;
  If MANIK.Checked = false then EX.WorkBooks[1].WorkSheets[3].Delete;
  If SUPER.Checked = false then EX.WorkBooks[1].WorkSheets[2].Delete;

  EX.WorkBooks[1].WorkSheets[f1+1].Move (Before:=EX.WorkBooks[1].WorkSheets[2]);
  EX.WorkBooks[1].WorkSheets[f1+1].Move (Before:=EX.WorkBooks[1].WorkSheets[3]);
  EX.WorkBooks[1].WorkSheets[1].select;
  EX.ActiveWindow.TabRatio := 0.850;
  EX.Visible := true;
  EX.WorkBooks[1].saveas(GetCurrentDir+'\'+naz);
  EX := Unassigned;
end;



procedure TForm1.Bezosh;
var
  f,f1,f2,f3,f4,f5,f6,xy1,xy2,m : integer;
  k : array[1..4] of integer;
  uk : array[1..4] of integer;
  nom : array[4..7] of integer;
  prov : boolean;
begin
  For f1:=1 to 4 do k[f1]:=0;

  //ширина ячеек\\
  EX.WorkBooks[1].WorkSheets[12].Columns[1].ColumnWidth := 3;
  EX.WorkBooks[1].WorkSheets[12].Columns[2].ColumnWidth := 7;
  EX.WorkBooks[1].WorkSheets[12].Columns[3].ColumnWidth := 6;
  EX.WorkBooks[1].WorkSheets[12].Columns[4].ColumnWidth := 18;
  EX.WorkBooks[1].WorkSheets[12].Columns[5].ColumnWidth := 8;
  for f1:=6 to 10 do EX.WorkBooks[1].WorkSheets[12].Columns[f1].ColumnWidth := 6;
  //выравнивание по центру\\
  For f1:=2 to 10 do EX.WorkBooks[1].WorkSheets[12].Columns[f1].HorizontalAlignment :=3;
  For f1:=2 to 10 do EX.WorkBooks[1].WorkSheets[12].Columns[f1].VerticalAlignment :=2;
  EX.WorkBooks[1].WorkSheets[12].Columns[4].HorizontalAlignment :=2;
  //жирный шрифт\\
  EX.WorkBooks[1].WorkSheets[12].Columns[3].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[12].Columns[4].Font.Bold := true;
  EX.WorkBooks[1].WorkSheets[12].Columns[5].Font.Bold := true;

  f3:=3;
  EX.WorkBooks[1].WorkSheets[12].Select;
  If (SUPER.Checked=true)and(bot[7]>0) then begin
    //Копирование ячеек\\
    EX.WorkBooks[1].WorkSheets[12].Range['m1:m'+inttostr(bot[7])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,3].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['n1:n'+inttostr(bot[7])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,5].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,9].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['o1:o'+inttostr(bot[7])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,10].PasteSpecial;

    m:=1500; f5:=0; nom[7]:=f3-1;
    If bot[7]>=5 then f4:=500 else f4:=500-((5-bot[7])*100);
    For f1:=1 to igrok do dan[f1][1][13]:=dan[f1][1][12];
    For f1:=1 to igrok do begin
      f:=1;
      For f2:=2 to igrok do
        If strtoint(dan[f][1][13])<strtoint(dan[f2][1][13]) then f:=f2 else
          If strtoint(dan[f][1][13])=strtoint(dan[f2][1][13]) then
            If strtoint(dan[f][1][11])<strtoint(dan[f2][1][11]) then f:=f2;

      If (dan[f][1][3]='7')and(strtoint(dan[f][1][13])>0)and
         (lsk[f][7]=6)and(dan[f][1][10]='2') then begin

        If f6<>strtoint(dan[f][1][13]) then f5:=f5+1;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,2] := dan[f][1][1];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,3] := f5;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,4] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,5] := dan[f][1][13];
        If dan[f][2][18]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6] := dan[f][2][18]+' '+dan[f][4][18];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][19]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7] := dan[f][2][19]+' '+dan[f][4][19];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][20]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8] := dan[f][2][20]+' '+dan[f][4][20];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If f6<>strtoint(dan[f][1][13]) then begin
          If f4=50 then f4:=0;
          If f4>100 then f4:=f4-100 else If f4=100 then f4:=50;
        end;
        If f4>0 then EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := f4
                else EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := '.';
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,10] := bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],2] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],3] := f4+bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],4] := 'Бонус - Безошибочный';
        y[5]:=y[5]+1;
        f3:=f3+1; k[1]:=k[1]+1; f6:=strtoint(dan[f][1][13]);
      end;
      dan[f][1][13]:='0';
    end;
    f3:=f3+2;
  end;

  If (MANIK.Checked=true)and(bot[6]>0) then begin
    //Копирование ячеек\\
    EX.WorkBooks[1].WorkSheets[12].Range['p1:p'+inttostr(bot[6])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,3].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['q1:q'+inttostr(bot[6])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,5].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,9].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['r1:r'+inttostr(bot[6])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,10].PasteSpecial;

    m:=1500; f5:=0; nom[6]:=f3-1;
    If bot[6]>=5 then f4:=500 else f4:=500-((5-bot[6])*100);
    For f1:=1 to igrok do dan[f1][1][13]:=dan[f1][1][12];
    For f1:=1 to igrok do begin
      f:=1;
      For f2:=2 to igrok do
        If strtoint(dan[f][1][13])<strtoint(dan[f2][1][13]) then f:=f2 else
          If strtoint(dan[f][1][13])=strtoint(dan[f2][1][13]) then
            If strtoint(dan[f][1][11])<strtoint(dan[f2][1][11]) then f:=f2;

      If (dan[f][1][3]='6')and(strtoint(dan[f][1][13])>0)and
         (lsk[f][7]=6)and(dan[f][1][10]='2') then begin
        If f6<>strtoint(dan[f][1][13]) then f5:=f5+1;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,2] := dan[f][1][1];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,3] := f5;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,4] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,5] := dan[f][1][13];
        If dan[f][2][18]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6] := dan[f][2][18]+' '+dan[f][4][18];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][19]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7] := dan[f][2][19]+' '+dan[f][4][19];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][20]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8] := dan[f][2][20]+' '+dan[f][4][20];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If f6<>strtoint(dan[f][1][13]) then begin
          If f4=50 then f4:=0;
          If f4>100 then f4:=f4-100 else If f4=100 then f4:=50;
        end;
        If f4>0 then EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := f4
                else EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := '.';
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,10] := bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],2] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],3] := f4+bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],4] := 'Бонус - Безошибочный';
        y[5]:=y[5]+1;
        f3:=f3+1; k[2]:=k[2]+1; f6:=strtoint(dan[f][1][13]);
      end;
      dan[f][1][13]:='0';
    end;
    f3:=f3+2;
  end;

  If (GONIK.Checked=true)and(bot[5]>0) then begin
    //Копирование ячеек\\
    EX.WorkBooks[1].WorkSheets[12].Range['s1:s'+inttostr(bot[5])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,3].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['t1:t'+inttostr(bot[5])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,5].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,9].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['u1:u'+inttostr(bot[5])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,10].PasteSpecial;

    m:=1500; f5:=0; nom[5]:=f3-1;
    If bot[5]>=5 then f4:=500 else f4:=500-((5-bot[5])*100);
    For f1:=1 to igrok do dan[f1][1][13]:=dan[f1][1][12];
    For f1:=1 to igrok do begin
      f:=1;
      For f2:=2 to igrok do
        If strtoint(dan[f][1][13])<strtoint(dan[f2][1][13]) then f:=f2 else
          If strtoint(dan[f][1][13])=strtoint(dan[f2][1][13]) then
            If strtoint(dan[f][1][11])<strtoint(dan[f2][1][11]) then f:=f2;

      If (dan[f][1][3]='5')and(strtoint(dan[f][1][13])>0)and
         (lsk[f][7]=6)and(dan[f][1][10]='2') then begin
        If f6<>strtoint(dan[f][1][13]) then f5:=f5+1;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,2] := dan[f][1][1];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,3] := f5;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,4] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,5] := dan[f][1][13];
        If dan[f][2][18]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6] := dan[f][2][18]+' '+dan[f][4][18];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][19]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7] := dan[f][2][19]+' '+dan[f][4][19];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][20]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8] := dan[f][2][20]+' '+dan[f][4][20];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If f6<>strtoint(dan[f][1][13]) then begin
          If f4=50 then f4:=0;
          If f4>100 then f4:=f4-100 else If f4=100 then f4:=50;
        end;
        If f4>0 then EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := f4
                else EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := '.';
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,10] := bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],2] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],3] := f4+bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],4] := 'Бонус - Безошибочный';
        y[5]:=y[5]+1;
        f3:=f3+1; k[3]:=k[3]+1; f6:=strtoint(dan[f][1][13]);
      end;
      dan[f][1][13]:='0';
    end;
    f3:=f3+2;
  end;

  If (PROFI.Checked=true)and(bot[4]>0) then begin
    //Копирование ячеек\\
    EX.WorkBooks[1].WorkSheets[12].Range['v1:v'+inttostr(bot[4])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,3].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['w1:w'+inttostr(bot[4])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,5].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,9].PasteSpecial;
    EX.WorkBooks[1].WorkSheets[12].Range['x1:x'+inttostr(bot[4])].Copy;
    EX.WorkBooks[1].WorkSheets[12].Cells[f3,10].PasteSpecial;

    m:=1500; f5:=0; nom[4]:=f3-1;
    If bot[4]>=5 then f4:=500 else f4:=500-((5-bot[4])*100);
    For f1:=1 to igrok do dan[f1][1][13]:=dan[f1][1][12];
    For f1:=1 to igrok do begin
      f:=1;
      For f2:=2 to igrok do
        If strtoint(dan[f][1][13])<strtoint(dan[f2][1][13]) then f:=f2 else
          If strtoint(dan[f][1][13])=strtoint(dan[f2][1][13]) then
            If strtoint(dan[f][1][11])<strtoint(dan[f2][1][11]) then f:=f2;

      If (dan[f][1][3]='4')and(strtoint(dan[f][1][13])>0)and
         (lsk[f][7]=6)and(dan[f][1][10]='2') then begin
        If f6<>strtoint(dan[f][1][13]) then f5:=f5+1;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,2] := dan[f][1][1];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,3] := f5;
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,4] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,5] := dan[f][1][13];
        If dan[f][2][18]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6] := dan[f][2][18]+' '+dan[f][4][18];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,6].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][19]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7] := dan[f][2][19]+' '+dan[f][4][19];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,7].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If dan[f][2][20]<>'0' then begin
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8] := dan[f][2][20]+' '+dan[f][4][20];
          EX.WorkBooks[1].WorkSheets[12].Cells[f3,8].Characters(Start:=5, Length:=1).Font.Superscript := True;
        end;
        If f6<>strtoint(dan[f][1][13]) then begin
          If f4=50 then f4:=0;
          If f4>100 then f4:=f4-100 else If f4=100 then f4:=50;
        end;
        If f4>0 then EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := f4
                else EX.WorkBooks[1].WorkSheets[12].Cells[f3,9] := '.';
        EX.WorkBooks[1].WorkSheets[12].Cells[f3,10] := bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],2] := dan[f][1][2];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],3] := f4+bo[f][4];
        EX.WorkBooks[1].WorkSheets[7].Cells[y[5],4] := 'Бонус - Безошибочный';
        y[5]:=y[5]+1;
        f3:=f3+1; k[4]:=k[4]+1; f6:=strtoint(dan[f][1][13]);
      end;
      dan[f][1][13]:='0';
    end;
    f3:=f3+1;
  end;


  For f1:=7 downto 4 do If bot[f1]>0 then begin
    //Заливка таблицы\\
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1])+':j'+inttostr(nom[f1])].interior.color:=rgb(131,174,255);
    //выравнивание по центру\\
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],4].HorizontalAlignment :=3;
    //жирный шрифт\\
    EX.WorkBooks[1].WorkSheets[12].Rows[nom[f1]].Font.Bold := true;
    //заполнение ячеек в шапке таблицы\\
    EX.WorkBooks[1].WorkSheets[12].Rows[nom[f1]].font.size := 9;
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],5].Wraptext:=true;
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],9].Wraptext:=true;
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],10].Wraptext:=true;
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],2] := 'ID';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],3] := 'Место';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],4] := 'Ник';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],5] := 'Лучшая скорость';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],6] := '1';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],7] := '2';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],8] := '3';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],9] := 'Награж дение';
    EX.WorkBooks[1].WorkSheets[12].Cells[nom[f1],10] := 'Бонус ные';
    //границы таблицы\\
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1])+':j'+inttostr(nom[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1])+':b'+inttostr(nom[f1]+bot[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['d'+inttostr(nom[f1])+':d'+inttostr(nom[f1]+bot[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['f'+inttostr(nom[f1])+':f'+inttostr(nom[f1]+bot[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['g'+inttostr(nom[f1])+':g'+inttostr(nom[f1]+bot[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['h'+inttostr(nom[f1])+':h'+inttostr(nom[f1]+bot[f1])].Borders.Weight := 2;
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1])+':j'+inttostr(nom[f1])].Borders[3].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1])+':b'+inttostr(nom[f1]+bot[f1])].Borders[1].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['j'+inttostr(nom[f1])+':j'+inttostr(nom[f1]+bot[f1])].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['b'+inttostr(nom[f1]+bot[f1])+':j'+inttostr(nom[f1]+bot[f1])].Borders[4].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['c'+inttostr(nom[f1])+':c'+inttostr(nom[f1]+bot[f1])].Borders[1].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['c'+inttostr(nom[f1])+':c'+inttostr(nom[f1]+bot[f1])].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['e'+inttostr(nom[f1])+':e'+inttostr(nom[f1]+bot[f1])].Borders[1].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['e'+inttostr(nom[f1])+':e'+inttostr(nom[f1]+bot[f1])].Borders[2].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['i'+inttostr(nom[f1])+':i'+inttostr(nom[f1]+bot[f1])].Borders[1].Weight := 3;
    EX.WorkBooks[1].WorkSheets[12].Range['i'+inttostr(nom[f1])+':i'+inttostr(nom[f1]+bot[f1])].Borders[2].Weight := 3;
  end;

  //границы таблицы зачета рангов\\
  EX.WorkBooks[1].WorkSheets[13].Range['b2:g2'].interior.color:=rgb(206,253,87);
  EX.WorkBooks[1].WorkSheets[13].Range['b2:g2'].Wraptext:=true;
  EX.WorkBooks[1].WorkSheets[13].Cells[2,2]:='Ранг';
  EX.WorkBooks[1].WorkSheets[13].Cells[2,3]:='Количество игроков';
  EX.WorkBooks[1].WorkSheets[13].Cells[2,4]:='% попавших в Бонусный';
  EX.WorkBooks[1].WorkSheets[13].Cells[2,5]:='очки-количество';
  EX.WorkBooks[1].WorkSheets[13].Cells[2,6]:='очки-% в Бонусном';
  EX.WorkBooks[1].WorkSheets[13].Cells[2,7]:='Награж дение';
  EX.WorkBooks[1].WorkSheets[13].Range['b2:b5'].Borders.Weight := 2;
  EX.WorkBooks[1].WorkSheets[13].Range['e2:f5'].Borders.Weight := 2;
  EX.WorkBooks[1].WorkSheets[13].Range['b2:g2'].Borders.Weight := 2;
  EX.WorkBooks[1].WorkSheets[13].Range['b2:g2'].Borders[3].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['b2:b5'].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['g2:g5'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['b5:g5'].Borders[4].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['c2:c5'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['d2:d5'].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[13].Range['f2:f5'].Borders[2].Weight := 3;



  //границы таблицы в награждениях\\
  EX.WorkBooks[1].WorkSheets[7].Range['b11:b'+inttostr(y[5]-1)].Borders[1].Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['e11:e'+inttostr(y[5]-1)].Borders[2].Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['b11:e11'].Borders.Weight := 3;
  EX.WorkBooks[1].WorkSheets[7].Range['b'+inttostr(y[5]-1)+':e'+inttostr(y[5]-1)].Borders[4].Weight := 3;

end;


procedure TForm1.TimerTimer(Sender: TObject);
begin
  if IsConnectedToInternet then begin
    Timer.Enabled:=false;
    START.Enabled:=true;
    START.Caption:='Выбрать первый заезд';
    START.SetFocus;
  end;
end;



procedure TForm1.Cvet;
begin
  Case rang of
    '4' : begin cv1:=255; cv2:=247; cv3:=143; cr:='#8C8100' end;
    '5' : begin cv1:=253; cv2:=199; cv3:=139; cr:='#BA5800' end;
    '6' : begin cv1:=255; cv2:=159; cv3:=182; cr:='#BC0143' end;
    '7' : begin cv1:=219; cv2:=147; cv3:=255; cr:='#5E0B9E' end;
  end;
  Case cvetop of
    1 : begin ct1:=255; ct2:=255; ct3:=155; end;
    2 : begin ct1:=255; ct2:=192; ct3:=0; end;
    3 : begin ct1:=246; ct2:=138; ct3:=238; end;
    4 : begin ct1:=194; ct2:=139; ct3:=255; end;
    5 : begin ct1:=138; ct2:=237; ct3:=242; end
    else begin ct1:=117; ct2:=113; ct3:=113; end;
  end;
end;



procedure TForm1.FormCreate(Sender: TObject);
begin
  papka:=GetCurrentDir;
  If fileexists('set.txt') then begin
    Okno.Lines.LoadFromFile('set.txt');
    IU.Text:=okno.Lines[0];
    If okno.Lines[1]='0' then PROFI.Checked:=false else PROFI.Checked:=true;
    POP0.Position:=strtoint(okno.Lines[2]);
    PPO0.Position:=strtoint(okno.Lines[3]);
    If okno.Lines[4]='0' then GONIK.Checked:=false else GONIK.Checked:=true;
    GOP0.Position:=strtoint(okno.Lines[5]);
    GPO0.Position:=strtoint(okno.Lines[6]);
    If okno.Lines[7]='0' then MANIK.Checked:=false else MANIK.Checked:=true;
    MOP0.Position:=strtoint(okno.Lines[8]);
    MPO0.Position:=strtoint(okno.Lines[9]);
    If okno.Lines[10]='0' then SUPER.Checked:=false else SUPER.Checked:=true;
    SOP0.Position:=strtoint(okno.Lines[11]);
    SPO0.Position:=strtoint(okno.Lines[12]);
    KOS.Text:=okno.Lines[13];
    KOO.Text:=okno.Lines[14];
    KOC.Text:=okno.Lines[15];
    KOB.Text:=okno.Lines[16];
    KOK.Text:=okno.Lines[17];
    KOM.Text:=okno.Lines[18];
    If okno.Lines[19]='0' then ZR.Checked:=false else ZR.Checked:=true;
  end;
end;



procedure TForm1.POPKeyPress(Sender: TObject; var Key: Char);
begin
  Case key of #8,'0','1','2','3','4','5','6','7','8','9' : else key:=char(0); end;
end;

procedure TForm1.KOSKeyPress(Sender: TObject; var Key: Char);
begin
  Case key of #8,',','0','1','2','3','4','5','6','7','8','9' : else key:=char(0); end;
end;

procedure TForm1.PRODOLClick(Sender: TObject);
  var f1,f2 : integer;
begin
  Panel.Visible:=false;
  For f1:=1 to igrok do if dan[f1][1][10]='3' then
    if box[f1].Checked=false then dan[f1][1][10]:='1' else dan[f1][1][10]:='2';
  /////Проверяем установлен ли Excel\\\\\
  if not IsOLEObjectInstalled('Excel.Application') then begin
    Showmessage('На компьютере не установлен MS Excel.'+#10+#13+'Вывод результаров невозможен.');
    exit;
  end else Tablica;
  START.Caption:='Задача выполнена'; START.Enabled:=false;
  SAVE.Visible:=true;
end;




procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  If Save.Visible=true then begin
    CanClose:=MessageBox(Self.Handle, PChar('Сохранить данные топа?'), PChar(''), MB_YESNO + MB_ICONINFORMATION  + MB_APPLMODAL)=IDNO;
  end;
end;



procedure TForm1.Nezach;
var
  f1,f2,f3 : integer;
  v : array [1..5000,1..3] of string;
begin
  t:=tstringlist.Create;
  If fileexists(papka+'\1.txt') then begin
    t.LoadFromFile(papka+'\1.txt');
    Win.Clear;
    For f1:=0 to t.Count-1 do begin
      t.Strings[f1]:=trim(t.Strings[f1])+'	';
      f3:=1;
      For f2:=1 to length(t.Strings[f1]) do
        If t.Strings[f1][f2]<>'	'
          then Win.SelText:=t.Strings[f1][f2]
          else begin
            If Win.Lines.Text<>'' then v[f1+1][f3]:=Win.Lines.Text;
            Win.Clear; f3:=f3+1;
          end;
    end;
  end;

  For f1:=1 to t.Count do
    For f2:=1 to igrok do
      If v[f1][2]=dan[f2][1][1] then begin
         dan[f2][1][3]:=v[f1][1];
         break;
      end;

  Win.Clear;
end;



procedure TForm1.SaveClick(Sender: TObject);
begin
  Okno.Lines.SaveToFile(papka+'/top.txt');
  Okno.Lines.SaveToFile(GetCurrentDir+'/top.txt');
  Win.Lines.SaveToFile(papka+'/save.txt');
  Win.Lines.SaveToFile(GetCurrentDir+'/save.txt');
  Okno.Clear;
    Okno.Lines.Add(IU.Text);
    Okno.Lines.Add('0');
    Okno.Lines.Add(POP.Text);
    Okno.Lines.Add(PPO.Text);
    If GONIK.Checked=false then Okno.Lines.Add('0') else Okno.Lines.Add('1');
    Okno.Lines.Add(GOP.Text);
    Okno.Lines.Add(GPO.Text);
    If MANIK.Checked=false then Okno.Lines.Add('0') else Okno.Lines.Add('1');
    Okno.Lines.Add(MOP.Text);
    Okno.Lines.Add(MPO.Text);
    Okno.Lines.Add('0');
    Okno.Lines.Add(SOP.Text);
    Okno.Lines.Add(SPO.Text);
    Okno.Lines.Add(KOS.Text);
    Okno.Lines.Add(KOO.Text);
    Okno.Lines.Add(KOC.Text);
    Okno.Lines.Add(KOB.Text);
    Okno.Lines.Add(KOK.Text);
    Okno.Lines.Add(KOM.Text);
    If ZR.Checked=false then Okno.Lines.Add('0') else Okno.Lines.Add('1');
    Okno.Lines.SaveToFile(papka+'/set.txt');
  SAVE.Visible:=false;
end;

end.

