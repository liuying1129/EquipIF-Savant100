unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  Menus, StdCtrls, Buttons, ADODB,
  ComCtrls, ToolWin, ExtCtrls,
  inifiles,Dialogs,
  StrUtils, DB, ComObj,Variants,CPort, CoolTrayIcon;

type
  TfrmMain = class(TForm)
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton8: TToolButton;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button1: TButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ComPort1: TComPort;
    ComDataPacket1: TComDataPacket;
    ToolButton7: TToolButton;
    SaveDialog1: TSaveDialog;
    LYTray1: TCoolTrayIcon;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    //增加病人信息表中记录,返回该记录的唯一编号作为检验结果表的外键
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ComDataPacket1Packet(Sender: TObject; const Str: String);
    procedure ToolButton7Click(Sender: TObject);
    procedure ComPort1AfterOpen(Sender: TObject);
  private
    { Private declarations }
    procedure UpdateConfig;{配置文件生效}
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction, PerlRegEx;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  H_DTR_RTS:boolean;//DTR/RTS高电位
  ifRecLog:boolean;//是否记录调试日志
  StartString:String;
  StopString:String;
  RegExSpecNo:String;//匹配联机号的正则
  RegExDlttype:String;//匹配联机标识的正则
  RegExValue:String;//匹配检验结果的正则

//  RFM:STRING;       //返回数据
  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  ConnectString:=GetConnectString;
  UpdateConfig;

  //笨方法替换.todo:通用替换方法
  StartString:=StringReplace(StartString, '$02', #$02, [rfReplaceAll]);
  StartString:=StringReplace(StartString, '$1B', #$1B, [rfReplaceAll]);//热景Hotgen-UPT2800
  StartString:=StringReplace(StartString, '$16', #$16, [rfReplaceAll]);//热景Hotgen-UPT2800
  StopString:=StringReplace(StopString, '$03', #$03, [rfReplaceAll]);
  StopString:=StringReplace(StopString, '$0A', #$0A, [rfReplaceAll]);//热景Hotgen-UPT2800
  StopString:=StringReplace(StopString, '$1B', #$1B, [rfReplaceAll]);//热景Hotgen-UPT2800
  StopString:=StringReplace(StopString, '$01', #$01, [rfReplaceAll]);//热景Hotgen-UPT2800
  StopString:=StringReplace(StopString, '$05', #$05, [rfReplaceAll]);//热景Hotgen-UPT2800

  ComDataPacket1.StartString:=StartString;//变量StartString在UpdateConfig中赋值,故该代码在UpdateConfig之后
  ComDataPacket1.StopString:=StopString;//变量StopString在UpdateConfig中赋值,故该代码在UpdateConfig之后
  
  if ifRegister then bRegister:=true else bRegister:=false;

  Caption:='数据接收服务'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=caNone;
  LYTray1.HideMainForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  if (MessageDlg('退出后将不再接收设备数据,确定退出吗？', mtWarning, [mbYes, mbNo], 0) <> mrYes) then exit;
  application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  LYTray1.ShowMainForm;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  CommName,BaudRate,DataBit,StopBit,ParityBit:string;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  CommName:=ini.ReadString(IniSection,'串口选择','COM1');
  BaudRate:=ini.ReadString(IniSection,'波特率','9600');
  DataBit:=ini.ReadString(IniSection,'数据位','8');
  StopBit:=ini.ReadString(IniSection,'停止位','1');
  ParityBit:=ini.ReadString(IniSection,'校验位','None');
  H_DTR_RTS:=ini.readBool(IniSection,'DTR/RTS高电位',false);
  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);
  StartString:=ini.ReadString(IniSection,'StartString','');
  if StartString='' then StartString:='$02';
  StopString:=ini.ReadString(IniSection,'StopString','');
  if StopString='' then StopString:='$03';
  RegExSpecNo:=ini.ReadString(IniSection,'匹配联机号的正则','');
  RegExDlttype:=ini.ReadString(IniSection,'匹配联机标识的正则','');
  RegExValue:=ini.ReadString(IniSection,'匹配检验结果的正则','');

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);
  ComPort1.Close;
  ComPort1.Port:=CommName;
  if BaudRate='1200' then
    ComPort1.BaudRate:=br1200
    else if BaudRate='2400' then
      ComPort1.BaudRate:=br2400
    else if BaudRate='4800' then
      ComPort1.BaudRate:=br4800
      else if BaudRate='9600' then
        ComPort1.BaudRate:=br9600
        else if BaudRate='19200' then
          ComPort1.BaudRate:=br19200
        else if BaudRate='57600' then
          ComPort1.BaudRate:=br57600
          else ComPort1.BaudRate:=br9600;
  if DataBit='5' then
    ComPort1.DataBits:=dbFive
    else if DataBit='6' then
      ComPort1.DataBits:=dbSix
      else if DataBit='7' then
        ComPort1.DataBits:=dbSeven
        else if DataBit='8' then
          ComPort1.DataBits:=dbEight
          else ComPort1.DataBits:=dbEight;
  if StopBit='1' then
    ComPort1.StopBits:=sbOneStopBit
    else if StopBit='2' then
      ComPort1.StopBits:=sbTwoStopBits
      else if StopBit='1.5' then
        ComPort1.StopBits:=sbOne5StopBits
        else ComPort1.StopBits:=sbOneStopBit;
  if ParityBit='None' then
    ComPort1.Parity.Bits:=prNone
    else if ParityBit='Odd' then
      ComPort1.Parity.Bits:=prOdd
      else if ParityBit='Even' then
        ComPort1.Parity.Bits:=prEven
        else if ParityBit='Mark' then
          ComPort1.Parity.Bits:=prMark
          else if ParityBit='Space' then
            ComPort1.Parity.Bits:=prSpace
            else ComPort1.Parity.Bits:=prNone;
  try
    ComPort1.Open;
  except
    showmessage('串口'+ComPort1.Port+'打开失败!');
  end;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
  lsComPort:TStrings;
  sComPort:String;
begin
  //获取串口列表 begin
  lsComPort := TStringList.Create;
  EnumComPorts(lsComPort);
  sComPort:=lsComPort.Text;
  lsComPort.Free;
  //获取串口列表 end

    ss:='串口选择'+#2+'Combobox'+#2+sComPort+#2+'0'+#2+#2+#3+
      '波特率'+#2+'Combobox'+#2+'57600'+#13+'19200'+#13+'9600'+#13+'4800'+#13+'2400'+#13+'1200'+#2+'0'+#2+#2+#3+
      '数据位'+#2+'Combobox'+#2+'8'+#13+'7'+#13+'6'+#13+'5'+#2+'0'+#2+#2+#3+
      '停止位'+#2+'Combobox'+#2+'1'+#13+'1.5'+#13+'2'+#2+'0'+#2+#2+#3+
      '校验位'+#2+'Combobox'+#2+'None'+#13+'Even'+#13+'Odd'+#13+'Mark'+#13+'Space'+#2+'0'+#2+#2+#3+
      'DTR/RTS高电位'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      'StartString'+#2+'Edit'+#2+#2+'1'+#2+'16进制必须2位.如Begin $02'+#2+#3+
      'StopString'+#2+'Edit'+#2+#2+'1'+#2+'16进制必须2位.如End $03'+#2+#3+
      '匹配联机号的正则'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '匹配联机标识的正则'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '匹配检验结果的正则'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Memo1.Lines.Clear;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  SaveDialog1.DefaultExt := '.txt';
  SaveDialog1.Filter := 'txt (*.txt)|*.txt';
  if not SaveDialog1.Execute then exit;
  memo1.Lines.SaveToFile(SaveDialog1.FileName);
  showmessage('保存成功!');
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
  ls:Tstrings;
  ss:String;
begin
  OpenDialog1.DefaultExt := '.txt';
  OpenDialog1.Filter := 'txt (*.txt)|*.txt';
  if not OpenDialog1.Execute then exit;
  ls:=Tstringlist.Create;
  ls.LoadFromFile(OpenDialog1.FileName);
  ss:=ls.Text;
  {ss:=#$02#$0A'20170113001'#$0A'25-OH-D         '#$0A'9.15  '#$0A'ng/ml'#$0A#$03;//华科泰-savant-100测试数据
  ss:=#$1B6#$1BC#$16+
      '2025-02-25 11:01'#$0A+
      #$1C'&项'#$1B'f'#$00#$02'目'#$1C'.:CCP'#$0A+
      #$1C'&样品号'#$1C'.:2502250002'#$0A+
      #$1C'&结'#$1B'f'#$00#$02'果'#$1C'.:466.541ng/mL'#$0A+
      #$1B'f'#$01#$05;//热景Hotgen-UPT2800测试数据}
  ComDataPacket1Packet(nil,ss);
  ls.Free;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.ComDataPacket1Packet(Sender: TObject;
  const Str: String);
VAR
  SpecNo:string;
  dlttype:String;
  sValue:String;
  FInts:OleVariant;
  ReceiveItemInfo:OleVariant;
  PerlRegEx:TPerlRegEx;
  ifMatch:Boolean;
begin
  if length(memo1.Lines.Text)>=60000 then memo1.Lines.Clear;//memo只能接受64K个字符
  memo1.Lines.Add(Str);

  //获得联机号 begin
  PerlRegEx:=TPerlRegEx.Create;
  PerlRegEx.RegEx:=RegExSpecNo;
  //PerlRegEx.Options:=PerlRegEx.Options+[preUnGreedy];//正则表达式中控制贪婪模式,以便更好的灵活性
  PerlRegEx.Subject:=Str;
  ifMatch:=False;//初始化
  Try
    ifMatch:=PerlRegEx.Match;//正则表达式为空、语法不正确，Match方法会抛出异常
  except
    on E:Exception do
    begin
      memo1.Lines.Add('匹配联机号报错:'+E.Message);
    end;
  end;
  if ifMatch then
  begin
    SpecNo:=PerlRegEx.MatchedText;//Groups[0]与MatchedText功能一样
    //GroupCount为捕获组数量
    //Groups[1] 第一个捕获组匹配的文本
    //Groups[2] 第二个捕获组匹配的文本，以此类推
    if PerlRegEx.GroupCount>0 then SpecNo:=PerlRegEx.Groups[1];//支持捕获组匹配.如使用捕获组,获取结果一定是Groups[1]
    SpecNo:=RightStr('0000'+trim(SpecNo),4);
  end;
  FreeAndNil(PerlRegEx);
  //获得联机号 end

  //获得联机标识 begin
  PerlRegEx:=TPerlRegEx.Create;
  PerlRegEx.RegEx:=RegExDlttype;
  //PerlRegEx.Options:=PerlRegEx.Options+[preUnGreedy];//正则表达式中控制贪婪模式,以便更好的灵活性
  PerlRegEx.Subject:=Str;
  ifMatch:=False;//初始化
  Try
    ifMatch:=PerlRegEx.Match;//正则表达式为空、语法不正确，Match方法会抛出异常
  except
    on E:Exception do
    begin
      memo1.Lines.Add('匹配联机标识报错:'+E.Message);
    end;
  end;
  if ifMatch then
  begin
    dlttype:=PerlRegEx.MatchedText;//Groups[0]与MatchedText功能一样
    if PerlRegEx.GroupCount>0 then dlttype:=PerlRegEx.Groups[1];//支持捕获组匹配.如使用捕获组,获取结果一定是Groups[1]
  end;
  FreeAndNil(PerlRegEx);
  //获得联机标识 end

  //获得检验结果 begin
  PerlRegEx:=TPerlRegEx.Create;
  PerlRegEx.RegEx:=RegExValue;
  //PerlRegEx.Options:=PerlRegEx.Options+[preUnGreedy];//正则表达式中控制贪婪模式.因为获取检验结果有时需要贪婪模式
  PerlRegEx.Subject:=Str;
  ifMatch:=False;//初始化
  Try
    ifMatch:=PerlRegEx.Match;//正则表达式为空、语法不正确，Match方法会抛出异常
  except
    on E:Exception do
    begin
      memo1.Lines.Add('匹配检验结果报错:'+E.Message);
    end;
  end;
  if ifMatch then
  begin
    sValue:=PerlRegEx.MatchedText;//Groups[0]与MatchedText功能一样
    if PerlRegEx.GroupCount>0 then sValue:=PerlRegEx.Groups[1];//支持捕获组匹配.如使用捕获组,获取结果一定是Groups[1]
    sValue:=StringReplace(sValue,'RU/ml','',[rfReplaceAll, rfIgnoreCase]);//热景Hotgen-UPT2800
    sValue:=StringReplace(sValue,'ng/mL','',[rfReplaceAll, rfIgnoreCase]);//热景Hotgen-UPT2800
    sValue:=StringReplace(sValue,'pg/ml','',[rfReplaceAll, rfIgnoreCase]);//热景Hotgen-UPT2800
    sValue:=trim(sValue);
  end;
  FreeAndNil(PerlRegEx);
  //获得检验结果 end

  ReceiveItemInfo:=VarArrayCreate([0,0],varVariant);
  ReceiveItemInfo[0]:=VarArrayof([dlttype,sValue,'','']);

  if bRegister then
  begin
    FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
    FInts.fData2Lis(ReceiveItemInfo,(SpecNo),'',
      (GroupName),(SpecType),(SpecStatus),(EquipChar),
      (CombinID),'',(LisFormCaption),(ConnectString),
      (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
      ifRecLog,true,'常规',
      '',
      -1,
      '','','','',
      -1,-1,-1,-1,
      -1,-1,-1,-1,
      false,false,false,false);
    if not VarIsEmpty(FInts) then FInts:= unAssigned;
  end;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.ComPort1AfterOpen(Sender: TObject);
begin
  if H_DTR_RTS then
  begin
    ComPort1.SetDTR(true);
    ComPort1.SetRTS(true);
  end;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
