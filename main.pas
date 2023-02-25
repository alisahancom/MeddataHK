unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, dxBarBuiltInMenu, cxGraphics,
  cxControls, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore, dxSkinBasic,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkroom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinOffice2016Colorful,
  dxSkinOffice2016Dark, dxSkinOffice2019Black, dxSkinOffice2019Colorful,
  dxSkinOffice2019DarkGray, dxSkinOffice2019White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringtime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinTheBezier, dxSkinsDefaultPainters, dxSkinValentine,
  dxSkinVisualStudio2013Blue, dxSkinVisualStudio2013Dark,
  dxSkinVisualStudio2013Light, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxContainer, cxEdit, dxGDIPlusClasses, cxImage, cxLabel,
  cxPC, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls, dxCore, cxDateUtils, cxTextEdit,
  cxMaskEdit, cxDropDownEdit, cxCalendar, Vcl.StdCtrls, cxButtons, cxClasses,
  dxSkinsForm, cxGroupBox, Winapi.ShellAPI, Data.DB, MemDS, DBAccess, Ora,
  OraCall, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxNavigator,
  dxDateRanges, dxScrollbarAnnotations, cxDBData, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGridLevel, cxGridCustomView, cxGrid,
  Vcl.Clipbrd, System.Win.Registry, cxCurrencyEdit;

type
  TForm1 = class(TForm)
    cxPageControl1: TcxPageControl;
    Panel1: TPanel;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    Panel2: TPanel;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxImage1: TcxImage;
    cxImage2: TcxImage;
    cxDateEdit1: TcxDateEdit;
    cxDateEdit2: TcxDateEdit;
    cxLabel3: TcxLabel;
    cxLabel4: TcxLabel;
    dxSkinController1: TdxSkinController;
    cxLabel5: TcxLabel;
    cxLabel6: TcxLabel;
    cxButton2: TcxButton;
    OraSession1: TOraSession;
    qr_gss: TOraQuery;
    ds_gss: TDataSource;
    qr_gssADI: TStringField;
    qr_gssSOYADI: TStringField;
    qr_gssDOSYA_NO: TFloatField;
    qr_gssPROTOKOL_NO: TFloatField;
    qr_gssBOLUM_ADI: TStringField;
    qr_gssADI_SOYADI: TStringField;
    qr_gssKURUM_ADI: TStringField;
    qr_gssGSS_TAKIP_NO: TStringField;
    qr_gssGSS_BASVURU_NO: TStringField;
    qr_gssKULLANICI_ACAN: TStringField;
    qr_gssKULLANICI_DEGISTIREN: TStringField;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1ADI: TcxGridDBColumn;
    cxGrid1DBTableView1SOYADI: TcxGridDBColumn;
    cxGrid1DBTableView1DOSYA_NO: TcxGridDBColumn;
    cxGrid1DBTableView1PROTOKOL_NO: TcxGridDBColumn;
    cxGrid1DBTableView1BOLUM_ADI: TcxGridDBColumn;
    cxGrid1DBTableView1ADI_SOYADI: TcxGridDBColumn;
    cxGrid1DBTableView1KURUM_ADI: TcxGridDBColumn;
    cxGrid1DBTableView1GSS_TAKIP_NO: TcxGridDBColumn;
    cxGrid1DBTableView1GSS_BASVURU_NO: TcxGridDBColumn;
    cxGrid1DBTableView1KULLANICI_ACAN: TcxGridDBColumn;
    cxGrid1DBTableView1KULLANICI_DEGISTIREN: TcxGridDBColumn;
    PopupMenu1: TPopupMenu;
    DosyaKopyala1: TMenuItem;
    KopyalaProtokol1: TMenuItem;
    cxLabel7: TcxLabel;
    cxLabel8: TcxLabel;
    qr_fatura: TOraQuery;
    ds_fatura: TDataSource;
    cxGrid2: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxGridLevel1: TcxGridLevel;
    cxGridDBTableView1ADI: TcxGridDBColumn;
    cxGridDBTableView1SOYADI: TcxGridDBColumn;
    cxGridDBTableView1DOSYA_NO: TcxGridDBColumn;
    cxGridDBTableView1PROTOKOL_NO: TcxGridDBColumn;
    cxGridDBTableView1BOLUM_ADI: TcxGridDBColumn;
    cxGridDBTableView1ADI_SOYADI: TcxGridDBColumn;
    cxGridDBTableView1KURUM_ADI: TcxGridDBColumn;
    cxGridDBTableView1FATURATIPI: TcxGridDBColumn;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    cxGroupBox1: TcxGroupBox;
    Image1: TImage;
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    edt_server: TEdit;
    edt_pass: TEdit;
    edt_username: TcxComboBox;
    btnloginbuton: TcxButton;
    btn1: TcxButton;
    cxTabSheet3: TcxTabSheet;
    ds_rehber: TDataSource;
    qr_rehber: TOraQuery;
    qr_rehberAD: TStringField;
    qr_rehberGOREV: TStringField;
    qr_rehberKAT: TStringField;
    qr_rehberDAHILI: TStringField;
    qr_rehberCEP: TStringField;
    qr_rehberMAIL: TStringField;
    qr_rehberTUR: TStringField;
    cxGrid3DBTableView1: TcxGridDBTableView;
    cxGrid3Level1: TcxGridLevel;
    cxGrid3: TcxGrid;
    cxGrid3DBTableView1AD: TcxGridDBColumn;
    cxGrid3DBTableView1GOREV: TcxGridDBColumn;
    cxGrid3DBTableView1KAT: TcxGridDBColumn;
    cxGrid3DBTableView1DAHILI: TcxGridDBColumn;
    cxGrid3DBTableView1CEP: TcxGridDBColumn;
    cxGrid3DBTableView1MAIL: TcxGridDBColumn;
    cxGrid3DBTableView1TUR: TcxGridDBColumn;
    qr_faturaADI: TStringField;
    qr_faturaSOYADI: TStringField;
    qr_faturaDOSYA_NO: TFloatField;
    qr_faturaPROTOKOL_NO: TFloatField;
    qr_faturaBOLUM_ADI: TStringField;
    qr_faturaADI_SOYADI: TStringField;
    qr_faturaKURUM_ADI: TStringField;
    qr_faturaFATURATIPI: TStringField;
    qr_faturaCIRO: TFloatField;
    cxGridDBTableView1CIRO: TcxGridDBColumn;
    qr_faturaODETMEALAN: TStringField;
    qr_faturaEKLEYEN: TStringField;
    cxGridDBTableView1ODETMEALAN: TcxGridDBColumn;
    cxGridDBTableView1EKLEYEN: TcxGridDBColumn;
    cxTabSheet4: TcxTabSheet;
    ds_afet: TDataSource;
    qr_afet: TOraQuery;
    cxGrid4: TcxGrid;
    cxGridDBTableView2: TcxGridDBTableView;
    cxGridLevel2: TcxGridLevel;
    cxGridDBTableView2TARIH: TcxGridDBColumn;
    cxGridDBTableView2DOSYA_NO: TcxGridDBColumn;
    cxGridDBTableView2ADI_SOYADI: TcxGridDBColumn;
    cxGridDBTableView2PROTOKOL_NO: TcxGridDBColumn;
    cxGridDBTableView2BOLUM: TcxGridDBColumn;
    cxGridDBTableView2BOLUM_ADI: TcxGridDBColumn;
    cxGridDBTableView2DR_KODU: TcxGridDBColumn;
    cxGridDBTableView2DOKTOR: TcxGridDBColumn;
    cxGridDBTableView2PROVIZYON: TcxGridDBColumn;
    cxGridDBTableView2TEDAVI: TcxGridDBColumn;
    qr_afetTARIH: TStringField;
    qr_afetDOSYA_NO: TFloatField;
    qr_afetADI_SOYADI: TStringField;
    qr_afetPROTOKOL_NO: TFloatField;
    qr_afetBOLUM: TFloatField;
    qr_afetBOLUM_ADI: TStringField;
    qr_afetDR_KODU: TFloatField;
    qr_afetDOKTOR: TStringField;
    qr_afetPROVIZYON: TStringField;
    qr_afetTEDAVI: TStringField;
    procedure cxLabel1Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure cxImage1Click(Sender: TObject);
    procedure cxLabel2Click(Sender: TObject);
    procedure cxImage2Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure DosyaKopyala1Click(Sender: TObject);
    procedure KopyalaProtokol1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btn1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.cxButton2Click(Sender: TObject);
begin
  //GSS Sorgusu
  qr_gss.ParamByName('tar1').Value := cxDateEdit1.Text + ' 00:00:00';
  qr_gss.ParamByName('tar2').Value := cxDateEdit2.Text + ' 23:59:59';
  qr_gss.Execute;
  cxGrid1DBTableView1.ApplyBestFit();

  //fatura Sorgusu
  qr_fatura.ParamByName('tar1').Value := cxDateEdit1.Text + ' 00:00:00';
  qr_fatura.ParamByName('tar2').Value := cxDateEdit2.Text + ' 23:59:59';
  qr_fatura.Execute;
  cxGridDBTableView1.ApplyBestFit();

  //Afet Sorgusu
  qr_afet.ParamByName('tar1').Value := cxDateEdit1.Text + ' 00:00:00';
  qr_afet.ParamByName('tar2').Value := cxDateEdit2.Text + ' 23:59:59';
  qr_afet.Execute;
  cxGridDBTableView2.ApplyBestFit();
end;

procedure TForm1.cxImage1Click(Sender: TObject);
begin
  ShellExecute(self.WindowHandle, 'open',
    PChar('https://www.instagram.com/alisahancom/'), nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.cxImage2Click(Sender: TObject);
begin
  ShellExecute(self.WindowHandle, 'open', PChar('https://www.alisahan.com'),
    nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.cxLabel1Click(Sender: TObject);
begin
  ShellExecute(self.WindowHandle, 'open',
    PChar('https://www.instagram.com/alisahancom/'), nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.cxLabel2Click(Sender: TObject);
begin
  ShellExecute(self.WindowHandle, 'open', PChar('https://www.alisahan.com'),
    nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.DosyaKopyala1Click(Sender: TObject);
begin
  Clipboard.AsText := cxGrid1DBTableView1DOSYA_NO.EditValue;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Application.Terminate;
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  cxLabel5.Left := Form1.Width - 70;
  cxLabel5.Top := Panel2.Top - 200 - 28;
  cxLabel6.Left := Form1.Width - 40;
  cxLabel6.Top := Panel2.Top - 200;

  cxLabel7.Left := Form1.Width - 70;
  cxLabel7.Top := Panel2.Top - 200 - 28;
  cxLabel8.Left := Form1.Width - 50;
  cxLabel8.Top := Panel2.Top - 200;

  cxGroupBox1.Top := round(Form1.Height / 2 - 100);
  cxGroupBox1.Left := round(Form1.Width / 2 - 200);
end;

procedure TForm1.FormShow(Sender: TObject);

begin

  edt_server.Text := 'orcl';
  edt_username.Text := 'meddatapacs';
  edt_pass.Text := 'meddatapacs';
  //
//  edt_server.Text := 'xe';
//  edt_username.Text := 'hastane';
//  edt_pass.Text := 'hastane';

  try
    OraSession1.Server := edt_server.Text;
    OraSession1.Username := edt_username.Text;
    OraSession1.Password := edt_pass.Text;
    OraSession1.Open;
    qr_gss.ParamByName('tar1').Value := cxDateEdit1.Text + ' 00:00:00';
    qr_gss.ParamByName('tar2').Value := cxDateEdit2.Text + ' 23:59:59';
    qr_gss.Execute;
    cxGrid1DBTableView1.ApplyBestFit();

    qr_fatura.ParamByName('tar1').Value := cxDateEdit1.Text + ' 00:00:00';
    qr_fatura.ParamByName('tar2').Value := cxDateEdit2.Text + ' 23:59:59';
    qr_fatura.Execute;
    cxGridDBTableView1.ApplyBestFit();

    qr_rehber.Open;
    cxGrid3DBTableView1.ApplyBestFit();
  except
    on E: Exception do
      ShowMessage('Hatalý Giriþ Problemi, Olasý Neden ' + ''#13'' + E.Message);
  end;
  // Panel1.Enabled := false;
  // cxPageControl1.Enabled := false;

end;

procedure TForm1.KopyalaProtokol1Click(Sender: TObject);
begin
  Clipboard.AsText := cxGrid1DBTableView1PROTOKOL_NO.EditValue;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin
  Clipboard.AsText := cxGridDBTableView1DOSYA_NO.EditValue;
end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
  Clipboard.AsText := cxGridDBTableView1PROTOKOL_NO.EditValue;
end;

end.
