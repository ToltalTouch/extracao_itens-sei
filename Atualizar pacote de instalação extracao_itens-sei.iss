[Setup]
AppName= ExtraÁ„o Itens Almoxarifado - SEI
AppVersion=1.0.0
DefaultDirName={localappdata}\ExtraÁ„o Itens Almoxarifado - SEI
DefaultGroupName=ExtraÁ„o Itens Almoxarifado - SEI
OutputBaseFilename=Extracao_Itens_Almoxarifado-SEIN_Setup
// N√ÉO MEXA NESSAS OP√á√ïES //////////////////
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
/////////////////////////////////////////////
[Files]
Source: "C:\python\extracao_itens-sei\dist\extracao_itens-sei\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
// MANTENHA O PADR√ÉO "\*" NO FINAL DO CAMINHO 
// PARA INDICAR QUE VAI SER UTILIZADO OS ARQUIVOS 
// DAQUELA PASTA EM ESPECIFICO
/////////////////////////////////////////////

// IDIOMA DO INSTALADOR, PADR√ÉO PORTUGUES ///
[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
/////////////////////////////////////////////
// DECLARAR ATALHOS DO PROGRAMA ((PADR√ÉO))
[Icons]
Name: "{userdesktop}\ExtraÁ„o Itens Almoxarifado - SEI"; Filename: "{app}\extracao_itens-sei.exe"   
Name: "{userdesktop}\Desinstalar ExtraÁ„o Itens Almoxarifado - SEI"; Filename: "{uninstallexe}"

// ATALHOS OPCIONAIS DEPENDENDO DO PROGRAMA, COMO PASTA DE DOWNLOAD OU PLANILHAS EXCEL
Name: "{userdesktop}\Planilha Extraidos"; Filename: "{app}\_internal\excel\itens_extraidos.xlsx"
Name: "{userdesktop}\Manual ExtraÁ„o Itens Almoxarifado - SEI"; Filename: "{app}\_internal\README.pdf"
/////////////////////////////////////////////
// CODIGO PADR√ÉO APENAS MUDE A FUN√á√ÉO DO PROGRAMA
[Code]
var
  CustomPage: TWizardPage;
  TermsPage: TInputOptionWizardPage;
  MemoText: TMemo; // Declara a vari√°vel MemoText

procedure InitializeWizard;
begin
  // Cria uma p√°gina personalizada
  CustomPage := CreateCustomPage(wpWelcome, 'InformaÁıes Importantes', 'Leia as informaÁıes abaixo antes de continuar.');

  // Adiciona um memo na p√°gina de informa√ß√µes
  MemoText := TMemo.Create(CustomPage);
  MemoText.Parent := CustomPage.Surface;
  MemoText.ScrollBars := ssVertical; // Adiciona uma barra de rolagem vertical
  MemoText.ReadOnly := True; // Torna o texto somente leitura
  MemoText.WordWrap := True; // Habilita quebra de linha autom√°tica
  MemoText.TabStop := False; // Impede que o controle receba foco ao pressionar Tab
  MemoText.Text := 'Este programa n„o armazena nenhuma informaÁ„o como Login, Senha ou dados sensÌveis respeitando a LGPD' +
    ' e foi desenvolvido para automatizar ExtraÁ„o Itens Almoxarifado - SEI, ' +
    'permitindo verificar o Processos presnete no SEI e gerar rel·rio no formato de Log. Ao utilizar este software, vocÍ concorda com os seguintes termos:' + #13#10#13#10 +
    '1. Uso Respons·vel: O programa deve ser utilizado apenas para fins legais e autorizados. O usu·rio È respons·vel por garantir que possui permiss„o ExtraÁ„o Itens Almoxarifado - SEI.' + #13#10#13#10 +
    '2. LimitaÁ„o de Responsabilidade: O desenvolvedor n„o se responsabiliza por quaisquer danos, perdas ou consequÍncias decorrentes do uso do programa, incluindo erros no sistema que resultam falhas na automaÁ„o.' + #13#10#13#10 +
    '3. Privacidade e SeguranÁa: O programa utiliza credenciais fornecidas pelo usu·rio para realizar login no sistema. O usu·rio È respons·vel por proteger suas credenciais e garantir que elas n„o sejam compartilhadas ou utilizadas de forma inadequada.';
    
  MemoText.Top := 10;
  MemoText.Left := 10;
  MemoText.Width := CustomPage.SurfaceWidth - 20;
  MemoText.Height := CustomPage.SurfaceHeight - 20;

  // Cria uma nova p√°gina para os Termos de Uso
  TermsPage := CreateInputOptionPage(CustomPage.ID,
    'Termos de Uso',
    'Leia e aceite os Termos de Uso para continuar',
    'VocÍ deve aceitar os Termos de Uso para instalar o programa.',
    True, False);

  // Adiciona a caixa de sele√ß√£o para aceitar os Termos de Uso
  TermsPage.Add('Eu li e aceito os Termos de Uso.');

  // Define a caixa de sele√ß√£o como desmarcada por padr√£o
  TermsPage.Values[0] := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  // Verifica se o usu√°rio est√° na p√°gina de Termos de Uso
  if CurPageID = TermsPage.ID then
  begin
    // Se a caixa de sele√ß√£o n√£o estiver marcada, exibe uma mensagem de erro
    if not TermsPage.Values[0] then
    begin
      MsgBox('Para instalar esse programa, È necess·rio aceitar os Termos de Uso.', mbError, MB_OK);
      Result := False; // Impede que o usu√°rio avance
      Exit;
    end;
  end;

  Result := True; // Permite que o usu√°rio avance
end;

// O READEME.pdf DEVE SER ESPECIFICAMENTE DO PROGRAMA NO FORMATO PDF
[Run]
Filename: "{app}\_internal\README.pdf"; Description: "Abrir instru√ß√µes do programa"; Flags: postinstall shellexec skipifsilent
