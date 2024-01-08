unit uImportadorIquine;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.SqlExpr, Vcl.StdCtrls,
  Data.DBXFirebird, Data.FMTBcd, DBXDevartOracle, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, Data.Win.ADODB, FireDAC.Comp.Client,
  Data.DBXMSSQL, System.DateUtils, Data.DBXCommon, cxControls, cxProgressBar,
  Controls, uboIquine, uboTintaMCP, uboTintometrico, uVsClientDataSet;

type
  TImportadorIquine = class(TObject)
    constructor Create(AMemo: TMemo; UpdateProgress, StartProgress: TProgressBarUpdate);

  private
    { Private declarations }
    memLog: TMemo;
    FIquine: TTintaIquine;
    FMCP: TTintaMCP;
    ProgressBarEventStart, ProgressBarEventUpdate: TProgressBarUpdate;
    procedure AddMsg(AMsg: string);
    function SecondToTime(Segundos: Cardinal): string;
  public
    { Public declarations }
    procedure MigrarSuvinilColecao;
    procedure MigrarSuvinilCorante;
    procedure MigrarSuvinilProduto;
    procedure MigrarSuvinilMsgTinta;
    procedure MigrarSuvinilEmbTinta;
    procedure MigrarSuvinilGrupoTinta;
    procedure MigrarSuvinilBase;
    procedure MigrarSuvinilCorTinta;
    procedure MigrarSuvinilFormulaTinta;
    procedure MigrarSuvinilFormulaCorante;
    procedure MigrarSuvinilItemBase;
    procedure MigrarSuvinilItemCorante;
    procedure MigrarSuvinil;
  end;

implementation

uses
  uboBaseTinta, uboFormulaTinta, uDataSetHelper, uboColecao,
  uboCorTinta, uboCorante, uboGrupoTinta, uboItemCorante, uboProdutoTinta,
  uboEmbTinta, uboMsgTinta, uboFormulaCorante, uboItemBase, uDmConexao, BibNum;

{ TImportadorIquine }

procedure TImportadorIquine.AddMsg(AMsg: string);
begin
  memLog.Lines.Add(DateTimeToStr(Now) + ': ' + AMsg);
end;

constructor TImportadorIquine.Create(AMemo: TMemo; UpdateProgress, StartProgress: TProgressBarUpdate);
begin
  inherited Create;

  if not Assigned(memLog) then
    memLog := TMemo.Create(Nil);

  memLog := AMemo;

  ProgressBarEventStart := StartProgress;
  ProgressBarEventUpdate := UpdateProgress;

  FIquine := TTintaIquine.Create(UpdateProgress, StartProgress);
  FMCP := TTintaMCP.Create(UpdateProgress, StartProgress);

end;

function TImportadorIquine.SecondToTime(Segundos: Cardinal): string;
var
  Seg, Min, Hora: Cardinal;
begin
  Hora := Segundos div 3600;
  Seg := Segundos mod 3600;
  Min := Seg div 60;
  Seg := Seg mod 60;
  Result := FormatFloat(',00', Hora) + ':' + FormatFloat('00', Min) + ':' + FormatFloat('00', Seg);
end;

procedure TImportadorIquine.MigrarSuvinil;
begin
  try
    AddMsg('Iniciada Atualização Tinta Suvinil');

    MigrarSuvinilColecao;
    MigrarSuvinilMsgTinta;
    MigrarSuvinilCorante;
    MigrarSuvinilGrupoTinta;
    MigrarSuvinilProduto;
    MigrarSuvinilEmbTinta;
    MigrarSuvinilBase;
    MigrarSuvinilCorTinta;
    MigrarSuvinilFormulaTinta;
    MigrarSuvinilFormulaCorante;

    AddMsg('Finalizada Atualização Tinta Suvinil');
  except
    raise;
  end;
end;

procedure TImportadorIquine.MigrarSuvinilBase;
var
  oListBaseSuvinil, oListBaseMcp: TListBaseTinta;
  oBaseSuvinil, oBaseMcp: TBaseTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oBaseMcp.IdColecao  := oBaseSuvinil.IdColecao;
    oBaseMcp.Descricao  := oBaseSuvinil.Descricao;
    oBaseMcp.IdBase     := oBaseSuvinil.IdBase;
    oBaseMcp.IdEmbTinta := oBaseSuvinil.IdEmbTinta;
    oBaseMcp.IdProduto  := oBaseSuvinil.IdProduto;
    oBaseMcp.Inativo    := oBaseSuvinil.Inativo;
    oBaseMcp.IdOrigem   := oBaseSuvinil.IdOrigem;
  end;

begin
  oListBaseSuvinil := TListBaseTinta.Create;
  oListBaseMcp := TListBaseTinta.Create;

  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento Base SQL Server');
    dNow := Now;
    oListBaseSuvinil := FSuvinil.CarregarBaseTinta(DmConexao.sqSuvinil);
    AddMsg('Carregou Base SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento Base MCP');
    dNow := Now;
    oListBaseMcp := FMCP.CarregarBaseTinta(DmConexao.sqMCP);
    AddMsg('Carregou Base MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela Base');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListBaseSuvinil.Count);

    for oBaseSuvinil in oListBaseSuvinil do
    begin
      oBaseMcp := oListBaseMcp.FindKey(oBaseSuvinil.IdColecao, oBaseSuvinil.IdBase, oBaseSuvinil.IdEmbTinta, oBaseSuvinil.IdProduto);
      if (not Assigned(oBaseMcp)) then
      begin
        oBaseMcp := TBaseTinta.Create;
        oListBaseMcp.Add(oBaseMcp);
        oBaseMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oBaseMcp.Descricao = oBaseSuvinil.Descricao) then
      begin
        oBaseMcp.Status := dsIgual;
      end
      else
      begin
        oBaseMcp.Status := dsAltera;
        CopiarDados;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListBaseSuvinil.IndexOf(oBaseSuvinil));
    end;
    AddMsg('Tempo (Base(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela Base');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListBaseMcp.Count);

    for oBaseMcp in oListBaseMcp do
    begin
      if not (oBaseMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oBaseSuvinil := oListBaseSuvinil.FindKey(oBaseMcp.IdColecao, oBaseMcp.IdBase, oBaseMcp.IdEmbTinta, oBaseMcp.IdProduto);
        if (not Assigned(oBaseSuvinil)) then
        begin
          oBaseMcp.Status := dsDeleta;
        end;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListBaseMcp.IndexOf(oBaseMcp));
    end;
    AddMsg('Tempo (Base(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListBaseSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela Base');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListBaseMcp.Count);
      for oBaseMcp in oListBaseMcp do
      begin
        cSql := oBaseMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListBaseMcp.IndexOf(oBaseMcp));
      end;
      AddMsg('Tempo (Antes commit base): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit base): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit base): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar base: ' + E.Message + ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListBaseSuvinil);
    FreeAndNil(oListBaseMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilColecao;
var
  oListColecaoSuvinil, oListColecaoMcp: TListColecao;
  oColecaoSuvinil, oColecaoMcp: TColecao;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;
begin
  //Não fazer nada, pois era executado comando na mão, para identificar a versão...
  oListColecaoSuvinil := TListColecao.Create;
  oListColecaoMcp := TListColecao.Create;

  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento Coleção SQL Server');
    dNow := Now;
    oListColecaoSuvinil := FSuvinil.CarregarColecao(DmConexao.sqSuvinil);
    AddMsg('Carregou Coleção SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento Coleção MCP');
    dNow := Now;
    oListColecaoMcp := FMCP.CarregarColecao(DmConexao.sqMCP);
    AddMsg('Carregou Coleção MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela Coleção');
    dNow := Now;

    for oColecaoSuvinil in oListColecaoSuvinil do
    begin
      oColecaoMcp := oListColecaoMcp.FindKey(oColecaoSuvinil.IdColecao);
      if (not Assigned(oColecaoMcp)) then
      begin
        oColecaoMcp := TColecao.Create;
        oListColecaoMcp.Add(oColecaoMcp);
        oColecaoMcp.Status := dsInsere;

        oColecaoMcp.IdColecao := oColecaoSuvinil.IdColecao;
        oColecaoMcp.Versao := oColecaoSuvinil.Versao;
      end
      else if (oColecaoMcp.Versao = oColecaoSuvinil.Versao) then
      begin
        oColecaoMcp.Status := dsIgual;
      end
      {else if (oColecaoMcp.Versao > oColecaoSuvinil.Versao) then
      begin
        AddMsg('Versão no Construshow é maior que da Suvinil, atualização cancelada');
        AddMsg('Versão no Construshow: ' + oColecaoMcp.Versao + ' Versão Suvinil: ' + oColecaoSuvinil.Versao);
        Abort;
      end }
      else
      begin
        oColecaoMcp.Status := dsAltera;

        oColecaoMcp.IdColecao := oColecaoSuvinil.IdColecao;
        oColecaoMcp.Descricao := oColecaoSuvinil.Descricao;
        oColecaoMcp.Versao    := oColecaoSuvinil.Versao;
      end;
    end;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela Coleção');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      for oColecaoMcp in oListColecaoMcp do
      begin
        cSql := oColecaoMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

      end;
      AddMsg('Tempo (Antes commit Coleção): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit Coleção): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit Coleção): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar Coleção: ' + E.Message + ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListColecaoSuvinil);
    FreeAndNil(oListColecaoMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilCorante;
var
  oListCoranteSuvinil, oListCoranteMcp: TListCorante;
  oCoranteSuvinil, oCoranteMcp: TCorante;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oCoranteMcp.IdColecao := oCoranteSuvinil.IdColecao;
    oCoranteMcp.Descricao := oCoranteSuvinil.Descricao;
    oCoranteMcp.IdCorante := oCoranteSuvinil.IdCorante;
    oCoranteMcp.Sigla     := oCoranteSuvinil.Sigla;
    oCoranteMcp.Cor       := oCoranteSuvinil.Cor;
    oCoranteMcp.Inativo   := oCoranteSuvinil.Inativo;
    oCoranteMcp.IdOrigem  := oCoranteSuvinil.IdOrigem;
    oCoranteMcp.Densidade := oCoranteSuvinil.Densidade;
  end;

begin
  oListCoranteSuvinil := TListCorante.Create;
  oListCoranteMcp := TListCorante.Create;
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento Corante SQL Server');
    dNow := Now;
    oListCoranteSuvinil := FSuvinil.CarregarCorantes(DmConexao.sqSuvinil);
    AddMsg('Carregou Corante SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento Corante MCP');
    dNow := Now;
    oListCoranteMcp := FMCP.CarregarCorantes(DmConexao.sqMCP);
    AddMsg('Carregou Corante MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela Corante');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListCoranteSuvinil.Count);

    for oCoranteSuvinil in oListCoranteSuvinil do
    begin
      oCoranteMcp := oListCoranteMcp.FindKey(oCoranteSuvinil.IdColecao, oCoranteSuvinil.IdCorante);
      if (not Assigned(oCoranteMcp)) then
      begin
        oCoranteMcp := TCorante.Create;
        oListCoranteMcp.Add(oCoranteMcp);
        oCoranteMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oCoranteMcp.Descricao = oCoranteSuvinil.Descricao) then
      begin
        oCoranteMcp.Status := dsIgual;
      end
      else
      begin
        oCoranteMcp.Status := dsAltera;
        CopiarDados;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListCoranteSuvinil.IndexOf(oCoranteSuvinil));
    end;
    AddMsg('Tempo (Corantes alterados): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela Corante');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListCoranteMcp.Count);

    for oCoranteMcp in oListCoranteMcp do
    begin
      if not (oCoranteMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oCoranteSuvinil := oListCoranteSuvinil.FindKey(oCoranteMcp.IdColecao, oCoranteMcp.IdCorante);
        if (not Assigned(oCoranteSuvinil)) then
        begin
          oCoranteMcp.Status := dsDeleta;
        end;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListCoranteMcp.IndexOf(oCoranteMcp));
    end;
    AddMsg('Tempo (Corantes excluidos): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListCoranteSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela Corante');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListCoranteMcp.Count);
      for oCoranteMcp in oListCoranteMcp do
      begin
        cSql := oCoranteMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListCoranteMcp.IndexOf(oCoranteMcp));
      end;
      AddMsg('Tempo (Antes commit corante): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit corante): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit corante): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar corantes: ' + E.Message + ' ' + cSql);
        raise;
      end;
    end;
  finally
    FreeAndNil(oListCoranteSuvinil);
    FreeAndNil(oListCoranteMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilCorTinta;
begin
  //Não fazer nada, pois era executado comando na mão, para identificar a versão...
end;

procedure TImportadorIquine.MigrarSuvinilEmbTinta;
var
  oListEmbTintaSuvinil, oListEmbTintaMcp: TListEmbTinta;
  oEmbTintaSuvinil, oEmbTintaMcp: TEmbTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oEmbTintaMcp.IdColecao  := oEmbTintaSuvinil.IdColecao;
    oEmbTintaMcp.Descricao  := oEmbTintaSuvinil.Descricao;
    oEmbTintaMcp.IdEmbTinta := oEmbTintaSuvinil.IdEmbTinta;
    oEmbTintaMcp.Capacidade := oEmbTintaSuvinil.Capacidade;
    oEmbTintaMcp.IdOrigem   := oEmbTintaSuvinil.IdOrigem;
  end;

begin
  oListEmbTintaSuvinil := TListEmbTinta.Create;
  oListEmbTintaMcp := TListEmbTinta.Create;
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento EmbTinta SQL Server');
    dNow := Now;
    oListEmbTintaSuvinil := FSuvinil.CarregarEmbTinta(DmConexao.sqSuvinil);
    AddMsg('Carregou EmbTinta SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento EmbTinta MCP');
    dNow := Now;
    oListEmbTintaMcp := FMCP.CarregarEmbTinta(DmConexao.sqMCP);
    AddMsg('Carregou EmbTinta MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela EmbTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
     ProgressbarEventStart(oListEmbTintaSuvinil.Count);
    for oEmbTintaSuvinil in oListEmbTintaSuvinil do
    begin
      oEmbTintaMcp := oListEmbTintaMcp.FindKey(oEmbTintaSuvinil.IdColecao, oEmbTintaSuvinil.IdEmbTinta);
      if (not Assigned(oEmbTintaMcp)) then
      begin
        oEmbTintaMcp := TEmbTinta.Create;
        oListEmbTintaMcp.Add(oEmbTintaMcp);
        oEmbTintaMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oEmbTintaMcp.Descricao = oEmbTintaSuvinil.Descricao) then
      begin
        oEmbTintaMcp.Status := dsIgual;
      end
      else
      begin
        oEmbTintaMcp.Status := dsAltera;
        CopiarDados;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListEmbTintaSuvinil.IndexOf(oEmbTintaSuvinil));
    end;
    AddMsg('Tempo (EmbTinta(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela EmbTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListEmbTintaMcp.Count);

    for oEmbTintaMcp in oListEmbTintaMcp do
    begin
      if not (oEmbTintaMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oEmbTintaSuvinil := oListEmbTintaSuvinil.FindKey(oEmbTintaMcp.IdColecao, oEmbTintaMcp.IdEmbTinta);
        if (not Assigned(oEmbTintaSuvinil)) then
        begin
          oEmbTintaMcp.Status := dsDeleta;
        end;

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListEmbTintaMcp.IndexOf(oEmbTintaMcp));
      end;
    end;
    AddMsg('Tempo (EmbTinta(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListEmbTintaSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela EmbTinta');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListEmbTintaMcp.Count);

      for oEmbTintaMcp in oListEmbTintaMcp do
      begin
        cSql := oEmbTintaMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListEmbTintaMcp.IndexOf(oEmbTintaMcp));
      end;
      AddMsg('Tempo (Antes commit EmbTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit EmbTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit EmbTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar EmbTinta: ' + E.Message+ ' ' + cSql);
        raise;
      end;
    end;
  finally
    FreeAndNil(oListEmbTintaSuvinil);
    FreeAndNil(oListEmbTintaMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilFormulaCorante;
var
  oListFormulaCoranteSuvinil, oListFormulaCoranteMcp: TListFormulaCorante;
  oFormulaCoranteSuvinil, oFormulaCoranteMcp: TFormulaCorante;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;
  cdsIndex: TVsClientDataSet;
  nCount, nRet: Integer;
  MyClass: TComponent;
  lFormulaCoranteVazia : Boolean;
  qrMcp : TSqlQuery;

  procedure CopiarDados;
  begin
    oFormulaCoranteMcp.IdColecao     := oFormulaCoranteSuvinil.IdColecao;
    oFormulaCoranteMcp.IdFormula     := oFormulaCoranteSuvinil.IdFormula;
    oFormulaCoranteMcp.IdBase        := oFormulaCoranteSuvinil.IdBase;
    oFormulaCoranteMcp.IdEmbTinta    := oFormulaCoranteSuvinil.IdEmbTinta;
    oFormulaCoranteMcp.IdGrupoTinta  := oFormulaCoranteSuvinil.IdGrupoTinta;
    oFormulaCoranteMcp.IdOrigem      := oFormulaCoranteSuvinil.IdOrigem;
    oFormulaCoranteMcp.IdCorante     := oFormulaCoranteSuvinil.IdCorante;
    oFormulaCoranteMcp.IdColecaoPers := oFormulaCoranteSuvinil.IdColecaoPers;
    oFormulaCoranteMcp.IdProduto     := oFormulaCoranteSuvinil.IdProduto;
    oFormulaCoranteMcp.QtdeoZ1       := oFormulaCoranteSuvinil.QtdeoZ1;
    oFormulaCoranteMcp.QtdeoZ2       := oFormulaCoranteSuvinil.QtdeoZ2;
    oFormulaCoranteMcp.QtdeMl        := oFormulaCoranteSuvinil.QtdeMl;
    oFormulaCoranteMcp.QtdeGr        := oFormulaCoranteSuvinil.QtdeGr;
    oFormulaCoranteMcp.QtdeGotas     := oFormulaCoranteSuvinil.QtdeGotas;
    oFormulaCoranteMcp.SeqCorante    := oFormulaCoranteSuvinil.SeqCorante;

  end;

begin
  nCount := 0;
  oListFormulaCoranteSuvinil := TListFormulaCorante.Create;
  oListFormulaCoranteMcp := TListFormulaCorante.Create;
  cdsIndex := TVsClientDataSet.Create(nil);
  cdsIndex.FieldDefs.Add('Indice', ftInteger);
  cdsIndex.FieldDefs.Add('IdColecao', ftInteger);
  cdsIndex.FieldDefs.Add('IdColecaoPers', ftInteger);
  cdsIndex.FieldDefs.Add('IdFormula', ftInteger);
  cdsIndex.FieldDefs.Add('IdBase', ftInteger);
  cdsIndex.FieldDefs.Add('IdProduto', ftInteger);
  cdsIndex.FieldDefs.Add('IdEmbTinta', ftInteger);
  cdsIndex.FieldDefs.Add('IdGrupoTinta', ftInteger);
  cdsIndex.FieldDefs.Add('IdCorante', ftInteger);
  cdsIndex.FieldDefs.Add('SeqCorante', ftInteger);
  cdsIndex.CreateDataSet;
  cdsIndex.LogChanges := False;
  cdsIndex.DisableControls;
  cdsIndex.IndexFieldNames := 'IdColecao;IdColecaoPers;IdFormula;IdBase;IdProduto;IdEmbTinta;IdGrupoTinta;IdCorante;SeqCorante';
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento FormulaCorante SQL Server');
    dNow := Now;
    oListFormulaCoranteSuvinil := FSuvinil.CarregarFormulaCorante(DmConexao.sqSuvinil, cdsIndex);
    AddMsg('Carregou Formula Corante SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento FormulaCorante MCP');
    dNow := Now;
    oListFormulaCoranteMcp := FMCP.CarregarFormulaCorante(DmConexao.sqMCP);
    AddMsg('Carregou FormulaCorante MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela FormulaCorante');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListFormulaCoranteSuvinil.Count);

    lFormulaCoranteVazia := oListFormulaCoranteMcp.IsEmpty;
    for oFormulaCoranteSuvinil in oListFormulaCoranteSuvinil do
    begin
      oFormulaCoranteMcp := nil;
      Inc(nCount);
      if (not lFormulaCoranteVazia) and (cdsIndex.FindKey([oFormulaCoranteSuvinil.IdColecao, oFormulaCoranteSuvinil.IdColecaoPers, oFormulaCoranteSuvinil.IdFormula, oFormulaCoranteSuvinil.IdBase, oFormulaCoranteSuvinil.IdProduto, oFormulaCoranteSuvinil.IdEmbTinta, oFormulaCoranteSuvinil.IdGrupoTinta, oFormulaCoranteSuvinil.IdCorante, oFormulaCoranteSuvinil.SeqCorante])) then
        oFormulaCoranteMcp := oListFormulaCoranteMcp[cdsIndex.AsInt('Indice')];
      //oFormulaCoranteMcp := oListFormulaCoranteMcp.FindKey(oFormulaCoranteSuvinil.IdColecao, oFormulaCoranteSuvinil.IdColecaoPers, oFormulaCoranteSuvinil.IdFormula, oFormulaCoranteSuvinil.IdBase, oFormulaCoranteSuvinil.IdProduto, oFormulaCoranteSuvinil.IdEmbTinta, oFormulaCoranteSuvinil.IdGrupoTinta, oFormulaCoranteSuvinil.IdCorante, oFormulaCoranteSuvinil.SeqCorante);
      if (not Assigned(oFormulaCoranteMcp)) then
      begin
        oFormulaCoranteMcp := TFormulaCorante.Create;
        oListFormulaCoranteMcp.Add(oFormulaCoranteMcp);
        oFormulaCoranteMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oFormulaCoranteMcp.IdOrigem = oFormulaCoranteSuvinil.IdOrigem) and (oFormulaCoranteMcp.QtdeMl = oFormulaCoranteSuvinil.QtdeMl)
          and (oFormulaCoranteMcp.QtdeGotas = oFormulaCoranteSuvinil.QtdeGotas) then
      begin
        oFormulaCoranteMcp.Status := dsIgual;
      end
      else
      begin
        oFormulaCoranteMcp.Status := dsAltera;
        CopiarDados;
      end;

      if nCount mod 500 = 1 then
        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListFormulaCoranteSuvinil.IndexOf(oFormulaCoranteSuvinil));
    end;
    AddMsg('Tempo (FormulaCorante(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela FormulaCorante');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListFormulaCoranteMcp.Count);

    nCount := 0;
    for oFormulaCoranteMcp in oListFormulaCoranteMcp do
    begin
      if not (oFormulaCoranteMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        Inc(nCount);
        if cdsIndex.FindKey([oFormulaCoranteMcp.IdColecao, oFormulaCoranteMcp.IdColecaoPers, oFormulaCoranteMcp.IdFormula, oFormulaCoranteMcp.IdBase, oFormulaCoranteMcp.IdProduto, oFormulaCoranteMcp.IdEmbTinta, oFormulaCoranteMcp.IdGrupoTinta, oFormulaCoranteMcp.IdCorante, oFormulaCoranteMcp.SeqCorante]) then
          oFormulaCoranteSuvinil := oListFormulaCoranteSuvinil[cdsIndex.AsInt('Indice')];
        if (not Assigned(oFormulaCoranteSuvinil)) then
        begin
          oFormulaCoranteMcp.Status := dsDeleta;
        end;
      end;
      if nCount mod 500 = 1 then
        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListFormulaCoranteMcp.IndexOf(oFormulaCoranteMcp));
    end;
    AddMsg('Tempo (FormulaCorante(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListFormulaCoranteSuvinil.Clear;
    nCount := 0;
    AddMsg('Iniciou Gravação no Banco de Dados Tabela FormulaCorante');
    dNow := Now;

    qrMcp := TSQLQuery.Create(nil);
    qrMcp.SQLConnection := DmConexao.sqMcp;
    TD := nil;
    try

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListFormulaCoranteMcp.Count);

      for oFormulaCoranteMcp in oListFormulaCoranteMcp do
      begin
        Inc(nCount);
        cSql := oFormulaCoranteMcp.ToSQL;

        if ( cSql <> '' ) then
        begin
          if not Assigned(TD) then
            TD := DmConexao.sqMcp.BeginTransaction;
          qrMcp.SQL.Text := cSql;
          qrMcp.ExecSQL(True);
          //DmConexao.sqMcp.Execute(cSql, nil);

          if ( nRet <= 500 ) then
            Inc(nRet)
          else
          begin
            DmConexao.sqMcp.CommitFreeAndNil(TD);
            AddMsg('Tempo Commit Parcial Formula.'+nCount.ToString+' de '+oListFormulaCoranteMcp.Count.ToString);
            nRet := 0;
          end;
        end;

        if nCount mod 500 = 1 then
          if Assigned(ProgressbarEventUpdate) then
            ProgressbarEventUpdate(oListFormulaCoranteMcp.IndexOf(oFormulaCoranteMcp));
      end;
      if nRet >= 0 then
      begin
        AddMsg('Tempo (Antes commit FormulaCorante): ' + SecondToTime(SecondsBetween(Now, dNow)));
        dNow := Now;
        if Assigned(TD) then
          DmConexao.sqMcp.CommitFreeAndNil(TD);
        AddMsg('Tempo (Depois commit FormulaCorante): ' + SecondToTime(SecondsBetween(Now, dNow)));
      end;
      FreeAndNil(qrMcp);
    except
      on E: Exception do
      begin
        if Assigned(TD) then
          DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit FormulaCorante): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar FormulaCorante: ' + E.Message+ ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(cdsIndex);
    FreeAndNil(oListFormulaCoranteSuvinil);
    FreeAndNil(oListFormulaCoranteMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilFormulaTinta;
var
  oListFormulaTintaSuvinil, oListFormulaTintaMcp: TListFormulaTinta;
  oFormulaTintaSuvinil, oFormulaTintaMcp: TFormulaTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;
  cdsIndex: TVsClientDataSet;
  nCount, nRet: Integer;
  lFormulaMCPVazia : Boolean;
  qrMcp: TSQLQuery;

  procedure CopiarDados;
  begin
    oFormulaTintaMcp.IdColecao     := oFormulaTintaSuvinil.IdColecao;
    oFormulaTintaMcp.Descricao     := oFormulaTintaSuvinil.Descricao;
    oFormulaTintaMcp.IdFormula     := oFormulaTintaSuvinil.IdFormula;
    oFormulaTintaMcp.IdBase        := oFormulaTintaSuvinil.IdBase;
    oFormulaTintaMcp.IdMsgTinta    := oFormulaTintaSuvinil.IdMsgTinta;
    oFormulaTintaMcp.CodCatalogo   := oFormulaTintaSuvinil.CodCatalogo;
    oFormulaTintaMcp.IdEmbTinta    := oFormulaTintaSuvinil.IdEmbTinta;
    oFormulaTintaMcp.IdGrupoTinta  := oFormulaTintaSuvinil.IdGrupoTinta;
    oFormulaTintaMcp.IdProduto     := oFormulaTintaSuvinil.IdProduto;
    oFormulaTintaMcp.Personalizada := oFormulaTintaSuvinil.Personalizada;
    oFormulaTintaMcp.IdColecaoPers := oFormulaTintaSuvinil.IdColecaoPers;
    oFormulaTintaMcp.Inativo       := oFormulaTintaSuvinil.Inativo;
    oFormulaTintaMcp.IdOrigem      := oFormulaTintaSuvinil.IdOrigem;
  end;

begin
  nCount := 0;
  oListFormulaTintaSuvinil := TListFormulaTinta.Create;
  oListFormulaTintaMcp := TListFormulaTinta.Create;
  cdsIndex := TVsClientDataSet.Create(nil);
  cdsIndex.FieldDefs.Add('Indice', ftInteger);
  cdsIndex.FieldDefs.Add('IdColecao', ftInteger);
  cdsIndex.FieldDefs.Add('IdColecaoPers', ftInteger);
  cdsIndex.FieldDefs.Add('IdFormula', ftInteger);
  cdsIndex.FieldDefs.Add('IdBase', ftInteger);
  cdsIndex.FieldDefs.Add('IdProduto', ftInteger);
  cdsIndex.FieldDefs.Add('IdEmbTinta', ftInteger);
  cdsIndex.FieldDefs.Add('IdGrupoTinta', ftInteger);
  cdsIndex.CreateDataSet;
  cdsIndex.LogChanges := False;
  cdsIndex.DisableControls;
  cdsIndex.IndexFieldNames := 'IdColecao;IdColecaoPers;IdFormula;IdBase;IdProduto;IdEmbTinta;IdGrupoTinta';
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento Formula SQL Server');
    dNow := Now;
    oListFormulaTintaSuvinil := FSuvinil.CarregarFormulaTinta(DmConexao.sqSuvinil, cdsIndex);
    AddMsg('Carregou Formula Tinta SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento Formula MCP');
    dNow := Now;
    oListFormulaTintaMcp := FMCP.CarregarFormulaTinta(DmConexao.sqMCP);
    AddMsg('Carregou Formula MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou alteração tabela Formula');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListFormulaTintaSuvinil.Count);

    lFormulaMCPVazia := oListFormulaTintaMcp.IsEmpty;

    for oFormulaTintaSuvinil in oListFormulaTintaSuvinil do
    begin
      oFormulaTintaMcp := nil;
      Inc(nCount);

      if (not lFormulaMCPVazia) and (cdsIndex.FindKey([oFormulaTintaSuvinil.IdColecao, oFormulaTintaSuvinil.IdColecaoPers, oFormulaTintaSuvinil.IdFormula, oFormulaTintaSuvinil.IdBase, oFormulaTintaSuvinil.IdProduto, oFormulaTintaSuvinil.IdEmbTinta, oFormulaTintaSuvinil.IdGrupoTinta])) then
        oFormulaTintaMcp := oListFormulaTintaMcp[cdsIndex.AsInt('Indice')];
//      oFormulaTintaMcp := oListFormulaTintaMcp.FindKey(oFormulaTintaSuvinil.IdColecao, oFormulaTintaSuvinil.IdColecaoPers, oFormulaTintaSuvinil.IdFormula, oFormulaTintaSuvinil.IdBase, oFormulaTintaSuvinil.IdProduto, oFormulaTintaSuvinil.IdEmbTinta, oFormulaTintaSuvinil.IdGrupoTinta);

      if (not Assigned(oFormulaTintaMcp)) then
      begin
        oFormulaTintaMcp := TFormulaTinta.Create;
        oListFormulaTintaMcp.Add(oFormulaTintaMcp);
        oFormulaTintaMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oFormulaTintaMcp.Descricao = oFormulaTintaSuvinil.Descricao) and (oFormulaTintaMcp.CodCatalogo = oFormulaTintaSuvinil.CodCatalogo)
          and (oFormulaTintaMcp.IdMsgTinta = oFormulaTintaSuvinil.IdMsgTinta ) and (oFormulaTintaMcp.IdOrigem = oFormulaTintaSuvinil.IdOrigem)  then
      begin
        oFormulaTintaMcp.Status := dsIgual;
      end
      else
      begin
        oFormulaTintaMcp.Status := dsAltera;
        CopiarDados;
      end;
      if nCount mod 500 = 1 then
        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListFormulaTintaSuvinil.IndexOf(oFormulaTintaSuvinil));
    end;
    AddMsg('Tempo (Formula(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou exclusão tabela Formula');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListFormulaTintaMcp.Count);

    nCount := 0;
    for oFormulaTintaMcp in oListFormulaTintaMcp do
    begin
      if not (oFormulaTintaMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        Inc(nCount);
        if cdsIndex.FindKey([oFormulaTintaMcp.IdColecao, oFormulaTintaMcp.IdColecaoPers, oFormulaTintaMcp.IdFormula, oFormulaTintaMcp.IdBase, oFormulaTintaMcp.IdProduto, oFormulaTintaMcp.IdEmbTinta, oFormulaTintaMcp.IdGrupoTinta]) then
          oFormulaTintaSuvinil := oListFormulaTintaSuvinil[cdsIndex.AsInt('Indice')];
//        oFormulaTintaSuvinil := oListFormulaTintaSuvinil.FindKey(oFormulaTintaMcp.IdColecao, oFormulaTintaMcp.IdColecaoPers, oFormulaTintaMcp.IdFormula, oFormulaTintaMcp.IdBase, oFormulaTintaMcp.IdProduto, oFormulaTintaMcp.IdEmbTinta, oFormulaTintaMcp.IdGrupoTinta);
        if (not Assigned(oFormulaTintaSuvinil)) then
        begin
          oFormulaTintaMcp.Status := dsDeleta;
        end;
      end;
      if nCount mod 500 = 1 then
        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListFormulaTintaMcp.IndexOf(oFormulaTintaMcp));
    end;
    AddMsg('Tempo (Formula(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListFormulaTintaSuvinil.Clear;
    nCount := 0;
    AddMsg('Iniciou Gravação no Banco de Dados Tabela Formula');
    dNow := Now;
    qrMcp := TSQLQuery.Create(nil);
    qrMcp.SQLConnection := DmConexao.sqMcp;

    TD := nil;
    try
      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListFormulaTintaMcp.Count);

      for oFormulaTintaMcp in oListFormulaTintaMcp do
      begin
        Inc(nCount);
        cSql := oFormulaTintaMcp.ToSQL;
        if ( cSql <> '' ) then
        begin
          if not Assigned(TD) then
            TD := DmConexao.sqMcp.BeginTransaction;
          qrMcp.SQL.Text := cSql;
          qrMcp.ExecSQL(True);
          if ( nRet <= 500 ) then
            Inc(nRet)
          else
          begin
            DmConexao.sqMcp.CommitFreeAndNil(TD);
            AddMsg('Tempo Commit Parcial Formula.'+nCount.ToString+' de '+oListFormulaTintaMcp.Count.ToString);
            nRet := 0;
          end;
        end;
        if nCount mod 500 = 1 then
          if Assigned(ProgressbarEventUpdate) then
            ProgressbarEventUpdate(nCount-1);
      end;
      AddMsg('Tempo (Antes commit Formula): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      if Assigned(TD) then
        DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit Formula): ' + SecondToTime(SecondsBetween(Now, dNow)));

      FreeAndNil(qrMcp);
    except
      on E: Exception do
      begin
        if Assigned(TD) then
          DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit Formula): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar Formula: ' + E.Message+ ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListFormulaTintaSuvinil);
    FreeAndNil(oListFormulaTintaMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilGrupoTinta;
var
  oListGrupoTintaSuvinil, oListGrupoTintaMcp: TListGrupoTinta;
  oGrupoTintaSuvinil, oGrupoTintaMcp: TGrupoTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oGrupoTintaMcp.IdColecao    := oGrupoTintaSuvinil.IdColecao;
    oGrupoTintaMcp.IdGrupoTinta := oGrupoTintaSuvinil.IdGrupoTinta;
    oGrupoTintaMcp.Descricao    := oGrupoTintaSuvinil.Descricao;
    oGrupoTintaMcp.Situacao     := oGrupoTintaSuvinil.Situacao;
    oGrupoTintaMcp.IdOrigem     := oGrupoTintaSuvinil.IdOrigem;
    oGrupoTintaMcp.Padrao       := oGrupoTintaSuvinil.Padrao;
  end;

begin
  oListGrupoTintaSuvinil := TListGrupoTinta.Create;
  oListGrupoTintaMcp := TListGrupoTinta.Create;
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento GrupoTinta SQL Server');
    dNow := Now;
    oListGrupoTintaSuvinil := FSuvinil.CarregarGrupoTinta(DmConexao.sqSuvinil);
    AddMsg('Carregou Grupo Tinta SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento GrupoTinta MCP');
    dNow := Now;
    oListGrupoTintaMcp := FMCP.CarregarGrupoTinta(DmConexao.sqMCP);
    AddMsg('Carregou GrupoTinta MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela GrupoTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListGrupoTintaSuvinil.Count);

    for oGrupoTintaSuvinil in oListGrupoTintaSuvinil do
    begin
      oGrupoTintaMcp := oListGrupoTintaMcp.FindKey(oGrupoTintaSuvinil.IdColecao, oGrupoTintaSuvinil.IdGrupoTinta);
      if (not Assigned(oGrupoTintaMcp)) then
      begin
        oGrupoTintaMcp := TGrupoTinta.Create;
        oListGrupoTintaMcp.Add(oGrupoTintaMcp);
        oGrupoTintaMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oGrupoTintaMcp.Descricao = oGrupoTintaSuvinil.Descricao) then
      begin
        oGrupoTintaMcp.Status := dsIgual;
      end
      else
      begin
        oGrupoTintaMcp.Status := dsAltera;
        CopiarDados;
      end;
      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListGrupoTintaSuvinil.IndexOf(oGrupoTintaSuvinil));
    end;
    AddMsg('Tempo (GrupoTinta(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela GrupoTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListGrupoTintaMcp.Count);

    for oGrupoTintaMcp in oListGrupoTintaMcp do
    begin
      if not (oGrupoTintaMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oGrupoTintaSuvinil := oListGrupoTintaSuvinil.FindKey(oGrupoTintaMcp.IdColecao, oGrupoTintaMcp.IdGrupoTinta);
        if (not Assigned(oGrupoTintaSuvinil)) then
        begin
          oGrupoTintaMcp.Status := dsDeleta;
        end;
        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListGrupoTintaMcp.IndexOf(oGrupoTintaMcp));
      end;
    end;
    AddMsg('Tempo (GrupoTinta(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListGrupoTintaSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela GrupoTinta');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListGrupoTintaMcp.Count);
      for oGrupoTintaMcp in oListGrupoTintaMcp do
      begin
        cSql := oGrupoTintaMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListGrupoTintaMcp.IndexOf(oGrupoTintaMcp));
      end;
      AddMsg('Tempo (Antes commit GrupoTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit GrupoTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit GrupoTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar GrupoTinta: ' + E.Message+ ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListGrupoTintaSuvinil);
    FreeAndNil(oListGrupoTintaMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilItemBase;
begin
  //Não fazer nada, pois era executado comando na mão, para identificar a versão...
end;

procedure TImportadorIquine.MigrarSuvinilItemCorante;
begin
  //Não fazer nada, pois era executado comando na mão, para identificar a versão...
end;

procedure TImportadorIquine.MigrarSuvinilMsgTinta;
var
  oListMsgTintaSuvinil, oListMsgTintaMcp: TListMsgTinta;
  oMsgTintaSuvinil, oMsgTintaMcp: TMsgTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oMsgTintaMcp.IdColecao  := oMsgTintaSuvinil.IdColecao;
    oMsgTintaMcp.Descricao  := oMsgTintaSuvinil.Descricao;
    oMsgTintaMcp.IdMsgTinta := oMsgTintaSuvinil.IdMsgTinta;
    oMsgTintaMcp.IdOrigem   := oMsgTintaSuvinil.IdOrigem;
  end;

begin
  oListMsgTintaSuvinil := TListMsgTinta.Create;
  oListMsgTintaMcp := TListMsgTinta.Create;
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento MsgTinta SQL Server');
    dNow := Now;
    oListMsgTintaSuvinil := FSuvinil.CarregarMsgTinta(DmConexao.sqSuvinil);
    AddMsg('Carregou Msg Tinta SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento MsgTinta MCP');
    dNow := Now;
    oListMsgTintaMcp := FMCP.CarregarMsgTinta(DmConexao.sqMCP);
    AddMsg('Carregou MsgTinta MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela MsgTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListMsgTintaSuvinil.Count);

    for oMsgTintaSuvinil in oListMsgTintaSuvinil do
    begin
      oMsgTintaMcp := oListMsgTintaMcp.FindKey(oMsgTintaSuvinil.IdColecao, oMsgTintaSuvinil.IdMsgTinta);
      if (not Assigned(oMsgTintaMcp)) then
      begin
        oMsgTintaMcp := TMsgTinta.Create;
        oListMsgTintaMcp.Add(oMsgTintaMcp);
        oMsgTintaMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oMsgTintaMcp.Descricao = oMsgTintaSuvinil.Descricao) then
      begin
        oMsgTintaMcp.Status := dsIgual;
      end
      else
      begin
        oMsgTintaMcp.Status := dsAltera;
        CopiarDados;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListMsgTintaSuvinil.IndexOf(oMsgTintaSuvinil));
    end;
    AddMsg('Tempo (MsgTinta(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela MsgTinta');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListMsgTintaMcp.Count);

    for oMsgTintaMcp in oListMsgTintaMcp do
    begin
      if not (oMsgTintaMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oMsgTintaSuvinil := oListMsgTintaSuvinil.FindKey(oMsgTintaMcp.IdColecao, oMsgTintaMcp.IdMsgTinta);
        if (not Assigned(oMsgTintaSuvinil)) then
        begin
          oMsgTintaMcp.Status := dsDeleta;
        end;
      end;
      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListMsgTintaMcp.IndexOf(oMsgTintaMcp));
    end;
    AddMsg('Tempo (MsgTinta(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListMsgTintaSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela MsgTinta');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListMsgTintaMcp.Count);
      for oMsgTintaMcp in oListMsgTintaMcp do
      begin
        cSql := oMsgTintaMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListMsgTintaMcp.IndexOf(oMsgTintaMcp));
      end;
      AddMsg('Tempo (Antes commit MsgTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit MsgTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit MsgTinta): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar Msg Tinta: ' + E.Message+ ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListMsgTintaSuvinil);
    FreeAndNil(oListMsgTintaMcp);
  end;
end;

procedure TImportadorIquine.MigrarSuvinilProduto;
var
  oListProdutoSuvinil, oListProdutoMcp: TListProdutoTinta;
  oProdutoSuvinil, oProdutoMcp: TProdutoTinta;
  dNow: TDateTime;
  cSql: string;
  TD: TDBXTransaction;

  procedure CopiarDados;
  begin
    oProdutoMcp.IdColecao := oProdutoSuvinil.IdColecao;
    oProdutoMcp.Descricao := oProdutoSuvinil.Descricao;
    oProdutoMcp.Imagem    := oProdutoSuvinil.Imagem;
    oProdutoMcp.IdProduto := oProdutoSuvinil.IdProduto;
    oProdutoMcp.IdOrigem  := oProdutoSuvinil.IdOrigem;
    oProdutoMcp.Inativo   := oProdutoSuvinil.Inativo;

  end;

begin
  oListProdutoSuvinil := TListProdutoTinta.Create;
  oListProdutoMcp := TListProdutoTinta.Create;
//  memLog.Clear;
  try
    DmConexao.sqSuvinil.Connected := False;
    DmConexao.sqSuvinil.Connected := True;

    AddMsg('Iniciou Carregamento Produto SQL Server');
    dNow := Now;
    oListProdutoSuvinil := FSuvinil.CarregarProduto(DmConexao.sqSuvinil);
    AddMsg('Carregou Produto Tinta SQL Server em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Carregamento Produto MCP');
    dNow := Now;
    oListProdutoMcp := FMCP.CarregarProduto(DmConexao.sqMCP);
    AddMsg('Carregou Produto MCP em: ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Alteração tabela Produto');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListProdutoSuvinil.Count);

    for oProdutoSuvinil in oListProdutoSuvinil do
    begin
      oProdutoMcp := oListProdutoMcp.FindKey(oProdutoSuvinil.IdColecao, oProdutoSuvinil.IdProduto);
      if (not Assigned(oProdutoMcp)) then
      begin
        oProdutoMcp := TProdutoTinta.Create;
        oListProdutoMcp.Add(oProdutoMcp);
        oProdutoMcp.Status := dsInsere;
        CopiarDados;
      end
      else if (oProdutoMcp.Descricao = oProdutoSuvinil.Descricao) then
      begin
        oProdutoMcp.Status := dsIgual;
      end
      else
      begin
        oProdutoMcp.Status := dsAltera;
        CopiarDados;
      end;
    end;
    AddMsg('Tempo (Produto(s) alterada(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    AddMsg('Iniciou Exclusão Tabela Produto');
    dNow := Now;

    if Assigned(ProgressbarEventStart) then
      ProgressbarEventStart(oListProdutoMcp.Count);
    for oProdutoMcp in oListProdutoMcp do
    begin
      if not (oProdutoMcp.Status in [dsInsere, dsIgual, dsAltera]) then
      begin
        oProdutoSuvinil := oListProdutoSuvinil.FindKey(oProdutoMcp.IdColecao, oProdutoMcp.IdProduto);
        if (not Assigned(oProdutoSuvinil)) then
        begin
          oProdutoMcp.Status := dsDeleta;
        end;
      end;

      if Assigned(ProgressbarEventUpdate) then
        ProgressbarEventUpdate(oListProdutoMcp.IndexOf(oProdutoMcp));
    end;
    AddMsg('Tempo (Produto(s) excluida(s)): ' + SecondToTime(SecondsBetween(Now, dNow)));

    oListProdutoSuvinil.Clear;

    AddMsg('Iniciou Gravação no Banco de Dados Tabela Produto');
    dNow := Now;
    try
      TD := DmConexao.sqMcp.BeginTransaction;

      if Assigned(ProgressbarEventStart) then
        ProgressbarEventStart(oListProdutoMcp.Count);
      for oProdutoMcp in oListProdutoMcp do
      begin
        cSql := oProdutoMcp.ToSQL;
        if (cSql <> '') then
          DmConexao.sqMcp.Execute(cSql, nil);

        if Assigned(ProgressbarEventUpdate) then
          ProgressbarEventUpdate(oListProdutoMcp.IndexOf(oProdutoMcp));
      end;
      AddMsg('Tempo (Antes commit Produto): ' + SecondToTime(SecondsBetween(Now, dNow)));
      dNow := Now;
      DmConexao.sqMcp.CommitFreeAndNil(TD);
      AddMsg('Tempo (Depois commit Produto): ' + SecondToTime(SecondsBetween(Now, dNow)));
    except
      on E: Exception do
      begin
        DmConexao.sqMcp.RollbackFreeAndNil(TD);
        AddMsg('Tempo (Erro commit Produto): ' + SecondToTime(SecondsBetween(Now, dNow)));
        AddMsg('Erro ao migrar Produto: ' + E.Message + ' ' + cSql);
        raise;
      end;
    end;

  finally
    FreeAndNil(oListProdutoSuvinil);
    FreeAndNil(oListProdutoMcp);
  end;
end;

end.
