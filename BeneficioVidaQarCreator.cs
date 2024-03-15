
using ClosedXML.Excel;

namespace ClosedXmlTest
{
    public class BeneficioVidaQarCreator: QarCreator
    {
        #region Endereços Células Seção Formulario Cotacao Vida
        const string
            consultor = "B7",
            razaoSocialEstipulante = "B8",     
            cnpjEstipulante = "I8",  
            seguradoraAtual = "B9",
            possuiCocorretor = "B10",
            percentualCocorretor = "C10",            
            aniversarioContrato = "B11",            
            relatorioSinistralidade="B12",            
            valorSinistro="D12",         
            modalidadeContrato="B13",
            modalidadeContratoSimNao="D13",            
            elegibildiade="B14",
            elegibildiadeSimNao="D14",
            taxaAtual="B15",
            valorFatura="D15",
            modeloCapital="B16",
            valorCapital="B17",
            custeioSeguro="B18", 
            sinistralidade="B19",

        #endregion

        #region Endereços Células Seção Categoria Planos Atuais Valores Percapta
            categoriaPlanosAtuaisValoresPercaptaObservacao="I24",
            basicaTitularPercentual="B24",
            basicaTitularValor="C24",
            basicaConjugePercentual="D24",
            basicaConjugeValor="E24",
            basicaFilhosPercentual="F24",
            basicaFilhosValor="G24",

            indenizacaoAcidenteTitularPercentual="B25",
            indenizacaoAcidenteTitularValor="C25",
            indenizacaoAcidenteConjugePercentual="D25",
            indenizacaoAcidenteConjugeValor="E25",
            indenizacaoAcidenteFilhosPercentual="F25",
            indenizacaoAcidenteFilhosValor="G25",

            invalidezTotalParcialTitularPercentual="B26",
            invalidezTotalParcialTitularValor="C26",
            invalidezTotalParcialConjugePercentual="D26",
            invalidezTotalParcialConjugeValor="E26",
            invalidezTotalParcialFilhosPercentual="F26",
            invalidezTotalParcialFilhosValor="G26",

            invalidezTotalTitularPercentual="B27",
            invalidezTotalTitularValor="C27",
            invalidezTotalConjugePercentual="D27",
            invalidezTotalConjugeValor="E27",
            invalidezTotalFilhosPercentual="F27",
            invalidezTotalFilhosValor="G27",

            invalidezFuncionalTitularPercentual="B28",
            invalidezFuncionalTitularValor="C28",
            invalidezFuncionalConjugePercentual="D28",
            invalidezFuncionalConjugeValor="E28",
            invalidezFuncionalFilhosPercentual="F28",
            invalidezFuncionalFilhosValor="G28",

            antecipacaoEspecialTitularPercentual="B29",
            antecipacaoEspecialTitularValor="C29",
            antecipacaoEspecialConjugePercentual="D29",
            antecipacaoEspecialConjugeValor="E29",
            antecipacaoEspecialFilhosPercentual="F29",
            antecipacaoEspecialFilhosValor="G29",

            indenizacaoFilhosTitularPercentual="B30",
            indenizacaoFilhosTitularValor="C30",
            indenizacaoFilhosConjugePercentual="D30",
            indenizacaoFilhosConjugeValor="E30",
            indenizacaoFilhosFilhosPercentual="F30",
            indenizacaoFilhosFilhosValor="G30",

            adaptacaoCasaVeiculoTitularPercentual="B31",
            adaptacaoCasaVeiculoTitularValor="C31",
            adaptacaoCasaVeiculoConjugePercentual="D31",
            adaptacaoCasaVeiculoConjugeValor="E31",
            adaptacaoCasaVeiculoFilhosPercentual="F31",
            adaptacaoCasaVeiculoFilhosValor="G31",

            invalidezLaborativaTitularPercentual="B32",
            invalidezLaborativaTitularValor="C32",
            invalidezLaborativaConjugePercentual="D32",
            invalidezLaborativaConjugeValor="E32",
            invalidezLaborativaFilhosPercentual="F32",
            invalidezLaborativaFilhosValor="G32",

            despachanteFuneralTitularPercentual="B33",
            despachanteFuneralTitularValor="C33",
            despachanteFuneralConjugePercentual="D33",
            despachanteFuneralConjugeValor="E33",
            despachanteFuneralFilhosPercentual="F33",
            despachanteFuneralFilhosValor="G33",

            verbaRecisoriasTitularPercentual="B34",
            verbaRecisoriasTitularValor="C34",
            verbaRecisoriasConjugePercentual="D34",
            verbaRecisoriasConjugeValor="E34",
            verbaRecisoriasFilhosPercentual="F34",
            verbaRecisoriasFilhosValor="G34",

            assistenciaFuneralTitularPercentual="B35",
            assistenciaFuneralTitularValor="C35",
            assistenciaFuneralConjugePercentual="D35",
            assistenciaFuneralConjugeValor="E35",
            assistenciaFuneralFilhosPercentual="F35",
            assistenciaFuneralFilhosValor="G35",

            assistenciaFuneralPaisTitularPercentual="B36",
            assistenciaFuneralPaisTitularValor="C36",
            assistenciaFuneralPaisConjugePercentual="D36",
            assistenciaFuneralPaisConjugeValor="E36",
            assistenciaFuneralPaisFilhosPercentual="F36",
            assistenciaFuneralPaisFilhosValor="G36",

            assistenciaFuneralIndividualTitularPercentual="B37",
            assistenciaFuneralIndividualTitularValor="C37",
            assistenciaFuneralIndividualConjugePercentual="D37",
            assistenciaFuneralIndividualConjugeValor="E37",
            assistenciaFuneralIndividualFilhosPercentual="F37",
            assistenciaFuneralIndividualFilhosValor="G37",

            auxilioFuneralFamiliarTitularPercentual="B38",
            auxilioFuneralFamiliarTitularValor="C38",
            auxilioFuneralFamiliarConjugePercentual="D38",
            auxilioFuneralFamiliarConjugeValor="E38",
            auxilioFuneralFamiliarFilhosPercentual="F38",
            auxilioFuneralFamiliarFilhosValor="G38",

            auxilioFuneralIndividualTitularPercentual="B39",
            auxilioFuneralIndividualTitularValor="C39",
            auxilioFuneralIndividualConjugePercentual="D39",
            auxilioFuneralIndividualConjugeValor="E39",
            auxilioFuneralIndividualFilhosPercentual="F39",
            auxilioFuneralIndividualFilhosValor="G39",

            excedenteTecnicoTitularPercentual="B40",
            excedenteTecnicoTitularValor="C40",
            excedenteTecnicoConjugePercentual="D40",
            excedenteTecnicoConjugeValor="E40",
            excedenteTecnicoFilhosPercentual="F40",
            excedenteTecnicoFilhosValor="G40",

            cestaBasicaTitularPercentual="B41",
            cestaBasicaTitularValor="C41",
            cestaBasicaConjugePercentual="D41",
            cestaBasicaConjugeValor="E41",
            cestaBasicaFilhosPercentual="F41",
            cestaBasicaFilhosValor="G41",

            cestaNatalidadeTitularPercentual="B42",
            cestaNatalidadeTitularValor="C42",
            cestaNatalidadeConjugePercentual="D42",
            cestaNatalidadeConjugeValor="E42",
            cestaNatalidadeFilhosPercentual="F42",
            cestaNatalidadeFilhosValor="G42",

            incapacidadeAcidenteDoencaTitularPercentual="B43",
            incapacidadeAcidenteDoencaTitularValor="C43",
            incapacidadeAcidenteDoencaConjugePercentual="D43",
            incapacidadeAcidenteDoencaConjugeValor="E43",
            incapacidadeAcidenteDoencaFilhosPercentual="F43",
            incapacidadeAcidenteDoencaFilhosValor="G43",

            incapacidadeAcidenteTitularPercentual="B44",
            incapacidadeAcidenteTitularValor="C44",
            incapacidadeAcidenteConjugePercentual="D44",
            incapacidadeAcidenteConjugeValor="E44",
            incapacidadeAcidenteFilhosPercentual="F44",
            incapacidadeAcidenteFilhosValor="G44",

            nomeCoberturaExtra="A45",
            coberturaExtraTitularPercentual="B45",
            coberturaExtraTitularValor="C45",
            coberturaExtraConjugePercentual="D45",
            coberturaExtraConjugeValor="E45",
            coberturaExtraFilhosPercentual="F45",
            coberturaExtraFilhosValor="G45",          
        #endregion

        #region Endereços Células Seção Informações Segurados
            agregados="B49",
            agregadosQuantidade="C49",     
            informacoesSeguradoObservacao="I49",
        #endregion

        #region Endereços  Células Seção Condições da Apolice Atual
            regraDpsImplantacao="B54",            
            regraDpsNovasAdesoes="B55",
            limiteIdadeNovasInclusoes="B56",    
        #endregion

        #region Endereços Endereços Células Seção SubEstipulantes
         subEstipulanteItemCellRangeTemplate="A60:H60",        
        #endregion
 
        #region Endereços Células Seção Importante

            importanteRange ="A62:H68",
        #endregion
        
        #region Endereços Células Seção Estratégia

            estrategiaRange ="A70:H81";

         #endregion
        private readonly IXLWorksheet _estrategiaWorkSheet;
        private readonly IXLWorksheet _baseDadosEstudosWorkSheet;     
        private readonly IXLRange _subEstipulanteItemBlockTemplate;               
        private readonly IXLRange _importanteBlock;         
        private readonly IXLRange _estrategiaBlock; 
        public BeneficioVidaQarCreator(Stream templateStream,string? outputFileAddress=null)
            :base(templateStream,outputFileAddress){
           _estrategiaWorkSheet = _workbook.Worksheets.First(x=>x.Name=="ESTRATEGIA");
           _baseDadosEstudosWorkSheet = _workbook.Worksheets.First(x=>x.Name=="BASE DE DADOS  ESTUDOS");
           _subEstipulanteItemBlockTemplate = _estrategiaWorkSheet.Range(subEstipulanteItemCellRangeTemplate);          
           _importanteBlock=_estrategiaWorkSheet.Range(importanteRange);                       
           _estrategiaBlock=_estrategiaWorkSheet.Range(estrategiaRange);           
        }

        public override MemoryStream GenerateExcelFile(){        
            BuildSectionFormularioContacaoVida();
            BuildSectionCategoriaPlanosAtuaisValoresPercapta();
            BuildSectionInformacoesSegurados();
            BuildSectionCondicoesApoliceAtual();
            BuildSectionSubEstipulante();
            BuildSectionEstrategias();
            BuildSectionBaseDadosEstudo();
            _subEstipulanteItemBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);            
            return base.GenerateExcelFile();
        }
        private void BuildSectionFormularioContacaoVida(){
            _estrategiaWorkSheet.Cell(consultor).SetValue("José da Silva Santos");
            _estrategiaWorkSheet.Cell(razaoSocialEstipulante).SetValue("Razão Social X");
            _estrategiaWorkSheet.Cell(cnpjEstipulante).SetValue("11.969.923/0001-72");
            _estrategiaWorkSheet.Cell(seguradoraAtual).SetValue("Sul América");
            _estrategiaWorkSheet.Cell(possuiCocorretor).SetValue("Sim");
            _estrategiaWorkSheet.Cell(percentualCocorretor).SetValue(0.05);
            _estrategiaWorkSheet.Cell(aniversarioContrato).SetValue(new DateTime(2025,1,31));
            _estrategiaWorkSheet.Cell(relatorioSinistralidade).SetValue("Sim");
            _estrategiaWorkSheet.Cell(valorSinistro).SetValue(2000);
            _estrategiaWorkSheet.Cell(modalidadeContrato).SetValue("Modalidade");
            _estrategiaWorkSheet.Cell(modalidadeContratoSimNao).SetValue("Sim");
            _estrategiaWorkSheet.Cell(elegibildiade).SetValue("Elegibilidade");
            _estrategiaWorkSheet.Cell(elegibildiadeSimNao).SetValue("Não");
            _estrategiaWorkSheet.Cell(taxaAtual).SetValue(0.08);
            _estrategiaWorkSheet.Cell(valorFatura).SetValue(560);
            _estrategiaWorkSheet.Cell(modeloCapital).SetValue("Modelo Capital");       
            _estrategiaWorkSheet.Cell(valorCapital).SetValue(3000);
            _estrategiaWorkSheet.Cell(custeioSeguro).SetValue("Sim");            
            _estrategiaWorkSheet.Cell(sinistralidade).SetValue(0.25);
        }
        private void BuildSectionCategoriaPlanosAtuaisValoresPercapta() {
            _estrategiaWorkSheet.Cell(categoriaPlanosAtuaisValoresPercaptaObservacao).SetValue("Observações...");            
            _estrategiaWorkSheet.Cell(basicaTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(basicaTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(basicaConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(basicaConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(basicaFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(basicaFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(indenizacaoAcidenteTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(indenizacaoAcidenteTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(indenizacaoAcidenteConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(indenizacaoAcidenteConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(indenizacaoAcidenteFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(indenizacaoAcidenteFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(invalidezTotalParcialTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(invalidezTotalParcialTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(invalidezTotalParcialConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(invalidezTotalParcialConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(invalidezTotalParcialFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(invalidezTotalParcialFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(invalidezTotalTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(invalidezTotalTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(invalidezTotalConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(invalidezTotalConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(invalidezTotalFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(invalidezTotalFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(invalidezFuncionalTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(invalidezFuncionalTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(invalidezFuncionalConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(invalidezFuncionalConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(invalidezFuncionalFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(invalidezFuncionalFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(antecipacaoEspecialTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(antecipacaoEspecialTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(antecipacaoEspecialConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(antecipacaoEspecialConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(antecipacaoEspecialFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(antecipacaoEspecialFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(indenizacaoFilhosTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(indenizacaoFilhosTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(indenizacaoFilhosConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(indenizacaoFilhosConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(indenizacaoFilhosFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(indenizacaoFilhosFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(adaptacaoCasaVeiculoFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(invalidezLaborativaTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(invalidezLaborativaTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(invalidezLaborativaConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(invalidezLaborativaConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(invalidezLaborativaFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(invalidezLaborativaFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(despachanteFuneralTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(despachanteFuneralTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(despachanteFuneralConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(despachanteFuneralConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(despachanteFuneralFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(despachanteFuneralFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(verbaRecisoriasTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(verbaRecisoriasTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(verbaRecisoriasConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(verbaRecisoriasConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(verbaRecisoriasFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(verbaRecisoriasFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(assistenciaFuneralTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(assistenciaFuneralTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(assistenciaFuneralConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(assistenciaFuneralConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(assistenciaFuneralFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(assistenciaFuneralFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(assistenciaFuneralPaisFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(assistenciaFuneralIndividualFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(auxilioFuneralFamiliarFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(auxilioFuneralIndividualFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(excedenteTecnicoTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(excedenteTecnicoTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(excedenteTecnicoConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(excedenteTecnicoConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(excedenteTecnicoFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(excedenteTecnicoFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(cestaBasicaTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(cestaBasicaTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(cestaBasicaConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(cestaBasicaConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(cestaBasicaFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(cestaBasicaFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(cestaNatalidadeTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(cestaNatalidadeTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(cestaNatalidadeConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(cestaNatalidadeConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(cestaNatalidadeFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(cestaNatalidadeFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteDoencaFilhosValor).SetValue(300);

            _estrategiaWorkSheet.Cell(incapacidadeAcidenteTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(incapacidadeAcidenteFilhosValor).SetValue(300);


            _estrategiaWorkSheet.Cell(nomeCoberturaExtra).SetValue("ALGUMA COBERTURA EXTRA");
            _estrategiaWorkSheet.Cell(coberturaExtraTitularPercentual).SetValue(0.25);
            _estrategiaWorkSheet.Cell(coberturaExtraTitularValor).SetValue(100);
            _estrategiaWorkSheet.Cell(coberturaExtraConjugePercentual).SetValue(0.35);
            _estrategiaWorkSheet.Cell(coberturaExtraConjugeValor).SetValue(200);
            _estrategiaWorkSheet.Cell(coberturaExtraFilhosPercentual).SetValue(0.22);
            _estrategiaWorkSheet.Cell(coberturaExtraFilhosValor).SetValue(300);
     
        }
        private void BuildSectionInformacoesSegurados() {
            _estrategiaWorkSheet.Cell(agregados).SetValue("Sim");
            _estrategiaWorkSheet.Cell(agregadosQuantidade).SetValue(100);
            _estrategiaWorkSheet.Cell(informacoesSeguradoObservacao).SetValue("Observações...");
        }
        private void BuildSectionCondicoesApoliceAtual() {
            _estrategiaWorkSheet.Cell(regraDpsImplantacao).SetValue("Regra 1");
            _estrategiaWorkSheet.Cell(regraDpsNovasAdesoes).SetValue("Regra 2");
            _estrategiaWorkSheet.Cell(limiteIdadeNovasInclusoes).SetValue("Regra 3");
        }           
        private void BuildSectionSubEstipulante() {
            List<SubEstipulanteVida> subEstipulantes = MockSubEstimulantes();
            int referenceLine=_subEstipulanteItemBlockTemplate.RangeAddress.LastAddress.RowNumber+1;            
            foreach(var subEstipulante in subEstipulantes){
                _importanteBlock.InsertRowsAbove(_subEstipulanteItemBlockTemplate.RowCount());                
                _subEstipulanteItemBlockTemplate.CopyTo(_estrategiaWorkSheet.Cell(referenceLine,"A"));
                _estrategiaWorkSheet.Cell(referenceLine,"B").SetValue(subEstipulante.RazaoSocial);
                _estrategiaWorkSheet.Cell(referenceLine,"H").SetValue(subEstipulante.CNPJ);
                referenceLine++;                                 
            }
        }        
        private void BuildSectionEstrategias() {
            var blockLines = _estrategiaBlock.Rows().ToList();
            blockLines[1].Cell(1).SetValue("Estrategia...");            
        }        
        private void BuildSectionBaseDadosEstudo(){
            _baseDadosEstudosWorkSheet.Cell("A2").InsertData(MockBaseDadosVida());
        }

        #region Mock de Dados
        List<PessoaVida> MockBaseDadosVida(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15", "M","11111111", new DateTime(2000,1,1),20,"Faixa Etaria", "Cargo",100,100,"Situação","CID"),
                new("Empresa da Maria", "41.646.207/0001-15", "M","11111111", new DateTime(2000,1,1),20,"Faixa Etaria", "Cargo",100,100,"Situação","CID"),
                new("Empresa do Roberto", "41.646.207/0001-15", "M","11111111", new DateTime(2000,1,1),20,"Faixa Etaria", "Cargo",100,100,"Situação","CID"),                                
            ];
 
        }
        List<SubEstipulanteVida> MockSubEstimulantes(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15")
                ,new("Empresa da Maria", "41.646.207/0001-14")
                ,new("Empresa do José", "41.646.207/0001-16")
            ];
        }           
        #endregion
    }

    #region Classes para Mock
    record PessoaVida(string Empresa, string CNPJ,string Sexo,string Identificacao,DateTime DataNascimento,int Idade,string FaixaEtaria,string Cargo, decimal Salario, decimal CapitalSegurado, string Situacao,string CID);
    record SubEstipulanteVida(string RazaoSocial, string CNPJ);    
    #endregion
}
