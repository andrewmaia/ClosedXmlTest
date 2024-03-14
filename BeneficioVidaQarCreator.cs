
using ClosedXML.Excel;

namespace ClosedXmlTest
{
    public class BeneficioVidaQarCreator: QarCreator
    {
        #region Endereços de Informação 
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
            coringa="";
        #endregion
        private readonly IXLWorksheet _estrategiaWorkSheet;
        private readonly IXLWorksheet _baseDadosEstudosWorkSheet;            
        public BeneficioVidaQarCreator(Stream templateStream,string? outputFileAddress=null)
            :base(templateStream,outputFileAddress){
           _estrategiaWorkSheet = _workbook.Worksheets.First(x=>x.Name=="ESTRATEGIA");
           _baseDadosEstudosWorkSheet = _workbook.Worksheets.First(x=>x.Name=="BASE DE DADOS  ESTUDOS");
        }

        public override MemoryStream GenerateExcelFile(){        
            BuildSectionFormularioContacaoVida();
            BuildSectionCategoriaPlanosAtuaisValoresPercapta();
            BuildSectionInformacoesSegurados();
            BuildSectionCondicoesApoliceAtual();
            BuildSectionSubEstipulante();
            BuildSectionEstrategias();
            BuildSectionBaseDadosEstudo();
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
        }

        private void BuildSectionInformacoesSegurados() {

        }
        private void BuildSectionCondicoesApoliceAtual() {

        }           

        private void BuildSectionSubEstipulante() {

        }        

        private void BuildSectionEstrategias() {
        }        

        private void BuildSectionBaseDadosEstudo(){

        }

        #region Mock de Dados

         static List<PessoaVida> MockBaseDadosVida(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Maria", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Jose", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Ricardo", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
            ];
 
        }

        static List<SubEstipulanteVida> MockSubEstimulantes(){
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
    record PessoaVida(string Empresa, string CNPJ,string Matricuka,string Parentesco, string Situacao, string CID, string Municipio, string UF, string Operadora, string Plano, int ValorAtual);
    record SubEstipulanteVida(string RazaoSocial, string CNPJ);    
    #endregion
}
