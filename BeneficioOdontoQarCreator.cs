
using ClosedXML.Excel;

namespace ClosedXmlTest
{
    public class BeneficioOdontoQarCreator: QarCreator
    {
        #region Endereços de Informação 
        const string 
            consultor = "B7",
            razaoSocialEstipulante = "B8",
            cnpjEstipulante = "F8",            
            operadoraAtual = "B10",
            tempoContrato = "F10",            
            classificacaoCliente = "B11",
            possuiCocorretor = "B12",            
            cocorretorPercentual = "C12",
            aniversarioContrato = "B13",
            breakEven = "B14",            
            possuiMulta = "B15",
            regraMulta = "D15",
            relatorioSinistralidade="B16",            
            relatorioSinistralidadePercentual="D16",
            modalidadeContrato="B17",   
            haDependentes="D17",
            haReembolso="B18",            
            reembolsoValores="D18",                        
            regraUpDownGrade="B19",    
            contribuicaoTitular="B20",    
            contribuicaoTitularEmpresa="F20",        
            contribuicaoDependente="B21",    
            contribuicaoDependenteEmpresa="F21", 
            elegibilidade = "B22",
            coringa="";
        #endregion

        private readonly IXLWorksheet _estrategiaWorkSheet;
        private readonly IXLWorksheet _baseDadosEstudosWorkSheet;            
        public BeneficioOdontoQarCreator(Stream templateStream,string? outputFileAddress=null)
            :base(templateStream,outputFileAddress){
           _estrategiaWorkSheet = _workbook.Worksheets.First(x=>x.Name=="ESTRATEGIA");
           _baseDadosEstudosWorkSheet = _workbook.Worksheets.First(x=>x.Name=="BASE DE DADOS  ESTUDOS");                      
        }

        public override MemoryStream GenerateExcelFile(){        
            BuildSectionFormularioContacaoDental();
            BuildSectionCategoriaPlanosAtuaisValoresPercapta();
            BuildSectionInformacoesSegurados();
            BuildSectionInformacoesSegurados();
            BuildSectionSubEstipulante();
            BuildSectionEstrategias();
            BuildSectionBaseDadosEstudo();
            return base.GenerateExcelFile();
        }
        private void BuildSectionFormularioContacaoDental(){
            _estrategiaWorkSheet.Cell(consultor).SetValue("José da Silva Santos");
            _estrategiaWorkSheet.Cell(razaoSocialEstipulante).SetValue("Razão Social Z");
            _estrategiaWorkSheet.Cell(cnpjEstipulante).SetValue("11.969.923/0001-72");
            _estrategiaWorkSheet.Cell(operadoraAtual).SetValue("Amil");
            _estrategiaWorkSheet.Cell(tempoContrato).SetValue("2 anos");
            _estrategiaWorkSheet.Cell(classificacaoCliente).SetValue("Ouro");            
            _estrategiaWorkSheet.Cell(possuiCocorretor).SetValue("Sim");
            _estrategiaWorkSheet.Cell(cocorretorPercentual).SetValue(0.5);
            _estrategiaWorkSheet.Cell(aniversarioContrato).SetValue(new DateTime(2025,1,31));
            _estrategiaWorkSheet.Cell(breakEven).SetValue("Break Even");
            _estrategiaWorkSheet.Cell(possuiMulta).SetValue("Sim");
            _estrategiaWorkSheet.Cell(regraMulta).SetValue("Regra XPTO");            
            _estrategiaWorkSheet.Cell(relatorioSinistralidade).SetValue("Sim");                        
            _estrategiaWorkSheet.Cell(relatorioSinistralidadePercentual).SetValue(0.1);                                    
            _estrategiaWorkSheet.Cell(modalidadeContrato).SetValue("Adesão");                        
            _estrategiaWorkSheet.Cell(haDependentes).SetValue("Sim");      
            _estrategiaWorkSheet.Cell(haReembolso).SetValue("Sim");     
            _estrategiaWorkSheet.Cell(regraUpDownGrade).SetValue("Regra XPTO");                 
            _estrategiaWorkSheet.Cell(contribuicaoTitular).SetValue(0.1);                             
            _estrategiaWorkSheet.Cell(contribuicaoTitularEmpresa).SetValue(0.9);                      
            _estrategiaWorkSheet.Cell(contribuicaoDependente).SetValue(0.2);                             
            _estrategiaWorkSheet.Cell(contribuicaoDependenteEmpresa).SetValue(0.8);
            _estrategiaWorkSheet.Cell(elegibilidade).SetValue("Elegibilidade XPTO");
        }
        
        private void BuildSectionCategoriaPlanosAtuaisValoresPercapta() {
            
        }

        private void BuildSectionInformacoesSegurados() {
            
        }

        private void BuildSectionSubEstipulante() {
            //repete linhas
        }        

        private void BuildSectionEstrategias() {
            
        }        

        private void BuildSectionBaseDadosEstudo(){

        }




    }
}
