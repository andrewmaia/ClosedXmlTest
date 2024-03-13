
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
            titularesQuantidade = "B27",
            titularesNomePlano = "D27",
            titularesQuantidadeVidas = "E27",
            titularesReembolsoConsulta = "F27",

            dependentesQuantidade = "B28",
            dependentesNomePlano = "D28",
            dependentesQuantidadeVidas = "E28",
            dependentesReembolsoConsulta = "F28",

            agregadosQuantidade = "B29",
            agregadosNomePlano = "D29",
            agregadosQuantidadeVidas = "E29",
            agregadosReembolsoConsulta = "F29",
            totalVidasPlanoQuantidade = "B30",
            totalVidasPlanoNomePlano = "D30",
            totalVidasPlanoQuantidadeVidas = "E30",
            totalVidasPlanoReembolsoConsulta = "F30",
            totalFuncionariosFgtsQuantidade = "B31",
            totalFuncionariosFgtsNomePlano = "D31",
            totalFuncionariosFgtsQuantidadeVidas = "E31",
            totalFuncionariosFgtsReembolsoConsulta = "F31",
            categoriaPlanosAtuaisObservacoes="G27",
            informacoesSeguradosAgregadosSimNao="B36",
            informacoesSeguradosAgregadosGrauParentesco="C36",
            informacoesSeguradosAgregadosQuantidade="D36",
            prestadorServicoSimNao="B37",
            prestadorServicoGrauParentesco="C37",
            prestadorServicoQuantidade="D37",
            informacoesSeguradosObservacoes="G36",

            subEstipulanteItemCellRangeTemplate="A40:I40",
            estrategiaRange="A42:G61",            

            coringa="";
        #endregion

        private readonly IXLWorksheet _estrategiaWorkSheet;
        private readonly IXLWorksheet _baseDadosEstudosWorkSheet;            
        private readonly IXLRange _subEstipulanteItemBlockTemplate;
        private readonly IXLRange _estrategiaBlock;        
        public BeneficioOdontoQarCreator(Stream templateStream,string? outputFileAddress=null)
            :base(templateStream,outputFileAddress){
           _estrategiaWorkSheet = _workbook.Worksheets.First(x=>x.Name=="ESTRATEGIA");
           _baseDadosEstudosWorkSheet = _workbook.Worksheets.First(x=>x.Name=="BASE DE DADOS  ESTUDOS");    
           _subEstipulanteItemBlockTemplate = _estrategiaWorkSheet.Range(subEstipulanteItemCellRangeTemplate);
           _estrategiaBlock=_estrategiaWorkSheet.Range(estrategiaRange);
        }

        public override MemoryStream GenerateExcelFile(){        
            BuildSectionFormularioContacaoDental();
            BuildSectionCategoriaPlanosAtuaisValoresPercapta();
            BuildSectionInformacoesSegurados();
            BuildSectionInformacoesSegurados();
            BuildSectionSubEstipulante();
            BuildSectionEstrategias();
            BuildSectionBaseDadosEstudo();
            _subEstipulanteItemBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);                
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
            _estrategiaWorkSheet.Cell(titularesQuantidade).SetValue(10);
            _estrategiaWorkSheet.Cell(titularesNomePlano).SetValue("Plano A");            
            _estrategiaWorkSheet.Cell(titularesQuantidadeVidas).SetValue(10);                        
            _estrategiaWorkSheet.Cell(titularesReembolsoConsulta).SetValue(100);                                    
            _estrategiaWorkSheet.Cell(dependentesQuantidade).SetValue(20);                                                
            _estrategiaWorkSheet.Cell(dependentesNomePlano).SetValue("Plano B");
            _estrategiaWorkSheet.Cell(dependentesQuantidadeVidas).SetValue(20);
            _estrategiaWorkSheet.Cell(dependentesReembolsoConsulta).SetValue(150);
            _estrategiaWorkSheet.Cell(agregadosQuantidade).SetValue(30);            
            _estrategiaWorkSheet.Cell(agregadosNomePlano).SetValue("Plano C");                        
            _estrategiaWorkSheet.Cell(agregadosQuantidadeVidas).SetValue(30);                                    
            _estrategiaWorkSheet.Cell(agregadosReembolsoConsulta).SetValue(250);                                                
            _estrategiaWorkSheet.Cell(totalVidasPlanoQuantidade).SetValue(40);                                                            
            _estrategiaWorkSheet.Cell(totalVidasPlanoNomePlano).SetValue("Plano D");                                                                        
            _estrategiaWorkSheet.Cell(totalVidasPlanoQuantidadeVidas).SetValue(15);
            _estrategiaWorkSheet.Cell(totalVidasPlanoReembolsoConsulta).SetValue(100);
            _estrategiaWorkSheet.Cell(totalFuncionariosFgtsQuantidade).SetValue(50);            
            _estrategiaWorkSheet.Cell(totalFuncionariosFgtsNomePlano).SetValue("Plano E");
            _estrategiaWorkSheet.Cell(totalFuncionariosFgtsQuantidadeVidas).SetValue(40);            
            _estrategiaWorkSheet.Cell(totalFuncionariosFgtsReembolsoConsulta).SetValue(180);                        
            _estrategiaWorkSheet.Cell(categoriaPlanosAtuaisObservacoes).SetValue("Observações....");                                    
        }

        private void BuildSectionInformacoesSegurados() {

            _estrategiaWorkSheet.Cell(informacoesSeguradosAgregadosSimNao).SetValue("Sim");            
            _estrategiaWorkSheet.Cell(informacoesSeguradosAgregadosGrauParentesco).SetValue("Grau Parentesco");
            _estrategiaWorkSheet.Cell(informacoesSeguradosAgregadosQuantidade).SetValue(50);           
            _estrategiaWorkSheet.Cell(prestadorServicoSimNao).SetValue("Sim");            
            _estrategiaWorkSheet.Cell(prestadorServicoGrauParentesco).SetValue("Grau Parentesco");
            _estrategiaWorkSheet.Cell(prestadorServicoQuantidade).SetValue(50);                         
            _estrategiaWorkSheet.Cell(informacoesSeguradosObservacoes).SetValue("Observações...");                                     
        }

        private void BuildSectionSubEstipulante() {
            List<SubEstipulanteOdonto> subEstipulantes = MockSubEstimulantes();
            int referenceLine=_subEstipulanteItemBlockTemplate.RangeAddress.LastAddress.RowNumber+1;            
            foreach(var subEstipulante in subEstipulantes){
                _estrategiaBlock.InsertRowsAbove(_subEstipulanteItemBlockTemplate.RowCount());                
                _subEstipulanteItemBlockTemplate.CopyTo(_estrategiaWorkSheet.Cell(referenceLine,"A"));
                _estrategiaWorkSheet.Cell(referenceLine,"B").SetValue(subEstipulante.RazaoSocial);
                _estrategiaWorkSheet.Cell(referenceLine,"F").SetValue(subEstipulante.CNPJ);
                referenceLine++;                                 
            }
        }        

        private void BuildSectionEstrategias() {
            var blockLines = _estrategiaBlock.Rows().ToList();
            blockLines[1].Cell(1).SetValue("Estrategia");
        }        

        private void BuildSectionBaseDadosEstudo(){
            _baseDadosEstudosWorkSheet.Cell("A2").InsertData(MockBaseDadosSaude());
        }

        #region Mock de Dados

         static List<PessoaOdonto> MockBaseDadosSaude(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Maria", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Jose", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Ricardo", "41.646.207/0001-15", "Identificacao", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
            ];
 
        }

        static List<SubEstipulanteOdonto> MockSubEstimulantes(){
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
    record PessoaOdonto(string Empresa, string CNPJ,string Matricuka,string Parentesco, string Situacao, string CID, string Municipio, string UF, string Operadora, string Plano, int ValorAtual);
    record SubEstipulanteOdonto(string RazaoSocial, string CNPJ);    
    #endregion
}
