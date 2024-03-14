using ClosedXML.Excel;
using ClosedXmlTest;
using DocumentFormat.OpenXml.InkML;

namespace ClosedXmlTest
{

 
    public class BeneficioSaudeQarCreator: QarCreator
    {
        #region Endereços de Informação 

        const string 
            consultor = "B7", 
            razaoSocialEstipulante="B8",
            estipulanteCnpj="J8",
            operadoraAtual="B10",
            adaptadoLei="D10",
            tempoContrato="G10",
            agregados="I10",
            agregadosComplemento="J10",
            aniversarioContrato="B11",
            breakEven="D11",
            afastados="I11",
            afastadosComplemento="J11",
            possuiMulta="B12",
            regraMulta="D12",
            aposentadosPor="I12",    
            aposentadosPorComplemento="J12",    
            relatorioSinistralidade="B13",
            percentualSinistralidade="D13",
            gestantes="I13",
            gestantesComplemento="J13",    
            modalidadeContrato="B14",
            haDependentes="D14",
            cronicos="I14",    
            cronicosComplemento="J14",
            haReeembolso="B15",    
            remidos="I15",    
            remidosComplemento="J15",    
            coparticipacao="B16",
            regraCoparticacao="D16",
            inativos="I16",
            inativosComplemento="J16",
            regraUpDownGrade="B17",
            prestadorServico="I17",
            prestadorServicoComplemento="J17",
            contribuicaoTitular="B18",    
            contribuicaoTitularEmpresa="E18",        
            estagiarios="I18",
            estagiariosComplemento="J18",
            contribuicaoDependente="B19",    
            contribuicaoDependenteEmpresa="E19",        
            homeCare="I19",
            homeCareComplemento="J19",
            elegibilidade="B20",    
            elegibilidadeCnpj="J20",
            estrategiaCellRangeTemplate="A22:J30",
            subEstipulanteTitleCellRangeTemplate="A32:I32",
            subEstipulanteItemCellRangeTemplate="A33:I33",    
            estrategiaWorkSheetName = "ESTRATEGIA",
            baseDadosEstudosWorkSheetName = "BASE DE DADOS  ESTUDOS";                   

        const int estrategiaReferenceLine = 31;

        #endregion
        private int referenceLine;        

        private readonly IXLWorksheet _estrategiaWorkSheet;
        private readonly IXLWorksheet _baseDadosEstudosWorkSheet;    
        private readonly IXLRange _estrategiaBlockTemplate;
        private readonly IXLRange _subEstipulanteTitleBlockTemplate;             
        private readonly IXLRange _subEstipulanteItemBlockTemplate;                     
        public BeneficioSaudeQarCreator(Stream templateStream,string? outputFileAddress=null)
            :base(templateStream,outputFileAddress){
           _estrategiaWorkSheet = _workbook.Worksheets.First(x=>x.Name==estrategiaWorkSheetName);                
           _baseDadosEstudosWorkSheet = _workbook.Worksheets.First(x=>x.Name==baseDadosEstudosWorkSheetName);           
           _estrategiaBlockTemplate = _estrategiaWorkSheet.Range(estrategiaCellRangeTemplate);
           _subEstipulanteTitleBlockTemplate = _estrategiaWorkSheet.Range(subEstipulanteTitleCellRangeTemplate);          
           _subEstipulanteItemBlockTemplate = _estrategiaWorkSheet.Range(subEstipulanteItemCellRangeTemplate);                                    
        }

        public override MemoryStream GenerateExcelFile(){            
            BuildSectionFormularioContacaoSaude();
            BuildSectionEstrategia();
            BuildSectionSubEstipulante();
            BuildSectionBaseDadosEstudo();
            _estrategiaBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);            
            _subEstipulanteItemBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);            

            return base.GenerateExcelFile();
        }

        private void BuildSectionFormularioContacaoSaude()
        {
            _estrategiaWorkSheet.Cell(consultor).SetValue("José da Silva");
            _estrategiaWorkSheet.Cell(razaoSocialEstipulante).SetValue("Razao social Ficticia");
            _estrategiaWorkSheet.Cell(estipulanteCnpj).SetValue("91.786.878/0001-50");
            _estrategiaWorkSheet.Cell(operadoraAtual).SetValue("Bradesco");
            _estrategiaWorkSheet.Cell(adaptadoLei).SetValue("Sim");
            _estrategiaWorkSheet.Cell(tempoContrato).SetValue("2 anos");
            _estrategiaWorkSheet.Cell(agregados).SetValue("Sim");
            _estrategiaWorkSheet.Cell(agregadosComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(aniversarioContrato).SetValue(new DateTime(2025,01,31));
            _estrategiaWorkSheet.Cell(breakEven).SetValue("Break Even");
            _estrategiaWorkSheet.Cell(afastados).SetValue("Não");
            _estrategiaWorkSheet.Cell(afastadosComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(possuiMulta).SetValue("Sim");
            _estrategiaWorkSheet.Cell(regraMulta).SetValue("Alguma Regra");
            _estrategiaWorkSheet.Cell(aposentadosPor).SetValue("Sim");
            _estrategiaWorkSheet.Cell(aposentadosPorComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(relatorioSinistralidade).SetValue("Sim");
            _estrategiaWorkSheet.Cell(percentualSinistralidade).SetValue( 0.2d);
            _estrategiaWorkSheet.Cell(gestantes).SetValue("Sim");
            _estrategiaWorkSheet.Cell(gestantesComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(modalidadeContrato).SetValue("Compulsório");
            _estrategiaWorkSheet.Cell(haDependentes).SetValue("Sim");
            _estrategiaWorkSheet.Cell(cronicos).SetValue("Não");
            _estrategiaWorkSheet.Cell(cronicosComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(haReeembolso).SetValue("Sim");
            _estrategiaWorkSheet.Cell(remidos).SetValue("Não");
            _estrategiaWorkSheet.Cell(remidosComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(coparticipacao).SetValue("Sim");
            _estrategiaWorkSheet.Cell(regraCoparticacao).SetValue("Alguma Regra");
            _estrategiaWorkSheet.Cell(inativos).SetValue("Sim");
            _estrategiaWorkSheet.Cell( inativosComplemento).SetValue("Complementos");
            _estrategiaWorkSheet.Cell( regraUpDownGrade).SetValue("Alguma regra");
            _estrategiaWorkSheet.Cell( prestadorServico).SetValue("Não");
            _estrategiaWorkSheet.Cell( prestadorServicoComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(contribuicaoTitular).SetValue(0.1);
            _estrategiaWorkSheet.Cell(contribuicaoTitularEmpresa).SetValue( 0.9);
            _estrategiaWorkSheet.Cell( estagiarios).SetValue("Sim");
            _estrategiaWorkSheet.Cell( estagiariosComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell(contribuicaoDependente).SetValue(0.3);
            _estrategiaWorkSheet.Cell(contribuicaoDependenteEmpresa).SetValue(0.7);
            _estrategiaWorkSheet.Cell( homeCare).SetValue("Sim");
            _estrategiaWorkSheet.Cell( homeCareComplemento).SetValue("Complemento");
            _estrategiaWorkSheet.Cell( elegibilidade).SetValue("Sim");
            _estrategiaWorkSheet.Cell( elegibilidadeCnpj).SetValue("91.786.878/0001-50");
        }
        private void BuildSectionEstrategia()
        {

            List<EstrategiaSaude> estrategias = MockEstrategias();

            referenceLine=estrategiaReferenceLine;
            foreach(var estrategia in estrategias){
                _subEstipulanteTitleBlockTemplate.InsertRowsAbove(_estrategiaBlockTemplate.RowCount());
                _estrategiaBlockTemplate.CopyTo(_estrategiaWorkSheet.Cell(referenceLine,"A"));
                referenceLine+=2;
                _estrategiaWorkSheet.Cell(referenceLine+1,"J").SetValue(estrategia.Observacoes);
                for(int i=0;i<estrategia.Produtos.Count && i<8;i++){
                    var rowPosition= referenceLine;                    
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].Operadora);
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].Plano); 
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].ReembolsoConsulta);
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].Elegibilidade);
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].Vidas);
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].ValorPerCaptita);
                    _estrategiaWorkSheet.Cell(rowPosition++,i+2).SetValue(estrategia.Produtos[i].Coparticipacao);
                }
                referenceLine+=7;
            }            
        }
        private void BuildSectionSubEstipulante()
        {
            referenceLine+=_subEstipulanteTitleBlockTemplate.RowCount()+ _subEstipulanteItemBlockTemplate.RowCount() + 1;
            List<SubEstipulanteSaude> subEstipulantes = MockSubEstimulantes();
            foreach(var subEstipulante in subEstipulantes){
                _subEstipulanteItemBlockTemplate.CopyTo(_estrategiaWorkSheet.Cell(referenceLine,"A"));
                _estrategiaWorkSheet.Cell(referenceLine,"B").SetValue(subEstipulante.RazaoSocial);
                _estrategiaWorkSheet.Cell(referenceLine,"H").SetValue(subEstipulante.CNPJ);
                referenceLine++;                                 
            }
        }          
        private void BuildSectionBaseDadosEstudo()
        {
            _baseDadosEstudosWorkSheet.Cell("A2").InsertData(MockBaseDadosSaude());
        }

        #region Mock de Dados
        List<EstrategiaSaude> MockEstrategias(){
            return
            [
                 new (MockProdutos1(),"Observacao estrategia 1")
                ,new (MockProdutos2(),"Observacao estrategia 2")
                ,new (MockProdutos3(),"Observacao estrategia 3")
            ];
        }     

        List<ProdutoSaude> MockProdutos1(){
            return
            [
                 new ("Operadora A", "Plano A", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora B", "Plano B", "Sim","Elegibilidade","Vidas",10,0.15M)
                ,new ("Operadora C", "Plano C", "Sim","Elegibilidade","Vidas",10,0.3M)
            ];            
        }

        List<ProdutoSaude> MockProdutos2(){
            return
            [
                 new ("Operadora A", "Plano A", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora B", "Plano B", "Sim","Elegibilidade","Vidas",10,0.15M)
                ,new ("Operadora C", "Plano C", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora D", "Plano D", "Sim","Elegibilidade","Vidas",10,0.15M)                
                ,new ("Operadora E", "Plano E", "Sim","Elegibilidade","Vidas",10,0.2M)                
            ];            
        }

        List<ProdutoSaude> MockProdutos3(){
            return
            [
                 new ("Operadora A", "Plano A", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora B", "Plano B", "Sim","Elegibilidade","Vidas",10,0.15M)
                ,new ("Operadora C", "Plano C", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora D", "Plano D", "Sim","Elegibilidade","Vidas",10,0.15M)                
                ,new ("Operadora E", "Plano E", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora F", "Plano F", "Sim","Elegibilidade","Vidas",10,0.15M)
                ,new ("Operadora G", "Plano G", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora H", "Plano H", "Sim","Elegibilidade","Vidas",10,0.15M)
            ];            
        }


        List<PessoaSaude> MockBaseDadosSaude(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15", "Masculino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa da Maria", "41.646.207/0001-15", "Feminino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Roberto", "41.646.207/0001-15", "Masculino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do José", "41.646.207/0001-15", "Feminino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500), 
            ];
 
        }

        List<SubEstipulanteSaude> MockSubEstimulantes(){
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
    record EstrategiaSaude(List<ProdutoSaude> Produtos,string Observacoes);
    record ProdutoSaude(string Operadora, string Plano,string ReembolsoConsulta, string Elegibilidade, string Vidas,decimal ValorPerCaptita, decimal Coparticipacao);
    record PessoaSaude(string Empresa, string CNPJ, string Sexo, string Identificacao, DateTime DataNascimento, int Idade, string FaixaEtaria, string Parentesco, string Situacao, string CID, string Municipio, string UF, string Operadora, string Plano, int ValorAtual);
    record SubEstipulanteSaude(string RazaoSocial, string CNPJ);    
    #endregion
    
}