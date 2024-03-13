using ClosedXML.Excel;
using ClosedXmlTest;
using DocumentFormat.OpenXml.InkML;

namespace ClosedXmlTest
{

 
    public class BeneficioSaudeTest
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
    
        public  static MemoryStream GenerateExcelFile(string templateAddress,string? outputFileAddress=null){
            using FileStream fs = new (templateAddress,FileMode.Open);
            return GenerateExcelFile(fs,outputFileAddress);
        }
        public  static MemoryStream GenerateExcelFile(Stream templateStream,string? outputFileAddress=null){            
 
            var workbook = new XLWorkbook(templateStream);
            var wsEstrategia = workbook.Worksheets.First(x=>x.Name==estrategiaWorkSheetName);

            wsEstrategia.Cell(consultor).SetValue("José da Silva");
            wsEstrategia.Cell(razaoSocialEstipulante).SetValue("Razao social Ficticia");
            wsEstrategia.Cell(estipulanteCnpj).SetValue("91.786.878/0001-50");
            wsEstrategia.Cell(operadoraAtual).SetValue("Bradesco");
            wsEstrategia.Cell(adaptadoLei).SetValue("Sim");
            wsEstrategia.Cell(tempoContrato).SetValue("2 anos");
            wsEstrategia.Cell(agregados).SetValue("Sim");
            wsEstrategia.Cell(agregadosComplemento).SetValue("Complemento");
            wsEstrategia.Cell(aniversarioContrato).SetValue(new DateTime(2025,01,31));
            wsEstrategia.Cell(breakEven).SetValue("Break Even");
            wsEstrategia.Cell(afastados).SetValue("Não");
            wsEstrategia.Cell(afastadosComplemento).SetValue("Complemento");
            wsEstrategia.Cell(possuiMulta).SetValue("Sim");
            wsEstrategia.Cell(regraMulta).SetValue("Alguma Regra");
            wsEstrategia.Cell(aposentadosPor).SetValue("Sim");
            wsEstrategia.Cell(aposentadosPorComplemento).SetValue("Complemento");
            wsEstrategia.Cell(relatorioSinistralidade).SetValue("Sim");
            wsEstrategia.Cell(percentualSinistralidade).SetValue( 0.2d);
            wsEstrategia.Cell(gestantes).SetValue("Sim");
            wsEstrategia.Cell(gestantesComplemento).SetValue("Complemento");
            wsEstrategia.Cell(modalidadeContrato).SetValue("Compulsório");
            wsEstrategia.Cell(haDependentes).SetValue("Sim");
            wsEstrategia.Cell(cronicos).SetValue("Não");
            wsEstrategia.Cell(cronicosComplemento).SetValue("Complemento");
            wsEstrategia.Cell(haReeembolso).SetValue("Sim");
            wsEstrategia.Cell(remidos).SetValue("Não");
            wsEstrategia.Cell(remidosComplemento).SetValue("Complemento");
            wsEstrategia.Cell(coparticipacao).SetValue("Sim");
            wsEstrategia.Cell(regraCoparticacao).SetValue("Alguma Regra");
            wsEstrategia.Cell(inativos).SetValue("Sim");
            wsEstrategia.Cell( inativosComplemento).SetValue("Complementos");
            wsEstrategia.Cell( regraUpDownGrade).SetValue("Alguma regra");
            wsEstrategia.Cell( prestadorServico).SetValue("Não");
            wsEstrategia.Cell( prestadorServicoComplemento).SetValue("Complemento");
            wsEstrategia.Cell(contribuicaoTitular).SetValue(0.1);
            wsEstrategia.Cell(contribuicaoTitularEmpresa).SetValue( 0.9);
            wsEstrategia.Cell( estagiarios).SetValue("Sim");
            wsEstrategia.Cell( estagiariosComplemento).SetValue("Complemento");
            wsEstrategia.Cell(contribuicaoDependente).SetValue(0.3);
            wsEstrategia.Cell(contribuicaoDependenteEmpresa).SetValue(0.7);
            wsEstrategia.Cell( homeCare).SetValue("Sim");
            wsEstrategia.Cell( homeCareComplemento).SetValue("Complemento");
            wsEstrategia.Cell( elegibilidade).SetValue("Sim");
            wsEstrategia.Cell( elegibilidadeCnpj).SetValue("91.786.878/0001-50");

            List<Estrategia> estrategias = MockEstrategias();
            IXLRange estrategiaBlockTemplate = wsEstrategia.Range(estrategiaCellRangeTemplate);
            IXLRange subEstipulanteTitleBlockTemplate = wsEstrategia.Range(subEstipulanteTitleCellRangeTemplate);
            IXLRange subEstipulanteItemBlockTemplate = wsEstrategia.Range(subEstipulanteItemCellRangeTemplate);

            int referenceLine=estrategiaReferenceLine;
            foreach(var estrategia in estrategias){
                subEstipulanteTitleBlockTemplate.InsertRowsAbove(9);
                estrategiaBlockTemplate.CopyTo(wsEstrategia.Cell(referenceLine,"A"));
                referenceLine+=2;
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.Operadora);
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.Plano); 
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.ReembolsoConsulta); 
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.Elegibilidade);
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.Vidas); 
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.ValorPerCaptita);
                wsEstrategia.Cell(referenceLine++,"B").SetValue(estrategia.Coparticipacao);
            }

            referenceLine+=3;
            List<SubEstipulante> subEstipulantes = MockSubEstimulantes();
            foreach(var subEstipulante in subEstipulantes){
                subEstipulanteItemBlockTemplate.CopyTo(wsEstrategia.Cell(referenceLine,"A"));
                wsEstrategia.Cell(referenceLine,"B").SetValue(subEstipulante.RazaoSocial);
                wsEstrategia.Cell(referenceLine,"H").SetValue(subEstipulante.CNPJ);
                referenceLine++;                                 
            }

            subEstipulanteItemBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);
            estrategiaBlockTemplate.Delete(XLShiftDeletedCells.ShiftCellsUp);

           
            //ABA BASE DE DADOS  ESTUDOS
            var wsBaseDados = workbook.Worksheets.First(x=>x.Name==baseDadosEstudosWorkSheetName);
            wsBaseDados.Cell("A2").InsertData(MockBaseDadosSaude());

            if(!string.IsNullOrEmpty(outputFileAddress))
                workbook.SaveAs(outputFileAddress);

            MemoryStream outputFileStream= new();
            workbook.SaveAs(outputFileStream);

            
            return outputFileStream;
        }
        
        #region Mock de Dados
        static List<Estrategia> MockEstrategias(){
            return
            [
                 new ("Operadora", "Plano A", "Sim","Elegibilidade","Vidas",10,0.2M)
                ,new ("Operadora", "Plano B", "Sim","Elegibilidade","Vidas",10,0.15M)
                ,new ("Operadora", "Plano C", "Sim","Elegibilidade","Vidas",10,0.3M)
            ];
        }          

         static List<Pessoa> MockBaseDadosSaude(){
            return
            [
                new("Empresa do Joao", "41.646.207/0001-15", "Masculino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa da Maria", "41.646.207/0001-15", "Feminino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do Roberto", "41.646.207/0001-15", "Masculino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa do José", "41.646.207/0001-15", "Feminino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500), 
            ];
 
        }

        static List<SubEstipulante> MockSubEstimulantes(){
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
    record Estrategia(string Operadora, string Plano,string ReembolsoConsulta, string Elegibilidade, string Vidas,decimal ValorPerCaptita, decimal Coparticipacao);
    record Pessoa(string Empresa, string CNPJ, string Sexo, string Identificacao, DateTime DataNascimento, int Idade, string FaixaEtaria, string Parentesto, string Situacao, string CID, string Municipio, string UF, string Operadora, string Plano, int ValorAtual);
    record SubEstipulante(string RazaoSocial, string CNPJ);    
    #endregion
    
}