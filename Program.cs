using ClosedXML.Excel;
using ClosedXmlTest;


using (var workbook = new XLWorkbook(@"C:\Users\andrew.maia\Desktop\QAR\QarSaude_CamposFixos.xlsx"))
{

    var wsEstrategia = workbook.Worksheets.First(x=>x.Name=="ESTRATEGIA");
    wsEstrategia.Cell("B10").SetValue(50);
    wsEstrategia.Cell("B12").SetValue("Sim");
    wsEstrategia.Cell("B11").SetValue(new DateTime(2000,01,01));    
    wsEstrategia.Cell("E18").SetValue(0.5); 

    var wsBaseDados = workbook.Worksheets.First(x=>x.Name=="BASE DE DADOS  ESTUDOS");
    int linha=2;
    foreach( var p in Mock.MockarBaseDadosSaude()){
        wsBaseDados.Cell(linha,1).SetValue(p.Empresa);
        wsBaseDados.Cell(linha,2).SetValue(p.CNPJ);
        wsBaseDados.Cell(linha,3).SetValue(p.Sexo);
        wsBaseDados.Cell(linha,4).SetValue(p.Identificacao);
        wsBaseDados.Cell(linha,5).SetValue(p.DataNascimento);
        wsBaseDados.Cell(linha,6).SetValue(p.Idade);        
        wsBaseDados.Cell(linha,7).SetValue(p.FaixaEtaria);
        wsBaseDados.Cell(linha,8).SetValue(p.Parentesto);
        wsBaseDados.Cell(linha,9).SetValue(p.Situacao);
        wsBaseDados.Cell(linha,10).SetValue(p.CID);
        wsBaseDados.Cell(linha,11).SetValue(p.Municipio);
        wsBaseDados.Cell(linha,12).SetValue(p.UF);
        wsBaseDados.Cell(linha,13).SetValue(p.Operadora);
        wsBaseDados.Cell(linha,14).SetValue(p.Plano);
        wsBaseDados.Cell(linha,15).SetValue(p.ValorAtual);        
        linha++;
    }

    workbook.SaveAs(@"C:\Users\andrew.maia\Desktop\QAR\HelloWorld.xlsx");
}