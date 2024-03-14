using System.Text;
using ClosedXmlTest;

// Beneficio - Saude
/*const string beneficioSaudeTemplateAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarSaude_01_Dinamico.xlsx"; 
const string beneficioSaudeOutputFileAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarSaude_03_Preenchido.xlsx"; 
using FileStream beneficioSaudeTemplateStream = new (beneficioSaudeTemplateAddress,FileMode.Open);
QarCreator  bs= new BeneficioSaudeQarCreator(beneficioSaudeTemplateStream,beneficioSaudeOutputFileAddress);
using Stream beneficioSaudeSpreadSheet = bs.GenerateExcelFile();*/


// Beneficio - Odonto
/*const string beneficioOdontoTemplateAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarOdonto_01_Dinamico.xlsx"; 
const string beneficioOdontoOutputFileAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarOdonto_02_Preenchido.xlsx"; 
using FileStream beneficioOdontoTemplateStream = new(beneficioOdontoTemplateAddress,FileMode.Open);
QarCreator  bo= new BeneficioOdontoQarCreator(beneficioOdontoTemplateStream,beneficioOdontoOutputFileAddress);
using Stream beneficioOdontoSpreadSheet = bo.GenerateExcelFile();*/

// Beneficio - Vida
const string beneficioVidaTemplateAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarVida_02_Dinamico.xlsx"; 
const string beneficioVidaOutputFileAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarVida_03_Preenchido.xlsx"; 
using FileStream beneficioVidaTemplateStream = new(beneficioVidaTemplateAddress,FileMode.Open);
QarCreator  bv= new BeneficioVidaQarCreator(beneficioVidaTemplateStream,beneficioVidaOutputFileAddress);
using Stream beneficioVidaSpreadSheet = bv.GenerateExcelFile();

