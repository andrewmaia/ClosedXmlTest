using ClosedXmlTest;

// Beneficios - Saude

const string beneficioSaudeTemplateAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarSaude_01_Dinamico.xlsx"; 
const string beneficioSaudeOutputFileAddress = @"C:\Users\andrew.maia\Desktop\QAR\QarSaude_03_Preenchido.xlsx"; 
using FileStream templateStream = new (beneficioSaudeTemplateAddress,FileMode.Open);
var  bs= new BeneficioSaudeQarGenerator(templateStream,beneficioSaudeOutputFileAddress);
using Stream beneficioSaudeSpreadSheet = bs.GenerateExcelFile();


