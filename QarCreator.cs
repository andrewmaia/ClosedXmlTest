using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ClosedXmlTest
{
    public abstract class QarCreator
    {
        protected readonly XLWorkbook _workbook;
        protected readonly string? _outputFileAddress;
        public QarCreator(Stream templateStream,string? outputFileAddress=null){
           _workbook= new XLWorkbook(templateStream);
           _outputFileAddress = outputFileAddress;
        }
        public  virtual MemoryStream GenerateExcelFile(){
            if(!string.IsNullOrEmpty(_outputFileAddress))
                _workbook.SaveAs(_outputFileAddress);

            MemoryStream outputFileStream= new();
            _workbook.SaveAs(outputFileStream);
            
            return outputFileStream;            
        }
    }
}