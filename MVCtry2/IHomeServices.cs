using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace MVCtry2
{
    public interface IHomeServices
    {
        List<string> DataFile(IFormFile file, string _dir);
        List<string> HeaderFile(IFormFile file, string _dir);
        List<string> FindFileFromDataList(string serialNumber, List<string> data);
        public XLWorkbook CreateWorkbook(string serialNumber, List<string> head, List<string> data);
    }
}