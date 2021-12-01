using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;


namespace MVCtry2
{
    public class HomeServices  :  IHomeServices
    {
        
        private List<string> headerList = new List<string>();
        private List<string> dataList = new List<string>();
        private List<string> findedList = new List<string>();

        public XLWorkbook CreateWorkbook(string? serialNumber, List<string> header, List<string> data)
        {
            if(serialNumber[0].Equals('H'))
                return Excel.ExcelCreateTranspondedFile(header, data);
            else
                return Excel.ExcelCreateFile(header, data);
        }


        private List<string> SplitString(string serialNumber, List<string> data)
        {
            findedList.Clear();
            var ListOfStringsToFind = serialNumber.Split(',');
            foreach (var singleString in ListOfStringsToFind)
            {
                if(singleString.Length > 1)
                    findedList.AddRange(data.Where(x => x.Contains(singleString)).Select(x => Regex.Replace(x, @"\s+", string.Empty)));
            }
            return findedList;
        }

        public List<string> FindFileFromDataList(string serialNumber, List<string> data)
        {
            if(serialNumber != null && data != null)
            {
                if (serialNumber.Contains(","))
                    findedList = SplitString(serialNumber, data);
                else
                    findedList = data.Where(x => x.Contains(serialNumber)).ToList();
                if (findedList.Count > 0)
                {
                    return findedList;
                }
            }
            
            return null;
        }

        public List<string> HeaderFile(IFormFile file, string _dir)
        {
            headerList.Clear();
            //using (var fileStream = new FileStream(Path.Combine(_dir, "Files", "header.txt"), FileMode.Create, FileAccess.Write))
            //{
            //    file.CopyTo(fileStream);
            //}
            using (StreamReader sr = new StreamReader(file.OpenReadStream()))
            {
                while (sr.Peek() >= 0)
                {
                    headerList.Add(sr.ReadLine());
                }
            }
            return headerList;
            
        }

        public List<string> DataFile(IFormFile file, string _dir)
        {
            dataList.Clear();
            //using (var fileStream = new FileStream(Path.Combine(_dir, "Files", "data.txt"), FileMode.Create, FileAccess.Write))
            //{
            //    file.CopyTo(fileStream);
            //}

            using (StreamReader sr = new StreamReader(file.OpenReadStream()))
            {
                while (sr.Peek() >= 0)
                {
                    dataList.Add(sr.ReadLine());
                }

            }
            return dataList;
            //    Current.Session["list"] = dataList;
        }

    }
}
