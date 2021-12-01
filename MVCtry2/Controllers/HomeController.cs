using ClosedXML.Excel;
using LazyCache;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using MVCtry2.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;


namespace MVCtry2.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _env;
        private string _dir;
        private readonly IHomeServices _services;
        private readonly IMemoryCache _cache;
        private string _userIp; 

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env, IHomeServices services, IMemoryCache memoryCache)
        {
            _logger = logger;
            _env = env;
            _dir = _env.ContentRootPath;
            _services = services;
            _cache = memoryCache;           
        }

        public FileResult DownloadFile2(string fileName)
        {
            _userIp = HttpContext.Connection.RemoteIpAddress.ToString();
            using (MemoryStream stream = new MemoryStream())
            {
                var resultHeader = (List<string>)_cache.Get(_userIp+"head");
                var resultFindedData = (List<string>)_cache.Get(_userIp+"finded");
                var buf = _services.CreateWorkbook(fileName,resultHeader,resultFindedData);
                buf.SaveAs(stream);

                //Return xlsx Excel File  
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
        }
        private List<string> headerList = new List<string>();
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult SearchSerialFromDataList(FormModel someForm)
        {
            _userIp = HttpContext.Connection.RemoteIpAddress.ToString();
            var result2 = (List<string>)_cache.Get(_userIp+"data");
          //  List<string> buf = (List<string>)result2;
            var viewModel = new FormModel();

            var result = _services.FindFileFromDataList(someForm.StringToFind, result2);
            if(result != null)
            {
                TempData.Clear();
                _cache.Set(_userIp+"finded", result, DateTimeOffset.Now.AddMinutes(10));
                DateTime localDate = DateTime.Now;
                var nameOfFile = someForm.StringToFind + @localDate.ToString("_yyyy'.'MM'.'dd'_['HH'.'mm'.'ss]") + @".xlsx";

                TempData["FileReturned"] = nameOfFile;
                viewModel.Name = nameOfFile;
            }
            else
            {
                TempData["FileReturned"] = "Error";
                _cache.Remove("finded");
                return RedirectToAction("Index");
            }

            return View("Index", viewModel);
        }

        private List<string> headList = new List<string>();
        public IActionResult HeaderFile(FormModel headerFile)
        {
            _userIp = HttpContext.Connection.RemoteIpAddress.ToString();
            if (headerFile.File != null)
            {
                var result = _services.HeaderFile(headerFile.File, _dir);

                _cache.Set(_userIp+"head", result, DateTimeOffset.Now.AddMinutes(10));

                TempData.Keep("data");
                TempData["header"] = headerFile.File.FileName;
            }
            else
            {
                TempData.Keep("data");
                TempData["header"] = "Nie podano pliku! Spróbuj jeszcze raz";
            }
            return View("Index");

        }
        private List<string> dataList = new List<string>();
        public IActionResult DataFile(FormModel DataFile)
        {
            _userIp = HttpContext.Connection.RemoteIpAddress.ToString();
            if (DataFile.File != null)
            {
                var result =_services.DataFile(DataFile.File, _dir);
                dataList = result;

                _cache.Set(_userIp+"data", dataList, DateTimeOffset.Now.AddMinutes(10));

                TempData.Keep("header");
                TempData["data"] = DataFile.File.FileName;
            }
            else
            {
                TempData.Keep("header");
                TempData["data"] = "Nie podano pliku! Spróbuj jeszcze raz";
            }

            return View("Index");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
