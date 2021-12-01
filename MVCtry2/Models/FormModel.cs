using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MVCtry2.Models
{
    public class FormModel
    {
        public string Name { get ; set; }
        public string StringToFind { get; set; }
        public IFormFile File { get; set; }

      //  public List<string> dataList = new List<string>();

    }
}
