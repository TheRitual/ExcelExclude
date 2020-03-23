using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using msExcel = Microsoft.Office.Interop.Excel;

namespace ExcelExclude
{
    class ExcelFile
    {
        private String path;
        private int sheet;

        public ExcelFile(String path)
        {
            this.Path = path;
        }

        public string Path { get => path; set => path = value; }
        public int Sheet { get => sheet; set => sheet = value; }
    }
}
