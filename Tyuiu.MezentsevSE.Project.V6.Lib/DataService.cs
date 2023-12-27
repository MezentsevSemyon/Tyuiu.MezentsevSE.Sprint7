using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
namespace Tyuiu.MezentsevSE.Project.V6.Lib
{
    public class DataService
    {
        private DataTableCollection tableCollection = null;
        public void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration() 
            {
                 ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                 {
                     UseHeaderRow = true
                 }
            
            
            
            });
            tableCollection = db.Tables;

            



        }


        
    }
}

