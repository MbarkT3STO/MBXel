using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace MBXel
{

    /// <summary>
    /// Import data from an Excel file
    /// </summary>
    public partial class XLImporter
    {

        private IQueryable<Row> _Import(string filePath, string sheetName)
        {
            //Load the workbook
            var Wbook = new ExcelQueryFactory(filePath);

            //Collect data from the worksheet
            var R = Wbook.Worksheet(sheetName);

            return R;
        }






        /// <summary>
        /// Load an excel file (Sheet) data
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetName">The Worksheet name to be recovered data from</param>
        /// <returns></returns>
        public IQueryable<Row> Import(string filePath, string sheetName)
        {
            return _Import(filePath, sheetName);
        }



        /// <summary>
        /// Load an excel file (Sheet) data, Asynchronously
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetName">The Worksheet name to be recovered data from</param>
        /// <returns></returns>
        public Task<IQueryable<Row>> ImportAsync(string filePath, string sheetName)
        {
            return Task.Factory.StartNew(() => _Import(filePath, sheetName));
        }

    }
}
