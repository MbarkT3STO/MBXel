using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace MBXel
{
    class XLReader
    {

        private IQueryable<Row> _Read(string filePath, string sheetName)
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
        public IQueryable<Row> Read(string filePath, string sheetName)
        {
            return _Read(filePath, sheetName);
        }



        /// <summary>
        /// Load an excel file (Sheet) data, Asynchronously
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetName">The Worksheet name to be recovered data from</param>
        /// <returns></returns>
        public Task<IQueryable<Row>> ReadAsync(string filePath, string sheetName)
        {
            return Task.Factory.StartNew(() => _Read(filePath, sheetName));
        }

    }
}
