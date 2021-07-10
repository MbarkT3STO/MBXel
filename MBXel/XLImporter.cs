using LinqToExcel;

using System.Linq;
using System.Threading.Tasks;

namespace MBXel
{

    /// <summary>
    /// Import data from an Excel file
    /// </summary>
    public partial class XLImporter
    {

        #region Private methods

        private IQueryable<Row> _Import(string filePath, string sheetName)
        {
            //Load the workbook
            var Wbook = new ExcelQueryFactory(filePath);

            //Collect data from the worksheet
            var R = Wbook.Worksheet(sheetName);

            return R;
        }

        #endregion


        /// <summary>
        /// Load an excel file (Sheet) data
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetName">The Worksheet name to be recovered data from</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public IQueryable<Row> Import(string filePath, string sheetName)
        {
            return _Import(filePath, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="Import(string, string)"/>
        /// </summary>
        /// <inheritdoc cref="Import(string, string)"/>
        /// <returns><see cref="Task{IQueryable{Row}}"/></returns>
        public Task<IQueryable<Row>> ImportAsync(string filePath, string sheetName)
        {
            return Task.Factory.StartNew(() => _Import(filePath, sheetName));
        }

    }
}
