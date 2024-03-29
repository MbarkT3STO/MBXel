﻿using MBXel.Enum;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace MBXel
{
    /// <summary>
    /// Export an Excel file
    /// </summary>
    public partial class XLExporter
    {

        #region Private methods

        private void _StylingTheWorkSheet(ref Excel.Worksheet WSheet, int ColumnsNumber, int RowsNumber)
        {
            //Columns styling
            WSheet.Range["A1", "BB1"].Cells.Font.Size = 20;
            WSheet.Range["A1", "BB1"].Cells.Font.Bold = true;
            WSheet.Range["A1", "BB1"].Cells.Font.Color = Color.White;
            WSheet.Range["A1", "BB1"].Cells.Interior.Color = ColorTranslator.FromHtml("#5352ed");
            WSheet.Range["A1", "BB1"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            WSheet.Range["A1", "BB1"].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            //Rows styling
            WSheet.Range["A2", $"BB{RowsNumber + 1}"].Cells.Font.Size = 14;
            WSheet.Range["A2", $"BB{RowsNumber + 1}"].Cells.Font.Color = Color.White;
            WSheet.Range["A2", $"BB{RowsNumber + 1}"].Cells.Interior.Color = ColorTranslator.FromHtml("#2ed573");
            WSheet.Range["A2", $"BB{RowsNumber + 1}"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            WSheet.Range["A2", $"BB{RowsNumber + 1}"].Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            //Other Columns styling
            WSheet.Columns.ColumnWidth = 20;
        }


        private async void _Export<T>(List<T> data, string path, Enums.XLExtension extension)
        {
            //Get data parameter type properties
            PropertyInfo[] data_Properties = typeof(T).GetProperties();
            var data_PropertiesCount = data_Properties.Length;

            //Prepaire the Excel app
            Excel.Application XlApp = new Excel.ApplicationClass();

            try
            {
                //Prepaire the Workbook and Worksheet
                var Wbook = XlApp.Workbooks.Add();
                var Wsheet = (Excel.Worksheet)Wbook.Sheets[1];

                //Put data into worksheet
                await Task.Run(() =>
                {
                    for (int i = 0; i < data_PropertiesCount; i++)
                    {
                        Wsheet.Cells[1, i + 1] = data_Properties[i].Name;
                    }

                    int rowIndex = 2;

                    foreach (T d in data)
                    {
                        for (int i = 0; i < data_PropertiesCount; i++)
                        {
                            Wsheet.Cells[rowIndex, i + 1] = data_Properties[i].GetValue(d);
                        }

                        rowIndex++;
                    }
                });

                //Styling the worksheet
                await Task.Run(() => _StylingTheWorkSheet(ref Wsheet, data_PropertiesCount, data.Count));

                //Save the workbook 
                Wbook.SaveAs(path + (extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls"));

                //Close all workbooks
                XlApp.Workbooks.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                //Quit from XlApp
                XlApp.Quit();

                //Kill all the opened untitled EXCEL processes
                Process[] processes = Process.GetProcessesByName("EXCEL");

                foreach (Process p in processes)
                {
                    if (p.MainWindowTitle.Length == 0)
                    {
                        p.Kill();
                    }
                }
            }
        }


        #endregion


        /// <summary>
        /// Export a <see cref="List{T}"/> of data to an excel file
        /// </summary>
        /// <param name="data">Data to be exported</param>
        /// <param name="path">Path to be save in</param>
        /// <param name="extension">Excel file extension</param>
        /// <returns><see cref="bool"/></returns>
        public bool Export<T>(List<T> data, string path, Enums.XLExtension extension)
        {
            _Export(data, path, extension);

            return true;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension)"/>
        /// </summary>
        /// <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension)"/>
        /// <returns><see cref="Task{TResult}"/></returns>
        public Task<bool> ExportAsync<T>(List<T> data, string path, Enums.XLExtension extension)
        {
            return Task.Factory.StartNew(() =>
                                         {
                                             _Export(data, path, extension);
                                             return true;
                                         });
        }
    }
}