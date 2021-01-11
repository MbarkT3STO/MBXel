using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using MBXel.Enum;
using Excel =Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Threading.Tasks;

namespace MBXel
{
        /// <summary>
        /// Export an Excel file
        /// </summary>
        public partial class XLExporter
        {

            private void _StylingTheWorkSheet(ref Excel.Worksheet WSheet, int ColumnsNumber, int RowsNumber)
            {
                //Columns styling
                WSheet.Range["A1","AZ1"].Cells.Font.Size            = 20;
                WSheet.Range["A1","AZ1"].Cells.Font.Bold            = true;
                WSheet.Range["A1", "AZ1"].Cells.Font.Color          = Color.White;
                WSheet.Range["A1", "AZ1"].Cells.Interior.Color      = ColorTranslator.FromHtml("#5352ed");
                WSheet.Range["A1", "AZ1"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                WSheet.Range["A1", "AZ1"].Cells.VerticalAlignment   = Excel.XlVAlign.xlVAlignCenter;


                //Rows styling
                WSheet.Range["A2", "AZ" + RowsNumber +2].Cells.Font.Size           = 14;
                WSheet.Range["A2", "AZ" + RowsNumber +2].Cells.Font.Color          = Color.White;
                WSheet.Range["A2", "AZ" + RowsNumber +2].Cells.Interior.Color      = ColorTranslator.FromHtml("#2ed573");
                WSheet.Range["A2", "AZ" + RowsNumber +2].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                WSheet.Range["A2", "AZ" + RowsNumber + 2].Cells.VerticalAlignment  = Excel.XlVAlign.xlVAlignCenter;

                
                //Other Columns styling
                WSheet.Columns.ColumnWidth = 20;



            }





            private async void _Export<T>(List<T> data, string path, Enums.XLExtension extension)
            {

                    //Get data parameter type properties
                    PropertyInfo[] data_Properties      = typeof(T).GetProperties();
                    var            data_PropertiesCount = data_Properties.Length;


                    //Prepaire the Excel app
                    Excel.Application XlApp = new Excel.ApplicationClass();

                    try
                    {
                        //Prepaire the Workbook and Worksheet
                        var Wbook  = XlApp.Workbooks.Add();
                        var Wsheet = (Excel.Worksheet) Wbook.Sheets[1];


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
                        await Task.Run( () => _StylingTheWorkSheet( ref Wsheet , data_PropertiesCount , data.Count ) );



                        //Save the workbook 
                        Wbook.SaveAs(path + (extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls"));

                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    finally
                    {
                        //Quit from XlApp
                        XlApp.Quit();
                    }

            }














            /// <summary>
            /// Export a data to excel file and save it
            /// </summary>
            /// <param name="data">Data to be exported</param>
            /// <param name="path">Path to be save in</param>
            /// <param name="extension">Excel file extension</param>
            /// <returns></returns>
            public bool Export<T>(List<T> data, string path, Enums.XLExtension extension )
            {

                    _Export(data, path, extension);

                    return true;

            }
            
            

            /// <summary>
            /// Export a data to excel file and save it, Asynchronously
            /// </summary>
            /// <param name="data">Data to be exported</param>
            /// <param name="path">Path to be save in</param>
            /// <param name="extension">Excel file extension</param>
            /// <returns></returns>
            public Task<bool> ExportAsync<T>(List<T> data, string path, Enums.XLExtension extension )
            {

                return Task.Factory.StartNew(() =>
                                             {
                                                 _Export(data, path, extension);
                                                 return true;
                                             });
                }

            



        }
}
