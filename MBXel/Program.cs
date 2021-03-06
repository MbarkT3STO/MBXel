using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using LinqToExcel;
using MBXel.Enum;
using Microsoft.Office.Interop.Excel;

namespace MBXel
{
    class Program
    {

        private static readonly List<Order> Orders = new List<Order>
                                                     {
                                                         new Order(1, "Ennasiri Ali", "PRD-1", 1500),
                                                         new Order(2, "Badaoui Inas", "PRD-1", 2000),
                                                         new Order(3, "Baddouh Ali", "PRD-3", 1000),
                                                         new Order(4, "Mouslim Kawtar", "PRD-2", 3500),
                                                         new Order(5, "Essalmi Karim", "PRD-1", 2000),
                                                         new Order(6, "Nousayr Ahmed", "PRD-1", 2000),
                                                         new Order(7, "Mersaoui Fatima", "PRD-3", 1000),
                                                         new Order(8, "Fanar Adil", "PRD-1", 2200),
                                                         new Order(9, "Eddawdi Nawal", "PRD-2", 3200),
                                                         new Order(10, "Houmam Karim", "PRD-1", 2400),
                                                         new Order(11, "Ennasiri Ali", "PRD-2", 2000),
                                                         new Order(12, "Ennasiri Ali", "PRD-3", 3500),
                                                         new Order(13, "Essalmi Karim", "PRD-2", 1500),
                                                         new Order(14, "Eddawdi Nawal", "PRD-1", 2000)
                                                     };


        static async System.Threading.Tasks.Task Main()
        {

            //---------------------------------------------------------------------------------------------------------
            //Examples
            //---------------------------------------------------------------------------------------------------------
           


            /*--------------------*/
            /*Export data*/
            /*--------------------*/

            //XLExporter exporter = new XLExporter();

            //await exporter.ExportAsync( Orders , Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\XXXX" , Enums.XLExtension.Xlsx );
            //Console.WriteLine( "Saved" );


            /*--------------------*/
            /*Import data*/
            /*--------------------*/

            //var        importer = new XLImporter();

            //var Wbook = await importer.ImportAsync( Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\XXXX.xlsx" , "Feuil1" );


            /*--------------------*/
            /*Use LINQ with the imported data*/
            /*--------------------*/
            //var R = (from x in Wbook select new { ID = x["ID"] , Client = x["Client"] , Product = x["Product"] , Total = x["Total"] }).ToList();

            //R.ForEach( x =>
            //           {
            //               Console.WriteLine( $"{x.ID}, {x.Client}, {x.Product}, {x.Total}" );
            //           } );
            

            Console.ReadKey();
        }


    }
}
