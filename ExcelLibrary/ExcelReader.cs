using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelLibrary
{
    public class ExcelReader
    {
        #region Private Vars
        private String FileLocation { get; set; }
        private Excel.Workbook ExcelWorkbook { get; set; }
        private Excel.Application ExcelApplication { get; set; }
        private Excel.Worksheet ExcelWorkSheet { get; set; }
        private List<Excel.Worksheet> ExcelWorkSheets { get; set;}
        #endregion 
        /// <summary>
        /// Constructor of Excel Reader.  
        /// </summary>
        /// <param name="AbsoluteFileLocation">File location of your excel workbook.  Must be xls format.</param>
        public ExcelReader(String AbsoluteFileLocation)
        {
            this.FileLocation = AbsoluteFileLocation;
            Initiate();
        }

        /// <summary>
        /// Opens the Excel document and sets class variables.
        /// </summary>
        private void Initiate()
        {
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelWB = null;

            try
            {
                ExcelWB = ExcelApp.Workbooks.Open(this.FileLocation, Type.Missing, Type.Missing,
                                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                  Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine("Excel Workbook opened");
            } 
            catch (FileNotFoundException fex)
            {
                Console.WriteLine("File not found at: " + this.FileLocation);
                Console.Write(fex.StackTrace);
            } 
            catch (Exception ex)
            {
              
                Console.Write(ex.StackTrace);
            }
            if (ExcelWB != null)
            {
                this.ExcelWorkbook = ExcelWB;
                this.ExcelWorkSheet = ExcelWorkbook.Worksheets.get_Item(1);
                Console.WriteLine("Extracted " + ExcelWorkSheet.Name + " From workbook.");
            }
            else
            {
                Console.Error.WriteLine("Error opening excel file.  Exiting");
                Environment.Exit(1);
            }
            
           

        }

        /// <summary>
        /// Hands off the current worksheet.
        /// </summary>
        /// <returns>Excel worksheet from the excel workbook.</returns>
        public Worksheet getWorkSheet()
        {
            if(ExcelWorkSheet == null)
            {
                throw new NullReferenceException();
            }
            else
            {
                return this.ExcelWorkSheet;
            }
        }

        public void close()
        {
            
        }


    }
}
