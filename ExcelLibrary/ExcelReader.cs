using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
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
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);
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
        /// <summary>
        /// Releases COM Objects.  Credit: http://www.claudiobernasconi.ch/2014/02/13/painless-office-interop-using-visual-c-sharp/
        /// </summary>
        /// <param name="obj">COM Object to release</param>
        public static void ReleaseCOMObject(object obj)
        {
            if(obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnProcessExit(object sender, EventArgs e)
        {
            ReleaseCOMObject(this.ExcelWorkSheet);
            ReleaseCOMObject(this.ExcelWorkbook);

            this.ExcelWorkSheet = null;
            this.ExcelWorkbook = null;

            for (int i = 0; i < this.ExcelWorkSheets.Count; i++) {
                ReleaseCOMObject(ExcelWorkSheets[i]);
                ExcelWorkSheets[i] = null;
            }

            if(this.ExcelApplication != null)
            {
                this.ExcelApplication.Quit();
            }

            ReleaseCOMObject(ExcelApplication);

            ExcelApplication = null;
        }


    }
}
