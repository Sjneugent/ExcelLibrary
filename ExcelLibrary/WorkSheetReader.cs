using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ExcelLibrary
{
    public class WorkSheetReader
    {
        #region Private Variables
        private Excel.Worksheet workSheet { get; set; }
        private Excel.Range xlsRange { get; set; }
        private int _usedColumns { get; set; }
        private int _usedRows { get; set; }
        private List<ValueLocation> locations { get; set; }
        private List<String> location_toStr { get; set; }
        #endregion 

        /// <summary>
        /// Initiates the class.
        /// </summary>
        /// <param name="workSheet"></param>
        public WorkSheetReader(Excel.Worksheet workSheet)
        {
            this.workSheet = workSheet;
            this.xlsRange =  workSheet.UsedRange;
            this._usedColumns = getActualColumns();
            this._usedRows = getActualRows();
            this.locations = mapValues();
            this.location_toStr = getHashes(this.locations);
        }

        /// <summary>
        /// Prints out all the data on the workbook.
        /// </summary>
        public void printValues()
        {
            int col = 0, row = 0;
            for(row =1; row <= xlsRange.Rows.Count; row++)
            {
                for(col=1; col <= xlsRange.Columns.Count; col++)
                {
                 
                    try
                    {         
                        var p = (xlsRange.Cells[row, col] as Excel.Range).Value;
                        if (p != null || p != "" || p != " "){
                            Console.WriteLine(p);
                            locations.Add(new ValueLocation(row, col));
                        }
                       
                    }
                    catch (Exception ex) { 
                    }
                    
                }
            }
        }
        /// <summary>
        /// Catalogue all row, col coordinates that hold real data.
        /// </summary>
        /// <returns>List of a value location class that has a row and col value</returns>
        public List<ValueLocation> mapValues()
        {
            List<ValueLocation> valueLocation = new List<ValueLocation>();
            int col = 0, row = 0;
            for (row = 1; row <= xlsRange.Rows.Count; row++)
            {
                for (col = 1; col <= xlsRange.Columns.Count; col++)
                {

                    try
                    {
                        var p = (xlsRange.Cells[row, col] as Excel.Range).Value;
                        String s = Convert.ToString(p);
                        if (!String.IsNullOrWhiteSpace(s))
                        {

                            ValueLocation value_location = new ValueLocation(row, col);
                            valueLocation.Add(value_location);
                            this.location_toStr.Add(value_location.ToString());
                        }

                    }
                    catch (Exception ex)
                    {
                    }
                    
                }
            }
            return valueLocation;
        }

        /// <summary>
        /// Get the valid ValueLocations toString() method recorded
        /// </summary>
        /// <param name="v_list">List of valid locations</param>
        /// <returns>List with the ValueLocations toString() method</returns>
        private List<String> getHashes(List<ValueLocation> v_list)
        {
            List<String> to_str = new List<string>();
            foreach(ValueLocation vl in v_list)
            {
                to_str.Add(vl.ToString());
            }
            return to_str;
        }

        /// <summary>
        /// Maps all valid values, including whitespace or null
        /// </summary>
        /// <returns>All valid col, row combinations</returns>
        public List<ValueLocation> mapValuesAll()
        {
            List<ValueLocation> vl = new List<ValueLocation>();
            int col = 0, row = 0;
            for (row = 1; row <= xlsRange.Rows.Count; row++)
            {
                for (col = 1; col <= xlsRange.Columns.Count; col++)
                {

                    try
                    {
                        var p = (xlsRange.Cells[row, col] as Excel.Range).Value;
                        String s = Convert.ToString(p);
                        ValueLocation value_location = new ValueLocation(row, col);
                        vl.Add(value_location);
                        
                    }
                    catch (Exception ex) { }
                }
            }
            return vl;
        }
        /// <summary>
        /// Loops through the legitimate values of Cells[row_1, x] and saves the values.
        /// </summary>
        /// <returns>Values of all the columns</returns>
        public List<String> getColumnNames()
        {
            List<String> colNames = new List<String>();
            for (int col = 1; col <= xlsRange.Columns.Count; col++)
            {
                try
                {
                    String s = (String) (xlsRange.Cells[1, col] as Excel.Range).Value;
                    colNames.Add(s);
                }catch(Exception ex){
                    Console.Error.WriteLine(ex.StackTrace);    
                }
            }
                return colNames;
        }
        /// <summary>
        /// Gets all the columns that are not null, whitespace or a single space.
        /// </summary>
        /// <returns>Count of the columns with real data.</returns>
        private int getActualColumns()
        {
            int columns = 0;
            for(int col = 1; col <= xlsRange.Columns.Count; col++)
            {
                try
                {
                    String s = (String)(xlsRange.Cells[1, col] as Excel.Range).Value;
                    columns++;
                }catch(Exception ex)
                {

                }
            }
            return columns;

        }
       
        /// <summary>
        /// Get all rows with data that isnt null, whitespace, or a single string.
        /// </summary>
        /// <returns>Count of the rows used in the Excel range</returns>
        private int getActualRows()
        {
            int rows = 0;
            for(int row = 1; row <= xlsRange.Rows.Count; row++)
            {
                try
                {
                    String s = (String)(xlsRange.Cells[row, 1] as Excel.Range).Value;
                    if(!String.IsNullOrWhiteSpace(s))
                        rows++;
                }
                catch (Exception ex) { }
               
            }
            return rows;
        }
        /// <summary>
        /// Make Column count publicly accessible.
        /// </summary>
        /// <returns>Returns columns from getActualColumns</returns>
        public int getRealColumns()
        {
            return this._usedColumns;
        }
        /// <summary>
        /// Make Row count publicly accessible.
        /// </summary>
        /// <returns>Returns rows from getActualRows</returns>
        public int getRealRows()
        {
            return this._usedRows;
        }
        /// <summary>
        /// Get an individual cell's value in the spreadsheet.
        /// </summary>
        /// <param name="row">Integer position of the row. >= 1</param>
        /// <param name="col">Integer position of the column. >= 1</param>
        /// <returns>String value of columns</returns>
        public String getValue(int row, int col)
        {
            
            var val = (xlsRange.Cells[row, col] as Excel.Range).Value;
            String chk = new ValueLocation(row, col).ToString();
            if (this.location_toStr.Contains(chk)) { 
                String val_s = Convert.ToString(val);
                 return val_s;
            }
             return "";
        }
       
           
        
           
        
        /// <summary>
        /// ToString method 
        /// </summary>
        /// <returns>Overview of WorkSheetReader's properties</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(this.workSheet.Name + "\n");
            sb.Append(this._usedColumns + " Columns\n");
            sb.Append(this._usedRows + " Rows \n");
            sb.Append(this._usedRows * this._usedColumns + " Values\n");
            return sb.ToString();
        }
    }
}
