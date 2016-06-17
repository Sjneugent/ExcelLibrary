using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class ValueLocation
    {
        public int row { get; set;}
        public int col {get; set;}
 
        public ValueLocation(int row, int col)
        {
            this.row = row;
            this.col = col;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Convert.ToString(row));
            sb.AppendLine();
            sb.Append(Convert.ToString(col));
            return sb.ToString();
        }
    }
}
