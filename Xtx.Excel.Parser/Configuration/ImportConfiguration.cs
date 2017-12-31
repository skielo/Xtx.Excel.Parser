using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xtx.Excel.Parser.Configuration
{
    public abstract class ImportConfiguration
    {
        /// <summary>
        /// The names of the worksheets to import.
        /// If no value is provided then all sheets are processed.
        /// </summary>
        public IList<string> WorksheetNames { get; set; }

        public bool FirstRowHasHeaders { get; set; }

        protected ImportConfiguration()
            : this(true)
        {
        }

        protected ImportConfiguration(bool firstRowHasHeaders)
        {
            FirstRowHasHeaders = firstRowHasHeaders;
        }
    }
}
