using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Dexiom.EPPlusExporter
{
    public class EnumerableExporter 
        : EnumerableExporter<object>
    {
        #region Constructors
        public EnumerableExporter(IEnumerable<object> data) 
            : base(data)
        {
        }
        #endregion
    }

    public class EnumerableExporter<T>
        where T : class
    {
        #region Constructors
        public EnumerableExporter(IEnumerable<T> data)
        {
            Data = data;
        }
        #endregion

        #region Public Functions
        public ExcelPackage CreateExcelPackage()
        {
            var retVal = new ExcelPackage();
            WorksheetHelper.AddWorksheet(retVal, Data, WorksheetName, TableStyle);

            return retVal;
        }

        public ExcelWorksheet AppendToExistingPackage(ExcelPackage package)
        {
            return WorksheetHelper.AddWorksheet(package, Data, WorksheetName, TableStyle);
        }
        #endregion

        #region Properties

        public string WorksheetName { get; set; } = "Data";

        public TableStyles TableStyle { get; set; } = TableStyles.Medium4;

        public IEnumerable<T> Data { get; set; }

        #endregion
    }
}
