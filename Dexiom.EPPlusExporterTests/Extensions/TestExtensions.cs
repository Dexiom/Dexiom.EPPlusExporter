using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Dexiom.EPPlusExporterTests.Extensions
{
    internal static class TestExtensions
    {
        public static ExcelRange FlagInfo(this ExcelRange source)
        {
            source.Style.Fill.PatternType = ExcelFillStyle.Solid;
            source.Style.Fill.BackgroundColor.SetColor(Color.CornflowerBlue);

            return source;
        }

        public static ExcelRange FlagSuccess(this ExcelRange source)
        {
            source.Style.Fill.PatternType = ExcelFillStyle.Solid;
            source.Style.Fill.BackgroundColor.SetColor(Color.Green);

            return source;
        }

        public static ExcelRange FlagWarning(this ExcelRange source)
        {
            source.Style.Fill.PatternType = ExcelFillStyle.Solid;
            source.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            return source;
        }

        public static ExcelRange FlagCritical(this ExcelRange source)
        {
            source.Style.Fill.PatternType = ExcelFillStyle.Solid;
            source.Style.Fill.BackgroundColor.SetColor(Color.Orange);

            return source;
        }
    }
}
