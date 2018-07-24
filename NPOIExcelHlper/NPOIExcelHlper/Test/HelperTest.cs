using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using NPOIExcelHlper.Core.Helper;
namespace NPOIExcelHlper.Test
{
    public class HelperTest
    {
        [Fact]
        public void DoTest()
        {
            NPOIExcelHelper helper = new NPOIExcelHelper(Environment.CurrentDirectory + "\\TestResource\\SampleOne.xls");
            List<Dictionary<string, string>> itms = helper.GetAllRowsValuesWithColumnMapping("data");
            //helper.Dohelow();
        }
    }
}
