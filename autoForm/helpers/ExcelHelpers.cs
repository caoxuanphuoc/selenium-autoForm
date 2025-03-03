using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace autoForm.helpers
{
    public static class ExcelHelpers
    {
        public static string getAnswerString(IXLRangeRow listValue, int index)
        {
            return listValue.Cell(index).Value.ToString();
        }

        public static int getAnswer(IXLRangeRow listValue, int index)
        {
            string value = listValue.Cell(index).Value.ToString();
            return int.Parse(value);
        }

        public static List<int > getListAnswer(IXLRangeRow listValue, int from, int to)
        {
            List<int> result = new List<int>();
            for(int i= from; i<= to; i++)
            {
                string value = listValue.Cell(i).Value.ToString();
                var valueInt = int.Parse(value);
                result.Add(valueInt);
            }
            return result;
        }
    }
}
