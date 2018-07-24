using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections;

namespace NPOIExcelHlper.Core.Helper
{
    public class NPOIExcelHelper
    {
        private HSSFWorkbook wb = null;
        public NPOIExcelHelper(string filePath)
        {
            try
            {
                using (FileStream targetFile = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    wb = new HSSFWorkbook(targetFile);
                }
            }
            catch (Exception er)
            {
                wb = null;
            }
        }
        public List<Dictionary<string, string>> GetAllRowsValuesWithColumnMapping(string sheetName, int columnRowIndex = 0)
        {
            HSSFSheet targetSheet = null;
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                if (wb.GetSheetName(i) == sheetName)
                {
                    targetSheet = (HSSFSheet)wb.GetSheetAt(i);
                    break;
                }
            }
            if (targetSheet == null)
                return null;
            return GetAllRowsValuesWithColumns(targetSheet, columnRowIndex);
        }
        private List<Dictionary<string, string>> GetAllRowsValuesWithColumns(HSSFSheet curSheet, int columnRowIndex = 0)
        {
            List<Dictionary<string, string>> retval = new List<Dictionary<string, string>>();
            if (curSheet == null)
                return null;
            Dictionary<int, string> columnIndexes = GetSheeColumnsIndex(curSheet, columnRowIndex);
            List<List<string>> allRowsValues = GetAllRowsValues(curSheet);
            foreach (List<string> curList in allRowsValues)
            {
                Dictionary<string, string> curRowColumnValues = new Dictionary<string, string>();
                int i = 0;
                foreach (string curentItemValue in curList)
                {
                    i++;
                    try
                    {
                        curRowColumnValues.Add(columnIndexes[i], curentItemValue);
                    }
                    catch (ArgumentException ar)
                    {
                        curRowColumnValues.Add(columnIndexes[i] + i, curentItemValue);
                    }
                    catch (Exception er)
                    {

                    }
                }
                retval.Add(curRowColumnValues);

            }
            return retval;
        }
        private Dictionary<int, string> GetSheeColumnsIndex(HSSFSheet curSheet, int rowIndex = 0)
        {
            Dictionary<int, string> retval = new Dictionary<int, string>();
            if (curSheet == null)
                return null;
            int i = 0;

            List<string> curRowValues = GetSingleRowValue((HSSFRow)curSheet.GetRow(rowIndex));
            foreach (string curitem in curRowValues)
            {
                i++;
                retval = AddColumnsToList(curitem, i, retval);
            }
            return retval;
        }
        private Dictionary<int, string> AddColumnsToList(string columnName, int index, Dictionary<int, string> curList)
        {
            try
            {
                curList.Add(index, columnName);
            }
            catch (Exception er)
            {
                return curList;
            }

            return curList;
        }
        public List<List<string>> GetAllRowsValues(HSSFSheet curSheet)
        {
            List<List<string>> retval = new List<List<string>>();
            if (curSheet == null)
                return null;
            IEnumerator rows = curSheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                HSSFRow curRow = (HSSFRow)rows.Current;
                retval.Add(GetSingleRowValue(curRow));
            }
            return retval;
        }
        public List<string> GetSingleRowValue(HSSFRow curRow)
        {
            List<string> curList = new List<string>();
            if (curRow == null)
                return null;
            IEnumerator columns = curRow.GetEnumerator();
            while (columns.MoveNext())
            {
                HSSFCell curCell = (HSSFCell)columns.Current;
                curList.Add(GetCellValue(curCell));
            }
            return curList;
        }
        public string GetCellValue(HSSFCell curCell)
        {
            string retval = null;
            if (curCell.CellType == CellType.String)
                retval = (curCell.RichStringCellValue).ToString();
            else if (curCell.CellType == CellType.Numeric)
                retval = (curCell.NumericCellValue).ToString();
            else if (curCell.CellType == CellType.Boolean)
                retval = (curCell.BooleanCellValue).ToString();
            else
                retval = "";
            return retval;
        }


    }
}
