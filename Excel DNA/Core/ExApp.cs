using Excel_DNA.Models;
using ExcelDna.Integration;
using System;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;

namespace Excel_DNA.Core
{
    public class ExApp
    {
        private static ExcelApplicaton instance;

        public static ExcelApplicaton GetInstance()
        {
            instance ??= (ExcelApplicaton)ExcelDnaUtil.Application;
            return instance;
        }

        internal static object? ReadCellValue(string? name)
        {
            if (string.IsNullOrEmpty(name)) { return string.Empty; }

            var cell = GetInstance().Range[name];

            Debug.WriteLine($"ReadCellValue for {name}: text {cell.Text}, value {cell.Value2} type {cell.Value2?.GetType()}");

            var cellValue = cell.Value2;

            if (cellValue is int i)
            {
                return i switch
                {
                    -2146826288 => "#NULL!",
                    -2146826281 => "#DIV/0!",
                    -2146826273 => "#VALUE!",
                    -2146826265 => "#REF!",
                    -2146826259 => "#NAME?",
                    -2146826252 => "#NUM!",
                    -2146826246 => "#N/A",
                    _ => i
                };
            }
            return cellValue;
        }

        internal static object? GetValue(string? name)
        {
            var excel = GetInstance();

            //var cellValue = excel.Evaluate(formula); // не работает в случае формулы =INDIRECT(C7&D7)

            excel.Range["BBB1000"].Formula = $"={name}";
            var cell = excel.Range["BBB1000"];

            var evaluateValue = excel.Evaluate($"={name}");

            //GetValue (evaluate) for INDIRECT(C7&D7): value System.__ComObject type System.__ComObject, type2 Range
            //GetValue (set cell) for INDIRECT(C7&D7): text 0, value 0 type System.Double

            //GetValue (evaluate) for OFFSET(G16,C16,D16): value System.__ComObject type System.__ComObject, type2 Range
            //GetValue(set cell) for OFFSET(G16, C16, D16): text 0, value 0 type System.Double

            //GetValue (evaluate) for OFFSET(C16,C17,D17,E17,F17): value System.__ComObject type System.__ComObject, type2 Range
            //GetValue(set cell) for OFFSET(C16, C17, D17, E17, F17): text #VALUE!, value -2146826273 type System.Int32

            //GetValue (evaluate) for IF(C21:D21>0,E21,F21): value System.Object[*] type System.Object[*], type2 <not com obj>
            //Value is array of type System.Object[*]
            //String <
            //String >
            //GetValue (set cell) for IF(C18>E18,E18,F18): text 5, value 5 type System.Double
            //GetValue(set cell) for IF(C21: D21 > 0, E21, F21): text #VALUE!, value -2146826273 type System.Int32

            //GetValue(evaluate) for C21:D21 > 0: value System.Object[*] type System.Object[*], type2 < not com obj >
            //Value is array of type System.Object[*]
            //Boolean False
            //Boolean True
            //GetValue(set cell) for C21:D21 > 0: text #VALUE!, value -2146826273 type System.Int32

            //GetValue (evaluate) for стр20*2: value System.Object[*] type System.Object[*], type2 <not com obj>
            //Double 0
            //Int32 - 2146826273
            //Double 0
            //Double 0
            //Double 0
            //Double 0
            //Double 26
            //Int32 - 2146826273
            //Double 0 and etc, 16k items
            //GetValue(set cell) for стр20 * 2: text 0, value 0 type System.Double

            Debug.WriteLine($"GetValue (evaluate) for {name}: value {evaluateValue} type {evaluateValue?.GetType()}, type2 {ComHelper.GetTypeName(evaluateValue)}");
            if (evaluateValue is Array)
            {
                var array = (Array)(object)evaluateValue;// https://stackoverflow.com/a/41227268 https://learn.microsoft.com/en-us/archive/blogs/mshneer/oh-that-mysteriously-broken-visiblesliceritemslist
                for (int i = 0; i < Math.Min(array.Length, 10); i++)
                {
                    object? item = array.GetValue(i+1);
                    string strValue = item?.ToString() ?? "<null>";
                    Debug.WriteLine($"{item?.GetType().Name ?? "<null>"} {strValue}");
                }
            }
            Debug.WriteLine($"GetValue (set cell) for {name}: text {cell.Text}, value {cell.Value2} type {cell.Value2?.GetType()}");

            return cell.Value2;
        }
    }
}
