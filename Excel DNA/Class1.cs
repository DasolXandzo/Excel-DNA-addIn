using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;
using XLParser;
using Irony.Parsing;
using System.Security.Policy;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace Excel_DNA
{

    public class Node
    {
        public string? Name { get; set; }
        public string? Depth { get; set; }
        public string? Result { get; set; }
        public List<ParseTreeNode>? ChildNodes { get; set; }
    }
    public class NodeDepth
    {
        public ParseTreeNode? Node { get; set; }
        public int? Depth { set; get; }
    }



    [ComVisible(true)]
    public class MyFunctions: ExcelRibbon
    {
        static List<Node> res = new List<Node>();

        public static Microsoft.Office.Interop.Excel.Application application1 = new Microsoft.Office.Interop.Excel.Application();
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='Дерево'>
            <group id='group1' label='Надстройка'>
              <button id='button1' label='Создать дерево' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            //var url = "https://test-excel.vercel.app/?dialogID=15&lettersFormula=" + res
            //MyForm form = new MyForm();
            //form.Show();
            RangeGet();
        }
        public static void RangeGet()
        {
            //application1.Visible = true;
            res.Clear();
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;
            res.Add(new Node { Name = range.AddressLocal.Replace("$",""), Depth = "0" });
            string lettersFormula = range.Formula; // Замените на вашу строку с формулой

            //var valueTest = range.Formula[0];
            //var stop5 = 5;
            string valuesFormulaPattern = @"^=-*\d+(\.\d+)?$"; //@"^=-(\d+\.\d+|\d+)$";

            // пустая ячейка
            if (range.Formula == "" && range.Value == null && range.Text == "")
            {
                MessageBox.Show("Ячейка не может быть пустой или содержать текст.");
                return;
            }
            else if (range.Formula[0] != '=')
            {
                // ячейка с числом
                if (range.Value.GetType() == typeof(int) || range.Value.GetType() == typeof(float) || range.Value.GetType() == typeof(double))
                {
                    res[0].Result = range.Text;
                    var earlyJson = JsonSerializer.Serialize(res);
                    var earlyUrl = "http://localhost:3000/?dialogID=15&lettersFormula=" + lettersFormula + "&valuesFormula = " + lettersFormula + "&jsonString=" + earlyJson;
                    MyForm earlyForm = new MyForm(earlyUrl);
                    earlyForm.Show();
                    return;
                }
                // ячейка с текстом, без "=" в начале
                else if (range.Value.GetType() == typeof(string))
                {
                    MessageBox.Show("Ячейка не может быть пустой или содержать текст.");
                    return;
                }
            }
            // ячейка с формулой формата "=число"
            else if (Regex.IsMatch(range.Formula, valuesFormulaPattern))
            {
                res[0].Result = range.Text;
                res.Add(new Node { Name = range.Text, Result = range.Text, Depth = "1" });
                var earlyJson = JsonSerializer.Serialize(res);
                var earlyUrl = "http://localhost:3000/?dialogID=15&lettersFormula=" + lettersFormula + "&valuesFormula = " + lettersFormula + "&jsonString=" + earlyJson;
                MyForm earlyForm = new MyForm(earlyUrl);
                earlyForm.Show();
                return;
            }
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА "=текст"


            //string pattern = @"([A-Z]\d+)\s*([<>]=?|!=)\s*([A-Z]\d+)\s*&\s*([A-Z]\d+)\s*([<>]=?|!=)\s*([A-Z]\d+)";
            //Regex regex = new Regex(pattern);

            //string transformedString = regex.Replace(lettersFormula, match =>
            //{
            //    return $"AND({match.Groups[1].Value}{match.Groups[2].Value}{match.Groups[3].Value}, {match.Groups[4].Value}{match.Groups[5].Value}{match.Groups[6].Value})";
            //});


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);
            DepthFirstSearch(node, excelApp, 1);
            res[0].Result = res[1].Result;
            var json = JsonSerializer.Serialize(res);
            var url = "http://localhost:3000/?dialogID=15&lettersFormula=" + lettersFormula + "&valuesFormula = " + lettersFormula + "&jsonString=" + json;
            MyForm form = new MyForm(url);
            form.Show();

        }
        public static void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth)
        {
            if(root.Term.Name == "CellToken")
            {
                var name = root.Token.Text;
                CellSet(name, depth);
                //res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result});
                return;
            }
            if(root.Term.Name == "ReferenceFunctionCall" && root.ChildNodes.Count() == 3)
            {
                var name = root.Print();
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = "<диапазон>" });
                return;
            }
            if (root.IsFunction())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                application.Range["BBB1000"].Formula = name;
                var range = application.Range["BBB1000"];
                Tuple<string,string> result = RangeSet("=" + name);
                name = result.Item1;
                res.Add(new Node{ Name = name, Depth = depth.ToString(), Result = result.Item2 });
                var stop = 5;
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1);
                }
                return;
            }
            if(root.IsRange())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var result = "range";
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result });
                //foreach (var child in root.ChildNodes)
                //{
                //    DepthFirstSearch(child, application, depth + 1);
                //}
                return;
            }
            if (root.IsParentheses())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
              //  var depth = analyzer.Depth().ToString();
                var result = RangeSet("=" + name);
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result.Item2 });
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1);
                }
                return;
            }
            

            foreach (var child in root.ChildNodes)
            {
                DepthFirstSearch(child, application, depth);
            }
        }

        public static Tuple<string,string> RangeSet(string formula)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            excelApp.Range["K13"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["K13"];
            return Tuple.Create(range.FormulaLocal.Substring(1), range.Text);
            //var test = 5;
        }

        public static void CellSet(string cellName, int cellDepth)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range[cellName];
            if (range.Value == null)
            {
                res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = "<пусто>" });
                return;
            }
            else if (range.Value.GetType() == typeof(string))
            {
                res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = "<текст>" });
                return;
            }
            res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text });
            string pattern = "^=[A-Z]+\\d*$"; //"^=([0-9A-Z&^:;(),/. *+-]*)?$";
            Regex regex = new Regex(pattern);
            if (range.Formula.GetType() == typeof(string) && regex.IsMatch(range.Formula))
            {
                res.Add(new Node { Name = range.Formula.ToString(), Depth = (cellDepth+1).ToString(), Result = range.Text });
            }
            return;
        }

    }

}