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
            res.Clear();
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application; 
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;
            //var res2 = excelApp.Evaluate("A1");

            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);
            //FormulaAnalyzer analyzer = new FormulaAnalyzer(range.Formula);
            DepthFirstSearch(node, excelApp, 0);
            var leters = JsonSerializer.Serialize(res[0].Name);
            var json = JsonSerializer.Serialize(res);
            var url = "http://localhost:3000/?dialogID=15&lettersFormula=" + leters + "&valuesFormula = " + leters + "&jsonString=" + json;
            MyForm form = new MyForm(url);
            form.Show();

        }
        public static void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth)
        {
            if(root.Term.Name == "CellToken")
            {
                var name = root.Token.Text;
                var result = RangeSet("=" + name);
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result});
            }
            if (root.IsFunction())
            {
                root.Print();
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                //var depth = analyzer.Depth().ToString();
                //var wb = application.Workbooks.Add(Type.Missing);
                //var ws = wb.Worksheets[1];
                //application.Windows[1].Visible= false;
                var result = RangeSet("="+name);
                res.Add(new Node{ Name = name, Depth = depth.ToString(), Result = result});
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
               // var depth = analyzer.Depth().ToString();
                var result = "range";
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result });
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1);
                }
                return;
            }
            if (root.IsParentheses())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
              //  var depth = analyzer.Depth().ToString();
                var result = RangeSet("=" + name);
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result });
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

        public static string RangeSet(string formula)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            excelApp.Range["K13"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["K13"];
            return range.Value.ToString();
            var test = 5;
        }

    }

}