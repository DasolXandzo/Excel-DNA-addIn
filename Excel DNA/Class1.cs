using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;
using XLParser;
using Irony.Parsing;
using System.Security.Policy;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace Excel_DNA
{

    public class Node
    {
        public string? Name { get; set; }
        public string? Depth { get; set; }
        public string? Result { get; set; }
    }
    public static class MyFunctions
    {
        static List<Node> res = new List<Node>();

        [ExcelCommand(MenuName = "Test", MenuText = "Range Set")]
        public static void RangeGet()
        {
            res.Clear();
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application; 
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;
            //var res2 = excelApp.Evaluate("A1");

            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);
            FormulaAnalyzer analyzer = new FormulaAnalyzer(range.Formula);
            DepthFirstSearch(node, excelApp);
            var leters = JsonSerializer.Serialize(res[0].Name);
            var json = JsonSerializer.Serialize(res);
            var url = "http://localhost:3000/?dialogID=15&lettersFormula=" + leters + "&valuesFormula = " + leters + "&jsonString=" + json;
            MyForm form = new MyForm(url);
            form.Show();

        }
        public static void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application)
        {
            if (root == null)
                return;

            if (root.IsFunction())
            {
                root.Print();
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var depth = analyzer.Depth().ToString();
                //var wb = application.Workbooks.Add(Type.Missing);
                //var ws = wb.Worksheets[1];
                //application.Windows[1].Visible= false;
                var result = RangeSet("="+name);
                res.Add(new Node{ Name = name, Depth = depth, Result = result});
                var stop = 5;
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application);
                }
                return;
            }
            if(root.IsRange())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var depth = analyzer.Depth().ToString();
                var result = "range";
                res.Add(new Node { Name = name, Depth = depth, Result = result });
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application);
                }
                return;
            }
            if (root.IsParentheses())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var depth = analyzer.Depth().ToString();
                var result = RangeSet("=" + name);
                res.Add(new Node { Name = name, Depth = depth, Result = result });
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application);
                }
                return;
            }
            

            foreach (var child in root.ChildNodes)
            {
                DepthFirstSearch(child, application);
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

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
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
        }
    }
}