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
using ExcelDna.Integration;

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
                  <tab id='tab1' label='Darkcell'>
                    <group id='treeGroup' label='Formula tree'>
                      <button id='createTreeButton' label='Create tree' onAction='createTreeButtonPressed'/>
                    </group >
                    <group id='moreGroup' label='More'>
                      <button id='settingsButton' label='Settings' onAction='settingsButtonPressed'/>
                      <button id='errorFormButton' label='Send error form' onAction='errorFormButtonPressed'/>
                      <button id='helpButton' label='Help' onAction='helpButtonPressed'/>
                      <button id='aboutButton' label='About us' onAction='aboutButtonPressed'/>
                    </group >
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }
        public void settingsButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Раздел временно неактивен.");
        }
        public void errorFormButtonPressed(IRibbonControl control)
        {
            var url = "http://localhost:3000/?windowType=errorFormPage";
            MyForm errorForm = new MyForm(url);
            errorForm.Show();
        }

        // ShortCut for call tree creator
        [ExcelCommand(ShortCut = "^H")]
        public static void CallShortCutHelp()
        {
            MessageBox.Show("Руководство по надстройке Darkcell:\n\n" +
                "Раздел 'Formula tree'\n" +
                "1) Create tree - представляет формулу, лежащую в выбранной ячейке в виде таблицы. (Ctrl+Shift+Q)\n\n" +
                "Раздел 'More'\n" +
                "1) Settings - открывает панель настроек.\n" +
                "2) Send error form - открывает страницу с формой, для сообщения об обнаруженных ошибках.\n" +
                "3) Help - открывает окно с кратким описанием интерфейса надстройки и её функционала. (Ctrl+Shift+H)\n" +
                "4) About us - открывает страницу с подробной информацией о нашем расширении.");
        }

        public void helpButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Руководство по надстройке Darkcell:\n\n" +
                "Раздел 'Formula tree'\n" +
                "1) Create tree - представляет формулу, лежащую в выбранной ячейке в виде таблицы. (Ctrl+Shift+Q)\n\n" +
                "Раздел 'More'\n" +
                "1) Settings - открывает панель настроек.\n" +
                "2) Send error form - открывает страницу с формой, для сообщения об обнаруженных ошибках.\n" +
                "3) Help - открывает окно с кратким описанием интерфейса надстройки и её функционала. (Ctrl+Shift+H)\n" +
                "4) About us - открывает страницу с подробной информацией о нашем расширении.");
        }
        public void aboutButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Раздел временно неактивен.");
        }
        public void createTreeButtonPressed(IRibbonControl control)
        {
            //var url = "https://test-excel.vercel.app/?dialogID=15&lettersFormula=" + res
            //MyForm form = new MyForm();
            //form.Show();
            RangeGet();
        }

        // ShortCut for call tree creator
        [ExcelCommand(ShortCut = "^Q")]
        public static void CallShortCutTree()
        {
            RangeGet();
        }

        public static void RangeGet()
        {
            //application1.Visible = true;
            res.Clear();
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;
            res.Add(new Node { Name = range.AddressLocal.Replace("$",""), Result = range.Text.Replace("#", "@"), Depth = "0" });
            string lettersFormula = range.FormulaLocal; // Замените на вашу строку с формулой

            //var valueTest = range.Value;
            //var stop5 = 5;
            string valuesFormulaPattern = @"^=-*\d+(\.\d+)?$"; //@"^=-(\d+\.\d+|\d+)$";
            string stringValuePattern = @"^=""[^""]*""$";
            string allSymbolsPattern = @"^=[^\d]*[a-z]+[^\d]*$";

            // пустая ячейка
            if (range.Formula == "" && range.Value == null && range.Text == "")
            {
                MessageBox.Show("Ячейка не может быть пустой.");
                return;
            }
            else if (range.Formula[0] != '=')
            {
                // ячейка с числом
                if (range.Value.GetType() == typeof(int) || range.Value.GetType() == typeof(float) || range.Value.GetType() == typeof(double))
                {
                    //res[0].Result = range.Text;
                    var earlyJson = JsonSerializer.Serialize(res);
                    var earlyUrl = "http://localhost:3000/?windowType=treePage&jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                    MyForm earlyTreeForm = new MyForm(earlyUrl);
                    earlyTreeForm.Show();
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
                //res[0].Result = range.Text;
                res.Add(new Node { Name = range.Text, Result = range.Text, Depth = "1" });
                var earlyJson = JsonSerializer.Serialize(res);
                var earlyUrl = "http://localhost:3000/?windowType=treePage&jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                MyForm earlyTreeForm = new MyForm(earlyUrl);
                earlyTreeForm.Show();
                return;
            }
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА =text, ="text"  (ну или обработка ошибки #ИМЯ?)
            else if (Regex.IsMatch(range.Formula, allSymbolsPattern) || Regex.IsMatch(range.Formula, stringValuePattern))
            {
                res.Add(new Node { Name = range.FormulaLocal.Substring(1), Result = range.Text.Replace("#", "@"), Depth = "1" });
                var earlyJson = JsonSerializer.Serialize(res);
                var earlyUrl = "http://localhost:3000/?windowType=treePage&jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                MyForm earlyTreeForm = new MyForm(earlyUrl);
                earlyTreeForm.Show();
                return;
            }


            //string pattern = @"([A-Z]\d+)\s*([<>]=?|!=)\s*([A-Z]\d+)\s*&\s*([A-Z]\d+)\s*([<>]=?|!=)\s*([A-Z]\d+)";
            //Regex regex = new Regex(pattern);

            //string transformedString = regex.Replace(lettersFormula, match =>
            //{
            //    return $"AND({match.Groups[1].Value}{match.Groups[2].Value}{match.Groups[3].Value}, {match.Groups[4].Value}{match.Groups[5].Value}{match.Groups[6].Value})";
            //});


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);
            DepthFirstSearch(node, excelApp, 1);
            //res[0].Result = res[1].Result;
            var json = JsonSerializer.Serialize(res);
            var url = "http://localhost:3000/?windowType=treePage&jsonString=" + json + "&lettersFormula" + lettersFormula;
            MyForm treeForm = new MyForm(url);
            treeForm.Show();

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

                //application.Range["BBB1000"].Formula = "=" + name;
                //var range = application.Range["BBB1000"];
                //var rangeValue = range.Value;
                //var rangeFormula = range.Formula;
                //var rangeText = range.Text;
                //var test = 5;

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
            excelApp.Range["BBB1000"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["BBB1000"];
            return Tuple.Create(range.FormulaLocal.Substring(1), range.Text.Replace("#", "@"));
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
                res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text });
                return;
            }
            res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text.Replace("#", "@") });
            //string pattern = "^=[A-Z]+\\d*$"; //"^=([0-9A-Z&^:;(),/. *+-]*)?$";
            //Regex regex = new Regex(pattern);
            //if (range.Formula.GetType() == typeof(string) && regex.IsMatch(range.Formula))
            //{
            //    res.Add(new Node { Name = range.Formula.ToString(), Depth = (cellDepth+1).ToString(), Result = range.Text.Replace("#", "@") });
            //}
            return;
        }

    }

}