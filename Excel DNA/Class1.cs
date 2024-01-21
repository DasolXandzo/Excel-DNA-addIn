using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using XLParser;
using Irony.Parsing;
using System.Text.Json;
using System.Text.RegularExpressions;
using IRibbonControl = Microsoft.Office.Core.IRibbonControl;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Windows;
using Microsoft.AspNetCore.Components;

namespace Excel_DNA
{

    public class Node
    {
        public string? Name { get; set; }
        public string? Result { get; set; }
        public string? Depth { get; set; }
    }


    public class Cell
    {
        public string Adress { get; set; }
        public string Fun { get; set; }
    }


    [ComVisible(true)]
    public class MyFunctions: ExcelRibbon
    {
        static List<Node> res = new List<Node>();

        static HttpListener server = new HttpListener();

        static List<Cell> cells = new List<Cell>();

        static HubConnection connection;



        public static Microsoft.Office.Interop.Excel.Application application1 = new Microsoft.Office.Interop.Excel.Application();
        public override string GetCustomUI(string RibbonID)
        {
            server.Prefixes.Add("http://127.0.0.1:8888/connection/");
            connection = new HubConnectionBuilder()
            .WithUrl("https://localhost:7108/chat")
            .Build();
            connection.On<string, string>("Receive", (user, message) =>
            {
                
            });
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
        public void settingsButtonPressed(ExcelDna.Integration.CustomUI.IRibbonControl control)
        {
            MessageBox.Show("Раздел временно неактивен.");
        }

        public void errorFormButtonPressed(IRibbonControl control)
        {
            var url = "http://localhost:3000/ErrorFormPage";
            MyForm errorForm = new MyForm(url);
            errorForm.Show();
        }

        // ShortCut for call help window
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
            RangeGet();
        }

        // ShortCut for call tree creator
        [ExcelCommand(ShortCut = "^Q")]
        public static void CallShortCutTree()
        {
            RangeGet();
        }

        public async static void RangeGet()
        {
            //application1.Visible = true;
            //
            //await hub.Send("asd");
            res.Clear();

            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;
            res.Add(new Node { Name = range.AddressLocal.Replace("$",""), Result = range.Text.Replace("#", "@"), Depth = "0" });
            string lettersFormula = range.FormulaLocal.Replace(" ", ""); // Замените на вашу строку с формулой

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
                    range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                    //res[0].Result = range.Text;
                    var earlyJson = JsonSerializer.Serialize(res);
                    var earlyUrl = "http://localhost:3000/CreateTreePage/?jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                    MyForm earlyTreeForm = new MyForm(earlyUrl);
                    //TestForm testform = new TestForm();
                    //testform.Show();
                    earlyTreeForm.Show();
                    return;
                }
                // ячейка с текстом, без "=" в начале
                else if (range.Value.GetType() == typeof(string))
                {
                    MessageBox.Show("Ячейка не может содержать текст.");
                    return;
                }
            }
            // ячейка с формулой формата "=число"
            else if (Regex.IsMatch(range.Formula, valuesFormulaPattern))
            {
                range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                //res[0].Result = range.Text;
                res.Add(new Node { Name = range.Text, Result = range.Text, Depth = "1" });
                var earlyJson = JsonSerializer.Serialize(res);
                var earlyUrl = "http://localhost:3000/CreateTreePage/?jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                MyForm earlyTreeForm = new MyForm(earlyUrl);
                earlyTreeForm.Show();
                return;
            }
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА =text, ="text"  (ну или обработка ошибки #ИМЯ?)
            else if (Regex.IsMatch(range.Formula, allSymbolsPattern) || Regex.IsMatch(range.Formula, stringValuePattern))
            {
                range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                res.Add(new Node { Name = range.FormulaLocal.Substring(1), Result = range.Text.Replace("#", "@"), Depth = "1" });
                var earlyJson = JsonSerializer.Serialize(res);
                var earlyUrl = "http://localhost:3000/CreateTreePage/?jsonString=" + earlyJson + "&lettersFormula" + lettersFormula;
                MyForm earlyTreeForm = new MyForm(earlyUrl);
                earlyTreeForm.Show();
                return;
            }

            range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);

            DepthFirstSearch(node, excelApp, 1);
            
            var json = JsonSerializer.Serialize(res);
            var url = "http://localhost:3000/CreateTreePage/?jsonString=" + json.Substring(1,100) + "&lettersFormula" + lettersFormula;
            MyForm treeForm = new MyForm(url);
            treeForm.Show();
            await connection.StartAsync();
            await connection.InvokeAsync("Send", "user", "mess");

            //var context = await server.GetContextAsync();

            //var response = context.Response;


            //byte[] buffer = Encoding.UTF8.GetBytes(json);
            //response.ContentType = "application/json";
            //response.ContentLength64 = buffer.Length;
            //using (Stream output = response.OutputStream)
            //{
            //    // отправляем данные
            //    await output.WriteAsync(buffer);
            //    await output.FlushAsync();
            //}
            //context.Response.Close();
            server.Stop();
          //  await hub.Send("asd");

        }
        public static void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, bool flag = false)
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
                //var depth = analyzer.Depth().ToString();
                var result = RangeSet("=" + name);
                if (flag)
                {
                    if (root.ChildNodes.Count == 1 && root.ChildNodes[0].IsParentheses()) //проверка внутри только скобки
                    {
                        DepthFirstSearch(root.ChildNodes[0], application, depth, true);
                        return;
                    }
                }
                else
                {
                    res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result.Item2 });
                }
                if (root.ChildNodes.Count == 1 && root.ChildNodes[0].IsParentheses()) //проверка внутри только скобки
                {
                    DepthFirstSearch(root.ChildNodes[0], application, depth + 1,true);
                    return;
                }
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
            var result = range.Text.Replace("#", "@");
            res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text.Replace("#", "@") });
            cells.Add(new Cell { Adress = cellName, Fun = result });
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