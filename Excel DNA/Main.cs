using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.SignalR.Client;
using Excel_DNA.Core;
using Excel_DNA.Models;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Irony.Parsing;
using XLParser;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using IDnaRibbonControl = ExcelDna.Integration.CustomUI.IRibbonControl;

namespace Excel_DNA
{
    [ComVisible(true)]
    public class CoreFunctions: ExcelRibbon, IExcelAddIn
    {
        static List<FormulaNode> res = new List<FormulaNode>();

        static List<Cell> cells = new List<Cell>();

        static HubConnection connection;

        static ExcelApplicaton exApp = ExApp.GetInstance();

        static MyForm treeForm = new MyForm($"http://194.87.74.186:3000/CreateTreePage/?userName={exApp.UserName}");

        static bool minus = true;

        public void AutoOpen()
        {
            //server.Prefixes.Add("http://127.0.0.1:8888/connection/");
            connection = new HubConnectionBuilder()
            .WithUrl("https://darkcell.ru/chathub")
            .Build();
            connection.On<string, string>("Receive", async (message, username) =>
            {
                // await Task.Delay(2000);
                // await connection.InvokeAsync("Send", username, message);
                // await Task.Delay(2000);
            });
        }

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
        public void settingsButtonPressed(IDnaRibbonControl control)
        {
            MessageBox.Show("Раздел временно неактивен.");
        }

        public void errorFormButtonPressed(IDnaRibbonControl control)
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

        public void helpButtonPressed(IDnaRibbonControl control)
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
        public void aboutButtonPressed(IDnaRibbonControl control)
        {
            MessageBox.Show("Раздел временно неактивен.");
        }
        public void createTreeButtonPressed(IDnaRibbonControl control)
        {
            RangeGet();
        }

        // ShortCut for call tree creator
        [ExcelCommand(ShortCut = "^Q")]
        public static void CallShortCutTree()
        {
            RangeGet();
        }
        // ShortCut for hide tree
        [ExcelCommand(ShortCut = "{ESC}")]
        public static void HideShortCutTree()
        {
            treeForm.Hide();
        }

        public async static void RangeGet()
        {
            res.Clear();

            ExcelApplicaton exApp = ExApp.GetInstance();
            Range range = exApp.ActiveCell;

            //range.Text.Replace("#", "@")
            res.Add(new FormulaNode { Name = range.AddressLocal.Replace("$",""), Result = string.Format("{0:F2}", range.Value), Depth = "0", Type = "function" });
            string lettersFormula = range.FormulaLocal.Replace(" ", ""); // Замените на вашу строку с формулой

            string valuesFormulaPattern = @"^=-*\d+(\.\d+)?$"; //@"^=-(\d+\.\d+|\d+)$";
            string stringValuePattern = @"^=""[^""]*""$";
            string allSymbolsPattern = @"^=[^\d]*[a-z]+[^\d]*$";
            string ONEmorePATTERN = @"^=""[^""]*""$";

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
                    SendMessage();
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
                res.Add(new FormulaNode { Name = range.Text, Result = range.Text, Depth = "1" });
                SendMessage();
                return;
            }
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА =text, ="text"
            else if (Regex.IsMatch(range.Formula, allSymbolsPattern) || Regex.IsMatch(range.Formula, stringValuePattern))
            {
                range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                res.Add(new FormulaNode { Name = range.FormulaLocal.Substring(1), Result = range.Text.Replace("#", "@"), Depth = "1" });
                SendMessage();
                return;
            }

            range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);

            DepthFirstSearch(node, exApp, 1);


            SendMessage();

        }

        public async static void SendMessage()
        {
            var nodesToRemove = new List<FormulaNode>();
            foreach (var temp_node in res)
            {
                temp_node.Childrens.AddRange(res.Where(x => x.Parent == temp_node));
                nodesToRemove.AddRange(res.Where(x => x.Parent == temp_node));
            }
            foreach (var nodeToRemove in nodesToRemove)
            {
                res.Remove(nodeToRemove);
            }
            if(res.Count > 1)
            {
                res[0].Childrens.Add(res[1]);
            }

            var options = new JsonSerializerOptions
            {
                IncludeFields = true,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
                WriteIndented = true
            };
            var json = JsonSerializer.Serialize(res[0], options);
            res.Clear();
            
            try
            {
                await connection.StartAsync();
                await connection.InvokeAsync("Send", exApp.UserName, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error invoking hub method: {ex.Message}");
                // Дополнительная обработка ошибки по вашему усмотрению
            }

            //await connection.InvokeAsync("SendJsonChunk", chunks.First(), true);

            //foreach (var chunk in chunks.Skip(1))
            //{
            //    await connection.InvokeAsync("SendJsonChunk", chunk, false);
            //}

            await connection.StopAsync();

            treeForm.Show();
        }

        public async static void SendSpecialMessage()
        {
            var nodesToRemove = new List<FormulaNode>();
            foreach (var temp_node in res)
            {
                temp_node.Childrens.AddRange(res.Where(x => x.Parent == temp_node));
                nodesToRemove.AddRange(res.Where(x => x.Parent == temp_node));
            }
            foreach (var nodeToRemove in nodesToRemove)
            {
                res.Remove(nodeToRemove);
            }
            if (res.Count > 1) res[0].Childrens.Add(res[1]);
                
            var options = new JsonSerializerOptions
            {
                IncludeFields = true,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
                WriteIndented = true
            };
            var json = JsonSerializer.Serialize(res[0], options);
            res.Clear();

            int chunkSize = 500;

            var chunks = Enumerable.Range(0, json.Length / chunkSize)
                               .Select(i => json.Substring(i * chunkSize, chunkSize));

            await connection.StartAsync();
            try
            {
                await connection.InvokeAsync("Send", exApp.UserName, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error invoking hub method: {ex.Message}");
                // Дополнительная обработка ошибки по вашему усмотрению
            }

            //await connection.InvokeAsync("SendJsonChunk", chunks.First(), true);

            //foreach (var chunk in chunks.Skip(1))
            //{
            //    await connection.InvokeAsync("SendJsonChunk", chunk, false);
            //}

            await connection.StopAsync();

            treeForm.Show();
        }

        public static void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, bool flag = false, FormulaNode parent = null, bool minus = false, bool binary_operation = false)
        {
            //if (parent != null && parent.Childrens == null) parent.Childrens = new List<Node>();
            if (root.Term.Name == "Number_new")
            {
                var name_node = root.Token.ValueString;
                var result_node = RangeSet("=" + name_node);
                res.Add(new FormulaNode { Name = name_node, Result = result_node.Item2, Depth = depth.ToString(), Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
            }
            if (root.Term.Name == "CellToken")
            {
                if (binary_operation)
                {
                    var name_node = root.Print();
                    var result_node = RangeSet("=" + name_node);
                    res.Add(new FormulaNode { Name = name_node, Result = result_node.Item2, Depth = depth.ToString(), Parent = parent });
                    return;
                }
                var name = root.Token.Text;
                CellSet(name, depth, parent);
                return;
            }
            if (root.Term.Name == "NumberToken")
            {
                var name = root.Token.Text;
                res.Add(new FormulaNode { Name = name, Depth = depth.ToString(), Result = name, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.Term.Name == "ReferenceFunctionCall" && root.ChildNodes.Count() == 3)
            {
                var name = root.Print();
                res.Add(new FormulaNode { Name = name, Depth = depth.ToString(), Result = "<диапазон>", Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.IsUnaryOperation())
            {
                var name = root.Print();

                var result = RangeSet("=" + name);
                name = result.Item1;
                //return;
                if (root.ChildNodes[0].Term.Name == "-")
                {
                    if (minus)
                    {
                        return;
                    }
                    else
                    {
                        res.Add(new FormulaNode { Name = name, Depth = depth.ToString(), Result = result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null)});
                        minus = true;
                        return;
                    }
                }
                return;
            } 
            if (root.IsFunction())
            {
                if (root.IsBinaryOperation())
                {
                    if(root.ChildNodes.Count == 3)
                    {
                        var bin_name = root.Print();
                        Tuple<string, string> bin_result = RangeSet("=" + bin_name);
                        bin_name = bin_result.Item1;

                        //BinaryFunZero(root.ChildNodes[0]);

                        var cell_args = new List<ParseTreeNode>();

                        BinaryFunZero(root, ref cell_args);

                        BinaryFunSecond(root, ref cell_args);

                        //if (CheckIsBinary(root.ChildNodes[0]))
                        //{
                        //    if (CheckIsNoBinFun(root.ChildNodes[0]))
                        //    {
                        //        cell_args.Add(root.ChildNodes[0]);
                        //    }
                        //    FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[0]);
                        //    var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                        //    var col = root.GetFunctionArguments();
                        //    var num = analyzer.Numbers();
                        //    cell_args.AddRange(cells);
                        //    foreach(var arg in num)
                        //    {
                        //        cell_args.Add(new ParseTreeNode(new Token(new Terminal("Number_new"), new SourceLocation(), "test", arg)));
                        //    }
                        //    //root.ChildNodes.AddRange(num);
                        //}

                        //if (CheckIsBinary(root.ChildNodes[2]))
                        //{
                        //    FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[2]);
                        //    var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                        //    var col = root.GetFunctionArguments();
                        //    var num = analyzer.Numbers();
                        //    cell_args.AddRange(cells);
                        //    foreach (var arg in num)
                        //    {
                        //        cell_args.Add(new ParseTreeNode(new Token(new Terminal("Number_new"), new SourceLocation(), "test", arg)));
                        //    }
                        //    root.ChildNodes.AddRange(num);
                        //}

                        res.Add(new FormulaNode { Name = bin_name, Depth = depth.ToString(), Result = bin_result.Item2, 
                            Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), 
                            Type = "function" });

                        if(cell_args.Count > 0)
                        {
                            for (int i = 0; i < cell_args.Count; i++) {
                                DepthFirstSearch(cell_args[i], application, depth + 1, false, res.Last());
                            }
                        }
                        else
                        {
                            for (int i = 0; i < root.ChildNodes.Count; i++)
                            {
                                if (i == 1)
                                {
                                    continue;
                                }
                                DepthFirstSearch(root.ChildNodes[i], application, depth + 1, false, res.Last());
                            }
                        }
                        return;
                    }
                    return;
                }
                
                var name = root.Print();

                Tuple<string,string> result = RangeSet("=" + name);
                name = result.Item1;
                
                res.Add(new FormulaNode{ Name = name, Depth = depth.ToString(), Result = result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                var stop = 5;
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1, false, parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) <= depth) : null));
                }
                return;
            }
            if(root.IsRange())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var result = "range";
                res.Add(new FormulaNode { Name = name, Depth = depth.ToString(), Result = result, Parent = parent });
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
                        DepthFirstSearch(root.ChildNodes[0], application, depth, true, parent);
                        return;
                    }
                }
                else
                {
                    res.Add(new FormulaNode { Name = name, Depth = depth.ToString(), Result = result.Item2 , Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                }
                if (root.ChildNodes.Count == 1 && root.ChildNodes[0].IsParentheses()) //проверка внутри только скобки
                {
                    DepthFirstSearch(root.ChildNodes[0], application, depth,true, parent);
                    return;
                }
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1, false, res.Last());
                }
                return;
            }
            

            foreach (var child in root.ChildNodes)
            {
                DepthFirstSearch(child, application, depth, false,parent, minus,binary_operation);
            }
        }

        public static Tuple<string,string> RangeSet(string formula)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            excelApp.Range["BBB1000"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["BBB1000"];
            var res = string.Format("{0:F2}", range.Value);
            //range.Text.Replace("#", "@")
            return Tuple.Create(range.FormulaLocal.Substring(1), res);
            //var test = 5;
        }

        public static void CellSet(string cellName, int cellDepth, FormulaNode parent)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Range range = excelApp.Range[cellName];
            if (range.Value == null)
            {
                res.Add(new FormulaNode { Name = cellName, Depth = cellDepth.ToString(), Result = "<пусто>", Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            else if (range.Value.GetType() == typeof(string))
            {
                res.Add(new FormulaNode { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text, Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            var result = range.Text.Replace("#", "@");
            res.Add(new FormulaNode { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text.Replace("#", "@"), Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
            cells.Add(new Cell { Adress = cellName, Fun = result });
            //string pattern = "^=[A-Z]+\\d*$"; //"^=([0-9A-Z&^:;(),/. *+-]*)?$";
            //Regex regex = new Regex(pattern);
            //if (range.Formula.GetType() == typeof(string) && regex.IsMatch(range.Formula))
            //{
            //    res.Add(new Node { Name = range.Formula.ToString(), Depth = (cellDepth+1).ToString(), Result = range.Text.Replace("#", "@") });
            //}
            return;
        }

        private static void AddBinaryToRes(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, FormulaNode parent = null)
        {
            var bin_name = root.Print();
            Tuple<string, string> bin_result = RangeSet("=" + bin_name);
            bin_name = bin_result.Item1;

            res.Add(new FormulaNode { Name = bin_name, Depth = depth.ToString(), Result = bin_result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
            DepthFirstSearch(root.ChildNodes[0], application, depth + 1, false, parent = res.Last());
            DepthFirstSearch(root.ChildNodes[2], application, depth + 1, false, parent = res.Last());
            return;
        }

        private static bool CheckIsBinary(ParseTreeNode node)
        {
            if (node.IsBinaryOperation()) return true;
            foreach (var child in node.ChildNodes)
            {
                if (CheckIsBinary(child))
                {
                    return true;
                }
            }
            return false;
        }
        private static bool CheckIsNoBinFun(ParseTreeNode node)
        {
            if (node.IsFunction() && !node.IsBinaryOperation()) return true;
            foreach (var child in node.ChildNodes)
            {
                if (CheckIsNoBinFun(child))
                {
                    return true;
                }
            }
            return false;
        }

        private static void BinaryFunZero(ParseTreeNode root, ref List<ParseTreeNode> cell_args)
        {
            //var cell_args = new List<ParseTreeNode>();
            if (CheckIsBinary(root.ChildNodes[0]))
            {
                if (CheckIsNoBinFun(root.ChildNodes[0]))
                {
                    cell_args.Add(root.ChildNodes[0]);
                    return;
                }
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[0]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cell_args.AddRange(cells);
                foreach (var arg in num)
                {
                    cell_args.Add(new ParseTreeNode(new Token(new Terminal("Number_new"), new SourceLocation(), "test", arg)));
                }
                //root.ChildNodes.AddRange(num);
            }
        }
        private static void BinaryFunSecond(ParseTreeNode root, ref List<ParseTreeNode> cell_args)
        {
            //var cell_args = new List<ParseTreeNode>();
            if (CheckIsBinary(root.ChildNodes[2]))
            {
                if (CheckIsNoBinFun(root.ChildNodes[2]))
                {
                    cell_args.Add(root.ChildNodes[2]);
                    return;
                }
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[2]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cell_args.AddRange(cells);
                foreach (var arg in num)
                {
                    cell_args.Add(new ParseTreeNode(new Token(new Terminal("Number_new"), new SourceLocation(), "test", arg)));
                }
                //root.ChildNodes.AddRange(num);
            }
        }

        public void AutoClose()
        {
            
        }
    }

}