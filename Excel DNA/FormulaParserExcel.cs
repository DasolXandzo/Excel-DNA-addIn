using Excel_DNA.Models;
using ExcelDna.Integration;
using Irony.Parsing;
using System.Text.Json.Serialization;
using System.Text.Json;
using XLParser;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;
using System.Diagnostics;

namespace Excel_DNA
{
    public class FormulaParserExcel
    {
        public List<FormulaNode> res = new List<FormulaNode>();
        public List<Cell> cells = new List<Cell>();
        static ExcelApplicaton exApp;
        public FormulaParserExcel(ExcelApplicaton App)
        {
            exApp = App; 
        }
        public List<FormulaNode> GetRes()
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
            if (res.Count > 1)
            {
                res[0].Childrens.Add(res[1]);
            }
            return res;
        }
        public List<Cell> GetCells() { return cells; }
        public void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, bool flag = false, FormulaNode parent = null, bool minus = false, bool binary_operation = false)
        {
            switch (root.Term.Name)
            {
                case "Number_new": break;
                case "CellToken": break;
                case "NumberToken": break;

            }
            if (root.Term.Name == "Number_new")
            {
                var name_node = root.Token.ValueString;
                var result_node = RangeSet("=" + name_node);
                res.Add(new FormulaNode { Name = name_node, Result = result_node, Depth = depth, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
            }
            if (root.Term.Name == "CellToken")
            {
                if (binary_operation)
                {
                    var name_node = root.Print();
                    var result_node = RangeSet("=" + name_node);
                    res.Add(new FormulaNode { Name = name_node, Result = result_node, Depth = depth, Parent = parent });
                    return;
                }
                var name = root.Token.Text;
                CellSet(name, depth, parent);
                return;
            }
            if (root.Term.Name == "NumberToken")
            {
                var name = root.Token.Text;
                res.Add(new FormulaNode { Name = name, Depth = depth, Result = root.Token.Value, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.Term.Name == "ReferenceFunctionCall" && root.ChildNodes.Count() == 3)
            {
                var name = root.Print();
                res.Add(new FormulaNode { Name = name, Depth = depth, Result = "<диапазон>", Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.IsUnaryOperation())
            {
                var name = root.Print();

                var result = RangeSet("=" + name);
                //name = result.Item1; //todo зачем перезаписывать? Такая же перезапись есть в нескольких местах ниже
                //return;
                if (root.ChildNodes[0].Term.Name == "-")
                {
                    if (minus)
                    {
                        return;
                    }
                    else
                    {
                        res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
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
                    if (root.ChildNodes.Count == 3)
                    {
                        var bin_name = root.Print();
                        var bin_result = RangeSet("=" + bin_name);
                        //bin_name = bin_result.Item1;

                        var cell_args = new List<ParseTreeNode>();

                        res.Add(new FormulaNode
                        {
                            Name = bin_name,
                            Depth = depth,
                            Result = bin_result,
                            Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null),
                            Type = "function"
                        });

                        BinaryFunZero(root, ref cell_args, depth);

                        BinaryFunSecond(root, ref cell_args, depth);

                        if (cell_args.Count > 0)
                        {
                            for (int i = 0; i < cell_args.Count; i++)
                            {
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

                var result = RangeSet("=" + name);
                //name = result.Item1;

                res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                var stop = 5;
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1, false, parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) <= depth) : null));
                }
                return;
            }
            if (root.IsRange())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var result = "range";
                res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = parent });
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
                    res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                }
                if (root.ChildNodes.Count == 1 && root.ChildNodes[0].IsParentheses()) //проверка внутри только скобки
                {
                    DepthFirstSearch(root.ChildNodes[0], application, depth, true, parent);
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
                DepthFirstSearch(child, application, depth, false, parent, minus, binary_operation);
            }
        }

        public static object RangeSet(string formula)
        {
            exApp.Range["BBB1000"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = exApp.Range["BBB1000"];
            var value = exApp.Evaluate(formula);
            if (range.FormulaLocal.Substring(1).ToString() != formula.Substring(1))
            {
                Debug.WriteLine($"{range.FormulaLocal.Substring(1)} != {formula.Substring(1)}");
            }
            return value;
        }

        public void CellSet(string cellName, int cellDepth, FormulaNode parent)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range[cellName];
            if (range.Value == null)
            {
                res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, Result = "<пусто>", Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            if (range.Value is string)
            {
                res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, 
                    Result = range.Text, //todo подумать над тем, чтобы сохранять и range.Text и range.Value
                    Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            var result = range.Text.Replace("#", "@");
            //var result = range.Text;
            if (result.ToString() != range.Text.ToString())
            {
                Debug.WriteLine($"{result.ToString()} != {range.Text.ToString()}");
            }
            res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, 
                Result = range.Value, //todo тут было range.Text.Replace("#", "@"), найти случаи, когда это было нужно
                Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
            cells.Add(new Cell { Adress = cellName, Fun = result });
            return;
        }
        //private static void AddBinaryToRes(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, FormulaNode parent = null)
        //{
        //    var bin_name = root.Print();
        //    Tuple<string, string> bin_result = RangeSet("=" + bin_name);
        //    bin_name = bin_result.Item1;

        //    res.Add(new FormulaNode { Name = bin_name, Depth = depth.ToString(), Result = bin_result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
        //    DepthFirstSearch(root.ChildNodes[0], application, depth + 1, false, parent = res.Last());
        //    DepthFirstSearch(root.ChildNodes[2], application, depth + 1, false, parent = res.Last());
        //    return;
        //}
        private static bool CheckIsBinary(ParseTreeNode FormulaNode)
        {
            if (FormulaNode.IsBinaryOperation()) return true;
            foreach (var child in FormulaNode.ChildNodes)
            {
                if (CheckIsBinary(child))
                {
                    return true;
                }
            }
            return false;
        }
        private static bool CheckIsNoBinFun(ParseTreeNode FormulaNode)
        {
            if (FormulaNode.IsFunction() && !FormulaNode.IsBinaryOperation()) return true;
            foreach (var child in FormulaNode.ChildNodes)
            {
                if (CheckIsNoBinFun(child))
                {
                    return true;
                }
            }
            return false;
        }
        private static bool CheckIsSameSign(ParseTreeNode root, ParseTreeNode child)
        {
            foreach (var FormulaNode in child.ChildNodes)
            {
                if (FormulaNode.Term.Name == "FunctionCall" && FormulaNode.IsBinaryOperation() && FormulaNode.ChildNodes[1].Term.Name != root.ChildNodes[1].Term.Name)
                {
                    return false;
                }
                // CheckIsSameSign(root, FormulaNode);
            }
            return true;
        }
        private void BinaryFunZero(ParseTreeNode root, ref List<ParseTreeNode> cell_args, int depht)
        {
            //var cell_args = new List<ParseTreeNode>();
            if (CheckIsBinary(root.ChildNodes[0]))
            {
                if (CheckIsNoBinFun(root.ChildNodes[0]))
                {
                    cell_args.Add(root.ChildNodes[0]);
                    return;
                }
                if (!CheckIsSameSign(root, root.ChildNodes[0]))
                {
                    var name = root.ChildNodes[0].Print();
                    res.Add(new FormulaNode { Name = name, Parent = res.Last(), Depth = depht, Result = "2" });
                }
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[0]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cell_args.AddRange(cells);
                foreach (var arg in num)
                {
                    cell_args.Add(new ParseTreeNode(new Token(new Terminal("NumberToken"), new SourceLocation(), "test", arg)));
                }
            }
        }
        private void BinaryFunSecond(ParseTreeNode root, ref List<ParseTreeNode> cell_args, int depth)
        {
            if (CheckIsBinary(root.ChildNodes[2]))
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[2]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cell_args.AddRange(cells);
                if (!CheckIsSameSign(root, root.ChildNodes[2]))
                {
                    var name = root.ChildNodes[2].Print();
                    res.Add(new FormulaNode { Name = name, Parent = res.Last(), Depth = depth, Result = "2" });
                }
                if (CheckIsNoBinFun(root.ChildNodes[2]))
                {
                    cell_args.Add(root.ChildNodes[2]);
                    return;
                }

                foreach (var arg in num)
                {
                    cell_args.Add(new ParseTreeNode(new Token(new Terminal("NumberToken"), new SourceLocation(), "test", arg)));
                }
            }
        }
        public string GetJson() 
        {
            var options = new JsonSerializerOptions
            {
                IncludeFields = true,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
                WriteIndented = true
            };
            var json = JsonSerializer.Serialize(res[0], options);
            res.Clear();
            return json;
        }
    }
}
