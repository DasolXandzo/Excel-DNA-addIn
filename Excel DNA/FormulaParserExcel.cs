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
        private List<FormulaNode> _res = new List<FormulaNode>();
        private List<Cell> _cells = new List<Cell>();
        static ExcelApplicaton _exApp;
        public FormulaParserExcel(ExcelApplicaton app)
        {
            _exApp = app; 
        }
        public List<FormulaNode> GetRes()
        {
            var nodesToRemove = new List<FormulaNode>();
            foreach (var tempNode in _res)
            {
                tempNode.Childrens.AddRange(_res.Where(x => x.Parent == tempNode));
                nodesToRemove.AddRange(_res.Where(x => x.Parent == tempNode));
            }
            foreach (var nodeToRemove in nodesToRemove)
            {
                _res.Remove(nodeToRemove);
            }
            if (_res.Count > 1)
            {
                _res[0].Childrens.Add(_res[1]);
            }
            return _res;
        }
        public void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, bool flag = false, FormulaNode parent = null, bool minus = false, bool binaryOperation = false)
        {
            switch (root.Term.Name)
            {
                case "Number_new": break;
                case "CellToken": break;
                case "NumberToken": break;

            }
            if (root.Term.Name == "Number_new")
            {
                var nameNode = root.Token.ValueString;
                var resultNode = RangeSet("=" + nameNode);
                _res.Add(new FormulaNode { Name = nameNode, Result = resultNode, Depth = depth, Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
            }
            if (root.Term.Name == "CellToken")
            {
                if (binaryOperation)
                {
                    var nameNode = root.Print();
                    var resultNode = RangeSet("=" + nameNode);
                    _res.Add(new FormulaNode { Name = nameNode, Result = resultNode, Depth = depth, Parent = parent });
                    return;
                }
                var name = root.Token.Text;
                CellSet(name, depth, parent);
                return;
            }
            if (root.Term.Name == "NumberToken")
            {
                var name = root.Token.Text;
                _res.Add(new FormulaNode { Name = name, Depth = depth, Result = root.Token.Value, Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.Term.Name == "ReferenceFunctionCall" && root.ChildNodes.Count() == 3)
            {
                var name = root.Print();
                _res.Add(new FormulaNode { Name = name, Depth = depth, Result = "<диапазон>", Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
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
                        _res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
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
                        var binName = root.Print();
                        var binResult = RangeSet("=" + binName);
                        //bin_name = bin_result.Item1;

                        var cellArgs = new List<ParseTreeNode>();

                        _res.Add(new FormulaNode
                        {
                            Name = binName,
                            Depth = depth,
                            Result = binResult,
                            Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null),
                            Type = "function"
                        });

                        BinaryFunZero(root, ref cellArgs, depth);

                        BinaryFunSecond(root, ref cellArgs, depth);

                        if (cellArgs.Count > 0)
                        {
                            for (int i = 0; i < cellArgs.Count; i++)
                            {
                                DepthFirstSearch(cellArgs[i], application, depth + 1, false, _res.Last());
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
                                DepthFirstSearch(root.ChildNodes[i], application, depth + 1, false, _res.Last());
                            }
                        }
                        return;
                    }
                    return;
                }

                var name = root.Print();

                var result = RangeSet("=" + name);
                //name = result.Item1;

                _res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1, false, parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) <= depth) : null));
                }
                return;
            }
            if (root.IsRange())
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root);
                var name = root.Print();
                var result = "range";
                _res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = parent });
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
                    _res.Add(new FormulaNode { Name = name, Depth = depth, Result = result, Parent = (depth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
                }
                if (root.ChildNodes.Count == 1 && root.ChildNodes[0].IsParentheses()) //проверка внутри только скобки
                {
                    DepthFirstSearch(root.ChildNodes[0], application, depth, true, parent);
                    return;
                }
                foreach (var child in root.ChildNodes)
                {
                    DepthFirstSearch(child, application, depth + 1, false, _res.Last());
                }
                return;
            }


            foreach (var child in root.ChildNodes)
            {
                DepthFirstSearch(child, application, depth, false, parent, minus, binaryOperation);
            }
        }

        public static object RangeSet(string formula)
        {
            _exApp.Range["BBB1000"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = _exApp.Range["BBB1000"];
            var value = _exApp.Evaluate(formula);
            if (range.FormulaLocal.Substring(1).ToString() != formula.Substring(1))
            {
                Debug.WriteLine($"{range.FormulaLocal.Substring(1)} != {formula.Substring(1)}");
            }
            if (value == -2146826281) // #DIV/0!
                return "error";
            return value;
        }

        public void CellSet(string cellName, int cellDepth, FormulaNode parent)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range[cellName];
            if (range.Value == null)
            {
                _res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, Result = "<пусто>", Parent = (cellDepth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            if (range.Value is string)
            {
                _res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, 
                    Result = range.Text, //todo подумать над тем, чтобы сохранять и range.Text и range.Value
                    Parent = (cellDepth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            var result = range.Text.Replace("#", "@");
            //var result = range.Text;
            if (result.ToString() != range.Text.ToString())
            {
                Debug.WriteLine($"{result.ToString()} != {range.Text.ToString()}");
            }
            _res.Add(new FormulaNode { Name = cellName, Depth = cellDepth, 
                Result = range.Value, //todo тут было range.Text.Replace("#", "@"), найти случаи, когда это было нужно
                Parent = (cellDepth >= 2 ? _res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
            _cells.Add(new Cell { Adress = cellName, Fun = result });
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
        private static bool CheckIsBinary(ParseTreeNode formulaNode)
        {
            if (formulaNode.IsBinaryOperation()) return true;
            foreach (var child in formulaNode.ChildNodes)
            {
                if (CheckIsBinary(child))
                {
                    return true;
                }
            }
            return false;
        }
        private static bool CheckIsNoBinFun(ParseTreeNode formulaNode)
        {
            if (formulaNode.IsFunction() && !formulaNode.IsBinaryOperation()) return true;
            foreach (var child in formulaNode.ChildNodes)
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
            foreach (var formulaNode in child.ChildNodes)
            {
                if (formulaNode.Term.Name == "FunctionCall" && formulaNode.IsBinaryOperation() && formulaNode.ChildNodes[1].Term.Name != root.ChildNodes[1].Term.Name)
                {
                    return false;
                }
                // CheckIsSameSign(root, FormulaNode);
            }
            return true;
        }
        private void BinaryFunZero(ParseTreeNode root, ref List<ParseTreeNode> cellArgs, int depht)
        {
            //var cell_args = new List<ParseTreeNode>();
            if (CheckIsBinary(root.ChildNodes[0]))
            {
                if (CheckIsNoBinFun(root.ChildNodes[0]))
                {
                    cellArgs.Add(root.ChildNodes[0]);
                    return;
                }
                if (!CheckIsSameSign(root, root.ChildNodes[0]))
                {
                    var name = root.ChildNodes[0].Print();
                    _res.Add(new FormulaNode { Name = name, Parent = _res.Last(), Depth = depht, Result = "2" });
                }
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[0]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cellArgs.AddRange(cells);
                foreach (var arg in num)
                {
                    cellArgs.Add(new ParseTreeNode(new Token(new Terminal("NumberToken"), new SourceLocation(), "test", arg)));
                }
            }
        }
        private void BinaryFunSecond(ParseTreeNode root, ref List<ParseTreeNode> cellArgs, int depth)
        {
            if (CheckIsBinary(root.ChildNodes[2]))
            {
                FormulaAnalyzer analyzer = new FormulaAnalyzer(root.ChildNodes[2]);
                var cells = analyzer.AllNodes.Where(x => x.Term.ToString() == "Cell");
                var col = root.GetFunctionArguments();
                var num = analyzer.Numbers();
                cellArgs.AddRange(cells);
                if (!CheckIsSameSign(root, root.ChildNodes[2]))
                {
                    var name = root.ChildNodes[2].Print();
                    _res.Add(new FormulaNode { Name = name, Parent = _res.Last(), Depth = depth, Result = "2" });
                }
                if (CheckIsNoBinFun(root.ChildNodes[2]))
                {
                    cellArgs.Add(root.ChildNodes[2]);
                    return;
                }

                foreach (var arg in num)
                {
                    cellArgs.Add(new ParseTreeNode(new Token(new Terminal("NumberToken"), new SourceLocation(), "test", arg)));
                }
            }
        }
        public static string GetJson(FormulaNode node) 
        {
            var options = new JsonSerializerOptions
            {
                IncludeFields = true,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
                WriteIndented = true
            };
            var json = JsonSerializer.Serialize(node, options);
            return json;
        }
    }
}
