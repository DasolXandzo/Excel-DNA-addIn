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
            //var value = _exApp.Evaluate(formula); // не работает в случае формулы =INDIRECT(C7&D7)
            var value = range.Value2;
            if (range.FormulaLocal.Substring(1).ToString() != formula.Substring(1))
            {
                Debug.WriteLine($"{range.FormulaLocal.Substring(1)} != {formula.Substring(1)}");
            }
            //https://xldennis.wordpress.com/2006/11/22/dealing-with-cverr-values-in-net-%E2%80%93-part-i-the-problem/
            //#NULL!       -2146826288
            //#DIV/0!      -2146826281
            //#VALUE!      -2146826273
            //#REF!        -2146826265
            //#NAME?       -2146826259
            //#NUM!        -2146826252
            //#N/A         -2146826246
            if (value is int && (value == -2146826281 || value == -2146826273))
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

        public static FormulaNode? Parse(ParseTreeNode node, FormulaNode? parent = null)
        {
            if (node.Term.GetType() == typeof(KeyTerm)) //Возможно, будет лучше использовать if (node.Term is Terminal)
                //return new FormulaNode { Name = node.Print(), Type = "Operator", Depth = (parent?.Depth ?? 0) + 1, Result = node.Token.ValueString};
                return null;
            
            FormulaNode formulaNode;
            switch (node.Term.Name)
            {
                case GrammarNames.FunctionName: //ссылка на функцию
                case GrammarNames.RefFunctionName: //ссылка на функцию
                case GrammarNames.TokenUDF: //ссылка на пользовательсюку функцию
                    //как отдельную ноду выводить не надо
                    return null;
                case GrammarNames.Constant:
                    return new FormulaNode { Name = node.Print(), Type = "Constant", Depth = (parent?.Depth ?? 0) + 1, Result = node.Print()};
                case GrammarNames.FormulaWithEq:
                    return Parse(node.ChildNodes[1], parent);
                case GrammarNames.Arguments:
                    //Нода Arguments - фиктивная, она только группирует дочерние Argument
                    //Так как аргументов может быть много, а возвращать мы можем только одну ноду, то делаем финт - добавляем аргурменты сразу к родителю, и возвращаем null
                    parent?.Childrens.AddRange(node.ChildNodes.Select(x => Parse(x, parent)).Where(x => x != null).Select(x => x!));
                    return null;
                case GrammarNames.Formula:
                case GrammarNames.Reference:
                case GrammarNames.Cell:
                case GrammarNames.Argument:
                case GrammarNames.NamedRange:
                case GrammarNames.UDFName:
                    return Parse(node.ChildNodes[0], parent);
                case GrammarNames.ReferenceFunctionCall:
                case GrammarNames.FunctionCall:
                case GrammarNames.UDFunctionCall:
                    if (node.IsBinaryOperation())
                    {
                        if (node.IsRange())
                            return new FormulaNode { Name = node.Print(), Type = "Range", Depth = (parent?.Depth ?? 0) + 1 };
                        
                        formulaNode = new FormulaNode { Name = node.Print(), Type = "Expression", Depth = (parent?.Depth ?? 0) + 1 };
                        formulaNode.Childrens.AddRange(node.ChildNodes.Select(x => Parse(x, formulaNode)).Where(x => x != null).Select(x => x!));
                        return formulaNode;
                    }

                    if (node.IsNamedFunction())
                    {
                        formulaNode = new FormulaNode { Name = node.Print(), Type = "Function", Depth = (parent?.Depth ?? 0) + 1 };
                        formulaNode.Childrens.AddRange(node.ChildNodes.Select(x => Parse(x, formulaNode)).Where(x => x != null).Select(x => x!));
                        return formulaNode;
                    }
                    if (node.IsUnaryOperation())
                    {   //бывают вида -1, и бывают вида -C10. Возможно, их стоит отличать
                        return new FormulaNode { Name = node.Print(), Type = "UnaryOperation", Depth = (parent?.Depth ?? 0) + 1 };
                    }
                    //TODO IsUnion
                    return new FormulaNode { Name = node.Print(), Type = "Union?", Depth = (parent?.Depth ?? 0) + 1 };
                case GrammarNames.TokenCell:
                    formulaNode = new FormulaNode{Name = node.Token.Text, Type = "CellLink", Depth = (parent?.Depth ?? 0) + 1};
                    return formulaNode;
                case GrammarNames.TokenName:
                    return new FormulaNode { Name = node.Token.Text, Type = "NamedRangeLink", Depth = (parent?.Depth ?? 0) + 1 };
                    
                default:
                    throw new ApplicationException($"Unknown term {node.Term.Name}");
            }
        }
    }
}
