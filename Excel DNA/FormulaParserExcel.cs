﻿using ExcelDna.Integration;
using Irony.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLParser;

namespace Excel_DNA
{
    public class FormulaParserExcel
    {
        public static List<Node> res = new List<Node>();
        public static List<Cell> cells = new List<Cell>();

        public List<Node> GetRes()
        {
            return res;
        }
        public List<Cell> GetCells() { return cells; }
        public void DepthFirstSearch(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, bool flag = false, Node parent = null, bool minus = false, bool binary_operation = false)
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
                res.Add(new Node { Name = name_node, Result = result_node.Item2, Depth = depth.ToString(), Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
            }
            if (root.Term.Name == "CellToken")
            {
                if (binary_operation)
                {
                    var name_node = root.Print();
                    var result_node = RangeSet("=" + name_node);
                    res.Add(new Node { Name = name_node, Result = result_node.Item2, Depth = depth.ToString(), Parent = parent });
                    return;
                }
                var name = root.Token.Text;
                CellSet(name, depth, parent);
                return;
            }
            if (root.Term.Name == "NumberToken")
            {
                var name = root.Token.Text;
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = name, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
                return;
            }
            if (root.Term.Name == "ReferenceFunctionCall" && root.ChildNodes.Count() == 3)
            {
                var name = root.Print();
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = "<диапазон>", Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
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
                        res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null) });
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
                        Tuple<string, string> bin_result = RangeSet("=" + bin_name);
                        bin_name = bin_result.Item1;

                        var cell_args = new List<ParseTreeNode>();

                        res.Add(new Node
                        {
                            Name = bin_name,
                            Depth = depth.ToString(),
                            Result = bin_result.Item2,
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

                Tuple<string, string> result = RangeSet("=" + name);
                name = result.Item1;

                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
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
                res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result, Parent = parent });
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
                    res.Add(new Node { Name = name, Depth = depth.ToString(), Result = result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
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

        public static Tuple<string, string> RangeSet(string formula)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            excelApp.Range["BBB1000"].Formula = formula;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["BBB1000"];
            var res = string.Format("{0:F2}", range.Value);
            return Tuple.Create(range.FormulaLocal.Substring(1), res);
        }

        public static void CellSet(string cellName, int cellDepth, Node parent)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range[cellName];
            if (range.Value == null)
            {
                res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = "<пусто>", Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            else if (range.Value.GetType() == typeof(string))
            {
                res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text, Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
                return;
            }
            var result = range.Text.Replace("#", "@");
            res.Add(new Node { Name = cellName, Depth = cellDepth.ToString(), Result = range.Text.Replace("#", "@"), Parent = (cellDepth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < cellDepth) : null) });
            cells.Add(new Cell { Adress = cellName, Fun = result });
            return;
        }
        //private static void AddBinaryToRes(ParseTreeNode root, Microsoft.Office.Interop.Excel.Application application, int depth, Node parent = null)
        //{
        //    var bin_name = root.Print();
        //    Tuple<string, string> bin_result = RangeSet("=" + bin_name);
        //    bin_name = bin_result.Item1;

        //    res.Add(new Node { Name = bin_name, Depth = depth.ToString(), Result = bin_result.Item2, Parent = (depth >= 2 ? res.Last(x => x.Type == "function" && Convert.ToInt32(x.Depth) < depth) : null), Type = "function" });
        //    DepthFirstSearch(root.ChildNodes[0], application, depth + 1, false, parent = res.Last());
        //    DepthFirstSearch(root.ChildNodes[2], application, depth + 1, false, parent = res.Last());
        //    return;
        //}
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
        private static bool CheckIsSameSign(ParseTreeNode root, ParseTreeNode child)
        {
            foreach (var node in child.ChildNodes)
            {
                if (node.Term.Name == "FunctionCall" && node.IsBinaryOperation() && node.ChildNodes[1].Term.Name != root.ChildNodes[1].Term.Name)
                {
                    return false;
                }
                // CheckIsSameSign(root, node);
            }
            return true;
        }
        private static void BinaryFunZero(ParseTreeNode root, ref List<ParseTreeNode> cell_args, int depht)
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
                    res.Add(new Node { Name = name, Parent = res.Last(), Depth = depht.ToString(), Result = "2" });
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
        private static void BinaryFunSecond(ParseTreeNode root, ref List<ParseTreeNode> cell_args, int depth)
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
                    res.Add(new Node { Name = name, Parent = res.Last(), Depth = depth.ToString(), Result = "2" });
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
    }
}