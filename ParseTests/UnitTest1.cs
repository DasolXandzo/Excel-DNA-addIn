using Excel_DNA;
using Excel_DNA.Core;
using Excel_DNA.Models;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;
using Moq;
using Newtonsoft.Json;
using System;
using System.IO;
using XLParser;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;

namespace ParseTests
{
    public class ParseTests
    {
        ExcelApplicaton excelApplicaton;
        public ParseTests() {
            excelApplicaton = new ExcelApplicaton();
            Workbook wb = excelApplicaton.Workbooks.Open(@"C:\Users\ivank\source\repos\Excel-DNA-addIn\ParseTests\bin\Debug\net6.0-windows\test_v2.xlsm");
            excelApplicaton = ExApp.GetInstance();
        }

        public void Dispose()
        {
            excelApplicaton.Quit();
        }

        [Theory]
        [MemberData(nameof(GetTestCases))]
        
        public void GetRes_Returns_Expected_Json(string cellWithFormula, FormulaNode expected)
        {
            // Arrange
            var node = ExcelFormulaParser.Parse((string)excelApplicaton.Range[cellWithFormula].Formula);

            var formulaParser = new FormulaParserExcel(excelApplicaton);

            formulaParser.DepthFirstSearch(node, excelApplicaton, 1);

            // Act
            var actual = formulaParser.GetRes()[0];

            // Assert
            AssertFormulaNode(expected, actual);
        }

        public static IEnumerable<object[]> GetTestCases()
        {
            
            yield return new object[]
            {
                "G4",
                new FormulaNode
                {
                    Depth = 1, Name = "-----1/3", Result = -0.33333333333333331d, Parent = null, Type = "function",
                    Childrens = new List<FormulaNode>
                    {
                        new() { Depth = 2, Name = "-----1", Result = -1d, Parent = null, Type = null },
                        new() { Depth = 2, Name = "3", Result = 3, Parent = null, Type = null }
                    }
                }
            };
            
            yield return new object[]
            {
                "G5",
                new FormulaNode
                {
                    Depth = 1, Name = "-(-(-(-1)))+C5", Result = 2d, Parent = null, Type = "function",
                    Childrens = new List<FormulaNode>
                    {
                        new() { Depth = 2, Name = "-(-(-(-1)))", Result = 1d, Parent = null, Type = null },
                        new() { Depth = 2, Name = "C5", Result = 1d, Parent = null, Type = null }
                    }
                }
            };
            
            yield return new object[]
            {
                "G6",
                new FormulaNode
                {
                    Depth = 1, Name = "1/0", Result = "error", Parent = null, Type = "function",
                    Childrens = new List<FormulaNode>
                    {
                        new() { Depth = 2, Name = "1", Result = 1, Parent = null, Type = null },
                        new() { Depth = 2, Name = "0", Result = 0, Parent = null, Type = null }
                    }
                }
            };
            
            yield return new object[]
            {
                "G10",
                new FormulaNode
                {
                    Depth = 1, Name = "SUM(AVERAGE(C10:D10),SUM(C10:D10,AVERAGE(C10:D10)),-C10,-1)", Result = 4d, Parent = null, Type = "function",
                    Childrens = new List<FormulaNode>
                    {
                        new() { Depth = 2, Name = "AVERAGE(C10:D10)", Result = 1.5d, Parent = null, Type = "function", Childrens = new List<FormulaNode>
                        {
                            new() { Depth = 3, Name = "C10:D10", Result = "<диапазон>", Parent = null, Type = null }
                        }},
                        new() { Depth = 2, Name = "SUM(C10:D10,AVERAGE(C10:D10))", Result = 4.5d, Parent = null, Type = "function", Childrens = new List<FormulaNode>
                        {
                            new() { Depth = 3, Name = "C10:D10", Result = "<диапазон>", Parent = null, Type = null },
                            new() { Depth = 3, Name = "AVERAGE(C10:D10)", Result = 1.5d, Parent = null, Type = "function", Childrens = new List<FormulaNode>
                            {
                                new() { Depth = 4, Name = "C10:D10", Result = "<диапазон>", Parent = null, Type = null }
                            }}
                        }},
                        new() { Depth = 2, Name = "-C10", Result = -1d, Parent = null, Type = null },
                        new() { Depth = 2, Name = "-1", Result = -1d, Parent = null, Type = null },
                    }
                }
            };
            
            yield return new object[]
            {
                "G26",
                new FormulaNode
                {
                    Depth = 1, Name = "C26+1", Result = 4d, Parent = null, Type = "function",
                    Childrens = new List<FormulaNode>
                    {
                        new() { Depth = 2, Name = "C26", Result = 3d, Parent = null, Type = null },
                        new() { Depth = 2, Name = "1", Result = 1, Parent = null, Type = null }
                    }
                }
            };
        }

        /// <summary>
        /// Сравнивает две ноды с учетом наследников. 
        /// Отличается от стандартного Assert.Equal(expected, actual) тем, что сранивает только некоторые поля, но зато выводит более понятное сообщение при различиях
        /// </summary>
        /// <param name="expected"></param>
        /// <param name="actual"></param>
        private static void AssertFormulaNode(FormulaNode expected, FormulaNode actual)
        {
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.Depth, actual.Depth);
            Assert.Equal(expected.Type, actual.Type);
            Assert.Equal(expected.Result, actual.Result);
            
            Assert.Equal(expected.Childrens.Count, actual.Childrens.Count);
            for (var i = 0; i < expected.Childrens.Count; i++)
            {
                AssertFormulaNode(expected.Childrens[i], actual.Childrens[i]);
            }
        }
    }
}