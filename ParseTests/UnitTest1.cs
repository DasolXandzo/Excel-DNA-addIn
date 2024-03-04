using Excel_DNA;
using Excel_DNA.Core;
using Irony.Parsing;
using Newtonsoft.Json;
using XLParser;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;

namespace ParseTests
{
    public class ParseTests
    {
        [Theory]
        [InlineData(@"
        {
            ""Name"": ""G5"",
            ""Result"": ""2,00"",
            ""Depth"": ""0"",
            ""Parent"": null,
            ""Type"": ""function"",
            ""Childrens"": [
                {
                    ""Name"": ""-(-(-(-1)))\u002BC5"",
                    ""Result"": ""2,00"",
                    ""Depth"": ""1"",
                    ""Parent"": null,
                    ""Type"": ""function"",
                    ""Childrens"": [
                        {
                            ""Name"": ""-(-(-(-1)))"",
                            ""Result"": ""1,00"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": null,
                            ""Childrens"": []
                        },
                        {
                            ""Name"": ""C5"",
                            ""Result"": ""1"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": null,
                            ""Childrens"": []
                        }
                    ]
                }
            ]
        }", "=-(-(-(-1)))+C5")]
        [InlineData(@"
        {
            ""Name"": ""G10"",
            ""Result"": ""4,00"",
            ""Depth"": ""0"",
            ""Parent"": null,
            ""Type"": ""function"",
            ""Childrens"": [
                {
                    ""Name"": ""\u0421\u0423\u041C\u041C(\u0421\u0420\u0417\u041D\u0410\u0427(C10:D10);\u0421\u0423\u041C\u041C(C10:D10;\u0421\u0420\u0417\u041D\u0410\u0427(C10:D10));-C10;-1)"",
                    ""Result"": ""4,00"",
                    ""Depth"": ""1"",
                    ""Parent"": null,
                    ""Type"": ""function"",
                    ""Childrens"": [
                        {
                            ""Name"": ""\u0421\u0420\u0417\u041D\u0410\u0427(C10:D10)"",
                            ""Result"": ""1,50"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": ""function"",
                            ""Childrens"": [
                                {
                                    ""Name"": ""C10:D10"",
                                    ""Result"": ""\u003C\u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D\u003E"",
                                    ""Depth"": ""3"",
                                    ""Parent"": null,
                                    ""Type"": null,
                                    ""Childrens"": []
                                }
                            ]
                        },
                        {
                            ""Name"": ""\u0421\u0423\u041C\u041C(C10:D10;\u0421\u0420\u0417\u041D\u0410\u0427(C10:D10))"",
                            ""Result"": ""4,50"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": ""function"",
                            ""Childrens"": [
                                {
                                    ""Name"": ""C10:D10"",
                                    ""Result"": ""\u003C\u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D\u003E"",
                                    ""Depth"": ""3"",
                                    ""Parent"": null,
                                    ""Type"": null,
                                    ""Childrens"": []
                                },
                                {
                                    ""Name"": ""\u0421\u0420\u0417\u041D\u0410\u0427(C10:D10)"",
                                    ""Result"": ""1,50"",
                                    ""Depth"": ""3"",
                                    ""Parent"": null,
                                    ""Type"": ""function"",
                                    ""Childrens"": [
                                        {
                                            ""Name"": ""C10:D10"",
                                            ""Result"": ""\u003C\u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D\u003E"",
                                            ""Depth"": ""4"",
                                            ""Parent"": null,
                                            ""Type"": null,
                                            ""Childrens"": []
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            ""Name"": ""-C10"",
                            ""Result"": ""-1,00"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": null,
                            ""Childrens"": []
                        },
                        {
                            ""Name"": ""-1"",
                            ""Result"": ""-1,00"",
                            ""Depth"": ""2"",
                            ""Parent"": null,
                            ""Type"": null,
                            ""Childrens"": []
                        }
                    ]
                }
            ]
        }", "=SUM(AVERAGE(C10:D10), SUM(C10:D10, AVERAGE(C10:D10)), -C10, -1)")]
        public void GetRes_Returns_Expected_Json(string expectedJson, string formula)
        {
            // Arrange
            ExcelApplicaton exApp = ExApp.GetInstance();

            ParseTreeNode node = ExcelFormulaParser.Parse(formula);

            FormulaParserExcel formulaParser = new FormulaParserExcel();


            formulaParser.DepthFirstSearch(node, exApp, 1);

            // Act
            string resultJson = formulaParser.GetJson();

            // Assert
            Assert.NotNull(resultJson);

            // Сравнение строк JSON
            Assert.Equal(
                JsonConvert.SerializeObject(JsonConvert.DeserializeObject(expectedJson)),
                JsonConvert.SerializeObject(JsonConvert.DeserializeObject(resultJson))
            );
        }
    }
}