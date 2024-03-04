using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using XLParser;
using Irony.Parsing;
using System.Text.Json;
using System.Text.RegularExpressions;
using IRibbonControl = Microsoft.Office.Core.IRibbonControl;
using Microsoft.AspNetCore.SignalR.Client;
using System.Text.Json.Serialization;
using System.Collections;
using System.Reflection;
using System.Resources;
using System.IO.Compression;
using System.Text;

namespace Excel_DNA
{

    public class Node
    {
        public string? Name { get; set; }
        public string? Result { get; set; }
        public string? Depth { get; set; }

        [JsonInclude]
        public List<Node>? Childrens = new List<Node>();
        public Node? Parent { get; set; }
        public string? Type { get; set; }
    }


    public class Cell
    {
        public string Adress { get; set; }
        public string Fun { get; set; }
    }


    [ComVisible(true)]
    public class MyFunctions: ExcelRibbon, IExcelAddIn
    {
        static List<Node> res = new List<Node>();

        static List<Cell> cells = new List<Cell>();

        static HubConnection connection;

        public static Microsoft.Office.Interop.Excel.Application application1 = new Microsoft.Office.Interop.Excel.Application();

        static MyForm treeForm = new MyForm($"http://localhost:3000/CreateTreePage/?userName={application1.UserName}");

        static bool minus = true;

        public void AutoOpen()
        {
            connection = new HubConnectionBuilder()
            .WithUrl("https://localhost:7108/chathub")
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
        // ShortCut for hide tree
        [ExcelCommand(ShortCut = "{ESC}")]
        public static void HideShortCutTree()
        {
            treeForm.Hide();
        }

        public async static void RangeGet()
        {
            res.Clear();

            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.ActiveCell;

            //range.Text.Replace("#", "@")
            res.Add(new Node { Name = range.AddressLocal.Replace("$",""), Result = string.Format("{0:F2}", range.Value), Depth = "0", Type = "function" });
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
                res.Add(new Node { Name = range.Text, Result = range.Text, Depth = "1" });
                SendMessage();
                return;
            }
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА =text, ="text"
            else if (Regex.IsMatch(range.Formula, allSymbolsPattern) || Regex.IsMatch(range.Formula, stringValuePattern))
            {
                range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                res.Add(new Node { Name = range.FormulaLocal.Substring(1), Result = range.Text.Replace("#", "@"), Depth = "1" });
                SendMessage();
                return;
            }

            range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);

            FormulaParserExcel parser = new FormulaParserExcel();

            parser.DepthFirstSearch(node, excelApp, 1);

            res = parser.GetRes();

            cells = parser.GetCells();

            SendMessage();

        }

        public async static void SendMessage()
        {
            var nodesToRemove = new List<Node>();
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
            var json = System.Text.Json.JsonSerializer.Serialize(res[0], options);
            res.Clear();
            
            try
            {
                await connection.StartAsync();
                await connection.InvokeAsync("Send", application1.UserName, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error invoking hub method: {ex.Message}");
                // Дополнительная обработка ошибки по вашему усмотрению
            }

            await connection.StopAsync();

            treeForm.Show();
        }


        
        public void AutoClose()
        {
            
        }
    }

}