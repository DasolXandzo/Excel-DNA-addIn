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
        static MyForm treeForm = new MyForm($"http://localhost:3000/CreateTreePage/?userName={exApp.UserName}");


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

            // TODO:Ivanco:регулярки обьявляем так - Regex regex = new Regex(@"туп(\w*)"); 
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
                // TODO: Ivanco: range.Value.GetType() - один раз вычисляем в переменную - затем используем
                // TODO: Ivanco: isNumeric - переменная - вычисляем bool затем используем
                if (range.Value.GetType() == typeof(int) || range.Value.GetType() == typeof(float) || range.Value.GetType() == typeof(double))
                {
                    // TODO:Ivanco:окрашиваем начальную ячейку в розовый - 100500 комментов одного и того же в коде.
                    // это что какая то сложная строка? - все эти комменты удалить
                    // TODO:Ivanco: Color.Pink - если это какой то стандартный цвет, в static read only , на уровне приложения.
                    range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                    SendMessage("");
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
                SendMessage("");
                return;
            }
            // TODO: Ivanco: если нужно что то доделать - используйте TODO
            // ТУТ ДОЛЖНА БЫТЬ ПРОВЕРКА НА ЗНАЧЕНИЕ ЯЧЕЙКИ ФОРМАТА =text, ="text"
            else if (Regex.IsMatch(range.Formula, allSymbolsPattern) || Regex.IsMatch(range.Formula, stringValuePattern))
            {
                range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый
                // TODO: Ivanco: сделать конструкторы для класса.
                // инициализация через именованные параметры выглядит очень громоздко, здесь и по всему коду дальше.
                res.Add(new FormulaNode { Name = range.FormulaLocal.Substring(1), Result = range.Text.Replace("#", "@"), Depth = "1" });
                SendMessage("");
                return;
            }

            range.Interior.Color = Color.Pink; // окрашиваем начальную ячейку в розовый


            ParseTreeNode node =  ExcelFormulaParser.Parse(range.Formula);

            FormulaParserExcel parser = new FormulaParserExcel();


            parser.DepthFirstSearch(node, exApp, 1);

            res = parser.GetRes();

            cells = parser.GetCells();

            string json = parser.GetJson();

            SendMessage(json);

        }

        public async static void SendMessage(string json)
        {
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
            await connection.StopAsync();

            treeForm.Show();
        }

        
        public void AutoClose()
        {
            
        }
    }

}