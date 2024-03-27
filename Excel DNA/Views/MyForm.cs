using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using static System.Windows.Forms.Design.AxImporter;
using System.Text.Json.Serialization;
using System.Text.Json;
using Excel_DNA.Models;


namespace Excel_DNA
{
    public partial class MyForm : Form
    {
        private readonly JsonSerializerOptions _options;

        public MyForm()
        {
            TopMost = true;
            InitializeComponent();
            _options = new JsonSerializerOptions
            {
                IncludeFields = true,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
                WriteIndented = true
            };
        }

        public MyForm(string url) : this()
        {
            var test = InitAsync(url);
        }

        private async Task InitAsync(string url)
        {
            // Предполагается, что у вас есть экземпляр CoreWebView2Environment.
            // Если у вас нет webView, удостоверьтесь, что он объявлен и проинициализирован перед вызовом этого метода.
            // Например: WebView webView = new WebView();
            var path = Path.Combine(Path.GetTempPath(), $"{Environment.UserName}");

            var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(userDataFolder: path);

            // Предполагается, что у вас есть экземпляр webView.
            string test = "test";

            await webView21.EnsureCoreWebView2Async(env);
            string req = $@"javascript:localStorage.setItem(test,test)";
            var res = await webView21.CoreWebView2.ExecuteScriptAsync("Math.sin(Math.PI/2)");

            webView21.CoreWebView2.Navigate(url);
        }

        private void webView21_Click(object sender, EventArgs e)
        {
            // webView21.CoreWebView2.Navigate("https://www.google.com/");
        }

        private void webView21_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var jsonObject = e.TryGetWebMessageAsString();
            var stop = 5;
        }

        private async void webView21_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
        }

        private void MyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Microsoft.Office.Interop.Excel.Range range = excelApp.Range["A1:BBB1000"];
            range.Interior.ColorIndex = 0;
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            // Отменить закрытие формы
            e.Cancel = true;

            // Скрываем форму вместо закрытия
            this.Hide();
        }

        private void webView21_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode== Keys.Escape) {
                this.Hide();
            }
        }

        public void Show(FormulaNode node)
        {
            textBox1.Lines = new[] { JsonSerializer.Serialize(node, _options) };
            Show();
        }
    }
}
