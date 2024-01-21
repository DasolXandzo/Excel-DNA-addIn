using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Wpf;


namespace Excel_DNA
{
    public partial class MyForm : Form
    {
        public MyForm(string url)
        {
            InitializeComponent();
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
    }
}
