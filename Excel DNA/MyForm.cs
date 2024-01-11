using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Wpf;


namespace Excel_DNA
{
    public partial class MyForm : Form
    {
        public MyForm()
        {
            InitializeComponent();
            var test = InitAsync();
        }

        private async Task InitAsync()
        {
            // Предполагается, что у вас есть экземпляр CoreWebView2Environment.
            // Если у вас нет webView, удостоверьтесь, что он объявлен и проинициализирован перед вызовом этого метода.
            // Например: WebView webView = new WebView();
            var path = Path.Combine(Path.GetTempPath(), $"{Environment.UserName}");

            var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(userDataFolder: path);

            // Предполагается, что у вас есть экземпляр webView.
            await webView21.EnsureCoreWebView2Async(env);
            webView21.CoreWebView2.Navigate("https://www.google.com/");
        }

        private void webView21_Click(object sender, EventArgs e)
        {
            webView21.CoreWebView2.Navigate("https://www.google.com/");
        }
    }
}
