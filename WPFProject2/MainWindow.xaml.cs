using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WPFProject2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<Node> nodes;
        public MainWindow(List<Node> nodes2)
        {
            InitializeComponent();
            nodes = new ObservableCollection<Node>
        {
            new Node
            {
                Name ="Европа",
                Nodes = new ObservableCollection<Node>
                {
                    new Node {Name="Германия", Result = "2" },
                    new Node {Name="Франция", Result = "2" },
                    new Node
                    {
                        Name ="Великобритания",
                        Nodes = new ObservableCollection<Node>
                        {
                            new Node {Name="Англия", Result = "2" },
                            new Node {Name="Шотландия", Result = "2" },
                            new Node {Name="Уэльс", Result = "2" },
                            new Node {Name="Сев. Ирландия", Result = "2" },
                        }
                    }
                }
            },
            new Node
            {
                Name ="Азия",
                Nodes = new ObservableCollection<Node>
                {
                    new Node {Name="Китай", Result = "2" },
                    new Node {Name="Япония" },
                    new Node { Name ="Индия" }
                }
            },
            new Node { Name="Африка" },
            new Node { Name="Америка" },
            new Node { Name="Австралия" }
        };
            treeView1.ItemsSource = nodes2;
        }

        private void treeView1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
    public class Node
    {
        public string? Name { get; set; }
        public string? Depth { get; set; }
        public string? Result { get; set; }
        public ObservableCollection<Node> Nodes { get; set; }
    }
}