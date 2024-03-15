using System.Text.Json.Serialization;

namespace Excel_DNA.Models
{
    public class FormulaNode
    {
        public string? Name { get; set; }
        public object Result { get; set; }
        public int Depth { get; set; }

        [JsonInclude]
        public List<FormulaNode> Childrens = new List<FormulaNode>();
        public FormulaNode? Parent { get; set; }
        public string? Type { get; set; }
    }
}
