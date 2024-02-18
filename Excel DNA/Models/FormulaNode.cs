using Excel_DNA.Core;
using System.Text.Json.Serialization;

namespace Excel_DNA.Models
{
    public class FormulaNode
    {
        public string? Name { get; set; }
        public string? Result { get; set; }
        public string? Depth { get; set; }

        [JsonInclude]
        public List<FormulaNode>? Childrens = new List<FormulaNode>();
        public FormulaNode? Parent { get; set; }
        public string? Type { get; set; }
    }
}
