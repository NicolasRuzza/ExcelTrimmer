using System.Windows.Controls;

namespace ExtratorDeConteudo.Class
{
    public class RegexField
    {
        public TextBox TxtColumn { get; set; }
        public TextBox TxtRegexFilter { get; set; }
        public TextBox TxtRegexReplace { get; set; }
        public TextBox TxtOverride { get; set; }
        public StackPanel PnlLineContainer { get; set; }

        public RegexField(TextBox txtColumn, TextBox txtRegexFilter, TextBox txtRegexReplace, TextBox txtOverride, StackPanel pnlLineContainer) { 
            this.TxtColumn = txtColumn;
            this.TxtRegexFilter = txtRegexFilter;
            this.TxtRegexReplace = txtRegexReplace;
            this.TxtOverride = txtOverride;
            this.PnlLineContainer = pnlLineContainer;
        }
    }
}