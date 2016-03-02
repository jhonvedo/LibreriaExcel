namespace ExcelLibrary.Writer
{
    public class StyleGeneric
    {
        private string _begin;
        private string _end;
        private bool? _verticalAlign;
        private bool? _horizontalAlign;
        private bool? _wrapText;
        private bool? _bold;
        private System.Drawing.Color? _color;
        private bool? _lineStyle;
        private dynamic _lineWeight;

        public dynamic LineWeight
        {
            get { return _lineWeight; }
            set { _lineWeight = value; }
        }

        public bool? LineStyle
        {
            get { return _lineStyle; }
            set { _lineStyle = value; }
        }

        public System.Drawing.Color? Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public bool? Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }

        public bool? WrapText
        {
            get { return _wrapText; }
            set { _wrapText = value; }
        }

        public bool? HorizontalAlign
        {
            get { return _horizontalAlign; }
            set { _horizontalAlign = value; }
        }

        public bool? VerticalAlign
        {
            get { return _verticalAlign; }
            set { _verticalAlign = value; }
        }

        public string End
        {
            get { return _end; }
            set { _end = value; }
        }

        public string Begin
        {
            get { return _begin; }
            set { _begin = value; }
        }
    }
}