
namespace ExcelLibrary.Writer
{
    public class PagePrinterGeneric
    {
        public const int VERTICAL = 1;
        public const int HORIZONTAL = 2;

        private int? PWide;
        private int? Orientation;
        private int? PTall;
        private double? Left;
        private double? Right;
        private double? Top;
        private double? Bottom;
        private double? Footer;
        private double? Header;

        public double? MarginHeader
        {
            get { return Header; }
            set { Header = value; }
        }

        public double? MarginFooter
        {
            get { return Footer; }
            set { Footer = value; }
        }

        public double? MarginBottom
        {
            get { return Bottom; }
            set { Bottom = value; }
        }
      
        public double? MarginTop
        {
            get { return Top; }
            set { Top = value; }
        }

        public double? MarginRight
        {
            get { return Right; }
            set { Right = value; }
        }

        public double? MarginLeft
        {
            get { return Left; }
            set { Left = value; }
        } 

        public int? PageTall
        {
            get { return PTall; }
            set { PTall = value; }
        }

        public int? PageOrientation
        {
            get { return Orientation; }
            set { Orientation = value; }
        }

        public int? PageWide
        {
            get { return PWide; }
            set { PWide = value; }
        }

    }
}
