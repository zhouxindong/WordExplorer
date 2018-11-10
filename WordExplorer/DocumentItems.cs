using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

namespace WordExplorer
{
    // Classes inherited from the Item class provide specialized representation of particular 
    // document nodes by overriding virtual methods and properties of the base class.
    public class DocumentItem : Item
    {
        public DocumentItem(Node node) : base(node)
        {
        }

        public override bool IsRemovable => false;
    }

    public class SectionItem : Item
    {
        public SectionItem(Node node) : base(node)
        {
        }

        public override bool IsRemovable => !Node.Equals(Node.Document.LastSection);
    }

    public class HeaderFooterItem : Item
    {
        public HeaderFooterItem(Node node) : base(node)
        {
        }

        protected override string IconName => ((HeaderFooter)Node).IsHeader ? "Header" : "Footer";

        public override string Name => $"{base.Name} - {((HeaderFooter) Node).HeaderFooterType}";
    }

    public class BodyItem : Item
    {
        public BodyItem(Node node) : base(node)
        {
        }
        public override bool IsRemovable => false;
    }

    public class TableItem : Item
    {
        public TableItem(Node node) : base(node)
        {
        }
    }

    public class RowItem : Item
    {
        public RowItem(Node node) : base(node)
        {
        }
    }

    public class CellItem : Item
    {
        public CellItem(Node node) : base(node)
        {
        }
    }

    public class ParagraphItem : Item
    {
        public ParagraphItem(Node node) : base(node)
        {
        }

        public override bool IsRemovable
        {
            get
            {
                var para = (Paragraph)Node;
                return !para.IsEndOfSection;
            }
        }
    }

    public class RunItem : Item
    {
        public RunItem(Node node) : base(node)
        {
        }
    }

    public class FieldStartItem : Item
    {
        public FieldStartItem(Node node) : base(node)
        {
        }
    }

    public class FieldSeparatorItem : Item
    {
        public FieldSeparatorItem(Node node) : base(node)
        {
        }
    }

    public class FieldEndItem : Item
    {
        public FieldEndItem(Node node) : base(node)
        {
        }
    }

    public class BookmarkStartItem : Item
    {
        public BookmarkStartItem(Node node) : base(node)
        {
        }

        public override string Name => $"{base.Name} - \"{((BookmarkStart) Node).Name}\"";
    }

    public class BookmarkEndItem : Item
    {
        public BookmarkEndItem(Node node) : base(node)
        {
        }

        public override string Name => $"{base.Name} - \"{((BookmarkEnd) Node).Name}\"";
    }

    public class FootnoteItem : Item
    {
        public FootnoteItem(Node node) : base(node)
        {
        }
    }

    public class ShapeItem : Item
    {
        public ShapeItem(Node node) : base(node)
        {
        }

        public override string Name
        {
            get
            {
                Shape shape = (Shape)Node;
                switch (shape.ShapeType)
                {
                    case ShapeType.OleObject:
                    case ShapeType.OleControl:
                        return shape.OleFormat.ProgId;
                    default:
                        return base.IconName;
                }
            }
        }

        protected override string IconName
        {
            get
            {
                Shape shape = (Shape)Node;
                switch (shape.ShapeType)
                {
                    case ShapeType.OleObject:
                        return "OleObject";
                    case ShapeType.OleControl:
                        return "OleControl";
                    default:
                        return base.IconName;
                }
            }
        }

    }

    public class GroupShapeItem : Item
    {
        public GroupShapeItem(Node node) : base(node)
        {
        }
    }

    public class FormFieldItem : Item
    {
        public FormFieldItem(Node node) : base(node)
        {
        }

        public override string Name => $"{base.Name} - \"{((FormField) Node).Name}\"";

        protected override string IconName
        {
            get
            {
                switch (((FormField)Node).Type)
                {
                    case FieldType.FieldFormCheckBox:
                        return "FormCheckBox";
                    case FieldType.FieldFormDropDown:
                        return "FormDropDown";
                    case FieldType.FieldFormTextInput:
                        return "FormTextInput";
                    default:
                        return base.IconName;
                }
            }
        }
    }

    public class SpecialCharItem : Item
    {
        public SpecialCharItem(Node node) : base(node)
        {
        }
    }
}