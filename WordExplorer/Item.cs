using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;

namespace WordExplorer
{
    /// <summary>
    /// Base class to provide GUI representation for document nodes.
    /// </summary>
    public class Item
    {
        #region statics

        private const string Namespace = "WordExplorer";

        /// <summary>
        /// the class name set whose derived from Item class
        /// </summary>
        private static readonly HashSet<string> ItemClassNameSet = new HashSet<string>();

        /// <summary>
        /// Map of character to string that we use to display control MS Word control characters.
        /// </summary>
        private static readonly Dictionary<char, string> CtlCharacterDic = new Dictionary<char, string>();

        private static readonly List<string> IconNameList = new List<string>();

        private static ImageList _s_image_list;

        public static ImageList ImageList
        {
            get
            {
                if (_s_image_list != null) return _s_image_list;
                _s_image_list = new ImageList
                {
                    ColorDepth = ColorDepth.Depth32Bit,
                    ImageSize = new Size(16, 16)
                };
                return _s_image_list;
            }
        }

        /// <summary>
        /// Static ctor.
        /// </summary>
        static Item()
        {
            foreach (var type in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (type.IsSubclassOf(typeof(Item)) && !type.IsAbstract)
                {
                    ItemClassNameSet.Add(type.Name);
                }
            }

            // Fill control chars fields set
            CtlCharacterDic.Add(ControlChar.CellChar, "[!Cell!]");
            CtlCharacterDic.Add(ControlChar.ColumnBreakChar, "[!ColumnBreak!]\r\n");
            CtlCharacterDic.Add(ControlChar.FieldEndChar, "[!FieldEnd!]");
            CtlCharacterDic.Add(ControlChar.FieldSeparatorChar, "[!FieldSeparator!]");
            CtlCharacterDic.Add(ControlChar.FieldStartChar, "[!FieldStart!]");
            CtlCharacterDic.Add(ControlChar.LineBreakChar, "[!LineBreak!]\r\n");
            CtlCharacterDic.Add(ControlChar.LineFeedChar, "[!LineFeed!]");
            CtlCharacterDic.Add(ControlChar.NonBreakingHyphenChar, "[!NonBreakingHyphen!]");
            CtlCharacterDic.Add(ControlChar.NonBreakingSpaceChar, "[!NonBreakingSpace!]");
            CtlCharacterDic.Add(ControlChar.OptionalHyphenChar, "[!OptionalHyphen!]");
            CtlCharacterDic.Add(ControlChar.ParagraphBreakChar, "[ParagraphBreak]r\n");
            CtlCharacterDic.Add(ControlChar.SectionBreakChar, "[!SectionBreak!]\r\n");
            CtlCharacterDic.Add(ControlChar.TabChar, "[!Tab!]");
        }

        /// <summary>
        /// Item class factory implementation.
        /// </summary>
        public static Item CreateItem(Node node)
        {
            var type_name = node.NodeType + "Item";
            if (ItemClassNameSet.Contains(type_name))
                return (Item) Activator.CreateInstance(Type.GetType(Namespace + "." + type_name), node);
            return new Item(node);
        }

        #endregion

        #region class

        public Node Node { get; }

        /// <summary>
        /// Creates Item for the document node.
        /// </summary>
        public Item(Node node)
        {
            Node = node;
        }

        #endregion

        #region virtual members

        /// <summary>
        ///  DisplayName for this Item. Can be customized by overriding in inheriting classes.
        /// </summary>
        public virtual string Name => Node.NodeType.ToString();

        /// <summary>
        /// Icon for this node can be customized by overriding this property in the inheriting classes.
        /// The name represents name of .ico file without extension located in the Icons folder of the project.
        /// </summary>
        protected virtual string IconName => GetType().Name.Replace("Item", "");

        public virtual bool IsRemovable => true;

        #endregion

        #region public members

        /// <summary>
        /// Text contained by the corresponding document node.
        /// </summary>
        public string Text
        {
            get
            {
                var result = new StringBuilder();

                // All control characters are converted to human readable form.
                // E.g. [!PageBreak!], [!ParagraphBreak!], etc.
                foreach (var c in Node.GetText())
                {
                    var control_char_display = CtlCharacterDic[c];
                    if (control_char_display == null)
                        result.Append(c);
                    else
                        result.Append(control_char_display);
                }

                return result.ToString();
            }
        }

        private TreeNode _tree_node;

        /// <summary>
        /// Creates TreeNode for this item to be displayed in Document Explorer TreeView control.
        /// </summary>
        public TreeNode TreeNode
        {
            get
            {
                if (_tree_node != null) return _tree_node;
                _tree_node = new TreeNode(Name);
                if (!IconNameList.Contains(IconName))
                {
                    IconNameList.Add(IconName);
                    ImageList.Images.Add(Icon);
                }
                int index = IconNameList.IndexOf(IconName);
                _tree_node.ImageIndex = index;
                _tree_node.SelectedImageIndex = index;
                _tree_node.Tag = this;
                if (Node is CompositeNode && ((CompositeNode)Node).ChildNodes.Count > 0)
                {
                    _tree_node.Nodes.Add("#dummy");
                }
                return _tree_node;
            }
        }


        private Icon _icon;
        /// <summary>
        /// Icon to display in the Document Explorer TreeView control.
        /// </summary>
        public Icon Icon
        {
            get
            {
                if (_icon == null)
                {
                    _icon = LoadIcon(IconName);
                    if (_icon == null)
                        _icon = LoadIcon("Node");
                }
                return _icon;
            }
        }
        /// <summary>
        /// Loads icon from assembly resource stream.
        /// </summary>
        /// <param name="icon_name">Name of the icon to load.</param>
        /// <returns>Icon object or null if icon was not found in the resources.</returns>
        private Icon LoadIcon(string icon_name)
        {
            var resource_name = "Icons." + icon_name + ".ico";
            var icon_stream = FetchResourceStream(resource_name);

            return icon_stream != null ? new Icon(icon_stream) : null;
        }

        /// <summary>
        /// Returns a resource stream from the executing assembly or throws if the resource cannot be found.
        /// </summary>
        /// <param name="resource_name">The name of the resource without the name of the assembly.</param>
        /// <returns>The stream. Don't forget to close it when finished.</returns>
        internal static Stream FetchResourceStream(string resource_name)
        {
            var asm = Assembly.GetExecutingAssembly();
            var full_name = $"{asm.GetName().Name}.{resource_name}";
            var stream = asm.GetManifestResourceStream(full_name);

            // Ugly optimization so conversion to VB.NET can work.
            while (stream == null)
            {
                var dot_pos = full_name.IndexOf(".", StringComparison.Ordinal);
                if (dot_pos < 0)
                    return null;

                full_name = full_name.Substring(dot_pos + 1);
                stream = asm.GetManifestResourceStream(full_name);
            }

            return stream;
        }

        /// <summary>
        /// Provides lazy on-expand loading of underlying tree nodes.
        /// </summary>
        public void OnExpand()
        {
            // Optimized to allow automatic conversion to VB.NET
            if (TreeNode.Nodes[0].Text.Equals("#dummy"))
            {
                TreeNode.Nodes.Clear();
                foreach (Node n in ((CompositeNode)Node).ChildNodes)
                {
                    TreeNode.Nodes.Add(Item.CreateItem(n).TreeNode);
                }
            }
        }

        public void Remove()
        {
            if (IsRemovable)
            {
                Node.Remove();
                _tree_node.Remove();
            }
        }

        #endregion
    }
}