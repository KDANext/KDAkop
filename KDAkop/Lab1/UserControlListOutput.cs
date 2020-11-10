using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace Var35n6n18n27
{
    public partial class UserControlListOutput : UserControl
    {
        List<object> items = new List<object>();
        String[] Order;
        public UserControlListOutput()
        {
            InitializeComponent();
        }

        public void Add(object item)
        {
            items.Add(item);
            UpdateTree();
        }

        public void AddAll(IEnumerable<object> items)
        {
            this.items.AddRange(items);
        }

        public void SetOrder(String[] order)
        {
            this.Order = order;
            UpdateTree();
        }

        public void UpdateTree()
        {
            if (Order == null) return;
            treeView.BeginUpdate();
            var Node = treeView.Nodes;
            Type t = items[0].GetType();
            PropertyInfo[] props = t.GetProperties();
            foreach(var obj in items)
            {
                foreach(var next in Order)
                {
                    var prop = props.First(x => x.Name == next);
                    Node = AddNode(Node, prop.GetValue(obj).ToString()).Nodes;
                }
                Node = treeView.Nodes;
            }
        }

        private TreeNode AddNode(TreeNodeCollection node, String text)
        {
            if (node== null) return null;
            if (node.Count == 0)
            {
                return node.Add(text);
            }
            else
            {
                for (int i = 0; i < node.Count; i++)
                {
                    if (node[i].Text == text)
                    {
                        return node[i];
                    }
                }
                return node.Add(text);
            }
        }
    }
}
