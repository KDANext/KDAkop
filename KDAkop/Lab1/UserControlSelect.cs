using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using static System.Windows.Forms.CheckedListBox;

namespace Var35n6n18n27
{

    public partial class UserControlSelect : UserControl
    {
        public Object[] List { 
            set 
            {
                checkedListBox.Items.Clear();
                checkedListBox.Items.AddRange(value); 
            } 
        }

        public List<int> CheckIndexs
        {
            get
            {
                return checkedListBox.CheckedIndices.OfType<int>().ToList();
            }
            set
            {
                foreach(var index in value)
                {
                    checkedListBox.SetItemChecked(index, true);
                }
            }
        }

        public UserControlSelect()
        {
            InitializeComponent();
        }        
    }
}
