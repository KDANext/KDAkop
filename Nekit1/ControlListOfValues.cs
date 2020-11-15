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
using System.Text.RegularExpressions;

namespace ClassLibraryControlSelected
{
    public partial class ControlListOfValues : UserControl
    {
        private string mainLine;

        private string ChangeValue;

        private event EventHandler _listBoxSelectedElementChange;

        public string SetMainLine
        {
            set 
            { 
                mainLine = value; 
            }
        }

        public string SetChangeValue
        {
            set
            {
                ChangeValue = value;
            }
        }

        public  void GetProperty(Object obj)
        {
            Type t = obj.GetType();
            string _work = null;
            List<object> values =new List<object>();
            PropertyInfo[] props = t.GetProperties();
            string[] main = mainLine.Split(' ');
            foreach (string s in main)
            {
                foreach (var prop in props)
                {
                    if(prop.Name == s)
                    {
                        values.Add(prop.GetValue(obj));
                    }
                }
            }
                _work = mainLine;
                listBoxValues.Items.Add(_work);
        }
        //example
        //валуе {SelectedValue} индекс {SelectedIndex}
        public void ChangeListValue(Object obj,String ChangeValue,int Index)
        {
            string line = Convert.ToString(listBoxValues.Items[Index]);
            string[] main = line.Split(' ');
            Type t = obj.GetType();
            PropertyInfo[] props = t.GetProperties();
            foreach (var prop in props)
            {
                for(int i = 0; i<main.Length;i++)
                {
                    int start = main[i].IndexOf("{");
                    int end = main[i].IndexOf("}");
                    if (start >= 0)
                    {
                        string result = main[i].Substring(start+1 , end-start-1);
                        if(result == ChangeValue && result == prop.Name)
                        {
                            main[i] = Convert.ToString(prop.GetValue(obj));
                        }
                    }
                }
            }
            listBoxValues.Items[Index] = string.Join(" ", main);
        }

        [Category("Спецификация"), Description("Текст выбранной записи")]
        public string SelectedText
        {
            get { return listBoxValues.Text; }
        }

        [Category("Спецификация"), Description("Событие выбора элемента из списка")]
        public event EventHandler ListBoxSelectedElementChange
        {
            add { _listBoxSelectedElementChange += value; }
            remove { _listBoxSelectedElementChange -= value; }
        }

        public ControlListOfValues()
        {
            InitializeComponent();
            listBoxValues.SelectedIndexChanged += (sender, e) => {
                _listBoxSelectedElementChange?.Invoke(sender, e);
            };
        }
    }
}
