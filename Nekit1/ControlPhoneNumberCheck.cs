using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ClassLibraryControlSelected
{
    public partial class ControlPhoneNumberCheck : UserControl
    {
        string NumberReg = @"[+7|8]{1}-[0-9]{3}-[0-9]{3}-[0-9]{2}-[0-9]{2}";
        private Color _сolor;
        private event EventHandler _phoneNumberWrite;


        [Category("Спецификация"), Description("Цвет поля")]
        public Color setColor
        {
            set
            {
                _сolor = value;
            }
        }
        [Category("Спецификация"), Description("Регулярное выражение")]
        public string regularValue
        {
            set
            {
                NumberReg = value;
            }
        }

        [Category("Спецификация"), Description("Событие ввода номера")]
        public event EventHandler NumberWrite
        {
            add { _phoneNumberWrite += value; }
            remove { _phoneNumberWrite -= value; }
        }

        public string GetNumber()
        {
            
            if (!Regex.IsMatch(textBoxNumber.Text, NumberReg, RegexOptions.IgnoreCase))
            {
                textBoxNumber.BackColor = _сolor;
                return "";
            }
            else
            {
                textBoxNumber.BackColor = Color.White;
                 return textBoxNumber.Text;
            }
        }

        public ControlPhoneNumberCheck()
        {
            InitializeComponent();
            textBoxNumber.TextChanged += (sender, e) => {
                _phoneNumberWrite?.Invoke(sender, e);
            };
        }
    }
}
