using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Var35n6n18n27
{
    public partial class UserControlInput : UserControl
    {
        public int Max { set; private get;}
        public int Min { set; private get;}
        public Nullable<int> Result
        {
            get => Result;
            set 
            {
                this.Result = value;
                textBoxInput.Text = Result.ToString();
            }
        }
        public UserControlInput()
        {
            InitializeComponent();
        }
        private void CheckInput()
        {
            if (Max == Min)
            {
                textBoxInformation.Text = "Некоректно макс не ровно мин";
                Result = null;
                return;
            }
            if (Max == Min && Max == 0)
            {
                textBoxInformation.Text = "Мак мин не выставлены";
                Result = null;
                return;
            }
            int Inpunt;
            if (int.TryParse(textBoxInput.Text,out Inpunt))
            {
                if (Inpunt < Max && Inpunt > Min)
                {
                    Result = Inpunt;
                    textBoxInformation.Text = "Число корректно";
                    return;
                }
            } 
            else if (textBoxInformation.Text == "")
            {
                textBoxInformation.Text = "Ничего не введено";
                Result = null;
                return;
            }
            else
            {
                textBoxInformation.Text = "ВВедено не число";
                Result = null;
                return;
            }
        }
        private void TextBoxInput_TextChanged(object sender, EventArgs e)
        {
            CheckInput();
        }

        private void textBoxInput_Enter(object sender, EventArgs e)
        {
            CheckInput();
        }
    }
}
