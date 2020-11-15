using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace components36
{
    public class Setting
    {
        public int value;
        public string legend;
        public Setting()
        {
            string rand_string = "beanyu96(Q)dkpbabe";
            Thread.Sleep(10);
            Random rand = new Random();
            value = rand.Next(5, 100);
            legend = rand_string.Substring(rand.Next(15));
        }
    }
}
