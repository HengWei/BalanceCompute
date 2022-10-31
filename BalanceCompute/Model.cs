using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BalanceCompute
{
    public class Model
    {
    }

    public class SystemData
    {
        public string Store { get; set; } = string.Empty;

        public decimal Cash { get; set; }
    }

    public class BalanceData
    {
        public string Store { get; set; } = string.Empty;

        public decimal LastBalance { get; set; }
            
        public decimal Cash { get; set; }

        public decimal NowBalance { get { return this.Cash + this.LastBalance; } }
    }
}
