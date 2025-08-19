using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools.bo
{
    public class CountyComboxBo
    {
        public string CountyId { get; set; }
        public string CountyName { get; set; }

        public override string ToString()
        {
            return "CountyId: " + CountyId + ", CountyName: " + CountyName;
        }
    }
}
