using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSAY_DipGen
{
    internal class DipGenInputClass
    {
        [Serializable]
        public class DipGenValueInfo 
        { 
            public string Project_Name_ser { get; set; }
            public string FY_ser { get; set; }
            public string Work_Completion_date_ser { get; set; }
            public string Final_Bill_GT_ser { get; set; }
            public string[] Date_value_ser { get; set; }
            public string[] CC_Office_value_ser { get; set; } 
            public string[] CC_Contractor1_value_ser { get; set; }
            public string[] CC_Contractor2_value_ser { get; set; }
            public string[] CC_Contractor3_value_ser { get; set; }
            public string Lvl1_ser { get; set; }
            public string Lvl2_ser { get; set; }
            public string Lvl3_ser { get; set; }
            public string Division_ser { get; set; }
            public string Lvl1_Context_ser { get; set; }
            public string Lvl2_Context_ser { get; set; }


        }
    }
}
