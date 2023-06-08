using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TXTLOG_TO_EXCEL_Project.Module
{
    public class Eff_Noise_Column
    {
        public static string SEQ { get; set; }
        public static string Name { get; set; }

        public static string Ripple12V { get; set; }
        public static string Ripple12Vsb { get; set; }
        public static string RipplePWOK { get; set; }
        public static string RippleVin_Good { get; set; }
        public static string RippleSMBAlert { get; set; }

        public static string Vin { get; set; }
        public static string Frequ { get; set; }

        public static string Vout_Min12V { get; set; }
        public static string Vout_Min12Vsb { get; set; }
        public static string Vout_MinPWOK { get; set; }
        public static string Vout_MinVin_Good { get; set; }
        public static string Vout_MinSMBAlert { get; set; }

        public static string Vout_Max12V { get; set; }
        public static string Vout_Max12Vsb { get; set; }
        public static string Vout_MaxPWOK { get; set; }
        public static string Vout_MaxVin_Good { get; set; }
        public static string Vout_MaxSMBAlert { get; set; }

        public static string Load12V { get; set; }
        public static string Load12Vsb { get; set; }
        public static string LoadPWOK { get; set; }
        public static string LoadVin_Good { get; set; }
        public static string LoadSMBAlert { get; set; }
    }
}
