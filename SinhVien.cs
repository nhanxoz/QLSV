using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLSV
{
    public class SinhVien
    {
        public string TenSV { get; set; }

        public string LopSV { get; set; }

        public string NSSV { get; set; }

        public string DiaChiSV { get; set; }

        public double diemcc { get; set; }

        public double diemgk { get; set; }

        public double diemck { get; set; }

        public SinhVien()
        {
        }
        public SinhVien(string t,string l,string ns, string dc,double cc, double gk, double ck)
        {
            TenSV = t;
            LopSV = l;
            NSSV = ns;
            DiaChiSV = dc;
            diemcc = cc;
            diemgk = gk;
            diemck = ck;
        }
    }
}
