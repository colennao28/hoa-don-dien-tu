using System;
using System.Collections.Generic;
using System.Text;

namespace hoa_don_dien_tu.Models
{
    public class HoaDon
    {
        public string khhdon { get; set; } // ki hieu hoa don
        public string nmten { get; set; }
        public string nmmst { get; set; }
        public string nmdchi { get; set; }
        public string shdon { get; set; }
        public DateTime? nky { get; set; } // ngay hop dong
        public string nbten { get; set; }
        public string nbmst { get; set; }
        public string nbdchi { get; set; }
        public float? TSTSauThue { get; set; }
        public float? TongVAT { get; set; }
        public string LoaiTien { get; set; }
        public float? TiGia { get; set; }
        public string TenHangHoa { get; set; }
        public string DonViTinh { get; set; }
        public int? SoLuong { get; set; }
        public float? DonGia { get; set; }
        public float? TiLeCK { get; set; }
        public float? SoTienCK { get; set; }
        public float? ThanhTien { get; set; }
        public float? ThueSuat { get; set; }
        public float? TienThueVAT { get; set; }
        public string CheckUnique { get; set; }
    }
}
