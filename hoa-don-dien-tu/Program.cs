using ClosedXML.Excel;
using hoa_don_dien_tu.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;

namespace hoa_don_dien_tu
{
    class Program
    {
        static void Main(string[] args)
        {
            // read file
            var lines = File.ReadAllLines("config.txt");
            var token = lines[0].ToString().Trim();
            var startDate = lines[1].ToString().Trim();
            var endDate = lines[2].ToString().Trim();
            // lay tong hoa don
            var url = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=1000000&search=tdlap=ge={startDate}T00:00:00;tdlap=le={endDate}T23:59:59";
            
            var danhsachhoadon = new List<HoaDon>();

            var firstTryCount = 0;

            while(firstTryCount < 3)
            {
                try
                {
                    using (var http = new HttpClient())
                    {
                        //http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                        var response = http.GetAsync(url).GetAwaiter().GetResult();

                        if (response.StatusCode != System.Net.HttpStatusCode.OK)
                        {
                            throw new Exception($"Mã lỗi: {response.StatusCode} - {response.Content.ToString()}");
                        }
                        var result = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                        var datas = result.datas;

                        danhsachhoadon = JsonConvert.DeserializeObject<List<HoaDon>>(JsonConvert.SerializeObject(datas));
                        Console.WriteLine($"Co tat ca {danhsachhoadon.Count} hoa don");
                        break;
                    }
                }
                catch
                {
                    firstTryCount++;
                }
            }

            if(firstTryCount >= 3)
            {
                Console.WriteLine("Xin lay lai token");
                Console.ReadKey();
                return;
            }
            

            Console.WriteLine("Dang lay chi tiet hoa don ...");
            int index = 0;
            foreach(var hoadon in danhsachhoadon)
            {
                index++;
                Console.WriteLine($"Dang lay chi tiet hoa don ... {index} / {danhsachhoadon.Count}");
                var detailUrl = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={hoadon.nbmst}&khhdon={hoadon.khhdon}&shdon={hoadon.shdon}&khmshdon=1";
                var secondTryCount = 0;
                while (secondTryCount < 5)
                {
                    try
                    {
                        using (var http = new HttpClient())
                        {
                            //http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                            http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                            http.Timeout = TimeSpan.FromSeconds(5);
                            var response = http.GetAsync(detailUrl).GetAwaiter().GetResult();

                            if (response.StatusCode != System.Net.HttpStatusCode.OK)
                            {
                                throw new Exception($"Mã lỗi: {response.StatusCode} - {response.Content.ToString()}");
                            }
                            var result = JsonConvert.DeserializeObject<dynamic>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());

                            hoadon.TSTSauThue = float.Parse(result?.tgtttbso.ToString());
                            hoadon.TongVAT = float.Parse(result?.tgtthue.ToString());
                            hoadon.LoaiTien = result?.dvtte;
                            hoadon.TiGia = result?.tgia;
                            hoadon.TenHangHoa = result?.hdhhdvu[0]?.ten;
                            hoadon.DonViTinh = result?.hdhhdvu[0]?.dvtinh;
                            hoadon.SoLuong = result?.hdhhdvu[0]?.sluong;
                            hoadon.DonGia = result?.hdhhdvu[0]?.dgia;
                            hoadon.TiLeCK = result?.hdhhdvu[0]?.tlckhau;
                            hoadon.ThanhTien = result?.hdhhdvu[0]?.thtien;
                            hoadon.ThueSuat = result?.hdhhdvu[0]?.tsuat;
                            hoadon.TienThueVAT = hoadon.TongVAT;
                            break;
                        }
                    }
                    catch
                    {
                        secondTryCount++;
                    }
                }

                if (secondTryCount >= 5)
                {
                    Console.WriteLine("Xin lay lai token");
                    Console.ReadKey();
                    return;
                }
            }

            var workbook = new XLWorkbook();
            IXLWorksheet workSheet = workbook.Worksheets.Add($"Hoa Don Dien Tu");
            workSheet.Cell(1, 1).Value = "Tên người mua";
            workSheet.Cell(1, 2).Value = "MST người mua";
            workSheet.Cell(1, 3).Value = "Địa chỉ người mua";
            workSheet.Cell(1, 4).Value = "Số HĐ";
            workSheet.Cell(1, 5).Value = "Ngày HĐ";
            workSheet.Cell(1, 6).Value = "Tên người bán";
            workSheet.Cell(1, 7).Value = "MST Người bán";
            workSheet.Cell(1, 8).Value = "Địa chỉ người bán";
            workSheet.Cell(1, 9).Value = "Tổng số tiền sau thuế";
            workSheet.Cell(1, 10).Value = "Tổng VAT";
            workSheet.Cell(1, 11).Value = "Loại tiền";
            workSheet.Cell(1, 12).Value = "Tỷ giá";
            workSheet.Cell(1, 13).Value = "STT";
            workSheet.Cell(1, 14).Value = "Tên hàng hóa dịch vụ";
            workSheet.Cell(1, 15).Value = "Đơn vị tính";
            workSheet.Cell(1, 16).Value = "Số lượng";
            workSheet.Cell(1, 17).Value = "Đơn giá";
            workSheet.Cell(1, 18).Value = "Tỷ lệ CK";
            workSheet.Cell(1, 19).Value = "Số tiền Ck";
            workSheet.Cell(1, 20).Value = "Thành tiền";
            workSheet.Cell(1, 21).Value = "Thuế Suất";
            workSheet.Cell(1, 22).Value = "Tiền thuế VAT";
            workSheet.Cell(1, 23).Value = "Check Unique";

            for (int i = 0; i < danhsachhoadon.Count; i++)
            {
                var row = danhsachhoadon[i];
                workSheet.Cell(2 + i, 1).Value = row.nmten;
                workSheet.Cell(2 + i, 2).Value = row.nmmst;
                workSheet.Cell(2 + i, 3).Value = row.nmdchi;
                workSheet.Cell(2 + i, 4).Value = row.shdon;
                workSheet.Cell(2 + i, 5).Value = row.ncma.ToString();
                workSheet.Cell(2 + i, 6).Value = row.nbten;
                workSheet.Cell(2 + i, 7).Value = row.nbmst;
                workSheet.Cell(2 + i, 8).Value = row.nbdchi;
                workSheet.Cell(2 + i, 9).Value = row.TSTSauThue.ToString();
                workSheet.Cell(2 + i, 10).Value = row.TongVAT.ToString();
                workSheet.Cell(2 + i, 11).Value = row.LoaiTien;
                workSheet.Cell(2 + i, 12).Value = row.TiGia.ToString();
                workSheet.Cell(2 + i, 13).Value = (i + 1).ToString();
                workSheet.Cell(2 + i, 14).Value = row.TenHangHoa;
                workSheet.Cell(2 + i, 15).Value = row.DonViTinh;
                workSheet.Cell(2 + i, 16).Value = row.SoLuong.ToString();
                workSheet.Cell(2 + i, 17).Value = row.DonGia.ToString();
                workSheet.Cell(2 + i, 18).Value = row.TiLeCK == null ? "-" : row.TiLeCK.ToString();
                workSheet.Cell(2 + i, 19).Value = row.SoTienCK.ToString();
                workSheet.Cell(2 + i, 20).Value = row.ThanhTien.ToString();
                workSheet.Cell(2 + i, 21).Value = row.ThueSuat.ToString();
                workSheet.Cell(2 + i, 22).Value = row.TienThueVAT.ToString();
                workSheet.Cell(2 + i, 23).Value = "";
            }

            var filePath = Path.Combine(startDate.Replace("/", "-") + " đến " + endDate.Replace("/", "-") + "_" + Guid.NewGuid().ToString() + ".xlsx");
            using (var stream = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                workbook.SaveAs(stream);
            }

            Thread.Sleep(100);
        }
    }
}
