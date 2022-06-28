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
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            // read file
            var lines = File.ReadAllLines("config.txt");
            var token = lines[0].ToString().Trim();
            var startDate = lines[1].ToString().Trim();
            var endDate = lines[2].ToString().Trim();
            // lay tong hoa don
            var urlban = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=1000000&search=tdlap=ge={startDate}T00:00:00;tdlap=le={endDate}T23:59:59";
            var urlmua = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/purchase?sort=tdlap:desc,khmshdon:asc,shdon:desc&size=1000000&search=tdlap=ge={startDate}T00:00:00;tdlap=le={endDate}T23:59:59";

            var danhsachhoadonban = new List<HoaDon>();
            var danhsachhoadonmua = new List<HoaDon>();
            Console.WriteLine($"Đang lấy danh sách hóa đơn bán và mua ...");

            var firstTryCount = 0;

            while(firstTryCount < 3)
            {
                try
                {
                    using (var http = new HttpClient())
                    {
                        //http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                        // lay danh sach ban
                        var responseban = http.GetAsync(urlban).GetAwaiter().GetResult();
                        if (responseban.StatusCode != System.Net.HttpStatusCode.OK)
                        {
                            throw new Exception($"Mã lỗi: {responseban.StatusCode} - {responseban.Content.ToString()}");
                        }
                        var resultban = JsonConvert.DeserializeObject<dynamic>(responseban.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                        var datasban = resultban.datas;
                        danhsachhoadonban = JsonConvert.DeserializeObject<List<HoaDon>>(JsonConvert.SerializeObject(datasban));

                        // lay danh sach mua
                        var responsemua = http.GetAsync(urlmua).GetAwaiter().GetResult();
                        if (responsemua.StatusCode != System.Net.HttpStatusCode.OK)
                        {
                            throw new Exception($"Mã lỗi: {responsemua.StatusCode} - {responsemua.Content.ToString()}");
                        }
                        var resultmua = JsonConvert.DeserializeObject<dynamic>(responsemua.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                        var datasmua = resultmua.datas;
                        danhsachhoadonmua = JsonConvert.DeserializeObject<List<HoaDon>>(JsonConvert.SerializeObject(datasmua));


                        Console.WriteLine($"Có tất cả {danhsachhoadonban.Count} hóa đơn bán và {danhsachhoadonmua.Count} hóa đơn mua!");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    firstTryCount++;
                }
            }

            if (firstTryCount >= 3)
            {
                Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                Console.ReadKey();
                return;
            }
            

            Console.WriteLine("Đang lấy thông tin chi tiết hóa đơn ...");
            int indexban = 0;
            foreach(var hoadon in danhsachhoadonban)
            {
                indexban++;
                Console.WriteLine($"Đang lấy thông tin chi tiết hóa đơn bán ... {indexban} / {danhsachhoadonban.Count}");
                var detailUrl = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={hoadon.nbmst}&khhdon={hoadon.khhdon}&shdon={hoadon.shdon}&khmshdon=1";
                var secondTryCount = 0;
                while (secondTryCount < 10)
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

                if (secondTryCount >= 10)
                {
                    Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                    Console.ReadKey();
                    return;
                }
            }


            int indexmua = 0;
            foreach (var hoadon in danhsachhoadonmua)
            {
                indexmua++;
                Console.WriteLine($"Đang lấy thông tin chi tiết hóa đơn mua ... {indexmua} / {danhsachhoadonmua.Count}");
                var detailUrl = $"https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail?nbmst={hoadon.nbmst}&khhdon={hoadon.khhdon}&shdon={hoadon.shdon}&khmshdon=1";
                var secondTryCount = 0;
                while (secondTryCount < 10)
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
                    catch (Exception ex)
                    {
                        secondTryCount++;
                    }
                }

                if (secondTryCount >= 10)
                {
                    Console.WriteLine("Lỗi!!! Xin lấy lại token và chạy lại!");
                    Console.ReadKey();
                    return;
                }
            }

            Console.WriteLine($"Lấy thông tin thành công => in ra file ...");
            // ban
            var workbook = new XLWorkbook();
            IXLWorksheet workSheetBan = workbook.Worksheets.Add($"Hóa Đơn Bán");
            workSheetBan.Cell(1, 1).Value = "Tên người mua";
            workSheetBan.Cell(1, 2).Value = "MST người mua";
            workSheetBan.Cell(1, 3).Value = "Địa chỉ người mua";
            workSheetBan.Cell(1, 4).Value = "Số HĐ";
            workSheetBan.Cell(1, 5).Value = "Ngày HĐ";
            workSheetBan.Cell(1, 6).Value = "Tên người bán";
            workSheetBan.Cell(1, 7).Value = "MST Người bán";
            workSheetBan.Cell(1, 8).Value = "Địa chỉ người bán";
            workSheetBan.Cell(1, 9).Value = "Tổng số tiền sau thuế";
            workSheetBan.Cell(1, 10).Value = "Tổng VAT";
            workSheetBan.Cell(1, 11).Value = "Loại tiền";
            workSheetBan.Cell(1, 12).Value = "Tỷ giá";
            workSheetBan.Cell(1, 13).Value = "STT";
            workSheetBan.Cell(1, 14).Value = "Tên hàng hóa dịch vụ";
            workSheetBan.Cell(1, 15).Value = "Đơn vị tính";
            workSheetBan.Cell(1, 16).Value = "Số lượng";
            workSheetBan.Cell(1, 17).Value = "Đơn giá";
            workSheetBan.Cell(1, 18).Value = "Tỷ lệ CK";
            workSheetBan.Cell(1, 19).Value = "Số tiền Ck";
            workSheetBan.Cell(1, 20).Value = "Thành tiền";
            workSheetBan.Cell(1, 21).Value = "Thuế Suất";
            workSheetBan.Cell(1, 22).Value = "Tiền thuế VAT";
            workSheetBan.Cell(1, 23).Value = "Check Unique";

            for (int i = 0; i < danhsachhoadonban.Count; i++)
            {
                var row = danhsachhoadonban[i];
                workSheetBan.Cell(2 + i, 1).Value = row.nmten;
                workSheetBan.Cell(2 + i, 2).Value = row.nmmst;
                workSheetBan.Cell(2 + i, 3).Value = row.nmdchi;
                workSheetBan.Cell(2 + i, 4).Value = row.shdon;
                workSheetBan.Cell(2 + i, 5).Value = row.nky.ToString();
                workSheetBan.Cell(2 + i, 6).Value = row.nbten;
                workSheetBan.Cell(2 + i, 7).Value = row.nbmst;
                workSheetBan.Cell(2 + i, 8).Value = row.nbdchi;
                workSheetBan.Cell(2 + i, 9).Value = row.TSTSauThue.ToString();
                workSheetBan.Cell(2 + i, 10).Value = row.TongVAT.ToString();
                workSheetBan.Cell(2 + i, 11).Value = row.LoaiTien;
                workSheetBan.Cell(2 + i, 12).Value = row.TiGia.ToString();
                workSheetBan.Cell(2 + i, 13).Value = (i + 1).ToString();
                workSheetBan.Cell(2 + i, 14).Value = row.TenHangHoa;
                workSheetBan.Cell(2 + i, 15).Value = row.DonViTinh;
                workSheetBan.Cell(2 + i, 16).Value = row.SoLuong.ToString();
                workSheetBan.Cell(2 + i, 17).Value = row.DonGia.ToString();
                workSheetBan.Cell(2 + i, 18).Value = row.TiLeCK == null ? "-" : row.TiLeCK.ToString();
                workSheetBan.Cell(2 + i, 19).Value = row.SoTienCK.ToString();
                workSheetBan.Cell(2 + i, 20).Value = row.ThanhTien.ToString();
                workSheetBan.Cell(2 + i, 21).Value = row.ThueSuat.ToString();
                workSheetBan.Cell(2 + i, 22).Value = row.TienThueVAT.ToString();
                workSheetBan.Cell(2 + i, 23).Value = "";
            }


            // mua
            IXLWorksheet workbookMua = workbook.Worksheets.Add($"Hóa Đơn Mua");
            workbookMua.Cell(1, 1).Value = "Tên người mua";
            workbookMua.Cell(1, 2).Value = "MST người mua";
            workbookMua.Cell(1, 3).Value = "Địa chỉ người mua";
            workbookMua.Cell(1, 4).Value = "Số HĐ";
            workbookMua.Cell(1, 5).Value = "Ngày HĐ";
            workbookMua.Cell(1, 6).Value = "Tên người bán";
            workbookMua.Cell(1, 7).Value = "MST Người bán";
            workbookMua.Cell(1, 8).Value = "Địa chỉ người bán";
            workbookMua.Cell(1, 9).Value = "Tổng số tiền sau thuế";
            workbookMua.Cell(1, 10).Value = "Tổng VAT";
            workbookMua.Cell(1, 11).Value = "Loại tiền";
            workbookMua.Cell(1, 12).Value = "Tỷ giá";
            workbookMua.Cell(1, 13).Value = "STT";
            workbookMua.Cell(1, 14).Value = "Tên hàng hóa dịch vụ";
            workbookMua.Cell(1, 15).Value = "Đơn vị tính";
            workbookMua.Cell(1, 16).Value = "Số lượng";
            workbookMua.Cell(1, 17).Value = "Đơn giá";
            workbookMua.Cell(1, 18).Value = "Tỷ lệ CK";
            workbookMua.Cell(1, 19).Value = "Số tiền Ck";
            workbookMua.Cell(1, 20).Value = "Thành tiền";
            workbookMua.Cell(1, 21).Value = "Thuế Suất";
            workbookMua.Cell(1, 22).Value = "Tiền thuế VAT";
            workbookMua.Cell(1, 23).Value = "Check Unique";

            for (int i = 0; i < danhsachhoadonmua.Count; i++)
            {
                var row = danhsachhoadonmua[i];
                workbookMua.Cell(2 + i, 1).Value = row.nmten;
                workbookMua.Cell(2 + i, 2).Value = row.nmmst;
                workbookMua.Cell(2 + i, 3).Value = row.nmdchi;
                workbookMua.Cell(2 + i, 4).Value = row.shdon;
                workbookMua.Cell(2 + i, 5).Value = row.nky.ToString();
                workbookMua.Cell(2 + i, 6).Value = row.nbten;
                workbookMua.Cell(2 + i, 7).Value = row.nbmst;
                workbookMua.Cell(2 + i, 8).Value = row.nbdchi;
                workbookMua.Cell(2 + i, 9).Value = row.TSTSauThue.ToString();
                workbookMua.Cell(2 + i, 10).Value = row.TongVAT.ToString();
                workbookMua.Cell(2 + i, 11).Value = row.LoaiTien;
                workbookMua.Cell(2 + i, 12).Value = row.TiGia.ToString();
                workbookMua.Cell(2 + i, 13).Value = (i + 1).ToString();
                workbookMua.Cell(2 + i, 14).Value = row.TenHangHoa;
                workbookMua.Cell(2 + i, 15).Value = row.DonViTinh;
                workbookMua.Cell(2 + i, 16).Value = row.SoLuong.ToString();
                workbookMua.Cell(2 + i, 17).Value = row.DonGia.ToString();
                workbookMua.Cell(2 + i, 18).Value = row.TiLeCK == null ? "-" : row.TiLeCK.ToString();
                workbookMua.Cell(2 + i, 19).Value = row.SoTienCK.ToString();
                workbookMua.Cell(2 + i, 20).Value = row.ThanhTien.ToString();
                workbookMua.Cell(2 + i, 21).Value = row.ThueSuat.ToString();
                workbookMua.Cell(2 + i, 22).Value = row.TienThueVAT.ToString();
                workbookMua.Cell(2 + i, 23).Value = "";
            }

            var filePath = Path.Combine(startDate.Replace("/", "-") + " đến " + endDate.Replace("/", "-") + "_" + Guid.NewGuid().ToString() + ".xlsx");
            using (var stream = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                workbook.SaveAs(stream);
            }

            Console.WriteLine($"DONE!!!");
            Console.ReadKey();
        }
    }
}
