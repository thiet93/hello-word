using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iVBDH.Utils;
using iVBDH.iWeb35.Base;
using iVBDH.iWeb35.ivbdh.qtht.Phongbans;
using Microsoft.SharePoint;
using iVBDH.iWeb35.ivbdh.Vanban.vbdbutphe;
using iVBDH.iWeb35.ivbdh.Vanban.vbdsentinfo;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using iVBDH.iWeb35.ivbdh.Danhmuc.coquanbanhanh;
using System.IO;
using iVBDH.iWeb35.ivbdh.Vanban.vbdgopy;
using iVBDH.iWeb35.ivbdh.Hoso.hosocv;
using iVBDH.iWeb35.ivbdh.qtht.Logs;
using iVBDH.iWeb35.ivbdh.Vanban.vanbandientu;
using System.Text.RegularExpressions;
using iVBDH.iWeb35.ivbdh.Danhmuc.loaivanban;
using iVBDH.iWeb35.ivbdh.Danhmuc.tinhchatvanban;
using iVBDH.iWeb35.ivbdh.Danhmuc.nguoiky;
using iVBDH.iWeb35.ivbdh.Danhmuc.ChucVu;
using Microsoft.Office.Server.Search.Query;
using System.Data;
using iVBDH.iWeb35.ivbdh.qtht.Users;
using Microsoft.Sharepoint.SandIndex.ivbdh.qtht.Config;
using Ionic.Zip;
using iVBDH.iWeb35.ivbdh.Vanban.Luongxuly;
using System;
using TUtils;
using iVBDH.iWeb35.ivbdh.qtht.Config;
using iVBDH.iWeb35.ivbdh.Base;
using iVBDH.iWeb35.ivbdh.Vanban.vbdisentinfo;
using iVBDH.iWeb35.ivbdh.Vanban.vanbandi;
using System.Globalization;
using iVBDH.iWeb35.ivbdh.Danhmuc.socongvan;



namespace iVBDH.iWeb35.ivbdh.Vanban.vanbanden
{
    public partial class action : pActionBase
    {
        // khoi tao doi tuong
        public LVanBanDen vbdItem { get; set; }
        public LVanBanDenDA vbdDA { get; set; }

        public LThongTinGuiVanBanDenDA sentInforDA { get; set; }
        public LThongTinGuiVanBanDen sentInforItem { get; set; }

        public LButPheVanBanDen butpheItem { get; set; }
        public LButPheVanBanDenDA butpheDA { get; set; }
        public LGroupDA groupDA { get; set; }
        // List object khoi phat hanh - Phuc vu render Json
        public List<LVanBanDenJson> jsonLVANBANDEN;

        /// <summary>
        /// Giá trị value
        /// </summary>
        public LVanBanDenQuery SearchVanBan { get; set; }

        public action()
        {

        }
        #region code fulltext cũ
        //public LVanBanDenFSQuery FullSearchVanBan { get; set; }
        //public void get_request_fs()
        //{
        //    FullSearchVanBan = new LVanBanDenFSQuery();
        //    FullSearchVanBan.DonVis = currentUser.CurrentTenDonVi;
        //    FullSearchVanBan.CurrentUserData = new LookupData(currentUser.ID, currentUser.userTenTruyCap);
        //    FullSearchVanBan.IsXemTatCaDonVi = currentUser.ListPermission.XemVanBanDenCuaDonVi;
        //    /// xem tat ca van ban den cua don vi
        //    if (currentUser.ListPermission.XemTatCaVanBanDen)
        //    {
        //        FullSearchVanBan.TrangThaiVanBan = 1;
        //        FullSearchVanBan.isXemTatca = true;

        //        // bo sung 25/03
        //        FullSearchVanBan.NguoiTao = currentUser.userTenTruyCap;
        //        FullSearchVanBan.DSTGXuLy = currentUser.userTenTruyCap;

        //        /// tất cả văn bản đều phải duyệt mới xem được
        //        FullSearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt

        //    }
        //    else if (currentUser.ListPermission.XemTatCaVanBanDenDaDuyet)
        //    {
        //        FullSearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt
        //        FullSearchVanBan.isXemTatca = true;
        //    }
        //    else
        //    {

        //        FullSearchVanBan.NguoiTao = currentUser.userTenTruyCap;
        //        FullSearchVanBan.DSTGXuLy = currentUser.userTenTruyCap;

        //        /// tất cả văn bản đều phải duyệt mới xem được
        //        FullSearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt
        //    }

        //}
        #endregion
        /// <summary>
        /// Lấy tham số
        /// </summary>
        private void get_request()
        {
            SearchVanBan = new LVanBanDenQuery(this.Page.Request);
            //1- Là văn thư đơn vị - là văn thư phòng (All đơn vị + id phòng)
            //2- Là văn thư của đơn vị - là chuyên vien phòng (      All đơn vị + id phòng + các văn bản gửi cho cá nhân trong phòng)
            //3- Là văn thư phong - la chuyên viên đơn vị   (các văn bản gửi cho đơn vị + all văn bản phòng)
            //4- Là chuyên viên cả đơn vị lẫn phòng  (chỉ các văn bản được gửi cho cá nhân)

            
            SearchVanBan.IsXemCuaDonVi = currentUser.ListPermission.XemVanBanDenCuaDonVi;
            SearchVanBan.DisplayDonviInfo = currentUser.DisplayDonviInfo;
            SearchVanBan.IDPhongBan = currentUser.CurrentIDPhongBan;

            SearchVanBan.lstIDScvDonVi = currentUser.lstIDScvDonVi;

            // xacs dinh vai tro cua nguoi dung de phuc vu truy van van ban trong phong ban hoac don vi            
            if (currentUser.ListPermission.ThemVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count > 0)
            {
                SearchVanBan.intTrangThaiVanThu = 1; // van thu cua ca don vi va phong ban
            }
            else if (currentUser.ListPermission.ThemVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count == 0)
            {
                SearchVanBan.intTrangThaiVanThu = 2; // Là văn thư của đơn vị - là chuyên vien phòng 
            }
            else if (!currentUser.ListPermission.ThemVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count > 0)
            {
                SearchVanBan.intTrangThaiVanThu = 3; // Là văn thư phong - la chuyên viên đơn vị 
            }
            else if (!currentUser.ListPermission.ThemVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count == 0)
            {
                SearchVanBan.intTrangThaiVanThu = 4; // Là chuyên viên cả đơn vị lẫn phòng
            }


            SearchVanBan.IdPhongBans = currentUser.lstPhongBan;
            SearchVanBan.DonVis = currentUser.CurrentIDPhongBanCap1;
            SearchVanBan.CurrentUserData = new LookupData(currentUser.ID, currentUser.userTenTruyCap);
            if (currentUser.lstPBHasVanBan.Count > 0)
            {
                SearchVanBan.lstPBCoVanBan = currentUser.lstPBHasVanBan; // truyền tham số vào để truy vấn
                SearchVanBan.lstTaiKhoanDDPB = currentUser.lstTaiKhoanDDPB; // truyền thông tin tài khoản đại diện để truy vấn

            }


            /// xem tat ca van ban den cua don vi
            if (currentUser.ListPermission.XemTatCaVanBanDen)
            {
                SearchVanBan.TrangThaiVanBan = 1;
                SearchVanBan.isXemTatca = true;

                // bo sung 25/03
                SearchVanBan.NguoiTaoID = currentUser.ID;
                SearchVanBan.NguoiTao = currentUser.userTenTruyCap;
                SearchVanBan.DSTGXuLy = currentUser.userTenTruyCap;

                /// tất cả văn bản đều phải duyệt mới xem được
                SearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt


                /// truy van danh sach phong ban dai dien theo don vi
                /// edit 0708
                //SPListItemCollection items = groupDA.GetDVPBByTKDaiDien(currentUser.ID, currentUser.CurrentIDDonVi);
                //if (items != null)
                //{
                //    foreach (SPListItem item in items)
                //    {
                //        SearchVanBan.DSPBDVTGXuLy.Add(Convert.ToInt32(item["ID"])); //  lay danh sach phong ban nguoi dung duoc dai dien
                //    }
                //}

                /// bo sung 0708
                SearchVanBan.DSPBDVTGXuLy = currentUser.DSPBDVTGXuLy;
                SearchVanBan.DSPBDVTGXulyTitle = currentUser.DSPBDVTGXuLyTitle;

            }
            else if (currentUser.ListPermission.XemTatCaVanBanDenDaDuyet)
            {
                SearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt
                SearchVanBan.isXemTatca = true;
            }
            else
            {
                // phong ban dai dien
                // nguoi nguoi tao               
                // danh sach tham gia xu ly
                SearchVanBan.NguoiTaoID = currentUser.ID;
                SearchVanBan.NguoiTao = currentUser.userTenTruyCap;
                SearchVanBan.DSTGXuLy = currentUser.userTenTruyCap;

                /// tất cả văn bản đều phải duyệt mới xem được
                SearchVanBan.TrangThaiVanBan = 1; // trạng thái đã duyệt

                /// bo sung 0708
                SearchVanBan.DSPBDVTGXuLy = currentUser.DSPBDVTGXuLy;
                SearchVanBan.DSPBDVTGXulyTitle = currentUser.DSPBDVTGXuLyTitle;
            }

        }

        /// <summary>
        /// phucj vu tim kiem toan van
        /// </summary>
        /// <returns></returns>
        public List<int> FSVanBanDen(string q, string strDonViTitle, string isToanVan)
        {
            try
            {
                List<int> lstID = new List<int>();
                string strTuKhoa = LocChuoi(q);

                FullTextSqlQuery query = new FullTextSqlQuery(SPContext.Current.Site);
                string strSelect = "";
                string queryField = "ID";
                string strSearchDonVi = " and FREETEXT(vbdGroupLookup,'\"" + strDonViTitle + "\"')";

                string strSearchToanVan = "";
                if (!string.IsNullOrEmpty(strTuKhoa))
                {
                    strSearchToanVan = " and CONTAINS(*,'\"" + strTuKhoa + "\"')";
                }

                strSelect = "SELECT " + queryField + " from SCOPE() WHERE \"scope\"='Lvanbanden' " + strSearchToanVan + strSearchDonVi + " ORDER BY Rank DESC";


                query.QueryText = strSelect;
                query.ResultTypes = ResultType.RelevantResults;
                //query.StartRow = (currentPage - 1) * rowLimit;
                //query.RowLimit = rowLimit;
                //query.TotalRowsExactMinimum = 50 * rowLimit;
                query.TrimDuplicates = true;
                query.EnableStemming = false;
                query.IgnoreAllNoiseQuery = true;
                query.HighlightedSentenceCount = 3;

                ResultTableCollection results = query.Execute();

                if (results.Count > 0)
                {
                    ResultTable rsTable = results[ResultType.RelevantResults];
                    DataSet ds = GetResultData(rsTable);

                    foreach (DataRow spItem in ds.Tables["SearchResultDataSet"].Rows)
                    {
                        lstID.Add(Convert.ToInt32(spItem["ID"].ToString()));
                    }
                }

                return lstID;
            }
            catch (Exception ex)
            {
                return new List<int>();
            }
        }
        public List<LVanBanDenJson> GetTHongTinXL(List<LVanBanDenJson> lsttemp, string TaiKhoanChuyenVB, int idThongKe)
        {
            LThongTinGuiVanBanDenDA oThongTinGuiVBDA = new LThongTinGuiVanBanDenDA();
            var oThongTinGui = oThongTinGuiVBDA.LayThongTinTheoDanhSachVanBan(idThongKe, TaiKhoanChuyenVB, lsttemp.Select(item => item.ID).ToArray());
            oThongTinGui = oThongTinGui.Select(item => { item.infoUserNameReceived = (item.infoUserNameReceived != null && item.infoUserNameReceived.LookupId > 0) ? item.infoUserNameReceived : item.infoGroupNameReceived; return item; }).GroupBy(item => item.infoVanBanDenID.LookupId).
                Select(item => new LThongTinGuiVanBanDen() { ID = item.Key, Title = string.Join(",", item.Select(xl => xl.infoUserNameReceived.LookupValue).Distinct()) }).ToList();
            lsttemp = lsttemp.GroupJoin(oThongTinGui, vb => vb.ID, tt => tt.ID, (vb, tt) => new { vb = vb, tt = tt }).SelectMany(temp => temp.tt.DefaultIfEmpty(),
                (temp, tt) => { temp.vb.strTitles = (tt == null ? string.Empty : tt.Title); return temp.vb; }).ToList();
            return lsttemp;
        }
        /// <summary>
        /// ham load trang
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            vbdDA = new LVanBanDenDA();
            groupDA = new LGroupDA();
            butpheDA = new LButPheVanBanDenDA();
            sentInforDA = new LThongTinGuiVanBanDenDA();
            ///nếu tồn tại cả action và array ID
            ///
            try
            {
                if (!string.IsNullOrEmpty(DoAction) && DoAction.ToLower() == "json")
                {
                    // truy xuất tham số cho văn bản query
                    get_request();
                    JsonGrid data = new JsonGrid();
                    if (SearchVanBan.XLVBLanhDao)
                    {
                        LGroupDA oGroupDA = new LGroupDA();
                        var listld = oGroupDA.GetById(90).groupUsers;
                        if (SearchVanBan.XemTatCaVanBanLanhDao)
                        {
                            foreach (SPFieldLookupValue spItem in listld)
                            {
                                SearchVanBan.vbdDSTGXuLyText.Add(string.Format(";#{0};#", spItem.LookupId));
                            }
                        }
                        else
                        {
                            foreach (SPFieldLookupValue spItem in listld)
                            {

                                SearchVanBan.vbdDSChuaXuLyText.Add(string.Format(";#{0};#", spItem.LookupId));
                            }
                        }

                    }
                    if (Request["ld"] != null) // theo doi ho so cong viec cua lanh đạo
                    {
                        /// phuc vu tim kiem toan van tim kiem toan van vbdToanVan
                        if (!string.IsNullOrEmpty(Request["vbdToanVan"]))
                        {
                            SearchVanBan.lstIDVanBans = FSVanBanDen(Request["vbdToanVan"].ToString(), currentUser.CurrentTenDonVi, "1");
                        }

                        data.Data = vbdDA.GetJsonVanBanDenByLanhDao(SearchVanBan, GridRequest.pageSize, GridRequest.page, Field, FieldOption);
                    }
                    else
                    {
                        /// phuc vu tim kiem toan van tim kiem toan van vbdToanVan
                        if (!string.IsNullOrEmpty(Request["vbdToanVan"]))
                        {
                            SearchVanBan.lstIDVanBans = FSVanBanDen(Request["vbdToanVan"].ToString(), currentUser.CurrentTenDonVi, "1");
                        }
                        var lsttemp = vbdDA.GetJsonVanBanDenByQuery(SearchVanBan, GridRequest.pageSize, GridRequest.page, Field, FieldOption);
                        //lsttemp = GetTHongTinXL(lsttemp, taikhoan, 0);
                        data.Data = lsttemp;
                    }
                    data.Request = GridRequest;
                    data.Total = vbdDA.TongSoBanGhiSauKhiQuery;
                    RenderJson(data);
                }
                else if (!string.IsNullOrEmpty(DoAction) && DoAction.ToLower() == "jsonall")
                {
                    //GetJsonVanBanDenByQuery2
                    SearchVanBan = new LVanBanDenQuery();
                    JsonGrid data = new JsonGrid();
                    int year = 0;
                    if (!string.IsNullOrEmpty(Request["year"]))
                        year = Convert.ToInt32(Request["year"]);
                    else
                        year = DateTime.Now.Year;
                    SearchVanBan.Year = year;
                    /*Nếu là thống kê của phòng Tổng hợp: thì lấy danh sách tài khoản là Lãnh đạo bộ, 
                     * sau đó tìm kiếm những văn bản mà trường vbđSTGXL có tài khoản lãnh đạo bộ tham gia*/
                    List<vanbanden.LVanBanDenJson> lsttemp = new List<LVanBanDenJson>();
                    int idThongKe = currentUser.CurrentIDPhongBan;
                    LConfigDA config = new LConfigDA();
                    SPFieldLookupValue TaiKhoanChuyenVB = new SPFieldLookupValue();
                    SPListItem spItemConfig = null;
                    SPListItem itemtkchuyenvb = config.GetConfigByType("taikhoanchuyenvb");
                    if (itemtkchuyenvb != null && itemtkchuyenvb["configValue"] != null)
                    {
                        TaiKhoanChuyenVB = new SPFieldLookupValue(Convert.ToString(itemtkchuyenvb["configValue"]));
                    }
                    // butphe != null = báo cáo phòng tổng hợp
                    if (Request["butphe"] != null)
                    {
                        //nếu lấy bút phê thì bổ sung thêm thông tin bút phê
                        //List<int> lstvbid = jsonLVANBANDEN.Select(x => x.ID).ToList();
                        //lấy danh sách tài khoản trong group lãnh đạo bộ
                        SPListItem spConfigLDB = config.GetConfigByType("chonphongldb");
                        List<int> idsLD = new List<int>();
                        if (spConfigLDB != null && spConfigLDB["configValue"] != null)
                        {
                            int idGroup = Convert.ToInt32(spConfigLDB["configValue"]);
                            LGroupDA oGroupDA = new LGroupDA();
                            LGroup oGroupLD = oGroupDA.GetById(idGroup);
                            //Xóa tài khoản bộ khoa công nghệ khỏi danh sách tài khoản lãnh đạo bộ
                            int idTK = TConvert.ToInt(TaiKhoanChuyenVB.LookupValue);
                            SearchVanBan.vbdDSTGXuLyText = oGroupLD.groupUsers.Where(item => item.LookupId != idTK).Select(item => string.Format(";#{0};#", item.LookupId)).ToList();
                            idsLD = oGroupLD.groupUsers.Where(item => item.LookupId != idTK).Select(item => item.LookupId).ToList();
                        }
                        lsttemp = vbdDA.GetJsonVanBanDenByQuery2(SearchVanBan, GridRequest.pageSize, GridRequest.page, Field, FieldOption);
                        lsttemp = butpheDA.GetListButPheByListVanBanId(lsttemp);
                        lsttemp = lsttemp.Select(item => { item.UpdateDSTGXuLy(idsLD); return item; }).ToList();
                        //spItemConfig = config.GetConfigByType("chonphongtonghop");
                        // reset tài khoản chuyển văn bản để phục vụ cho query thông tin gửi nhận văn bản của phòng tổng hợp
                        TaiKhoanChuyenVB = new SPFieldLookupValue();
                        int start = (GridRequest.page == 1) ? 0 : ((GridRequest.page - 1) * Convert.ToInt32(GridRequest.pageSize)); ;

                        for (int i = 0; i < lsttemp.Count; i++)
                        {
                            lsttemp[i].STT = start + i + 1;
                        }
                    }
                    else
                    {
                        lsttemp = vbdDA.GetJsonVanBanDenByQuery2(SearchVanBan, GridRequest.pageSize, GridRequest.page, Field, FieldOption);
                        spItemConfig = config.GetConfigByType("chonphonghanhchinh");
                        if (spItemConfig != null && spItemConfig["configValue"] != null)
                        {
                            idThongKe = Convert.ToInt32(spItemConfig["configValue"]);
                        }
                        //Lấy danh sách thông tin gửi nhận
                        lsttemp = GetTHongTinXL(lsttemp, TaiKhoanChuyenVB.LookupValue, idThongKe);
                    }
                    
                    data.Data = lsttemp;
                    data.Request = GridRequest;
                    data.Total = vbdDA.TongSoBanGhiSauKhiQuery;
                    RenderJson(data);
                }
                else
                {
                    if (!string.IsNullOrEmpty(DoAction))
                    {

                        switch (DoAction.ToLower())
                        {
                            case "delete":
                                _objMesage = Delete(ltsID);
                                break;
                            case "searchindex":
                                _objMesage = SearchIndex();
                                break;
                            case "add":
                                _objMesage = Add();
                                break;
                            case "addemail":
                                _objMesage = AddFromEmail();
                                break;

                            case "edit":
                                _objMesage = Edit();
                                break;
                            case "done":
                                _objMesage = Done();
                                break;
                            case "know":
                                _objMesage = Know();
                                break;
                            case "hxl":
                                _objMesage = CapNhatHXL();
                                break;
                            case "phanloai":
                                _objMesage = PhanLoaiVanBan();
                                break;

                            case "unfolow":
                                UnFolow();
                                break;
                            case "checkthuhoi":
                                _objMesage = CheckThuHoiVB();
                                break;
                            case "thuhoivbwithusers":
                                _objMesage = ThuHoiVBWithUsers();
                                break;
                            case "thuhoi":
                                _objMesage = ThuHoiVB();
                                break;
                            case "folow":
                                Folow();
                                break;

                            case "attach":
                                _objMesage = Attach();
                                break;

                            case "getnum":
                                _objMesage = GetNum();
                                break;
                            case "check":

                                _objMesage = CheckTrungVB(Request["vbdSoKyHieu"], Request["vbdNgayDen"]);
                                break;

                            case "exportedxml":
                                ExportEdXML();
                                break;
                            case "import":
                                _objMesage = ImportEdXML();
                                break;
                            case "addbutphe":
                                _objMesage = AddButPhe();
                                break;
                            case "thuhoivanbandi":
                                _objMesage = thuhoivanbandi();
                                break;
                            case "checkvbmang":
                                _objMesage = CheckExitVanBanQuaMang();
                                break;
                            case "sosanh":
                                _objMesage = SoSanhVanBanNhapVaQuaMang();
                                break;
                            case "vaosodonvi":
                                _objMesage = VaoSoVBDonVi();
                                break;

                        }
                    }
                    RenderMessage(_objMesage);
                }
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.Message;
            }
        }


        /// <summary>
        /// vào sổ văn bản cho văn bản của đơn vị
        /// </summary>
        /// <returns></returns>
        private Utils.objMessage VaoSoVBDonVi()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            LSoCongVanDA scvDA = new LSoCongVanDA();
            LVanBanDen vbdItem = vbdDA.GetById(ItemID);
            try
            {
                int scvId = Convert.ToInt32(Request["vbdSoCongVanLookup"].ToString());
                int maxNum = Convert.ToInt32(Request["vbdSoVanBan"].ToString());

                DateTime? ngayden;

                DateTimeFormatInfo dtfiParser;
                dtfiParser = new DateTimeFormatInfo();
                dtfiParser.ShortDatePattern = "dd/MM/yyyy";
                dtfiParser.DateSeparator = "/";

                if (!string.IsNullOrEmpty(Request["vbdNgayDen"]))
                    ngayden = Convert.ToDateTime(Request["vbdNgayDen"], dtfiParser);
                else
                    ngayden = null;


                scvDA.UpdateSCVMAXNum(scvId, maxNum + 1);// update số lớn nhất hiện tại
               
                // Cập nhật sổ và số văn bản mới vào item văn bản đến
                // vbdMultiSCV
                // strSoVBPhongBan ==> ID;#Num;#ID;#num
                vbdDA.UpdateSCVDonVi(vbdItem, scvId, maxNum, ngayden);


                _objMesage.Erros = false;
                _objMesage.Message = "Đã vào sổ cho sổ văn bản.";
                _objMesage.ID = "0";



            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        private objMessage SoSanhVanBanNhapVaQuaMang()
        {
            Utils.objMessage objMsg = new Utils.objMessage();
            if (!string.IsNullOrEmpty(Request["ItemIDQM"]))
            {
                //vbdDA
                int ItemIDQM = Convert.ToInt32(Request["ItemIDQM"]);
                //xu ly file attach truoc
                if (!string.IsNullOrEmpty(Request["ListFileAttach"]))
                {
                    List<int> lstid = SPUtils.GetDanhSachIDsQuaFormPost(Request["ListFileAttach"]);
                    if (lstid.Count > 1)//gộp 2 file của 2 văn bản 
                    {
                        #region
                        SPListItem vbdenqm = vbdDA.SpListProcess.GetItemByIdSelectedFields(ItemIDQM, "Attachments");
                        SPListItem vbdennhap = vbdDA.SpListProcess.GetItemByIdSelectedFields(ItemID, "Attachments");
                        foreach (string fileName in vbdenqm.Attachments)
                        {
                            SPFile file = vbdenqm.ParentList.ParentWeb.GetFile(vbdenqm.Attachments.UrlPrefix + fileName);
                            byte[] imageData = file.OpenBinary();
                            vbdennhap.Attachments.Add(fileName, imageData);
                        }
                        vbdennhap.SystemUpdate();
                        #endregion

                    }
                    else if (lstid.Count == 1) //chỉ lấy 1 file
                    {
                        #region code
                        if (lstid[0] == ItemIDQM)
                        {
                            SPListItem vbdenqm = vbdDA.SpListProcess.GetItemByIdSelectedFields(ItemIDQM, "Attachments");
                            SPListItem vbdennhap = vbdDA.SpListProcess.GetItemByIdSelectedFields(ItemID, "Attachments");
                            SPAttachmentCollection tempfile = vbdennhap.Attachments;
                            foreach (string fileName in tempfile)
                            {
                                vbdennhap.Attachments.DeleteNow(fileName);
                            }
                            foreach (string fileName in vbdenqm.Attachments)
                            {
                                SPFile file = vbdenqm.ParentList.ParentWeb.GetFile(vbdenqm.Attachments.UrlPrefix + fileName);
                                byte[] imageData = file.OpenBinary();
                                vbdennhap.Attachments.Add(fileName, imageData);
                            }
                            vbdennhap.SystemUpdate();
                        }
                        #endregion
                    }

                }
                //các trường thông tin khác
                vbdItem = vbdDA.GetById(ItemID);
                vbdItem.ChangeObject(this.Request);
                string strThongBao = vbdDA.Update(vbdItem);
                vbdDA.Delete(ItemIDQM);
                if (string.IsNullOrEmpty(strThongBao))
                {
                    objMsg.Erros = false;
                    objMsg.Message = string.Format("Đã cập nhật {0}: <b>{1}</b>", TitleWebpart, vbdItem.vbdSoKyHieu);
                    objMsg.ID = vbdItem.ID.ToString();
                }
            }
            return objMsg;
        }

        private objMessage CheckExitVanBanQuaMang()
        {
            Utils.objMessage objMsg = new Utils.objMessage();

            string SoKyHieu = Request["vbdSoKyHieu"];
            DateTimeFormatInfo dtfi = new DateTimeFormatInfo();
            dtfi.ShortDatePattern = "dd/MM/yyyy";
            dtfi.DateSeparator = "/";
            DateTime vbdNgayBanHanh = Convert.ToDateTime(Request["vbdNgayBanHanh"], dtfi);

            int intCheckVB = vbdDA.CheckExitsReturnID(SoKyHieu, ItemID, currentUser.CurrentIDDonVi, vbdNgayBanHanh, vbdNgayBanHanh.Year.ToString());
            if (intCheckVB > 0)
            {
                objMsg.Erros = true;
                objMsg.Message = "Văn bản đến đã tồn tại. Bạn có muốn nhập văn bản tiếp tục hay ko?";
                objMsg.ID = Convert.ToString(intCheckVB);
                vbdItem = vbdDA.GetById(intCheckVB);
                objMsg.Message = string.Format("Đã tồn tại văn bản <b> số {0}, do {1} ban hành ngày {2}, về việc: {3}, đến ngày {4} </b>. </br> Bạn có muốn sử lý trên văn bản đã tồn tại không?", vbdItem.vbdSoKyHieu,
                    string.Join(",", vbdItem.vbdCoQuanBanHanh.Select(x => x.LookupValue)), (vbdItem.vbdNgayBanHanh != null ? String.Format("{0:dd/MM/yyyy}", vbdItem.vbdNgayBanHanh) : ""), vbdItem.vbdTrichYeu,
                    (vbdItem.vbdNgayDen != null ? String.Format("{0:dd/MM/yyyy}", vbdItem.vbdNgayDen) : ""));
            }
            else
            {
                objMsg.Erros = false;
                objMsg.Message = "Văn bản đến đã tồn tại. Bạn có muốn nhập văn bản tiếp tục hay ko?";
            }

            return objMsg;
            //throw new NotImplementedException();
        }

        private Utils.objMessage thuhoivanbandi()
        {
            Utils.objMessage objMsg = new Utils.objMessage();
            try
            {
                string list = Request["List"];
                string[] lst = list.Split(new Char[] {','});
                foreach (var item in lst)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        string id = Request["id"];
                        int infoPB = Convert.ToInt32(item);                        
                        LThongTinGuiVanBanDiDA ttgvbdDA= new LThongTinGuiVanBanDiDA();                       
                        int idpb = ttgvbdDA.GetinfoGroupbyID(infoPB);
                        vbdDA = new LVanBanDenDA();
                        int idvbd = vbdDA.GetIdByIdVBDi(id, idpb);
                        
                        if(idvbd > 0)
                        vbdDA.Delete(idvbd);
                        ttgvbdDA.Delete(infoPB);
                    }
                }
                objMsg.Erros = false;
                objMsg.Message = "Thu hồi văn bản thành công";
            }
            catch (Exception)
            {
                objMsg.Erros = true;
                objMsg.Message = "Lỗi thu hồi";
            }
            return objMsg;
        }
        public Utils.objMessage SearchIndex()
        {
            Utils.objMessage _objMsg = new Utils.objMessage();

            LVanBanDenDA vbdDA = new LVanBanDenDA();
            //DateTime ngayDen = TConvert.ToDateTimeFormat(Request["NgayDen"]).Value;
            var SearchVanBan = new LVanBanDenQuery();

            SearchVanBan.Year = Convert.ToInt16(Request["Years"]);
            /*Nếu là thống kê của phòng Tổng hợp: thì lấy danh sách tài khoản là Lãnh đạo bộ, 
             * sau đó tìm kiếm những văn bản mà trường vbđSTGXL có tài khoản lãnh đạo bộ tham gia*/
            List<vanbanden.VanBanDenSolr> lsttemp = new List<VanBanDenSolr>();
            int idThongKe = currentUser.CurrentIDPhongBan;
            SPFieldLookupValue TaiKhoanChuyenVB = new SPFieldLookupValue();
            string itemtkchuyenvb = SysParms.SysConfigs.GetValueConfigByType("taikhoanchuyenvb");
            if (!string.IsNullOrEmpty(itemtkchuyenvb))
            {
                TaiKhoanChuyenVB = new SPFieldLookupValue(itemtkchuyenvb);
            }
            // butphe != null = báo cáo phòng tổng hợp
            if (!string.IsNullOrEmpty(Request["butphe"]))
            {
                //nếu lấy bút phê thì bổ sung thêm thông tin bút phê
                //List<int> lstvbid = jsonLVANBANDEN.Select(x => x.ID).ToList();
                //lấy danh sách tài khoản trong group lãnh đạo bộ
                string spConfigLDB = SysParms.SysConfigs.GetValueConfigByType("chonphongldb");
                List<int> idsLD = new List<int>();
                if (!string.IsNullOrEmpty(spConfigLDB))
                {
                    int idGroup = Convert.ToInt32(spConfigLDB);
                    LGroupDA oGroupDA = new LGroupDA();
                    LGroup oGroupLD = oGroupDA.GetById(idGroup);
                    //Xóa tài khoản bộ khoa công nghệ khỏi danh sách tài khoản lãnh đạo bộ
                    int idTK = TConvert.ToInt(TaiKhoanChuyenVB.LookupValue);
                    SearchVanBan.vbdDSTGXuLyText = oGroupLD.groupUsers.Where(item => item.LookupId != idTK).Select(item => string.Format(";#{0};#", item.LookupId)).ToList();
                    idsLD = oGroupLD.groupUsers.Where(item => item.LookupId != idTK).Select(item => item.LookupId).ToList();
                }
                lsttemp = vbdDA.GetJsonVanBanDenAll(SearchVanBan, "vbdSoVanBan", true);
                //spItemConfig = config.GetConfigByType("chonphongtonghop");
                // reset tài khoản chuyển văn bản để phục vụ cho query thông tin gửi nhận văn bản của phòng tổng hợp
                //TaiKhoanChuyenVB = new SPFieldLookupValue();
            }
            else
            {
                lsttemp = vbdDA.GetJsonVanBanDenAll(SearchVanBan, "vbdSoVanBan", true);

            }
            int soden = TConvert.ToInt(Request["SoDen"]);
            string soKyHieu = Request["SoKyHieu"];
            List<int> lstIndex = new List<int>();
            for (int i = 0; i < lsttemp.Count; i++)
            {
                if (lsttemp[i].vbdSoVanBan == soden)
                {
                    lstIndex.Add(i + 1);
                }
            }
            if (lstIndex.Count > 0)
            {
                StringBuilder sblIndex = new StringBuilder(string.Join(",",lstIndex));
               
                _objMsg.Message = sblIndex.ToString();
            }
            else
                _objMsg.Message = "Không tìm thấy văn bản nào";

            return _objMsg;
        }
        private Utils.objMessage AddButPhe()
        {
            Utils.objMessage objMsg = new Utils.objMessage();
            try
            {
                vbdDA = new LVanBanDenDA();
                int IdNguoiButPhe = currentUser.ID;
                string tenNguoiButPhe = currentUser.Title;
                if (currentUser.ListPermission.XuLyVanBanLanhDao && !string.IsNullOrEmpty(Request["NguoiButPhe"]))
                {
                    IdNguoiButPhe = Convert.ToInt32(Request["NguoiButPhe"]);
                    tenNguoiButPhe = Request["TenNguoiButPhe"];
                }
                butpheItem = new LButPheVanBanDen();
                butpheItem.butpheIDVanBanDen.LookupId = ItemID;
                butpheItem.butpheLanhDao.LookupId = IdNguoiButPhe;
                if (!string.IsNullOrEmpty(Request["imgcrop"]) && Request["imgcrop"] != "undefined")
                {
                    string fileCrop = Request["imgcrop"];
                    string filefullUrl = Server.MapPath("~") + fileCrop;
                    butpheItem.ListFileAttachAdd.Add(new FileAttach(Path.GetFileName(fileCrop), Utils.SPUtils.ReadFile(filefullUrl)));
                }
                butpheItem.butpheNoiDung = Request["vNoiDungBP"].ToString();
                string strThongBao = butpheDA.Add(butpheItem);

                int wfid = 0; // cong việc của người dùng
                if (Request["wfid"] != null && !string.IsNullOrEmpty(Request["wfid"]))
                {
                    wfid = Convert.ToInt32(Request["wfid"]);
                }

                vbdDA.UpdateTrangThaiXuLyDone(ItemID, IdNguoiButPhe, wfid); // cap nhat trang thai xu ly xong cua nguoi dung
                LVanBanDen tem = vbdDA.GetSoKyHieuById(ItemID);
                /// add log: Xử lý xong văn bản
                addLog(EnumVanBanDen.strDone + " số: " + tem.vbdSoKyHieu, EnumVanBanDen.Done, EnumDoiTuong.VanBanDen, ItemID, tem.vbdDSTGXuLy, false, false, false);
                if (Request["b"].ToString().Equals("False"))
                { /// cap nhat but phe thông tin đã bút phê
                    vbdDA.UpdateButPhe(ItemID);
                }
                else
                {
                    vbdDA.UpdateLastTimeButPhe(ItemID);
                }
                if (string.IsNullOrEmpty(strThongBao))
                {
                    addLuongXulyVanBanDen(string.Format("{0} nhập bút phê", tenNguoiButPhe), (int)enumMaThaoTac.butphe,
                            clsViewLink.vanbanden, ItemID, ItemID, null);

                    objMsg.Erros = false;
                    objMsg.Message = string.Format("Đã thêm mới ý kiến bút phê: <b>{0}</b>", Server.HtmlEncode(butpheItem.Title));
                    objMsg.ID = "0";

                    #region add for log
                    LVanBanDen tem2 = vbdDA.GetSoKyHieuById(ItemID);
                    /// add log: 
                    addLog(EnumVanBanDen.strButPhe + " số: " + tem2.vbdSoKyHieu, EnumVanBanDen.ButPhe, EnumDoiTuong.VanBanDen, ItemID, tem2.vbdDSTGXuLy, true, true, true);
                    #endregion
                }
                else
                {
                    objMsg.Erros = true;
                    objMsg.Message = "Thông tin quá dài, chỉ nên dưới 255 ký tự";
                }
            }
            catch (Exception ex)
            {
                objMsg.Erros = true;
                objMsg.Message = "Thông tin quá dài, chỉ nên dưới 255 ký tự";
            }
            return objMsg;
        }

        /// <summary>
        /// kieemr tra van ban bij trung tren hej thong
        /// </summary>
        /// <returns></returns>
        public Utils.objMessage CheckTrungVB(string strSoKyHieu, string ngayDen)
        {

            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                DateTime? NgayDen = TConvert.ToDateTimeFormat(ngayDen);
                bool boolCheckVB = vbdDA.CheckExits(strSoKyHieu, 0, currentUser.CurrentIDDonVi, NgayDen.Value.Year.ToString());
                if (boolCheckVB)
                {
                    _objMesage.Erros = true;
                    _objMesage.Message = "Văn bản đến đã tồn tại. Bạn có muốn nhập văn bản tiếp tục hay ko?";
                }
                else
                {
                    _objMesage.Erros = false;
                    _objMesage.Message = "Văn bản đến đã tồn tại. Bạn có muốn nhập văn bản tiếp tục hay ko?";
                }
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// Thêm mới danh mục
        /// </summary>
        /// <returns></returns>
        private Utils.objMessage ImportEdXML()
        {
            Utils.objMessage objMsg = new Utils.objMessage();
            try
            {
                if (!string.IsNullOrEmpty(Request["listValueFileAttach"]))
                {

                    string strListFileAttach = Request["listValueFileAttach"];
                    System.Web.Script.Serialization.JavaScriptSerializer oSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    List<Utils.FileAttachForm> ltsFileForm = oSerializer.Deserialize<List<Utils.FileAttachForm>>(strListFileAttach);

                    string filePath = "/Uploads/ajaxUpload/";

                    LVanBanDen vanbanImport = new LVanBanDen();

                    if (ltsFileForm[0].FileName.IndexOf(".xml") > 0)
                    {
                        string filepathUpload = HttpContext.Current.Server.MapPath(filePath + ltsFileForm[0].FileServer);

                        var content = string.Empty;
                        using (StreamReader reader = new StreamReader(filepathUpload))
                        {
                            content = reader.ReadToEnd();
                            reader.Close();
                        }

                        var contentTemp = content;
                        content = Regex.Replace(content, "edXML:", "");
                        content = Regex.Replace(content, "soap-env:", "");


                        string replare = "<Envelope>";// xmlns:soap-env=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xsi:schemaLocation=\"http://schemas.xmlsoap.org/soap/envelope/ \"
                        string toReplace = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Envelope>";

                        content = Regex.Replace(content, replare, toReplace);


                        using (StreamWriter writer = new StreamWriter(filepathUpload))
                        {
                            writer.Write(content);
                            writer.Close();
                        }

                        XmlTextReader readerXML = new XmlTextReader(filepathUpload);


                        readerXML.WhitespaceHandling = WhitespaceHandling.None;
                        XmlDocument xd = new XmlDocument();
                        xd.Load(readerXML);

                        #region import co quan ban hanh
                        /// ten to chuc gui van ban/Co quan ban hanh
                        string OrganName = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/From/OrganName").InnerText.ToString();

                        LCoQuanBanHanhDA coquanDA = new LCoQuanBanHanhDA();
                        /// kiểm tra sự tồn tại của cơ quan ban hành
                        bool isExits = coquanDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, OrganName);
                        string idCoQuanBanHanh;
                        if (!isExits)
                        {
                            LCoQuanBanHanh coquanItem = new LCoQuanBanHanh();
                            coquanItem.Title = OrganName;
                            coquanItem.cqbhDonViLookup.LookupId = currentUser.CurrentIDDonVi;
                            idCoQuanBanHanh = coquanDA.AddReturnID(coquanItem);
                        }
                        else
                        {
                            idCoQuanBanHanh = coquanDA.GetByTitle(OrganName).ID.ToString();
                        }
                        vanbanImport.vbdCoQuanBanHanh = new SPFieldLookupValueCollection(idCoQuanBanHanh + ";#");

                        #endregion co quan ban hanh

                        /// danh sach nhan van ban
                        XmlNodeList OrganIdNodes = xd.DocumentElement.SelectNodes("/Envelope/Header/MessageHeader/To/OrganId");
                        XmlNodeList OrganNameNodes = xd.DocumentElement.SelectNodes("/Envelope/Header/MessageHeader/To/OrganName");
                        XmlNodeList EmailNodes = xd.DocumentElement.SelectNodes("/Envelope/Header/MessageHeader/To/Email");


                        /// so van ban + so ky hieu
                        string CodeNumberNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/Code/CodeNumber").InnerText.ToString();
                        string CodeNotationNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/Code/CodeNotation").InnerText.ToString();

                        vanbanImport.vbdSoKyHieu = CodeNotationNode;



                        /// dia chi + ngay ban hanh
                        string PlaceNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/PromulgationInfo/Place").InnerText.ToString();
                        string DateNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/PromulgationInfo/Date").InnerText.ToString();
                        vanbanImport.vbdNgayBanHanh = Convert.ToDateTime(DateNode);


                        #region loaij vawn banr
                        /// kieu van ban
                        string descriptionNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/type/description").InnerText.ToString();
                        /// kiem tra su ton tai cua loại văn bản
                        LPhanLoaiVanBanDA loaivanbanDA = new LPhanLoaiVanBanDA();
                        bool isLoaiVanBanExits;
                        string idLoaiVanBan;
                        isLoaiVanBanExits = loaivanbanDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, descriptionNode, "Văn bản đến");
                        if (!isLoaiVanBanExits)
                        {
                            LPhanLoaiVanBan loaivanban = new LPhanLoaiVanBan();
                            loaivanban.Title = descriptionNode;
                            loaivanban.phanloaiVanBan = "Văn bản đến";
                            loaivanban.phanloaiDonViLookup.LookupId = currentUser.CurrentIDDonVi;
                            loaivanban.phanloaiSTT = 0;
                            idLoaiVanBan = loaivanbanDA.AddReturnID(loaivanban);
                        }
                        else
                        {
                            idLoaiVanBan = loaivanbanDA.GetByTitle_VanBanDen(currentUser.CurrentIDDonVi, descriptionNode).ID.ToString();
                        }

                        vanbanImport.vbdLoaiVanBan.LookupId = Convert.ToInt32(idLoaiVanBan); // cap nhat loai van ban
                        #endregion Loai van ban



                        /// trich yeu
                        string SubjectNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/Subject").InnerText.ToString();
                        vanbanImport.vbdTrichYeu = SubjectNode;


                        #region nguoi ky + chuc danh
                        /// nguoi ky + chuc danh                        
                        string FullNameNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/Author/FullName").InnerText.ToString();
                        string FunctionNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/Author/Function").InnerText.ToString();

                        /// kiem tra su ton tai cua nguoi ky
                        LNguoiKyDA nguoikyDA = new LNguoiKyDA();
                        bool nguoikyTonTai;
                        string idNguoiKy;
                        nguoikyTonTai = nguoikyDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, FullNameNode);
                        if (!nguoikyTonTai)
                        {
                            LNguoiKy nguoiky = new LNguoiKy();
                            nguoiky.Title = FullNameNode;
                            nguoiky.nguoikyDonViLookup.LookupId = currentUser.CurrentIDDonVi;
                            nguoiky.nguoikySTT = 0;
                            idNguoiKy = nguoikyDA.AddReturnID(nguoiky);
                        }
                        else
                        {
                            idNguoiKy = nguoikyDA.GetByTitle(currentUser.CurrentIDDonVi, FullNameNode).ID.ToString();
                        }

                        vanbanImport.vbdNguoiKyVB.Add(new SPFieldLookupValue(idNguoiKy + ";#"));


                        /// kiem tra su ton tai cua chuc danh
                        LChucVuDA chucvuDA = new LChucVuDA();
                        bool chucvuTonTai;
                        string idChucVu;
                        chucvuTonTai = chucvuDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, FunctionNode);
                        if (!chucvuTonTai)
                        {
                            LChucVu chucvu = new LChucVu();
                            chucvu.Title = FunctionNode;
                            chucvu.chucdanhDonViLookup.LookupId = currentUser.CurrentIDDonVi;

                            idChucVu = chucvuDA.AddReturnID(chucvu);
                        }
                        else
                        {
                            idChucVu = chucvuDA.GetByTitle(currentUser.CurrentIDDonVi, FunctionNode).ID.ToString();
                        }

                        vanbanImport.vbdChucVuNguoiKy.Add(new SPFieldLookupValue(idChucVu + ";#"));
                        #endregion nguoi ky + chuc danh



                        #region do mat do khan
                        /// do mat + do khan
                        string SecretNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/OtherInfo/Secret").InnerText.ToString();
                        string PriorityNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/OtherInfo/Priority").InnerText.ToString();

                        // kiem tra thong tin do mat
                        LTinhChatVanBanDA tinhchatDA = new LTinhChatVanBanDA();
                        bool isTinhChatVBDen_MatExits;
                        string idTinhChatVanBanDen_Mat;

                        bool isTinhChatVBDen_KhanExits;
                        string idTinhChatVanBanDen_Khan;
                        /// kiem tra thông tiin  văn ban mat
                        isTinhChatVBDen_MatExits = tinhchatDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, SecretNode, "Văn bản mật");
                        if (!isTinhChatVBDen_MatExits)
                        {
                            LTinhChatVanBan tinhchat_Mat = new LTinhChatVanBan();
                            tinhchat_Mat.Title = SecretNode;
                            tinhchat_Mat.tinhchatLoaiTinhChatVB = "Văn bản mật";
                            tinhchat_Mat.tinhchatDonViLookup.LookupId = currentUser.CurrentIDDonVi;
                            idTinhChatVanBanDen_Mat = tinhchatDA.AddReturnID(tinhchat_Mat);
                        }
                        else
                        {
                            idTinhChatVanBanDen_Mat = tinhchatDA.GetByTitle(currentUser.CurrentIDDonVi, SecretNode, "Văn bản mật").ID.ToString();
                        }

                        //vanbanmoi.vbdDoMat = vanbanChuyenTiep.vbdDoMat;
                        vanbanImport.vbdDoMat.LookupId = Convert.ToInt32(idTinhChatVanBanDen_Mat);

                        /// kiem tra thông tiin  văn ban khẩn
                        isTinhChatVBDen_KhanExits = tinhchatDA.CheckExitsForAdd(currentUser.CurrentIDDonVi, PriorityNode, "Văn bản khẩn");
                        if (!isTinhChatVBDen_KhanExits)
                        {
                            LTinhChatVanBan tinhchat_Khan = new LTinhChatVanBan();
                            tinhchat_Khan.Title = PriorityNode;
                            tinhchat_Khan.tinhchatLoaiTinhChatVB = "Văn bản khẩn";
                            tinhchat_Khan.tinhchatDonViLookup.LookupId = currentUser.CurrentIDDonVi;
                            idTinhChatVanBanDen_Khan = tinhchatDA.AddReturnID(tinhchat_Khan);
                        }
                        else
                        {
                            idTinhChatVanBanDen_Khan = tinhchatDA.GetByTitle(currentUser.CurrentIDDonVi, PriorityNode, "Văn bản khẩn").ID.ToString();
                        }

                        vanbanImport.vbdDoKhan.LookupId = Convert.ToInt32(idTinhChatVanBanDen_Khan);

                        #endregion Do mat, do khan



                        /// so ban + so to
                        string PromulgationAmountNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/OtherInfo/PromulgationAmount").InnerText.ToString();
                        string PageAmountNode = xd.DocumentElement.SelectSingleNode("/Envelope/Header/MessageHeader/OtherInfo/PageAmount").InnerText.ToString();
                        vanbanImport.vbdSoBan = PromulgationAmountNode;
                        vanbanImport.vbdSoTo = PageAmountNode;

                        #region lay file dinh kem
                        /// get danh sach file dinh kem
                        XmlNodeList AttachmentsNodes = xd.DocumentElement.SelectNodes("/Envelope/document/attach/attachment");

                        List<FileAttach> lstFile = new List<FileAttach>();
                        foreach (XmlNode node in AttachmentsNodes)
                        {
                            FileAttach file = new FileAttach();
                            string filename = node.SelectSingleNode("name").InnerText.ToString();

                            file.Name = filename;

                            string value = node.SelectSingleNode("value").InnerText.ToString();

                            //byte[] toDecodeByte = System.Text.UnicodeEncoding.Unicode.GetBytes(value);
                            byte[] toDecodeByte = Convert.FromBase64String(value);


                            file.DataFile = toDecodeByte;

                            lstFile.Add(file);
                        }
                        vanbanImport.ListFileAttachAdd = lstFile;
                        #endregion lay file dinh kem

                        vanbanImport.vbdGroupLookup.LookupId = currentUser.CurrentIDDonVi;
                        vbdDA.Add(vanbanImport);
                        readerXML.Close();

                        objMsg.Erros = false;
                        objMsg.Message = "Import văn bản thành công: " + vanbanImport.vbdSoKyHieu;

                    }
                    else
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Bạn phải nhập file xml!";
                    }
                }

            }
            catch (Exception ex)
            {
                objMsg.Erros = true;
                objMsg.Message = "Định dạng file XML không chuẩn, mời bạn kiểm tra lại. <br>" + ex.Message;
            }
            return objMsg;
        }

        /// <summary>
        ///  thực hiện attach văn bản vào ho so 
        /// </summary>
        /// <returns></returns>
        public Utils.objMessage Attach()
        {
            Utils.objMessage objMsg = new Utils.objMessage();
            try
            {

                LHoSoCongViecDA hsDA = new LHoSoCongViecDA();
                string strIdHoSo = Request["hosoid"].ToString();

                string strThongBao = hsDA.AttachVanBanToHoso(ItemID, strIdHoSo);


                if (string.IsNullOrEmpty(strThongBao))
                {
                    objMsg.Message = "Chuyển văn bản vào hồ sơ thành công";
                    objMsg.Erros = false;

                    LVanBanDen tem = vbdDA.GetSoKyHieuById(ItemID);
                    /// add log: chuyển văn bản vào hồ sơ
                    addLog(EnumVanBanDen.strChuyenVBVaoHS + " số: " + tem.vbdSoKyHieu + " vào hồ sơ", EnumVanBanDen.ChuyenVBVaoHS, EnumDoiTuong.VanBanDen, ItemID, new SPFieldLookupValueCollection(""), false, false, false);

                }
                else
                {
                    objMsg.Message = strThongBao;
                    objMsg.Erros = true;
                }

                return objMsg;
            }
            catch (Exception ex)
            {
                objMsg.Erros = true;
                objMsg.Message = ex.ToString();
            }
            return objMsg;
        }


        private Utils.objMessage PhanLoaiVanBan()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                int hanxuly = 0;

                if (!string.IsNullOrEmpty(Request["vbdDuAn"]))
                    hanxuly = Convert.ToInt32(Request["vbdDuAn"]);
                if (hanxuly > 0)
                    vbdDA.UpdateLoaiVanBan(ItemID, hanxuly);

                _objMesage.Erros = false;
                _objMesage.Message = "Đã cập nhật loại văn bản thành công.";
                _objMesage.ID = "0";

            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// CapNhatHXL
        /// </summary>
        /// <returns>Thông báo</returns>
        private Utils.objMessage CapNhatHXL()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                DateTime? hanxuly= null;

                if (!string.IsNullOrEmpty(Request["vbdHanXuLy"]))
                    hanxuly = Convert.ToDateTime(Request["vbdHanXuLy"], dtfi);

                vbdDA.UpdateHanXuLy(ItemID, hanxuly, currentUser.ID);

                LVanBanDen tem = vbdDA.GetSoKyHieuById(ItemID);
                /// add log: cập nhật hạn xử lý
                addLog(EnumVanBanDen.strCapNhatHXL + " số: " + tem.vbdSoKyHieu, EnumVanBanDen.CapNhatHXL, EnumDoiTuong.VanBanDen, ItemID, tem.vbdDSTGXuLy, false, false, false);

                _objMesage.Erros = false;
                _objMesage.Message = "Đã cập nhật hạn xử lý thành công.";
                _objMesage.ID = "0";

            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// Done
        /// </summary>
        /// <returns>Thông báo</returns>
        private Utils.objMessage Know()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                vbdDA.UpdateTrangThaiXemDeBiet(ItemID, currentUser.ID); // cap nhat trang thai xem de biet cua nguoi dung                
                _objMesage.Erros = false;
                _objMesage.Message = "Đã chuyển sang xem để biết.";
                _objMesage.ID = "0";



            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// Done
        /// </summary>
        /// <returns>Thông báo</returns>
        private Utils.objMessage Done()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                int wfid = 0; // cong việc của người dùng
                if (Request["wfid"] != null && !string.IsNullOrEmpty(Request["wfid"]))
                {
                    wfid = Convert.ToInt32(Request["wfid"]);
                }

                SPListItem vanbandenItem = vbdDA.UpdateTrangThaiXuLyDone(ItemID, currentUser.ID, wfid); // cap nhat trang thai xu ly xong cua nguoi dung

                //LVanBanDen tem = vbdDA.GetSoKyHieuById(ItemID);
                LVanBanDen tem = new LVanBanDen();
                tem.vbdSoKyHieu = (vanbandenItem["vbdSoKyHieu"] != null) ? vanbandenItem["vbdSoKyHieu"].ToString()  : "";
                tem.vbdDSTGXuLy = (vanbandenItem["vbdDSTGXuLy"] != null) ? new SPFieldLookupValueCollection(vanbandenItem["vbdDSTGXuLy"].ToString()) : null;
                
                addLuongXulyVanBanDen("đã xử lý xong văn bản", (int)enumMaThaoTac.XuLyXong, clsViewLink.vanbanden, ItemID, ItemID, null);
                /// add log: Xử lý xong văn bản
                addLog(EnumVanBanDen.strDone + " số: " + tem.vbdSoKyHieu, EnumVanBanDen.Done, EnumDoiTuong.VanBanDen, ItemID, tem.vbdDSTGXuLy, false, false, false);

                _objMesage.Erros = false;
                _objMesage.Message = "Đã xử lý xong.";
                _objMesage.ID = "0";
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// Xuất file văn bản ra file EdXML
        /// </summary>
        /// <returns></returns>
        private void ExportEdXML()
        {
            vbdItem = vbdDA.GetById(ItemID);

            StringBuilder xml = new StringBuilder();
            //xml.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
            xml.AppendLine("<soap-env:Envelope xmlns:soap-env=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xsi:schemaLocation=\"http://schemas.xmlsoap.org/soap/envelope/ \">");
            xml.AppendLine("<soap-env:Header>");
            xml.AppendLine("    <edXML:MessageHeader  OriginalBodyRequested=\"false\" ImmediateResponseRequired=\"true\">");

            xml.AppendLine("        <edXML:From>"); //Thông tin nơi ban hành văn bản
            // TODO: thiếu
            xml.AppendLine("            <edXML:OrganId edXML:type=\"edXMLString\">EOF1102</edXML:OrganId>"); //ID cơ quan, tổ chức ban hành văn bản
            // TODO: thiếu
            xml.AppendLine("            <edXML:OrganizationInCharge></edXML:OrganizationInCharge>"); //Tên cơ quan, tổ chức chủ quản trực tiếp
            xml.AppendLine("            <edXML:OrganName>" + vbdItem.vbdCoQuanBanHanh[0].LookupValue + "</edXML:OrganName>"); //Tên cơ quan, tổ chức ban hành văn bản
            // TODO: thiếu                
            xml.AppendLine("            <edXML:Email></edXML:Email>"); //Thư điện tử đại diện cho nơi ban hành
            xml.AppendLine("        </edXML:From>");

            xml.AppendLine("        <edXML:To>"); //Thông tin nơi nhận văn bản(có thể có nhiều trường mỗi nơi nhận văn bản tương ứng mới một < edXML:to>)
            // TODO: thiếu
            xml.AppendLine("            <edXML:OrganId edXML:type=\"edXMLString\"></edXML:OrganId>"); //ID nơi nhận văn bản
            // TODO: thiếu
            xml.AppendLine("            <edXML:OrganName></edXML:OrganName>"); //Tên nơi nhận văn bản
            // TODO: thiếu
            xml.AppendLine("            <edXML:Email></edXML:Email>"); //Thư điện tử đại diện cho nhận gửi                      
            xml.AppendLine("        </edXML:To>");

            xml.AppendLine("        <edXML:DocumentId>"); //ID của văn bản
            xml.AppendLine("            " + vbdItem.ID + "");
            xml.AppendLine("        </edXML:DocumentId>");

            xml.AppendLine("        <edXML:Code>"); //Số, kí hiệu văn bản
            xml.AppendLine("            <edXML:CodeNumber>" + vbdItem.vbdSoVanBan + "</edXML:CodeNumber>"); //Số của văn bản
            xml.AppendLine("            <edXML:CodeNotation>" + vbdItem.vbdSoKyHieu + "</edXML:CodeNotation>"); //Ký hiệu của văn bản, công văn                        
            xml.AppendLine("        </edXML:Code>");

            xml.AppendLine("        <edXML:PromulgationInfo>"); //Địa danh và ngày, tháng, năm ban hành văn bản
            // TODO: Thiếu -> bổ sung là địa chỉ của cơ quan ban hành
            xml.AppendLine("            <edXML:Place></edXML:Place>"); // Địa danh
            xml.AppendLine("            <edXML:Date>" + vbdItem.vbdNgayBanHanh + "</edXML:Date>"); //Ngày, tháng, năm ban hành
            xml.AppendLine("        </edXML:PromulgationInfo>");

            xml.AppendLine("        <edXML:type>"); //Tên loại văn bản
            xml.AppendLine("            <value>" + vbdItem.vbdLoaiVanBan.LookupId + "</value>");
            xml.AppendLine("            <description>" + vbdItem.vbdLoaiVanBan.LookupValue + "</description>");
            xml.AppendLine("        </edXML:type>");

            xml.AppendLine("        <edXML:Subject>" + vbdItem.vbdTrichYeu + "</edXML:Subject>"); // Trích yếu nội dung của văn bản
            // TODO: thiếu
            xml.AppendLine("        <edXML:Content></edXML:Content>"); // Nội dung văn bản

            // TODO: Thiếu
            xml.AppendLine("        <edXML:Author>"); //Quyền hạn, chức vụ, họ tên và chữ kí của người có thẩm quyền
            xml.AppendLine("            <edXML:Competence></edXML:Competence>"); //Quyền hạn của người ký
            xml.AppendLine("            <edXML:Function></edXML:Function>"); //Chức vụ của người ký
            xml.AppendLine("            <edXML:FullName></edXML:FullName>"); //Họ và tên người ký văn bản
            xml.AppendLine("        </edXML:Author>");

            // TODO: Thiếu
            xml.AppendLine("        <edXML:ToPlaces>"); // Danh sách Nơi nhận văn bản
            xml.AppendLine("            <edXML:Place></edXML:Place>"); //Nơi nhận văn bản.
            xml.AppendLine("        </edXML:ToPlaces>");

            xml.AppendLine("        <edXML:OtherInfo>"); //Các thành phần khác
            xml.AppendLine("            <edXML:Secret>" + vbdItem.vbdDoMat.LookupValue + "</edXML:Secret>"); //Mức độ mật văn bản
            xml.AppendLine("            <edXML:Priority>" + vbdItem.vbdDoKhan.LookupValue + "</edXML:Priority>"); //Độ khẩn văn bản
            // TODO: Thiếu
            xml.AppendLine("            <edXML:SphereOfPromulgation></edXML:SphereOfPromulgation>"); //Phạm vi, đối tượng được phổ biến, sử dụng hạn chế, sử dụng các hạn chế về phạm vi lưu hành
            // TODO: Thiếu
            xml.AppendLine("            <edXML:ContactInfo></edXML:ContactInfo>"); //Đối với công văn, ngoài các thành phần được quy định có thể bổ sung địa chỉ cơ quan, tổ chức; địa chỉ thư điện tử (E-Mail); số điện thoại, số Telex, số Fax; địa chỉ trang thông tin điện tử (Website).
            // TODO: Thiếu
            xml.AppendLine("            <edXML:TyperNotation></edXML:TyperNotation>"); //Ký hiệu người đánh máy
            xml.AppendLine("            <edXML:PromulgationAmount>" + vbdItem.vbdSoBan + "</edXML:PromulgationAmount>"); //Số lượng bản phát hành
            xml.AppendLine("            <edXML:PageAmount>" + vbdItem.vbdSoTo + "</edXML:PageAmount>"); //Số trang của văn bản
            xml.AppendLine("        </edXML:OtherInfo>");
            xml.AppendLine("    </edXML:MessageHeader>");

            //xml.AppendLine("    <edXML:TraceHeaderList edXML:id=\"3490sdo9\" edXML:version=\"1.0\" SOAP-ENV:mustUnderstand=\"1\" SOAP-ENV:actor=\"http://schemas.xmlsoap.org/soap/actor/next\">");
            //xml.AppendLine("        <edXML:TraceHeader>");
            //xml.AppendLine("            <edXML:Sender>");
            //xml.AppendLine("                <edXML:OrganId edXML:type=\"edXMLString\">EOF1102</edXML:OrganId>");
            //xml.AppendLine("                <edXML:OrganizationInCharge>PHONG CNTT</edXML:OrganizationInCharge>");
            //xml.AppendLine("                <edXML:Email>from@example3.com</edXML:Email>");
            //xml.AppendLine("            </edXML:Sender>");
            //xml.AppendLine("            <edXML:Receiver>");
            //xml.AppendLine("                <edXML:OrganId edXML:type=\"edXMLString\">EOF1103_1</edXML:OrganId>");
            //xml.AppendLine("                <edXML:OrganName>SO 1 Tinh A</edXML:OrganName>");
            //xml.AppendLine("                <edXML:Email>to1@example3.com</edXML:Email>");
            //xml.AppendLine("            <edXML:Receiver>");
            //xml.AppendLine("            <edXML:Timestamp>2000 -12-16T21:19:35Z</edXML:Timestamp>");
            //xml.AppendLine("        </edXML:TraceHeader>");
            //xml.AppendLine("    </edXML:TraceHeaderList>");

            //xml.AppendLine("    <edXML:ErrorList edXML:id=\"3490sdo9\" edXML:highestSeverity=\"error\" edXML:version=\"1.0\"  SOAP-ENV:mustUnderstand=\"1\">");
            //xml.AppendLine("    </edXML:ErrorList>");
            xml.AppendLine("</soap-env:Header>");

            xml.AppendLine("<soap-env:Body>");
            xml.AppendLine("    <edXML:Manifest edXML:version=\"1.0\">"); //Định nghĩa các file đính kèm
            xml.AppendLine("        <edXML:Reference xlink:href=\"cid:edXMLatt1@example.com\" xlink:role=\"XLinkRole\" xlink:type=\"simple\">");
            xml.AppendLine("            <edXML:Description xml:lang=\"Vi-vn\">file dinh kem theo cong van</edXML:Description>"); //Thông tin miêu tả cho file đính kèm
            xml.AppendLine("        </edXML:Reference>");
            xml.AppendLine("    </edXML:Manifest>");
            xml.AppendLine("</soap-env:Body>");
            xml.AppendLine("</soap-env:Envelope>");

            // Thiếu chữ ký số

            StringBuilder xmlDocument = new StringBuilder();

            if (vbdItem.ListFileAttach.Count > 0)
            {

                xmlDocument.AppendLine("<edXML:document>");
                xmlDocument.AppendLine("    <edXML:attach>"); //Chứa file đính kèm
                for (int i = 0; i < vbdItem.ListFileAttach.Count; i++)
                {
                    SPFile attachmentFile = vbdDA.SpListProcess.ParentWeb.GetFile(vbdItem.ListFileAttach[i].Url);
                    Stream stream = attachmentFile.OpenBinaryStream();
                    StreamReader reader = new StreamReader(stream);
                    String fileContent = reader.ReadToEnd();

                    xmlDocument.AppendLine("        <edXML:attachment Content-Type=\"application/zip\" Content-Transfer-Encoding=\"base64\">"); //Kiểu của file đính kèm
                    xmlDocument.AppendLine("            <name>ten_file_1</name>"); //Tên file đính kèm
                    xmlDocument.AppendLine("            <value>"); //Nội dung file đính kèm đã được mã hóa
                    xmlDocument.AppendLine("                " + Convert.ToBase64String(System.Text.UnicodeEncoding.Unicode.GetBytes(fileContent)) + "");
                    xmlDocument.AppendLine("            </value>");
                    xmlDocument.AppendLine("        </edXML:attachment>");
                }
                xmlDocument.AppendLine("    </edXML:attach>");
                xmlDocument.AppendLine("</edXML:document>");
            }

            string tempPath = Server.MapPath("~/Uploads/Temp"); // Thư mục lưu file
            XmlDocument docAttachment;

            XmlDocument docHeader = GetXmlDocumentFromString(xml.ToString());
            if (vbdItem.ListFileAttach.Count > 0)
            {
                docAttachment = GetXmlDocumentFromString(xmlDocument.ToString());
                docAttachment.Save(tempPath + "\\" + vbdItem.ID + "_Attachment.xml"); // Lưu file xml chứa file đính kèm văn bản
            }

            docHeader.Save(tempPath + "\\" + vbdItem.ID + "_header.xml"); // Lưu file xml chứa thông tin cơ bản của văn bản

            // Nén các file xml thành 1 file zip
            using (ZipFile zip = new ZipFile())
            {
                zip.AddFile(tempPath + "\\" + vbdItem.ID + "_header.xml", "");
                if (vbdItem.ListFileAttach.Count > 0)
                    zip.AddFile(tempPath + "\\" + vbdItem.ID + "_Attachment.xml", "");
                zip.Save(tempPath + "\\" + vbdItem.ID.ToString() + ".zip");
            }

            // Trả file đính kèm về trình duyệt để người dùng lưu lại
            Response.Clear();
            Response.AppendHeader("Content-Disposition", "attachment; filename=\"" + vbdItem.vbdSoKyHieu + ".zip\"");
            // Chú ý: Nếu filename không nằm trong cặp "" thì Firefox không lưu file             
            Response.ContentType = "Content-Type: application/octet-stream";
            Response.WriteFile(tempPath + "\\" + vbdItem.ID.ToString() + ".zip");
            Response.End();
        }




        /// <summary>
        /// get so cua so văn bản
        /// </summary>
        public Utils.objMessage GetNum()
        {

            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                int intSovanban = vbdDA.GetSoVanBanMax(currentUser.CurrentIDDonVi, Convert.ToInt32(Request["scv"].ToString()));
                _objMesage.Erros = false;
                _objMesage.Message = intSovanban.ToString();
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;

            

        }
        public Utils.objMessage CheckThuHoiVB()
        {

            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                var spItem = vbdDA.GetCurrentUserAndGroupReceived(ItemID); // cap nhat trang thai xu ly xong cua nguoi dung
                SPFieldLookupValueCollection vbdCurrentUserReceived = new SPFieldLookupValueCollection(), vbdCurrentGroupsReceived = new SPFieldLookupValueCollection();
                if (spItem["vbdCurrentUserReceived"] != null)
                    vbdCurrentUserReceived = new SPFieldLookupValueCollection(spItem["vbdCurrentUserReceived"].ToString());
                //if (spItem["vbdCurrentGroupsReceived"] != null)
                //    vbdCurrentGroupsReceived = new SPFieldLookupValueCollection(spItem["vbdCurrentGroupsReceived"].ToString());
                //

                //
                if (vbdCurrentUserReceived.Count == 1)
                {
                    _objMesage.ID = "0";
                }
                else
                {
                    _objMesage.Erros = false;
                    _objMesage.ID = ItemID.ToString();
                }
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }
        /// <summary>
        /// thu hoi van ban
        /// </summary>
        /// ThuHoiVBWithUsers
        /// 
        public Utils.objMessage ThuHoiVBWithUsers()
        {

            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                if (Request["CurrentUsers"] != null)
                {
                    var listIDUser = Request["CurrentUsers"].Split(',').Where(item => !string.IsNullOrEmpty(item)).Select(item => Convert.ToInt32(item)).ToArray();
                    var vbdCurrentUserReceivedRemove = vbdDA.ThuHoiVBWithUsers(ItemID, listIDUser); // cap nhat trang thai xu ly xong cua nguoi dung
                    foreach (var spFieldLK in vbdCurrentUserReceivedRemove)
                    {
                        addLuongXulyVanBanDen("thu hồi văn bản gửi đến", (int)enumMaThaoTac.thuhoiVbDen,
                            clsViewLink.vanbanden, ItemID, ItemID, spFieldLK);
                    }
                    _objMesage.Erros = false;
                    _objMesage.Message = "Đã thu hồi văn bản thành công.";
                    _objMesage.ID = "0";
                }
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }
        public Utils.objMessage ThuHoiVB()
        {
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                var spItem = vbdDA.GetCurrentUserAndGroupReceived(ItemID); // cap nhat trang thai xu ly xong cua nguoi dung
                SPFieldLookupValueCollection vbdCurrentUserReceived = new SPFieldLookupValueCollection();
                if (spItem["vbdCurrentUserReceived"] != null)
                    vbdCurrentUserReceived = new SPFieldLookupValueCollection(spItem["vbdCurrentUserReceived"].ToString());
                vbdDA.ThuHoiVB(ItemID); // cap nhat trang thai xu ly xong cua nguoi dung
                foreach (var spFieldLK in vbdCurrentUserReceived)
                {
                    addLuongXulyVanBanDen("thu hồi văn bản gửi đến", (int)enumMaThaoTac.thuhoiVbDen,
                        clsViewLink.vanbanden, ItemID, ItemID, spFieldLK);
                }
                _objMesage.Erros = false;
                _objMesage.Message = "Đã thu hồi văn bản thành công.";
                _objMesage.ID = "0";

            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }

        /// <summary>
        /// theo dõi văn bản
        /// </summary>
        public void UnFolow()
        {

            try
            {
                vbdDA.RemoveFaverite(ItemID, currentUser.ID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// theo dõi văn bản
        /// </summary>
        public void Folow()
        {
            if (!currentUser.ListPermission.TheoDoiGuiNhanVanBanDen)
            {
                throw new Exception("Bạn không có quyền thực hiện thao tác này");
            }
            vbdDA.AddFaverite(ItemID, currentUser.ID);

        }

        /// <summary>
        /// Thêm mới danh mục
        /// </summary>
        /// <returns></returns>
        private Utils.objMessage Add()
        {
            if (!currentUser.ListPermission.ThemVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count == 0)
            {
                throw new Exception("Bạn không có quyền thực hiện thao tác này");
            }
            Utils.objMessage objMsg = new Utils.objMessage();
            List<LUser> lstChanhVP = new List<LUser>();
            try
            {

                vbdItem = new LVanBanDen();
                butpheItem = new LButPheVanBanDen();
                vbdItem.UpdateObject(this.Request);


                if (vbdItem.vbdCoQuanBanHanh.Count == 0)
                {
                    objMsg.Erros = true;
                    objMsg.Message = "Cơ quan ban hành chưa được nhập";
                }

                if (Request["vbdSoKyHieu"] == null || string.IsNullOrEmpty(Request["vbdSoKyHieu"].ToString()))
                {
                    objMsg.Erros = true;
                    objMsg.Message = "Số ký hiệu chưa được điền";
                }

                if (Request["vbdYkienLanhDaoPhuTrach"] != null && Request["vbdYkienLanhDaoPhuTrach"].ToString().Length > 0)
                {
                    if (Request["vbdLanhDaoPhuTrach"].ToString().Equals("0"))
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Lãnh đạo chưa được chọn!";
                    }
                }

                if (Request["vbdYKienLanhDaoVanPhong"] != null && Request["vbdYKienLanhDaoVanPhong"].ToString().Length > 0)
                {
                    if (Request["vbdLanhDaoVanPhong"].ToString().Equals("0"))
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Chánh văn phòng chưa được chọn!";
                    }
                }

                if (Request["vbdLanhDaoPhuTrach"] != null)
                {
                    // add lanh dao vao danh sach xu ly
                    if (!Request["vbdLanhDaoPhuTrach"].ToString().Equals("0"))
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoPhuTrach"].ToString()));
                        vbdItem.vbdUserChuaXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoPhuTrach"].ToString()));
                    }
                }
                if (Request["vbdLanhDaoVanPhong"] != null)
                {
                    /// add chanh van phong dao vao danh sach xu ly
                    if (!Request["vbdLanhDaoVanPhong"].ToString().Equals("0"))
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoVanPhong"].ToString()));
                        vbdItem.vbdUserChuaXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoVanPhong"].ToString()));
                    }
                }


                //check với trường hợp văn bản phòng, hoặc văn bản đơn vị.
                if (Request["vanbandv"] != null) //trường hợp có thể là văn bản phòng
                {
                    int intvanbandv = Convert.ToInt32(Request["vanbandv"]);
                    if (intvanbandv == 0) //van ban phong
                    {
                        vbdItem.vbdPBLookup = new SPFieldLookupValue(Request["vbdPBLookup1"]);
                    }
                    else
                        vbdItem.vbdPBLookup = null;
                }
                else
                    vbdItem.vbdPBLookup = null;

                #region tự động chuyển văn bản đến người dùng hoặc phòng ban được cấu hình
                // huylm Edit 31/05
                // Get danh sách các phong ban hoạc người dùng tự động nhận văn bản.
                //if (vbdItem.vbdPBLookup != null && vbdItem.vbdPBLookup.LookupId > 0) // neeus laf van ban cua phong, thi ko thuc hien thao tac chuyen tu dong
                //{
                try
                {
                    #region Code tu dong nhan van ban
                    //LNhanVanBanDA nhanvanbanDA = new LNhanVanBanDA();
                    //LNhanVanBan lnhanvanbanItem = nhanvanbanDA.GetById(1);
                    //vbdItem.vbdDSPBDVTGXuLy.AddRange(lnhanvanbanItem.dsPBNhanVB);

                    #endregion  /// timf taif khoan dai dien de add vao van ban
                    //foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                    //{
                    //    SPFieldLookupValue daidien = groupDA.GetTaiKhoanDaiDien(item.LookupId);
                    //    if (daidien.LookupId > 0)
                    //    {
                    //        vbdItem.vbdDSTGXuLy.Add(daidien);
                    //        vbdItem.vbdUserChuaXuLy.Add(daidien);


                    //        vbdItem.vbdCurrentUserReceived.Add(daidien);
                    //    }
                    //}
                    //if (vbdItem.vbdDSPBDVTGXuLy.Count > 0) /// neeus co phong ban thi da sent
                    //    vbdItem.vbdIsSentVanBan = true;


                    //vbdItem.vbdDSTGXuLy.AddRange(lnhanvanbanItem.dsNguoiDungNhanVB);
                    //vbdItem.vbdCurrentUserReceived.AddRange(lnhanvanbanItem.dsNguoiDungNhanVB);
                    //vbdItem.vbdUserChuaXuLy.AddRange(lnhanvanbanItem.dsNguoiDungNhanVB);

                }
                catch (Exception exx)
                {
                    //trường hợp ko có List Nhaanvanban thì ko làm gì cả
                }
                //}
                #endregion

                #region ko sử dụng chánh văn phòng nữa
                //if(đơn vị của người dùng là ubnd tự động add lanhdaovp vào list vbdDSTGXuLy tham gia xử lý
                //if (groupDA.CheckDonViIsUBND(currentUser.CurrentIDDonVi))
                //{
                //    //get danh sách user là lãnh đạo văn phòng
                //    lstChanhVP = new LUserDA().GetUserChanhVanPhong();
                //    foreach (LUser objUser in lstChanhVP)
                //    {
                //        vbdItem.vbdDSTGXuLy.Add(new SPFieldLookupValue(objUser.ID, objUser.Title));
                //        vbdItem.vbdUserChuaXuLy.Add(new SPFieldLookupValue(objUser.ID, objUser.Title));

                //    }
                //    if (lstChanhVP.Count > 0)
                //        vbdItem.vbdIsSentVanBan = true;


                //}
                #endregion ko sử dụng chánh văn phòng nữa

                if (Request["intTaoHo"] != null) /// truong hop tao ho so
                {
                    if (Request["intchuyen"].ToString().Equals("1"))
                    {
                        if (Request["intphongban"] == null || string.IsNullOrEmpty(Request["intphongban"]))
                        {
                            objMsg.Erros = true;
                            objMsg.Message = "Chưa chọn phòng ban để chuyển văn bản!";
                        }

                        // chuyển giá trị phòng ban thành lookup value
                        vbdItem.vbdDSPBDVTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intphongban"].ToString()));

                    }
                    else
                    { //intchuyen = 0
                        if (Request["intcanbo"] == null || string.IsNullOrEmpty(Request["intcanbo"]))
                        {
                            objMsg.Erros = true;
                            objMsg.Message = "Chưa chọn cán bộ để chuyển văn bản!";
                        }
                        // chuyển giá trị cán bộ thành lookup value
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                    }
                }
                else
                { // truong hop ko tao ho so / Trường hợp sử dụng chính
                    if (Request["intchuyen"].ToString().Equals("1"))
                    {
                        vbdItem.vbdDSPBDVTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intphongban"].ToString()));

                        /// timf taif khoan dai dien de add vao van ban
                        foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                        {
                            SPFieldLookupValue daidien = groupDA.GetTaiKhoanDaiDien(item.LookupId);
                            if (daidien.LookupId > 0)
                            {
                                vbdItem.vbdDSTGXuLy.Add(daidien);
                                vbdItem.vbdUserChuaXuLy.Add(daidien);
                                // add danh sách người được gửi hiện thời: edit 0403
                                vbdItem.vbdCurrentUserReceived.Add(daidien);
                            }
                        }
                        if (vbdItem.vbdDSPBDVTGXuLy.Count > 0) /// neeus co phong ban thi da sent
                            vbdItem.vbdIsSentVanBan = true;
                    }
                    else
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                        // add danh sách người được gửi hiện thời: edit 0403
                        vbdItem.vbdCurrentUserReceived.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));

                        vbdItem.vbdUserChuaXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                        if (vbdItem.vbdDSTGXuLy.Count > 0)   // neu co nguoi dung thi chua sent
                            vbdItem.vbdIsSentVanBan = true;
                    }
                }

                if (!objMsg.Erros)/// nếu không có lỗi xảy ra
                {

                    // đổi tên file đính kèm trong trường hợp scan
                    if (Page.Request["isscan"] == "true")
                    {
                        if (!string.IsNullOrEmpty(Page.Request["listFileScan"]))
                        {
                            string[] listFileName = Page.Request["listFileScan"].Split('*');
                            int i = 1;
                            foreach (FileAttach fileAttach in vbdItem.ListFileAttachAdd)
                            {
                                if (listFileName.Contains(fileAttach.Name))
                                {
                                    fileAttach.Name = vbdItem.vbdSoKyHieu + "-" + i.ToString() + ".tif";
                                    // Đổi các ký tự đặc biệt thành '-'
                                    string[] kytudacbiet = new string[] { "\\", "/", ":", "*", "?", "\"", "<", ">", "&", "%", "!", "^", "#", "$" };
                                    foreach (string kytu in kytudacbiet)
                                    {
                                        if (fileAttach.Name.Contains(kytu))
                                            fileAttach.Name = fileAttach.Name.Replace(kytu, "-");
                                    }
                                    i++;
                                }

                            }
                        }
                    }

                    #region add for log: danh sách người được chuyển văn bản khi tạo mới
                    SPFieldLookupValueCollection dsNguoiLienQuan = vbdItem.vbdUserChuaXuLy;
                    #endregion

                    // thiet lap nguoi thuc hien thao tac hien thoi
                    vbdItem.vbdCurrentUserAction = new SPFieldLookupValue(currentUser.ID.ToString() + ";#");
                    //
                    SPFieldLookupValueCollection temp = new SPFieldLookupValueCollection();
                    foreach (var spValue in vbdItem.vbdCurrentUserReceived)
                    {
                        if (!temp.Any(item => item.LookupId == spValue.LookupId))
                            temp.Add(spValue);
                    }
                    vbdItem.vbdCurrentUserReceived = temp;
                    //
                    temp = new SPFieldLookupValueCollection();
                    foreach (var spValue in vbdItem.vbdDSTGXuLy)
                    {
                        if (!temp.Any(item => item.LookupId == spValue.LookupId))
                            temp.Add(spValue);
                    }
                    vbdItem.vbdDSTGXuLy = temp;
                    //
                    temp = new SPFieldLookupValueCollection();
                    foreach (var spValue in vbdItem.vbdUserChuaXuLy)
                    {
                        if (!temp.Any(item => item.LookupId == spValue.LookupId))
                            temp.Add(spValue);
                    }
                    vbdItem.vbdUserChuaXuLy = temp;

                    string strThongBao = vbdDA.Add(vbdItem);

                    if (!strThongBao.Equals("Văn bản đã tồn tại trong hệ thống"))
                    {
                        int intIDVB = Convert.ToInt32(strThongBao);
                        #region cập nhật log xử lý
                        string strCurrentListIdInfoSent = string.Empty;
                        if (Convert.ToInt32(strThongBao) > 0)
                        {
                            ///add log: thêm mới văn bản
                            addLog(EnumVanBanDen.strThemMoi + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ThemMoi, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), new Microsoft.SharePoint.SPFieldLookupValueCollection(""), false, false, true);
                            if (Request["intchuyen"].ToString().Equals("1"))
                            {
                                /// dien thong tin gui van ban cho don vi
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoGroupNameReceived = item;
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = intIDVB;// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }

                                /// add log: chuyển văn bản
                                if (dsNguoiLienQuan.Count > 0)
                                {
                                    addLog(EnumVanBanDen.strChuyenNoiBo + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ChuyenNoiBo, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), dsNguoiLienQuan, true, true, true);
                                }
                            }
                            else
                            {
                                ///dien thong tin gui van ban cho can bo
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSTGXuLy)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoUserNameReceived = item;
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = intIDVB;// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }

                                /// add log: chuyển văn bản
                                if (dsNguoiLienQuan.Count > 0)
                                {
                                    addLog(EnumVanBanDen.strChuyenNoiBo + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ChuyenNoiBo, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), dsNguoiLienQuan, true, true, true);
                                }
                            }


                            ///bổ sung lưu luồng xử lý
                            ///05082015
                            ///Luồng xử lý thì sẽ xử lý theo nguwoif tham gia xử lý
                            ///gửi cho phòng đơn vị thì cũng quy ra người
                            ///
                            LLuongxulyDA objLuongXLDA = new LLuongxulyDA();
                            //log thêm mới
                            LLuongxuly objLuongXL = new LLuongxuly();
                            objLuongXL.VanBanDenLk = intIDVB;
                            objLuongXL.NguoiGui = new SPFieldLookupValue(currentUser.ID, "");
                            objLuongXL.NguoiNhan = new SPFieldLookupValue(currentUser.ID, "");
                            objLuongXL.MaTree = "001";
                            objLuongXL.MaThaoTac = (int)enumMaThaoTac.addvb;
                            objLuongXL.Title = string.Format("{0} {1}", currentUser.Title, objLuongXLDA.getTenthaotac(objLuongXL.MaThaoTac));
                            objLuongXL.ViewLink = objLuongXLDA.getLinkView(clsViewLink.vanbanden, intIDVB);
                            ///get thông tin về mức độ thụ thò của dòng log xử lý.
                            ///
                            LUserDA objuserDA = new LUserDA();
                            if (objLuongXL.NguoiNhan.LookupId > 0)
                                objLuongXL.MucDoThut = objuserDA.GetCapBacUser(objLuongXL.NguoiNhan.LookupId);
                            else objLuongXL.MucDoThut = objuserDA.GetCapBacUser(objLuongXL.NguoiGui.LookupId);
                            int intlogparent = objLuongXLDA.AddReturnID(objLuongXL);


                            ///
                            if (vbdItem.vbdDSTGXuLy.Count > 0)
                            {
                                objLuongXL = new LLuongxuly();
                                objLuongXL.VanBanDenLk = intIDVB;
                                objLuongXL.NguoiGui = new SPFieldLookupValue(currentUser.ID, "");
                                objLuongXL.MaThaoTac = (int)enumMaThaoTac.chuyenvb;

                                objLuongXL.ViewLink = objLuongXLDA.getLinkView(clsViewLink.vanbanden, intIDVB);
                                objLuongXL.LogParent = new SPFieldLookupValue(intlogparent, "");
                                int i = 1;
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSTGXuLy)
                                {
                                    objLuongXL.NguoiNhan = item;
                                    objLuongXL.MaTree = string.Format("00100{0}", i);
                                    if (objLuongXL.NguoiNhan.LookupId > 0)
                                        objLuongXL.MucDoThut = objuserDA.GetCapBacUser(objLuongXL.NguoiNhan.LookupId);
                                    else objLuongXL.MucDoThut = objuserDA.GetCapBacUser(objLuongXL.NguoiGui.LookupId);
                                    i++;
                                    objLuongXLDA.AddReturnID(objLuongXL);
                                }
                            }





                            if (lstChanhVP.Count > 0)
                            {
                                foreach (LUser item in lstChanhVP)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoUserNameReceived = new SPFieldLookupValue(item.ID, item.Title);
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = Convert.ToInt32(strThongBao);// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }
                            }

                            if (vbdItem.vbdDSXemDeBiet.Count > 0)
                            {
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSXemDeBiet)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoUserNameReceived = item;
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = Convert.ToInt32(strThongBao);// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }
                            }


                            #region cập nhật thông tin vừa nhận: edit 0403
                            vbdDA.UpdateCurrentListIdInfoSent(strCurrentListIdInfoSent, Convert.ToInt32(strThongBao), 0);
                            #endregion
                            if (vbdItem.vbdIsSentVanBan)
                            {
                                vbdDA.UpdateNguoiGiaoVB(Convert.ToInt32(strThongBao), currentUser.ID, currentUser.userTenTruyCap);
                            }
                            objMsg.Erros = false;
                            objMsg.Message = string.Format("Đã thêm mới {0}: <b>{1}</b>", TitleWebpart, Server.HtmlEncode(vbdItem.vbdSoKyHieu));
                            objMsg.ID = "0";
                        }
                        else
                        {
                            objMsg.Erros = true;
                            objMsg.Message = strThongBao;
                        }
                        #endregion
                    }
                    else
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Văn bản đã tồn tại trong hệ thống";
                    }
                }
            }
            catch (Exception ex)
            {
                objMsg.Erros = true;
                objMsg.Message = ex.ToString();
            }
            return objMsg;
        }


        /// <summary>
        /// Thêm mới văn bản từ email
        /// </summary>
        /// <returns></returns>
        private Utils.objMessage AddFromEmail()
        {
            LVanBanEmail emailItem = new LVanBanEmail();
            LVanBanEmailDA emailDA = new LVanBanEmailDA();
            if (Request["itememail"] != null && !string.IsNullOrEmpty(Request["itememail"].ToString()))
            {
                emailItem = emailDA.GetFileById(Convert.ToInt32(Request["itememail"].ToString()));
            }
            Utils.objMessage objMsg = new Utils.objMessage();
            try
            {

                vbdItem = new LVanBanDen();
                butpheItem = new LButPheVanBanDen();
                vbdItem.UpdateObject(this.Request);
                vbdItem.ListFileAttachAdd.AddRange(emailItem.ListFileAttach);
                #region kiểm tra các điều kiện
                if (vbdItem.vbdCoQuanBanHanh.Count == 0)
                {
                    objMsg.Erros = true;
                    objMsg.Message = "Cơ quan ban hành chưa được nhập";
                }

                if (Request["vbdSoKyHieu"] == null || string.IsNullOrEmpty(Request["vbdSoKyHieu"].ToString()))
                {
                    objMsg.Erros = true;
                    objMsg.Message = "Số ký hiệu chưa được điền";
                }

                if (Request["vbdYkienLanhDaoPhuTrach"] != null && Request["vbdYkienLanhDaoPhuTrach"].ToString().Length > 0)
                {
                    if (Request["vbdLanhDaoPhuTrach"].ToString().Equals("0"))
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Lãnh đạo chưa được chọn!";
                    }
                }



                if (Request["vbdYKienLanhDaoVanPhong"] != null && Request["vbdYKienLanhDaoVanPhong"].ToString().Length > 0)
                {
                    if (Request["vbdLanhDaoVanPhong"].ToString().Equals("0"))
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Chánh văn phòng chưa được chọn!";
                    }
                }

                if (Request["vbdLanhDaoPhuTrach"] != null)
                {
                    // add lanh dao vao danh sach xu ly
                    if (!Request["vbdLanhDaoPhuTrach"].ToString().Equals("0"))
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoPhuTrach"].ToString()));
                    }
                }
                if (Request["vbdLanhDaoVanPhong"] != null)
                {
                    /// add chanh van phong dao vao danh sach xu ly
                    if (!Request["vbdLanhDaoVanPhong"].ToString().Equals("0"))
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["vbdLanhDaoVanPhong"].ToString()));
                    }
                }
                //if(đơn vị của người dùng là ubnd tự động add lanhdaovp vào list vbdDSTGXuLy tham gia xử lý
                if (groupDA.CheckDonViIsUBND(currentUser.CurrentIDDonVi))
                {
                    //get danh sách user là lãnh đạo văn phòng
                    List<LUser> lstChanhVP = new LUserDA().GetUserChanhVanPhong();
                    foreach (LUser objUser in lstChanhVP)
                    {
                        vbdItem.vbdDSTGXuLy.Add(new SPFieldLookupValue(objUser.ID, objUser.Title));
                        vbdItem.vbdUserChuaXuLy.Add(new SPFieldLookupValue(objUser.ID, objUser.Title));

                    }
                    if (lstChanhVP.Count > 0)
                        vbdItem.vbdIsSentVanBan = true;
                }


                if (Request["intTaoHo"] != null) /// truong hop tao ho so
                {
                    if (Request["intchuyen"].ToString().Equals("1"))
                    {
                        if (Request["intphongban"] == null || string.IsNullOrEmpty(Request["intphongban"]))
                        {
                            objMsg.Erros = true;
                            objMsg.Message = "Chưa chọn phòng ban để chuyển văn bản!";
                        }

                        // chuyển giá trị phòng ban thành lookup value
                        vbdItem.vbdDSPBDVTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intphongban"].ToString()));

                    }
                    else
                    { //intchuyen = 0
                        if (Request["intcanbo"] == null || string.IsNullOrEmpty(Request["intcanbo"]))
                        {
                            objMsg.Erros = true;
                            objMsg.Message = "Chưa chọn cán bộ để chuyển văn bản!";
                        }
                        // chuyển giá trị cán bộ thành lookup value
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                    }
                }
                else
                { // truong hop ko tao ho so
                    if (Request["intchuyen"].ToString().Equals("1"))
                    {
                        vbdItem.vbdDSPBDVTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intphongban"].ToString()));
                        /// timf taif khoan dai dien de add vao van ban
                        foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                        {
                            SPFieldLookupValue daidien = groupDA.GetTaiKhoanDaiDien(item.LookupId);
                            if (daidien.LookupId > 0)
                            {
                                vbdItem.vbdDSTGXuLy.Add(daidien);
                                vbdItem.vbdUserChuaXuLy.Add(daidien);

                                // add danh sách người được gửi hiện thời: edit 0403
                                vbdItem.vbdCurrentUserReceived.Add(daidien);
                            }
                        }
                        if (vbdItem.vbdDSPBDVTGXuLy.Count > 0) /// neeus co phong ban thi da sent
                            vbdItem.vbdIsSentVanBan = true;
                    }
                    else
                    {
                        vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                        // add danh sách người được gửi hiện thời: edit 0403
                        vbdItem.vbdCurrentUserReceived.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));

                        vbdItem.vbdUserChuaXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                        if (vbdItem.vbdDSTGXuLy.Count > 0)   // neu co nguoi dung thi chua sent
                            vbdItem.vbdIsSentVanBan = true;
                    }
                }
                #endregion


                if (!objMsg.Erros)/// nếu không có lỗi xảy ra
                {

                    #region add for log: danh sách người được chuyển văn bản khi tạo mới
                    SPFieldLookupValueCollection dsNguoiLienQuan = vbdItem.vbdUserChuaXuLy;
                    #endregion

                    // thiet lap nguoi thuc hien thao tac hien thoi
                    vbdItem.vbdCurrentUserAction = new SPFieldLookupValue(currentUser.ID.ToString() + ";#");

                    string strThongBao = vbdDA.Add(vbdItem);

                    if (!strThongBao.Equals("Văn bản đã tồn tại trong hệ thống"))
                    {

                        string strCurrentListIdInfoSent = string.Empty;
                        if (Convert.ToInt32(strThongBao) > 0)
                        {
                            /// add log: thêm mới văn bản
                            addLog(EnumVanBanDen.strThemMoi + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ThemMoi, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), new Microsoft.SharePoint.SPFieldLookupValueCollection(""), false, false, true);
                            if (Request["intchuyen"].ToString().Equals("1"))
                            { /// dien thong tin gui van ban cho don vi
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoGroupNameReceived = item;
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = Convert.ToInt32(strThongBao);// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }

                                /// add log: chuyển văn bản
                                if (dsNguoiLienQuan.Count > 0)
                                {
                                    addLog(EnumVanBanDen.strChuyenNoiBo + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ChuyenNoiBo, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), dsNguoiLienQuan, true, true, true);
                                }
                            }
                            else
                            { /// dien thong tin gui van ban cho can bo
                                foreach (SPFieldLookupValue item in vbdItem.vbdDSTGXuLy)
                                {
                                    sentInforItem = new LThongTinGuiVanBanDen();
                                    sentInforItem.Title = strThongBao;
                                    sentInforItem.infoUserNameReceived = item;
                                    sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                                    sentInforItem.infoVanBanDenID.LookupId = Convert.ToInt32(strThongBao);// id van ban den duoc gui
                                    strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                                }

                                /// add log: chuyển văn bản
                                if (dsNguoiLienQuan.Count > 0)
                                {
                                    addLog(EnumVanBanDen.strChuyenNoiBo + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.ChuyenNoiBo, EnumDoiTuong.VanBanDen, Convert.ToInt32(strThongBao), dsNguoiLienQuan, true, true, true);
                                }
                            }
                            #region cập nhật thông tin vừa nhận: edit 0403
                            vbdDA.UpdateCurrentListIdInfoSent(strCurrentListIdInfoSent, Convert.ToInt32(strThongBao), 0);
                            #endregion
                            ///cập nhật trạng thái vào sổ văn bản cho văn bản qua email
                            emailItem.IsVaoSoVB = true;
                            emailItem.vanbandenlienquan.LookupId = Convert.ToInt32(strThongBao);
                            emailDA.UpdateTrangThai(emailItem);
                            objMsg.Erros = false;
                            objMsg.Message = string.Format("Đã thêm mới {0}: <b>{1}</b>", TitleWebpart, Server.HtmlEncode(vbdItem.vbdSoKyHieu));
                            objMsg.ID = "0";
                        }
                        else
                        {
                            objMsg.Erros = true;
                            objMsg.Message = strThongBao;
                        }
                    }
                    else
                    {
                        objMsg.Erros = true;
                        objMsg.Message = "Văn bản đã tồn tại trong hệ thống";
                    }
                }
            }
            catch (Exception ex)
            {
                objMsg.Erros = true;
                objMsg.Message = ex.ToString();
            }
            return objMsg;

        }

        /// <summary>
        /// Hàm xóa nhiều
        /// </summary>
        /// <param name="ltsID"></param>
        private Utils.objMessage Delete(List<int> ltsID)
        {
            if (!currentUser.ListPermission.XoaVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count == 0)
            {
                throw new Exception("Bạn không có quyền thực hiện thao tác này");
            }
            Utils.objMessage _objMesage = new Utils.objMessage();

            if (ltsID.Count > 0)
            {
                try
                {
                    string output = string.Empty;

                    foreach (int itemID in ltsID)
                    {
                        LVanBanDen tem = vbdDA.GetSoKyHieuById(itemID);
                        output += vbdDA.Delete(itemID);

                        /// add log
                        addLog(EnumVanBanDen.strXoa + " số: " + tem.vbdSoKyHieu, EnumVanBanDen.Xoa, EnumDoiTuong.VanBanDen, itemID, tem.vbdDSTGXuLy, false, false, false);
                    }


                    if (string.IsNullOrEmpty(output))
                    {
                        _objMesage.Erros = false;
                        if (ltsID.Count > 1)
                            _objMesage.Message = "Xóa thành công các bản ghi đã chọn";
                        else
                            _objMesage.Message = "Xóa thành công bản ghi đã chọn";
                    }
                    else
                    {
                        _objMesage.Erros = true;
                        _objMesage.Message = output;
                    }
                }
                catch (Exception ex)
                {
                    _objMesage.Erros = true;
                    _objMesage.Message = ex.Message + ex.StackTrace;
                }
            }
            else
            {
                _objMesage.Erros = true;
                _objMesage.Message = "Bạn chưa chọn bản ghi nào";
            }
            return _objMesage;
        }

        /// <summary>
        /// edit
        /// </summary>
        /// <returns>Thông báo</returns>
        private Utils.objMessage Edit()
        {
            if (!currentUser.ListPermission.SuaVanBanDen && currentUser.lstPhongBanLaVanThuVBDEN.Count == 0)
            {
                throw new Exception("Bạn không có quyền thực hiện thao tác này");
            }
            Utils.objMessage _objMesage = new Utils.objMessage();
            try
            {
                vbdItem = vbdDA.GetById(ItemID);
                vbdItem.ChangeObject(this.Request);


                // đổi tên file đính kèm trong trường hợp scan
                if (Page.Request["isscan"] == "true")
                {
                    if (!string.IsNullOrEmpty(Page.Request["listFileScan"]))
                    {
                        string[] listFileName = Page.Request["listFileScan"].Split('*');
                        int i = 1;
                        foreach (FileAttach fileAttach in vbdItem.ListFileAttachAdd)
                        {
                            if (listFileName.Contains(fileAttach.Name))
                            {
                                fileAttach.Name = vbdItem.vbdSoKyHieu + "-" + i.ToString() + ".tif";
                                // Đổi các ký tự đặc biệt thành '-'
                                string[] kytudacbiet = new string[] { "\\", "/", ":", "*", "?", "\"", "<", ">", "&", "%", "!", "^", "#", "$" };
                                foreach (string kytu in kytudacbiet)
                                {
                                    if (fileAttach.Name.Contains(kytu))
                                        fileAttach.Name = fileAttach.Name.Replace(kytu, "-");
                                }
                                i++;
                            }

                        }
                    }
                }

                /// bổ sung trường hợp chuyển văn bản khi sửa văn bản đến
                //if (Request["intchuyen"].ToString().Equals("1"))
                //{
                //    vbdItem.vbdDSPBDVTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intphongban"].ToString()));

                //    /// timf taif khoan dai dien de add vao van ban
                //    foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                //    {
                //        SPFieldLookupValue daidien = groupDA.GetTaiKhoanDaiDien(item.LookupId);
                //        if (daidien.LookupId > 0)
                //        {
                //            vbdItem.vbdDSTGXuLy.Add(daidien);
                //            vbdItem.vbdUserChuaXuLy.Add(daidien);

                //            // add danh sách người được gửi hiện thời: edit 0403
                //            vbdItem.vbdCurrentUserReceived.Add(daidien);
                //        }
                //    }
                //    if (vbdItem.vbdDSPBDVTGXuLy.Count > 0) /// neeus co phong ban thi da sent
                //        vbdItem.vbdIsSentVanBan = true;
                //}
                //else
                //{
                //    vbdItem.vbdDSTGXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                //    // add danh sách người được gửi hiện thời: edit 0403
                //    vbdItem.vbdCurrentUserReceived.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));

                //    vbdItem.vbdUserChuaXuLy.AddRange(Utils.SPUtils.StringToLookup(Request["intcanbo"].ToString()));
                //    if (vbdItem.vbdDSTGXuLy.Count > 0)   // neu co nguoi dung thi chua sent
                //        vbdItem.vbdIsSentVanBan = true;
                //}

                //// thiet lap nguoi thuc hien thao tac hien thoi
                //vbdItem.vbdCurrentUserAction = new SPFieldLookupValue(currentUser.ID.ToString() + ";#");

                if (vbdItem.vbdCoQuanBanHanh.Count == 0)
                {
                    _objMesage.Erros = true;
                    _objMesage.Message = "Cơ quan ban hành chưa được nhập";
                }

                if (!_objMesage.Erros)
                {
                    if (Request["vbdIsVanBanPhapQuy"] != null && Request["vbdIsVanBanPhapQuy"].ToString().Equals("true"))
                    {
                        vbdItem.vbdIsVanBanPhapQuy = true;
                    }
                    else
                    {
                        vbdItem.vbdIsVanBanPhapQuy = false;
                    }

                    string strThongBao = vbdDA.Update(vbdItem);
                    if (string.IsNullOrEmpty(strThongBao))
                    {
                        _objMesage.Erros = false;
                        _objMesage.Message = string.Format("Đã cập nhật {0}: <b>{1}</b>", TitleWebpart, vbdItem.vbdSoKyHieu);
                        _objMesage.ID = vbdItem.ID.ToString();

                        string strCurrentListIdInfoSent = string.Empty;

                        //if (Request["intchuyen"].ToString().Equals("1"))
                        //{ /// dien thong tin gui van ban cho don vi
                        //    foreach (SPFieldLookupValue item in vbdItem.vbdDSPBDVTGXuLy)
                        //    {
                        //        sentInforItem = new LThongTinGuiVanBanDen();
                        //        sentInforItem.Title = vbdItem.ID.ToString();
                        //        sentInforItem.infoGroupNameReceived = item;
                        //        sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                        //        sentInforItem.infoVanBanDenID.LookupId = vbdItem.ID;// id van ban den duoc gui
                        //        strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                        //    }


                        //}
                        //else
                        //{ /// dien thong tin gui van ban cho can bo
                        //    foreach (SPFieldLookupValue item in vbdItem.vbdDSTGXuLy)
                        //    {
                        //        sentInforItem = new LThongTinGuiVanBanDen();
                        //        sentInforItem.Title = vbdItem.ID.ToString();
                        //        sentInforItem.infoUserNameReceived = item;
                        //        sentInforItem.infoSentByUser.LookupId = currentUser.ID;
                        //        sentInforItem.infoVanBanDenID.LookupId = vbdItem.ID;// id van ban den duoc gui
                        //        strCurrentListIdInfoSent += sentInforDA.AddReturnID(sentInforItem) + ","; /// gán giá trị vừa được ghi lại
                        //    }

                        //}

                        //if (vbdItem.vbdIsSentVanBan)
                        //{
                        //    vbdDA.UpdateNguoiGiaoVB(vbdItem.ID, currentUser.ID, currentUser.userTenTruyCap); // cap nhat ngươi giao văn bản
                        //}

                        /// add log: sửa văn bản
                        addLog(EnumVanBanDen.strSua + " số: " + vbdItem.vbdSoKyHieu, EnumVanBanDen.Sua, EnumDoiTuong.VanBanDen, vbdItem.ID, vbdItem.vbdDSTGXuLy, false, false, false);
                    }
                    else
                    {
                        _objMesage.Erros = true;
                        _objMesage.Message = strThongBao;
                    }
                }
            }
            catch (Exception ex)
            {
                _objMesage.Erros = true;
                _objMesage.Message = ex.StackTrace;
            }
            return _objMesage;
        }
    }
}