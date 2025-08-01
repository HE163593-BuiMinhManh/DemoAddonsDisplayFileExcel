using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Xml;

using System.Data;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;


namespace DemoAddonsDisplayFileExcel
{
    [FormAttribute("DemoAddonsDisplayFileExcel.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        private SAPbouiCOM.Application oApp;
        private SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Company oCurrentDICompany;
        public static SAPbobsCOM.CompanyService oCurrentServiceCompany;

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_4").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        /// 
        // day la override cho class UserFormBase
        public override void OnInitializeFormEvents()
        {
        }

        public static SAPbobsCOM.Company GetCurrentDICompany()
        {
            if (oCurrentDICompany == null)
            {
                oCurrentDICompany = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                oCurrentServiceCompany = oCurrentDICompany.GetCompanyService();
            }

            return oCurrentDICompany;
        }

        private void OnCustomInitialize()
        {
            //Lấy đối tượng ứng dụng SAP B1 hiện tại (UI API).
            oApp = SAPbouiCOM.Framework.Application.SBO_Application;


            this.oCompany = GetCurrentDICompany();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;


        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            string file = Common.openFileDialog("Select Excel File", "Excel Files|*.xlsx");
            EditText0.Value = file;

        }


        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            // kiem tra file
            
            BubbleEvent = true;
            string filePath = EditText0.Value;

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                //Application.SBO_Application.MessageBox("Vui lòng chọn file Excel hợp lệ."); 

                oApp.SetStatusBarMessage("Vui lòng chọn file Excel hợp lệ.");
                oApp.MessageBox("Vui lòng chọn file Excel hợp lệ.");
              
                return;
            }


            try
            {
                // dky bo ma hoa de doc file excel
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


                // doc file
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))//
                    {
                        var result = reader.AsDataSet();
                        DataTable dt = result.Tables[0]; // lấy sheet đầu tiên
                        Grid0.DataTable = null;
                        // Tạo DataTable SAP để đổ dữ liệu vào Grid
                        SAPbouiCOM.DataTable sapDT = this.UIAPIRawForm.DataSources.DataTables.Item("DT_0");//


                        // Xóa toàn bộ dòng
                        while (sapDT.Rows.Count > 0)
                            sapDT.Rows.Remove(0);

                        // Xóa toàn bộ cột
                        while (sapDT.Columns.Count > 0)
                            sapDT.Columns.Remove(sapDT.Columns.Item(0).Name);

                        // sua header

                        // Lấy header từ dòng đầu tiên
                        List<string> columnNames = new List<string>();
                        HashSet<string> usedNames = new HashSet<string>();

                        for (int col = 0; col < dt.Columns.Count; col++)
                        {
                            string baseName = dt.Rows[0][col]?.ToString().Trim();
                            if (string.IsNullOrEmpty(baseName))
                                baseName = $"Column{col}";

                            // Đảm bảo tên không trùng
                            string uniqueName = baseName;
                            int count = 1;
                            while (usedNames.Contains(uniqueName))
                            {
                                uniqueName = baseName + "_" + count++;
                            }

                            usedNames.Add(uniqueName);
                            columnNames.Add(uniqueName);

                            sapDT.Columns.Add(uniqueName, SAPbouiCOM.BoFieldsType.ft_Text);
                        }

                        // Gán dữ liệu từ dòng thứ 2 trở đi
                        for (int row = 1; row < dt.Rows.Count; row++)
                        {
                            sapDT.Rows.Add();
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                string value = dt.Rows[row][col]?.ToString() ?? "";
                                sapDT.SetValue(columnNames[col], row - 1, value);
                            }
                        }

                        // Gán DataTable vào Grid
                        Grid0.DataTable = sapDT;

                        // Gán caption cột Grid
                        for (int col = 0; col < columnNames.Count; col++)
                        {
                            Grid0.Columns.Item(col).TitleObject.Caption = columnNames[col];
                        }

                        Grid0.AutoResizeColumns();



                        //// Thêm cột
                        //for (int col = 0; col < dt.Columns.Count; col++)
                        //{
                        //    sapDT.Columns.Add($"Col{col}", SAPbouiCOM.BoFieldsType.ft_Text);
                        //}

                        //// Thêm dữ liệu từng dòng

                        //for (int row = 0; row < dt.Rows.Count; row++)
                        //{
                        //    sapDT.Rows.Add();
                        //    for (int col = 0; col < dt.Columns.Count; col++)
                        //    {
                        //        string value = dt.Rows[row][col]?.ToString() ?? "";
                        //        sapDT.SetValue($"Col{col}", row, value);
                        //    }
                        //}

                        // Gán dữ liệu vào Grid
                        //Grid0.DataTable = sapDT;
                        //Grid0.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.MessageBox("Lỗi khi đọc file Excel: " + ex.Message);

                oApp.MessageBox("Lỗi khi đọc file Excel: " + ex.Message);
                BubbleEvent = false;

            }

            //BubbleEvent = true;

        }


    }
}