using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace AmExEmailer
{
    public class EmailerData
    {
        public string UserName { get; set; }
        public string EmployeeName { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string ECardFile { get; set; }

    }

    public class ProcessedEmailerData : EmailerData
    {
        public int EmailStatus { get; set; }
        public string EmailerException { get; set; }
    }

    public class ClaySysEmailerData
    {
        public List<EmailerData> EmailerList { get; set; }
        public ClaySysEmailerData()
        {
            this.EmailerList = new List<EmailerData>();
            // Load from Excel
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string excelPath = $"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClaySysGMC.xlsx")}";
            DataSet Excelds = this.GetExcelDataSet(excelPath);
            DataTable dtClaySys = Excelds.Tables[0];
            foreach (DataRow drow in dtClaySys.Rows)
            {
                this.EmailerList.Add(new EmailerData()
                {
                    EmployeeName = Convert.ToString(drow[1])
                                                        ,
                    ECardFile = Convert.ToString(drow[2])
                                                        ,
                    Email = Convert.ToString(drow[4])
                                                        ,
                    Password = Convert.ToDateTime(drow[3]).ToString("dd-MM-yyyy")
                                                        ,
                    UserName = Convert.ToString(drow[0]) + "@Claysys"
                });
            }
        }

        internal DataSet GetExcelDataSet(string path)
        {
            DataSet ds = new DataSet();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    ds = reader.AsDataSet();
                    // The result of each spreadsheet is in result.Tables
                }
            }
            return ds;
        }
    }
}
