using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace YourNamespace
{
    class Program
    {
        static void Main(string[] args)
        {
            SendMail();
        }

        private static void SendMail()
        {
            try
            {
                var date = DateTime.Now.AddDays(-1).Date.ToString("dd/MM/yyyy");
                string connStr = "Data Source=178.128.217.191:1521/MALAYSIA;User ID=PPL1;Password=ppl;";
                DataTable dt = new DataTable();

                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    var query = $"SELECT COUNTRY, ACMP_NAME, ACCH_NAME AS ACCOUNT_HEAD_NAME, JDT,  DESD as DESCRIPTION,PAYMENT AS EXPENSE FROM PPL1.TT WHERE JDT = '{date}'";

                    OracleCommand cmd = new OracleCommand(query, conn);
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(dt);
                    //asdfasdf

                }

                var attachment = GetAttachment(dt);

                try
                {
                    var _email = "automation@mis.prangroup.com";
                    var _epass = "aaaaAAAA0000";

                    SmtpClient sc = new SmtpClient("mail.mis.prangroup.com");
                    sc.EnableSsl = false;
                    sc.Credentials = new NetworkCredential(_email, _epass);
                    sc.Port = 25;

                    MailMessage mail = new MailMessage();

                    //mail.To.Add("Samia@prangroup.com, mis3@prangroup.com, mis33@mis.prangroup.com");
                    mail.To.Add("mis33@mis.prangroup.com");
                    //mail.To.Add("automation18@mis.prangroup.com");
                    mail.Bcc.Add("mis33@mis.prangroup.com");
                    mail.From = new MailAddress("automation@mis.prangroup.com", "IBS (Financial Report)");
                    mail.Attachments.Add(attachment);
                    mail.Subject = $"LAST DAY EXPENSE | IBS | {date}";

                    StringBuilder sb = new StringBuilder();
                    sb.Append("<div style='font-family:calibri;font-size:15px;'>");
                    sb.Append("This is a computer generated mail. Please do not reply.</div>");

                    decimal totalExpense = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        totalExpense = dt.Compute("SUM(EXPENSE)", string.Empty) != DBNull.Value ? Convert.ToDecimal(dt.Compute("SUM(EXPENSE)", string.Empty)) : 0;
                    }

                    // Generate the HTML table with the total expense included
                    string htmlTable = ConvertDataTableToHTML(dt, totalExpense);

                    //string htmlTable = ConvertDataTableToHTML(dt);
                    sb.Append(htmlTable); // Append the generated HTML table to your email body

                    mail.Body = sb.ToString();
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.High;

                    sc.Send(mail);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                }


       

            }
            catch (Exception ex)
            {
            }
        }

        public static Attachment GetAttachment(DataTable dataTable)
        {
            MemoryStream outputStream = new MemoryStream();

            using (ExcelPackage package = new ExcelPackage(outputStream))
            {
                ExcelWorksheet facilityWorksheet = package.Workbook.Worksheets.Add("Sheet1");
                facilityWorksheet.Cells.LoadFromDataTable(dataTable, true);

                using (ExcelRange headerCells = facilityWorksheet.Cells[1, 1, 1, dataTable.Columns.Count])
                {
                    headerCells.Style.Font.Bold = true;
                }

                using (ExcelRange allCells = facilityWorksheet.Cells[1, 1, dataTable.Rows.Count + 1, dataTable.Columns.Count])
                {
                    allCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    allCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    allCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    allCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                facilityWorksheet.Cells.AutoFitColumns();

                package.Save();
            }



            outputStream.Position = 0;
            Attachment attachment = new Attachment(outputStream, $"IBS_FINANCIAL_REPORT_{DateTime.Now.AddDays(-1).Date.ToString("dd-MM-yyyy")}.xlsx", "application/vnd.ms-excel");

            return attachment;
        }

        private static string ConvertDataTableToHTML(DataTable dt)
        {
            var sb = new StringBuilder();
            sb.Append("<table border='1' cellspacing='0' cellpadding='5'>");
            sb.Append("<tr>");
            foreach (DataColumn column in dt.Columns)
            {
                sb.Append($"<th style='background-color: #D3D3D3;'>{column.ColumnName}</th>");
            }
            sb.Append("</tr>");

            foreach (DataRow row in dt.Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    sb.Append($"<td>{row[column]}</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            return sb.ToString();
        }

        private static string ConvertDataTableToHTML(DataTable dt, decimal totalExpense)
        {
            var sb = new StringBuilder();
            sb.Append("<table border='1' cellspacing='0' cellpadding='5'>");
            sb.Append("<tr>");
            foreach (DataColumn column in dt.Columns)
            {
                sb.Append($"<th style='background-color: #D3D3D3;'>{column.ColumnName}</th>");
            }
            sb.Append("</tr>");

            foreach (DataRow row in dt.Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    sb.Append($"<td>{row[column]}</td>");
                }
                sb.Append("</tr>");
            }

            // Add a footer row for the total expense
            sb.Append("<tr>");
            sb.Append($"<td colspan='{dt.Columns.Count - 1}' style='text-align: right; font-weight: bold;'>Total EXPENSE</td>");
            sb.Append($"<td>{totalExpense}</td>");
            sb.Append("</tr>");

            sb.Append("</table>");
            return sb.ToString();
        }

    }
}



