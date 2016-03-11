using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Net.Mail;

namespace EmployeeMonitoring
{
    class Program
    {
        static clsSQL clsSQL = new clsSQL();
        static string pathLog = System.Configuration.ConfigurationManager.AppSettings["PathLog"];
        static string pathTemp = System.Configuration.ConfigurationManager.AppSettings["PathTemp"];
        static string startDateIN = DateTime.Now.AddDays(double.Parse(System.Configuration.ConfigurationManager.AppSettings["DayBeforeIN"])).ToString("yyyy-MM-dd");
        static string endDateIN = DateTime.Now.AddDays(double.Parse(System.Configuration.ConfigurationManager.AppSettings["DayAfterIN"])).ToString("yyyy-MM-dd");
        static string startDateOUT = DateTime.Now.AddDays(double.Parse(System.Configuration.ConfigurationManager.AppSettings["DayBeforeOUT"])).ToString("yyyy-MM-dd");
        static string endDateOUT = DateTime.Now.AddDays(double.Parse(System.Configuration.ConfigurationManager.AppSettings["DayAfterOUT"])).ToString("yyyy-MM-dd");
        static string tempStaffInName = "StaffIn.csv";
        static string tempStaffOutName = "StaffOut.csv";
        static string siteShortName = System.Configuration.ConfigurationManager.AppSettings["siteShortName"];
        static int countStaffIn = 0;
        static int countStaffOut = 0;

        static void Main(string[] args)
        {
            Console.WriteLine("Processing : Get Staff IN");
            GetStaffIn();
            Console.WriteLine("Processing : Get Staff OUT");
            GetStaffOut();
            Console.WriteLine("Processing : SendMail");
            SendMail();
            Console.WriteLine("Processing : Delete Temp File");
            #region Delete Temp File
            if (File.Exists(clsGlobal.ExecutePathBuilder() + pathTemp+tempStaffInName))
            {
                try
                {
                    FileInfo fiIN = new FileInfo(clsGlobal.ExecutePathBuilder() + pathTemp + tempStaffInName);
                    fiIN.Delete();
                }
                catch (Exception ex) { Console.WriteLine("Error on Delete TempFile : " + ex.Message); }
            }
            if (File.Exists(clsGlobal.ExecutePathBuilder() + pathTemp + tempStaffOutName))
            {
                try
                {
                    FileInfo fiOut = new FileInfo(clsGlobal.ExecutePathBuilder() + pathTemp + tempStaffOutName);
                    fiOut.Delete();
                }
                catch (Exception ex) { Console.WriteLine("Error on Delete TempFile : " + ex.Message); }
            }
            #endregion
            Console.WriteLine("Complete");
        }

        static bool SendMail()
        {
            string mail_to = System.Configuration.ConfigurationManager.AppSettings["MailTo"];
            string mail_from = "AutoSystem@glsict.com";
            string mail_from_aliasname = siteShortName+" Employee Monitoring";
            string mail_cc = System.Configuration.ConfigurationManager.AppSettings["MailCc"];
            string mail_bcc = System.Configuration.ConfigurationManager.AppSettings["MailBcc"];
            string mail_subject = siteShortName+" Employee Monitoring : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            string mail_body = "<h2>รายงานสรุปข้อมูลพนักงาน เริ่มงาน และ ลาออก</h2>"+
                "<div>"+
                "<b>เริ่มงาน</b>: "+countStaffIn.ToString()+" คน <span style='color:#30C9AA;'>(" + startDateIN + " ถึง " + endDateIN + ")</span>" +
                "</div><div>"+
                "<b>ลาออก</b>: "+countStaffOut.ToString()+" คน <span style='color:#C9303E;'>(" + startDateOUT + " ถึง " + endDateOUT + ")</span>" +
                "</div>";

            bool rtnValue = false;
            bool useAuthen = false;
            string pathFullStaffIn = clsGlobal.ExecutePathBuilder() + pathTemp + tempStaffInName;
            string pathFullStaffOut = clsGlobal.ExecutePathBuilder() + pathTemp + tempStaffOutName;

            if (!File.Exists(pathFullStaffIn))
            {
                pathFullStaffIn = "";
            }
            if (!File.Exists(pathFullStaffOut))
            {
                pathFullStaffOut = "";
            }

            if (string.IsNullOrEmpty(mail_to))  //ถ้าไม่ระบุ mail_to ให้คืนค่าและออกจากโปรแกรม
            {
                return false;
            }

            #region Send Mail
            string smtpMail_Host = System.Configuration.ConfigurationManager.AppSettings["SmtpServer"];
            string smtpMail_Username = "smtpMail_Username";
            string smtpMail_Password = "smtpMail_Password";

            MailMessage myMail = new MailMessage();
            StringBuilder strMailBody = new StringBuilder();

            //########### Mail From ###########
            if (mail_from_aliasname != "")
            {
                myMail.From = new MailAddress(mail_from, mail_from_aliasname);
            }
            else
            {
                myMail.From = new MailAddress(mail_from);
            }

            //########### Mail To ###########
            #region Mail To
            if (!string.IsNullOrEmpty(mail_to))
            {
                List<string> lstMailTo = new List<string>();
                lstMailTo = mail_to.Split(',').ToList();

                for (int i = 0; i < lstMailTo.Count(); i++)
                {
                    myMail.To.Add(lstMailTo[i]);
                }
            }
            #endregion

            //########### Mail Cc ###########
            #region Mail Cc
            if (!string.IsNullOrEmpty(mail_cc))
            {
                List<string> lstMailCc = new List<string>();
                lstMailCc = mail_cc.Split(',').ToList();

                for (int i = 0; i < lstMailCc.Count; i++)
                {
                    myMail.CC.Add(lstMailCc[i]);
                }
            }
            #endregion

            //########### Mail Bcc ###########
            #region Mail Bcc
            if (!string.IsNullOrEmpty(mail_bcc))
            {
                List<string> lstMailBcc = new List<string>();
                lstMailBcc = mail_bcc.Split(',').ToList();

                for (int i = 0; i < lstMailBcc.Count; i++)
                {
                    myMail.Bcc.Add(lstMailBcc[i]);
                }
            }
            #endregion

            //########### Mail Subject ###########
            if (!string.IsNullOrEmpty(mail_subject))
            {
                myMail.Subject = mail_subject;
            }
            else
            {
                myMail.Subject = "";
            }

            //########### Mail Body ###########
            strMailBody.Append("<html>");
            strMailBody.Append("<head></head>");
            strMailBody.Append("<body>");
            if (!string.IsNullOrEmpty(mail_body))
            {
                strMailBody.Append(mail_body);
            }
            strMailBody.Append("<div>จากเซิฟเวอร์ : " + clsGlobal.IPAddressBuilder() + "</div>");
            strMailBody.Append(@"<div>Log Path : \\" + clsGlobal.IPAddressBuilder() + @"\Application\EmployeeMonitoring\" + pathLog.Replace("/",@"\") + "</div>");

            strMailBody.Append("<hr/>");
            strMailBody.Append("<b>ServerIP</b> : " + clsGlobal.IPAddressBuilder() + "<br/>");
            strMailBody.Append("<b>ExecutablePath</b> : " + clsGlobal.ExecutePathBuilder() + "</br>");
            strMailBody.Append("<b>Version</b> : " + clsGlobal.VersionBuilder() + "<br/>");
            strMailBody.Append("<b>LastUpdate</b> : " + clsGlobal.LastUpdateBuilder().ToString("dd/MM/yyyy HH:mm") + "");

            strMailBody.Append("</body>");
            strMailBody.Append("</html>");

            myMail.Body = strMailBody.ToString();

            myMail.IsBodyHtml = true;
            myMail.BodyEncoding = System.Text.Encoding.GetEncoding("windows-874");

            //########### Mail Attachment ###########

            if (!string.IsNullOrEmpty(pathFullStaffIn))
            {
                Attachment attachFile1 = new Attachment(pathFullStaffIn);
                myMail.Attachments.Add(attachFile1);
                //attachFile1.Dispose();
            }
            if (!string.IsNullOrEmpty(pathFullStaffOut))
            {
                Attachment attachFile2 = new Attachment(pathFullStaffOut);
                myMail.Attachments.Add(attachFile2);
                //attachFile2.Dispose();
            }

            SmtpClient smtpMail = new SmtpClient();
            smtpMail.Host = smtpMail_Host;
            if (useAuthen)
            {
                System.Net.NetworkCredential Auth = new System.Net.NetworkCredential(smtpMail_Username, smtpMail_Password);
                smtpMail.UseDefaultCredentials = false;
                smtpMail.Credentials = Auth;
            }
            else
            {
                smtpMail.UseDefaultCredentials = true;
            }

            try
            {
                smtpMail.Send(myMail);
                rtnValue = true;
                ExportLog("SendMail", "Complete");
            }
            catch (Exception ex)
            {
                rtnValue = false;
                ExportLog("SendMail", "Error: "+ex.Message);
            }
            #endregion

            return rtnValue;
        }

        static bool GetStaffIn()
        {
            string outMessage;
            bool rtnValue = false;
            StringBuilder strSQL = new StringBuilder();
            DataTable dt = new DataTable();
            outMessage = "";

            #region SQL Query
            strSQL.Append("SELECT ");
            strSQL.Append("* ");
            strSQL.Append("FROM ");
            strSQL.Append("STAFF ");
            strSQL.Append("WHERE ");
            strSQL.Append("st_startdate >= '"+startDateIN+"' ");
            strSQL.Append("AND st_startdate <= '"+endDateIN+"' ");
            #endregion

            dt = clsSQL.Bind(strSQL.ToString(), "SQL", "cs");
            strSQL.Length = 0; strSQL.Capacity = 0;

            if (dt != null && dt.Rows.Count > 0)
            {
                countStaffIn = dt.Rows.Count;
                outMessage = "พบพนักงานที่เริ่มงานระหว่างวันที่ " +
                    startDateIN +
                    " ถึง " +
                    endDateIN + 
                    " จำนวน "+dt.Rows.Count.ToString()+" คน";

                ExportLog("GetStaffIn", outMessage);

                if (ExportTemp(dt, tempStaffInName))
                {
                    ExportLog("ExportStaffIn", "Complete");
                    rtnValue = true;
                }
                else
                {
                    ExportLog("ExportStaffIn", "Error");
                }
            }
            else
            {
                outMessage = "ไม่พบข้อมูลพนักงานที่เริ่มงานระหว่างวันที่ " + 
                    startDateIN + 
                    " ถึง " + 
                    endDateIN;

                ExportLog("GetStaffIn", outMessage);
            }

            return rtnValue;
        }

        static bool GetStaffOut()
        {
            string outMessage;
            bool rtnValue = false;
            StringBuilder strSQL = new StringBuilder();
            DataTable dt = new DataTable();
            outMessage = "";

            #region SQL Query
            strSQL.Append("SELECT ");
            strSQL.Append("* ");
            strSQL.Append("FROM ");
            strSQL.Append("STAFF ");
            strSQL.Append("WHERE ");
            strSQL.Append("st_enddate >= '" + startDateOUT + "' ");
            strSQL.Append("AND st_enddate <= '" + endDateOUT + "' ");
            #endregion

            dt = clsSQL.Bind(strSQL.ToString(), "SQL", "cs");
            strSQL.Length = 0; strSQL.Capacity = 0;

            if (dt != null && dt.Rows.Count > 0)
            {
                countStaffOut = dt.Rows.Count;
                outMessage = "พบพนักงานที่ลาออกระหว่างวันที่ " +
                    startDateOUT +
                    " ถึง " +
                    endDateOUT +
                    " จำนวน " + dt.Rows.Count.ToString() + " คน";

                ExportLog("GetStaffOut", outMessage);

                rtnValue = true;

                if (ExportTemp(dt, tempStaffOutName))
                {
                    ExportLog("ExportStaffOut", "Complete");
                    rtnValue = true;
                }
                else
                {
                    ExportLog("ExportStaffOut", "Error");
                }
            }
            else
            {
                outMessage = "ไม่พบข้อมูลพนักงานที่ลาออกระหว่างวันที่ " +
                    startDateOUT +
                    " ถึง " +
                    endDateOUT;

                ExportLog("GetStaffOut", outMessage);
            }

            return rtnValue;
        }

        static private bool ExportTemp(DataTable dt,string fileName)
        {
            bool rtnBool = false;
            string fullFileName = clsGlobal.ExecutePathBuilder() + pathTemp + fileName;
            FileInfo fiLog = new FileInfo(fullFileName);
            StringBuilder strOutput = new StringBuilder();
            try
            {
                #region Delete Temp File
                if (File.Exists(fullFileName))
                {
                    FileInfo fiIN = new FileInfo(fullFileName);
                    fiIN.Delete();
                }
                #endregion

                #region Add Header
                if (!fiLog.Exists)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (i == 0)
                        {
                            strOutput.Append(dt.Columns[i].ColumnName);
                        }
                        else
                        {
                            strOutput.Append("," + dt.Columns[i].ColumnName);
                        }
                    }
                    using (StreamWriter sw1 = new StreamWriter(fullFileName, true, System.Text.Encoding.UTF8))
                    {
                        sw1.WriteLine(
                            strOutput.ToString()
                        );
                    }
                }
                #endregion
                #region Add Value
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    strOutput.Length = 0; strOutput.Capacity = 0;

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j > 0)
                        {
                            strOutput.Append(",");
                        }
                        strOutput.Append(dt.Rows[i][j].ToString());
                    }

                    using (StreamWriter sw1 = new StreamWriter(fullFileName, true, System.Text.Encoding.UTF8))
                    {
                        sw1.WriteLine(
                            strOutput.ToString()
                        );
                    }
                }
                rtnBool = true;
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                rtnBool = false;
            }

            return rtnBool;
        }

        static private void ExportLog(string MethodName,string Message)
        {
            try
            {
                string pathLogTemp = clsGlobal.ExecutePathBuilder()+pathLog;
                FileInfo fiLog = new FileInfo(pathLogTemp);
                if (!fiLog.Exists)
                {
                    using (StreamWriter sw1 = new StreamWriter(pathLogTemp, true, System.Text.Encoding.UTF8))
                    {
                        sw1.WriteLine(
                            "MonitorDate" + "," +
                            "MethodName" + "," +
                            "Message"
                        );
                    }
                }

                StreamWriter sw = new StreamWriter(pathLogTemp, true, System.Text.Encoding.UTF8);

                sw.WriteLine(
                    DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "," +
                    MethodName + "," +
                    Message
                );

                sw.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                System.Threading.Thread.Sleep(10000);
            }
        }
    }
}
