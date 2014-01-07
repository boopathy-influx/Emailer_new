using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System;
using System.Diagnostics;
using System.IO;
using System.Web;
namespace Emailer
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            //Sentmail("hat@influx.co.in");  
           BlastEmail();
           //Response.End();
        }


        void BlastEmail()
        {
           // DataTable dt = getEmailids();
           int count = 0;
           // if (dt.Rows.Count > 0)
           // {

               // foreach (DataRow r in dt.Rows)
               // {
           //harishanand@hotmail.com
          //string[] spiltemail = "nagendraboopathy.murugan@influx.co.in;aakrit.vaish@gmail.com;harishanand@gmail.com;harishanand@hotmail.com;swapan.rajdev@gmail.com;miten.sampat@gmail.com;aakrit@haptik.co;hat@haptik.co".ToString().Split(';');
          // string[] spiltemail = "harishanand@hotmail.com;hat@haptik.co;nagendraboopathy.murugan@influx.co.in".Split(';');
           string[] spiltemail = "nagendraboopathy.murugan@influx.co.in".Split(';');
                    if (spiltemail.Length > 0)
                    {
                        for (int i = 0; i < spiltemail.Length; i++)
                        {
                            if (spiltemail[i].ToString() != "")
                            {
                                Sentmail(spiltemail[i]);
                                count++;
                                LogFile("Last emailid is  : " + spiltemail[i] + "   Send mail count is : " + count);
                            }
                        }
                    }
              //  }
          //  }
        }
        private DataTable getEmailids()
        {


            string mFolder = System.IO.Path.GetDirectoryName(Server.MapPath("~/bangaloredatabase.csv"));
            string connStrings = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mFolder + ";Extended Properties='text;HDR=Yes'";

            // Create the connection object

            OleDbConnection con = new OleDbConnection();
            OleDbConnection oledbConn = new OleDbConnection(connStrings);
            
            OleDbCommand cmd;
            oledbConn.Open();

            cmd = new OleDbCommand("SELECT IDS FROM bangaloredatabase.csv", oledbConn);
            // Response.Write("SELECT * FROM "+Request["selTheatre"]+".csv where screenid="+int.Parse(sclass[2].ToString()));

            OleDbDataAdapter oleda = new OleDbDataAdapter();

            oleda.SelectCommand = cmd;

            // Create a DataSet which will hold the data extracted from the worksheet.
            DataSet ds = new DataSet();

            // Fill the DataSet from the data extracted from the worksheet.
            oleda.Fill(ds, "emailids");
            return ds.Tables[0];

        }


        void Sentmail(string emailid)
        {
            try
            {
                MailMessage mailMsg = new MailMessage();
                // From
                mailMsg.From = new MailAddress("hat@haptik.co", "haptik");
                // To
             //   mailMsg.To.Add(new MailAddress("hat@influx.co.in", "Harish"));

                mailMsg.To.Add(new MailAddress(emailid, ""));
                
                // Subject and multipart/alternative Body
                mailMsg.Subject = " Invitation: Beta access - Haptik - on behalf of Aakrit Vaish";
               // string text = "text body";
                string html = string.Empty;
                //html = html + "<html> ";
                //html = html + "<head>";
                //html = html + "<title>Vote Maadi - Mailer</title>";
                //html = html + "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'> </head>";
                //html = html + "<body bgcolor='#FFFFFF' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'> ";
                //html = html + "<table width='650' border='0' cellpadding='0' cellspacing='0'>";
                //html = html + "<tr>";
                //html = html + "<td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_01.jpg' width='650' height='50' border='0' usemap='#Map2' style='vertical-align:top'></td>";
                //html = html + "</tr> <tr>";
                //html = html + "<td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_02.jpg' width='650' height='137' border='0' style='vertical-align:top'></td>";
                //html = html + "</tr>";
                //html = html + "<tr>";
                //html = html + "<td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_03.jpg' width='650' height='181' border='0' style='vertical-align:top'></td>";
                //html = html + "</tr>";
                //html = html + "<tr>";
                //html = html + " <td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_04.jpg' width='650' height='202' border='0' usemap='#Map4' style='vertical-align:top'></td>";
                //html = html + " </tr>";
                //html = html + "<tr>";
                //html = html + " <td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_05.jpg' width='650' height='113' border='0' usemap='#Map3' style='vertical-align:top'></td>";
                //html = html + " </tr>";
                //html = html + "<tr>";
                //html = html + " <td><img src='http://www.votemaadi.com/VoteMaadi-mailer/images/index_06.png' width='650' height='126' border='0' usemap='#Map' style='vertical-align:top'></td>";
                //html = html + "</tr>";
                //html = html + "</tr> ";
                //html = html + "<tr>";
                //html = html + "<td style=\"font-family: 'Open Sans', sans-serif;color:#000;font-size:14px;font-weight:300;text-align:left;padding:10px 0 10px 25px\" >Use your discretion. Go out & #VoteMaadi | 05/05 - Karnataka Votes.</td>";
                //html = html + "</tr>";


                //html = html + "</table>";
                //html = html + "<map name='Map'>";
                //html = html + "<area shape='rect' coords='404,5,629,30' href='https://www.facebook.com/VoteMaadi2013' target='_blank'>";
                //html = html + " <area shape='rect' coords='22,86,157,109' href='http://www.votemaadi.com' target='_blank'>";
                //html = html + " <area shape='rect' coords='21,5,257,47' href='https://twitter.com/VoteMaadi2013' target='_blank'>";
                //html = html + "</map>";
                //html = html + "<map name='Map2'>";
                //html = html + " <area shape='rect' coords='445,15,631,47' href='http://timesofindia.indiatimes.com/' target='_blank'>";
                //html = html + "</map>";
                //html = html + "<map name='Map3'>";
                //html = html + " <area shape='rect' coords='461,26,611,86' href='https://www.facebook.com/VoteMaadi2013' target='_blank'>";
                //html = html + "</map>";
                //html = html + "<map name='Map4'>";
                //html = html + " <area shape='rect' coords='134,155,299,197' href='https://www.facebook.com/VoteMaadi2013' target='_blank'>";
              //  html = html + "</map></body></html>";

               // mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(text, null, MediaTypeNames.Text.Plain));

                html = html + "This is a test email from Haptik";



                mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(html, null, MediaTypeNames.Text.Plain));

                // Init SmtpClient and send
                SmtpClient smtpClient = new SmtpClient("smtp.sendgrid.net", Convert.ToInt32(587));
                System.Net.NetworkCredential credentials = new System.Net.NetworkCredential("influxint", "dzine*586");
                smtpClient.Credentials = credentials;

                smtpClient.Send(mailMsg);
             
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void LogFile(string msg)
        {
       
            try
            {

                string path = HttpContext.Current.Server.MapPath("~/Count/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                path = path + DateTime.Today.ToString("dd-MMM-yy") + ".log";
                if (!File.Exists(path))
                {
                    File.Create(path).Dispose();
                }
                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.WriteLine(msg);                  
                    writer.Flush();
                    writer.Close();
                }
               
            }
            catch
            {
                throw;
            }
           
        }

    }
}



