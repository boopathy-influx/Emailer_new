using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using System.Diagnostics;
using System.IO;


    public partial class Newmailer : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
         // Sentmail("harishanand@gmail.com", "Harish");
        //    Sentmail("aakrit@haptik.co", "Aakrit");
          //  Sentmail("hat@influx.co.in", "Harish");

          //Sentmail("mohandass_mn@apollohospitals.com", "Mohan Dass");
          //  Sentmail("harinath_m@apollohospitals.com", "Harinath");
          // Sentmail("akash.narang@influx.co.in", "Akash");
          
         // Sentmail("nagendraboopathy.murugan@influx.co.in", "Nagendra Boopathy");
       //     Sentmail("harishanand@gmail.com", "Harish Anand");
        //    Sentmail("mohandassmn@gmail.com", "Mr. Mohandass MN");


        //  Sentmail("giridharan.rangasamy@influx.co.in", " Giridharan");           
           BlastEmail();
        //    Sentmail("koliyotrohit@gmail.com", "Rohit");
          //  Sentmail("sunil@budfurn.com", "Uncle");
         ///   Sentmail("info@orbgen.com", "Ram");
           //Response.End();
     //   DataTable dt =    getEmailids();
           
        }
        void BlastEmail()
        {
            DataTable dt = getEmailids();
           int count = 0;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow r in dt.Rows)
                {
                    if (r["Email"].ToString() != "")
                    {
                         Sentmail(r["Email"].ToString(), r["Name"].ToString());
                        count++;
                        LogFile(" Mail sent to " + r["Email"].ToString() + "   Send count is : " + count  + Environment.NewLine);
                    }
                }
            }
        }
        private DataTable getEmailids()
        {

            string mFolder = System.IO.Path.GetDirectoryName(Server.MapPath("~/EmailersDB/FOH_Dometic_2014.csv"));

          //  string connStrings = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+mFolder+";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            string connStrings = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mFolder + ";Extended Properties='text;HDR=Yes'";
            // Create the connection object
            OleDbConnection con = new OleDbConnection();
            OleDbConnection oledbConn = new OleDbConnection(connStrings);            
            OleDbCommand cmd;
            oledbConn.Open();
            cmd = new OleDbCommand("SELECT [Prefix] + ' ' + [Full Name] as Name,[Email] FROM FOH_Dometic_2014.csv where [Email] <>'' ", oledbConn);
            // Response.Write("SELECT * FROM "+Request["selTheatre"]+".csv where screenid="+int.Parse(sclass[2].ToString()));
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            oleda.SelectCommand = cmd;
            // Create a DataSet which will hold the data extracted from the worksheet.
            DataSet ds = new DataSet();
            // Fill the DataSet from the data extracted from the worksheet.
            oleda.Fill(ds, "emailids");
            cmd.Dispose();
            oledbConn.Close();
            return ds.Tables[0];
        }
        void Sentmailold(string emailid,string name)
        {
            try
            {
                MailMessage mailMsg = new MailMessage();
                // From
               // mailMsg.From = new MailAddress("aakrit@haptik.co", "Aakrit Vaish");

                mailMsg.From = new MailAddress("hat@haptik.co", "Harish Anand Thilakan");
                // To
             //   mailMsg.To.Add(new MailAddress("hat@influx.co.in", "Harish"));

                mailMsg.To.Add(new MailAddress(emailid, name));                
                // Subject and multipart/alternative Body
                mailMsg.Subject = "Invitation: Beta access - Haptik";
                string html = string.Empty;
                string imagepath = "http://haptik.co/mailer/images/"; 
                html = html + "<html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' />";
                html = html + "<title>Haptik - Welcome e-mail</title>";
                html = html + "<style type='text/css'>body{margin:0;padding:0;font-family:'Rambla', sans-serif;}</style></head><body>";
                html = html + "<table width='600' border='0' cellspacing='0' cellpadding='0'>";
                html = html + "<tr>  <td height='20'>&nbsp;</td>  </tr>  <tr>    <td align='left' valign='top' height='97'>";
                html = html + "<table width='600' border='0' cellspacing='0' cellpadding='0' ";
                html = html + "<tr> <td align='left' valign='top'><img src='" + imagepath + "logo.gif' alt='' border=''/></td>";
                html = html + "<td align='left' valign='top'><img src='" + imagepath + "bg.gif' width='315' height='68' alt='' border='0'/></td>";
                html = html + "<td align='left' valign='top'><a href='http://www.haptik.co' target='_blank' title='haptik.co'><img src='" + imagepath + "link.gif' width='100' height='68' alt='' border='0'/></a></td>";
                html = html + "</tr>   </table>   </td>  </tr>   <tr><td>  <table width='600' border='0' cellspacing='0' cellpadding='0'>  <tr>";
                html = html + "<td align='left' valign='top' style='border-bottom:1px solid #b9dae8;'>";
                html = html + "<p style='padding-top:10px;padding-bottom:15px;color:#000;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Hi " + name + ",</p>";
                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>Hope all is well. </p>";
                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>The last few months have been an amazing ride working on Haptik, and today I am excited to extend a special invite to you to be a part of the private beta. In case we haven't talked about it yet, you can know more about Haptik <a href='http://www.haptik.co' style='color:#000;'>here</a>. You don't need to sign up on the site or do anything else, since we already got you down.</p>";
                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>I will send you an SMS in the coming week with a link to download the app. Watch out for that from my cell phone number. It will be available on both iOS and Android. </p>";
                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>We'd like to thank you for your support so far, and look forward to you being among the first set of users to use the service as we imagine it. We have received an overwhelming response on beta sign ups, and will gradually open up access to more users as we scale operations and development. </p>";
                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>Please do send across any and all feedback you have, as that will be critical in helping shape the next version of the product. You can email/call/message/tweet/whatsapp, whatever works for you :)</p>";
                //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Thanks,<br />Aakrit Vaish<br />Co-founder &amp; CEO</p>";

                html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Thanks,<br />Harish Anand Thilakan <br />Design &amp; Marketing Partner</p>";


                html = html + "</td> </tr> <tr><td align='right' valign='top' style='padding-top:12px;padding-bottom:30px;'>";
                html = html + "<table width='140' border='0' cellspacing='0' cellpadding='0'>";
                html = html + "<tr>";
                html = html + "<td align='left' valign='top' width='37'><a href='http://facebook.com/haptik' target='_blank' title='Facebook'><img src='" + imagepath + "facebook.png' width='31' height='31' alt='' border='0'/></a></td>";
                html = html + "<td align='left' valign='top' width='35'><a href='http://twitter.com/hellohaptik' target='_blank' title='Twitter'><img src='" + imagepath + "twitter.png' width='30' height='30' alt='' border='0'/></a></td>";
                html = html + "<td align='left' valign='top' width='37'><a href='http://www.youtube.com/channel/UCEeT8gTq3PKhbA-vB7k7GiQ' target='_blank' title='Youtube'><img src='" + imagepath + "youtube.png' width='32' height='31' alt='' border='0'/></a></td>";
                html = html + "<td align='left' valign='top'><a href='http://www.angel.co/haptik' target='_blank' title='a'><img src='" + imagepath + "a.png' width='31' height='31' alt='' border='0'/></a></td>";
                html = html + "</tr>";
                html = html + "</table>";
                html = html + "</td>";
                html = html + "</tr>";
                html = html + "</table>   ";
                html = html + "</td>";
                html = html + "</tr>";
                html = html + "</table>";
                html = html + "</body>";
                html = html + "</html>";

               mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(html, null, MediaTypeNames.Text.Html));
                // Init SmtpClient and send
                SmtpClient smtpClient = new SmtpClient("smtp.sendgrid.net", Convert.ToInt32(587));
                System.Net.NetworkCredential credentials = new System.Net.NetworkCredential("haptik1305", "Batman1305");
                smtpClient.Credentials = credentials;
               //smtpClient.Send(mailMsg);             
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
                using (StreamWriter writer = File.AppendText(path))
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

        void Sentmail(string emailid, string name)
        {
            try
            {
                MailMessage mailMsg = new MailMessage();
                // From      

                mailMsg.From = new MailAddress("contactus@future-of-healthcare.org", "The Healthcare Alliance");

                // To
                //   mailMsg.To.Add(new MailAddress("hat@influx.co.in", "Harish"));

                mailMsg.To.Add(new MailAddress(emailid, name));
                mailMsg.ReplyTo= new MailAddress("apollo@influx.co.in", "The Future of Healthcare Conference");              
               
                string Title = "<span style='font-family: Arial, Helvetica, sans-serif;  color:#231f20; font-size:18px; text-align:left;font-weight:bold;line-height:44px'>Dear " + name + ",</span><br />";




                // Subject and multipart/alternative Body
                mailMsg.Subject = "Invitation: The Future of Healthcare Conference";
                string html = string.Empty;
                string imagepath = "http://www.future-of-healthcare.org/emaillerimg/";
                html = html + "<html xmlns='http://www.w3.org/1999/xhtml'>";
                html = html + "<head>";
                html = html + "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />";
                html = html + "<title>The Future of Healthcare Conference - Mailer</title>";
                html = html + "</head>";
                html = html + "<body bgcolor='#FFFFFF' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>";
                html = html + "<table width='650' border='0' cellspacing='0' cellpadding='0'>";
                html = html + "<tr>";
                html = html + "<td><img src='" + imagepath + "/top.png' width='650' height='107' border='0' style='vertical-align:top' /></td>";
                html = html + "</tr>";
                html = html + "<tr>";
                html = html + "<td><img src='" + imagepath + "/alliances-logo.png' width='650' height='76' border='0' style='vertical-align:top' /></td>";
                html = html + "</tr>";
                html = html + "<tr>";
                html = html + "<td><img src='" + imagepath + "/banner.png' width='650' height='191' border='0' style='vertical-align:top' /></td>";
                html = html + "</tr>";
                html = html + "<tr>";
                html = html + "<td><img src='" + imagepath + "/together-hdg.png' width='650' height='30' border='0' style='vertical-align:top' /></td>";
                html = html + "</tr>";
                html = html + "<tr>";
                html = html + " <td style='font-family: Arial, Helvetica, sans-serif;  color:#231f20; font-size:12px; text-align:justify;line-height:24px;padding:5px 25px 10px'> "+ Title + "<strong>The Future of Healthcare Conference</strong> to be held on the <strong>3rd and 4th of March</strong> is an opportunity for change. A change the world is in serious need of. For we stand at a crucial juncture, where the world is facing unprecedented healthcare challenges. One billion people in the world still lack access to healthcare systems. Communicable and infectious diseases are growing at alarming rates and non-communicable diseases amount to almost two-thirds of all deaths worldwide. In fact, diseases kill more people each year than all conflicts and natural or man-made catastrophes put together. <br />";
                html = html + "Therefore, it's the choices we make today, that will determine the future of healthcare in our nations and the world at large. Let's take the right steps to make the world a healthier place for our future to thrive in. It is time for all important healthcare stakeholders to work together to find global solutions that are systemic, sustainable and impactful. It is time to look for a definitive way forward with an action plan to impact the future of healthcare. <br />";
                
                html = html + "<strong>The Healthcare Alliance</strong> is the coming together of an august set of organisations from important industry bodies to knowledge partners of global eminence, to facilitate deliberation over globally-relevant themes in healthcare. <br />";
                html = html + "In the first conference on <strong>The Future of Healthcare: A Collective Vision</strong>,six important topics will be discussed: <br />";
                html = html + "<ul>";
                html = html + "<li><strong>A framework for primary and preventive health</strong></li>";
                html = html + "<li><strong>Non-communicable diseases</strong> </li>";
                html = html + "<li><strong>New age health systems for innovation and universal access</strong></li>";
                html = html + "<li><strong>Talent fast-forward: Upskill and upscale</strong> </li>";
                html = html + "<li><strong>Public-private-people partnership to address the huge need for funding</strong></li>";
                html = html + "<li><strong>Hospitals of the future.</strong></li>";
                html = html + "</ul>";
                html = html + "The conference will provide innovative solutions and formulate guidelines for tackling the healthcare challenges faced by the world today. As someone who always inspires change and is at the forefront of healthcare development initiatives, we believe your knowledge and perspective would add tremendous depth to this global exchange of Ideas. <br />";
                html = html + "Your participation is vital to the success of this global initiative. <strong>So block your dates and join us at The Future of Healthcare: A Collective Vision</strong>. Together let us transform healthcare and secure the strength of our nations. </td>";
                html = html + "</tr>";          
                html = html + " <tr>";
                html = html + "<td valign='top'><table width='650' border='0' cellspacing='0' cellpadding='0'>";
                html = html + "<tr>";
                html = html + "<td width='327' height='60' align='left' valign='top' style='font-family: Arial, Helvetica, sans-serif;  color:#231f20; font-size:12px; text-align:justify;line-height:16px;padding:0 25px 0'>Thank you,<br />";
                html = html + "<strong>Team Healthcare Alliance</strong></td>";
                html = html + " <td align='left' valign='middle' style='padding-top:0px'><a href='http://www.future-of-healthcare.org/index.aspx' target='_blank'><img src='" + imagepath + "/btn-visit.png' alt='Visit Future-Of-Healthcare.org' width='247' height='32' border='0' style='vertical-align:top' /></a></td>";
                html = html + "</tr>      </table></td>  </tr>  <tr>";
                html = html + "<td><img src='" + imagepath + "/footer.png' width='650' height='121' border='0' style='vertical-align:top' /></td>  </tr>";

                html = html + "</table>";
                html = html + "</body>";
                html = html + "</html>";             

                mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(html, null, MediaTypeNames.Text.Html));
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
    }


    //html = html + "<html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' />";
    //html = html + "<title>Haptik - Welcome e-mail</title>";
    //html = html + "<style type='text/css'>body{margin:0;padding:0;font-family:'Rambla', sans-serif;}</style></head><body>";
    //html = html + "<table width='600' border='0' cellspacing='0' cellpadding='0'>";
    //html = html + "<tr>  <td height='20'>&nbsp;</td>  </tr>  <tr>    <td align='left' valign='top' height='97'>";
    //html = html + "<table width='600' border='0' cellspacing='0' cellpadding='0' ";
    //html = html + "<tr> <td align='left' valign='top'><img src='" + imagepath + "logo.gif' alt='' border=''/></td>";
    //html = html + "<td align='left' valign='top'><img src='" + imagepath + "bg.gif' width='315' height='68' alt='' border='0'/></td>";
    //html = html + "<td align='left' valign='top'><a href='http://www.haptik.co' target='_blank' title='haptik.co'><img src='" + imagepath + "link.gif' width='100' height='68' alt='' border='0'/></a></td>";
    //html = html + "</tr>   </table>   </td>  </tr>   <tr><td>  <table width='600' border='0' cellspacing='0' cellpadding='0'>  <tr>";
    //html = html + "<td align='left' valign='top' style='border-bottom:1px solid #b9dae8;'>";
    //html = html + "<p style='padding-top:10px;padding-bottom:15px;color:#000;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Hi " + name + ",</p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>Hope all is well. </p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>The last few months have been an amazing ride working on Haptik, and today I am excited to extend a special invite to you to be a part of the private beta. In case we haven't talked about it yet, you can know more about Haptik <a href='http://www.haptik.co' style='color:#000;'>here</a>. You don't need to sign up on the site or do anything else, since we already got you down.</p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>I will send you an SMS in the coming week with a link to download the app. Watch out for that from my cell phone number. It will be available on both iOS and Android. </p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>We'd like to thank you for your support so far, and look forward to you being among the first set of users to use the service as we imagine it. We have received an overwhelming response on beta sign ups, and will gradually open up access to more users as we scale operations and development. </p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;text-align:justify;padding-left:15px;padding-right:15px;'>Please do send across any and all feedback you have, as that will be critical in helping shape the next version of the product. You can email/call/message/tweet/whatsapp, whatever works for you :)</p>";
    ////html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Thanks,<br />Aakrit Vaish<br />Co-founder &amp; CEO</p>";
    //html = html + "<p style='padding-bottom:15px;color:#333333;font-family:Arial, Helvetica, sans-serif;font-size:14px;margin:0;padding-left:15px;padding-right:15px;'>Thanks,<br />Harish Anand Thilakan <br />Design &amp; Marketing Partner</p>";
    //html = html + "</td> </tr> <tr><td align='right' valign='top' style='padding-top:12px;padding-bottom:30px;'>";
    //html = html + "<table width='140' border='0' cellspacing='0' cellpadding='0'>";
    //html = html + "<tr>";
    //html = html + "<td align='left' valign='top' width='37'><a href='http://facebook.com/haptik' target='_blank' title='Facebook'><img src='" + imagepath + "facebook.png' width='31' height='31' alt='' border='0'/></a></td>";
    //html = html + "<td align='left' valign='top' width='35'><a href='http://twitter.com/hellohaptik' target='_blank' title='Twitter'><img src='" + imagepath + "twitter.png' width='30' height='30' alt='' border='0'/></a></td>";
    //html = html + "<td align='left' valign='top' width='37'><a href='http://www.youtube.com/channel/UCEeT8gTq3PKhbA-vB7k7GiQ' target='_blank' title='Youtube'><img src='" + imagepath + "youtube.png' width='32' height='31' alt='' border='0'/></a></td>";
    //html = html + "<td align='left' valign='top'><a href='http://www.angel.co/haptik' target='_blank' title='a'><img src='" + imagepath + "a.png' width='31' height='31' alt='' border='0'/></a></td>";
    //html = html + "</tr>";
    //html = html + "</table>";
    //html = html + "</td>";
    //html = html + "</tr>";
    //html = html + "</table>   ";
    //html = html + "</td>";
    //html = html + "</tr>";
    //html = html + "</table>";
    //html = html + "</body>";
    //html = html + "</html>";




