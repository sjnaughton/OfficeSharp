using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace OfficeSharp.SMTP
{
    public static class SMTP
    {

        public static bool CreateEmail(string SendFrom, string SendTo, string Subject, string Body, SmtpServer Server)
        {

            MailMessage mail = new MailMessage();

            try
            {
                var _with1 = mail;
                _with1.From = new MailAddress(SendFrom);
                _with1.To.Add(new MailAddress(SendTo));
                _with1.Subject = Subject;

                if (Body.ToLower().Contains("<html>"))
                {
                    _with1.IsBodyHtml = true;
                }

                _with1.Body = Body;

                SmtpClient smtp = new SmtpClient(Server.SmtpServerName, Server.SmtpPort);
                smtp.Credentials = new NetworkCredential(Server.SmtpUserId, Server.SmtpPassword);
                smtp.Send(mail);

                return true;

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create Email.", ex);
            }

            return false;
        }

        public static bool CreateEmail(string SendFrom, string SendTo, string Subject, IEnumerable Body, SmtpServer Server)
        {

            MailMessage mail = new MailMessage();
            string sbody = null;

            // opening an html tag and creating a table 
            sbody = "<html><body style=\"font-family: Arial, Helvetica, sans-serif;\" ><table border=\"1\">";

            try
            {
                var _with2 = mail;
                _with2.From = new MailAddress(SendFrom);
                _with2.To.Add(new MailAddress(SendTo));
                _with2.Subject = Subject;

                if (sbody.ToLower().Contains("<html>"))
                {
                    _with2.IsBodyHtml = true;

                    // object row count
                    int iRowCount = 0;
                    // to extract attributes using reflection  
                    foreach (object rowValue in Body)
                    {
                        // extracting the header value [once] 
                        if (iRowCount == 0)
                        {
                            // creating new row for the header
                            sbody +=  "<tr>";
                            foreach (System.Reflection.PropertyInfo properties in rowValue.GetType().GetProperties())
                            {
                                // omitting the Detail and the Id reserved words from the header
                                if (properties.Name != "Details" & properties.Name != "Id")
                                {
                                    // creating new cell
                                    sbody += "<td>";
                                    sbody += " " + properties.Name;
                                    sbody += "</td>";
                                }
                            }
                            // row ends
                            sbody +=  "</tr>";
                        }

                        // creating a new row for the attribute values
                        sbody +=  "<tr>";
                        // extracting table attribute values 
                        foreach (System.Reflection.PropertyInfo properties in rowValue.GetType().GetProperties())
                        {
                            // omitting the ID & Details reserved words for attribute values
                            if (properties.Name != "Id" & properties.Name != "Details")
                            {
                                // creating new table cell
                                sbody += "<td>";
                                sbody += " " + properties.GetValue(Body, new Object[] { iRowCount }).ToString();
                                sbody += "</td>";
                            }
                        }
                        // row ends
                        sbody +=  "</tr>";
                        iRowCount += 1;
                    }

                    // closing the tags
                    sbody +=  "</table></body></html>";
                }

                _with2.Body = sbody;

                SmtpClient smtp = new SmtpClient(Server.SmtpServerName, Server.SmtpPort);
                smtp.Credentials = new NetworkCredential(Server.SmtpUserId, Server.SmtpPassword);
                smtp.Send(mail);

                return true;

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create Email.", ex);
            }

            return false;
        }

        public static bool CreateAppointment(string SendFrom, string SendTo, string Subject, string Body, string Location, System.DateTime StartTime, System.DateTime EndTime, string MsgID, int Sequence, bool IsCancelled,
        SmtpServer Server)
        {
            dynamic result = false;
            try
            {
                if (string.IsNullOrEmpty(SendTo) || string.IsNullOrEmpty(SendFrom))
                {
                    throw new InvalidOperationException("SendFrom and SendTo email addresses must be specified.");
                }

                dynamic fromAddress = new MailAddress(SendFrom);
                dynamic toAddress = new MailAddress(SendTo);
                MailMessage mail = new MailMessage();

                var _with3 = mail;
                _with3.Subject = Subject;
                _with3.From = fromAddress;

                //Need to send to both parties to organize the meeting
                _with3.To.Add(toAddress);
                _with3.To.Add(fromAddress);

                //Use the text/calendar content type 
                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
                ct.Parameters.Add("method", "REQUEST");
                //Create the iCalendar format and add it to the mail
                dynamic cal = CreateICal(SendFrom, SendTo, Subject, Body, Location, StartTime, EndTime, MsgID, Sequence, IsCancelled);
                mail.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(cal, ct));

                //Send the meeting request
                SmtpClient smtp = new SmtpClient(Server.SmtpServerName, Server.SmtpPort);
                smtp.Credentials = new NetworkCredential(Server.SmtpUserId, Server.SmtpPassword);
                smtp.Send(mail);

                result = true;

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to send Appointment.", ex);
            }

            return result;
        }

        private static string CreateICal(string SendFrom, string SendTo, string Subject, string Body, string Location, System.DateTime StartTime, System.DateTime EndTime, string MsgID, int Sequence, bool IsCancelled)
        {

            StringBuilder sb = new StringBuilder();
            if (string.IsNullOrEmpty(MsgID))
            {
                MsgID = Guid.NewGuid().ToString();
            }

            //See iCalendar spec here: http://tools.ietf.org/html/rfc2445
            //Abridged version here: http://www.kanzaki.com/docs/ical/
            sb.AppendLine("BEGIN:VCALENDAR");
            sb.AppendLine("PRODID:-//Microsoft LightSwitch");
            sb.AppendLine("VERSION:2.0");
            if (IsCancelled)
                sb.AppendLine("METHOD:CANCEL");
            else
                sb.AppendLine("METHOD:REQUEST");

            sb.AppendLine("BEGIN:VEVENT");
            if (IsCancelled)
            {
                sb.AppendLine("STATUS:CANCELLED");
                sb.AppendLine("PRIORITY:1");
            }

            sb.AppendLine(string.Format("ATTENDEE;RSVP=TRUE;ROLE=REQ-PARTICIPANT:MAILTO:{0}", SendTo));
            sb.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", SendFrom));
            sb.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", StartTime.ToUniversalTime()));
            sb.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", EndTime.ToUniversalTime()));
            sb.AppendLine(string.Format("LOCATION:{0}", Location));
            sb.AppendLine("TRANSP:OPAQUE");
            //You need to increment the sequence anytime you update the meeting request. 
            sb.AppendLine(string.Format("SEQUENCE:{0}", Sequence));
            //This needs to be a unique ID. A GUID is created when the appointment entity is inserted
            sb.AppendLine(string.Format("UID:{0}", MsgID));
            sb.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            sb.AppendLine(string.Format("DESCRIPTION:{0}", Body));
            sb.AppendLine(string.Format("SUMMARY:{0}", Subject));
            sb.AppendLine("CLASS:PUBLIC");
            //Create a 15min reminder
            sb.AppendLine("BEGIN:VALARM");
            sb.AppendLine("TRIGGER:-PT15M");
            sb.AppendLine("ACTION:DISPLAY");
            sb.AppendLine("DESCRIPTION:Reminder");
            sb.AppendLine("END:VALARM");

            sb.AppendLine("END:VEVENT");
            sb.AppendLine("END:VCALENDAR");

            return sb.ToString();
        }
    }
}