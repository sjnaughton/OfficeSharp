using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace OfficeSharp.SMTP
{
    public class SmtpServer
    {
        private string sSmtpServerName;
        private string sSmtpUserId;
        private string sSmtpPassword;
        private int iSmtpPort;

        public SmtpServer()
        {
        }

        public SmtpServer(string sSmtpUserId, string sSmtpServer, string sSmtpPassword, int iSmtpPort)
        {
            this.sSmtpUserId = sSmtpUserId;
            this.sSmtpServerName = sSmtpServer;
            this.sSmtpPassword = sSmtpPassword;
            this.iSmtpPort = iSmtpPort;
        }

        public string SmtpServerName
        {
            get { return sSmtpServerName; }
            set { sSmtpServerName = value; }
        }

        public string SmtpUserId
        {
            get { return sSmtpUserId; }
            set { sSmtpUserId = value; }
        }

        public string SmtpPassword
        {
            get { return sSmtpPassword; }
            set { sSmtpPassword = value; }
        }

        public int SmtpPort
        {
            get { return iSmtpPort; }
            set { iSmtpPort = value; }
        }

    }
}
