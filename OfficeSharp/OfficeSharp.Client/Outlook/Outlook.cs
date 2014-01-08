using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices.Automation;
using Microsoft.LightSwitch;
using Microsoft.LightSwitch.Client;
using Microsoft.LightSwitch.Details;
using Microsoft.LightSwitch.Framework;
using Microsoft.LightSwitch.Model;
using Microsoft.LightSwitch.Presentation;
using Microsoft.LightSwitch.Presentation.Extensions;
using Microsoft.LightSwitch.Runtime.Shell.Framework;
using Microsoft.LightSwitch.Sdk.Proxy;
using Microsoft.LightSwitch.Threading;
using Microsoft.VisualStudio.ExtensibilityHosting;

namespace OfficeSharp.Outlook
{
    public static class Outlook
    {
        // Outlook Appointment Constants
        const int olAppointmentItem = 1;
        const int olMeeting = 1;
        // Outlook Mail Item Constants
        const int olMailItem = 0;
        const int olFormatPlain = 1;

        const int olFormatHTML = 2;
        static dynamic outlook = null;

        static dynamic oNS = null;
        //Public Sub CreateAppointment(ByVal Address As String,
        public static dynamic CreateAppointment(string Address, string Subject, string Body, string Location, System.DateTime StartDateTime, System.DateTime EndDateTime)
        {
            dynamic result = false;

            try
            {
                if (GetOutlook())
                {
                    //Create the Appointment
                    dynamic appt = outlook.CreateItem(olAppointmentItem);
                    var _with1 = appt;
                    _with1.Body = Body;
                    _with1.Subject = Subject;
                    _with1.Start = StartDateTime;
                    _with1.End = EndDateTime;
                    _with1.Location = Location;
                    _with1.MeetingStatus = olMeeting;
                    _with1.Recipients.Add(Address);
                    _with1.Display();
                    result = true;

                    // Returning the dynamic 
                    return appt;
                }

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create Appointment.", ex);
            }
            return null;
        }

        public static dynamic CreateEmail(string Address, string Subject, string Body)
        {
            try
            {
                if (GetOutlook())
                {
                    dynamic mail = outlook.CreateItem(olMailItem);
                    var _with2 = mail;
                    if (Body.ToLower().Contains("<html>"))
                    {
                        _with2.BodyFormat = olFormatHTML;
                        _with2.HTMLBody = Body;
                    }
                    else
                    {
                        _with2.BodyFormat = olFormatPlain;
                        _with2.Body = Body;
                    }

                    _with2.Recipients.Add(Address);
                    _with2.Subject = Subject;
                    _with2.Display();
                    // Returning the dynamic 
                    return mail;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create email.", ex);
            }
            return null;
        }

        public static dynamic CreateEmail(string Address, string Subject, IVisualCollection Items)
        {
            try
            {
                string sBody = null;
                sBody = HtmlExport(Items);


                if (GetOutlook())
                {
                    dynamic mail = outlook.CreateItem(olMailItem);
                    var _with3 = mail;
                    // checking if it contains an html tags
                    if (sBody.ToLower().Contains("<html>"))
                    {
                        _with3.BodyFormat = olFormatHTML;
                        _with3.HTMLBody = sBody;
                    }
                    else
                    {
                        _with3.BodyFormat = olFormatPlain;
                        _with3.Body = sBody;
                    }

                    _with3.Recipients.Add(Address);
                    _with3.Subject = Subject;
                    _with3.Display();

                    // Returning the dynamic 
                    return mail;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create email.", ex);
            }
            return null;
        }

        // Create HTML table based on IVisualCollection
        public static string HtmlExport(IVisualCollection Collection, List<string> ColumnNames)
        {
            string sBody = "";

            if (Collection.Count > 0)
            {
                // opening an html tag and creating a table 
                //sBody = "<html>"
                //sBody = sBody + "</br></br>"
                //sBody = sBody + "<body style=""font-family: Arial, Helvetica, sans-serif;"" >"
                sBody = "<table border=\"1\">";

                string sColumnName = null;

                // add columns names to the list
                // row begins
                sBody = sBody + "<tr>";
                foreach (string sColumnName_loopVariable in ColumnNames)
                {
                    sColumnName = sColumnName_loopVariable;
                    sBody = sBody + "<td>";
                    sBody = sBody + " " + sColumnName;
                    sBody = sBody + "</td>";
                }
                // row ends
                sBody = sBody + "</tr>";

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in Collection)
                {
                    sBody = sBody + "<tr>";
                    for (int i = 0; i <= ColumnNames.Count - 1; i++)
                    {
                        sBody = sBody + "<td>";
                        sBody = sBody + LightSwitchHelper.GetValue(entityObj, ColumnNames[i]);
                        sBody = sBody + "</td>";
                    }
                    sBody = sBody + "</tr>";
                }

                // closing the tags
                sBody = sBody + "</table>";
                //sBody = sBody + "</body>"
                //sBody = sBody + "</html>"

            }
            return sBody;
        }

        // Create HTML table based on IVisualCollection
        public static string HtmlExport(IVisualCollection collection)
        {
            List<string> columnNames = new List<string>();

            string sBody = "";

            if (collection.Count > 0)
            {
                // opening an html tag and creating a table 
                //Body = "<html>"
                //sBody = sBody + "</br></br>"
                //sBody = sBody + "<body style=""font-family: Arial, Helvetica, sans-serif;"" >"
                sBody = "<table border=\"1\">";

                // get column properties
                IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

                // add columns names to the list
                // row begins
                sBody = sBody + "<tr>";
                foreach (IEntityProperty entityProperty in columnProperties)
                {
                    columnNames.Add(entityProperty.Name);
                    sBody = sBody + "<td>";
                    sBody = sBody + " " + entityProperty.DisplayName;
                    sBody = sBody + "</td>";
                }
                // row ends
                sBody = sBody + "</tr>";

                // iterate the collection and extract values by column name
                foreach (IEntityObject entityObj in collection)
                {
                    sBody = sBody + "<tr>";
                    for (int i = 0; i <= columnNames.Count - 1; i++)
                    {
                        sBody = sBody + "<td>";
                        sBody = sBody + LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                        sBody = sBody + "</td>";
                    }
                    sBody = sBody + "</tr>";
                }

                // closing the tags
                sBody = sBody + "</table>";
                //sBody = sBody + "</body>"
                //sBody = sBody + "</html>"
            }
            return sBody;
        }

        // Create HTML table based on IEntityCollection
        public static string HtmlExportEntityCollection(IEntityCollection Collection, List<string> ColumnNames)
        {
            bool bFirstItem = true;
            string sBody = "";

            // iterate the collection and extract values by column name
            foreach (IEntityObject entityObj in Collection)
            {
                if (bFirstItem)
                {
                    // opening an html tag and creating a table 
                    //sBody = "<html>"
                    //sBody = sBody + "</br></br>"
                    //sBody = sBody + "<body style=""font-family: Arial, Helvetica, sans-serif;"" >"
                    sBody = "<table border=\"1\">";

                    // add columns names to the list
                    // row begins
                    string sColumnName = null;
                    sBody = sBody + "<tr>";
                    foreach (string sColumnName_loopVariable in ColumnNames)
                    {
                        sColumnName = sColumnName_loopVariable;
                        sBody = sBody + "<td>";
                        sBody = sBody + " " + sColumnName;
                        sBody = sBody + "</td>";
                    }
                    // row ends
                    sBody = sBody + "</tr>";
                    bFirstItem = false;
                }

                sBody = sBody + "<tr>";
                for (int i = 0; i <= ColumnNames.Count - 1; i++)
                {
                    sBody = sBody + "<td>";
                    sBody = sBody + LightSwitchHelper.GetValue(entityObj, ColumnNames[i]);
                    sBody = sBody + "</td>";
                }
                sBody = sBody + "</tr>";
            }

            // Add closing tags if there was at least one item
            // bFirstItem = True by default
            // It is set to False when the first item is encountered
            if (!bFirstItem)
            {
                // closing the tags
                sBody = sBody + "</table>";
                //sBody = sBody + "</body>"
                //sBody = sBody + "</html>"
            }
            return sBody;
        }

        // Create HTML table based on IEntityCollection
        public static string HtmlExportEntityCollection(IEntityCollection collection)
        {
            List<string> columnNames = new List<string>();
            bool bFirstItem = true;
            string sBody = "";

            // get column properties
            IEnumerable<IEntityProperty> columnProperties = collection.OfType<IEntityObject>().First().Details.Properties.All();

            // iterate the collection and extract values by column name
            foreach (IEntityObject entityObj in collection)
            {
                if (bFirstItem)
                {
                    // opening an html tag and creating a table 
                    //sBody = "<html>"
                    //sBody = sBody + "</br></br>"
                    //sBody = sBody + "<body style=""font-family: Arial, Helvetica, sans-serif;"" >"
                    sBody = "<table border=\"1\">";

                    // add columns names to the list
                    // row begins
                    sBody = sBody + "<tr>";
                    foreach (IEntityProperty entityProperty in columnProperties)
                    {
                        columnNames.Add(entityProperty.Name);
                        sBody = sBody + "<td>";
                        sBody = sBody + " " + entityProperty.DisplayName;
                        sBody = sBody + "</td>";
                    }
                    // row ends
                    sBody = sBody + "</tr>";
                    bFirstItem = false;
                }

                sBody = sBody + "<tr>";
                for (int i = 0; i <= columnNames.Count - 1; i++)
                {
                    sBody = sBody + "<td>";
                    sBody = sBody + LightSwitchHelper.GetValue(entityObj, columnNames[i]);
                    sBody = sBody + "</td>";
                }
                sBody = sBody + "</tr>";
            }

            // Add closing tags if there was at least one item
            // bFirstItem = True by default
            // It is set to False when the first item is encountered
            if (!bFirstItem)
            {
                // closing the tags
                sBody = sBody + "</table>";
                //sBody = sBody + "</body>"
                //sBody = sBody + "</html>"
            }
            return sBody;
        }

        private static bool GetOutlook()
        {
            try
            {
                // If GetObject throws an exception, then Outlook is 
                // either not running or is not available.
                outlook = AutomationFactory.GetObject("Outlook.Application");
                return true;
            }
            catch
            {
                try
                {
                    // Start Outlook and display the Inbox, but minimize 
                    // it to avoid hiding the running application.

                    outlook = AutomationFactory.CreateObject("Outlook.Application");
                    outlook.Session.GetDefaultFolder(6).Display();
                    // 6 = Inbox
                    outlook.ActiveWindow.WindowState = 1;
                    // minimized
                    return true;
                }
                catch
                {
                    // Outlook is unavailable.
                    return false;
                }
            }
        }
    }
}
