// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Collections.Generic;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Web
{
    [TestClass]
    public class CreateLead
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void WEBTestCreateNewLead()
        {
            using (var xrmBrowser = new Api.Browser(TestSettings.Options))
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                //xrmBrowser.Navigation.SwitchToDialogFrame()

                xrmBrowser.ThinkTime(2000);

                xrmBrowser.Navigation.OpenSubArea("comercial", "Leads");
                
                xrmBrowser.Grid.SwitchView("All Leads");
                
                xrmBrowser.CommandBar.ClickCommand("New");

                List<Field> fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = "Test"},
                    new Field() {Id = "lastname", Value = "Lead"}
                };

                xrmBrowser.ThinkTime(2000);
                
                xrmBrowser.Entity.SetValue("firstname", string.Format("Martin{0}", TestSettings.GetRandomString(1, 5)));
                xrmBrowser.Entity.SetValue("lastname", string.Format("Cerdeira{0}", TestSettings.GetRandomString(1, 5)));

                string dni = string.Format("29{0}", TestSettings.GetRandomNumber(6, 6));

                xrmBrowser.Entity.SetValue(new OptionSet() { Name= "pnet_documenttype", Value = "102610000"});
                xrmBrowser.Entity.SetValue("pnet_documentnumber", dni);

                Random random = new Random();
                string gender = random.Next(1, 3).ToString();

                xrmBrowser.Entity.SetValue(new OptionSet() { Name = "pnet_gender", Value = gender });
                xrmBrowser.Entity.SetValue(new OptionSet() { Name = "leadsourcecode", Value = "102610015" });
                xrmBrowser.Entity.SetValue(new OptionSet() { Name = "pnet_phonetype1", Value = "102610000" });

                xrmBrowser.Entity.SetValue("pnet_areacode1", "1111");
                xrmBrowser.Entity.SetValue("pnet_phone1", "111111");

                xrmBrowser.Entity.SetValue("emailaddress1", "martin@gmail.com");

                xrmBrowser.Entity.SetValue("pnet_cnaeid","Acopio de algodón" );

                xrmBrowser.Entity.SetHeaderValue("pnet_subsidiaryid", "9 de Julio");

                xrmBrowser.CommandBar.ClickCommand("Save");
                xrmBrowser.ThinkTime(3000);

                Guid ID = xrmBrowser.Entity.GetRecordGuid(100);

                // Go to Commercial Background
                string url = String.Format(_xrmUri + "main.aspx?etc=10032&extraqs=%3f_CreateFromId%3d%257b{0}%257d%26_CreateFromType%3d4%26_searchText%3d%26etc%3d10032%26parentLookupControlId%3dlookup_CommercialBackgrounds_i&newWindow=true&pagetype=entityrecord", ID);
                
                xrmBrowser.Navigate(url);
                xrmBrowser.ThinkTime(2000);

                xrmBrowser.Entity.SetValue(new TwoOption() { Name = "pnet_classification", Value = "true" });

                xrmBrowser.Entity.SetValueJS("pnet_todate", "new Date()", 100);

                xrmBrowser.ThinkTime(1000);

                xrmBrowser.CommandBar.ClickCommand("Save");
                xrmBrowser.ThinkTime(3000);

                xrmBrowser.Entity.OpenEntity("lead", ID);

                xrmBrowser.ThinkTime(5000);

                xrmBrowser.CommandBar.ClickCommand("Qualify");
            }
        }
    }
}