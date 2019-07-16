// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;

namespace Microsoft.Dynamics365.UIAutomation.Api
{
    /// <summary>
    /// Xrm Guided Help Page
    /// </summary>
    public class GuidedHelp
        : XrmPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GuidedHelp"/> class.
        /// </summary>
        /// <param name="browser">The browser.</param>
        public GuidedHelp(InteractiveBrowser browser)
            : base(browser)
        {
            SwitchToDefault();
        }

        public bool IsEnabled
        {
            get
            {
                bool isGuidedHelpEnabled = false;
                try
                {
                    bool.TryParse(
                        this.Browser.Driver.ExecuteScript("return Xrm.Internal.isFeatureEnabled('FCB.GuidedHelp') && Xrm.Internal.isGuidedHelpEnabledForUser();").ToString(),
                        out isGuidedHelpEnabled);
                }
                catch (System.NullReferenceException)
                {
                    // do nothing
                }


                return isGuidedHelpEnabled;
            }
        }

        /// <summary>
        /// Closes the Guided Help
        /// </summary>
        /// <example>xrmBrowser.GuidedHelp.CloseGuidedHelp();</example>
        public BrowserCommandResult<bool> CloseGuidedHelp(int thinkTime = Constants.DefaultThinkTime)
        {
            this.Browser.ThinkTime(thinkTime);
            return this.Execute(GetOptions("Close Guided Help"), driver =>
            {
                bool returnValue = false;

                if (IsEnabled)
                {
                    driver.WaitUntilVisible(By.XPath(Elements.Xpath[Reference.GuidedHelp.MarsOverlay]), new TimeSpan(0, 0, 15), d =>
                    {
                        var allMarsElements = driver
                            .FindElement(By.XPath(Elements.Xpath[Reference.GuidedHelp.MarsOverlay]))
                            .FindElements(By.XPath(".//*"));

                        foreach (var element in allMarsElements)
                        {
                            var buttonId = driver.ExecuteScript("return arguments[0].id;", element).ToString();

                            if (buttonId.Equals(Elements.ElementId[Reference.GuidedHelp.Close], StringComparison.InvariantCultureIgnoreCase))
                            {
                                driver.WaitUntilClickable(By.Id(buttonId), new TimeSpan(0, 0, 5));

                                element.Click();
                            }
                        }

                        returnValue = true;
                    });
                }
                else
                {
                    IWebElement iframe = driver.FindElement(By.Id("InlineDialog_Iframe"));
                    driver.SwitchTo().Frame(iframe);
                    IReadOnlyCollection<IWebElement> el = driver.FindElements(By.CssSelector("*"));
                    foreach (IWebElement e in el)
                    {
                        if(e.TagName == "img")
                        {
                            e.Click();
                        }
                    }
                    driver.SwitchTo().DefaultContent();
                }

                return returnValue;
            });
        }
    }
}