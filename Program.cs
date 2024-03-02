using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;

namespace SQL_Connection
{
    internal class Program
    {
        public const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string Toad_App = "C:\\Program Files (x86)\\Quest Software\\Toad for Oracle 16.1\\Toad.exe"; //"Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";
        protected static WindowsDriver<WindowsElement> session;
        //  static bool  isRunning = Process.GetProcessesByName("sqldeveloper").Length > 0; //return true if application running in background
        Process[] processes = Process.GetProcessesByName("Toad.exe"); //return true if application running in background

        [DllImport("user32.dll")]
        static extern bool SetForgroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        static void Main(string[] args)
        {

            Program checker = new Program(); //Creating 'checker' object of Program class to 'calling non static method in static method i.e focus on 'static method'.
                                             //IntPtr hWnd = FindWindow(null, "sqldeveloper.exe"); //return true if application running in background
                                             //  if (hWnd!=IntPtr.Zero) //return true if application running in background
                                             //  {
                                             //      SetForgroundWindow(hWnd);//shifting control on existing application 
                                             //  }
                                             //  else
                                             //  {
                                             //      Debug.WriteLine("no app running");
                                             //      Thread.Sleep(10000);
                                             //  }
            Setup();
            TearDown();
            // checker.check(); //calling non static method in static method i.e focus on static method only.
        }
        public void check() //non static method
        {
            // ok();
        }
        void ok()//non static method
        {
            if (processes.Length == 0)
            {
                Process.Start("Toad.exe");
            }
            else
            {
                Debug.WriteLine(" Toad_app is  running already");
                Thread.Sleep(10000);
            }
        }
        public static void Setup()
        {
            if (session == null)// Launch Calculator application if it is not yet launched
            {
                try
                {
                    //Appium.WebDriver OR WinappDriver Code Start here
                    Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");//Starting WinAppDriver for windows, for android,ios,macos no need to use this line i.e you can comment this line. 
                    AppiumOptions op = new AppiumOptions();//creating object/instance of appium note:this is not actual appium this is c# appium which is different than real one.                               
                    op.AddAdditionalCapability("app", Toad_App); //defining which application we want to automate
                    op.AddAdditionalCapability("deviceName", "WindowsPC");//Adding platform like windows,macos,android,ios etc
                    op.AddAdditionalCapability("appArguments", "-internal -NoSplash");//skiping flash screen
                    session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), op); // Accessing WinAppDriver control into variable: session
                    session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5); //1.5 sec //don't use Thread.Sleep(2000);//2 sec as it increase RAM load
                    session.SwitchTo().Window(session.WindowHandles[0]);
                    session.Manage().Window.Maximize();
                    Connect_To_DB();//call automation method here
                    //Appium.WebDriver OR WinappDriver Code End here
                }
                catch
                {
                    AppiumOptions op = new AppiumOptions();//creating object/instance of appium note:this is not actual appium this is c# appium which is different than real one.                               
                    op.AddAdditionalCapability("app", "Root"); //Root mean we are pointing to application which is opened in try block. i.e avoiding duplication of application
                    op.AddAdditionalCapability("deviceName", "WindowsPC");//Adding platform like windows,macos,android,ios etc
                    op.AddAdditionalCapability("appArguments", "-internal -NoSplash");//skiping flash screen
                    session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), op); // Accessing WinAppDriver control into variable: session
                    Connect_To_DB();//call automation method here
                }
            }
        }

        public static void Connect_To_DB()
        {

            string json_data = File.ReadAllText(@"C:\Tools\EnvManager\Env_Details.json");
            JObject p = JObject.Parse(json_data);
            string db_name = p["DB_Name"].ToString();
            string vapp_name = p["Vapp_Name"].ToString();
            string username = p["Username"].ToString();
            string password = p["Password"].ToString();
            string dbhost = p["DBHost"].ToString();
            string instance = p["Instance"].ToString();

            Thread.Sleep(15000);
            //Windows control or focus get transfer automatically by windows kernal itself so no need to use 'session.SwitchTo()'; to shift from parent to child or vice versa
            Actions actions = new Actions(session); //Use this 'Actions' methods when FindElementByName etc not working correctly 


            /*   actions.KeyDown(Keys.Alt).SendKeys(Keys.F10).KeyUp(Keys.Alt).Build().Perform();
               //actions.KeyDown(Keys.Control).SendKeys("n").KeyUp(Keys.Control).Build().Perform();
               // actions.SendKeys(Keys.Tab).Perform();
               object value = session.Keyboard.SendKeys(Keys.Enter);

              // session.FindElementByClassName("Clear").Click();
               actions.SendKeys(vapp_name).Perform(); //Entering vapp_name 
               Thread.Sleep(5000);
                int i1 = 1;
               while (i1 <= 5)
               {
                   session.Keyboard.SendKeys(Keys.Tab);
                  // Thread.Sleep(100);
                   i1++;
               }
               actions.SendKeys(Keys.Space + username).Perform(); //Entering username 
               int i2 = 1;
               while (i2 <= 3)
               {
                    session.Keyboard.SendKeys(Keys.Tab);
                   i2++;
               }
                actions.SendKeys(Keys.Space + password).Perform();//Entering password 
                session.Keyboard.SendKeys(Keys.Tab);
                session.Keyboard.SendKeys(Keys.Space);
               int i3 = 1;
               while (i3 <= 2)
               {
                   session.Keyboard.SendKeys(Keys.Tab);
                   i3++;
               }
               actions.SendKeys(Keys.Space + dbhost).Perform();//Entering dbhost


               // actions.SendKeys("okk");
               actions.SendKeys(Keys.Enter).Perform();
               actions.SendKeys("$vapp_name").Perform();
               session.Keyboard.SendKeys(Keys.Tab); 


           }*/

            //please refer this below  FindElementByAccessibilityId,FindElementByXPath etc of calculator app becuase it will give u correct direction to built correct xpath,FindElementByName etc






            public static void TearDown()
            {
                if (session != null)
                {
                    session.Quit();// Close the application and delete the session
                    session = null;
                    Array.ForEach(Process.GetProcessesByName("WinAppDriver"), x => x.Kill());//close the winappdriver window 
                }
            }


        }
    }

