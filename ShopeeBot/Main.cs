using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ComponentFactory.Krypton.Toolkit;
using System.Configuration;
using System.Net.Http;
using System.Net;
using System.Globalization;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace ShopeeBot
{
    public partial class Main : KryptonForm
    {
        public Main()
        {
            InitializeComponent();
            this.Palette = MyPalette;
            InitializationSettingsChrome();
            Tb_LinkProduct.Text = ConfigurationManager.AppSettings["default_link_produk"];
            Tb_Cookies.Text = ConfigurationManager.AppSettings["default_cookiee"];
            Tb_Varian.Text = ConfigurationManager.AppSettings["default_varian"];
        }
        private ChromeOptions options = null;
        private IWebDriver driver = null;
        private bool bActive = false;
        private string sSetPointTime = "";
        private void InitializationSettingsChrome()
        {
            this.options = new ChromeOptions();
            if (ConfigurationManager.AppSettings["default_headless"] == "true") this.options.AddArgument("--headless");
            this.options.AddArgument("--disable-extensions");
            this.options.AddArgument("start-maximized");
            this.options.AddArgument("disable-infobars");
            this.options.AddArgument("--disable-extensions");
            this.options.AddUserProfilePreference("profile.default_content_setting_values.images", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.plugins", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.popups", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.auto_select_certificate", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.fullscreen", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.mouselock", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.mixed_script", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.media_stream_mic", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.media_stream_camera", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.protocol_handlers", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.ppapi_broker", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.midi_sysex", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.push_messaging", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.ssl_cert_decisions", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.metro_switch_to_desktop", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.protected_media_identifier", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.app_banner", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.site_engagement", 2);
            this.options.AddUserProfilePreference("profile.default_content_setting_values.durable_storage", 2);
        }
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (driver != null)
            {
                driver.Close();
                driver.Quit();
                driver = null;
            }
        }
        private string LogGenerator(string sMessage)
        {
            return $"[{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff", CultureInfo.InvariantCulture)}] {sMessage}";
        }
        private void AddLogs(string sMessage)
        {
            Lb_Logs.Items.Add(sMessage);
            Lb_Logs.TopIndex = Lb_Logs.Items.Count - 1;
        }
        private void Bt_Start_Click(object sender, EventArgs e)
        {
            LoadCookies();
            this.sSetPointTime = Dt_Picker.Value.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
            AddLogs(LogGenerator($"Time Config: {sSetPointTime}"));
            TimerRealtime.Start();
            TimerRefresh.Start();
            this.bActive = true;
        }
        private void LoadCookies()
        {
            try
            {
                driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\webdriver", options);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                driver.Navigate().GoToUrl("https://shopee.co.id");
                AddLogs(LogGenerator("Visting https://shopee.co.id"));
                driver.Manage().Cookies.AddCookie(new OpenQA.Selenium.Cookie("SPC_EC", Tb_Cookies.Text));
                AddLogs(LogGenerator("Adding Cookies"));
                driver.Navigate().GoToUrl(Tb_LinkProduct.Text);
                AddLogs(LogGenerator("Visting Link Product"));
            }
            catch (Exception ex)
            {
                AddLogs(LogGenerator($"Error on {ex.Message}"));
            }
        }
        private void Bt_Test_Click(object sender, EventArgs e)
        {
            try
            {
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(Tb_LinkProduct.Text);
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                if (myHttpWebResponse.StatusCode == HttpStatusCode.OK)
                {
                    AddLogs(LogGenerator("Link Product is OK"));
                }
                myHttpWebResponse.Close();
            }
            catch (WebException ex)
            {
                AddLogs(LogGenerator($"Link Product Raised. The following error occurred : {0} {ex.Status.ToString()}"));
            }
            catch (Exception ex)
            {
                AddLogs(LogGenerator($"The following Exception was raised from Link Product : {0} {ex.Message}"));
            }
        }
        private void PurchaseAProduct()
        {
            try
            {
                AddLogs(LogGenerator("Gow!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"));
                //Klik Tombol Varian
                if (Tb_Varian.Text != "")
                {
                    var varian = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                    varian.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath($"//button[contains(text(),'{Tb_Varian.Text}')]"))).Click();
                    AddLogs(LogGenerator("Berhasil memilih Varian!"));
                }

                //Klik Tombol Beli
                var beli = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                beli.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='main']/div/div[2]/div[1]/div/div[2]/div/div[1]/div[3]/div/div[5]/div/div/button[2]"))).Click();
                AddLogs(LogGenerator("Proses pembelian Barang!"));

                //Klik Tombol Checkout
                var checkout = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                checkout.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='main']/div/div[2]/div[2]/div/div[3]/div[2]/div[7]/button[4]/span"))).Click();
                AddLogs(LogGenerator("Proses berhasil di checkout!"));

                //Klik Tombol Buat Pesanan
                var pesanan = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                pesanan.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='main']/div/div[3]/div[2]/div[4]/div[2]/div[7]/button"))).Click();
                AddLogs(LogGenerator("Pesanan Dibuat!"));

                //Klik Tombol Bayar
                var bayar = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                bayar.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("pay-button"))).Click();
                AddLogs(LogGenerator("Barang sedang di Bayar!"));

                //Masukan Shopee Pay
                var shopeePay = new WebDriverWait(driver, TimeSpan.FromSeconds(100));
                shopeePay.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[9]/div[1]/div[3]/div[1]"))).Click();
                AddLogs(LogGenerator("Barang berhasil di Bayar!"));
            }
            catch (Exception ex)
            {
                AddLogs(LogGenerator($"Error on {ex.Message}"));
            }
        }
        private void TimerRealtime_Tick(object sender, EventArgs e)
        {
            if (bActive)
            {
                timer1.Stop();
                AddLogs($"Waktu yang di setting [{sSetPointTime}] dan Waktu Sekarang [{DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)}]");
                Lb_Time.Text = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
                if (DateTime.Compare(Dt_Picker.Value, DateTime.Now) < 0)
                {
                    TimerRefresh.Stop();
                    TimerRealtime.Stop();
                    // Here our Script
                    PurchaseAProduct();
                    timer1.Start();
                    bActive = false;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Lb_Time.Text = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
        }

        private void TimerRefresh_Tick(object sender, EventArgs e)
        {
            driver.Navigate().Refresh();
        }

        private void Bt_Batal_Click(object sender, EventArgs e)
        {
            if (driver != null)
            {
                driver.Close();
                driver.Quit();
                driver = null;
            }
            if (bActive)
            {
                AddLogs($"[{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff", CultureInfo.InvariantCulture)}] Proses di batalkan!");
                TimerRefresh.Stop();
                TimerRealtime.Stop();
                timer1.Start();
                bActive = false;
            }
        }
    }
}
