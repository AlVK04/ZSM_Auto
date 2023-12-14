using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Aspose.Cells;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Threading;
using System.Reflection;

namespace Automatization
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IWebDriver web;

        public string FilePath;

        Workbook wb;
        WorksheetCollection collection;

        List<string> worksheetNames = new List<string>();

        List<Products> productsList = new List<Products>();

        List<string> shopNames = new List<string>();

        private LogWindow logWindow;
        private Thread newWindowThread;


        public MainWindow()
        {
            InitializeComponent();
        }

        // Кнопка Найти магазин и заполнить
        private void ButtonSearchAndFill_Click(object sender, RoutedEventArgs e)
        {
            //Создание нового потока с запуском окна
            NewThreadForWindowLog();
            // Создаем новый экземпляр сервиса Chrome
            var service = ChromeDriverService.CreateDefaultService(@"chromedriver.exe");
            service.HideCommandPromptWindow = true; // Скрываем консоль ChromeDriver
            web = new ChromeDriver(service);

            web.Navigate().GoToUrl("https://monitoring.rk.gov.ru/admin/structure/shop");  //Переходим по URL

            web.FindElement(By.XPath("//input[@id='inputUsername']")).SendKeys(Login.Text); // Вводим логин
            web.FindElement(By.XPath("//input[@id='inputPassword']")).SendKeys(Password.Text); // Вводим пароль

            web.FindElement(By.XPath("//button[@class='button']")).Click(); // Кликаем войти

            if (!IsChainOfStores.IsChecked.Value)
            {
                ShopSearch();
            }

            OpenShopEdit(false);
            web.Quit();
            CloseLogWindow();
        }


        //метод, создающий поток для второго окна

        private void NewThreadForWindowLog()
        {
            try {
                newWindowThread = new Thread(new ThreadStart(() =>
                {
                    // Создание нового окна и запуск цикла обработки сообщений
                    logWindow = new LogWindow();
                    logWindow.Show();
                    System.Windows.Threading.Dispatcher.Run();



                }));

                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.Start();
            }
            catch (ThreadInterruptedException e)
            {
                
            }

        }

        private void TransmitText(string text, int index) 
       //0 - обновление магазина, 1 - заполнение магазина 2 - Конец заполнения сети 
        {
            switch (index)
            {
                case 0:
                    text += " обновлён";
                    logWindow.Dispatcher.Invoke(() =>
                    {
                        logWindow.ReceiveData(text, index);
                    });
                    break;
                case 1:
                    text += " заполнен";
                    logWindow.Dispatcher.Invoke(() =>
                    {
                        logWindow.ReceiveData(text, index);
                    });
                    break;
                case 2:
                    text = "Сеть заполненна";
                    logWindow.Dispatcher.Invoke(() =>
                    {
                        logWindow.ReceiveData(text, index);
                    });
                    break;
               
                default:
                    text = "Неизвестная ошибка!";
                    logWindow.Dispatcher.Invoke(() => { logWindow.ReceiveData(text, -1); });
                    break;
            }
        }

        // Закрыть браузер
        private void ButtonCloseShop(object sender, RoutedEventArgs e)
        {
            try
            {
                web.Quit();
                CloseLogWindow();
                Log.Text += "Браузер закрыт\n";
            }
            catch
            {
                Log.Text += "Не удалось закрыть браузер, возможно он уже закрыт!\n";
            }

        }

        void ShopSearch()
        {
            web.FindElement(By.XPath("//form[@class='search']/input")).Clear();
            web.FindElement(By.XPath("//form[@class='search']/input")).SendKeys(Name.Text);
            web.FindElement(By.XPath("//form[@class='search']/button[@class='btn']")).Click();
        }


        void OpenShopEdit(bool IsOnlyUpdate)
        {
            int filledShops = 0;
            List<string> FillShopNames = new List<string>();

            string LastUrl = web.Url;

            try
            {
                if (IsOnlyUpdate)
                {
                    for (int i = 2; ; i++)
                    {
                        web.Navigate().GoToUrl(LastUrl);
                        if (i <= 21)
                        {                            
                            if (!FillShopNames.Any(x => x.Contains(web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim().ToLower())))
                            {
                                FillShopNames.Add(web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim().ToLower());

                                LastUrl = web.Url;
                                web.FindElement(By.XPath($"//tr[{i}]/td[@class='row-btn']/a[@class='btn']")).Click(); // Открываем редактирование магазина

                                ShopEdit(true);
                                Log.Text += FillShopNames.Last() + " Обновлен\n";
                                TransmitText(FillShopNames.Last(), 0);

                                filledShops++;
                            }
                        }
                        else
                        {
                            i = 1;
                            web.FindElement(By.XPath("//a[@class='pagination__link pagination__link_arrow_right']")).Click(); // Переход на следующую страницу
                            LastUrl= web.Url;
                        }
                    }
                }
                else
                {
                    for (int i = 2; ; i++)
                    {
                        if (i <= 21)
                        {
                            if (IsChainOfStores.IsChecked.Value)
                            {
                                if (web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim().Split('(')[0].Trim().ToLower() == Name.Text.ToString().Trim().ToLower() &&
                                    !FillShopNames.Any(x => x.Contains(web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim().ToLower())))
                                {
                                    FillShopNames.Add(web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim().ToLower());

                                    web.FindElement(By.XPath($"//tr[{i}]/td[@class='row-btn']/a[@class='btn']")).Click(); // Открываем редактирование магазина

                                    ShopEdit(false);
                                    Log.Text += FillShopNames.Last() + " ЗАПОЛНЕН\n";
                                    TransmitText(FillShopNames.Last(), 1);
                                    filledShops++;
                                }
                            }
                            else
                            {
                                if (web.FindElement(By.XPath($"//tbody/tr[{i}]/td[2]")).Text.ToString().Trim() == Name.Text.ToString().Trim())
                                {
                                    web.FindElement(By.XPath($"//tr[{i}]/td[@class='row-btn']/a[@class='btn']")).Click(); // Открываем редактирование магазина

                                    ShopEdit(false);
                                    Log.Text += Name.Text.ToString().Trim() + " ОБНОВЛЕН\n";
                                    TransmitText(Name.Text.ToString().Trim(), 0);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            i = 1;
                            web.FindElement(By.XPath("//a[@class='pagination__link pagination__link_arrow_right']")).Click(); // Переход на следующую страницу
                        }
                    }
                }
            }
            catch
            {
                web.Quit();
                CloseLogWindow();
                if (filledShops > 0)
                {
                    Log.Text += "Сеть заполнена\n";
                    TransmitText("", 2);
                }
                else
                {
                    Log.Text += "Не удалось найти магазин!\n";
                }
            }
        }

        void ShopEdit(bool IsOnlyUpdate)
        {
            try
            {
                if (!IsOnlyUpdate)
                {
                    EnterText("//input[@id='shop_breadCost']", Bread.Text);
                    EnterText("//input[@id='shop_breadCost']", Bread.Text);
                    EnterText("//input[@id='shop_breadCost']", Bread.Text);
                    EnterText("//input[@id='shop_beefOnBoneCost']", BeefMeatOnBone.Text);
                    EnterText("//input[@id='shop_beefCost']", BeefMeat.Text);
                    EnterText("//input[@id='shop_porkShoulderCost']", PorkMeatOnBone.Text);
                    EnterText("//input[@id='shop_porkNeckCost']", PorkMeat.Text);
                    EnterText("//input[@id='shop_fowlCost']", Chickens.Text);
                    EnterText("//input[@id='shop_chickenEggCost']", Egg.Text);
                    EnterText("//input[@id='shop_potatoCost']", Potato.Text);
                    EnterText("//input[@id='shop_onionCost']", Onion.Text);
                    EnterText("//input[@id='shop_carrotsCost']", Carrot.Text);
                    EnterText("//input[@id='shop_beetCost']", Beet.Text);
                    EnterText("//input[@id='shop_cabbageCost']", Cabagge.Text);
                    EnterText("//input[@id='shop_wheatFlourCost']", Flour.Text);
                    EnterText("//input[@id='shop_sugarCost']", Sugar.Text);

                    EnterText("//input[@id='shop_buckwheatCost']", Buckwheat.Text);
                    EnterText("//input[@id='shop_groatsWheatCost']", Wheat.Text);
                    EnterText("//input[@id='shop_riceGroatsCost']", Rice.Text);
                    EnterText("//input[@id='shop_pastaCost']", Pasta.Text);
                    EnterText("//input[@id='shop_sunflowerOilCost']", Oil.Text);
                    EnterText("//input[@id='shop_milkCost']", Milk.Text);
                    EnterText("//input[@id='shop_curdCost']", Curd.Text);
                    EnterText("//input[@id='shop_boiledSausageProductsCost']", Sausage.Text);
                }
                //web.FindElement(By.XPath("//a[@class='btn']")).Click(); // Вернуться к списку
                web.FindElement(By.XPath("//button[@class='btn btn-success']")).Click();

            }
            catch (Exception ex)
            {
                Log.Text += ex.Message + "\n";
            }
        }

        private void EnterText(string ElementXPath ,string text)
        {
            web.FindElement(By.XPath(ElementXPath)).Clear();
            web.FindElement(By.XPath(ElementXPath)).SendKeys(text);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                wb = new Workbook(FilePath);
                collection = wb.Worksheets;
                GetWorksheetsNames();
            }
        }

        private void GetWorksheetsNames()
        {
            worksheetNames = new List<string>();
            foreach (Worksheet worksheet in collection)
            {
                worksheetNames.Add(worksheet.Name);
            }

            Types.ItemsSource = worksheetNames;
            Types.SelectedIndex = 0;
        }

        private void GetShopsAndProducts(string worksheetName)
        {            
            Worksheet worksheet = collection[0];

            foreach(Worksheet ws in collection)
            {
                if(ws.Name == worksheetName)
                {
                    worksheet = ws;
                    break;
                }
            }            

            productsList = new List<Products>();
            shopNames = new List<string>();

            for (int i = 2; ; i++)
            {
                if (worksheet.Cells[5, i].Value != null)
                {
                    Products product = new Products();

                    product.Name = worksheet.Cells[5, i].Value.ToString();

                    shopNames.Add(product.Name);

                    product.Bread = GetCost(worksheet.Cells[7, i].Value);
                    product.BeefMeatOnBone = GetCost(worksheet.Cells[8, i].Value);
                    product.BeefMeat = GetCost(worksheet.Cells[9, i].Value);
                    product.PorkMeatOnBone = GetCost(worksheet.Cells[10, i].Value);
                    product.PorkMeat = GetCost(worksheet.Cells[11, i].Value);
                    product.Chickens = GetCost(worksheet.Cells[12, i].Value);
                    product.Egg = GetCost(worksheet.Cells[13, i].Value);
                    product.Potato = GetCost(worksheet.Cells[14, i].Value);
                    product.Onion = GetCost(worksheet.Cells[15, i].Value);
                    product.Carrot = GetCost(worksheet.Cells[16, i].Value);
                    product.Beet = GetCost(worksheet.Cells[17, i].Value);
                    product.Cabagge = GetCost(worksheet.Cells[18, i].Value);
                    product.Flour = GetCost(worksheet.Cells[19, i].Value);
                    product.Sugar = GetCost(worksheet.Cells[20, i].Value);

                    product.Buckwheat = GetCost(worksheet.Cells[22, i].Value);
                    product.Wheat = GetCost(worksheet.Cells[23, i].Value);
                    product.Rice = GetCost(worksheet.Cells[24, i].Value);
                    product.Pasta = GetCost(worksheet.Cells[25, i].Value);
                    product.Oil = GetCost(worksheet.Cells[26, i].Value);
                    product.Milk = GetCost(worksheet.Cells[27, i].Value);
                    product.Curd = GetCost(worksheet.Cells[28, i].Value);
                    product.Sausage = GetCost(worksheet.Cells[29, i].Value);

                    productsList.Add(product);
                }
                else
                {
                    break;
                }
            }

            Shops.ItemsSource = shopNames;
        }

        private string GetCost(object DefaultCost)
        {
            if (DefaultCost != null)
            {
                if (DefaultCost.ToString() == "нет")
                {
                    return "0";
                }
                else
                {
                    return DefaultCost.ToString();
                }
            }
            else
            {
                return "0";
            }
        }

        private void SetProducts(string ShopName)
        {
            Products product = new Products();
            foreach(Products p in productsList)
            {
                if(p.Name == ShopName)
                {
                    product = p;
                    break;
                }
            }


            if (product.Name.Contains(","))
            {
                Name.Text = product.Name.Split(',')[0];
            }
            else if (product.Name.Contains("("))
            {
                Name.Text = product.Name.Split('(')[0];
            }
            else
            {
                Name.Text = product.Name;
            }
            
            Bread.Text = product.Bread;
            BeefMeatOnBone.Text = product.BeefMeatOnBone;
            BeefMeat.Text = product.BeefMeat;
            PorkMeatOnBone.Text = product.PorkMeatOnBone;
            PorkMeat.Text = product.PorkMeat;
            Chickens.Text = product.Chickens;
            Egg.Text = product.Egg;
            Potato.Text = product.Potato;
            Onion.Text = product.Onion;
            Carrot.Text = product.Carrot;
            Beet.Text = product.Beet;
            Cabagge.Text = product.Cabagge;
            Flour.Text = product.Flour;
            Sugar.Text = product.Sugar;

            Buckwheat.Text = product.Buckwheat;
            Wheat.Text = product.Wheat;
            Rice.Text = product.Rice;
            Pasta.Text = product.Pasta;
            Oil.Text = product.Oil;
            Milk.Text = product.Milk;
            Curd.Text = product.Curd;
            Sausage.Text = product.Sausage;
        }

        // Выбран вид магазина
        private void Types_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (Types.SelectedValue != null)
                {
                    GetShopsAndProducts(Types.SelectedValue.ToString());
                    if (Types.SelectedValue.ToString().ToLower().Contains("сети") || Types.SelectedValue.ToString().ToLower().Contains("сеть"))
                    {
                        IsChainOfStores.IsChecked = true;
                    }
                    else
                    {
                        IsChainOfStores.IsChecked = false;
                    }
                }                
            }
            catch { }
        }

        // Выбран магазин
        private void Names_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (Shops.SelectedValue != null)
                    SetProducts(Shops.SelectedValue.ToString());
            }
            catch{ }
        }

        // Обновить магазины
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Создаем новый экземпляр сервиса Chrome
            var service = ChromeDriverService.CreateDefaultService(@"chromedriver.exe");
            service.HideCommandPromptWindow = true; // Скрываем консоль ChromeDriver
            web = new ChromeDriver(service);

            web.Navigate().GoToUrl("https://monitoring.rk.gov.ru/admin/structure/shop");  //Переходим по URL

            web.FindElement(By.XPath("//input[@id='inputUsername']")).SendKeys(Login.Text); // Вводим логин
            web.FindElement(By.XPath("//input[@id='inputPassword']")).SendKeys(Password.Text); // Вводим пароль

            web.FindElement(By.XPath("//button[@class='button']")).Click(); // Кликаем войти

            if (!IsChainOfStores.IsChecked.Value)
            {
                ShopSearch();
            }

            OpenShopEdit(true);
            web.Quit();
            CloseLogWindow();
        }

        private void CloseLogWindow()
        {
            if (logWindow != null)
            {
            logWindow.Dispatcher.Invoke(() =>
            {
                logWindow.ProcessLog.Clear();
                logWindow.Close();

            });
            }
            newWindowThread.Interrupt();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            CloseLogWindow();
            Application.Current.Shutdown();
        }
    }
}

// Кто прочитал тот лох
