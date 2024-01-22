using System.Linq;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Data;
using System.Globalization;

namespace TestTask3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = "C:\\тз акелон\\Практическое задание для кандидата.xlsx";
            ExcelDocument excelDocument = new ExcelDocument(filePath);
            MainMenu mainMenu = new MainMenu(excelDocument);
            mainMenu.Execute();
            //  = GetFilePath();

            Console.ReadKey();
        }

        /// <summary>
        /// Получение пути к файлу таблицы из консоли
        /// </summary>
        /// <returns>Путь до файла таблицы</returns>
        private static string GetFilePath()
        {
            while (true)
            {
                Console.WriteLine("Введите путь до таблицы:");
                string? filePath = Console.ReadLine();
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Такого файла не существует");
                    continue;
                }
                if (!Path.GetExtension(filePath).Equals("xls", StringComparison.CurrentCultureIgnoreCase) ||
                    !Path.GetExtension(filePath).Equals("xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    Console.WriteLine("Файл имеет неверный формат");
                    continue;
                }

                return filePath;
            }
        }
    }

    interface IMenuItem
    {
        string Title { get; }
        void Execute();
    }

    /// <summary>
    /// Класс пункта меню
    /// </summary>
    class MenuItem : IMenuItem
    {
        public string Title { get; private set; }
        private Action action;

        public MenuItem(string title, Action action)
        {
            Title = title;
            this.action = action;
        }

        public void Execute()
        {
            action?.Invoke();
        }
    }

    /// <summary>
    /// Класс подменю
    /// </summary>
    class SubMenu : IMenuItem
    {
        public string Title { get; private set; }
        private List<IMenuItem> items = new List<IMenuItem>();

        public SubMenu(string title)
        {
            Title = title;
        }

        public void AddItem(IMenuItem item)
        {
            items.Add(item);
        }

        public void Execute()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine($"=== {Title} ===");

                for (int i = 0; i < items.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {items[i].Title}");
                }

                Console.WriteLine($"{items.Count + 1}. Вернуться");

                string? choice = Console.ReadLine();

                if (int.TryParse(choice, out int index) && index >= 1 && index <= items.Count)
                {
                    items[index - 1].Execute();
                }
                else if (index == items.Count + 1)
                {
                    return;
                }
                else
                {
                    Console.WriteLine("Некорректный ввод");
                }
            }
        }
    }

    /// <summary>
    /// Класс главного меню
    /// </summary>
    class MainMenu
    {
        private List<IMenuItem> items = new List<IMenuItem>();
        private ExcelDocument excelDocument;

        public MainMenu(ExcelDocument excelDoc)
        {

            InitializeMenu();
            excelDocument = excelDoc;
        }

        private void InitializeMenu()
        {
            items.Add(new MenuItem("Информация о клиентах заказавших товар", GetClientsInfoByProductName));
            items.Add(new MenuItem("Изменить контактное лицо клиента", ChangeClientInfo));

            SubMenu subMenu = new SubMenu("Определить золотого клиента");
            subMenu.AddItem(new MenuItem("За месяц", GetGoldenClientByMonth));
            subMenu.AddItem(new MenuItem("За год", GetGoldenClientByYear));

            items.Add(subMenu);

            items.Add(new MenuItem("Выйти", () => Environment.Exit(0)));
        }

        public string Title => "Главное меню";

        public void Execute()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine($"=== {Title} ===");

                for (int i = 0; i < items.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {items[i].Title}");
                }

                string? choice = Console.ReadLine();

                if (int.TryParse(choice, out int index) && index >= 1 && index <= items.Count)
                {
                    items[index - 1].Execute();
                }
                else
                {
                    Console.WriteLine("Некорректный ввод");
                }
            }
        }

        /// <summary>
        /// 1. По наименованию товара выводить информацию о клиентах, заказавших этот товар, 
        /// с указанием информации по количеству товара, цене и дате заказа.
        /// </summary>
        private void GetClientsInfoByProductName()
        {
            Console.WriteLine("Введите наименование товара");
            string? productName = Console.ReadLine();
            if (productName == null)
            {
                return;
            }

            var infosList = excelDocument.GetClientsInfoByProductName(productName);
            if (infosList == null)
            {
                Console.WriteLine("Такой товар не найден");
                Console.ReadKey();
                return;
            }

            foreach (var info in infosList)
            {
                Console.WriteLine($"Клиент: {info.ClientText}\nКоличество товара: {info.Quantity}\nЦена заказа: {info.Cost}\nДата заказа: {DateTime.Parse(info.Date).ToShortDateString()}");
            }

            Console.ReadKey();
        }

        /// <summary>
        /// 2. Запрос на изменение контактного лица клиента с указанием параметров: 
        /// Название организации, ФИО нового контактного лица. 
        /// Информация заносится в документ, в качестве ответа пользователю необходимо выдавать информацию о результате изменений.
        /// </summary>
        private void ChangeClientInfo()
        {
            Console.WriteLine("Введите название организации: ");
            string? orgName = Console.ReadLine();
            if (orgName == null)
            {
                return;
            }

            Console.WriteLine("Введите новое контактное лицо: ");
            string? newContactName = Console.ReadLine();
            if (newContactName == null)
            {
                return;
            }

            var changingResult = excelDocument.ChangeClientsContactInfo(orgName, newContactName);
            Console.WriteLine(changingResult.Message);
            if (changingResult.IsSuccess)
            {
                var savingResult = excelDocument.SaveClientInfoFromDataSetToDocument();
                Console.WriteLine(savingResult.Message);
            }
            Console.ReadKey();
        }

        /// <summary>
        /// подменю Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный месяц.
        /// </summary>
        private void GetGoldenClientByMonth()
        {
            Console.WriteLine("Запрос на определение золотого клиента, месяц");
            Console.ReadKey();
        }

        /// <summary>
        /// подменю Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год.
        /// </summary>
        private void GetGoldenClientByYear()
        {
            Console.WriteLine("Запрос на определение золотого клиента, год");
            Console.ReadKey();
        }
    }

    /// <summary>
    /// Класс для взаимодействия с таблицей и работы с данными
    /// </summary>
    class ExcelDocument
    {
        private string FilePath { get; set; }

        private DataSet? dataSet;

        public ExcelDocument(string filePath)
        {
            FilePath = filePath;
            Initialize();
        }

        private void Initialize()
        {
            dataSet = ConvertExcelToDataSet();
            AddRelationsToDataSet();
        }

        public DataSet? ConvertExcelToDataSet()
        {
            DataSet dataSet = new DataSet();
            try
            {
                using (var workBook = new XLWorkbook(FilePath))
                {
                    foreach (var worksheet in workBook.Worksheets)
                    {
                        // Создаем DataTable для каждого листа Excel
                        var dataTable = new DataTable(worksheet.Name);

                        // Добавляем колонки в DataTable
                        foreach (IXLColumn column in worksheet.ColumnsUsed())
                        {
                            dataTable.Columns.Add(column.CellsUsed().First().Value.ToString(), typeof(string));
                        }

                        // Заполняем DataTable данными из листа Excel
                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем первую строку, если она содержит заголовки
                        {
                            var dataRow = dataTable.NewRow();
                            foreach (var cell in row.CellsUsed())
                            {
                                dataRow[cell.Address.ColumnNumber - 1] = cell.Value.ToString();
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        // Добавляем DataTable в DataSet
                        dataSet.Tables.Add(dataTable);
                    }

                    return dataSet;
                }
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
                Console.ReadKey();
            }

            return null;
        }

        private void AddRelationsToDataSet()
        {
            if (dataSet == null)
            {
                return;
            }

            DataRelation productToRequest = new DataRelation("productRelation", dataSet.Tables["Товары"].Columns["Код товара"], dataSet.Tables["Заявки"].Columns["Код товара"]);
            DataRelation clientToRequest = new DataRelation("clientRelation", dataSet.Tables["Клиенты"].Columns["Код клиента"], dataSet.Tables["Заявки"].Columns["Код клиента"]);
            dataSet.Relations.Add(productToRequest);
            dataSet.Relations.Add(clientToRequest);
        }

        public List<ProductRequestInfo>? GetClientsInfoByProductName(string productName)
        {
            List<ProductRequestInfo> requestInfos = new List<ProductRequestInfo>();

            var productRow = dataSet.Tables["Товары"].Select($"Наименование = '{productName}'").FirstOrDefault();
            if (productRow == null)
            {
                return null;
            }

            var requestRows = productRow.GetChildRows("productRelation");
            if (requestRows == null)
            {
                return null;
            }

            foreach (var requestRow in requestRows)
            {
                var clientRow = requestRow.GetParentRow("clientRelation");
                string clientText = $"{clientRow["Наименование организации"]}  {clientRow["Адрес"]} {clientRow["Контактное лицо (ФИО)"]}";
                string requestDate = requestRow["Дата размещения"].ToString();
                int quantity;
                if (!int.TryParse(requestRow["Требуемое количество"].ToString(), out quantity))
                {
                    continue;
                }
                string stringPrice = productRow["Цена товара за единицу"].ToString();
                if (!float.TryParse(stringPrice.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out float price))
                {
                    continue;
                }

                var info = new ProductRequestInfo(clientText, price * quantity, quantity, requestDate);
                requestInfos.Add(info);
            }


            return requestInfos;
        }

        public Result ChangeClientsContactInfo(string orgName, string newContactName)
        {
            Result result = new Result(false, "Непредвиденная ошибка");
            var clientRow = dataSet.Tables["Клиенты"].Select($"[Наименование организации] = '{orgName}'").FirstOrDefault();
            if (clientRow == null)
            {
                result.Message = "Такая организация не найдена";
                return result;
            }

            clientRow["Контактное лицо (ФИО)"] = newContactName;

            result.IsSuccess = true;
            result.Message = "Контактное лицо успешно изменено!";

            return result;
        }

        public Result SaveClientInfoFromDataSetToDocument()
        {
            Result result = new Result(false, null);

            try
            {
                using (XLWorkbook workBook = new XLWorkbook(FilePath))
                {
                    var dataTable = dataSet.Tables["Клиенты"];
                    var worksheet = workBook.Worksheet("Клиенты");

                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 1; col < dataTable.Columns.Count; col++)
                        {
                            worksheet.Cell(row + 2, col + 1).Value = dataTable.Rows[row][col].ToString();
                        }
                    }

                    workBook.Save();
                }
            }
            catch (Exception e)
            {
                result.Message = $"Произошла ошибка при записи данных в файл: {e.Message}";
                return result;
            }

            result.IsSuccess = true;
            result.Message = "Изменения в документ прошли успешно!";

            return result;
        }
    }

    /// <summary>
    /// Класс результата выполнения операции
    /// </summary>
    class Result
    {
        public bool IsSuccess { get; set; }
        public string? Message { get; set; }

        public Result(bool isSuccess, string? message) 
        {
            IsSuccess = isSuccess;
            Message = message;
        }
    }

    /// <summary>
    /// Класс для хранения информации о заказе
    /// </summary>
    class ProductRequestInfo
    {
        public string ClientText { get; set; }
        public float Cost { get; set; }
        public int Quantity { get; set; }
        public string Date { get; set; }

        public ProductRequestInfo(string clientText, float cost, int quantity, string date)
        {
            ClientText = clientText;
            Cost = cost;
            Quantity = quantity;
            Date = date;
        }

        public ProductRequestInfo()
        {
        }
    }
}
