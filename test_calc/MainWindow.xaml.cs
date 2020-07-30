using System;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace FirstWpfApp
{
    public partial class MainWindow : Window
    {
        string leftop = ""; // Левый операнд
        string operation = ""; // Знак операции
        string rightop = ""; // Правый операнд
        int a = 1; // Строка для записей
        int b = 0; // Help

        // Создаём экземпляр нашего приложения
        Excel.Application excelApp = new Excel.Application();
        // Создаём экземпляр рабочий книги Excel
        Excel.Workbook workBook;
        // Создаём экземпляр листа Excel
        Excel.Worksheet workSheet;


        void DataWindow_Closing(object sender, CancelEventArgs e)
        {
            workBook.Close();
        }
        public MainWindow()
        {
            InitializeComponent();
            plus.IsEnabled = false;
            minus.IsEnabled = false;
            star.IsEnabled = false;
            slash.IsEnabled = false;
            res.IsEnabled = false;
            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            // Открываем созданный excel-файл
            excelApp.Visible = false;
            excelApp.UserControl = false;
            var docPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            System.IO.File.Delete(@docPath + @"\test_calc12345.xlsx");
            workBook.SaveAs(@docPath + @"\test_calc12345.xlsx");

            // Добавляем обработчик для всех кнопок на гриде
            foreach (UIElement c in LayoutRoot.Children)
            {
                if (c is Button)
                {
                    ((Button)c).Click += Button_Click;
                }
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Получаем текст кнопки
            string s = (string)((Button)e.OriginalSource).Content;
            // Добавляем его в текстовое поле
            textBlock.Text += s;
            int num;
            // Пытаемся преобразовать его в число
            bool result = Int32.TryParse(s, out num);
            // Если текст - это число
            if (result == true)
            {
                // Если операция не задана
                if (operation == "")
                {
                    // Добавляем к левому операнду
                    leftop += s;
                    plus.IsEnabled = true;
                    minus.IsEnabled = true;
                    star.IsEnabled = true;
                    slash.IsEnabled = true;
                }
                else
                {
                    // Иначе к правому операнду
                    rightop += s;
                }
            }
            // Если было введено не число
            else
            {
                // Если равно, то выводим результат операции
                if (s == "=")
                {
                    Update_RightOp();
                    textBlock.Text = rightop;
                    workSheet.Cells[a, 1] = leftop;
                    workSheet.Cells[a, 2] = operation;
                    workSheet.Cells[a, 3] = b;
                    workSheet.Cells[a, 4] = Convert.ToDouble(rightop);
                    workSheet.Cells[a, 6] = DateTime.Now.ToString();
                    operation = "";
                    one.IsEnabled = false;
                    two.IsEnabled = false;
                    three.IsEnabled = false;
                    four.IsEnabled = false;
                    five.IsEnabled = false;
                    six.IsEnabled = false;
                    seven.IsEnabled = false;
                    eight.IsEnabled = false;
                    nine.IsEnabled = false;
                    zero.IsEnabled = false;
                    plus.IsEnabled = false;
                    minus.IsEnabled = false;
                    star.IsEnabled = false;
                    slash.IsEnabled = false;
                    res.IsEnabled = false;
                }
                // Очищаем поле и переменные
                else if (s == "CLEAR")
                {
                    leftop = "";
                    rightop = "";
                    operation = "";
                    textBlock.Text = "";
                    a += 1;
                    one.IsEnabled = true;
                    two.IsEnabled = true;
                    three.IsEnabled = true;
                    four.IsEnabled = true;
                    five.IsEnabled = true;
                    six.IsEnabled = true;
                    seven.IsEnabled = true;
                    eight.IsEnabled = true;
                    nine.IsEnabled = true;
                    zero.IsEnabled = true;
                    plus.IsEnabled = false;
                    minus.IsEnabled = false;
                    star.IsEnabled = false;
                    slash.IsEnabled = false;
                    res.IsEnabled = false;
                    
                    workBook.Save();

                }
                // Получаем операцию
                else
                {
                    plus.IsEnabled = false;
                    minus.IsEnabled = false;
                    star.IsEnabled = false;
                    slash.IsEnabled = false;
                    res.IsEnabled = true;
                    // Если правый операнд уже имеется, то присваиваем его значение левому
                    // операнду, а правый операнд очищаем
                    if (rightop != "")
                    {
                        Update_RightOp();
                        leftop = rightop;
                        rightop = "";
                    }
                    operation = s;
                }
            }
        }
        
        // Обновляем значение правого операнда
        private void Update_RightOp()
        {
            int num1 = Int32.Parse(leftop);
            int num2 = Int32.Parse(rightop);
            b = num2;
            // И выполняем операцию
            switch (operation)
            {
                case "+":
                    rightop = (num1 + num2).ToString();
                    break;
                case "-":
                    rightop = (num1 - num2).ToString();
                    break;
                case "*":
                    rightop = (num1 * num2).ToString();
                    break;
                case "/":
                    if (num2 == 0)
                    {
                        textBlock.Text = "error";
                        rightop = "error";
                        break;
                    }
                    else {
                        rightop = ((double)num1 / num2).ToString();
                    }
                    break;
            }
        }
    }
}