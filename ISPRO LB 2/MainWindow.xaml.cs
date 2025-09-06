using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using ClosedXML.Excel;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Newtonsoft.Json;

namespace ISPRO_LB_2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Orders _currentMaterial = new Orders();
        public MainWindow()
        {
            InitializeComponent();

            DataContext = _currentMaterial;
            Data1.ItemsSource = ISPRO2Entities.GetContext().Orders.OrderBy(x => x.Id).ToList();
        }


        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog ofd = new SaveFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Выберите место для сохранения файла"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            var wb = new XLWorkbook();
            using (var context = ISPRO2Entities.GetContext())
            {
                var orders = context.Orders.ToList();

                var groupedByStatus = orders.GroupBy(o => o.Status ?? "Unknown");

                foreach (var group in groupedByStatus)
                {
                    var statusSheet = wb.Worksheets.Add(group.Key); 

                
                    statusSheet.Cell(1, 1).Value = "ID";
                    statusSheet.Cell(1, 2).Value = "Order Code";
                    statusSheet.Cell(1, 3).Value = "Client Code";
                    statusSheet.Cell(1, 4).Value = "Services";
                   
                    for (int i = 0; i < group.Count(); i++)
                    {
                        var order = group.ElementAt(i);
                        statusSheet.Cell(i + 2, 1).Value = order.Id;
                        statusSheet.Cell(i + 2, 2).Value = order.CodeOrder;
                        statusSheet.Cell(i + 2, 3).Value = order.CodeClient;
                        statusSheet.Cell(i + 2, 4).Value = order.Services;
                     
                    }
                }
            }

            try
            {
                
                wb.SaveAs(ofd.FileName);
                MessageBox.Show("Файл успешно экспортирован!", "Успешно", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }




        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            try
            {
                using (var workbook = new XLWorkbook(ofd.FileName))
                {
                    var worksheet = workbook.Worksheet(1);
                    var range = worksheet.RangeUsed();
                    var rows = range.RowsUsed().Skip(1); 

                    List<Orders> ordersList = new List<Orders>();

                    foreach (var row in rows)
                    {
                        ordersList.Add(new Orders()
                        {
                            Id = int.Parse(row.Cell(1).Value.ToString()),
                            CodeOrder = row.Cell(2).Value.ToString(),
                            CreaateDate = row.Cell(3).Value.ToString(),
                            CreateTime = row.Cell(4).Value.ToString(),
                            CodeClient = row.Cell(5).Value.ToString(),
                            Services = row.Cell(6).Value.ToString(),
                            Status = row.Cell(7).Value.ToString(),
                            ClosedDate = row.Cell(8).Value.ToString(),
                            ProkatTime = row.Cell(9).Value.ToString()
                        });
                    }

                    using (ISPRO2Entities usersEntities = new ISPRO2Entities())
                    {
         

                        foreach (var order in ordersList)
                        {
                            usersEntities.Orders.Add(order); 
                            usersEntities.SaveChanges(); 
                        }
                    }

                    MessageBox.Show("Данные успешно импортировались");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BnImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файл (*.json)|*.json",
                Title = "Выберите JSON файл для импорта"
            };

            if (!(ofd.ShowDialog() == true))
                return;

           
                string jsonContent = System.IO.File.ReadAllText(ofd.FileName);

               
                List<Orders> ordersList = JsonConvert.DeserializeObject<List<Orders>>(jsonContent);

                if (ordersList == null || !ordersList.Any())
                {
                    MessageBox.Show("Файл не содержит данных или данные некорректны.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (ISPRO2Entities usersEntities = new ISPRO2Entities())
                {
                    foreach (var order in ordersList)
                    {
                        usersEntities.Orders.Add(order);
                    }
                    usersEntities.SaveChanges(); 
                }

                MessageBox.Show("Данные успешно импортировались", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            
            
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            
            SaveFileDialog ofd = new SaveFileDialog()
            {
                DefaultExt = "docx",
                Filter = "Word Files (*.docx)|*.docx",
                Title = "Выберите место для сохранения файла"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            try
            {
                using (var context = ISPRO2Entities.GetContext())
                {
                    var orders = context.Orders.ToList();
                    var groupedByStatus = orders.GroupBy(o => o.Status ?? "Unknown");

                    
                    using (var doc = DocX.Create(ofd.FileName))
                    {
                        foreach (var group in groupedByStatus)
                        {
                           
                            var statusHeader = doc.InsertParagraph($"Статус: {group.Key}", true);
                            statusHeader.FontSize(14).Bold();

                          
                            var table = doc.AddTable(group.Count() + 1, 4);
                            table.Design = TableDesign.TableGrid;

                          
                            var headerRow = table.Rows[0];
                            headerRow.Cells[0].Paragraphs.First().Append("ID");
                            headerRow.Cells[1].Paragraphs.First().Append("Order Code");
                            headerRow.Cells[2].Paragraphs.First().Append("Client Code");
                            headerRow.Cells[3].Paragraphs.First().Append("Services");

                           
                            int rowIndex = 1;
                            foreach (var order in group)
                            {
                                var row = table.Rows[rowIndex];
                                row.Cells[0].Paragraphs.First().Append(order.Id.ToString());
                                row.Cells[1].Paragraphs.First().Append(order.CodeOrder);
                                row.Cells[2].Paragraphs.First().Append(order.CodeClient);
                                row.Cells[3].Paragraphs.First().Append(order.Services);

                                rowIndex++;
                            }

                            
                            doc.InsertTable(table);
                            doc.InsertParagraph(); 
                        }

                      
                        doc.Save();
                    }

                    
                    MessageBox.Show("Файл успешно экспортирован!", "Успешно", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
               
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
    }

