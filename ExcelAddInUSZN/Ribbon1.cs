using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;                    // пространство имен в .NET Framework, которое содержит классы для управления приложением Microsoft Word. Эти классы предоставляют функциональность для взаимодействия с документами Word, включая создание, чтение, редактирование и форматирование документов.
using Excel = Microsoft.Office.Interop.Excel;           // эта строка создает псевдоним Excel для пространства имен Microsoft.Office.Interop.Excel
using Word = Microsoft.Office.Interop.Word;             // эта строка создает псевдоним Excel для пространства имен Word Microsoft.Office.Interop.Word
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms.VisualStyles;
using System.Text.RegularExpressions;
using System.Security.Cryptography.X509Certificates;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using System.Runtime.InteropServices;
using System.IO;


namespace ExcelAddInUSZN
{
    // Определение класса Ribbon1
    public partial class Ribbon1
    {
        // Метод, который вызывается при загрузке ленты
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Здесь можно добавить код, который будет выполняться при загрузке ленты
        }
        // Обработчик события нажатия на кнопку
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            sample_to_use();
            //action();    
            MessageBox.Show("Процесс завершен!");
        }

        private void sample_to_use()
        {
            // Переменные для Таблицы 4
            double time_in_minutes;

            // переменные для раздела 3
            double minutes;

            // Получение активного листа в Excel
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            // Создание нового экземпляра приложения Word
            var wordApp = new Word.Application();
            // Отображение приложения Word
            wordApp.Visible = true;

            // Добавление нового документа в приложение Word
            Document wordDoc = wordApp.Documents.Add();

            // Активация нового документа
            wordDoc.Activate();

            // Настройка параметров страницы

            wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2); // Преобразование сантиметров в точки
            wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1); // Преобразование сантиметров в точки
            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f); // Преобразование сантиметров в точки
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.9f); // Преобразование сантиметров в точки

            ContentDocument cd = new ContentDocument();

            // данные к пункту 3
            minutes = Math.Round(activeWorksheet.Range["E42"].Value2);
            TimeSpan hours_minutes = TimeSpan.FromDays(activeWorksheet.Range["E43"].Value2);
            cd.minutes = sing_time((string)minutes.ToString(), false);
            cd.hours_minutes = sing_time((string)hours_minutes.ToString());

            // данные к таблице 4 пункта 4.5
            time_in_minutes = Math.Round(activeWorksheet.Range["E44"].Value2);
            TimeSpan time_in_hours = TimeSpan.FromDays(activeWorksheet.Range["E45"].Value2);
            cd.time_in_minutes = sing_time2((string)time_in_minutes.ToString(), false);
            cd.time_in_hours = sing_time2((string)time_in_hours.ToString());
            // 
            string[] footnotes = cd.create_footnotes();
            char[] symbols = new char[] { '\u00B9', '\u00B2', '\u00B3', '\u2074', '\u2075', '\u2076' };
            // Указываем диапазоны для считывания данных
            List<string> excelRanges = new List<string>
            {
                "B5:B37",
                "U5:U37",
                "AN5:AN37",
                "BG5:BG37",
                "BZ5:BZ37"
            };

            HashSet<string> hashSet_excelRanges = Create_a_list_table_ServicesProvided(activeWorksheet, excelRanges);
            cd.count_servise = hashSet_excelRanges.Count.ToString();
            // Создадим список кортежей для объединения ячеек
            List<Tuple<int, int>> mergeCellsTab4 = new List<Tuple<int, int>>
            {
                new Tuple<int, int>(3, 2),     // Объединить ячейку в третьей строке, первом столбце
                new Tuple<int, int>(3, 3)      // Объединить ячейку в третьей строке, втором столбце
            };




            //Создание временных таблиц

            System.Data.DataTable dt1 = cd.CreateTable1();
            System.Data.DataTable dt2 = cd.CreateTable2();
            System.Data.DataTable dt3 = cd.CreateTable3();
            System.Data.DataTable dt4 = cd.CreateTable4();
            System.Data.DataTable dt5 = cd.CreateTable5();
            System.Data.DataTable dt7 = cd.CreateTable7();
            System.Data.DataTable dt8 = cd.CreateTable8();
            System.Data.DataTable dt9 = cd.CreateTable9();

            // Создание структуры документа
            List<System.Action> documentStructure = new List<System.Action>
            {
                // Заголовок
                () => InsertText(wordDoc, cd.heading, false,"Times New Roman", Word.WdParagraphAlignment.wdAlignParagraphCenter,14, 1),
                //Таблица1 (Дата составления, Номер регистрации, Статус)
                () => CreateTableAndInsert(wordDoc,  dt1,false,null,"Times New Roman", 8,Word.WdParagraphAlignment.wdAlignParagraphCenter,Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,true ),
                // Фамилия
                () => InsertText(wordDoc, cd.surname),
                // Имя
                () => InsertText(wordDoc, cd.name),
                // Отчество
                () => InsertText(wordDoc, cd.patronymic,true),
                // Снилс
                () => InsertText(wordDoc,cd.d_r_gender_SNILS),
                // Часть1 Заголовок
                () => InsertText(wordDoc,cd.heading_aragraph,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphCenter,14,1),
                // пункт 1
                () => InsertText(wordDoc,cd.paragraph1),
                // пункт 2
                () => InsertText(wordDoc,cd.paragraph2,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphJustify),
                // пункт 3
                () => InsertText(wordDoc,cd.timeInString(),false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphJustify),
                // пункт 4
                () => InsertText(wordDoc,cd.paragraph4),
                // подпункт 4.1
                () => InsertText(wordDoc,cd.paragraph4_1,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphJustify),
                // подпункт 4.2
                () => InsertText(wordDoc,cd.paragraph4_2,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphJustify),
                // Таблица 2  
                () => CreateTableAndInsert(wordDoc, dt2,true),
                // подпункт 4.3
                () => InsertText(wordDoc,cd.paragraph4_3),
                // Таблица 3
                () => CreateTableAndInsert(wordDoc, dt3,true),
                // подпункт 4.4
                () => InsertText(wordDoc,cd.paragraph4_4),
                // Таблица 1 неделя 1
                () => CopyExcelTableToWord(activeWorksheet, wordDoc, "B1:S38"),
                // Таблица 2 неделя 2
                () => CopyExcelTableToWord(activeWorksheet, wordDoc, "U1:AL38"),
                // Таблица 3 неделя 3
                () => CopyExcelTableToWord(activeWorksheet, wordDoc, "AN1:BE38"),
                // Таблица 4 неделя 4
                () => CopyExcelTableToWord(activeWorksheet, wordDoc, "BG1:BX38"),
                // Таблица 5 неделя 5
                () => CopyExcelTableToWord(activeWorksheet, wordDoc, "BZ1:CQ38"),
                // примечание к таблицам недель1-5
                () => InsertText(wordDoc,cd.explanation,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphLeft,9),
                () => ClearClipboard(),
                // подпункт 4.5
                () => InsertText(wordDoc,cd.paragraph4_5),
                // таблица 4
                () => CreateTableAndInsert(wordDoc,dt4,true,mergeCells:mergeCellsTab4),
                // пункт 5
                () => InsertText(wordDoc,cd.paragraph5),
                // Таблица 5
                () => CreateTableAndInsert(wordDoc,dt5,true),
                // пункт 6
                () => InsertText(wordDoc,cd.paragraph6),
                // таблица услуг которые не требуются переписать 
                () => Create_a_list_table_not_included(activeWorksheet, wordDoc,hashSet_excelRanges),
                // пункт 7
                () => InsertText(wordDoc,cd.paragraph7),
                // пункт 8
                () => InsertText(wordDoc,cd.paragraph8),
                // подвал документа 
                () => InsertText(wordDoc, cd.text_pered_podpis1),
                () => CreateTableAndInsert(wordDoc,dt7,false,null,"Times New Roman", 8,Word.WdParagraphAlignment.wdAlignParagraphCenter,Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,true ),
                () => InsertText(wordDoc, cd.text_pered_podpis2),
                () => CreateTableAndInsert(wordDoc,dt8,false, null,"Times New Roman", 8,Word.WdParagraphAlignment.wdAlignParagraphCenter,Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,true ),
                () => CreateTableAndInsert(wordDoc,dt9,false, null,"Times New Roman", 8,Word.WdParagraphAlignment.wdAlignParagraphCenter,Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,true ),
                () => ProcessDocument2(wordDoc,symbols,footnotes)
            };

            // Вставка структуры документа в Word
            foreach (System.Action action in documentStructure)
            {
                action();
            }

            Marshal.ReleaseComObject(wordApp);


            GetTableWidth(wordDoc, 4);

            System.Media.SystemSounds.Asterisk.Play();
        }
        

        private void CopyExcelTableToWord(Excel.Worksheet activeWorksheet, Document wordDoc, string range, string styleName = "Times New Roman", float fontSize = 10)
        {
            Excel.Range excelRange = activeWorksheet.Range[range, Type.Missing];
            object[,] values = excelRange.Value2;

            // Определите количество непустых строк
            int rowCount = 0;
            for (int i = 1; i <= excelRange.Rows.Count; i++)
            {
                if (i <= 4 || !string.IsNullOrWhiteSpace(values[i, 1]?.ToString()))
                {
                    rowCount++;
                }
            }

            // Получите конечную позицию содержимого в документе Word
            Word.Range endRange = wordDoc.Content;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // Создайте новую таблицу в Word с количеством строк, равным количеству непустых строк, и количеством столбцов, равным количеству столбцов в Excel
            //Word.Table wordTable = wordDoc.Tables.Add(wordDoc.Range(wordDoc.Content.End - 1, wordDoc.Content.End - 1), rowCount, excelRange.Columns.Count);
            // Создайте новую таблицу в Word и добавьте ее к концу содержимого документа
            //Word.Table wordTable = wordDoc.Tables.Add((Word.Range)endRange, rowCount, excelRange.Columns.Count);
            Word.Table wordTable = wordDoc.Tables.Add(endRange, rowCount, excelRange.Columns.Count);
            // Добавьте границы к таблице
            Word.Borders borders = wordTable.Borders;
            borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            // Заполните таблицу данными из массива values, пропуская пустые строки
            int rowIndex = 1;
            for (int i = 1; i <= excelRange.Rows.Count; i++)
            {
                if (i <= 4 || !string.IsNullOrWhiteSpace(values[i, 1]?.ToString()))
                {
                    for (int j = 1; j <= excelRange.Columns.Count; j++)
                    {
                        Word.Cell cell = wordTable.Cell(rowIndex, j);
                        StringBuilder sb = new StringBuilder(values[i, j]?.ToString() ?? "");
                        cell.Range.Text = sb.ToString();
                    }
                    rowIndex++;
                }
            }

            // Объединение ячеек в соответствии с заданными условиями
            wordTable.Cell(1, 1).Merge(wordTable.Cell(1, 18));
            wordTable.Cell(1, 1).Range.Bold = 1; // Выделение текста первой строки жирным
            wordTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //объединение последнего столбика 2-4 строки
            wordTable.Cell(2, 18).Merge(wordTable.Cell(4, 18));
            // объединение ячеек с 4-17 столбик 3 и 4 строки
            for (int i = 4; i <= 17; i++)
            {
                wordTable.Cell(3, i).Merge(wordTable.Cell(4, i));
            }
            // объединение 1 столбика 2-4 строки 
            wordTable.Cell(2, 1).Merge(wordTable.Cell(4, 1));

            // Объединение четырех ячеек
             wordTable.Cell(2, 2).Merge(wordTable.Cell(3, 3));
            // повернуть текст на 90 градусов
            for (int i = 3; i <= 16; i++)
            {
                wordTable.Cell(3, i).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
            }
            wordTable.Cell(2, 17).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;

            // объединение столбцов во второй строке  с 5 по 16 столбик
            wordTable.Cell(2, 15).Merge(wordTable.Cell(2, 16));
            wordTable.Cell(2, 13).Merge(wordTable.Cell(2, 14));
            wordTable.Cell(2, 11).Merge(wordTable.Cell(2, 12));
            wordTable.Cell(2, 9).Merge(wordTable.Cell(2, 10));
            wordTable.Cell(2, 7).Merge(wordTable.Cell(2, 8));
            wordTable.Cell(2, 6).Merge(wordTable.Cell(2, 5));
            wordTable.Cell(2, 4).Merge(wordTable.Cell(2, 3));

            // последняя строка таблицы выделить жирным
            int lastRow = wordTable.Rows.Count;
            for (int i = 1; i <= wordTable.Columns.Count; i++)
            {
                wordTable.Cell(lastRow, i).Range.Font.Bold = 1;
            }

            // Получаем ширину листа
            float pageWidth = wordDoc.PageSetup.PageWidth - wordDoc.PageSetup.LeftMargin - wordDoc.PageSetup.RightMargin;

            // Устанавливаем ширину таблицы
            wordTable.PreferredWidth = pageWidth;
            wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow); // Для автоматического подгонки ширины столбцов

            Word.Font tableFont = wordTable.Range.Font;
            tableFont.Name = styleName; // Установка имени шрифта
            tableFont.Size = fontSize;  // Установка размера шрифта

           

            // Добавляем новый абзац в конец документа
            Word.Paragraph para = wordDoc.Content.Paragraphs.Add();
            para.Range.Text = "\n";
            para.Range.InsertParagraphAfter();

            // Освобождение ресурсов
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(wordTable);
        }



        // Определение услуг которые предоставляются переписать на метод с возвращающимся значением
        private HashSet<string> Create_a_list_table_ServicesProvided(Worksheet activeWorksheet, List<string> ranges)
        {
            // Создается HashSet для хранения уникальных данных.
            HashSet<string> uniqueData = new HashSet<string>();
            foreach (string range in ranges)
            {
                // Для каждого диапазона в списке диапазонов он получает соответствующий диапазон из листа Excel.
                Excel.Range excelRange = activeWorksheet.Range[range];

                // Извлекаем уникальные данные из каждого диапазона
                foreach (Excel.Range cell in excelRange)
                {
                    // Проверяем, не пустая ли ячейка
                    string cellValue = cell.Value2;
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        uniqueData.Add(RemoveNumbersAtStart(cellValue.ToString()));
                    }
                }
            }
            return uniqueData;
        }

        // определение услуг которые не предоставляются
        private void Create_a_list_table_not_included(Worksheet activeWorksheet, Document wordDoc, HashSet<string> uniqueData)
        {
            string rangeBase = "B56:B105";
            List<string> difference = new List<string>();
            HashSet<string> dataBase = new HashSet<string>();
            // Эта функция создает таблицу в документе Word из yt уникальных данных в листе Excel.
            Word.Range rng = wordDoc.Content;
            // Устанавливает точку вставки в конец документа Word.
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            try
            {
                // определение общего списка услуг
                Excel.Range excelRangeBase = activeWorksheet.Range[rangeBase];
                foreach (Excel.Range cell in excelRangeBase)
                {
                    // Проверяем, не пустая ли ячейка
                    string cellValue = cell.Value2;
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        dataBase.Add(RemoveNumbersAtStart(cellValue.ToString()));
                    }
                }
                // определяем услуги не вошедшие в перечень предоставляемых услуг
                difference.AddRange(dataBase.Except(uniqueData));

                // Создает таблицу в документе Word с числом строк, равным числу уникальных данных плюс 2 (для заголовка и подвала), и 2 столбца.
                Table wordTable = wordDoc.Tables.Add(rng, difference.Count + 2, 2);
                wordTable.Borders.Enable = 1; // Включаем границы таблицы

                // Объединяем ячейки в шапке
                wordTable.Cell(1, 1).Merge(wordTable.Cell(1, 2));
                // Заполняем шапку таблицы
                wordTable.Cell(1, 1).Range.Text = "Наименование социальной услуги по уходу";
                // Заполняем строки данными из списка List difference.
                int rowIndex = 2;
                foreach (string data in difference)
                {
                    // Объединяем ячейки таблицы
                    wordTable.Cell(rowIndex, 1).Merge(wordTable.Cell(rowIndex, 2));
                    wordTable.Cell(rowIndex, 1).Range.Text = data;
                    rowIndex++;
                }
                // Заполняем подвал таблицы ?
                wordTable.Cell(difference.Count + 2, 1).Range.Text = "Общее количество социальных услуг по уходу, не включенных в социальный пакет долговременного ухода" + '\u2075';
                wordTable.Cell(difference.Count + 2, 2).Range.Text = difference.Count.ToString();
            }
            catch (Exception e)
            {
                // Если возникает ошибка, она отображается в диалоговом окне с сообщением об ошибке.
                //MessageBox.Show($"Ошибка: {e.Message}");
            }
        }

        public int ColumnLetterToNumber(string columnLetter)
        {
            int columnNumber = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                columnNumber *= 26;
                columnNumber += columnLetter[i] - 'A' + 1;
            }
            return columnNumber;
        }
        private void InsertText(Document wordDoc, string text, bool nullString = false, string fontName = "Times New Roman",
                                    Word.WdParagraphAlignment alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                                    int fontSize = 12, int bold = 0)
        {
            //Word.Range rng = wordDoc.Content; - Получает содержимое всего документа Word и сохраняет его в переменную rng.
            Word.Range rng = wordDoc.Content;
            //“Схлопывает” диапазон до его конца. Это означает, что вместо того чтобы ссылаться на весь документ,
            //rng теперь ссылается только на конец документа.Это полезно, когда вы хотите вставить новый текст в конец документа.
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //Вставляет текст, переданный в метод, сразу после текущего диапазона(который теперь указывает на конец документа).
            rng.InsertAfter(text);
            if (nullString)
            {
                //  Вставляет новую строку после вставленного текста. Это создает пустую строку после каждого вставленного блока текста.
                rng.InsertAfter("\n");
            }
            //Устанавливает размер шрифта вставленного текста на значение, переданное в метод.
            rng.Font.Size = fontSize;
            //Устанавливает жирность шрифта вставленного текста.Если bold равно 1, текст будет жирным, если 0 - обычным.
            rng.Font.Bold = bold;
            rng.Font.Name = fontName;
            //Устанавливает выравнивание вставленного текста на значение, переданное в метод. Это может быть выравнивание по левому краю,
            //по центру или по правому краю.
            rng.ParagraphFormat.Alignment = alignment;
        }


        public void CreateTableAndInsert(Word.Document wordDoc, System.Data.DataTable dt, bool table_boundaries = true, List<Tuple<int, int>> mergeCells = null,
                     string fontName = "Times New Roman", int fontSize = 10,
                     Word.WdParagraphAlignment alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                     Word.WdCellVerticalAlignment verticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                     bool underline = false) // Добавлен новый параметр
        {
            // Создаем массив с заданными словами
            string[] words = new string[] { "№", "Статус", "М. П." };

            // Получаем текущий диапазон документа и перемещаем курсор в конец
            Word.Range rng = wordDoc.Content;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Добавляем таблицу в документ с количеством строк и столбцов, соответствующим DataTable
            Word.Table table = wordDoc.Tables.Add(rng, dt.Rows.Count, dt.Columns.Count);

            int row = 1;
            // Проходим по каждой строке DataTable
            foreach (DataRow dr in dt.Rows)
            {
                // Проходим по каждому столбцу DataTable
                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    // Получаем диапазон ячейки и устанавливаем текст, шрифт, размер шрифта, выравнивание и вертикальное выравнивание
                    Word.Range cellRange = table.Cell(row, col).Range;
                    cellRange.Text = dr[col - 1].ToString();
                    cellRange.Font.Name = fontName;
                    cellRange.Font.Size = fontSize;
                    cellRange.ParagraphFormat.Alignment = alignment;
                    table.Cell(row, col).VerticalAlignment = verticalAlignment;

                    // Если в ячейке есть текст и он не совпадает с заданными словами, то добавляем верхнюю границу
                    if (!string.IsNullOrEmpty(dr[col - 1].ToString()) && !words.Contains(dr[col - 1].ToString()))
                    {
                        table.Cell(row, col).Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }
                }
                row++;
            }

            // Если table_boundaries = true, то устанавливаем границы таблицы
            if (table_boundaries)
            {
                table.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                table.Borders.InsideColor = Word.WdColor.wdColorBlack;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }

            try
            {
                // Если mergeCells не null, то объединяем указанные ячейки
                if (mergeCells != null)
                {
                    foreach (var cell in mergeCells)
                    {
                        table.Cell(cell.Item1, cell.Item2).Merge(table.Cell(cell.Item1, cell.Item2 + 1));
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                // Обработка ошибки при объединении ячеек
            }

            // Добавляем новый абзац в конец документа
            Word.Paragraph para = wordDoc.Content.Paragraphs.Add();
            para.Range.Text = "\n";
            para.Range.InsertParagraphAfter();
        }

        public void GetTableWidth(Word.Document wordDoc, int tableNumber)
        {
            try
            {
                // Получаем таблицу по индексу (нумерация с 1)
                Table table = wordDoc.Tables[tableNumber];
                // Получаем ширину таблицы относительно страницы
                //float tableWidth = table.Range.Information[WdInformation.wdHorizontalPositionRelativeToPage];

                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

            }
            catch (Exception e)
            {
                // Обработка ошибки, если таблица не найдена или возникла другая проблема
                // Вместо этого можно вернуть float.MinValue или другое значение по умолчанию
                // в зависимости от требований вашего приложения.
                //MessageBox.Show($"Ошибка: {e.Message}");

            }
        }

        private string RemoveNumbersAtStart(string input)
        {
            // Удаляем числа и точку в начале строки
            return Regex.Replace(input.Trim(), @"^\d+\.?\s*", "");
        }

        private string sing_time(string time, bool hour = true)
        {
            string time_sing;
            if (hour)
            {
                // Разделить входную строку на часы, минуты и секунды
                string[] parts_of_the_time = time.Split(':');

                // Преобразовать значения строк в целые числа
                string hours = parts_of_the_time[0];
                string minutes = parts_of_the_time[1];

                // Форматировать выходную строку
                time_sing = $"{hours} часов {minutes} минут";
            }
            else
            {
                time_sing = $"{time} минут";
            }

            return time_sing;
        }
        private string sing_time2(string time, bool hour = true)
        {
            string time_sing;
            if (hour)
            {
                // Разделить входную строку на часы, минуты и секунды
                string[] parts_of_the_time = time.Split(':');

                // Преобразовать значения строк в целые числа
                string hours = parts_of_the_time[0];
                string minutes = parts_of_the_time[1];

                // Форматировать выходную строку
                time_sing = $"{hours},{minutes}";
            }
            else
            {
                time_sing = $"{time} ";
            }

            return time_sing;
        }
        
        private void AddFootnote(Word.Document wordDoc, Word.Range range, string footnoteText)
        {
            // Создаем сноску
            Word.Footnote footnote = wordDoc.Footnotes.Add(range, "", footnoteText);

            // Устанавливаем формат сноски
            footnote.Range.Font.Size = 8;
            footnote.Range.Font.Name = "Times New Roman";
        }
        public void ProcessDocument(Word.Document wordDoc, char[] symbols, string[] footnotes)
        {
            // Ищем в документе символы из массива symbols
            Word.Range range = wordDoc.Content; // Используем только основной текст документа
            for (int i = 0; i < symbols.Length; i++)
            {
                int start = 0;
                while (start < range.End)
                {
                    Word.Range searchRange = range.Duplicate;
                    searchRange.Start = start;
                    searchRange.Find.ClearFormatting();
                    searchRange.Find.Text = symbols[i].ToString();

                    if (searchRange.Find.Execute())
                    {
                        // Добавляем сноску
                        AddFootnote(wordDoc, searchRange, footnotes[i]);
                        start = searchRange.End;
                        
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        public void ProcessDocument2(Word.Document wordDoc, char[] symbols, string[] footnotes)
        {
            // Создаем список для отслеживания уже обработанных символов
            List<char> processedSymbols = new List<char>();

            // Ищем в документе символы из массива symbols
            Word.Range range = wordDoc.Content; // Используем только основной текст документа
            for (int i = 0; i < symbols.Length; i++)
            {
                // Если символ уже был обработан, пропускаем его
                if (processedSymbols.Contains(symbols[i]))
                {
                    continue;
                }

                int start = 0;
                bool symbolFound = false;
                while (start < range.End)
                {
                    Word.Range searchRange = range.Duplicate;
                    searchRange.Start = start;
                    searchRange.Find.ClearFormatting();
                    searchRange.Find.Text = symbols[i].ToString();

                    if (searchRange.Find.Execute())
                    {
                        // Если это первое обнаружение символа, добавляем сноску
                        if (!symbolFound)
                        {
                            searchRange.Text = ""; // Удаляем символ
                            AddFootnote(wordDoc, searchRange, footnotes[i]);
                            symbolFound = true;
                        }
                        start = searchRange.End;
                    }
                    else
                    {
                        break;
                    }
                }

                // Добавляем символ в список обработанных
                if (symbolFound)
                {
                    processedSymbols.Add(symbols[i]);
                }
            }
        }

        public void ClearClipboard()
        {
            Clipboard.Clear();
        }

        // Добавлено Создание таблиц по шаблону

        private void btnCreateTables_Click(object sender, RibbonControlEventArgs e)
        {
            using (Templates form = new Templates())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    string templateFilePath = form.SelectedTemplatePath;

                    if (!File.Exists(templateFilePath))
                    {
                        MessageBox.Show($"Файл шаблона не найден: {templateFilePath}");
                        return;
                    }

                    Excel.Application excelApp = Globals.ThisAddIn.Application;
                    Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;


                    CreateSheetsFromTemplateInActiveWorkbook2(templateFilePath);
                }
            }
        }
        private void CreateSheetsFromTemplateInActiveWorkbook2(string templateFilePath)
        {
            // Получаем текущее приложение Excel
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            // Открываем шаблон
            Excel.Workbook templateWorkbook = excelApp.Workbooks.Open(templateFilePath);

            try
            {
                // Копируем каждый лист шаблона в активную книгу
                foreach (Excel.Worksheet templateSheet in templateWorkbook.Sheets)
                {
                    // Если в активной книге нет листов, создаем новый
                    if (activeWorkbook.Sheets.Count == 0)
                    {
                        activeWorkbook.Sheets.Add();
                    }

                    // Получаем последний лист в активной книге
                    Excel.Worksheet targetSheet = (Excel.Worksheet)activeWorkbook.Sheets[activeWorkbook.Sheets.Count];

                    // Копируем данные из шаблона в целевой лист
                    //templateSheet.UsedRange.Copy(targetSheet.Range["A1"]);

                    // Копируем данные, ширину столбцов и высоту строк из шаблона в целевой лист
                    templateSheet.UsedRange.EntireRow.Copy(targetSheet.Range["A1"]);
                    templateSheet.UsedRange.EntireColumn.Copy(targetSheet.Range["A1"]);

                    // Устанавливаем имя целевого листа равным имени листа шаблона
                    targetSheet.Name = templateSheet.Name;
                    // Освобождаем COM-объекты для освобождения памяти и избежания блокировок
                    Marshal.ReleaseComObject(templateSheet);


                }

            }
            finally
            {
                // Закрываем шаблон без сохранения
                templateWorkbook.Close(false);
                // Освобождаем COM-объект шаблона
                Marshal.ReleaseComObject(templateWorkbook);
            }
        }

    }
}
