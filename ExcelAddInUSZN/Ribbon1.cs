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
        /*private void action()
        {
            // Получить активный лист
            Excel.Worksheet activeSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            // Создать список для хранения данных из таблиц
            List<List<string>> tablesData = new List<List<string>>();

            // Создать список для хранения количества столбцов для каждого диапазона
            List<int> columnsCountList = new List<int>();

            // Список диапазонов
            List<string> ranges = new List<string>() { "$B$2:$S$38", "$U$2:$AL$38", "$AN$2:$BE$38", "$BG$2:$BX$38", "$BZ$2:$CQ$38" };

            // Пройти по каждому диапазону (таблице)
            foreach (string range in ranges)
            {
                Excel.Range tableRange = activeSheet.Range[range];
                List<string> tableData = new List<string>();        // Создание списка для хранения данных таблицы
                columnsCountList.Add(tableRange.Columns.Count);     // Добавить количество столбцов в список
                foreach (Excel.Range row in tableRange.Rows)
                {
                    bool hasData = false;
                    string rowData = "";

                    foreach (Excel.Range cell in row.Cells)
                    {
                        // Проверить, является ли ячейка ссылкой
                        if (cell.HasFormula)
                        {
                            // Получить ссылку
                            string reference = cell.Formula.ToString().Substring(1); // Удалить знак "=" из формулы
                            Excel.Range referencedCell = activeSheet.Range[reference];

                            // Если ссылочная ячейка не пуста, добавить данные в строку
                            if (!string.IsNullOrEmpty(referencedCell.Value?.ToString()))
                            {
                                hasData = true;
                                rowData += referencedCell.Value + "\t";
                            }

                        }
                        else if (!string.IsNullOrEmpty(cell.Value?.ToString()))
                        {
                            hasData = true;
                            rowData += cell.Value + "\t";
                        }
                    }
                        if (hasData)
                        {
                            tableData.Add(rowData);
                        }
                }
                    tablesData.Add(tableData);
            }
            

            // Создать новый документ Word
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            // Перенести данные из таблиц в документ Word
            for (int i = 0; i < tablesData.Count; i++)
            {
                var tableData = tablesData[i];
                // Создать таблицу в документе Word
                Word.Table wordTable = wordDoc.Tables.Add(wordDoc.Range(), tableData.Count, columnsCountList[i]);

                for (int  j= 0; j < tableData.Count; j++)
                {
                    string[] rowData = tableData[j].Split('\t');

                    for (int k = 0; k < rowData.Length; k++)
                    {
                        wordTable.Cell(j + 1, k + 1).Range.Text = rowData[k];
                    }
                }

                wordDoc.Range().InsertParagraphAfter(); // Вставить параграф после каждой таблицы
                // Применить стиль к таблице
                wordTable.set_Style("Table Grid"); // Замените "Table Grid" на желаемый стиль таблицы: “Table Grid”, “Table Professional”, “Table List 1”, “Table List 2”, “Table Classic 1”,

                // Вставить параграф после каждой таблицы
                wordDoc.Range().InsertParagraphAfter();
            }

            // Отобразить документ Word
            wordApp.Visible = true;

        }*/
        private void sample_to_use()
        {
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
            
            ContentDocument cd = new ContentDocument();
            // Создадим список кортежей для объединения ячеек
            List<Tuple<int, int>> mergeCellsTab4 = new List<Tuple<int, int>>
            {
                new Tuple<int, int>(3, 2),     // Объединить ячейку в третьей строке, первом столбце
                new Tuple<int, int>(3, 3)      // Объединить ячейку в третьей строке, втором столбце
            };
            // Указываем диапазоны для считывания данных
            List<string> excelRanges = new List<string>
            {
                "B5:B37",
                "U5:U37",
                "AN5:AN37",
                "BG5:BG37",
                "BZ5:BZ37"
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
                () => CreateTableAndInsert(wordDoc,  dt1,false),
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
                () => InsertText(wordDoc,cd.paragraph3,false,"Times New Roman",Word.WdParagraphAlignment.wdAlignParagraphJustify),
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
                () => Create_a_list_table_Delete(activeWorksheet, wordDoc,excelRanges),
                // пункт 7
                () => InsertText(wordDoc,cd.paragraph7),
                // пункт 8
                () => InsertText(wordDoc,cd.paragraph8),
                // подвал документа 
                () => InsertText(wordDoc, cd.text_pered_podpis1),
                () => CreateTableAndInsert(wordDoc,dt7,false),
                () => InsertText(wordDoc, cd.text_pered_podpis2),
                () => CreateTableAndInsert(wordDoc,dt8,false),
                () => CreateTableAndInsert(wordDoc,dt9,false)
            };

            // Вставка структуры документа в Word
            foreach (System.Action action in documentStructure)
            {
                action();
            }

            GetTableWidth(wordDoc, 4);

            System.Media.SystemSounds.Asterisk.Play();
        }

        public void CopyExcelTableToWord(Excel.Worksheet activeWorksheet, Document wordDoc, string range,string styleName = "Times New Roman", float fontSize=10)
        {
            // Скопируйте указанный диапазон с активного листа
            Excel.Range excelRange = activeWorksheet.Range[range];
            string[] parts = range.Split(':');
            Regex regex = new Regex("[A-Za-z]+"); // Регулярное выражение для извлечения буквенной части
            string firstColumn =  regex.Match(parts[0]).Value; // Получаем "ABA" или "B"
            //string endColumn = regex.Match(parts[1]).Value; // Получаем "ABD" или "S"
           

            int firstColumnIndex = ColumnLetterToNumber(firstColumn);
            int lastColumn;
            // Пройдите по каждой строке в указанном диапазоне
            for (int i = 1; i <= excelRange.Rows.Count; i++)
            {
                // Получите ячейку в первом столбце
                Excel.Range firstCell = excelRange.Cells[i, 1];

                // Проверьте, заполнена ли ячейка
                if (!string.IsNullOrWhiteSpace(firstCell.Value))
                {
                    if(firstColumnIndex == 1)
                    {
                        lastColumn = 0;
                    }
                    else 
                    {
                        lastColumn = firstColumnIndex-2;
                    }
                    // Если ячейка заполнена, скопируйте эту строку
                    //Excel.Range rowRange = activeWorksheet.Range[activeWorksheet.Cells[i, firstColumnIndex], activeWorksheet.Cells[i, excelRange.Columns.Count+1]];
                    Excel.Range rowRange = activeWorksheet.Range[activeWorksheet.Cells[i, firstColumnIndex], activeWorksheet.Cells[i, excelRange.Columns.Count + 1+ lastColumn]];
                    rowRange.Copy(); 
                    // Вставьте скопированную строку в документ Word
                    Word.Range wordRange = wordDoc.Range(wordDoc.Content.End - 1, wordDoc.Content.End - 1);
                    wordRange.Paste();

                    // Примените стиль и размер шрифта к вставленной таблице
                    Table table = wordRange.Tables[wordRange.Tables.Count];
                    table.Range.Font.Name = styleName;
                    table.Range.Font.Size = fontSize;
                    // Установка цвета границ таблицы
                    table.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                    table.Borders.InsideColor = Word.WdColor.wdColorBlack;
                    table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                }
            }            
        }

        public void CopyExcelTableToWord2(Excel.Worksheet activeWorksheet, Document wordDoc, string range)
        {
            // Скопируйте указанный диапазон с активного листа
            Excel.Range excelRange = activeWorksheet.Range[range];
            excelRange.Copy();

            // Вставьте скопированный диапазон в документ Word
            Word.Range wordRange = wordDoc.Range(wordDoc.Content.End - 1, wordDoc.Content.End - 1);
            wordRange.Paste();

            // Получите первую таблицу в документе
            Word.Table wordTable = wordDoc.Tables[wordDoc.Tables.Count];

            // Измените ширину таблицы, чтобы она соответствовала ширине страницы
            wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

            // Пройдите по всем ячейкам в таблице и измените шрифт и размер шрифта
            foreach (Word.Cell cell in wordTable.Range.Cells)
            {
                cell.Range.Font.Name = "Times New Roman";
                cell.Range.Font.Size = 10;
            }
        }
       
        private void Create_a_list_table_Delete(Worksheet activeWorksheet, Document wordDoc, List<string> ranges)
        {
            // Эта функция создает таблицу в документе Word из уникальных данных в листе Excel.
            Word.Range rng = wordDoc.Content;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // Устанавливает точку вставки в конец документа Word.

            try
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
                        if (!string.IsNullOrWhiteSpace(cell.Value))
                        {
                            uniqueData.Add(RemoveNumbersAtStart(cell.Value.ToString()));
                        }
                    }
                }


                // Создает таблицу в документе Word с числом строк, равным числу уникальных данных плюс 2 (для заголовка и подвала), и 2 столбца.
                Table wordTable = wordDoc.Tables.Add(rng, uniqueData.Count + 2, 2);
                wordTable.Borders.Enable = 1; // Включаем границы таблицы

                // Объединяем ячейки в шапке
                wordTable.Cell(1, 1).Merge(wordTable.Cell(1, 2));
                // Заполняем шапку таблицы
                wordTable.Cell(1, 1).Range.Text = "Наименование социальной услуги по уходу";

                // Заполняем строки данными из списка HashSet.
                int rowIndex = 2;
                foreach (string data in uniqueData)
                {
                    // Объединяем ячейки таблицы
                    wordTable.Cell(rowIndex, 1).Merge(wordTable.Cell(rowIndex, 2));
                    wordTable.Cell(rowIndex, 1).Range.Text = data;
                    rowIndex++;
                }
                // Заполняем подвал таблицы
                wordTable.Cell(uniqueData.Count + 2, 1).Range.Text = "Общее количество социальных услуг по уходу, не включенных в социальный пакет долговременного ухода";
                wordTable.Cell(uniqueData.Count + 2, 2).Range.Text = uniqueData.Count.ToString();

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
                                    int fontSize = 11, int bold = 0)
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
                                    Word.WdCellVerticalAlignment verticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
        {
            Word.Range rng = wordDoc.Content;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Table table = wordDoc.Tables.Add(rng, dt.Rows.Count, dt.Columns.Count); // Используйте rng вместо wordDoc.Content

            int row = 1;
            foreach (DataRow dr in dt.Rows)
            {
                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    //table.Cell(row, col).Range.Text = dr[col - 1].ToString();
                    Word.Range cellRange = table.Cell(row, col).Range;
                    cellRange.Text = dr[col - 1].ToString();
                    cellRange.Font.Name = fontName;
                    cellRange.Font.Size = fontSize;
                    cellRange.ParagraphFormat.Alignment = alignment;
                    table.Cell(row, col).VerticalAlignment = verticalAlignment;
                }
                row++;
            }
            // Установка цвета границ таблицы
            if (table_boundaries)
            {
                table.Borders.OutsideColor = Word.WdColor.wdColorBlack;
                table.Borders.InsideColor = Word.WdColor.wdColorBlack;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }

            try
            {
                // Объединение указанных ячеек
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
                //MessageBox.Show($"Ошибка при объединении ячеек: {e.Message}");
            }

            // Перенос каретки на новую строку
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

        public static string RemoveNumbersAtStart(string input)
        {
            // Удаляем числа и точку в начале строки
            return Regex.Replace(input.Trim(), @"^\d+\.?\s*", "");
        }

    }
}
