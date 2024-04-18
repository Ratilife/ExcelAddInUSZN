using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInUSZN
{
    public class TableHandler
    {

        // оставил как образец выборки данных, но при этом медлено исполнялся
        public Dictionary<string, TableInfo> HandleTables(Excel.Worksheet worksheet, string cellRange)
        {
            string[] parts      = cellRange.Split(':');
            string lastColumn   = parts[1].Substring(0, 2);     // "CR"
            string startRow     = parts[0].Substring(1);        // "1"
            string endRow       = parts[1].Substring(2);        // "39"

            // Создаем словарь для хранения ячеек, которые являются частью таблицы
            Dictionary<string, TableInfo> tableCells = new Dictionary<string, TableInfo>
            {
                { "Table1", new TableInfo { Cells = new List<Excel.Range>(), BorderStyles = new List<Excel.XlLineStyle>() } },
                { "Table2", new TableInfo { Cells = new List<Excel.Range>(), BorderStyles = new List<Excel.XlLineStyle>() } },
                { "Table3", new TableInfo { Cells = new List<Excel.Range>(), BorderStyles = new List<Excel.XlLineStyle>() } },
                { "Table4", new TableInfo { Cells = new List<Excel.Range>(), BorderStyles = new List<Excel.XlLineStyle>() } },
                { "Table5", new TableInfo { Cells = new List<Excel.Range>(), BorderStyles = new List<Excel.XlLineStyle>() } },
            };

            // Создаем список для хранения столбцов-разделителей
            List<string> dividerColumns = new List<string>();

            // Создаем список для хранения ячеек, которые являются частью таблицы
            List<Excel.Range> listCells = new List<Excel.Range>();

            // Проходим по всем ячейкам на листе
            foreach (Excel.Range cell in worksheet.Range[cellRange].Cells)
            {
              
                // Проверяем, есть ли у ячейки границы
                if ((Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle == Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle == Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle == Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    // Если у ячейки нет границ со всех сторон, то считаем ее столбцом-разделителем
                    string column = GetColumnName(cell.Column);
                    if (!dividerColumns.Contains(column))
                    {
                        dividerColumns.Add(column);
                    }
                }
            }

            // Теперь у вас есть список столбцов-разделителей, и вы можете использовать его для определения начала и конца каждой таблицы
            for (int i = 0; i < Math.Min(dividerColumns.Count, 5); i++)
            {
                string startColumn = dividerColumns[i];
                string endColumn = i < dividerColumns.Count - 1 ? dividerColumns[i + 1] : lastColumn; // Если это последний столбец-разделитель, то конец таблицы будет в столбце "CR"
                int kolRow = 0; 
                int KolColumns = 0;
                
                // Здесь вы можете добавить логику для добавления ячеек в соответствующие списки tableCells
                foreach (Excel.Range cell in worksheet.Range[startColumn + startRow + ":" + endColumn + endRow].Cells)
                {

                        /// Проверяем, есть ли у ячейки границы
                        if ((Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                            (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                            (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                            (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle != Excel.XlLineStyle.xlLineStyleNone)
                        {
                            tableCells["Table" + (i + 1)].Cells.Add(cell);
                            tableCells["Table" + (i + 1)].BorderStyles.Add(Excel.XlLineStyle.xlContinuous); // Сохраните стиль границы
                            KolColumns = KolColumns + 1;
                    }
                        else if (cell.MergeCells)
                        {
                            tableCells["Table" + (i + 1)].Cells.Add(cell);
                            tableCells["Table" + (i + 1)].BorderStyles.Add(Excel.XlLineStyle.xlContinuous); // Сохраните стиль границы
                            KolColumns = KolColumns + 1;    
                        }
                    
                    string columnName = GetColumnName(cell.Column);
                    if (cell.Row == 1 && columnName == endColumn)
                    {
                        tableCells["Table" + (i + 1)].Columns = KolColumns;
                    }
                    if (kolRow == int.Parse(endRow)) 
                    {
                        tableCells["Table" + (i + 1)].Rows = kolRow;
                    }
                    if(columnName!=startColumn && columnName != endColumn)
                    {
                        kolRow = kolRow + 1;
                    }
                }
            }
            return tableCells;
        }
        private string GetColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        // Метод для определения таблиц на листе Excel по наличию границ
        // Теперь этот метод принимает диапазон ячеек в качестве параметра
        public void HandleTables2(Excel.Worksheet worksheet, string cellRange)
        {
            // Создаем список для хранения ячеек, которые являются частью таблицы
            List<Excel.Range> tableCells = new List<Excel.Range>();
            List<Excel.Range> tableEmptyCells = new List<Excel.Range>();
            // Проходим по всем ячейкам на листе
            foreach (Excel.Range cell in worksheet.Range[cellRange].Cells)
            {
                /// Проверяем, есть ли у ячейки границы
                if ((Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle != Excel.XlLineStyle.xlLineStyleNone &&
                    (Excel.XlLineStyle)cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle != Excel.XlLineStyle.xlLineStyleNone)
                {
                    // Если у ячейки есть границы со всех сторон, то считаем ее частью таблицы
                    tableCells.Add(cell);

                }

                else if (cell.MergeCells)
                {
                    // Если ячейка объединена по горизонтали или вертикали, то также считаем ее частью таблицы
                    tableCells.Add(cell);
                }
                else
                {
                    tableEmptyCells.Add(cell);
                }
            }

            // Записываем адреса ячеек, которые являются частью таблицы, в ячейки листа, начиная с ячейки "B107"
            int row = 107;
            int nom = 1;
            foreach (Excel.Range cell in tableCells)
            {
                worksheet.Cells[row, "A"].Value = nom;
                worksheet.Cells[row, "B"].Value = "Ячейка, являющаяся частью таблицы: " + cell.Address;
                // Окрашиваем ячейку в светло-желтый цвет
                cell.Interior.Color = System.Drawing.Color.Yellow;
                row++;
                nom++;
            }
            foreach (Excel.Range cell in tableEmptyCells) 
            {
                // Окрашиваем ячейку в голубой цвет
                cell.Interior.Color = System.Drawing.Color.BlueViolet;
            }

                // Создаем список для хранения номеров таблиц
                List<int> tableNumbers = new List<int>();
            int tableNumber = 1;
            
            // Проходим по списку ячеек и проверяем, являются ли они частью одной и той же таблицы
            for (int i = 0; i < tableCells.Count - 1; i++)
            {

                // Если следующая ячейка не соседствует с текущей, это означает, что начинается новая таблица
                /*
                    Если абсолютное значение разницы между номерами строк или столбцов больше 1, 
                    это означает, что между двумя ячейками есть хотя бы одна ячейка. В контексте  кода 
                    это указывает на то, что начинается новая таблица.
                */
                if (Math.Abs(tableCells[i].Row - tableCells[i + 1].Row) > 1 ||
                    Math.Abs(tableCells[i].Column - tableCells[i + 1].Column) > 1)
                {
                    tableNumber++;
                }
                    tableNumbers.Add(tableNumber);
            }

            // Добавляем номер последней таблицы
            tableNumbers.Add(tableNumber);

            // Выводим номера таблиц в ячейки листа, начиная с ячейки "C107"
            row = 107;
            foreach (int number in tableNumbers)
            {
                worksheet.Cells[row, "C"].Value = "Номер таблицы: " + number;
                row++;
            }
        }
    }

    public class TableInfo
    {
        public int Rows { get; set; }
        public int Columns { get; set; }
        public List<Excel.Range> Cells { get; set; }
        public List<Excel.XlLineStyle> BorderStyles { get; set; } // Стили границ для каждой ячейки
    }
}


/* это работает медленно
// начало проверить и перенести куда надо
// Создание нового экземпляра класса TableHandler
TableHandler tableHandler = new TableHandler();

// Вызов метода HandleTables
tableCells = tableHandler.HandleTables(activeWorksheet, "A1:CR39");

foreach (KeyValuePair<string, TableInfo> entry in tableCells) 
{
   string tableName = entry.Key;
   TableInfo tableInfo = entry.Value;
   // Перемещение курсора в конец документа
   wordApp.Selection.EndKey(Word.WdUnits.wdStory);
   // Создание новой таблицы в документе Word
   Word.Table wordTable = wordDoc.Tables.Add(wordApp.Selection.Range, tableInfo.Rows, tableInfo.Columns);
   // Установка границ для таблицы
   for (int i = 0; i < tableInfo.Cells.Count; i++)
   {
       Excel.Range cell = tableInfo.Cells[i];
       Excel.XlLineStyle borderStyle = tableInfo.BorderStyles[i];

       // Заполнение таблицы данными из Excel
       //if (cell.Value2 != null)
       //{
       int row = i / tableInfo.Columns + 1;
       int column = i % tableInfo.Columns + 1;
       if (cell.Value2 != null)
       {
          wordTable.Cell(row, column).Range.Text = cell.Value2.ToString();
       }
       else
       {
          wordTable.Cell(row, column).Range.Text = "";
       }
       // Установка стиля границы
       foreach (Word.Border border in wordTable.Cell(row, column).Borders)
       {
           border.LineStyle = (Word.WdLineStyle)borderStyle;
       }
           // Удаление диагональных границ
       wordTable.Borders[Word.WdBorderType.wdBorderDiagonalDown].LineStyle = Word.WdLineStyle.wdLineStyleNone;
       wordTable.Borders[Word.WdBorderType.wdBorderDiagonalUp].LineStyle = Word.WdLineStyle.wdLineStyleNone;
       //}
   }

   // Добавление пустой строки после каждой таблицы
   wordApp.Selection.EndKey(Word.WdUnits.wdStory);
   wordApp.Selection.TypeParagraph();
}
*/