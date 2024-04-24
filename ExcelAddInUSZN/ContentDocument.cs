using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelAddInUSZN
{
    internal class ContentDocument
    {
        // переменные для раздела 3
        public string minutes { get; set; }
        public string hours_minutes { get; set; }

        // переменные для таблицы 4
        public string time_in_minutes { get; set; }
        public string time_in_hours { get; set; }
        public string count_servise { get; set; }

        // номера сносок
        // private char superscriptOne   = '\u00B9';
        // private char superscriptTwo   = '\u00B2';
        // private char superscriptThree = '\u00B3';
        // private char superscriptFour  = '\u2074';
        // private char superscriptFive  = '\u2075';
        // private char superscriptSix   = '\u2076';

        public string heading { get; private set; } = "Дополнение \r\nк индивидуальной программе предоставления социальных услуг (ИППСУ)\r\n";
        //public int font_size_heading { get; private set; } = 14;
        public string surname { get; private set; } = "Фамилия:\n";
        public string name { get; private set; } = "Имя:\n";
        public string patronymic { get; private set; } = "Отчество:\n";
        public string d_r_gender_SNILS { get; private set; } = "Дата рождения ___________ Пол _______ СНИЛС _______________\n";
        public string heading_aragraph { get; private set; } = "Социальный пакет долговременного ухода, предоставляемый гражданину бесплатно в форме социального обслуживания на дому, \r\nусловия его предоставления\r\n";
        public string paragraph1 { get; private set; } = "1. Установлен уровень нуждаемости в уходе ______________________________________ \n";
        public string paragraph2 { get; private set; } = "2. Объем социального пакета долговременного ухода в неделю в соответствии с установленным уровнем нуждаемости в уходе (в часах)\n";
        public string paragraph3 { get; private set; } = "3. Объем назначенного социального пакета долговременного ухода в неделю (в минутах /часах): ";
        public string paragraph4 { get; private set; } = "4. Условия предоставления социального пакета долговременного ухода:\r\n";
        public string paragraph4_1 { get; private set; } = "4.1. Количество дней в неделю, в течение которых гражданину предоставляются социальные услуги по уходу___________\n";
        public string paragraph4_2 { get; private set; } = "4.2. Ежедневное распределение количества посещений гражданина помощником по уходу по дням недели:\n";
        public string paragraph4_3 { get; private set; } = "4.3. Ежемесячное распределение объема социального пакета долговременного ухода по неделям и дням недели:\r\n";
        public string paragraph4_4 { get; private set; } = "4.4. Еженедельное распределение перечня и объема социальных услуг по уходу" + '\u00B9' + ", включенных в социальный пакет долговременного ухода и предоставляемых в соответствии с рекомендуемыми стандартами" + '\u00B2' + ", на получение которых выражено согласие:\r\n";
        public string paragraph4_5 { get; private set; } = "4.5. Ежемесячный объем социального пакета долговременного ухода (в минутах /часах):\r\n";
        public string paragraph5 { get; private set; } = "5. Перечень социальных услуг по уходу, не включенных в социальный пакет долговременного ухода, поскольку их предоставление гарантируется гражданами, осуществляющими уход (из числа ближайшего окружения):\r\n";
        public string paragraph6 { get; private set; } = "6. Перечень социальных услуг по уходу, не включенных в социальный пакет долговременного ухода, предоставление которых гражданину не требуется:\r\n";
        public string paragraph7 { get; private set; } = "7. Сроки предоставления социальных услуг по уходу, включенных в пакет долговременного ухода: _______________________________________________________________________________\r\n";
        public string paragraph8 { get; private set; } = "Поставщик социальных услуг: _______________________________________________________________________________\r\n";
        public string text_pered_podpis1 { get; private set; } = "С содержанием социального пакета долговременного ухода, предоставляемого в форме социального обслуживания на дому, согласен (согласна):\r\n";
        public string text_pered_podpis2 { get; private set; } = "Правильность составления дополнения к индивидуальной программе предоставления социальных услуг подтверждаю:" + '\u2076' + "\r\n";
        public string paragraph3_date { get; private set; }
        public string explanation { get; private set; } = "*На 2 и 4 неделях месяца включаются социальные услуги по уходу, периодичность которых составляет 2 раза в месяц(гигиеническая обработка рук и ногтей, помощь в гигиенической обработке рук и ногтей).\r\n" +
                                                          "** На 3 неделе месяца включаются социальные услуги по уходу, периодичность которых составляет 1 раз в месяц \r\n(гигиеническая обработка ног и ногтей, помощь в гигиенической обработке ног и ногтей, гигиеническая стрижка).\r\n"; 


        public string timeInString()
        {
            return $"{paragraph3} {minutes}/{hours_minutes} \n";
        }
        public Word.Table CreateTableWord(Word.Document tempDoc, System.Data.DataTable dt)
        {
            Word.Table table = tempDoc.Tables.Add(tempDoc.Content, dt.Rows.Count, dt.Columns.Count);
            int row = 1;
            foreach (DataRow dr in dt.Rows)
            {
                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    table.Cell(row, col).Range.Text = dr[col - 1].ToString();
                }
                row++;
            }
            return table;
        }

        // создание таблицы1
        public System.Data.DataTable CreateTable1()
        {
            /*Word.Table table = tempDoc.Tables.Add(tempDoc.Content, 2, 5);
            table.Cell(1, 2).Range.Text = "№";
            table.Cell(1, 4).Range.Text = "Статус";
            table.Cell(2, 1).Range.Text = "(дата составления ИППСУ)";
            table.Cell(2, 3).Range.Text = "(ИППСУ)";
            table.Cell(2, 5).Range.Text = "(первичная, повторная, очередная ИППСУ)";

            for (int i = 1; i <= 5; i++)
            {
                table.Cell(1, i).Range.Font.Size = 11;
            }
            for (int i = 1; i <= 5; i++)
            {
                table.Cell(2, i).Range.Font.Size = 8;
            }

            // Перенос каретки на новую строку
            Word.Paragraph para = tempDoc.Content.Paragraphs.Add();
            para.Range.Text = "\n";
            para.Range.InsertParagraphAfter();*/
            // Создание DataTable с 5 столбцами
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < 5; i++)
            {
                dt.Columns.Add();
            }

            // Добавление двух строк в DataTable
            DataRow row1 = dt.NewRow();
            DataRow row2 = dt.NewRow();
            dt.Rows.Add(row1);
            dt.Rows.Add(row2);

            // Заполнение ячеек таблицы
            dt.Rows[0][1] = "№";
            dt.Rows[0][3] = "Статус";
            dt.Rows[1][0] = "(дата составления ИППСУ)";
            dt.Rows[1][2] = "(ИППСУ)";
            dt.Rows[1][4] = "(первичная, повторная, очередная ИППСУ)";

            return dt;
        }
        // создание таблицы2
        public System.Data.DataTable CreateTable2()
        {
            /*Word.Table table = tempDoc.Tables.Add(tempDoc.Content, 4, 8);
            table.Cell(1, 1).Range.Text = "Дни недели";
            table.Cell(2, 1).Range.Text = "1 раз в день";
            table.Cell(3, 1).Range.Text = "2 раза в день";
            table.Cell(4, 1).Range.Text = "3 раза в день";
            table.Cell(1, 2).Range.Text = "Пн"; 
            table.Cell(1, 2).Range.Text = "Вт";
            table.Cell(1, 2).Range.Text = "Ср";
            table.Cell(1, 2).Range.Text = "Чт";
            table.Cell(1, 2).Range.Text = "Пт";
            table.Cell(1, 2).Range.Text = "Сб";
            table.Cell(1, 2).Range.Text = "Вс";

            for (int i = 1; i <= 8; i++)
            {
                table.Cell(1, i).Range.Font.Size = 10;
                table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            return table;*/
            // Создание DataTable с 8 столбцами
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < 8; i++)
            {
                dt.Columns.Add();
            }
            // Добавление четырех строк в DataTable
            for (int i = 0; i < 4; i++)
            {
                dt.Rows.Add();
            }
            // Заполнение ячеек таблицы
            dt.Rows[0][0] = "Дни недели";
            dt.Rows[1][0] = "1 раз в день";
            dt.Rows[2][0] = "2 раза в день";
            dt.Rows[3][0] = "3 раза в день";
            dt.Rows[0][1] = "Пн";
            dt.Rows[0][2] = "Вт";
            dt.Rows[0][3] = "Ср";
            dt.Rows[0][4] = "Чт";
            dt.Rows[0][5] = "Пт";
            dt.Rows[0][6] = "Сб";
            dt.Rows[0][7] = "Вс";


            return dt;
        }
        // создание таблицы3
        public System.Data.DataTable CreateTable3()
        {

            // Создание DataTable с 6 столбцами
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < 6; i++)
            {
                dt.Columns.Add();
            }
            // Добавление четырех строк в DataTable
            for (int i = 0; i < 2; i++)
            {
                dt.Rows.Add();
            }
            // Заполнение ячеек таблицы
            dt.Rows[0][0] = "Количество расчетных недель в месяц – 5";
            dt.Rows[1][0] = "Количество расчетных дней – 30";
            dt.Rows[0][1] = "1 неделя";
            dt.Rows[1][1] = "5 дней";
            dt.Rows[0][2] = "2 неделя";
            dt.Rows[1][2] = "7 дней";
            dt.Rows[0][3] = "3 неделя";
            dt.Rows[1][3] = "7 дней";
            dt.Rows[0][4] = "4 неделя";
            dt.Rows[1][4] = "7 дней";
            dt.Rows[0][5] = "5 неделя";
            dt.Rows[1][5] = "4 дня";

            return dt;
        }
        // создание таблицы9
        public System.Data.DataTable CreateTable4()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            
            // Создание DataTable с 3 столбцами
            for (int i = 0; i < 3; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < 3; i++)
            {
                dt.Rows.Add();
            }
            // Заполнение ячеек таблицы
            dt.Rows[0][0] = "Ежемесячный объем";
            dt.Rows[1][0] = "Общая продолжительность времени на предоставление социальных услуг по уходу, включенных в социальный пакет долговременного ухода, в месяц";
            dt.Rows[2][0] = "Общее количество социальных услуг по уходу, включенных в социальный пакет долговременного ухода";
            dt.Rows[0][1] = "в мин";
            dt.Rows[0][2] = "в часах";
            dt.Rows[1][1] = time_in_minutes;
            dt.Rows[1][2] = time_in_hours;
            dt.Rows[2][1] = count_servise;



            return dt;
        }

        public System.Data.DataTable CreateTable5()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // Создание DataTable с 2 столбцами
            for (int i = 0; i < 2; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < 3; i++)
            {
                dt.Rows.Add();
            }
            // Заполнение ячеек таблицы
            dt.Rows[0][0] = "Наименование социальной услуги по уходу";
            dt.Rows[0][1] = "Фамилия, имя, отчество лица, гарантирующего предоставление социальной услуги по уходу, статус";
            dt.Rows[2][0] = "Общее количество социальных услуг по уходу, не включенных в социальный пакет долговременного ухода"+ '\u2074';

            return dt;
        }
        // ? реализация формирования таблицы 6 вынесено за приделы класса
        /*public System.Data.DataTable CreateTable6(int columns, int row)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < columns; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < row; i++)
            {
                dt.Rows.Add();
            }
            return dt;
        }*/
       
        public System.Data.DataTable CreateTable7()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // Создание DataTable с 3 столбцами
            for (int i = 0; i < 3; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < 2; i++)
            {
                dt.Rows.Add();
            }
            dt.Rows[1][0] = "(подпись гражданина или его законного представителя)";
            dt.Rows[1][2] = "(ФИО)";
            return dt;
        }
        public System.Data.DataTable CreateTable8()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // Создание DataTable с 5 столбцами
            for (int i = 0; i < 5; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < 2; i++)
            {
                dt.Rows.Add();
            }
            dt.Rows[1][0] = "(должность)";
            dt.Rows[1][2] = "(ФИО)";
            dt.Rows[1][4] = "(подпись)";
            return dt;
        }
        public System.Data.DataTable CreateTable9()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // Создание DataTable с 3 столбцами
            for (int i = 0; i < 3; i++)
            {
                dt.Columns.Add();
            }
            // Добавление 3-х строк в DataTable
            for (int i = 0; i < 2; i++)
            {
                dt.Rows.Add();
            }
            dt.Rows[1][0] = "М. П.";
            dt.Rows[1][2] = "(дата составления дополнения к ИППСУ)";
            

            return dt;
        }
        public string[] create_footnotes() 
        {

            string[] footnotes = new string[]
                {
                " Перечень социальных услуг по уходу заполняется в соответствии с перечнем социальных услуг по уходу, включаемых в социальный пакет долговременного ухода, предусмотренным приложением № 6 к Типовой модели системы долговременного ухода за гражданами пожилого возраста и инвалидами, нуждающимися в уходе (далее – модель).",
                " Рекомендуемые стандарты социальных услуг по уходу, включаемых в социальный пакет долговременного ухода, предусмотренные приложением № 7 к модели.",
                " В  графе  указывается  суммарный  объем  времени,  затрачиваемого  на  предоставление  социальной  услуги  по  уходу  с учетом ее кратности.", 
                " Вносятся услуги, в предоставлении которых помощник по уходу участия не принимает. Наименование услуг должно соответствовать исчерпывающему перечню социальных услуг по уходу, включаемых в социальный пакет долговременного ухода, предусмотренному приложением № 6 к модели.",
                " Общее количество социальных услуг по уходу, вносимых в разделы 4-6 настоящего дополнения к индивидуальной программе,  должно  соответствовать  исчерпывающему  перечню  социальных  услуг  по  уходу,  включаемых  в социальный пакет долговременного ухода, предусмотренному приложением № 6 к модели.",
                " Настоящее  дополнение  к  индивидуальной  программе  подписывается  уполномоченным  представителем  органа государственной  власти  субъекта  Российской  Федерации  в  сфере  социального  обслуживания  граждан  субъекта Российской Федерации или уполномоченной данным органом организации, не являющейся поставщиком социальных услуг"
                };
            return footnotes;
    }
    }
}
