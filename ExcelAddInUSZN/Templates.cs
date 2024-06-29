using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInUSZN
{
    public partial class Templates : Form
    {
        // Определение делегата и события
        public delegate void TemplateSelectedHandler(string templatePath);
        public event TemplateSelectedHandler TemplateSelected;

        //SelectedTemplatePath - это строковая переменная, которая хранит путь к выбранному шаблону.
        //Когда пользователь выбирает шаблон из списка, значение этой переменной обновляется, чтобы указать путь к выбранному файлу.
        //get позволяет получить значение переменной SelectedTemplatePath
        //private set означает, что значение можно установить только внутри класса (например, в методах этого класса), но не извне
        public string SelectedTemplatePath { get; private set; }

        public Templates()
        {
            InitializeComponent();
            // Назначаем обработчик события для изменения цвета выделенной строки
            listBoxTemplates.SelectedIndexChanged += ListBoxTemplates_SelectedIndexChanged;
        }

        private void Templates_Load(object sender, EventArgs e)
        {
            string templatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Настраиваемые шаблоны Office");

            if (Directory.Exists(templatesPath))
            {
                // Если папка с шаблонами существует, получаем список файлов с расширением .xltx
                var templateFiles = Directory.GetFiles(templatesPath, "*.xltx");
                foreach (var file in templateFiles)
                {

                    // Добавляем в список элементы, представляющие имена файлов без пути
                    listBoxTemplates.Items.Add(Path.GetFileName(file));
                }
            }
            else
            {
                // Если папка не существует, выводим сообщение об ошибке
                MessageBox.Show("Папка 'Настраиваемые шаблоны Office' не найдена.");
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (listBoxTemplates.SelectedItem != null)
            {
                // Получаем путь к папке с настраиваемыми шаблонами Office
                string templatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Настраиваемые шаблоны Office");
                // Составляем полный путь к выбранному шаблону
                SelectedTemplatePath = Path.Combine(templatesPath, listBoxTemplates.SelectedItem.ToString());
                // Устанавливаем результат диалога как OK и закрываем форму
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                // Если не выбран шаблон, выводим сообщение
                MessageBox.Show("Пожалуйста, выберите шаблон.");
            }
        }

        private void ListBoxTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Предполагая, что 'sender' это listBoxTemplates
            var listBox = (ListBox)sender;
            if (listBox.SelectedItem != null)
            {
                var selectedTemplate = listBox.SelectedItem.ToString();
                // Теперь 'selectedTemplate' содержит имя выбранного файла шаблона
                // Здесь можно добавить дополнительный код для работы с выбранным шаблоном
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
