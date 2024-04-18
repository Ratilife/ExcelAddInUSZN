using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelAddInUSZN
{
    internal class MainForm : Form   // Возможно удалю код как неудовлетворительный
    {
        private NotifyIcon notifyIcon;

        public MainForm()
        {
            InitializeComponent();
            InitializeNotifyIcon();
        }

        private void InitializeComponent()
        {
            throw new NotImplementedException();
        }

        private void InitializeNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = Properties.Resources.YourIcon; // Замените на свою иконку
            notifyIcon.Text = "Заполнение документа Word";
            notifyIcon.Visible = true;
        }

        private void UpdateProcessStatus(string status)
        {
            // Обновление состояния в иконке системного трея
            notifyIcon.Text = $"Заполнение: {status}";
        }

        // Метод заполнения документа Word данными
        private void FillWordDocument()
        {
            // Ваш код заполнения документа
            // Внутри цикла обновляйте состояние и вызывайте UpdateProcessStatus
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Запустите процесс заполнения документа при загрузке формы
            FillWordDocument();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Скрытие иконки в системном трее при закрытии формы
            notifyIcon.Visible = false;
        }
    }
}
