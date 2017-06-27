using System;
using System.IO;
using System.Windows.Forms;

namespace TaskScheduler
{
    public partial class TaskManager : Form
    {
        private TaskScheduler _taskScheduler;
        public TaskManager()
        {
            InitializeComponent();
            _taskScheduler = new TaskScheduler();

            // инициализируем объект синхронизации
            _taskScheduler.SynchronizingObject = this; 

            dateTimePickerStartDate.Value = DateTime.Today;
            dateTimePickerEndDate.Value = DateTime.Today.AddYears(1);
            dateTimePickerTriggerTime.Value = DateTime.Now.AddMinutes(10); // Add 10 Minutes for testing
        }

        private void UpdateTaskList()
        {
            listViewItems.Items.Clear();
            foreach (TaskScheduler.TriggerItem item in _taskScheduler.TriggerItems)
            {
                ListViewItem listItem = listViewItems.Items.Add(item.Tag.ToString());
                listItem.Tag = item;
                DateTime nextDate = item.GetNextTriggerDateTime();
                if (nextDate != DateTime.MaxValue)
                    listItem.SubItems.Add(nextDate.ToString());
                else
                    listItem.SubItems.Add("Никогда");
            }
        }

        private void ResetScheduler()
        {
            _taskScheduler.Enabled = false;
            _taskScheduler.TriggerItems.Clear();
            UpdateTaskList();
            //textBoxEvents.Clear();
        }

        private void CreateSchedulerItem()
        {
            TaskScheduler.TriggerItem triggerItem = new TaskScheduler.TriggerItem();
            triggerItem.Tag = textBoxlabelOneTimeOnlyTag.Text;
            triggerItem.StartDate = dateTimePickerStartDate.Value;
            triggerItem.EndDate = dateTimePickerEndDate.Value;
            triggerItem.TriggerTime = dateTimePickerTriggerTime.Value;
            // добавляем обработчик
            triggerItem.OnTrigger += new TaskScheduler.TriggerItem.OnTriggerEventHandler(triggerItem_OnTrigger); 

            // выставляем единовременный триггер
            triggerItem.TriggerSettings.OneTimeOnly.Active = checkBoxOneTimeOnlyActive.Checked;
            triggerItem.TriggerSettings.OneTimeOnly.Date = dateTimePickerOneTimeOnlyDay.Value.Date;

            // выставляем интервал для ежедневного триггера
            triggerItem.TriggerSettings.Daily.Interval = (ushort)numericUpDownDaily.Value;

            // устанавливаем активные дни для еженедельного триггера
            for (byte day = 0; day < 7; day++) // Set the active Days
                triggerItem.TriggerSettings.Weekly.DaysOfWeek[day] = checkedListBoxWeeklyDays.GetItemChecked(day);

            // устанавливаем активные месяцы для ежемесячного триггера
            for (byte month = 0; month < 12; month++)
                triggerItem.TriggerSettings.Monthly.Month[month] = checkedListBoxMonthlyMonths.GetItemChecked(month);

            // устанавливаем активные дни (0..30 = дни, 31=последний день) для месячного триггера
            for (byte day = 0; day < 32; day++)
                triggerItem.TriggerSettings.Monthly.DaysOfMonth[day] = checkedListBoxMonthlyDays.GetItemChecked(day);

            // устанавливаем активный день недели и день в неделю для месячного триггера
            // т.е. перый понедельник или последняя пятница
            for (byte weekNumber = 0; weekNumber < 5; weekNumber++)
                triggerItem.TriggerSettings.Monthly.WeekDay.WeekNumber[weekNumber] = checkedListBoxMonthlyWeekNumber.GetItemChecked(weekNumber);
            for (byte day = 0; day < 7; day++)
                triggerItem.TriggerSettings.Monthly.WeekDay.DayOfWeek[day] = checkedListBoxMonthlyWeekDay.GetItemChecked(day);

            // делаем активным
            triggerItem.Enabled = true;
            // Добавляем триггер в список
            _taskScheduler.AddTrigger(triggerItem);
            // стартуем планировщик
            _taskScheduler.Enabled = checkBoxEnabled.Checked;

            UpdateTaskList();
        }

        private void ShowAllTriggerDates()
        {
            if (listViewItems.SelectedItems.Count > 0)
            {
                TaskScheduler.TriggerItem item = (TaskScheduler.TriggerItem)listViewItems.SelectedItems[0].Tag;
                Form form = new Form();
                ListView listView = new ListView();
                listView.FullRowSelect = true;

                form.Text = "Полный список для задач: "+item.Tag.ToString();
                form.Width = 400;
                form.Height = 450;

                listView.Parent = form;
                listView.Dock = DockStyle.Fill;
                listView.View = View.Details;
                listView.Columns.Add("Дата", 200);

                DateTime date = dateTimePickerStartDate.Value.Date;
                while (date <= dateTimePickerEndDate.Value.Date)
                {
                    if (item.CheckDate(date)) // probe this date
                        listView.Items.Add(date.ToLongDateString());
                    date = date.AddDays(1);
                }
                form.Show();
            }
            else
                MessageBox.Show("Выберите триггер!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportItemToXML()
        {
            if (listViewItems.SelectedItems.Count > 0)
            {
                TaskScheduler.TriggerItem item = (TaskScheduler.TriggerItem)listViewItems.SelectedItems[0].Tag;
                //textBoxEvents.Clear();
            }
            else
                MessageBox.Show("Выберите триггер!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string ExportCollectionToXML()
        {
            String xmlString = String.Empty;
            try
            {
                xmlString = _taskScheduler.TriggerItems.ToXML();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: сериализация в XML: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return xmlString;
        }

        private void ImportCollectionFromXML(String xmlString)        
        {
            _taskScheduler.TriggerItems.Clear();
            try
            {
                TaskScheduler.TriggerItemCollection items = TaskScheduler.TriggerItemCollection.FromXML(xmlString);
                _taskScheduler.TriggerItems.AddRange(items, new TaskScheduler.TriggerItem.OnTriggerEventHandler(triggerItem_OnTrigger));
                _taskScheduler.Enabled = checkBoxEnabled.Checked; // Start the Scheduler
                UpdateTaskList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: парсинг XML: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportCollectionFromXML()
        {
            //ImportCollectionFromXML(textBoxEvents.Text);
        }

        private String GetServiceConfigFileName()
        {
            String commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            String configDirectory = commonAppData + Path.DirectorySeparatorChar + "TaskScheduler";
            return configDirectory + Path.DirectorySeparatorChar + "SchedulerItems.xml";
        }

        private void ReadServiceConfig()
        {
            ResetScheduler();

            String configFile = GetServiceConfigFileName();

            String xmlString = String.Empty;
            try
            {
                xmlString = System.IO.File.ReadAllText(configFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(("Невозможно прочитать конифгурацию: " + configFile + ": " + ex.Message), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                ImportCollectionFromXML(xmlString);
            }
        }

        private void SaveAsServiceConfig()
        {
            if (_taskScheduler.TriggerItems.Count == 0)
            {
                MessageBox.Show("Создайте триггеры!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            String xmlString = ExportCollectionToXML();
            String configFile = GetServiceConfigFileName();

            String directory = Path.GetDirectoryName(configFile);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using (StreamWriter outfile = new StreamWriter(configFile))
            {
                try
                {
                    outfile.Write(xmlString);
                    MessageBox.Show("Конфигурацие успешно сохранена!" + Environment.NewLine + Environment.NewLine + "Имя файла: " + configFile, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: сохранение в XML: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static void InstallService()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                TaskSchedulerServiceAssistant.Install();
                MessageBox.Show("Установка сервиса успешна", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: установка сервиса: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private static void UninstallService()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                TaskSchedulerServiceAssistant.Uninstall();
                MessageBox.Show("Удаление сервиса успешно", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: деинсталляция сервиса: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private static void StartService()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                TaskSchedulerServiceAssistant.StartService();
                MessageBox.Show("Запуск сервиса успешен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: запуск сервиса: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private static void StopService()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                TaskSchedulerServiceAssistant.StopService();
                MessageBox.Show("Остановка сервиса успешна", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: остановка сервиса: " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        void triggerItem_OnTrigger(object sender, TaskScheduler.OnTriggerEventArgs e)
        {
            String nextTrigger = String.Empty;
            if (e.Item.GetNextTriggerDateTime() != DateTime.MaxValue)
                nextTrigger = e.Item.GetNextTriggerDateTime().DayOfWeek.ToString() + ", " + e.Item.GetNextTriggerDateTime().ToString();
            else
                nextTrigger = "Никогда";
            UpdateTaskList();
        }

        private void buttonCreateTrigger_Click(object sender, EventArgs e)
        {
            CreateSchedulerItem();
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            ResetScheduler();
        }

        private void buttonShowAllTrigger_Click(object sender, EventArgs e)
        {
            ShowAllTriggerDates();
        }

        private void buttonToXML_Click(object sender, EventArgs e)
        {
            ExportItemToXML();
        }

        private void checkBoxEnabled_CheckedChanged(object sender, EventArgs e)
        {
            _taskScheduler.Enabled = checkBoxEnabled.Checked;
        }

        private void buttonCollectionToXML_Click(object sender, EventArgs e)
        {
            //textBoxEvents.Clear();
            //textBoxEvents.AppendText(ExportCollectionToXML());
        }

        private void buttonCollectionFromXML_Click(object sender, EventArgs e)
        {
            ImportCollectionFromXML();
        }

        private void buttonSaveAsServiceConfig_Click(object sender, EventArgs e)
        {
            SaveAsServiceConfig();
        }

        private void buttonReadServiceConfig_Click(object sender, EventArgs e)
        {
            ReadServiceConfig();
        }

        private void buttonInstallService_Click(object sender, EventArgs e)
        {
            InstallService();
        }

        private void buttonUninstallService_Click(object sender, EventArgs e)
        {
            UninstallService();
        }

        private void buttonStartService_Click(object sender, EventArgs e)
        {
            StartService();
        }

        private void buttonStopService_Click(object sender, EventArgs e)
        {
            StopService();
        }
    }
}