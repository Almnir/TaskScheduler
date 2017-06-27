using System;
using System.Collections;
using System.IO;
using System.Timers;
using System.Xml.Serialization;

namespace TaskScheduler
{
    public class TaskScheduler
    {
        /// <summary>
        /// Тип повтора
        /// </summary>
        public enum DayOccurrence
        {
            First = 0,
            Second = 1,
            Third = 2,
            Fourth = 3,
            Last = 4
        }

        /// <summary>
        /// День месяца
        /// </summary>
        public enum MonthOfTheYeay
        {
            January = 0,
            February = 1,
            March = 2,
            April = 3,
            May = 4,
            June = 5,
            July = 6,
            August = 7,
            September = 8,
            October = 9,
            November = 10,
            December = 11
        }

        /// <summary>
        /// Единовременный триггер
        /// </summary>
        public class TriggerSettingsOneTimeOnly
        {
            private DateTime _date;
            private bool _active;

            [XmlIgnore]
            public DateTime Date
            {
                get
                {
                    return _date;
                }
                set
                {
                    _date = value;
                }
            }

            [XmlElement("Date")]
            public string XMLDate
            {
                get { return this._date.ToString("yyyy-MM-dd"); }
                set { this.Date = DateTime.ParseExact(value, "yyyy-MM-dd", null).Date; }
            }

            public bool Active
            {
                get
                {
                    return _active;
                }
                set
                {
                    _active = value;
                }
            }
        }

        /// <summary>
        /// Ежедневный триггер
        /// </summary>
        public class TriggerSettingsDaily
        {
            private ushort _Interval;

            public ushort Interval
            {
                get
                {
                    return _Interval;
                }
                set
                {
                    _Interval = value;
                    if (_Interval < 0) _Interval = 0;
                }
            }
        }

        /// <summary>
        /// Еженедельный триггер
        /// </summary>
        public class TriggerSettingsWeekly
        {
            private bool[] _DaysOfWeek;
            /// <summary>
            /// Дней в неделю
            /// </summary>
            public bool[] DaysOfWeek
            {
                get
                {
                    return _DaysOfWeek;
                }
                set
                {
                    _DaysOfWeek = value;
                }
            }

            public TriggerSettingsWeekly()
            {
                _DaysOfWeek = new bool[7];
            }
        }

        /// <summary>
        /// Месячно-недельные настройки триггера
        /// </summary>
        public class TriggerSettingsMonthlyWeekDay
        {
            private bool[] _WeekNumber;
            private bool[] _DayOfWeek;

            /// <summary>
            /// Неделя
            /// </summary>
            public bool[] WeekNumber
            {
                get
                {
                    return _WeekNumber;
                }
                set
                {
                    _WeekNumber = value;
                }
            }
            /// <summary>
            /// Ежемесячные - будни
            /// </summary>
            public bool[] DayOfWeek
            {
                get
                {
                    return _DayOfWeek;
                }
                set
                {
                    _DayOfWeek = value;
                }
            }

            public TriggerSettingsMonthlyWeekDay()
            {
                _WeekNumber = new bool[5];
                _DayOfWeek = new bool[7];
            }
        }

        /// <summary>
        /// Настройки месячного триггера
        /// </summary>
        public class TriggerSettingsMonthly
        {
            private bool[] _Month;

            private bool[] _DaysOfMonth;
            private TriggerSettingsMonthlyWeekDay _WeekDay;

            /// <summary>
            /// Активные месяцы
            /// </summary>
            public bool[] Month
            {
                get
                {
                    return _Month;
                }
                set
                {
                    _Month = value;
                }
            }
            /// <summary>
            /// Дни в месяце
            /// </summary>
            public bool[] DaysOfMonth
            {
                get
                {
                    return _DaysOfMonth;
                }
                set
                {
                    _DaysOfMonth = value;
                }
            }
            /// <summary>
            /// Настройки триггера ежемесячно - будни
            /// </summary>
            public TriggerSettingsMonthlyWeekDay WeekDay
            {
                get
                {
                    return _WeekDay;
                }
                set
                {
                    _WeekDay = value;
                }
            }

            public TriggerSettingsMonthly()
            {
                _Month = new bool[12];
                _DaysOfMonth = new bool[32];
                _WeekDay = new TriggerSettingsMonthlyWeekDay();
            }

        }

        /// <summary>
        /// Настройки триггера
        /// </summary>
        public class TriggerSettings
        {
            private TriggerSettingsOneTimeOnly _OneTimeOnly;
            private TriggerSettingsDaily _Daily;
            private TriggerSettingsWeekly _Weekly;
            private TriggerSettingsMonthly _Monthly;

            public TriggerSettingsOneTimeOnly OneTimeOnly
            {
                get
                {
                    return _OneTimeOnly;
                }
                set
                {
                    _OneTimeOnly = value;
                }
            }
            public TriggerSettingsDaily Daily
            {
                get
                {
                    return _Daily;
                }
                set
                {
                    _Daily = value;
                }
            }
            public TriggerSettingsWeekly Weekly
            {
                get
                {
                    return _Weekly;
                }
                set
                {
                    _Weekly = value;
                }
            }
            public TriggerSettingsMonthly Monthly
            {
                get
                {
                    return _Monthly;
                }
                set
                {
                    _Monthly = value;
                }
            }

            public TriggerSettings()
            {
                _OneTimeOnly = new TriggerSettingsOneTimeOnly();
                _Daily = new TriggerSettingsDaily();
                _Weekly = new TriggerSettingsWeekly();
                _Monthly = new TriggerSettingsMonthly();
            }
        }

        /// <summary>
        /// Событие по триггеру
        /// </summary>
        public class OnTriggerEventArgs : EventArgs
        {
            public OnTriggerEventArgs(TriggerItem item, DateTime triggerDate)
            {
                _item = item;
                _triggerDate = triggerDate;
            }
            private TriggerItem _item;
            private DateTime _triggerDate;
            public TriggerItem Item
            {
                get { return _item; }
            }
            public DateTime TriggerDate
            {
                get { return _triggerDate; }
            }
        }

        /// <summary>
        /// Класс записи в панели задач - триггер как таковой
        /// </summary>
        public class TriggerItem
        {
            // начальная дата
            private DateTime _StartDate = DateTime.MinValue;
            // конечная дата
            private DateTime _EndDate = DateTime.MaxValue;
            // текущее выполнение триггера
            private DateTime _TriggerTime;
            // временное окно, в котором триггер активируется (Время пуска + HitTimeSpan)
            private TimeSpan _HitTimeSpan = new TimeSpan(0, 0, 1);
            // дата следующего триггера
            private DateTime _NextTriggerDate;
            // настройки когда триггер должен быть включен
            private TriggerSettings _TriggerSettings;
            // последний день месяца
            private const byte LastDayOfMonthID = 31;
            // вспомогательный объект
            private object _Tag;
            // активирован ли триггер
            private bool _Enabled;
            // делегат события
            public delegate void OnTriggerEventHandler(object sender, OnTriggerEventArgs e);
            // событие
            public event OnTriggerEventHandler OnTrigger;

            /// <summary>
            /// Создаёт экземпляр TriggerItem
            /// </summary>
            public TriggerItem()
            {
                _TriggerSettings = new TriggerSettings();
            }

            /// <summary>
            /// Сериализация объекта в XML-строку
            /// </summary>
            /// <returns></returns>
            public String ToXML()
            {
                XmlSerializer ser = new XmlSerializer(typeof(TriggerItem));
                TextWriter writer = new StringWriter();
                ser.Serialize(writer, this);
                writer.Close();
                return writer.ToString();
            }

            /// <summary>
            /// Десериализация из XML-строки
            /// </summary>
            /// <param name="Configuration"></param>
            /// <returns></returns>
            public static TriggerItem FromXML(string Configuration)
            {
                XmlSerializer ser = new XmlSerializer(typeof(TriggerItem));
                TextReader reader = new StringReader(Configuration);
                TriggerItem result = (TriggerItem)ser.Deserialize(reader);
                reader.Close();
                return result;
            }

            /// <summary>
            /// Вспомогательный объект
            /// </summary>
            [XmlElement(Order = 0)]
            public object Tag
            {
                get
                {
                    return _Tag;
                }
                set
                {
                    _Tag = value;
                }
            }

            /// <summary>
            /// начальная дата
            /// </summary>
            [XmlIgnore]
            public DateTime StartDate
            {
                get
                {
                    return _StartDate;
                }
                set
                {
                    _StartDate = value;
                    if (_EndDate < _StartDate) _EndDate = _StartDate;
                }
            }

            [XmlElement("StartDate", Order = 1)]
            public string XMLStartDate
            {
                get { return this._StartDate.ToString("yyyy-MM-dd"); }
                set { this.StartDate = DateTime.ParseExact(value, "yyyy-MM-dd", null); }
            }

            /// <summary>
            /// конечная дата
            /// </summary>
            [XmlIgnore]
            public DateTime EndDate
            {
                get
                {
                    return _EndDate;
                }
                set
                {
                    _EndDate = value.Date;
                }
            }
            
            [XmlElement("EndDate", Order = 2)]
            public string XMLEndDate
            {
                get { return this._EndDate.ToString("yyyy-MM-dd"); }
                set { this.EndDate = DateTime.ParseExact(value, "yyyy-MM-dd", null); }
            }

            /// <summary>
            /// текущее выполнение триггера
            /// </summary>
            [XmlIgnore]
            public DateTime TriggerTime
            {
                get
                {
                    return _TriggerTime;
                }
                set
                {
                    _TriggerTime = new DateTime(_TriggerTime.Year, _TriggerTime.Month, _TriggerTime.Day, value.Hour, value.Minute, value.Second);
                }
            }

            [XmlElement("TriggerTime", Order = 3)]
            public string XMLTriggerTime
            {
                get { return this.TriggerTime.ToString("HH:mm:ss"); } //yyyy-MM-dd 
                set { this.TriggerTime = DateTime.ParseExact(value, "HH:mm:ss", null); }
            }

            /// <summary>
            /// временное окно, в котором триггер активируется (Время пуска + HitTimeSpan)
            /// </summary>
            private TimeSpan HitTimeSpan
            {
                get
                {
                    return _HitTimeSpan;
                }
                set
                {
                    _HitTimeSpan = value;
                }
            }

            /// <summary>
            /// настройки когда триггер должен быть включен
            /// </summary>
            [XmlElement(Order = 4)]
            public TriggerSettings TriggerSettings
            {
                get
                {
                    return _TriggerSettings;
                }
                set
                {
                    _TriggerSettings = value;
                }
            }

            /// <summary>
            /// Активен или неактивен триггер
            /// </summary>
            [XmlElement(Order = 5)]
            public bool Enabled
            {
                get
                {
                    return _Enabled;
                }
                set
                {
                    _Enabled = value;
                    if (_Enabled)
                        _NextTriggerDate = FindNextTriggerDate(DateTime.Now);
                    else
                        _NextTriggerDate = DateTime.MaxValue;
                }
            }

            /// <summary>
            /// Определяет последний день месяца
            /// </summary>
            /// <param name="date"></param>
            /// <returns></returns>
            private DateTime LastDayOfMonth(DateTime date)
            {
                return new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            }

            /// <summary>
            /// Определите, какое количество раз в неделю в этом месяце на эту дату происходит
            /// </summary>
            /// <param name="date"></param>
            /// <returns></returns>
            private int WeekDayOccurInMonth(DateTime date)
            {
                byte count = 0;
                for (int day = 1; day <= date.Day; day++)
                    if (new DateTime(date.Year, date.Month, day).DayOfWeek == date.DayOfWeek)
                        count++;
                return count-1;
            }

            /// <summary>
            /// Определяет, является ли этот день последним рабочим днём в месяце
            /// </summary>
            /// <param name="date"></param>
            /// <returns></returns>
            private bool IsLastWeekDayInMonth(DateTime date)
            {
                return (WeekDayOccurInMonth(date.AddDays(7))==0);
            }

            /// <summary>
            /// Триггер однократного выполнения
            /// </summary>
            /// <returns></returns>
            private bool TriggerOneTimeOnly(DateTime date)
            {
                return (_TriggerSettings.OneTimeOnly.Active && (_TriggerSettings.OneTimeOnly.Date == date));
            }

            /// <summary>
            /// Ежедневный триггер
            /// </summary>
            private bool TriggerDaily(DateTime date)
            {
                if ((date < _StartDate.Date) || (date > _EndDate.Date))
                    return false;
                if (_TriggerSettings.Daily.Interval == 0)
                    return false;
                DateTime RunTime = _StartDate.Date;
                while (RunTime.Date < date)
                    RunTime = RunTime.AddDays(_TriggerSettings.Daily.Interval);
                return (RunTime == date);
            }

            /// <summary>
            /// Еженедельный триггер
            /// </summary>
            private bool TriggerWeekly(DateTime date)
            {
                if ((date < _StartDate.Date) || (date > _EndDate.Date))
                    return false;
                return (_TriggerSettings.Weekly.DaysOfWeek[(int)date.DayOfWeek]);
            }

            /// <summary>
            /// Ежемесячный триггер
            /// </summary>
            private bool TriggerMonthly(DateTime date)
            {
                // если вне диапазона дат - возвращаем false
                if ((date < _StartDate.Date) || (date > _EndDate.Date))
                    return false;

                bool result = false;
                // если в этом месяце
                if (_TriggerSettings.Monthly.Month[date.Month - 1])
                {
                    // попадает ли на последний день месяца
                    if (_TriggerSettings.Monthly.DaysOfMonth[LastDayOfMonthID])
                        // последний день месяца?
                        result = (result || (date == LastDayOfMonth(date)));

                    // день последний?
                    result = (result || (_TriggerSettings.Monthly.DaysOfMonth[date.Day - 1]));

                    // активен ли день?
                    if (_TriggerSettings.Monthly.WeekDay.DayOfWeek[(int)date.DayOfWeek])
                    {
                        // проверить последнее вхождение в день в течении месяца
                        if (_TriggerSettings.Monthly.WeekDay.WeekNumber[(int)DayOccurrence.Last])
                            result = (result || (IsLastWeekDayInMonth(date)));

                        // активируется N'ное появление?
                        result = (result || _TriggerSettings.Monthly.WeekDay.WeekNumber[WeekDayOccurInMonth(date)]);
                    }
                }
                return result;
            }

            /// <summary>
            /// Проверить дату по триггерам
            /// </summary>
            /// <returns></returns>
            public bool CheckDate(DateTime date)
            {
                return (TriggerOneTimeOnly(date) || TriggerDaily(date) || TriggerWeekly(date) || TriggerMonthly(date));
            }

            /// <summary>
            /// Происходит ли до определённого момента времени
            /// </summary>
            /// <param name="dateTime"></param>
            /// <returns></returns>
            public bool RunCheck(DateTime dateTimeToCheck)
            {
                if (dateTimeToCheck == DateTime.MaxValue)
                    return false; 

                if (_Enabled)
                {
                    DateTime triggerDateTime = GetNextTriggerDateTime();
                    if ((dateTimeToCheck >= triggerDateTime) && (dateTimeToCheck <= triggerDateTime.AddTicks(_HitTimeSpan.Ticks)))
                    {
                        OnTriggerEventArgs eventArgs = new OnTriggerEventArgs(this, triggerDateTime);
                        // продолжить триггер
                        _NextTriggerDate = FindNextTriggerDate(_NextTriggerDate.AddDays(1));
                        // проверяем и исполняем
                        OnTrigger?.Invoke(this, eventArgs);
                        return true;
                    }
                }
                return false;
            }

            /// <summary>
            /// Поиск следующего времени выполнения
            /// </summary>
            private DateTime FindNextTriggerDate(DateTime searchStartDateTime)
            {
                DateTime date = searchStartDateTime.Date;

                // если активно, продолжаем поиск следующим днём
                if (searchStartDateTime.TimeOfDay > _TriggerTime.TimeOfDay)
                    date = date.AddDays(1);

                while (date <= _EndDate)
                {
                    if (CheckDate(date))
                        return date;
                    date = date.AddDays(1);
                }
                return DateTime.MaxValue;
            }

            /// <summary>
            /// Определяет дату следующего исполнения, независимо от времени запуска
            /// </summary>
            /// <returns></returns>
            public DateTime GetNextTriggerDateTime()
            {
                if ((!_Enabled) || (_NextTriggerDate==DateTime.MaxValue))
                    return DateTime.MaxValue;
                return new DateTime(_NextTriggerDate.Year, _NextTriggerDate.Month, _NextTriggerDate.Day, _TriggerTime.Hour, _TriggerTime.Minute, _TriggerTime.Second);
            }
        }

        /// <summary>
        /// Коллекция TriggerItem
        /// </summary>
        [XmlRoot(ElementName = "TriggerItemCollection")]
        public class TriggerItemCollection : CollectionBase
        {
            public TriggerItem this[int index]
            {
                get
                {
                    return ((TriggerItem)List[index]);
                }
                set
                {
                    List[index] = value;
                }
            }

            public int Add(TriggerItem value)
            {
                return (List.Add(value));
            }

            public void AddRange(TriggerItemCollection items, TriggerItem.OnTriggerEventHandler handler)
            {
                foreach (TriggerItem item in items)
                {
                    item.OnTrigger += handler;
                    Add(item);
                }
            }

            public int IndexOf(TriggerItem value)
            {
                return (List.IndexOf(value));
            }

            public void Insert(int index, TriggerItem value)
            {
                List.Insert(index, value);
            }

            public void Remove(TriggerItem value)
            {
                List.Remove(value);
            }

            public bool Contains(TriggerItem value)
            {
                return (List.Contains(value));
            }

            protected override void OnInsert(int index, Object value)
            {
            }

            protected override void OnRemove(int index, Object value)
            {
            }

            protected override void OnSet(int index, Object oldValue, Object newValue)
            {
            }

            protected override void OnValidate(Object value)
            {
                if (value.GetType() != typeof(TaskScheduler.TriggerItem))
                    throw new ArgumentException("Указанный элемент не входит в планировщик задач!", "value");
            }

            /// <summary>
            /// Сериализация в XML-строку
            /// </summary>
            /// <returns></returns>
            public String ToXML()
            {
                XmlSerializer ser = new XmlSerializer(typeof(TriggerItemCollection));
                TextWriter writer = new StringWriter();
                ser.Serialize(writer, this);
                writer.Close();
                return writer.ToString();
            }

            /// <summary>
            /// Десериализация из XML
            /// </summary>
            /// <param name="Configuration"></param>
            /// <returns></returns>
            public static TriggerItemCollection FromXML(String Configuration)
            {
                XmlSerializer ser = new XmlSerializer(typeof(TriggerItemCollection));
                TextReader reader = new StringReader(Configuration);
                TriggerItemCollection result = (TriggerItemCollection)ser.Deserialize(reader);
                reader.Close();
                return result;
            }
        }

        /// <summary>
        /// Колекция TriggerItem
        /// </summary>
        private TriggerItemCollection _triggerItems;

        /// <summary>
        /// Пауза между триггерными проверками в миллисекундах
        /// </summary>
        private int _Interval = 500;

        /// <summary>
        /// Планировщик активен?
        /// </summary>
        private bool _Enabled = false;

        /// <summary>
        /// Проверочный таймер для триггера
        /// </summary>
        private Timer _triggerTimer;

        /// <summary>
        /// Основной конструктор класса
        /// </summary>
        public TaskScheduler()
        {
            _triggerItems = new TriggerItemCollection();
            _triggerTimer = new Timer();
            _triggerTimer.Elapsed += new ElapsedEventHandler(_triggerTimer_Tick);
        }

        /// <summary>
        /// Определяет интервал триггера
        /// </summary>
        public int Interval
        {
            get
            {
                return _Interval;
            }
            set
            {
                _Interval = value;
                _triggerTimer.Stop();
                _triggerTimer.Interval = _Interval;
                _triggerTimer.Start();
            }
        }

        /// <summary>
        /// Активация/деактивация планировщика
        /// </summary>
        public bool Enabled
        {
            get
            {
                return _Enabled;
            }
            set
            {
                _Enabled = value;
                if (_Enabled) Start();
                else
                    Stop();
            }
        }

        /// <summary>
        /// Вызывает объект, используемый для маршалинга вызовов обработчика событий
        /// </summary>
        public System.ComponentModel.ISynchronizeInvoke SynchronizingObject
        {
            get
            {
                return _triggerTimer.SynchronizingObject;
            }
            set
            {
                _triggerTimer.SynchronizingObject = value;
            }
        }

        /// <summary>
        /// Добавить триггер
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public TriggerItem AddTrigger(TriggerItem item)
        {
            return _triggerItems[_triggerItems.Add(item)];
        }

        /// <summary>
        /// Активация планировщика
        /// </summary>
        private void Start()
        {
            if (_triggerTimer.Enabled)
                _triggerTimer.Stop();
            _triggerTimer.Interval = _Interval;
            _triggerTimer.Start();
        }

        /// <summary>
        /// Остановить планировщик
        /// </summary>
        private void Stop()
        {
            _triggerTimer.Stop();
        }

        /// <summary>
        /// Представляет список элементов триггера
        /// </summary>
        public TriggerItemCollection TriggerItems
        {
            get
            {
                return _triggerItems;
            }
        }

        /// <summary>
        /// Обработка событий для таймер-тика
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _triggerTimer_Tick(object source, ElapsedEventArgs e)
        {
            _triggerTimer.Stop();
            foreach (TriggerItem item in TriggerItems)
                item.RunCheck(DateTime.Now);
            _triggerTimer.Start();
        }
    }
}
