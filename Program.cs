using System;
using System.ServiceProcess;
using System.Windows.Forms;

namespace TaskScheduler
{
    static class Program
    {
        private static void ShowHelp()
        {
            MessageBox.Show("Параметры:" + Environment.NewLine +
                "-i, --install: \tУстановить сервис" + Environment.NewLine +
                "-u, --uninstall: \tУдалить сервис" + Environment.NewLine +
                "-s, --start: \t\tЗапустить сервис" + Environment.NewLine +
                "-t, --stop: \t\tОстановить сервис" + Environment.NewLine +
                "-h, --help: \tОтобразить это сообщение",
                "Task Scheduler", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// Точка входа приложения
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            bool install = false, uninstall = false, start = false, stop = false,  service = false;
            bool runConfiguration = true;
            try
            {
                foreach (string arg in args)
                {
                    switch (arg)
                    {
                        case "-i":
                        case "--install":
                            install = true; break;
                        case "-u":
                        case "--uninstall":
                            uninstall = true; break;
                        case "-s":
                        case "--start":
                            start = true; break;
                        case "-t":
                        case "--stop":
                            stop = true; break;
                        case "--service":
                            service = true; break;
                        default:
                            ShowHelp();
                            return 0;
                    }
                }

                if (uninstall)
                {
                    runConfiguration = false;
                    TaskSchedulerServiceAssistant.Uninstall();
                }

                if (install)
                {
                    runConfiguration = false;
                    TaskSchedulerServiceAssistant.Install();
                }

                if (start)
                {
                    runConfiguration = false;
                    TaskSchedulerServiceAssistant.StartService();
                }

                if (stop)
                {
                    runConfiguration = false;
                    TaskSchedulerServiceAssistant.StopService();
                }

                if (service)
                {
                    runConfiguration = false;
                    ServiceBase[] services = { new TaskSchedulerService() };
                    ServiceBase.Run(services);
                }

                if (runConfiguration)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new TaskManager());
                }

                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Необработанное исключение: \r\n" + ex.ToString(), "Task Scheduler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }
    }
}