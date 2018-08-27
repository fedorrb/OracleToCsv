using System;
using System.Windows.Forms;
using System.IO;

namespace LsDumpToCsv
{
    public class stPath
    {
        public int delayMS { get; set; }//час очікування
        public string connectionString { get; set; }//строка зєднання з БД Оракл
        public string lsPath { get; set; } //шлях до папки обміну
        public int numberFilesWait { get; set; } //кількість файлів для очікування роботи модуля АСОПД
        public int stopDelay { get; set; } //час очікування до виходу з програми, якщо модуль АСОПД не видаляє файли

        public stPath()
        {
            delayMS = 5000;
            connectionString = @"Data Source=SPP181;User Id=IKIS;Password=IKIS;";
            lsPath = "C:\\WORK\\BASE\\ORACLE\\";
        }
        /// <summary>
        /// Чтение файла конфигурации
        /// </summary>
        public void LoadIniFile()
        {
            string path = String.Empty;
            path = System.IO.Directory.GetCurrentDirectory();
            path += @"\LsDumpToCsv.ini";
            int elements = 0;
            if (System.IO.File.Exists(path))
            {
                try
                {
                    string[] allLines = System.IO.File.ReadAllLines(path);
                    elements = allLines.Length;
                    switch (elements)
                    {
                        case 5:
                            stopDelay = int.Parse(allLines[4]);
                            goto case 4;
                        case 4:
                            numberFilesWait = int.Parse(allLines[3]);
                            goto case 3;
                        case 3:
                            lsPath = allLines[2];
                            goto case 2;
                        case 2:
                            delayMS = int.Parse(allLines[1]);
                            goto case 1;
                        case 1:
                            connectionString = allLines[0];
                            break;
                        case 0:
                            break;
                        default:
                            break;
                    }

                }
                catch
                {
                    delayMS = 5000;
                    connectionString = @"Data Source=SPP181;User Id=IKIS;Password=IKIS;";
                    lsPath = "C:\\WORK\\BASE\\ORACLE\\";
                }
            }

            if (delayMS < 10)
                delayMS = 5000;

        }

        /// <summary>
        /// Запись файла конфигурации
        /// </summary>
        public void SaveIniFile()
        {
            string[] s = { connectionString,
                             delayMS.ToString(),
                             lsPath,
                             numberFilesWait.ToString(),
                             stopDelay.ToString()
                         };
            string path = String.Empty;
            path = Application.StartupPath;
            path += @"\LsDumpToCsv.ini";
            System.IO.File.WriteAllLines(path, s);
        }
    }
}
