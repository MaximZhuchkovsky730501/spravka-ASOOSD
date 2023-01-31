using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;

namespace ConsoleApplication1
{
    class Server
    {
        public String name, ip, service_name;
        public Server(String n, String i, String s)
        {
            name = n; //имя подразделения
            ip = i; //ip-адрес для подключения
            service_name = s; //servise name базы
        }
    }

    class Settings
    {
        private String login, password, location_exaple, location_output;
        private List<String> server_lines;

        public Settings()
        {
            try // конструкций для обработки исключений try - catch
            {
                String text = read_settings(); //чтение данных из файла
                int ind = text.IndexOf("begin") + 7;
                text = text.Substring(ind); //удаляем слово "begin", и  весь текст до него
                String[] lines = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); // разбиваем текст по строкам
                location_exaple = lines[0]; //получаем путь к шаблону
                location_output = lines[1]; // получаем путь для выгрузки файлов
                login = lines[2].Split('/')[0]; // получаем логин
                password = lines[2].Split('/')[1]; // получаем пароль
                server_lines = new List<string>(); // создаём список с информацией о серверах
                foreach (String line in lines)
                    if (!line.StartsWith(location_exaple) && !line.StartsWith(location_output) && line != login + '/' + password)
                        server_lines.Add(line); //добовляем строки с информацией о серверах в список 
            }
            catch (Exception e)
            {
                Console.WriteLine("ошибка интерпритации данных из файла настроек, возможно данные записаны с ошибками или в другой последовательности\n" + e.ToString());
                Console.ReadKey(); // ожидание нажатия на любую клавишу, чтобы консоль не закрывалась автоматически
                System.Diagnostics.Process.GetCurrentProcess().Kill(); // завершение работы приложения
            }
        }

        public String get_login()
        {
            return login;
        }

        public String get_password()
        {
            return password;
        }

        public String get_location_exaple()
        {
            return location_exaple;
        }

        public String get_location_output()
        {
            return location_output;
        }

        public List<String> get_server_lines()
        {
            return server_lines;
        }

        static private String read_settings()
        {
            String file_text = null;
            try
            {
                FileStream file = File.OpenRead("settings.txt"); // открываем файл "settings.txt" в корневой папке приложения
                byte[] array = new byte[file.Length];
                int count = file.Read(array, 0, array.Length); // считываем байты из файла в массив "array", count - колличество успешно считанных байт
                file_text = System.Text.Encoding.UTF8.GetString(array); //декодируем байты в текст
                file.Close(); // закрываем файл
            }
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine("не удалось найти путь к файлу \"settings.txt\"\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("файл \"settings.txt\" не найден, он должен распологаться в папке приложения\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            catch (FileLoadException e)
            {
                Console.WriteLine("не удалось загрузить файл \"settings.txt\"\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            catch (Exception e)
            {
                Console.WriteLine("неизвестная ошибка при чтении файла \"settings.txt\"\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            return file_text;
        }
    }

    class Information
    {
        private String division, title; //имя подразделения, имя параметра
        private int row, col; //номер строки, номер колонки
        private DateTime date; //дата параметра
        private static DateTime null_date = new DateTime(0001, 01, 01, 00, 00, 00); //дата, получаемая при отсутствии данных (для сравнения): 01.01.0001 00:00:00
        private static Boolean is_col_init = false; //словарь соответствия названий подразделений и колонок создан? false - нет, true - да

        public static Dictionary<String, int> division_col = new Dictionary<String, int>(); //словарь соответствия названий подразделений и колонок
        public static Dictionary<String, int> title_row = new Dictionary<String, int>() //словарь соответствия названий параметров и строк
        {
            {"раздел 6", 6}, //название параметра, номер строки
            {"раздел 7", 7},
            {"раздел 8", 8},
            {"раздел 9", 9},
            {"раздел 10", 10},
            {"раздел 11", 11},
            {"раздел 12", 12},
            {"раздел 13", 13},
            {"раздел 14", 14},
            {"раздел 18", 18},
            {"раздел 27", 27},
            {"раздел 28", 28},
        };

        public Information(String d, String t, List<Server> s)
        {
            division = d;
            title = t;
            if (!is_col_init) //если словарь соответствия названий подразделений и колонок не создан
            {
                int i = 3; //начальный номер столбца
                foreach (Server tmp in s) //для каждого сервера из списка
                {
                    //в строке ниже оператор ++ выполняется перед функцией .Add()
                    division_col.Add(tmp.name, ++i); //добавление соответствия в словарь (имя сервера, номер столбца)
                }
                is_col_init = true; //указываем, что словарь соответствия названий подразделений и колонок создан
            }
            col = division_col[division]; //присваиваем номер колонки соответствующий имени подразделения
            row = title_row[title]; //присваиваем номер строки соответствующий названию параметра
        }

        public Information(String d, String t) //альтернативный конструктор, всё аналочно с описаным выше
        {
            division = d;
            title = t;
            col = division_col[division];
            row = title_row[title];
        }

        public void set_date(String d)
        {
            date = Convert.ToDateTime(d);  //преобразует символьное значение в дату                                                 
        }

        public String get_str_date()
        {
            if (date != null_date) //если дата не соответствует дате, получаемой при отсутствии данных
                if (date.Year < DateTime.Now.Year) //если дата не относится к текущему календарному году
                    return date.ToString("dd.MM.yyyy"); //возвращается дата в формате "дд.мм.гггг"
                else
                    return date.ToString("dd.MM"); //возвращается дата в формате "дд.мм"
            else
                return "null"; //дата соответствует дате, получаемой при отсутствии данных, значит возвращаем значение "null"
        }

        public static String date_to_str(DateTime d) //преобразует полученную дату в символьное значение. аналогично функции выше
        {
            if (d != null_date)
                if (d.Year < DateTime.Now.Year)
                    return d.ToString("dd.MM.yyyy");
                else
                    return d.ToString("dd.MM");
            else
                return "null";
        }

        public DateTime get_date()
        {
            return date;
        }

        public String get_time()
        {
            return date.ToString("HH:mm:ss"); // возвращает время в формате чч:мм:сс
        }

        public String get_title()
        {
            return title.ToString();
        }

        public String get_division()
        {
            return division.ToString();
        }

        public int get_row()
        {
            return row;
        }

        public int get_col()
        {
            return col;
        }
    }   

    class Program
    {
        static public List<String> init_script() //возвращает скрипт в ввиде списка команд
        {
            List<String> command = new List<String>();
            command.Add("SELECT max(TG_TIME) AS \"раздел 6\" FROM SCHEMA.P6");
            command.Add("SELECT max(CO_TIME) AS \"раздел 8\" FROM SCHEMA.P8");
            command.Add("SELECT max(RZ_TIME) AS \"раздел 9\" FROM SCHEMA.P9");
            command.Add("SELECT max(PP_TIME_IN) AS \"раздел 10\" FROM SCHEMA.P10");
            command.Add("SELECT max(AM_TIME) AS \"раздел 11\" FROM SCHEMA.P11");
            command.Add("SELECT max(DETECTION_DATE) AS \"раздел 12\" FROM SCHEMA.P12");
            command.Add("SELECT max(DETECTION_DATE) AS \"раздел 13\" FROM SCHEMA.P13");
            command.Add("SELECT max(AT_DATE_IN) AS \"раздел 14\" FROM SCHEMA.P14 WHERE OPTION = 1");
            command.Add("SELECT max(AT_DATE_OUT) AS \"раздел 7\" FROM SCHEMA.P7 WHERE OPTION = 1");
            command.Add("SELECT max(PL_MONTH) AS \"раздел 18\" FROM SCHEMA.P18");
            command.Add("SELECT max(AT_DATE_IN) AS \"раздел 27\" FROM SCHEMA.P27 WHERE OPTION = 0");
            command.Add("SELECT max(AT_DATE_OUT) AS \"раздел 28\" FROM SCHEMA.P28 WHERE OPTION = 0");
            return command;
        }

        static public List<Server> init_server_list(List<String> lines)
        {
            List<Server> server = new List<Server>(); //создаём список серверов
            foreach (String line in lines)
            {
                String[] words = line.Split('/'); //разделяем строку на подстроки по символу "/" (содержимое строки: "имя подразделения/ip-адрес для подключения/servise name базы")
                server.Add(new Server(words[0], words[1], words[2])); //создаём экземпляры класса Server и добавляем их в список
            }
            return server; // возвращаем список серверов
        }

        static public List<Information> script_run(List<Server> server_list, String login, String password)
        {
            List<Information> info_list = new List<Information>(); //список параметров, полученных из БД АСООСД
            OracleConnection con = new OracleConnection();  //текущее подключение                       \
            OracleCommand cmd = new OracleCommand();        //исполняемая команда                        > переменные необходимые для подключения к БД
            String connectionString;                        //стока для открытия БД (как в tnsnames)    /
            List<String> script = new List<String>();
            script = init_script(); //получаем список команд
            Console.WriteLine("получение данных...");
            foreach (Server serv in server_list) //для каждого сервера из списка
            {
                try
                {
                    Console.Write(serv.name);
                    connectionString = "Data Source = (DESCRIPTION = " +
                                                       "(ADDRESS = (PROTOCOL = TCP)(HOST =  " + serv.ip + ")(PORT = 8888)) " +
                                                       "(CONNECT_DATA = " +
                                                         "(SERVICE_NAME =  " + serv.service_name + ") " +
                                                       ") " +
                                                       ");User Id = " + login + ";password=" + password;
                    con.ConnectionString = connectionString;
                    cmd.Connection = con;
                    con.Open(); //подключется к БД с использованием ранее заданных свойств
                    Console.WriteLine(" подключено");
                    String title; //имя параметра из БД
                    int ind1, ind2;
                    foreach (String str in script) //для каждой команды из списка
                    {
                        cmd.CommandText = str; //определяем команду для выполнения на сервере 
                        ind1 = ind2 = 0;                                    // \
                        ind1 = str.IndexOf(" \"");                          //  достаём из текста комманды
                        ind2 = str.IndexOf("\" ");                          //  имя параметра
                        title = str.Substring(ind1 + 2, ind2 - ind1 - 2);   // /
                        OracleDataReader dr = cmd.ExecuteReader(0); //выполняем команду
                        dr.Read(); //получаем результат выполнения (дату последнего обновления параметра)
                        if (dr.IsDBNull(0)) //если результат отсутствует
                        {
                            Information info = new Information(serv.name, title, server_list); //создаём экземпляр класса Information с информацией о параметре
                            info.set_date(null); //устанавливаем нулевую дату
                            info_list.Add(info); //добавляем в список
                        }
                        else
                        {
                            String buf = dr.GetValue(0).ToString(); //записываем результат в буфер
                            Information info = new Information(serv.name, title, server_list); //создаём экземпляр класса Information с информацией о параметре
                            info.set_date(buf); //устанавливаем дату из буфера
                            info_list.Add(info); //добавляем в список
                        }
                    }
                    con.Close(); //закрываем соединение
                }
                catch (OracleException e)
                {
                    Console.WriteLine("\nПри подключении к " + serv.name + "возникла ошибка. проверьте параметры подключения в файле настроек");
                    Console.WriteLine(e.ToString());
                }
                catch (Exception e)
                {
                    Console.WriteLine("\n", e.ToString());
                }
            }
            return info_list; //возвращем список с информацией о параметрах
        }

        static public void create_table(List<Information> info_list, String location_exaple, String location_output)
        {
            String catalog = location_exaple; 
            try
            {
                Application excel = new Application();
                Workbook book = excel.Workbooks.Open(catalog, Type.Missing, true); //открываем .excel файл
                _Worksheet sheet = book.Sheets[1]; //выбираем первую страницу

                DateTime today = DateTime.Now; //получаем текущую дату
                while (today.DayOfWeek != DayOfWeek.Friday) //\
                {                                           // \
                    today = today.AddDays(1);               //  находим дату ближайшей пятницы
                }                                           // /
                DateTime last_date = today.AddDays(-7); //находим дату прошлой пятницы
                DateTime begin_of_month = new DateTime(today.Year, today.Month, 01); //первое число месяца ближайшей пятницы

                sheet.Cells[1][1] = "ВЕДОМОСТЬ КОНТРОЛЯ\nзаполнения разделов АСООСД подразделениями границы войсковой части 2044 по состоянию на период с 08:00 " + //замена 
                    last_date.ToString("dd.MM.yyyy") + " по 08:00 " + today.ToString("dd.MM.yyyy");                                                                 //загаловка

                foreach (Information inf in info_list) //для каждого параметра из списка
                {
                    if (inf.get_title() == "План календарь ООМ ОСД") //для параметра "План календарь ООМ ОСД"
                    {
                        if (inf.get_date() < begin_of_month) //сравниваем с первым числом месяца
                        {
                            sheet.Cells[inf.get_col()][inf.get_row()] = inf.get_str_date(); //меняем значение ячейки
                            sheet.Cells[inf.get_col()][inf.get_row()].Interior.Color = XlRgbColor.rgbSteelBlue; //меняем цвет ячейки
                        }
                        continue; //переходим к следующему
                    }
                    ////////////////////////////////////////
                    if (inf.get_division() == "рез. погз Поставы" || inf.get_division() == "рез. погз Гудогай")                                                 //
                        continue;                                                                                                                               //
                    if (inf.get_title() == "Выход транспортных средств" && (inf.get_division() == "опк Мольдевичи" || inf.get_division() == "опк Котловка" ||   //
                        inf.get_division() == "опк Гудогай" || inf.get_division() == "опк Лоша" || inf.get_division() == "опк Каменный лог"))                   //если поле не заполняется, переходим к следующему
                        continue;                                                                                                                               //
                    if ((inf.get_title() == "Действия по обстановке" || inf.get_title() == "Журнал наблюдений") && inf.get_division() == "опк Гудогай")         //
                        continue;                                                                                                                               //
                    ////////////////////////////////////////
                    if (inf.get_title() == "Мониторинг ОС") //в данном разделе в сравнении участвуют два параметра: "Мониторинг ОС" и "Мониторинг радиационного фона"
                    {
                        Information temp = info_list.Find(item => (item.get_title() == "Мониторинг радиационного фона" && item.get_division() == inf.get_division()));
                        DateTime max_date;
                        if (inf.get_date() < temp.get_date())   // 
                            max_date = temp.get_date();         //выбираем большую
                        else                                    //из дат
                            max_date = inf.get_date();          //
                        if (max_date < last_date) //сравниваем с прошлой пятницей
                        {
                            sheet.Cells[inf.get_col()][inf.get_row()] = Information.date_to_str(max_date);
                            sheet.Cells[inf.get_col()][inf.get_row()].Interior.Color = XlRgbColor.rgbSteelBlue;
                        }
                        continue;
                    }
                    if (inf.get_title() == "Мониторинг радиационного фона") //этот параметр был рассмотрен в предыдущем условии
                        continue;
                    ///////////////////////////////////////////
                    if (inf.get_title() == "Учет ДЛ вьезд")//в данном разделе в сравнении участвуют два параметра: "Учет ДЛ вьезд" и "Учет ДЛ выезд"
                    {
                        Information temp = info_list.Find(item => (item.get_title() == "Учет ДЛ выезд" && item.get_division() == inf.get_division()));
                        DateTime max_date;
                        if (inf.get_date() < temp.get_date())
                            max_date = temp.get_date();
                        else
                            max_date = inf.get_date();
                        if (max_date < last_date)
                        {
                            sheet.Cells[inf.get_col()][inf.get_row()] = Information.date_to_str(max_date);
                            sheet.Cells[inf.get_col()][inf.get_row()].Interior.Color = XlRgbColor.rgbSteelBlue;
                        }
                        continue;
                    }
                    if (inf.get_title() == "Учет ДЛ выезд") //этот параметр был рассмотрен в предыдущем условии
                        continue;
                    //////////////////////////////////////////
                    if (inf.get_title() == "Прикомандированные вьезд")//в данном разделе в сравнении участвуют два параметра: "Прикомандированные вьезд" и "Прикомандированные выезд"
                    {
                        Information temp = info_list.Find(item => (item.get_title() == "Прикомандированные выезд" && item.get_division() == inf.get_division()));
                        DateTime max_date;
                        if (inf.get_date() < temp.get_date())
                            max_date = temp.get_date();
                        else
                            max_date = inf.get_date();
                        if (max_date < last_date)
                        {
                            sheet.Cells[inf.get_col()][inf.get_row()] = Information.date_to_str(max_date);
                            sheet.Cells[inf.get_col()][inf.get_row()].Interior.Color = XlRgbColor.rgbSteelBlue;
                        }
                        continue;
                    }
                    if (inf.get_title() == "Прикомандированные выезд") //этот параметр был рассмотрен в предыдущем условии
                        continue;
                    //////////////////////////////////////////
                    if (inf.get_str_date() == "null") //если дата отсутствует, переходим к следующему
                        continue;
                    if (inf.get_date() < last_date)
                    {
                        sheet.Cells[inf.get_col()][inf.get_row()] = inf.get_str_date();
                        sheet.Cells[inf.get_col()][inf.get_row()].Interior.Color = XlRgbColor.rgbSteelBlue;
                    }
                }
                catalog = location_output; //путь для выгрузки файла
                if (Directory.Exists(catalog)) //проверка существует ли путь для выгрузки
                {
                    catalog = catalog + "\\" + today.ToString("yyyy");
                    if (Directory.Exists(catalog)) //проверка существует ли папка с годом
                    {
                        catalog = catalog + "\\" + today.ToString("MMMM");
                        if (Directory.Exists(catalog)) //проверка существует ли папка с месяцем
                        {
                            catalog = catalog + "\\Справка по АСООСД " + today.ToString("dd.MM.yyyy") + ".xlsx"; //окончательный путь к файлу
                        }
                        else
                        {
                            Directory.CreateDirectory(catalog); //создание каталога
                            catalog = catalog + "\\Справка по АСООСД " + today.ToString("dd.MM.yyyy") + ".xlsx";
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(catalog);
                        catalog = catalog + "\\" + today.ToString("MMMM");
                        Directory.CreateDirectory(catalog);
                        catalog = catalog + "\\Справка по АСООСД " + today.ToString("dd.MM.yyyy") + ".xlsx";
                    }
                }
                else
                {
                    Console.WriteLine("заданный путь не сеществует. файл сохранён на диске C");
                    catalog = "C:\\Справка по АСООСД " + today.ToString("dd.MM.yyyy") + ".xlsx";// если путь для выгрузки не существует, файл сохраняется в корень диска С
                }
                book.SaveAs(catalog); //сохранение файла
                book.Close();//закрытие книги excel
                excel.Quit();                                                   //окончание работы
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); //с excel
                Console.WriteLine("файл сохранён по адресу " + catalog);
            }
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine("не удалось найти путь к файлу " + catalog + ", проверьте путь к файлу указанный в файле настроек\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            catch (FileLoadException e)
            {
                Console.WriteLine("не удалось загрузить файл " + catalog + "\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            catch (Exception e)
            {
                Console.WriteLine("ошибка при создании таблицы\n" + e.ToString());
                Console.ReadKey();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }

        }

        static void Main(string[] args)
        {
            Settings settings = new Settings(); //получаем основные параметры

            ////////////////////// Oracle ///////////////////////

            List<Server> server_list = new List<Server>(); //создаём список серверов
            List<Information> info_list = new List<Information>(); //создаём список полученных с сервера данных
            server_list = init_server_list(settings.get_server_lines()); //заполняем список серверов
            info_list = script_run(server_list, settings.get_login(), settings.get_password()); //заполняем список полученных данных

            ////////////////////// Excel ///////////////////////

            create_table(info_list, settings.get_location_exaple(), settings.get_location_output()); //создание файла excel 

            Console.ReadKey(); //ожидание нажатия любой клавиши пользователем, чтобы консоль не закрывалась автоматически
        }
    }
}
