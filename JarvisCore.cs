using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using System.Net;
using NAudio;
using NAudio.Wave;
using CUETools.Codecs.FLAKE;
using CUETools.Codecs;
using System.Speech.Synthesis;
using System.Diagnostics;
using Shell32;
using System.Runtime.InteropServices;
using System.Timers;
using System.Net.NetworkInformation;
using System.Speech;
using System.Xml;
using System.Xml.XPath;
using System.Threading;
using JDialog;
using System.Runtime.Serialization.Json;
//using Google.GData.Calendar;
//using Google.GData.Extensions;
//using Google.GData.Client;
using System.IO.Ports;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;



// 
// 
namespace JarvisCore
{
    public class JCore
    {

        // Глобальные переменные
        public bool systemWaiting = false;
        public Random randomizer = new Random();
        public bool systemConnected = false;
        public bool systemPowered = false;

        // Запись и воспроизведение звука
        public WaveIn waveIn;
        public WaveFileWriter writer;
        //public static SpeechSynthesizer _synthesizer = new SpeechSynthesizer();

        // Доступ к системным функциям
        public Shell32.Shell shellAccess = new Shell();    //описание тут http://msdn.microsoft.com/en-us/library/windows/desktop/bb774094%28v=vs.85%29.aspx

        // Получение времени с последней активности
        [DllImport("User32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        public struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        //// Аптайм системы
        //int systemUptime = Environment.TickCount;

        // Управление звуком
        public bool upsound = false;
        public bool dwnsound = false;
        public bool onsound = false;
        public bool offsound = false;

        // Имя системы и значения параметров по умолчанию
        public string systemName = "";
        public string masterName = "";
        public string tempDir = "";
        public string systemLanguage = "";
        public string systemPassword = "";
        public string mobilePassword = "";

        public string textCommand = "";
        public string lastCommand = "";
        public string speechString = "";
        public List<string> commandList = new List<string>();
        public string outputFilename = "vctemp.wav";
        public string flacName = "vctemp.flac";
        public int currentFileNumber = 0;

        public DateTime timeNow = new DateTime();

        // Наличие интернет-соединения
        public bool isInternetConnection;
        public bool recognizeRealtime;

        // Флаг проверки времени
        public bool timeChecked;
        public bool useVoiceTimer = true;
        public int startVoiceTimer;
        public int endVoiceTimer;

        // Яндекс.Переводчик
        public string yandexKey = "";
        public string yandexSpeechKey = "";
        public string yandexTranslateKey = "";

        // Google аккаунт
        public string googleUsername = "";
        public string googleID = "";
        public string googleSecret = "";
        public string googlePassword = "";
        public int hoursToCheckEvents = 8;
        public int checkPeriod = 15;
        public string googleVRkey = "";

        public int bluetoothPort = 0;
        public int arduinoPort = 0;
        public static SerialPort _serialPort;

        public string phoneNumber = "";
        public string smsRuKey = "";
        public string bingKey = "";

        public string weatherAppID = "9495e626905d2f427771767ff754db3e";    // надо параметризовать в настройки xml, получается отсюда http://home.openweathermap.org/
        public string cityName = "Krasnodar";                               // надо параметризовать в настройки xml

        // Порог опознавания звука
        public double voicePorog = 0.01;

        ///////////////////////////////////////////////////////////////
        // Блок основных и вспомогательных функций

        // Проверка интернет-соединения
        public bool checkInternetConnection()
        {
            try
            {
                bool InternetConnection = false;
                IPStatus status = IPStatus.Unknown;

                List<string> adressList = new List<string>();
                adressList.Add("google.com");
                adressList.Add("yandex.ru");
                adressList.Add("bing.com");

                foreach (var currentAdress in adressList)
                {

                    if (InternetConnection)
                    { break; }

                    status = new Ping().Send(currentAdress.ToString()).Status;
                    if (status == IPStatus.Success)
                    {
                        InternetConnection = true;
                        isInternetConnection = true;
                    }
                }

                if (!InternetConnection)
                {
                    isInternetConnection = false;

                    int rnd = randomizer.Next(2);
                    if (rnd == 0)
                    {
                        speechString = "нет соединения с интернетом, ";
                    }
                    else if (rnd == 1)
                    {
                        speechString = "проверьте коннект, не могу достучаться до гугла, ";
                    }
                    else if (rnd == 2)
                    {
                        speechString = "простите, что-то с сетью, ";
                    }

                }

                return InternetConnection;
            }
            catch { return false; }
        }

        // Запуск потока - исполнителя команд
        private void parallelExecuteCommand(String commandToExecute)
        {
            executeCommand(commandToExecute);
            commandList.RemoveAt(0);
        }

        // Распознавание через гугл
        public void recognition(String wavFilename, String flacFilename)
        {
            string textCommand = "";

            // Конвертация Wav во Flac
            int sampleRate = Wav2Flac(wavFilename, flacFilename);

            // Распознавание Flac
            if (isInternetConnection)
            {
                textCommand = GoogleSpeechRequest(flacFilename, sampleRate);
                textCommand = ParseGoogleAnswer(textCommand);
            }
            else
            {
                // Тут должен быть движок оффлайн распознавания
                //textCommand = "";
            }

            if (textCommand.Length == 0 && isInternetConnection)
            {
                // попробуем распознать яндексом, или c некоторой вероятностью спросим что-либо
                //textCommand = YandexSpeechRequest(flacFilename, sampleRate);
            }

            if (textCommand.Length > 0)
            {
                commandList.Add(textCommand);
            }
            else
            {
                int rnd = randomizer.Next(10);
                if (rnd == 10)
                {
                    speechString += "не поняла, еще раз пожалуйста, " + masterName + ", ";
                }
                else if (rnd == 11)
                {
                    speechString = "простите, не могли бы вы повторить еще разок?, ";
                }
                else if (rnd == 12)
                {
                    speechString = "плохо расслышала, повторите, ";
                }
                else if (rnd == 13)
                {
                    speechString = "вы что-то хотите мне сказать?, ";
                }
                else if (rnd > 13)
                {
                    // какой-нибудь вопрос
                }

            }

        }

        public bool ProcessData(WaveInEventArgs e)
        {
            bool result = false;
            bool Tr = false;
            double Sum2 = 0;
            int Count = e.BytesRecorded / 2;
            for (int index = 0; index < e.BytesRecorded; index += 2)
            {
                double Tmp = (short)((e.Buffer[index + 1] << 8) | e.Buffer[index + 0]);
                Tmp /= 32768.0;
                Sum2 += Tmp * Tmp;
                if (Tmp > voicePorog)
                    Tr = true;
            }
            Sum2 /= Count;
            if (Tr || Sum2 > voicePorog)
            { result = true; }
            else
            { result = false; }
            return result;
        }

        // Конвертация Wav во Flac
        public static int Wav2Flac(String wavFileName, String flacFileName)
        {
            int sampleRate = 0;

            IAudioSource audioSource = new WAVReader(wavFileName, null);
            AudioBuffer buff = new AudioBuffer(audioSource, 0x10000);

            FlakeWriter flakewriter = new FlakeWriter(flacFileName, audioSource.PCM);
            sampleRate = audioSource.PCM.SampleRate;

            FlakeWriter audioDest = flakewriter;
            while (audioSource.Read(buff, -1) != 0)
            {
                audioDest.Write(buff);
            }

            flakewriter.Close();
            audioDest.Close();
            audioSource.Close();

            return sampleRate;
        }

        // Функция преобразования звука из flac-файла в текст
        public String GoogleSpeechRequest(String flacFileName, int sampleRate)
        {
            var request = (HttpWebRequest)WebRequest.Create("https://www.google.com/speech-api/v2/recognize?output=json&lang=ru-ru&key=" + googleVRkey + "&client=chromium&maxresults=1&pfilter=2");
            //WebRequest request = WebRequest.Create("https://www.google.com/speech-api/v1/recognize?xjerr=1&client=chromium&lang=ru-RU");

            request.Method = "POST";
            request.KeepAlive = true;
            request.SendChunked = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/15.0.874.121 Safari/535.2";
            request.Headers.Set(HttpRequestHeader.AcceptEncoding, "gzip,deflate,sdch");
            //request.Headers.Set(HttpRequestHeader.AcceptLanguage, "en-GB,en-US;q=0.8,en;q=0.6");
            request.Headers.Set(HttpRequestHeader.AcceptCharset, "ISO-8859-1,utf-8;q=0.7,*;q=0.3");

            byte[] byteArray = File.ReadAllBytes(flacFileName);

            // Set the ContentType property of the WebRequest.
            request.ContentType = "audio/x-flac; rate=" + sampleRate; //"16000";        
            request.ContentLength = byteArray.Length;
            request.Timeout = 15000;

            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);

            dataStream.Close();

            try
            {
                string responseFromServer = "";

                using (var response = request.GetResponse())
                {
                    using (var responseStream = response.GetResponseStream())
                    {
                        using (var zippedStream = new GZipStream(responseStream, CompressionMode.Decompress))
                        {
                            using (var sr = new StreamReader(zippedStream))
                            {
                                var res = sr.ReadToEnd();
                                responseFromServer = responseFromServer + res;
                            }
                        }
                    }
                }

                return responseFromServer;

            }
            catch
            { return ""; }

        }

        public String YandexSpeechRequest(String flacFileName, int sampleRate)
        {
            var request = (HttpWebRequest)WebRequest.Create("https://asr.yandex.net/asr_xml?uuid=" + yandexKey + "&key=" + yandexSpeechKey + "&topic=queries&lang=ru");

            request.Method = "POST";
            request.KeepAlive = true;
            request.SendChunked = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/15.0.874.121 Safari/535.2";
            request.Headers.Set(HttpRequestHeader.AcceptEncoding, "gzip,deflate,sdch");
            request.Headers.Set(HttpRequestHeader.AcceptCharset, "ISO-8859-1,utf-8;q=0.7,*;q=0.3");

            byte[] byteArray = File.ReadAllBytes(flacFileName);

            // Set the ContentType property of the WebRequest.
            request.ContentType = "audio/x-wav; rate=" + sampleRate; //"16000";        
            request.ContentLength = byteArray.Length;
            request.Timeout = 15000;

            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            try
            {
                string responseFromServer = "";

                using (var response = request.GetResponse())
                {
                    using (var responseStream = response.GetResponseStream())
                    {
                        //using (var zippedStream = new GZipStream(responseStream, CompressionMode.Decompress))
                        //{
                        //    using (var sr = new StreamReader(zippedStream))
                        //    {
                        //        var res = sr.ReadToEnd();
                        //        responseFromServer = responseFromServer + res;
                        //    }
                        //}

                        var rootElement = XElement.Parse(responseStream.ToString());
                        responseFromServer = rootElement.Element("recognitionResults").Element("variant").ToString();

                    }
                }

                return responseFromServer;
            }
            catch
            { return ""; }

        }

        // Парсинг ответа Гугла
        public String ParseGoogleAnswer(String googleAnswer)
        {
            String newGoogleAnswer = "";
            Single valid;
            //int startIndex = googleAnswer.IndexOf("utterance");
            int startIndex = googleAnswer.IndexOf("transcript");
            int endIndex = googleAnswer.IndexOf("confidence");
            if (startIndex == -1 || endIndex == -1)
            {
                return "";
            }

            newGoogleAnswer = googleAnswer.Substring(startIndex + 13, endIndex - startIndex - 16);

            // Обработка корректности полученного параметра, если ниже 50%, то опускаем
            String GoogleAnswerValid = googleAnswer.Substring(endIndex + 14, 1);

            Single.TryParse(GoogleAnswerValid, System.Globalization.NumberStyles.Number, System.Globalization.NumberFormatInfo.CurrentInfo, out valid);
            if (valid < 2)
            {

                int rnd = randomizer.Next(2);
                if (rnd == 0)
                {
                    speechString += "не поняла, еще раз пожалуйста, " + masterName + ", ";
                }
                else if (rnd == 1)
                {
                    speechString = "простите, не могли бы вы повторить еще разок?, ";
                }
                else if (rnd == 2)
                {
                    speechString = "плохо расслышала, повторите, ";
                }

                return "";
            }
            else if (newGoogleAnswer.Length < 3)
            {
                // считаем шумом, все команды должны начинаться с ключевого слова
                return "";
            }

            return newGoogleAnswer;
        }

        // Проверка на вхождение ключевого слова
        public bool isThereSystemNameInCommand(String commandToAnalyse)
        {
            if (commandToAnalyse == systemName.ToLowerInvariant())
            {
                return true;
            }

            if (commandToAnalyse.Length < (systemName.Length + 1))
            {
                return false;
            }

            if (commandToAnalyse.Substring(0, systemName.Length).ToLowerInvariant() == systemName.ToLowerInvariant())
            {
                commandToAnalyse = commandToAnalyse.Substring(systemName.Length + 1, (commandToAnalyse.Length - (systemName.Length + 1)));
                return true;
            }
            return false;
        }

        // Исключение системного имени из строки
        public String excludeSystemName(String textCommand)
        {
            textCommand = textCommand.Replace(systemName, "");
            return textCommand.Trim();
        }

        public string getStringFromLink(string requestUrl, bool utf = true)
        {
            if (!isInternetConnection)
            {
                return "";
            }

            WebClient webcl = new WebClient();
            if (utf == true)
            {
                webcl.Encoding = System.Text.Encoding.UTF8;
            }
            else
            {
                webcl.Encoding = System.Text.Encoding.GetEncoding(1251);
            }
            webcl.BaseAddress = requestUrl;
            string catchedString = webcl.DownloadString(webcl.BaseAddress);

            return catchedString;
        }

        // Пожелание доброго утра, прогноз погоды
        public void goodMorning()
        {
            speechString += "доброе утро, " + masterName + ", ";
            checkWeather();

            int rnd = randomizer.Next(5);
            if (rnd == 0)
            {
                speechString += ", неплохо бы заняться зарядкой, ";
            }
            else if (rnd == 1)
            {
                speechString += ", хорошего дня, " + masterName;
            }
            else if (rnd == 2)
            {
                speechString += ", предлагаю провести парочку физических уражнений, " + masterName;
            }
            else if (rnd == 3)
            {
                speechString += ", желаю доброго дня, " + masterName;
            }
            //
        }

        // Проверка погоды
        public void checkWeather()
        {
            bool weatcherChecked = false;
            if (checkInternetConnection())
            {
                try
                {
                    String xmlString = "";
                    String requestString = "http://api.openweathermap.org/data/2.5/weather?q=" + cityName + "&APPID=" + weatherAppID + "&mode=xml&units=metric&lang=ru";

                    xmlString = getStringFromLink(requestString, true);

                    var rootElement = XElement.Parse(xmlString);
                    String temperature = rootElement.Element("temperature").FirstAttribute.Value;
                    temperature = temperature.Replace('.', ',');
                    float resulttemperature;
                    bool succConvert = float.TryParse(temperature, out resulttemperature);
                    if (succConvert)
                    {
                        temperature = ((int)resulttemperature).ToString();
                    }
                    speechString += "температура " + temperature + " градусов" + ", ";

                    try
                    {
                        String weatherValue = rootElement.Element("weather").FirstAttribute.NextAttribute.Value;
                        if (weatherValue.Trim().Length != 0)
                        {
                            speechString += "ожидается " + weatherValue + ",";
                        }
                    }
                    catch 
                    { }

                    weatcherChecked = true;

                }
                catch
                {

                }
            }

            if (!weatcherChecked)
            {
                int rnd = randomizer.Next(1);
                if (rnd == 0)
                {
                    speechString += masterName + ", нет соединения с сервером, чтобы проверить погоду. Возможно, вы просто выгляните в окно?, ";
                }
                else if (rnd == 1)
                {
                    speechString = masterName + ", не могу проверить погоду, пропала связь с сервером, ";
                }

                return;
            }

        }

        // Получение времени неактивности системы
        // Метод отсюда http://kbyte.ru/ru/Programming/Sources.aspx?id=945&mode=show //Geekpedia.com 2007
        public static uint GetIdleTime()
        {
            int systemUptime = Environment.TickCount;
            int LastInputTicks = 0;
            int IdleTicks = 0;

            LASTINPUTINFO LastInputInfo = new LASTINPUTINFO();
            LastInputInfo.cbSize = (uint)Marshal.SizeOf(LastInputInfo);
            LastInputInfo.dwTime = 0;

            if (GetLastInputInfo(ref LastInputInfo))
            {
                LastInputTicks = (int)LastInputInfo.dwTime;
                IdleTicks = systemUptime - LastInputTicks;
            }

            return ((uint)Environment.TickCount - LastInputInfo.dwTime);
        }

        // Работа с Wiki
        public String returnWikiInfo(String textRequest)
        {
            String wikiReturned = "";

            if (!checkInternetConnection())
            {
                return wikiReturned;
            }

            try
            {

                String xmlString = "";

                string requesturl = "http://ru.wikipedia.org/w/api.php?action=query&list=search&format=xml&srsearch=TEXTREQUEST&srnamespace=0&srwhat=text&srredirects=&srlimit=10";
                requesturl = requesturl.Replace("TEXTREQUEST", textRequest);

                //WebClient webcl = new WebClient();
                //webcl.Encoding = System.Text.Encoding.UTF8; //GetEncoding(1251);
                //webcl.BaseAddress = requesturl;

                xmlString = getStringFromLink(requesturl); // webcl.DownloadString(webcl.BaseAddress);

                var rootElement = XElement.Parse(xmlString);
                wikiReturned = rootElement.Element("query").Element("search").ToString();

                // Удаляем лишнюю инфу
                wikiReturned = wikiReturned.Replace("<search>", " ");
                wikiReturned = wikiReturned.Replace("</search>", " ");
                wikiReturned = wikiReturned.Replace("&lt;", " ");
                wikiReturned = wikiReturned.Replace("&gt;", " ");
                wikiReturned = wikiReturned.Replace("<p>", " ");
                wikiReturned = wikiReturned.Replace("</p>", " ");
                wikiReturned = wikiReturned.Replace("<li>", " ");
                wikiReturned = wikiReturned.Replace("</li>", " ");
                wikiReturned = wikiReturned.Replace("<b>", " ");
                wikiReturned = wikiReturned.Replace("</b>", " ");
                wikiReturned = wikiReturned.Replace("<span class='searchmatch'>", " ");
                wikiReturned = wikiReturned.Replace("</span>", " ");
                wikiReturned = wikiReturned.Replace("span class='searchmatch'", " ");
                wikiReturned = wikiReturned.Replace("/span", " ");
                wikiReturned = wikiReturned.Replace("b … /b", " ");
                wikiReturned = wikiReturned.Replace("+", " ");
                wikiReturned = wikiReturned.Replace("|", " ");
                wikiReturned = wikiReturned.Replace("/", " ");

                string textReturned = "";
                string[] textArray = wikiReturned.Split('>');
                for (int ii = 0; ii < textArray.Length; ++ii)
                {
                    int startIndex = textArray[ii].IndexOf("snippet=");
                    int endIndex = textArray[ii].IndexOf("size");
                    if (startIndex > -1 && endIndex > -1)
                    {
                        textReturned += textArray[ii].Substring(startIndex + 8, endIndex - (startIndex + 12));
                    }
                }

                System.IO.File.WriteAllText(@"" + tempDir + "readytoread.txt", textReturned);

                return textReturned;

            }
            catch
            {
                //
            }

            return wikiReturned;
        }

        // Поиск в Bing
        public String returnSEInfo(String textRequest)
        {
            string SEInfo = "";

            string query = textRequest;
            string rootUri = "https://api.datamarket.azure.com/Bing/Search";
            var bingContainer = new Bing.BingSearchContainer(new Uri(rootUri));
            var accountKey = bingKey;

            string urlToFile = "";
            try
            {
                bingContainer.Credentials = new NetworkCredential(accountKey, accountKey);
                var webQuery = bingContainer.Web(query, null, null, null, null, null, null, null);

                var results = webQuery.Execute();

                foreach (var result in results)
                {
                    urlToFile += result.Url.ToString();
                    urlToFile = urlToFile.Trim();
                    SEInfo += result.Description.ToString();
                    SEInfo = SEInfo.Trim();

                    if (urlToFile.Length != 0)
                    {
                        break;
                    }
                }
            }
            catch { }

            System.IO.File.WriteAllText(@"" + tempDir + "readytoread.txt", urlToFile);

            return SEInfo;
        }

        public class GenericEvent : IEquatable<GenericEvent>
        {
            public string Title { get; set; }
            public string Contents { get; set; }
            public string Location { get; set; }
            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }

            #region IEquatable<GenericEvent> Members

            public bool Equals(GenericEvent other)
            {
                //Compare all fields to check equality
                if (this.Title == other.Title &&
                    this.Contents == other.Contents &&
                    this.Location == other.Location &&
                    this.StartTime == other.StartTime &&
                    this.EndTime == other.EndTime)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            #endregion
        }

        // Получаем ивенты в ближайшие часы
        public void checkGoogleCalendar(bool forcedAsk = false)
        {
            if (!isInternetConnection)
            {
                return;
            }

            try // V3
            {

                ClientSecrets secrets = new ClientSecrets
                {
                    ClientId = googleID,
                    ClientSecret = googleSecret,
                };


                UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                secrets,
                                new string[]
                                { 
                                        CalendarService.Scope.Calendar
                                },
                                "user",
                                CancellationToken.None)
                        .Result;

                var initializer = new BaseClientService.Initializer();
                initializer.HttpClientInitializer = credential;
                initializer.ApplicationName = "JProject";
                CalendarService calendarConnection;
                calendarConnection = new CalendarService(initializer);

                var calendars = calendarConnection.CalendarList.List().Execute().Items;
                foreach (CalendarListEntry entry in calendars)
                {
                    Events request = null;
                    EventsResource.ListRequest lr = calendarConnection.Events.List(googleUsername);

                    lr.TimeMin = DateTime.Now.AddDays(0);
                    lr.TimeMax = DateTime.Now.AddDays(31);

                    request = lr.Execute();
                    var eventsInRange = new List<Calendar>();
                    if (request != null)
                    {
                        if (request.Items != null)
                        {
                            foreach (var genericEvent in request.Items)
                            {
                                // вычисляем время до каждого ивента, если меньше 8 часов, то сообщаем
                                TimeSpan timeLeft = genericEvent.Start.DateTime.Value.Subtract(DateTime.Now);
                                if (timeLeft.Days == 0 && timeLeft.Hours < hoursToCheckEvents && (timeLeft.Seconds > 0 || timeLeft.Minutes > 0 || timeLeft.Hours > 0))
                                {
                                    speechString += "напоминаю, через " + timeLeft.Hours + "часов и " + timeLeft.Minutes + " минут, запланировано " + genericEvent.Summary + ", ";
                                    break;
                                }
                            }

                            if (forcedAsk && request.Items.Count == 0)
                            {
                                speechString += "в настоящее время ничего не запланировано, ";
                                break;
                            }

                        }
                    }

                }


            }
            catch
            {
                // не удалось достучаться в гугл
            }

        }

        public string getMonthByNumber(int monthNumber, string language = "RU")
        {
            string monthString = "";

            if (language == "RU")
            {
                if (monthNumber == 1)
                {
                    monthString = "января";
                }
                if (monthNumber == 2)
                {
                    monthString = "февраля";
                }
                if (monthNumber == 3)
                {
                    monthString = "марта";
                }
                if (monthNumber == 4)
                {
                    monthString = "апреля";
                }
                if (monthNumber == 5)
                {
                    monthString = "мая";
                }
                if (monthNumber == 6)
                {
                    monthString = "июня";
                }
                if (monthNumber == 7)
                {
                    monthString = "июля";
                }
                if (monthNumber == 8)
                {
                    monthString = "августа";
                }
                if (monthNumber == 9)
                {
                    monthString = "сентября";
                }
                if (monthNumber == 10)
                {
                    monthString = "октября";
                }
                if (monthNumber == 11)
                {
                    monthString = "ноября";
                }
                if (monthNumber == 12)
                {
                    monthString = "декабря";
                }
            }

            return monthString;
        }

        public void checkTime(bool checkMinutes)
        {
            timeNow = DateTime.Now;
            int hour = timeNow.Hour;
            int minute = timeNow.Minute;
            int second = timeNow.Second;

            // тут проверить вхождение часа и минуты в интервалы 
            // час = 1, 21 / часа = 2-4, 22-24 / часов = 5-20 и т.д.
            string shour = "";
            string sminute = "";
            string[] str1 = { "1", "21", "31", "41", "51" };
            string[] str2 = { "2", "3", "4", "22", "23", "24", "32", "33", "34", "42", "43", "44", "52", "53", "54" };

            if (str1.Contains(hour.ToString())) { shour = "час"; }
            else if (str2.Contains(hour.ToString())) { shour = "часа"; }
            else { shour = "часов"; }

            if (str1.Contains(minute.ToString())) { sminute = "минута"; }
            else if (str2.Contains(minute.ToString())) { sminute = "минуты"; }
            else { sminute = "минут"; }


            if (!checkMinutes)
            {
                if (hour > endVoiceTimer && hour < startVoiceTimer)
                {
                    return;
                }
                if (!useVoiceTimer)
                {
                    return;
                }

                if (minute == 0 && second < 2 && !timeChecked)
                {
                    if (hour == startVoiceTimer)
                    {
                        goodMorning();
                        speechString += " сегодня " + timeNow.Day + " " + getMonthByNumber(timeNow.Month, "RU") + ", ";
                        // можно еще апи праздников впилить сюда
                        // например: getStringFromLink(http://htmlweb.ru/service/api.php?holiday&country=RU&d_from=2014-11-14&d_to=2014-11-14&perpage=1)
                        // или так: getStringFromLink(http://kayaposoft.com/enrico/json/v1.0/?action=getPublicHolidaysForMonth&month=1&year=2014&country=rus)
                    }

                    speechString += "системное время " + hour + " " + shour + ", ";
                    timeChecked = true;

                }
                else
                {
                    timeChecked = false;
                }

                int currentCheck = 0;
                List<int> checkList = new List<int>();
                while (currentCheck < 60)
                {
                    checkList.Add(currentCheck);
                    currentCheck += checkPeriod;
                    if (checkList.Count > 60)
                    { break; }
                }

                if ((checkList.IndexOf(minute) != -1) && second < 2)
                {
                    checkGoogleCalendar();
                }

            }
            else
            {
                speechString += "системное время " + hour + " " + shour + " " + minute + " " + sminute + ", ";
            }

        }

        public string getJoke()
        {
            string jokeString = getStringFromLink("http://rzhunemogu.ru/Rand.aspx?CType=1", false);
            return jokeString;
        }

        public string translateText(string txtToTranslate, bool toRus = true)
        {
            string translated = "";
            string lang = "";

            if (toRus)
            { lang = "ru"; }
            else
            { lang = "en"; }

            string requesturl = "https://translate.yandex.net/api/v1.5/tr.json/translate?key=" + yandexTranslateKey + "&text=" + txtToTranslate + "&lang=" + lang + "&format=plain";
            string jsonString = getStringFromLink(requesturl);

            translated = parseJson(jsonString, "text");

            return translated;
        }

        // Десериализация JSON
        public string parseJson(String jsonString, string tagToSearch)
        {
            string strParsed = "";

            try
            {
                string[] strArray = jsonString.Split(',');
                for (int ii = 0; ii < strArray.Length; ++ii)
                {
                    if (strArray[ii].Substring(1, tagToSearch.Length) == tagToSearch)
                    {
                        strParsed += strArray[ii].Substring(tagToSearch.Length + 5, strArray[ii].Length - tagToSearch.Length - 8);
                    }
                }
            }
            catch
            { }

            return strParsed;
        }

        public bool bluetoothConnect(string stringToPut)
        {
            bool operationSuccessfull = false;

            try
            {
                string scom = "COM" + bluetoothPort.ToString();
                _serialPort = new SerialPort(scom, 9600, Parity.None, 8, StopBits.One);
                _serialPort.Open();
                if (stringToPut.Length != 0)
                {
                    _serialPort.ReadTimeout = 5000;
                    _serialPort.WriteTimeout = 5000;
                    _serialPort.Write(stringToPut);
                }
                operationSuccessfull = true;
                _serialPort.Close();
            }
            catch
            { }

            return operationSuccessfull;
        }

        public bool arduinoConnect(string stringToPut)
        {
            bool operationSuccessfull = false;

            try
            {
                string scom = "COM" + arduinoPort.ToString();
                _serialPort = new SerialPort(scom, 9600, Parity.None, 8, StopBits.One);
                _serialPort.Open();
                if (stringToPut.Length != 0)
                {
                    _serialPort.ReadTimeout = 5000;
                    _serialPort.WriteTimeout = 5000;
                    _serialPort.Write(stringToPut);
                }
                operationSuccessfull = true;
                _serialPort.Close();
            }
            catch
            { }

            return operationSuccessfull;
        }

        public bool smsSend(string textMessage)
        {
            bool operationSuccessfull = false;

            string requesturl = "http://sms.ru/sms/send?api_id=" + smsRuKey + "&to=" + phoneNumber + "&text=" + textMessage;

            WebRequest request = WebRequest.Create(requesturl);
            request.Method = "POST";
            request.Timeout = 15000;
            try
            {
                WebResponse response = request.GetResponse();
                response.Close();
                operationSuccessfull = true;
            }
            catch
            {
            }

            return operationSuccessfull;
        }

        public void checkBattery()
        {
            try
            {
                if (SystemInformation.PowerStatus.BatteryChargeStatus == BatteryChargeStatus.NoSystemBattery || SystemInformation.PowerStatus.BatteryChargeStatus == BatteryChargeStatus.Unknown)
                {
                    return;
                }
                // Если было питание, а теперь батарея ниже 99%, то оповещаем
                // Если не было питания, а теперь батарея заряжается, то оповещаем
                // Иначе ничего не делаем
                else if (systemPowered && SystemInformation.PowerStatus.BatteryLifePercent < 1 && SystemInformation.PowerStatus.PowerLineStatus == PowerLineStatus.Offline)
                {
                    systemPowered = false;
                    if (recognizeRealtime)
                    {
                        // Если пользователь рядом, смс не отправляем
                        speechString += "батарея разряжается, ";
                    }
                    else
                    {
                        smsSend("электричество пропало");
                    }

                }
                else if (!systemPowered && SystemInformation.PowerStatus.PowerLineStatus == PowerLineStatus.Online)
                {
                    systemPowered = true;
                    if (recognizeRealtime)
                    {
                        // Если пользователь рядом, смс не отправляем
                        speechString += "батарея заряжается, ";
                    }
                    else
                    {
                        smsSend("электричество появилось");
                    }
                }
                else if (SystemInformation.PowerStatus.BatteryChargeStatus == BatteryChargeStatus.Critical)
                {
                    speechString += "питание батареи на исходе, ";
                }

            }
            catch
            {
            }


        }

        public void checkNetwork()
        {
            isInternetConnection = checkInternetConnection();
            if (!systemConnected && isInternetConnection)
            {
                systemConnected = true;
                if (recognizeRealtime)
                {
                    // Если пользователь рядом, смс не отправляем
                    speechString += "соединение восстановлено, ";
                }
                else
                {
                    smsSend("соединение восстановлено");
                }
            }
            else if (systemConnected && !isInternetConnection)
            {
                systemConnected = false;
                if (recognizeRealtime)
                {
                    // Если пользователь рядом, смс не отправляем
                    speechString += "соединение потеряно, ";
                }
                else
                {
                    //smsSend("соединение потеряно"); // нет смысла, отсылка то по сети идет =)
                }
            }

        }

        // Получаем случайную цитату баша
        public void readBash()
        {
            if (!checkInternetConnection())
            {
                return;
            }

            try
            {
                string query = "http://bash.im/forweb/?u";

                // Начинаем читать отсюда
                // <' + 'div id="b_q_t" style="padding: 1em 0;">
                // и до сюда
                // <' + '/div><' + 'small>
                // Удаляем теги <' + 'br><' + 'br>

                String htmlString = "";
                //WebClient webcl = new WebClient();
                //webcl.Encoding = System.Text.Encoding.UTF8; //GetEncoding(1251);
                //webcl.BaseAddress = query;
                htmlString = getStringFromLink(query); // webcl.DownloadString(webcl.BaseAddress);

                // Удаляем лишнюю инфу
                htmlString = htmlString.Replace("<' + 'br>", " ");
                htmlString = htmlString.Replace("&lt;", " ");
                htmlString = htmlString.Replace("&gt;", " ");
                htmlString = htmlString.Replace("&quot;", " ");
                int startSym = htmlString.IndexOf("padding: 1em 0;");
                int endSym = htmlString.IndexOf("<' + '/div><' + 'small>");
                startSym += 18;
                htmlString = htmlString.Substring(startSym, endSym - startSym);

                System.IO.File.WriteAllText(@"" + tempDir + "bash.txt", htmlString);
                readTxt("bash.txt");

            }
            catch
            {
                //
            }

        }

        // Чтение
        public void readTxt(string filename = "readytoread.txt")
        {
            if (System.IO.File.Exists(tempDir + filename))
            {
                string txtReader = System.IO.File.ReadAllText(tempDir + filename);
                speechString += txtReader + ", ";
            }
            else
            {
                speechString += "не могу прочитать, так как нечего, ";
            }
        }

        // Открытие ссылки
        public void openUrl(string strUrl = "")
        {
            if (strUrl.Length == 0)
            {
                strUrl = System.IO.File.ReadAllText(tempDir + "readytoread.txt");
                strUrl = strUrl.Trim();
            }
            shellAccess.Open(strUrl);
        }


        //////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////

        // Обработка голосовых команд для ответов и запуска программ
        public void executeCommand(String textCommand)
        {
            if (textCommand.Length == 0)
            {
                return;
            }

            textCommand = textCommand.ToLowerInvariant();

            // Проверка на вхождение ключевого слова
            bool isThereSystemName = isThereSystemNameInCommand(textCommand);
            textCommand = excludeSystemName(textCommand);

            /////////////////////////////////////////////
            // Оповещаем что мы готовы к диалогу
            if (!systemWaiting && isThereSystemName)
            {
                speechString += "да " + masterName + "?, ";
                systemWaiting = isThereSystemName;
            }

            lastCommand = textCommand; // запоминаем предыдущую команду

            // Если система не в режиме ожидания и ключевого слова не прозвучало, то возврат
            if (systemWaiting || isThereSystemName)
            {

                var sqlConnection = new System.Data.SqlServerCe.SqlCeConnection();
                //sqlConnection.ConnectionString = "Data Source=\"C:\\Share\\db4vcs2.sdf\"";
                sqlConnection.ConnectionString = "Data Source=\"db4vcs2.sdf\"";
                sqlConnection.Open();

                var sqlCommand = new System.Data.SqlServerCe.SqlCeCommand();
                sqlCommand.Connection = sqlConnection;

                // Тут разобьем команду на слова и попробуем найти каждому сопоставление в БД
                string[] commandArray = textCommand.Split(' ');

                string strAction = "";      // действие
                string strObject = "";      // объект
                string strLink = "";        // принадлежность
                string strType = "";        // тип объекта
                string strPath = "";        // путь к объекту
                string strProcName = "";    // имя процесса объекта
                string strSearch = "";
                bool isThereHello = false;  // проверка наличия приветствия
                //bool isThereQuestion = false;   // проверка наличия вопроса
                int reaction = 0;           // проверка наличия ответа (-1 = нет, 0 - нет ответа, 1 - да)

                for (int ii = 0; ii < commandArray.Length; ++ii)
                {
                    commandArray[ii] = commandArray[ii].Trim();
                    if (commandArray[ii].Length == 0)
                    {
                        continue;
                    }

                    // Запрос в цикле конечно плохо, по свободе обдумаю другие варианты
                    //sqlCommand.Parameters.Add("textCommand", commandArray[ii]);
                    sqlCommand.CommandText = "SELECT * FROM SpeechAnalysys WHERE (Syntax LIKE '%" + commandArray[ii] + "%' ) ORDER BY ID";

                    var sqlReader = sqlCommand.ExecuteReader(System.Data.CommandBehavior.SingleResult);

                    if (sqlReader.Read()) //
                    {
                        // Если нашли подходящую строку, то из изначальной ее удалим
                        textCommand = textCommand.Replace(commandArray[ii], "");

                        if (sqlReader.GetString(4) == "exclude")
                        {
                            continue;
                        }

                        if (sqlReader.GetString(2) == "action")
                        {
                            strAction = sqlReader.GetString(4);
                        }
                        if (sqlReader.GetString(2) == "virtual" || sqlReader.GetString(2) == "material")
                        {
                            // Определимся, что и с кем/чем нужно сделать
                            strType = sqlReader.GetString(2);
                            strLink = sqlReader.GetString(3);
                            strObject = sqlReader.GetString(4);
                            if (!sqlReader.IsDBNull(5))
                            {
                                strPath = sqlReader.GetString(5);
                            }
                            if (!sqlReader.IsDBNull(6))
                            {
                                strProcName = sqlReader.GetString(6);
                            }
                        }
                        // Для диалога
                        if (sqlReader.GetString(2) == "dialog")
                        {
                            if (sqlReader.GetString(4) == "hello")
                            {
                                isThereHello = true;
                            }
                            if (sqlReader.GetString(4) == "ok")
                            {
                                reaction = 1;
                            }
                            else if (sqlReader.GetString(4) == "cancel")
                            {
                                reaction = -1;
                            }
                            if (sqlReader.GetString(4) == "question")
                            {
                                strAction = "question";
                                strObject = textCommand;
                                break;
                            }
                        }
                    }
                    // Исключительная ситуация для поиска 
                    else if (strAction == "search" || strAction == "question" || strAction == "check")
                    {
                        //strObject = commandArray[ii];
                        strSearch = textCommand.Trim();
                        break;
                    }


                    sqlReader.Close();
                }

                sqlConnection.Close();

                if (isThereHello)
                {
                    speechString += "здравствуте, " + masterName + ", ";
                }
                if (reaction != 0)
                {
                    // нет // не знаю // да
                }

                // Действия по умолчанию
                if (strObject == "bashim")
                {
                    strAction = "read";
                }
                // Если действие не задано, попробуем запустить
                else if (strAction.Length == 0 && strType == "virtual" && (strPath.Length != 0 || strObject.Length != 0))
                {
                    strAction = "start";
                }

                /////////////////////////////////////////////
                // Попытка исполнения по собранной информации
                try
                {
                    // Запуск объекта
                    if (strAction == "start")
                    {
                        // Ардуино
                        if (strType == "material" && strLink == "house")
                        {
                            if (strObject == "light")
                            {
                                arduinoConnect("1");
                            }
                            // в этом же стиле 10-ок остальных команд
                            // вопрос в формате приема
                        }
                        // Звук
                        else if (strObject == "sound" || strObject == "music")
                        {
                            onsound = true;
                        }
                        // Папки
                        else if (strObject == "folder" && strType == "virtual" && strPath.Length != 0)
                        {
                            shellAccess.Open(strPath);
                        }
                        // Приложения по пути
                        else if (strObject.Length != 0 && strObject != "folder" & strType == "virtual" && strPath.Length != 0)
                        {
                            Process.Start(strPath);
                        }
                        // Приложения по имени
                        else if (strObject.Length != 0 && strObject != "folder" & strType == "virtual" && strPath.Length == 0)
                        {
                            Process.Start(strObject);
                        }
                        // Приложения по имени
                        else if (strObject.Length != 0 && strType == "material")
                        {
                            // управление физическими устройствами //
                        }
                        // Пуск->Выполнить
                        else if (strObject.Length == 0 && strType.Length == 0)
                        {
                            try
                            {
                                openUrl();
                            }
                            catch
                            {
                                shellAccess.FileRun();
                            }
                        }


                    }

                    // Остановка объекта
                    if (strAction == "stop")
                    {
                        // Ардуино
                        if (strType == "material" && strLink == "house")
                        {
                            if (strObject == "light")
                            {
                                arduinoConnect("0");
                            }
                            // в этом же стиле 10-ок остальных команд
                            // вопрос в формате приема
                        }
                        // Звук
                        else if (strObject == "sound" || strObject == "music")
                        {
                            offsound = true;
                        }
                        // Закрытие процесса
                        else if (strProcName.Length != 0 && strType == "virtual")
                        {
                            Process[] processes = Process.GetProcessesByName(strProcName);
                            if (processes.Length == 0)
                            {
                                processes = Process.GetProcessesByName(strProcName + " *32");
                            }

                            foreach (Process tempProc in processes)
                            {
                                tempProc.CloseMainWindow();
                                tempProc.WaitForExit();
                            }
                        }
                        // Закрытие себя
                        //else if (strObject == "self")
                        //{
                        //    this.Close();
                        //}
                        // если не указано, что остановить, остановим голос и текущее исполнение
                        //else if (strObject.Length == 0)
                        //{
                        //    speechWorker.CancelAsync();
                        //    executionWorker.CancelAsync();
                        //    _synthesizer.SpeakAsyncCancelAll();
                        //}

                    }

                    // Повышение уровня
                    if (strAction == "up" && (strObject.Length == 0 || strObject == "sound" || strObject == "music"))
                    {
                        upsound = true;
                    }
                    // Иначе допишем по мере заполнения БД
                    // Возможно повышение чего-то еще, может быть

                    // Понижение уровня
                    if (strAction == "down" && (strObject.Length == 0 || strObject == "sound" || strObject == "music"))
                    {
                        dwnsound = true;
                    }
                    // Иначе допишем по мере заполнения БД
                    // Возможно понижение чего-то еще, может быть

                    // Поиск
                    if (strAction == "search")
                    {
                        if (strSearch.Length != 0)
                        {
                            String textRequestResult = returnWikiInfo(strSearch);
                            if (textRequestResult.Length == 0)
                            {
                                speechString += "не удалось осуществить поиск, технические проблемы, ";
                            }
                            else
                            {
                                speechString += "поиск в базе википедиа завершен, ";
                            }
                        }

                    }

                    // Вопрос - спросим у гугла или яндекса
                    if (strAction == "question")
                    {
                        if (strSearch.Length != 0)
                        {
                            String textRequestResult = returnSEInfo(strSearch);
                            if (textRequestResult.Length == 0)
                            {
                                speechString += "не знаю, " + masterName + ", ";
                            }
                            else
                            {
                                // Либо, как вариант, дописать тут предварительный показ в окне
                                speechString += "поисковик Bing закончил поиск, вот что найдено: ";
                                speechString += textRequestResult + ", ";

                            }
                        }

                    }

                    // Проверка
                    if (strAction == "check")
                    {
                        if (strObject == "weather")
                        {
                            checkWeather();
                        }
                        else if (strObject == "time")
                        {
                            checkTime(true);
                        }
                        else if (strObject == "browser")
                        {
                            checkNetwork();
                        }
                        else if (strType == "virtual" && strProcName.Length != 0 && strObject.Length != 0)
                        {
                            // Проверка, запущен ли объект
                            Process[] processes = Process.GetProcessesByName(strProcName);
                            if (processes.Length != 0)
                            {
                                processes = Process.GetProcessesByName(strProcName + " *32");
                            }
                            if (processes.Length != 0)
                            {
                                speechString += "Процесс " + strObject + " работает, ";
                            }
                            else
                            {
                                speechString += "Процесс " + strObject + " не работает, ";
                            }

                        }
                        else if (strProcName.Length == 0 && strSearch.Length != 0)
                        {
                            String textRequestResult = returnSEInfo(strSearch);
                            if (textRequestResult.Length == 0)
                            {
                                speechString += "не знаю, " + masterName + ", ";
                            }
                            else
                            {
                                // Либо, как вариант, дописать тут предварительный показ в окне
                                speechString += "поисковик Bing закончил поиск, вот что найдено: ";
                                speechString += textRequestResult + ", ";

                            }
                        }
                        else if (strType == "material")
                        {
                            if (strObject == "battery")
                            {
                                checkBattery();
                                if (systemPowered) { speechString += "батарея заряжается, "; } else { speechString += "батарея разряжается, "; }
                            }
                            else if (strObject == "network")
                            {
                                checkNetwork();
                                if (systemConnected) { speechString += "система подключена к сети, "; } else { speechString += "система не подключена к сети, "; }
                            }
                            else
                            {
                                // Пока не знаю что можно ответить про материальные объекты
                                speechString += "извините, я не знаю, ";
                            }
                        }
                    }

                    // Чтение
                    if (strAction == "read")
                    {
                        if (strObject == "bashim")
                        {
                            readBash();
                        }
                        else
                        {
                            readTxt();
                        }
                    }

                    // Перевод
                    if (strAction == "translate")
                    {
                        // Запрос с помощью API Яндекса
                        if (System.IO.File.Exists(tempDir + "readytoread.txt"))
                        {
                            string txtReader = System.IO.File.ReadAllText(tempDir + "readytoread.txt");
                            System.IO.File.WriteAllText(tempDir + "original.txt", txtReader);
                            string translated = translateText(txtReader);
                            speechString += translated + ", ";
                        }
                        else
                        {
                            speechString += "не могу перевести, не нашла исходный файл, ";
                        }

                    }

                    // Планирование
                    if (strAction == "plan")
                    {
                        // google calendar 3.0 api check
                    }
                    if (strAction == "checkplan")
                    {
                        // google calendar 3.0 api check
                        checkGoogleCalendar(true);
                    }

                    // Звонок
                    if (strAction == "call")
                    {
                        // Пока так
                        if (strPath.Length != 0)
                        { Process.Start(strPath); } // skype

                    }

                    // Свернуть
                    if (strAction == "collapse")
                    {
                        shellAccess.MinimizeAll();
                    }

                    // Развернуть
                    if (strAction == "expand")
                    {
                        shellAccess.UndoMinimizeALL();
                    }

                    // Завершение операции
                    if (strAction == "done")
                    {
                        int rnd = randomizer.Next(6);
                        if (rnd == 0)
                        {
                            speechString += "рада стараться, " + masterName + ", ";
                        }
                        else if (rnd == 1)
                        {
                            speechString += "буду рада как-нибудь повторить, " + masterName + ", ";
                        }
                        else if (rnd == 2)
                        {
                            speechString += "обращайтесь, " + masterName + ", ";
                        }
                        else if (rnd == 3)
                        {
                            speechString += "если что, я всегда на месте, " + masterName + ", ";
                        }
                        else if (rnd == 4)
                        {
                            speechString += "может, что-нибудь еще?, " + masterName + ", ";
                        }
                        else if (rnd == 5)
                        {
                            speechString += "это моя работа, " + masterName + ", ";
                        }
                        else if (rnd == 6)
                        {
                            speechString += "как пожелаете, ";
                        }

                        systemWaiting = false;
                    }

                    // Фото
                    if (strAction == "photo")
                    {
                        // Запуск драйвера видеокамеры
                        // Пока без него
                    }

                    // Перезапуск
                    if (strAction == "reload")
                    {
                        if (waveIn != null)
                        {
                            waveIn.Dispose();
                            writer.Close();
                            waveIn = null;
                            writer = null;
                        }
                        Application.Restart();
                    }

                    // Запомнить объект/команду
                    if (strAction == "savenew")
                    {
                        // Цепочка вопросов для внесения в БД
                        // Пока тут пусто
                    }

                    // Соединение 
                    if (strAction == "connect")
                    {
                        bool connectSuccess = bluetoothConnect("");
                        if (connectSuccess)
                        {
                            speechString += "соединение прошло успешно, ";
                        }
                        else
                        {
                            speechString += "произошла ошибка соединения, ";
                        }
                    }

                    // Соединение 
                    if (strAction == "message")
                    {
                        if (strObject == "android")
                        {
                            bool sendSuccess = smsSend("test");
                            if (sendSuccess)
                            {
                                speechString += "соединение прошло успешно, ";
                            }
                            else
                            {
                                speechString += "не удалось отправить СМС, ";
                            }

                        }
                    }


                    // Если скомандовали "повтор", повторим предыдущую команду
                    if (strAction == "repeat")
                    {
                        if (lastCommand.Length == 0)
                        {
                            speechString += " что еще? ";
                        }
                        else
                        {
                            executeCommand(lastCommand);
                        }
                        return;
                    }

                    // Анекдот
                    if (strAction == "joke")
                    {
                        speechString += getJoke();
                    }

                    // Резерв действий

                    //// По умолчанию - запуск
                    //if (strPath.Length != 0 && strType == "virtual")
                    //{
                    //    Process.Start(strPath);
                    // }
                    //if (strObject.Length != 0 && strType == "virtual")
                    //{
                    //    Process.Start(strObject);
                    //}


                    if (strAction == "start" || strAction == "stop" || strAction == "up" || strAction == "down")
                    {
                        int rndm = randomizer.Next(7);
                        if (rndm == 0)
                        {
                            speechString += "выполняю, " + masterName + ", ";
                        }
                        else if (rndm == 1)
                        {
                            speechString += "уже занимаюсь, " + masterName + ", ";
                        }
                        else if (rndm == 2)
                        {
                            speechString += "да, сейчас, " + masterName + ", ";
                        }
                        else if (rndm == 3)
                        {
                            speechString += "будет сделано, " + masterName + ", ";
                        }
                        else if (rndm == 4)
                        {
                            speechString += "слушаюсь, " + masterName + ", ";
                        }
                        else if (rndm == 5)
                        {
                            speechString += "есть, " + masterName + ", ";
                        }
                        else if (rndm == 6)
                        {
                            speechString += "как пожелаете, " + masterName + ", ";
                        }
                        else if (rndm == 7)
                        {
                            speechString += "как пожелаете, ";
                        }
                    }


                }
                catch
                {
                    int rnd = randomizer.Next(2);
                    if (rnd == 0)
                    {
                        speechString += "прошу прощения " + masterName + ", но я не могу этого сделать, ";
                    }
                    else if (rnd == 1)
                    {
                        speechString += "прошу прощения " + masterName + ", у меня не получилось, ";
                    }
                    else if (rnd == 2)
                    {
                        speechString += masterName + "что-то пошло не так, ";
                    }
                }

            }

            textCommand.Trim();
            if (textCommand.Length < 2) { return; }

            /////////////////////////////////////////////
            // Тут запускается алгоритм разговора

            // Подключение библиотеки разговора, на входе строка - вопрос, на выходе строка - ответ
            JarvisDialog jdialog = new JarvisDialog();
            string returnedTxt = jdialog.returnAnswer(textCommand);
            if (returnedTxt.Length != 0)
            {
                // Описываем переменные AIML тут
                returnedTxt = returnedTxt.Replace("$systemname$", systemName);

                speechString += returnedTxt + ", ";
            }


        }

        // Транслитерация строки из кириллицы, для работы бота, пока отключено
        //private string translite(string textCommand)
        //{
        //    string translited = textCommand;

        //    translited = translited.Replace('а', 'а');
        //    translited = translited.Replace('б', 'b');
        //    translited = translited.Replace('п', 'p');
        //    translited = translited.Replace('р', 'r');
        //    translited = translited.Replace('и', 'i');
        //    translited = translited.Replace('в', 'v');
        //    translited = translited.Replace('е', 'e');
        //    translited = translited.Replace('т', 't');
        //    translited = translited.Replace('к', 'k');
        //    translited = translited.Replace('д', 'd');
        //    translited = translited.Replace('л', 'l');

        //    return translited;
        //}


    }
}
