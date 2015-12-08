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
using JarvisCore;
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

// Пока отключены
//using System.Text.RegularExpressions;


// 
// 
namespace JarvisHome
{
    public partial class Form1 : Form
    {
        JCore jcore = new JCore();

        // Глобальные переменные
        Random randomizer = new Random();
        //public string speechString = "";
        public static SpeechSynthesizer _synthesizer = new SpeechSynthesizer();

        // Перемещение формы
        public bool formMoving;
        public int mousePointX;
        public int mousePointY;

        // Доступ к системным функциям
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        // Получение времени с последней активности
        [DllImport("User32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        internal struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        // Управление звуком
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        // Ограничитель вызовов
        private int numberOfCalls = 0;

        ////////////////
        // Инициализация
        public Form1()
        {
            InitializeComponent();
        }


        ///////////////////
        // Работа с формой

        // Обработчик при загрузки формы
        private void Form1_Load(object sender, EventArgs e)
        {

            // Загрузка параметров - имя системы, путь к временной папке, язык
            string settingsFile = "settings.xml";
            /////////////////////////////////////////////////
            // Загрузка настроек из XML
            bool settingsExist = System.IO.File.Exists(settingsFile);
            if (!settingsExist)
            {
                bool tempExist = System.IO.Directory.Exists("C:/Share/");
                if (tempExist)
                {
                    settingsExist = System.IO.File.Exists("C:/Share/" + settingsFile);
                    settingsFile = "C:/Share/" + settingsFile;
                }
            }
            if (!settingsExist)
            {
                // Создадим временную папку и загрузим настройки по-умолчанию
                try
                {
                    System.IO.Directory.CreateDirectory("C:/Share/");
                }
                catch
                { }

                jcore.speechString += "не найден файл с настройками, используются значения по умолчанию, ";
                jcore.systemName = "джарвис";
                jcore.masterName = "";
                jcore.tempDir = "C:/Share/";
                jcore.systemLanguage = "rus";
                jcore.yandexKey = "";
                jcore.yandexSpeechKey = "";
                jcore.yandexTranslateKey = "";
                jcore.googleUsername = "";
                jcore.googlePassword = "";
                jcore.startVoiceTimer = 8;
                jcore.endVoiceTimer = 0;
                jcore.hoursToCheckEvents = 8;
                jcore.voicePorog = 0.01;
                jcore.useVoiceTimer = true;
                jcore.systemPassword = "";
                jcore.mobilePassword = "";
                jcore.checkPeriod = 15;
                jcore.googleVRkey = "";

                try
                {
                    _synthesizer.SelectVoice("ScanSoft Katerina_Full_22kHz");
                }
                catch
                {
                    jcore.speechString += jcore.systemName + "selected voice not found, install properly voice, ";
                    //
                }
            }
            else
            {
                // Подгрузка переменных из XML
                var rootelement = XElement.Load(@"" + settingsFile);
                jcore.systemName = rootelement.Element("SystemName").Value.ToString().ToLower();
                jcore.masterName = rootelement.Element("MasterName").Value.ToString();
                jcore.systemLanguage = rootelement.Element("Language").Value.ToString(); // пока не используется
                jcore.tempDir = rootelement.Element("TempDirectory").Value.ToString();
                jcore.yandexKey = rootelement.Element("YandexKey").Value.ToString(); // ключ яндекса
                jcore.yandexSpeechKey = rootelement.Element("YandexSpeechKey").Value.ToString(); // ключ яндекса для speech api
                jcore.yandexTranslateKey = rootelement.Element("YandexTranslateKey").Value.ToString(); // ключ яндекс - переводчика
                jcore.googleUsername = rootelement.Element("GoogleUsername").Value.ToString();
                jcore.googlePassword = rootelement.Element("GooglePassword").Value.ToString();
                jcore.googleID = rootelement.Element("GoogleID").Value.ToString();
                jcore.googleSecret = rootelement.Element("GoogleSecret").Value.ToString();
                int.TryParse(rootelement.Element("StartVoiceTimer").Value, out jcore.startVoiceTimer); // с какого часа начать оповещать
                int.TryParse(rootelement.Element("EndVoiceTimer").Value, out jcore.endVoiceTimer); // с какого часа прекратить оповещать
                int.TryParse(rootelement.Element("HoursToCheckEvents").Value, out jcore.hoursToCheckEvents); // за сколько часов оповещать о событии
                double.TryParse(rootelement.Element("VoiceSensivity").Value, out jcore.voicePorog); // чувствительность голоса               
                bool.TryParse(rootelement.Element("UseVoiceTimer").Value, out jcore.useVoiceTimer); // использование голосового оповещения
                jcore.systemPassword = rootelement.Element("SystemPassword").Value.ToString();
                jcore.mobilePassword = rootelement.Element("MobilePassword").Value.ToString();
                int.TryParse(rootelement.Element("CheckPeriod").Value, out jcore.checkPeriod);
                int.TryParse(rootelement.Element("BlueToothPort").Value, out jcore.bluetoothPort);
                int.TryParse(rootelement.Element("ArduinoPort").Value, out jcore.arduinoPort);
                jcore.phoneNumber = rootelement.Element("PhoneNumber").Value.ToString();
                jcore.smsRuKey = rootelement.Element("SMSRUKey").Value.ToString();
                jcore.bingKey = rootelement.Element("BingKey").Value.ToString();
                jcore.googleVRkey = rootelement.Element("GoogleVRKey").Value.ToString();

                try
                {
                    //_synthesizer.SelectVoice("ScanSoft Katerina_Full_22kHz");
                    _synthesizer.SelectVoice(rootelement.Element("Voice").Value);
                }
                catch
                {
                    jcore.speechString += jcore.systemName + "selected voice not found, install properly voice, ";
                    //
                }

                // тут дописать остальные параметры при изменении файла настроек
                //
            }

            /////////////////////////////////////////////////
            // Инициализация элементов формы

            executionWorker.DoWork += new DoWorkEventHandler(executionWorker_DoWork);
            recognitionWorker.DoWork += new DoWorkEventHandler(recognitionWorker_DoWork);
            speechWorker.DoWork += new DoWorkEventHandler(speechWorker_DoWork);

            notifyIcon1.Icon = SystemIcons.Shield;
            notifyIcon1.Text = "Voice control system";
            notifyIcon1.Visible = false;
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;

            // Проверка соединения с интернетом, если нет, то использование оффлайн движка
            jcore.isInternetConnection = jcore.checkInternetConnection();
            jcore.recognizeRealtime = true;


            ////////////////////////////////////
            // Идентификация
            // Пока не реализована

            int rnd = randomizer.Next(3);
            if (rnd == 0)
            { jcore.speechString += jcore.systemName + " слушает, "; }
            else if (rnd == 1)
            { jcore.speechString += jcore.systemName + " тут, "; }
            else if (rnd == 2)
            { jcore.speechString += jcore.systemName + " на связи, "; }
            else if (rnd == 3)
            { jcore.speechString += " я перезагрузилась, " + jcore.masterName + ", "; }

            // Сначала определимся, что мы подключены,
            // затем проверим и сообщим, ежели что
            jcore.systemConnected = true;
            jcore.systemPowered = true;

            //microphoneTimer.Enabled = true;
            //executeTimer.Enabled = true;
            //SpeechTimer.Enabled = true;
            //idleTimer.Enabled = true;

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //myBot.saveToBinaryFile("brain.bin");
            //myBot.writeToLog("session data saved.");
            //myUser.Predicates.DictionaryAsXML.Save("UserSession.bin");
        }

        // Вернуть форму из трея
        private void notifyIcon1_MouseDoubleClick(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            this.Show();
        }

        // Выход из подменю в трее
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Dispose();
            Application.Exit();
        }

        private void стопЗвукToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _synthesizer.SpeakAsyncCancelAll();
        }

        // Свернуть в трей
        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            notifyIcon1.Visible = true;
            this.Hide();
        }

        // Свернуть в трей
        private void backImage_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = true;
            this.Hide();
        }

        // Перемещение формы
        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) { formMoving = false; }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                formMoving = true;
                mousePointX = e.X;
                mousePointY = e.Y;
            }
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (formMoving)
            {
                var point = new System.Drawing.Point();
                point.X = this.Location.X + (e.X - mousePointX);
                point.Y = this.Location.Y + (e.Y - mousePointY); ;
                this.Location = point;
            }
        }


        //////////////////////////////////////
        // Запуск таймеров

        // Таймер исполнения
        private void executeTimer_Tick(object sender, EventArgs e)
        {
            if (executionWorker.IsBusy)
            {
                return;
            }
            if (jcore.commandList != null && jcore.commandList.Count > 0)
            {
                //executionWorker.DoWork += new DoWorkEventHandler(executionWorker_DoWork);
                executionWorker.RunWorkerAsync();
            }

            // Сюда перенесены элементы, требующие для исполнения тот же поток, в котором создана форма
            if (jcore.upsound)
            {
                for (int ii = 1; ii < 11; ++ii)
                {
                    SendMessageW(Handle, WM_APPCOMMAND, this.Handle, (IntPtr)APPCOMMAND_VOLUME_UP);
                }
                jcore.upsound = false;
            }
            else if (jcore.dwnsound)
            {
                for (int ii = 1; ii < 11; ++ii)
                {
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)APPCOMMAND_VOLUME_DOWN);
                }
                jcore.dwnsound = false;
            }
            else if (jcore.onsound)
            {
                SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)APPCOMMAND_VOLUME_MUTE);
                jcore.onsound = false;
            }
            else if (jcore.offsound)
            {
                SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)APPCOMMAND_VOLUME_MUTE);
                jcore.offsound = false;
            }

        }

        // Таймер голоса
        private void SpeechTimer_Tick(object sender, EventArgs e)
        {
            if (speechWorker.IsBusy)
            {
                return;
            }
            speechWorker.RunWorkerAsync();
        }

        // Таймер событий
        private void addIventTimer_Tick(object sender, EventArgs e)
        {
            checkTime(false);

            // Проверка соединения и питания каждую минуту
            jcore.timeNow = DateTime.Now;
            int second = jcore.timeNow.Second;
            if (second == 0)
            {
                jcore.checkNetwork();
                jcore.checkBattery();
            }

        }

        // Таймер неактивности
        private void idleTimer_Tick(object sender, EventArgs e)
        {
            uint idleMs = GetIdleTime();
            uint idleMin = (idleMs / 1000 / 60);
            if (idleMin > 15)
            {
                jcore.recognizeRealtime = false;
            }
            else
            {
                jcore.recognizeRealtime = true;
            }
        }


        // Запуск воркеров в параллельных потоках

        // Воркер распознавания
        void recognitionWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!jcore.recognizeRealtime)
            { return; }

            String wavFilename = jcore.tempDir + jcore.currentFileNumber + jcore.outputFilename;
            String flacFilename = jcore.tempDir + jcore.currentFileNumber + jcore.flacName;

            jcore.recognition(wavFilename, flacFilename);
            //throw new NotImplementedException();
        }

        // Воркер исполнения
        void executionWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (jcore.commandList != null && jcore.commandList.Count > 0)
            {
                parallelExecuteCommand(jcore.commandList[0]);
            }
            //throw new NotImplementedException();
        }

        // Воркер голоса
        private void speechWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            PCSpeak();
            //jcore.speechString = "";
        }

        // Воспроизведение звука
        public void PCSpeak()
        {
            if (jcore.speechString == "")
            {
                return;
            }
            _synthesizer.Rate = 9;
            _synthesizer.SpeakAsync(jcore.speechString);
            jcore.speechString = "";
        }

        ///////////////////////////////////////////////////////////////
        // Блок основных и вспомогательных функций

        // Запуск потока - исполнителя команд
        private void parallelExecuteCommand(String commandToExecute)
        {
            jcore.executeCommand(commandToExecute);
            jcore.commandList.RemoveAt(0);
        }

        // Поиск в буфере
        void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {

            bool isThereVoice = false;
            isThereVoice = jcore.ProcessData(e);

            try
            {

                if (isThereVoice)
                {
                    numberOfCalls += 1;
                    jcore.writer.Write(e.Buffer, 0, e.BytesRecorded);
                }
                else
                {
                    if (numberOfCalls > 1)
                    {
                        stopRecord_Click(sender, e);
                    }
                    else
                    {
                    }
                }
            }
            catch { }

        }

        // Обработчик остановки записи
        void waveIn_RecordingStopped(object sender, EventArgs e)
        {
            if (jcore.waveIn == null) { return; }

            jcore.waveIn.Dispose();
            jcore.writer.Close();
            // Закрытие методов
            jcore.waveIn = null;
            jcore.writer = null;

            if (numberOfCalls > 0)
            {

                if (!recognitionWorker.IsBusy)
                {
                    recognitionWorker.RunWorkerAsync();
                }
                else
                {
                    jcore.speechString += "обрабатываю предыдущую команду" + ", ";
                }
                numberOfCalls = 0;
            }

            // Удаление временных файлов           
            //File.Delete(outputFilename);
            //File.Delete(flacName);


            //currentFileNumber = currentFileNumber + 1;
        }

        // Обработчик начала записи
        private void beginRecord_Click(object sender, EventArgs e)
        {
            if (!jcore.recognizeRealtime)
            {
                this.backImage.Image = JarvisHome.Properties.Resources.bw_iron_man_jarvis_200x300;
                return;
            }

            // Пока не придумаю как обходить несколько файлов в фоне
            if (recognitionWorker.IsBusy)
            {
                return;
            }

            if (jcore.waveIn == null)
            {
                this.backImage.Image = JarvisHome.Properties.Resources.iron_man_jarvis_200x300;
                String Filename = jcore.tempDir + jcore.currentFileNumber + jcore.outputFilename;

                jcore.waveIn = new WaveIn();
                jcore.waveIn.DeviceNumber = 0;
                jcore.waveIn.DataAvailable += waveIn_DataAvailable;
                jcore.waveIn.RecordingStopped += new EventHandler<NAudio.Wave.StoppedEventArgs>(waveIn_RecordingStopped);
                jcore.waveIn.WaveFormat = new WaveFormat(16000, 2); //44100        
                jcore.writer = new WaveFileWriter(Filename, jcore.waveIn.WaveFormat);
                jcore.waveIn.StartRecording();
            }
        }

        // Обработчик окончания записи
        private void stopRecord_Click(object sender, EventArgs e)
        {
            try
            {
                this.backImage.Image = JarvisHome.Properties.Resources.BR_iron_man_jarvis_200x300;

                jcore.waveIn.StopRecording();
            }
            catch
            {
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

        private void checkTime(bool checkMinutes)
        {
            jcore.timeNow = DateTime.Now;
            int hour = jcore.timeNow.Hour;
            int minute = jcore.timeNow.Minute;
            int second = jcore.timeNow.Second;

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
                if (hour > jcore.endVoiceTimer && hour < jcore.startVoiceTimer)
                {
                    return;
                }
                if (!jcore.useVoiceTimer)
                {
                    return;
                }

                if (minute == 0 && second < 2 && !jcore.timeChecked)
                {
                    if (hour == jcore.startVoiceTimer)
                    {
                        jcore.goodMorning();
                        jcore.speechString += " сегодня " + jcore.timeNow.Day + " " + jcore.getMonthByNumber(jcore.timeNow.Month, "RU") + ", ";
                        // можно еще апи праздников впилить сюда
                        // например: getStringFromLink(http://htmlweb.ru/service/api.php?holiday&country=RU&d_from=2014-11-14&d_to=2014-11-14&perpage=1)
                        // или так: getStringFromLink(http://kayaposoft.com/enrico/json/v1.0/?action=getPublicHolidaysForMonth&month=1&year=2014&country=rus)
                    }

                    jcore.speechString += "системное время " + hour + " " + shour + ", ";
                    jcore.timeChecked = true;

                }
                else
                {
                    jcore.timeChecked = false;
                }

                int currentCheck = 0;
                List<int> checkList = new List<int>();
                while (currentCheck < 60)
                {
                    checkList.Add(currentCheck);
                    currentCheck += jcore.checkPeriod;
                    if (checkList.Count > 60)
                    { break; }
                }

                if ((checkList.IndexOf(minute) != -1) && second < 2)
                {
                    jcore.checkGoogleCalendar();
                }

            }
            else
            {
                jcore.speechString += "системное время " + hour + " " + shour + " " + minute + " " + sminute + ", ";
            }

        }


    }
}
