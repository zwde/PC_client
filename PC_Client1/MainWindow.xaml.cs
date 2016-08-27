using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace PC_Client1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    class Data_LED : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private string data;
         
        public string Data
        {
            get { return data; }
            set
            {
                data = value;
                //触发事件
                if (PropertyChanged != null)
                {
                    PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Data"));
                }
            }
        }
    }
    public partial class MainWindow : Window
    {
       /* class ex_dll
        {
             [DllImport("extern_dll.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern double volcal(int bytel, int bytem, int byteh);
        }*/





        private Dictionary<int, string> mydic0 = new Dictionary<int, string>(){
            {25,"25mA"},
            {50,"50mA"},
            {0,"Off"}
            };
        
        private string CommNum;//串口号
        private int IntBdr;//波特率
        private int K=0;
        private int Q = 0;
        private double xK1, xK2, xK3, xK4, xC1, xC2, xC3, xC4;
        public MainWindow()
        {
            InitializeComponent();
            FileStream fs = new FileStream("xishu.ini", FileMode.Open);
            StreamReader sr = new StreamReader(fs);
            string str = sr.ReadLine();
            string[] strs = str.Split(' ');
            xK1 = double.Parse(strs[0]);
            xC1 = double.Parse(strs[1]);
            xK2 = double.Parse(strs[2]);
            xC2 = double.Parse(strs[3]);
            xK3 = double.Parse(strs[4]);
            xC3 = double.Parse(strs[5]);
            xK4 = double.Parse(strs[6]);
            xC4 = double.Parse(strs[7]);
            
            sr.Close();
            fs.Close();
           

                Grid_ADC_Capture_Analysis.Visibility = (Visibility)2;

            Grid_Global_Settings.Visibility = (Visibility)0;
            Grid_Tx_stage.Visibility = (Visibility)2;
            Grid_Rx_Stage.Visibility = (Visibility)2;
            Grid_Timing_Controls.Visibility = (Visibility)2;
      


           ComBox_DAC.ItemsSource = mydic0;
            ComBox_DAC.SelectedValuePath = "Key";
            ComBox_DAC.DisplayMemberPath = "Value";
            ComBox_DAC.SelectedIndex = 0;
        
        }
        private void btn_Device_Configuration_Click(object sender, RoutedEventArgs e)
        {
            Grid_Global_Settings.Visibility = (Visibility)0;
            Grid_Tx_stage.Visibility = (Visibility)2;
            Grid_Rx_Stage.Visibility = (Visibility)2;
            Grid_Timing_Controls.Visibility = (Visibility)2;
            Grid_Device_Configuration.Visibility = (Visibility)0;
            Grid_ADC_Capture_Analysis.Visibility = (Visibility)2;
            Grid_Setting.Visibility = (Visibility)2;
        }

        private void btn_ADC_Capture_Analysis_Click(object sender, RoutedEventArgs e)
        {
            Grid_Device_Configuration.Visibility = (Visibility)2;
            Grid_ADC_Capture_Analysis.Visibility = (Visibility)0;
            Grid_Setting.Visibility = (Visibility)2;
        }
       private Dictionary<string, string> dic = new Dictionary<string, string> { 
            {"ppm","ppm"},
            {"ppb","ppb"},
            {"ug/L","ug/L"},
            {"mg/L","mg/L"},
            {"g/L","g/L"},
            {"%","%"}
            };
        private void Setting_Click(object sender, RoutedEventArgs e)
        {
            
            Grid_Device_Configuration.Visibility = (Visibility)2;
            Grid_ADC_Capture_Analysis.Visibility = (Visibility)2;
            Grid_Sample_Test.Visibility = (Visibility)2;
            Grid_Setting.Visibility = (Visibility)0;
            Grid_Data_setting.Visibility = (Visibility)0;
            Grid_Bluetooth.Visibility = (Visibility)2;
            FileStream fs = new FileStream("Yuansu.ini", FileMode.Open);
            StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("gb2312"));
            string[] data = sr.ReadLine().Split(' ');
            Dictionary<string, string> Dic = new Dictionary<string, string>
            {

            };
            for(int i=0;i<data.Length;i++){
                Dic.Add(data[i], data[i]);
            }

           ComBox_yuansu.ItemsSource = Dic;
            ComBox_yuansu.SelectedValuePath = "Key";
            ComBox_yuansu.DisplayMemberPath = "Value";
            ComBox_yuansu.SelectedIndex = 0;
            sr.Close();
            fs.Close();

           
            ComBox_danwei.ItemsSource = dic;
            ComBox_danwei.SelectedValuePath = "Key";
            ComBox_danwei.DisplayMemberPath = "Value";
            ComBox_danwei.SelectedIndex = 0;

        }

        private void Btn_Global_Settings_Click(object sender, RoutedEventArgs e)
        {
            Grid_Global_Settings.Visibility = (Visibility)0;
            Grid_Tx_stage.Visibility = (Visibility)2;
            Grid_Rx_Stage.Visibility = (Visibility)2;
            Grid_Timing_Controls.Visibility = (Visibility)2;
        }

        private void Btn_Tx_stage_Click(object sender, RoutedEventArgs e)
        {
            Grid_Global_Settings.Visibility = (Visibility)2;
            Grid_Tx_stage.Visibility = (Visibility)0;
            Grid_Rx_Stage.Visibility = (Visibility)2;
            Grid_Timing_Controls.Visibility = (Visibility)2;
        }

        private void Btn_Rx_Stage_Click(object sender, RoutedEventArgs e)
        {
            Grid_Global_Settings.Visibility = (Visibility)2;
            Grid_Tx_stage.Visibility = (Visibility)2;
            Grid_Rx_Stage.Visibility = (Visibility)0;
            Grid_Timing_Controls.Visibility = (Visibility)2;
        }

        private void Btn_Timing_Controls_Click(object sender, RoutedEventArgs e)
        {
            Grid_Global_Settings.Visibility = (Visibility)2;
            Grid_Tx_stage.Visibility = (Visibility)2;
            Grid_Rx_Stage.Visibility = (Visibility)2;
            Grid_Timing_Controls.Visibility = (Visibility)0;
        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

            Grid_Enable_Separate_Gain_Mode_isChecked.Visibility = (Visibility)0;
            Grid_Enable_Separate_Gain_Mode_unChecked.Visibility = (Visibility)2;

        }
        private void CheckBox_unChecked(object sender, RoutedEventArgs e)
        {
            Grid_Enable_Separate_Gain_Mode_isChecked.Visibility = (Visibility)2;
            Grid_Enable_Separate_Gain_Mode_unChecked.Visibility = (Visibility)0;

        }
        private void CheckBox_Common_Checked(object sender, RoutedEventArgs e)
        {
            Common.IsEnabled = true;
            panel_Ambient_DAC.IsEnabled = true;
        }
        private void CheckBox_Common_unChecked(object sender, RoutedEventArgs e)
        {
            Common.IsEnabled = false;
            panel_Ambient_DAC.IsEnabled = false;
        }
        private void CheckBox_LED2_LED3_Checked(object sender, RoutedEventArgs e)
        {
            panel_LED2_LED3.IsEnabled = true;
            panel_Ambient_DAC.IsEnabled = true;
        }
        private void CheckBox_LED2_LED3_unChecked(object sender, RoutedEventArgs e)
        {
            panel_LED2_LED3.IsEnabled = false;
            if (CheckBox_LED1_STG2_EN.IsChecked == false)
            {

                panel_Ambient_DAC.IsEnabled = false;
            }

        }
        private void CheckBox_LED1_STG2_EN_Checked(object sender, RoutedEventArgs e)
        {
            panel_LED1_STG2_EN.IsEnabled = true;
            panel_Ambient_DAC.IsEnabled = true;
        }
        private void CheckBox_LED1_STG2_EN_unChecked(object sender, RoutedEventArgs e)
        {
            panel_LED1_STG2_EN.IsEnabled = false;
            ;
            if (CheckBox_LED2_LED3.IsChecked == false)
            {
                panel_Ambient_DAC.IsEnabled = false;
            }
        }

        private void Btn_Diagnostic(object sender, RoutedEventArgs e)
        {
            Byte[] bt = new Byte[10];
            bt[0] = 0x02; 
            bt[1]=bt[2]=bt[3]=bt[4]=bt[5]=bt[6]=bt[7]=0x30;
            bt[9]=0x0D;
            switch (Q)
            {
                case 0:
                    {
                        bt[8] = 0x34;
                        Q = 1;
                        break;
                    }
                case 1:
                    {
                        bt[8] = 0x34;
                        break;
                    }
                case 2:
                    {
                        bt[8] = 0x36;
                        Q = 3;
                        break;
                    }
                case 3:
                    {
                        bt[8] = 0x36;
                        break;
                    }
            }
            try
            {
               SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.Open();
                port.Write(bt, 0, 10);
                Thread.Sleep(30);
                check(port);
                port.Close();
                MessageBox.Show("success");
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message.ToString());
              
            }
            

        }

        private void check(SerialPort port){
                    Byte[] Fir_Ling = new Byte[4];
                    Fir_Ling[0] = 0x03; Fir_Ling[1] = 0x33;
                    Fir_Ling[2] = 0x30; Fir_Ling[3] = 0x0D;
                    port.Write(Fir_Ling, 0, 4); Thread.Sleep(30);
                    Byte[] bt_receive = new Byte[7];
                    port.Read(bt_receive, 0, 7);
                    int flag;
                    flag = bt_receive[2] & 0x01;
                    if (flag == 0)
                    {
                        Rect_PSC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else {
                        Rect_PSC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x02;
                    if (flag == 0)
                    {
                        Rect_SC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_SC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x04;
                    if (flag == 0)
                    {
                        Rect_PPG.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_PPG.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x08;
                    if (flag == 0)
                    {
                        Rect_PNG.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_PNG.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x10;
                    if (flag == 0)
                    {
                        Rect_PD_SC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_PD_SC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x20;
                    if (flag == 0)
                    {
                        Rect_PD_OC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_PD_OC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x40;
                    if (flag == 0)
                    {
                        Rect_LED_OPG.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED_OPG.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[2] & 0x80;
                    if (flag == 0)
                    {
                        Rect_LED_ONG.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED_ONG.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[3] & 0x01;
                    if (flag == 0)
                    {
                        Rect_LED_SC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED_SC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[3] & 0x02;
                    if (flag == 0)
                    {
                        Rect_LED1_OC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED1_OC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[3] & 0x04;
                    if (flag == 0)
                    {
                        Rect_LED2_OC.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED2_OC.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[3] & 0x08;
                    if (flag == 0)
                    {
                        Rect_LED_AS.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_LED_AS.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    flag = bt_receive[3] & 0x10;
                    if (flag == 0)
                    {
                        Rect_PD_AS.Fill = new SolidColorBrush(Colors.LawnGreen);
                    }
                    else
                    {
                        Rect_PD_AS.Fill = new SolidColorBrush(Colors.Gold);
                    }
                    
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.ReadBufferSize = 10000;
                port.Open();
                Byte[] bt1 = new Byte[10], bt2 = new Byte[10];
                bt1[0] = bt2[0] = 0x02;//命令标志
                bt1[9] = bt2[9] = 0x0D;//结束标志位
                bt1[1] = 0x32; bt1[2] = 0x30;//一号数组地址位
                bt2[1] = 0x32; bt2[2] = 0x31;//二号数组地址位
                int[,] t = new int[2, 10];
                for (int i = 0; i < 2; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        t[i, j] = 0;
                    }
                }
                if (ComBox_filter.SelectedIndex == 1)
                {
                    t[1, 5] += 8;
                }
                if (CheckBox_Enable_Separate_Gain_Mode.IsChecked == false)
                {
                    if (CheckBox_Common_STG2_EN.IsChecked == true)
                    {
                        t[1, 5] += 4;
                    }
                    t[1, 4] += ComBox_Ambient_DAC.SelectedIndex;
                    t[1, 6] += ComBox_Common_STG2_GAI.SelectedIndex;
                    t[1, 8] += ComBox_CommonR.SelectedIndex;
                    t[1, 7] += ComBox_CommonC.SelectedIndex / 2;
                    if ((ComBox_CommonC.SelectedIndex % 2) == 1)
                    {
                        t[1, 8] += 8;
                    }
                }
                else
                {
                    t[0, 5] += 8;
                    t[1, 8] += ComBox_LED_RF.SelectedIndex;
                    t[1, 7] += ComBox_LED_CF.SelectedIndex / 2;
                    if ((ComBox_LED_CF.SelectedIndex % 2) == 1)
                    {
                        t[1, 8] += 8;
                    }
                    t[0, 8] += ComBox_LED1_RF.SelectedIndex;
                    t[0, 7] += ComBox_LED1_CF.SelectedIndex / 2;
                    if ((ComBox_LED1_CF.SelectedIndex % 2) == 1)
                    {
                        t[0, 8] += 8;
                    }
                    if (CheckBox_LED1_STG2_EN.IsChecked == true)
                    {
                        t[0, 5] += 4;
                    }
                    if (CheckBox_LED2_LED3.IsChecked == true)
                    {
                        t[1, 5] += 4;
                    }
                    t[1, 6] += ComBox_LED2_LED3.SelectedIndex;
                    t[0, 6] += ComBox_LED1_STG2_GAIN.SelectedIndex;
                    t[1, 4] += ComBox_Ambient_DAC.SelectedIndex;
                }

                for (int j = 3; j <= 8; j++)
                {
                    bt1[j] = change(t[0, j]);
                    bt2[j] = change(t[1, j]);
                }
              
                    port.Write(bt1, 0, 10);
                    Thread.Sleep(50);
                    port.Write(bt2, 0, 10);
                    Thread.Sleep(50);
                    /*检测
                      for (int i = 0; i < 10; i++)
                      {
                          str1 += bt1[i].ToString() + " ";
                          str2 += bt2[i].ToString() + " ";
                      }*/
                    //MessageBox.Show(str1);
                    //MessageBox.Show(str2);
                    port.Close();
                    MessageBox.Show("success");
                    
            }
            catch
            {
                MessageBox.Show("error");

            }
        }
        private Byte change(int s)
        {
            if (s == 0)
            {
                return 0x30;
            }
            else if (s == 1)
            {
                return 0x31;
            }
            else if (s == 2)
            {
                return 0x32;
            }
            else if (s == 3)
            {
                return 0x33;
            }
            else if (s == 4)
            {
                return 0x34;
            }
            else if (s == 5)
            {
                return 0x35;
            }
            else if (s == 6)
            {
                return 0x36;
            }
            else if (s == 7)
            {
                return 0x37;
            }
            else if (s == 8)
            {
                return 0x38;
            }
            else if (s == 9)
            {
                return 0x39;
            }
            else if (s == 10)
            {
                return 0x41;
            }
            else if (s == 11)
            {
                return 0x42;
            }
            else if (s == 12)
            {
                return 0x43;
            }
            else if (s == 13)
            {
                return 0x44;
            }
            else if (s == 14)
            {
                return 0x45;
            }
            else if (s == 15)
            {
                return 0x46;
            }
            return 0;
        }
        private void Tx3_Mode_Click(object sender, RoutedEventArgs e)
        {
            ImageBrush imagebrush = new ImageBrush();
            if (K== 0)
            {
                imagebrush.ImageSource = new BitmapImage(new Uri("image/btn2.png", UriKind.Relative));
                K = 1;
                if (ComBox_Bridge.SelectedIndex == 0)
                {
                    image_Active1.Visibility = (Visibility)0;
                    image_Off1.Visibility = (Visibility)0;
                }
                else
                {
                    image_Active2.Visibility = (Visibility)0;
                    image_Off2.Visibility = (Visibility)0;
                }
            }
            else
            {
                imagebrush.ImageSource = new BitmapImage(new Uri("image/btn1.png", UriKind.Relative));
                K = 0;
                if (ComBox_Bridge.SelectedIndex == 0)
                {
                    image_Active1.Visibility = (Visibility)2;
                    image_Off1.Visibility = (Visibility)2;
                }
                else
                {
                    image_Active2.Visibility = (Visibility)2;
                    image_Off2.Visibility = (Visibility)2;
                }
            }
            btn_Tx3_Mode.Background = imagebrush;
        }

        private void ComBox_Bridge_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComBox_Bridge.SelectedIndex == 0)
            {
                Whole_image.Source = new BitmapImage(new Uri(@"image/Tx.jpg", UriKind.Relative));
                image_Active2.Visibility = (Visibility)2;
                image_Off2.Visibility = (Visibility)2;
            }
            else
            {
                Whole_image.Source = new BitmapImage(new Uri(@"image/Tx1.jpg", UriKind.Relative));
                image_Active1.Visibility = (Visibility)2;
                image_Off1.Visibility = (Visibility)2;
            }
        }

        private void ComBox_Voltage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            /*Dictionary<int, string> mydic0 = new Dictionary<int, string>(){
            {25,"25mA"},
            {50,"50mA"},
            {0,"Off"}
            };*/
            Dictionary<int, string> mydic1 = new Dictionary<int, string>(){
            {50,"50mA"},
            {100,"100mA"},
            {0,"Off"}
            };
            Dictionary<int, string> mydic2 = new Dictionary<int, string>(){
            {100,"100mA"},
            {0,"Off"}
            };
            Dictionary<int, string> mydic3 = new Dictionary<int, string>(){
            {75,"75mA"},
            {0,"Off"}
            };
            try{
               
           if (ComBox_Voltage.SelectedIndex == 0 && mydic0 != null)
            {
                ComBox_DAC.ItemsSource = mydic0;
                ComBox_DAC.SelectedValuePath = "Key";
                ComBox_DAC.DisplayMemberPath = "Value";
                ComBox_DAC.SelectedIndex = 0;
            }
            else if (ComBox_Voltage.SelectedIndex == 1 && mydic1 != null)
            {
                ComBox_DAC.ItemsSource = mydic1;
                ComBox_DAC.SelectedValuePath = "Key";
                ComBox_DAC.DisplayMemberPath = "Value";
                ComBox_DAC.SelectedIndex = 0;
            }
            else if (ComBox_Voltage.SelectedIndex == 2 && mydic2 != null)
            {
                ComBox_DAC.ItemsSource = mydic2;
                ComBox_DAC.SelectedValuePath = "Key";
                ComBox_DAC.DisplayMemberPath = "Value";
                ComBox_DAC.SelectedIndex = 0;
            }
            else if (ComBox_Voltage.SelectedIndex == 3 && mydic3 != null)
            {
                ComBox_DAC.ItemsSource = mydic3;
                ComBox_DAC.SelectedValuePath = "Key";
                ComBox_DAC.DisplayMemberPath = "Value";
                ComBox_DAC.SelectedIndex = 0;
            }
            }
            catch
            {
                //MessageBox.Show("error2");
            };
        }

        private void btn_send(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.Open();
                Byte[] bt1 = new Byte[10], bt2 = new Byte[10], bt3 = new Byte[10];
                bt1[0] = bt2[0] = bt3[0] = 0x02;
                bt1[1] = 0x32; bt1[2] = 0x32;
                bt2[1] = 0x32; bt2[2] = 0x33;
                bt3[1] = 0x33; bt3[2] = 0x31;
                bt1[9] = bt2[9] = bt3[9] = 0x0D;//结束标志位
                int[,] t = new int[3, 10];
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        t[i, j] = 0;
                    }
                }
                if (ComBox_Bridge.SelectedIndex == 1)
                {
                    t[1, 6] += 8;
                }
                if (K == 1)//Enabled
                {
                    t[2, 5] += 8;
                }
                t[1, 4] += ComBox_Voltage.SelectedIndex * 2;
                int r;
               // if (K == 1)//Enabled
                //{
                    r = Convert.ToInt32(LED2_Current.Text);
                    t[0, 7] += r / 16;
                    t[0, 8] += r % 16;
                   
                    r = Convert.ToInt32(LED1_Current.Text);
                    t[0, 5] += r / 16;
                    t[0, 6] += r % 16;
                /*}
                else
                {
                    r = Convert.ToInt32(LED2_Current.Text);
                    t[0, 7] += r / 16;
                    t[0, 8] += r % 16;

                    r = Convert.ToInt32(LED1_Current.Text);
                    t[0, 5] += r / 16;
                    t[0, 6] += r % 16;
                }*/
                if (ComBox_Voltage.SelectedIndex == 0 || ComBox_Voltage.SelectedIndex == 1)//0.25 or 0.5
                {
                    t[0, 4] += ComBox_DAC.SelectedIndex + 1;
                }
                else//1.0 or 0.75
                {
                    if (ComBox_DAC.SelectedIndex == 0)
                    {
                        t[0, 4] += 1;
                    }
                    else
                    {
                        t[0, 4] += 3;
                    }
                }
                for (int j = 3; j <= 8; j++)
                {
                    bt1[j] = change(t[0, j]);
                    bt2[j] = change(t[1, j]);
                    bt3[j] = change(t[2, j]);
                }

              
                   port.Write(bt1, 0, 10);
                    Thread.Sleep(50);
                   port.Write(bt2, 0, 10);
                    Thread.Sleep(50);
                    port.Write(bt3, 0, 10);
                    Thread.Sleep(50);    
                    //检测
                   /* string str1 ="", str2 ="", str3 = "";
                      for (int i = 0; i < 10; i++)
                      {
                          str1 += bt1[i].ToString() + " ";
                          str2 += bt2[i].ToString() + " ";
                          str3 += bt3[i].ToString() + " ";
                      }*/
                    port.Close();
                    MessageBox.Show("success");
                   //  MessageBox.Show(str1);
                   //  MessageBox.Show(str2);
                    // MessageBox.Show(str3);
                   
               
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message.ToString());
            }
        }
       
       

        private void TextBox_LED2_LostFocus(object sender, RoutedEventArgs e)
        {
            float f = float.Parse(TextBox_LED2.Text),t;
            int y;
            Data_LED dt = new Data_LED();
           if ((int)ComBox_DAC.SelectedValue != 0)
            {
                t = f / float.Parse(ComBox_DAC.SelectedValue.ToString()) * 255;
                y = (int)Math.Round(t, MidpointRounding.AwayFromZero);
                dt.Data = y.ToString();
            }
            else
            {
                dt.Data = "0";
            }
            Binding binding = new Binding();
            binding.Source = dt;
            binding.Path = new PropertyPath("Data");
            BindingOperations.SetBinding(LED2_Current, TextBox.TextProperty, binding);
        }

        private void TextBox_LED1_LostFocus(object sender, RoutedEventArgs e)
        {
            float f = float.Parse(TextBox_LED1.Text), t;
            int y;
            Data_LED dt = new Data_LED();
            if ((int)ComBox_DAC.SelectedValue != 0)
            {
                t = f / float.Parse(ComBox_DAC.SelectedValue.ToString()) * 255;
                y = (int)Math.Round(t, MidpointRounding.AwayFromZero);
                dt.Data = y.ToString();
            }
            else
            {
                dt.Data = "0";
            }
            Binding binding = new Binding();
            binding.Source = dt;
            binding.Path = new PropertyPath("Data");
            BindingOperations.SetBinding(LED1_Current, TextBox.TextProperty, binding);
        }

        private void ComBox_DAC_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            float f = float.Parse(TextBox_LED1.Text), t;
            int y;
            Data_LED dt = new Data_LED();
            try
            {
                if ((int)ComBox_DAC.SelectedValue != 0)
                {
                    t = f / (float)ComBox_DAC.SelectedValue * 255;
                    y = (int)Math.Round(t, MidpointRounding.AwayFromZero);
                    dt.Data = y.ToString();
                }
                else
                {
                    dt.Data = "0";
                }
                Binding binding = new Binding();
                binding.Source = dt;
                binding.Path = new PropertyPath("Data");
                BindingOperations.SetBinding(LED1_Current, TextBox.TextProperty, binding);
                f = float.Parse(TextBox_LED2.Text);
                dt = new Data_LED();
                if ((int)ComBox_DAC.SelectedValue != 0)
                {
                    t = f / (float)ComBox_DAC.SelectedValue * 255;
                    y = (int)Math.Round(t, MidpointRounding.AwayFromZero);
                    dt.Data = y.ToString();
                }
                else
                {
                    dt.Data = "0";
                }
                binding = new Binding();
                binding.Source = dt;
                binding.Path = new PropertyPath("Data");
                BindingOperations.SetBinding(LED2_Current, TextBox.TextProperty, binding);
            }
            catch
            {
                //MessageBox.Show("error");
            }
        }

        private void Btn_SaveFile(object sender, RoutedEventArgs e)
        {
            FileStream fs = new FileStream("data.ini",FileMode.Open);
            StreamWriter sw = new StreamWriter(fs);
            string str = "";
            str += Low_1.Text + " " + High_1.Text + " " + Low_2.Text + " " + High_2.Text + " " +
            Low_3.Text + " " + High_3.Text + " " + Low_4.Text + " " + High_4.Text + " " + Low_5.Text + " " +
            High_5.Text + " " + Low_6.Text + " " + High_6.Text + " " + Low_7.Text + " " + High_7.Text + " " +
            Low_8.Text + " " + High_8.Text + " " + Low_9.Text + " " + High_9.Text + " " + Low_10.Text + " " + High_10.Text + " " +
            Low_11.Text + " " + High_11.Text + " " + Low_12.Text + " " + High_12.Text + " " + Low_13.Text + " " + High_13.Text + " " +
            Low_14.Text + " " + High_14.Text + " " + Low_15.Text + " " + High_15.Text+" "+Pr_textBox.Text+" "+Nu_textBox.Text;
            sw.Write(str);
            sw.Flush();
            sw.Close();
            fs.Close();
            MessageBox.Show("Save success");
        }

        private void Btn_LoadFile(object sender, RoutedEventArgs e)
        {
            FileStream fs = new FileStream("data.ini", FileMode.Open);
            StreamReader sr = new StreamReader(fs);
            string str = sr.ReadLine();
            string[] strs = str.Split(' ');
            MessageBox.Show(strs.Length.ToString());
            sr.Close();
            fs.Close();
            Low_1.Text = strs[0]; High_1.Text = strs[1];
            Low_2.Text = strs[2]; High_2.Text = strs[3];
            Low_3.Text = strs[4]; High_3.Text = strs[5];
            Low_4.Text = strs[6]; High_4.Text = strs[7];
            Low_5.Text = strs[8]; High_5.Text = strs[9];
            Low_6.Text = strs[10]; High_6.Text = strs[11];
            Low_7.Text = strs[12]; High_7.Text = strs[13];
            Low_8.Text = strs[14]; High_8.Text = strs[15];
            Low_9.Text = strs[16]; High_9.Text = strs[17];
            Low_10.Text = strs[18]; High_10.Text = strs[19];
            Low_11.Text = strs[20]; High_11.Text = strs[21];
            Low_12.Text = strs[22]; High_12.Text = strs[23];
            Low_13.Text = strs[24]; High_13.Text = strs[25];
            Low_14.Text = strs[26]; High_14.Text = strs[27];
            Low_15.Text = strs[28]; High_15.Text = strs[29];
            Pr_textBox.Text = strs[30];
            Nu_textBox.Text = strs[31];
        }

        private void Btn_Send_Timing(object sender, RoutedEventArgs e)
        {
            
            Byte[] bt1 = new Byte[10], bt2 = new Byte[10], bt3 = new Byte[10],
                bt4 = new Byte[10],bt5 = new Byte[10],bt6 = new Byte[10],
                bt7 = new Byte[10], bt8 = new Byte[10], bt9 = new Byte[10],
                bt10 = new Byte[10], bt11 = new Byte[10], bt12 = new Byte[10],
                bt13 = new Byte[10], bt14 = new Byte[10], bt15 = new Byte[10],
                bt16 = new Byte[10], bt17 = new Byte[10], bt18 = new Byte[10],
                bt19 = new Byte[10], bt20 = new Byte[10], bt21 = new Byte[10],
                bt22 = new Byte[10], bt23 = new Byte[10], bt24 = new Byte[10],
                bt25 = new Byte[10], bt26 = new Byte[10], bt27 = new Byte[10],
                bt28 = new Byte[10], bt29 = new Byte[10], bt30 = new Byte[10];
            bt1[0] = bt2[0] = bt3[0] = bt4[0] = bt5[0] = bt6[0] = bt7[0] =
                 bt8[0] =  bt9[0] = bt10[0] = bt11[0] =
                  bt12[0] = bt13[0] = bt14[0] = bt15[0] =
                  bt16[0] = bt17[0] = bt18[0] = bt19[0] = bt20[0] = bt21[0] = bt22[0] =
                 bt23[0] = bt24[0] = bt25[0] = bt26[0] = bt27[0] =
                  bt28[0] = bt29[0] = bt30[0] =  0x02;//命令标志
            bt1[9] = bt2[9] = bt3[9] = bt4[9] = bt5[9] = bt6[9] = bt7[9] =
                bt8[9] = bt9[9] = bt10[9] = bt11[9] =
                 bt12[9] = bt13[9] = bt14[9] = bt15[9] =
                 bt16[9] = bt17[9] = bt18[9] = bt19[9] = bt20[9] = bt21[9] = bt22[9] =
                bt23[9] = bt24[9] = bt25[9] = bt26[9] = bt27[9] =
                 bt28[9] = bt29[9] = bt30[9] = 0x0D;//结束标志位
            bt1[1] = 0x30; bt1[2] = 0x33;
            bt2[1] = 0x30; bt2[2] = 0x34;
            bt3[1] = 0x30; bt3[2] = 0x39;
            bt4[1] = 0x30; bt4[2] = 0x41;
            bt5[1] = 0x30; bt5[2] = 0x31;
            bt6[1] = 0x30; bt6[2] = 0x32;
            bt7[1] = 0x30; bt7[2] = 0x37;
            bt8[1] = 0x30; bt8[2] = 0x38;
            bt9[1] = 0x30; bt9[2] = 0x35;
            bt10[1] = 0x30; bt10[2] = 0x36;
            bt11[1] = 0x30; bt11[2] = 0x42;
            bt12[1] = 0x30; bt12[2] = 0x43;
            bt13[1] = 0x30; bt13[2] = 0x44;
            bt14[1] = 0x30; bt14[2] = 0x45;
            bt15[1] = 0x31; bt15[2] = 0x35;

            bt16[1] = 0x31; bt16[2] = 0x36;
            bt17[1] = 0x30; bt17[2] = 0x46;
            bt18[1] = 0x31; bt18[2] = 0x30;
            bt19[1] = 0x31; bt19[2] = 0x37;
            bt20[1] = 0x31; bt20[2] = 0x38;
            bt21[1] = 0x31; bt21[2] = 0x31;
            bt22[1] = 0x31; bt22[2] = 0x32;
            bt23[1] = 0x31; bt23[2] = 0x39;
            bt24[1] = 0x31; bt24[2] = 0x41;
            bt25[1] = 0x31; bt25[2] = 0x33;
            bt26[1] = 0x31; bt26[2] = 0x34;
            bt27[1] = 0x31; bt27[2] = 0x42;
            bt28[1] = 0x31; bt28[2] = 0x43;
            bt29[1] = 0x33; bt29[2] = 0x32;
            bt30[1] = 0x33; bt30[2] = 0x33;
            int[,] t = new int[30, 10];
            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j <= 9; j++)
                {
                    t[i, j] = 0;
                }
            }
            int t1, t2;
            try
            {
            t1 =Convert.ToInt32(Low_1.Text)/256;
            t2 = Convert.ToInt32(Low_1.Text) % 256;
            t[0, 5] = t1 / 16;
            t[0, 6] = t1 % 16;
            t[0, 7] = t2 / 16;
            t[0, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_1.Text) / 256;
            t2 = Convert.ToInt32(High_1.Text) % 256;
            t[1, 5] = t1 / 16;
            t[1, 6] = t1 % 16;
            t[1, 7] = t2 / 16;
            t[1, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_2.Text) / 256;
            t2 = Convert.ToInt32(Low_2.Text) % 256;
            t[2, 5] = t1 / 16;
            t[2, 6] = t1 % 16;
            t[2, 7] = t2 / 16;
            t[3, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_2.Text) / 256;
            t2 = Convert.ToInt32(High_2.Text) % 256;
            t[3, 5] = t1 / 16;
            t[3, 6] = t1 % 16;
            t[3, 7] = t2 / 16;
            t[3, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_3.Text) / 256;
            t2 = Convert.ToInt32(Low_3.Text) % 256;
            t[4, 5] = t1 / 16;
            t[4, 6] = t1 % 16;
            t[4, 7] = t2 / 16;
            t[4, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_3.Text) / 256;
            t2 = Convert.ToInt32(High_3.Text) % 256;
            t[5, 5] = t1 / 16;
            t[5, 6] = t1 % 16;
            t[5, 7] = t2 / 16;
            t[5, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_4.Text) / 256;
            t2 = Convert.ToInt32(Low_4.Text) % 256;
            t[6, 5] = t1 / 16;
            t[6, 6] = t1 % 16;
            t[6, 7] = t2 / 16;
            t[6, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_4.Text) / 256;
            t2 = Convert.ToInt32(High_4.Text) % 256;
            t[7, 5] = t1 / 16;
            t[7, 6] = t1 % 16;
            t[7, 7] = t2 / 16;
            t[7, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_5.Text) / 256;
            t2 = Convert.ToInt32(Low_5.Text) % 256;
            t[8, 5] = t1 / 16;
            t[8, 6] = t1 % 16;
            t[8, 7] = t2 / 16;
            t[8, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_5.Text) / 256;
            t2 = Convert.ToInt32(High_5.Text) % 256;
            t[9, 5] = t1 / 16;
            t[9, 6] = t1 % 16;
            t[9, 7] = t2 / 16;
            t[9, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_6.Text) / 256;
            t2 = Convert.ToInt32(Low_6.Text) % 256;
            t[10, 5] = t1 / 16;
            t[10, 6] = t1 % 16;
            t[10, 7] = t2 / 16;
            t[10, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_6.Text) / 256;
            t2 = Convert.ToInt32(High_6.Text) % 256;
            t[11, 5] = t1 / 16;
            t[11, 6] = t1 % 16;
            t[11, 7] = t2 / 16;
            t[11, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_7.Text) / 256;
            t2 = Convert.ToInt32(Low_7.Text) % 256;
            t[12, 5] = t1 / 16;
            t[12, 6] = t1 % 16;
            t[12, 7] = t2 / 16;
            t[12, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_7.Text) / 256;
            t2 = Convert.ToInt32(High_7.Text) % 256;
            t[13, 5] = t1 / 16;
            t[13, 6] = t1 % 16;
            t[13, 7] = t2 / 16;
            t[13, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_8.Text) / 256;
            t2 = Convert.ToInt32(Low_8.Text) % 256;
            t[14, 5] = t1 / 16;
            t[14, 6] = t1 % 16;
            t[14, 7] = t2 / 16;
            t[14, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_8.Text) / 256;
            t2 = Convert.ToInt32(High_8.Text) % 256;
            t[15, 5] = t1 / 16;
            t[15, 6] = t1 % 16;
            t[15, 7] = t2 / 16;
            t[15, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_9.Text) / 256;
            t2 = Convert.ToInt32(Low_9.Text) % 256;
            t[16, 5] = t1 / 16;
            t[16, 6] = t1 % 16;
            t[16, 7] = t2 / 16;
            t[16, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_9.Text) / 256;
            t2 = Convert.ToInt32(High_9.Text) % 256;
            t[17, 5] = t1 / 16;
            t[17, 6] = t1 % 16;
            t[17, 7] = t2 / 16;
            t[17, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_10.Text) / 256;
            t2 = Convert.ToInt32(Low_10.Text) % 256;
            t[18, 5] = t1 / 16;
            t[18, 6] = t1 % 16;
            t[18, 7] = t2 / 16;
            t[18, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_10.Text) / 256;
            t2 = Convert.ToInt32(High_10.Text) % 256;
            t[19, 5] = t1 / 16;
            t[19, 6] = t1 % 16;
            t[19, 7] = t2 / 16;
            t[19, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_11.Text) / 256;
            t2 = Convert.ToInt32(Low_11.Text) % 256;
            t[20, 5] = t1 / 16;
            t[20, 6] = t1 % 16;
            t[20, 7] = t2 / 16;
            t[20, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_11.Text) / 256;
            t2 = Convert.ToInt32(High_11.Text) % 256;
            t[21, 5] = t1 / 16;
            t[21, 6] = t1 % 16;
            t[21, 7] = t2 / 16;
            t[21, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_12.Text) / 256;
            t2 = Convert.ToInt32(Low_12.Text) % 256;
            t[22, 5] = t1 / 16;
            t[22, 6] = t1 % 16;
            t[22, 7] = t2 / 16;
            t[22, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_12.Text) / 256;
            t2 = Convert.ToInt32(High_12.Text) % 256;
            t[23, 5] = t1 / 16;
            t[23, 6] = t1 % 16;
            t[23, 7] = t2 / 16;
            t[23, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_13.Text) / 256;
            t2 = Convert.ToInt32(Low_13.Text) % 256;
            t[24, 5] = t1 / 16;
            t[24, 6] = t1 % 16;
            t[24, 7] = t2 / 16;
            t[24, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_13.Text) / 256;
            t2 = Convert.ToInt32(High_13.Text) % 256;
            t[25, 5] = t1 / 16;
            t[25, 6] = t1 % 16;
            t[25, 7] = t2 / 16;
            t[25, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_14.Text) / 256;
            t2 = Convert.ToInt32(Low_14.Text) % 256;
            t[26, 5] = t1 / 16;
            t[26, 6] = t1 % 16;
            t[26, 7] = t2 / 16;
            t[26, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_14.Text) / 256;
            t2 = Convert.ToInt32(High_14.Text) % 256;
            t[27, 5] = t1 / 16;
            t[27, 6] = t1 % 16;
            t[27, 7] = t2 / 16;
            t[27, 8] = t2 % 16;

            t1 = Convert.ToInt32(Low_15.Text) / 256;
            t2 = Convert.ToInt32(Low_15.Text) % 256;
            t[28, 5] = t1 / 16;
            t[28, 6] = t1 % 16;
            t[28, 7] = t2 / 16;
            t[28, 8] = t2 % 16;

            t1 = Convert.ToInt32(High_15.Text) / 256;
            t2 = Convert.ToInt32(High_15.Text) % 256;
            t[29, 5] = t1 / 16;
            t[29, 6] = t1 % 16;
            t[29, 7] = t2 / 16;
            t[29, 8] = t2 % 16;
            for (int j = 3; j <= 8; j++)
            {
                bt1[j] = change(t[0, j]);
                bt2[j] = change(t[1, j]);
                bt3[j] = change(t[2, j]);
                bt4[j] = change(t[3, j]);
                bt5[j] = change(t[4, j]);
                bt6[j] = change(t[5, j]);
                bt7[j] = change(t[6, j]);
                bt8[j] = change(t[7, j]);
                bt9[j] = change(t[8, j]);
                bt10[j] = change(t[9, j]);
                bt11[j] = change(t[10, j]);
                bt12[j] = change(t[11, j]);
                bt13[j] = change(t[12, j]);
                bt14[j] = change(t[13, j]);
                bt15[j] = change(t[14, j]);
                bt16[j] = change(t[15, j]);
                bt17[j] = change(t[16, j]);
                bt18[j] = change(t[17, j]);
                bt19[j] = change(t[18, j]);
                bt20[j] = change(t[19, j]);
                bt21[j] = change(t[20, j]);
                bt22[j] = change(t[21, j]);
                bt23[j] = change(t[22, j]);
                bt24[j] = change(t[23, j]);
                bt25[j] = change(t[24, j]);
                bt26[j] = change(t[25, j]);
                bt27[j] = change(t[26, j]);
                bt28[j] = change(t[27, j]);
                bt29[j] = change(t[28, j]);
                bt30[j] = change(t[29, j]);
                
            }

            Byte[] Pr = new Byte[10], Te = new Byte[10], Nu = new Byte[10];
            Pr[0] = Te[0] = Nu[0] = 0x02;
            Pr[9] = Te[9] = Nu[9] = 0x0D;
            Pr[1] = 0x31; Pr[2] = 0x44;
            Te[1] = 0x31; Te[2] = 0x45;
            Nu[1] = 0x31; Nu[2] = 0x45;
           
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j <= 9; j++)
                {
                    t[i, j] = 0;
                }
            }
            t1 = Convert.ToInt32(Pr_textBox.Text)/256;
            t2 = Convert.ToInt32(Pr_textBox.Text) % 256;
            t[0, 5] = t1 / 16; t[0, 6] = t1 % 16;
            t[0, 7] = t2 / 16; t[0, 8] = t2 % 16;

            if (En_Check.IsChecked == true)
            {
                t[1, 6] = 1;
            }
            t[2, 7] = Convert.ToInt32(Nu_textBox.Text);
            t[2, 8] = Convert.ToInt32(Nu_textBox.Text);
            for (int i = 3; i <= 8; i++)
            {
                Pr[i] = change(t[0, i]);
                Te[i] = change(t[1, i]);
                Nu[i] = change(t[2, i]);
            }
           
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.Open();
               
                    port.Write(bt1, 0, 10); Thread.Sleep(30);
                    port.Write(bt2, 0, 10); Thread.Sleep(30);
                    port.Write(bt3, 0, 10); Thread.Sleep(30);
                    port.Write(bt4, 0, 10); Thread.Sleep(30);
                    port.Write(bt5, 0, 10); Thread.Sleep(30);
                    port.Write(bt6, 0, 10); Thread.Sleep(30);
                    port.Write(bt7, 0, 10); Thread.Sleep(30);
                    port.Write(bt8, 0, 10); Thread.Sleep(30);
                    port.Write(bt9, 0, 10); Thread.Sleep(30);
                    port.Write(bt10, 0, 10); Thread.Sleep(30);
                    port.Write(bt11, 0, 10); Thread.Sleep(30);
                    port.Write(bt12, 0, 10); Thread.Sleep(30);
                    port.Write(bt13, 0, 10); Thread.Sleep(30);
                    port.Write(bt14, 0, 10); Thread.Sleep(30);
                    port.Write(bt15, 0, 10); Thread.Sleep(30);
                    port.Write(bt16, 0, 10); Thread.Sleep(30);
                    port.Write(bt17, 0, 10); Thread.Sleep(30);
                    port.Write(bt18, 0, 10); Thread.Sleep(30);
                    port.Write(bt19, 0, 10); Thread.Sleep(30);
                    port.Write(bt20, 0, 10); Thread.Sleep(30);
                    port.Write(bt21, 0, 10); Thread.Sleep(30);
                    port.Write(bt22, 0, 10); Thread.Sleep(30);
                    port.Write(bt23, 0, 10); Thread.Sleep(30);
                    port.Write(bt24, 0, 10); Thread.Sleep(30);
                    port.Write(bt25, 0, 10); Thread.Sleep(30);
                    port.Write(bt26, 0, 10); Thread.Sleep(30);
                    port.Write(bt27, 0, 10); Thread.Sleep(30);
                    port.Write(bt28, 0, 10); Thread.Sleep(30);
                    port.Write(bt29, 0, 10); Thread.Sleep(30);
                    port.Write(bt30, 0, 10); Thread.Sleep(30);
                    port.Write(Pr, 0, 10); Thread.Sleep(30);
                    port.Write(Te, 0, 10); Thread.Sleep(30);
                    port.Write(Nu, 0, 10); Thread.Sleep(30);
                    /* 输出检测 
                     * string[] S = new string[33];
                      for (int i = 0; i < 9; i++)
                      {
                          S[0] += bt1[i].ToString("X") + " ";
                          S[1] += bt2[i].ToString("X") + " ";
                          S[2] += bt3[i].ToString("X") + " ";
                          S[3] += bt4[i].ToString("X") + " ";
                          S[4] += bt5[i].ToString("X") + " ";
                          S[5] += bt6[i].ToString("X") + " ";
                          S[6] += bt7[i].ToString("X") + " ";
                          S[7] += bt8[i].ToString("X") + " ";
                          S[8] += bt9[i].ToString("X") + " ";
                          S[9] += bt10[i].ToString("X") + " ";
                          S[10] += bt11[i].ToString("X") + " ";
                          S[11] += bt12[i].ToString("X") + " ";
                          S[12] += bt13[i].ToString("X") + " ";
                          S[13] += bt14[i].ToString("X") + " ";
                          S[14] += bt15[i].ToString("X") + " ";
                          S[15] += bt16[i].ToString("X") + " ";
                          S[16] += bt17[i].ToString("X") + " ";
                          S[17] += bt18[i].ToString("X") + " ";
                          S[18] += bt19[i].ToString("X") + " ";
                          S[19] += bt20[i].ToString("X") + " ";
                          S[20] += bt21[i].ToString("X") + " ";
                          S[21] += bt22[i].ToString("X") + " ";
                          S[22] += bt23[i].ToString("X") + " ";
                          S[23] += bt24[i].ToString("X") + " ";
                          S[24] += bt25[i].ToString("X") + " ";
                          S[25] += bt26[i].ToString("X") + " ";
                          S[26] += bt27[i].ToString("X") + " ";
                          S[27] += bt28[i].ToString("X") + " ";
                          S[28] += bt29[i].ToString("X") + " ";
                          S[29] += bt30[i].ToString("X") + " ";
                          S[30] += Pr[i].ToString("X") + " ";
                          S[31] += Te[i].ToString("X") + " ";
                          S[32] += Nu[i].ToString("X") + " ";
                          }
                      string r = "";
                      for (int y = 0; y < 33; y++)
                      {
                          r += S[y] + "\n";
                      }
                      MessageBox.Show(r);*/
                    port.Close();
                    MessageBox.Show("success");
                
             
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void btn_connect_Click(object sender, RoutedEventArgs e)
        {
            int[] IntBdrs = { 9600,115200,230400 };
            string[] CommNums = { "COM1", "COM2", "COM3", "COM4","COM5","COM6","COM7","COM8" };
            int num;
            num = ComBox_CommNum.SelectedIndex;
            CommNum = CommNums[num];
            num = ComBox_IntBdr.SelectedIndex;
            IntBdr = IntBdrs[num];
            SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
            if (!port.IsOpen)
            {
                try
                {
                    port.Open();
                   
                    check(port);
                    port.Close();
                    MessageBox.Show("打开成功");
                }
                catch
                {

                    MessageBox.Show("串口打开失败,可能是串口已被占用");
                }
            }
        }

        private void Btn_Data_setting_Click(object sender, RoutedEventArgs e)
        {
            Grid_Data_setting.Visibility = (Visibility)0;
            Grid_Sample_Test.Visibility = (Visibility)2;
            Grid_Bluetooth.Visibility = (Visibility)2;
           
         
        }

        private void Btn_Sample_Test_Click(object sender, RoutedEventArgs e)
        {
            Grid_Data_setting.Visibility = (Visibility)2;
            Grid_Sample_Test.Visibility = (Visibility)0;
            Grid_Bluetooth.Visibility = (Visibility)2;
        }

        private void Btn_Bluetooth_Click(object sender, RoutedEventArgs e)
        {
            Grid_Data_setting.Visibility = (Visibility)2;
            Grid_Sample_Test.Visibility = (Visibility)2;
            Grid_Bluetooth.Visibility = (Visibility)0;
        }
       /* private Dictionary<string, string> yuansu = new Dictionary<string, string>(){
            {"三聚氰胺","三聚氰胺"},
            {"亚硝酸盐","亚硝酸盐"},
            {"农药残留","农药残留"},
            {"黄曲霉素","黄曲霉素"},
            {"铅","铅"},
            {"汞","汞"},
            {"砷","砷"},
            {"镉","镉"}
            };*/

        private void New_Set_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FileStream fs = new FileStream("Yuansu.ini", FileMode.Append);
                StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("gb2312"));
                sw.Write(" " + TextBox_NewData.Text);
                sw.Close();
                fs.Close();
                FileStream fs1 = new FileStream("Yuansu.ini", FileMode.Open);
                StreamReader sr = new StreamReader(fs1, Encoding.GetEncoding("gb2312"));
                string[] data = sr.ReadLine().Split(' ');
                Dictionary<string, string> Dic = new Dictionary<string, string>
                {

                };
                //MessageBox.Show(data.ToString());
                for (int i = 0; i < data.Length; i++)
                {
                    Dic.Add(data[i], data[i]);
                }
                //StreamWriter sw = new StreamWriter(fs);
                MessageBox.Show("添加成功");
                ComBox_yuansu.ItemsSource = Dic;
                ComBox_yuansu.SelectedValuePath = "Key";
                ComBox_yuansu.DisplayMemberPath = "Value";
                ComBox_yuansu.SelectedIndex = 0;
                sr.Close();
                fs1.Close();
            }
            catch 
            {
                MessageBox.Show("添加错误(不能添加已有项)");
                
            }
            
        }

        private void Btn_save_DataSet_Click(object sender, RoutedEventArgs e)
        {
            string path = "Config.xml";
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode dateNode,projectNode,minNode,maxNode,limitNode,floatNode,danweiNode,PersonNode;
           XmlElement  deviceNode;
            XmlNodeList ProjectList,Devlist;
            //文件不存在
            if (!File.Exists(path))
            {
                XmlNode Declaration = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                xmlDoc.AppendChild(Declaration);
                deviceNode = xmlDoc.CreateElement("仪器编号");
                XmlAttribute DevName=xmlDoc.CreateAttribute("Name");
               DevName.Value=Instrument_number.Text;
                deviceNode.Attributes.Append(DevName);

                dateNode = xmlDoc.CreateElement("标定日期");
                dateNode.InnerText = Date.Text;
                projectNode = xmlDoc.CreateElement("检测项目");
                XmlAttribute projectName = xmlDoc.CreateAttribute("Name");
                projectName.Value= ComBox_yuansu.SelectedValue.ToString();
                projectNode.Attributes.Append(projectName);
                minNode = xmlDoc.CreateElement("最小值");
                minNode.InnerText = Min_number.Text;
                maxNode = xmlDoc.CreateElement("最大值");
                maxNode.InnerText = Max_number.Text;
                limitNode = xmlDoc.CreateElement("最高限量");
                limitNode.InnerText = Max_limit.Text;
                floatNode = xmlDoc.CreateElement("有效位数");
                floatNode.InnerText = ComBox_float_num.SelectedIndex.ToString();
                danweiNode = xmlDoc.CreateElement("单位");
                danweiNode.InnerText = ComBox_danwei.SelectedValue.ToString();
                PersonNode = xmlDoc.CreateElement("标定人员");
                PersonNode.InnerText = Person.Text;
                xmlDoc.AppendChild(deviceNode);
                deviceNode.AppendChild(projectNode);
                projectNode.AppendChild(dateNode);
                projectNode.AppendChild(PersonNode);
                projectNode.AppendChild(minNode);
                projectNode.AppendChild(maxNode);
                projectNode.AppendChild(limitNode);
                projectNode.AppendChild(floatNode);
                projectNode.AppendChild(danweiNode);
                
            }
             //文件存在
            else
            {
                xmlDoc.Load(path);
                bool b = false;
                Devlist = xmlDoc.GetElementsByTagName("仪器编号");
                deviceNode = (XmlElement)Devlist[0];
               
                deviceNode.SetAttribute("Name", Instrument_number.Text);
                ProjectList = xmlDoc.GetElementsByTagName("检测项目");
                foreach(XmlNode node in ProjectList){
                    if (((XmlElement)node).GetAttribute("Name") == ComBox_yuansu.SelectedValue.ToString())
                    {
                        b = true;
                        dateNode = node.SelectSingleNode("标定日期");
                        dateNode.InnerText = Date.Text;
                        PersonNode = node.SelectSingleNode("标定人员");
                        PersonNode.InnerText = Person.Text;
                        minNode = node.SelectSingleNode("最小值");
                        minNode.InnerText = Min_number.Text;
                        maxNode = node.SelectSingleNode("最大值");
                        maxNode.InnerText = Max_number.Text;
                        limitNode = node.SelectSingleNode("最高限量");
                        limitNode.InnerText = Max_limit.Text;
                        floatNode = node.SelectSingleNode("有效位数");
                        floatNode.InnerText = ComBox_float_num.SelectedIndex.ToString();
                        danweiNode = node.SelectSingleNode("单位");
                        danweiNode.InnerText = ComBox_danwei.SelectedValue.ToString();
                    }
                }
                if (b == false) {
                    dateNode = xmlDoc.CreateElement("标定日期");
                    dateNode.InnerText = Date.Text;
                    projectNode = xmlDoc.CreateElement("检测项目");
                    XmlAttribute projectName = xmlDoc.CreateAttribute("Name");
                    projectName.Value = ComBox_yuansu.SelectedValue.ToString();
                    projectNode.Attributes.Append(projectName);
                    minNode = xmlDoc.CreateElement("最小值");
                    minNode.InnerText = Min_number.Text;
                    maxNode = xmlDoc.CreateElement("最大值");
                    maxNode.InnerText = Max_number.Text;
                    limitNode = xmlDoc.CreateElement("最高限量");
                    limitNode.InnerText = Max_limit.Text;
                    floatNode = xmlDoc.CreateElement("有效位数");
                    floatNode.InnerText = ComBox_float_num.SelectedIndex.ToString();
                    danweiNode = xmlDoc.CreateElement("单位");
                    danweiNode.InnerText = ComBox_danwei.SelectedValue.ToString();
                    PersonNode = xmlDoc.CreateElement("标定人员");
                    PersonNode.InnerText = Person.Text;
                    deviceNode.AppendChild(projectNode);
                    projectNode.AppendChild(dateNode);
                    projectNode.AppendChild(PersonNode);
                    projectNode.AppendChild(minNode);
                    projectNode.AppendChild(maxNode);
                    projectNode.AppendChild(limitNode);
                    projectNode.AppendChild(floatNode);
                    projectNode.AppendChild(danweiNode);
                
}
            }
            xmlDoc.Save(path);
            MessageBox.Show("save success");
        }

       private double volcal(int bytel, int bytem, int byteh)
        {
            int w, flag, i = 0;
            double voltage;
            w = (byteh * 256 + bytem) * 256 + bytel;
            flag = w & Convert.ToInt32(0x800000);
            if (flag == 0)
            {
                voltage = 1.2 * w / (Math.Pow(2, 21) - 1);
            }
            else
            {
                w = w - 16777216;
                voltage = 1.2 * w / (Math.Pow(2, 21));
            }
            return voltage;
        }


        private double[,] Datas = new double[6, 1050];
        private void Btn_ReceiveData_Click(object sender, RoutedEventArgs e)
        {
            try { 
            SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
            port.ReadTimeout = 5000;
            port.ReadBufferSize = 10000;
           port.ReceivedBytesThreshold = 22;
           int Num = Convert.ToInt32(TexBox_Num.Text);

           if (!port.IsOpen)
           {

               port.Open();
               port.DiscardInBuffer();
               port.DiscardOutBuffer();
               Byte[] bt_send_start = new Byte[2], bt_receive = new Byte[22], bt_send_end = new Byte[2], bt_First = new Byte[10];
               bt_send_start[0] = 0x01;
               bt_send_start[1] = 0x0D;


               bt_send_end[0] = 0x06; bt_send_end[1] = 0x0D;
               bt_First[0] = 0x02; bt_First[1] = 0x33; bt_First[2] = 0x34; bt_First[3] = 0x30;
               bt_First[4] = 0x30; bt_First[5] = 0x30; bt_First[6] = 0x30; bt_First[7] = 0x30;
               bt_First[9] = 0x0D; bt_First[8] = change(Dev_Select.SelectedIndex);
               port.Write(bt_First, 0, 10);
               Thread.Sleep(30);
               port.Write(bt_send_start, 0, 2); Thread.Sleep(30);
               for (int i = 0; i <= Num; i++)
               {
                  
                   port.Read(bt_receive, 0, 22);
                 // Thread.Sleep(1);
                  
                  
               /* if (bt_receive[0] == 0x01)
                  {
                      Datas[0, i] = volcal(bt_receive[2], bt_receive[3], bt_receive[4]);

                      Datas[1, i] = volcal(bt_receive[5], bt_receive[6], bt_receive[7]);
                      Datas[2, i] = volcal(bt_receive[8], bt_receive[9], bt_receive[10]);
                      Datas[3, i] = volcal(bt_receive[11], bt_receive[12], bt_receive[13]);
                      Datas[4, i] = volcal(bt_receive[14], bt_receive[15], bt_receive[16]);
                      Datas[5, i] = volcal(bt_receive[17], bt_receive[18], bt_receive[19]);
                     Datas[0, i] = ex_dll.volcal(bt_receive[2], bt_receive[3], bt_receive[4]);
                       Datas[1, i] = ex_dll.volcal(bt_receive[5], bt_receive[6], bt_receive[7]);
                       Datas[2, i] = ex_dll.volcal(bt_receive[8], bt_receive[9], bt_receive[10]);
                       Datas[3, i] = ex_dll.volcal(bt_receive[11], bt_receive[12], bt_receive[13]);
                       Datas[4, i] = ex_dll.volcal(bt_receive[14], bt_receive[15], bt_receive[16]);
                       Datas[5, i] = ex_dll.volcal(bt_receive[17], bt_receive[18], bt_receive[19]);
                   }
                  else
                   {*/
                       Datas[0, i] = volcal(bt_receive[1], bt_receive[2], bt_receive[3]);
                       Datas[1, i] = volcal(bt_receive[4], bt_receive[5], bt_receive[6]);
                       Datas[2, i] = volcal(bt_receive[7], bt_receive[8], bt_receive[9]);
                       Datas[3, i] = volcal(bt_receive[10], bt_receive[11], bt_receive[12]);
                       Datas[4, i] = volcal(bt_receive[13], bt_receive[14], bt_receive[15]);
                       Datas[5, i] = volcal(bt_receive[16], bt_receive[17], bt_receive[18]);
                     /* Datas[0, i] = ex_dll.volcal(bt_receive[1], bt_receive[2], bt_receive[3]);
                       Datas[1, i] = ex_dll.volcal(bt_receive[4], bt_receive[5], bt_receive[6]);
                       Datas[2, i] = ex_dll.volcal(bt_receive[7], bt_receive[8], bt_receive[9]);
                       Datas[3, i] = ex_dll.volcal(bt_receive[10], bt_receive[11], bt_receive[12]);
                       Datas[4, i] = ex_dll.volcal(bt_receive[13], bt_receive[14], bt_receive[15]);
                       Datas[5, i] = ex_dll.volcal(bt_receive[16], bt_receive[17], bt_receive[18]);*/
                /// }
                  switch (ComBox_Voltage.SelectedIndex)
                   {
                       case 0:
                           {
                               Datas[0, i] = xK1 * Datas[0, i] + xC1;
                               Datas[1, i] = xK1 * Datas[1, i] + xC1;
                               Datas[2, i] = xK1 * Datas[2, i] + xC1;
                               Datas[3, i] = xK1 * Datas[3, i] + xC1;
                               Datas[4, i] = xK1 * Datas[4, i] + xC1;
                               Datas[5, i] = xK1 * Datas[5, i] + xC1;
                               break;
                           }
                       case 1:
                           {
                               Datas[0, i] = xK2 * Datas[0, i] + xC2;
                               Datas[1, i] = xK2 * Datas[1, i] + xC2;
                               Datas[2, i] = xK2 * Datas[2, i] + xC2;
                               Datas[3, i] = xK2 * Datas[3, i] + xC2;
                               Datas[4, i] = xK2 * Datas[4, i] + xC2;
                               Datas[5, i] = xK2 * Datas[5, i] + xC2;
                               break;
                           }
                       case 2:
                           {
                               Datas[0, i] = xK3 * Datas[0, i] + xC3;
                               Datas[1, i] = xK3 * Datas[1, i] + xC3;
                               Datas[2, i] = xK3 * Datas[2, i] + xC3;
                               Datas[3, i] = xK3 * Datas[3, i] + xC3;
                               Datas[4, i] = xK3 * Datas[4, i] + xC3;
                               Datas[5, i] = xK3 * Datas[5, i] + xC3;
                               break;
                           }
                       case 3:
                           {
                               Datas[0, i] = xK4 * Datas[0, i] + xC4;
                               Datas[1, i] = xK4 * Datas[1, i] + xC4;
                               Datas[2, i] = xK4 * Datas[2, i] + xC4;
                               Datas[3, i] = xK4 * Datas[3, i] + xC4;
                               Datas[4, i] = xK4 * Datas[4, i] + xC4;
                               Datas[5, i] = xK4 * Datas[5, i] + xC4;
                               break;
                           }
                   }
               }
           
               port.DiscardInBuffer();
               
               port.Write(bt_send_end, 0, 2);
               Thread.Sleep(30);
               port.Close();
               MessageBox.Show("success");
               //MessageBox.Show(str);
           }
            }
            catch 
            {
                MessageBox.Show("error");
            }
        }
      
        private void Btn_Test1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int i;
                int start = Convert.ToInt32(TextBox_start1.Text);
                int end = Convert.ToInt32(TextBox_end1.Text);

                if (ComBox_Select1.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select1.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }

                }
                else if (ComBox_Select1.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }

                }
                else if (ComBox_Select1.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select1.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select1.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver1.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start2.Text), i;
                int end = Convert.ToInt32(TextBox_end2.Text);

                if (ComBox_Select2.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select2.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }

                }
                else if (ComBox_Select2.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }

                }
                else if (ComBox_Select2.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }

                }
                else if (ComBox_Select2.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select2.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver2.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test3_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start3.Text), i;
                int end = Convert.ToInt32(TextBox_end3.Text);

                if (ComBox_Select3.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select3.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select3.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select3.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select3.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select3.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }

                ave = sum / (end - start + 1);
                TextBox_Aver3.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test4_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start4.Text), i;
                int end = Convert.ToInt32(TextBox_end4.Text);

                if (ComBox_Select4.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select4.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select4.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select4.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select4.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select4.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver4.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test5_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start5.Text), i;
                int end = Convert.ToInt32(TextBox_end5.Text);

                if (ComBox_Select5.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select5.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select5.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select5.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select5.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select5.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver5.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Save10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = "Analysis.xml";
                XmlDocument xmlDoc = new XmlDocument();
                if (!File.Exists(path))
                {
                    XmlNode node = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                    xmlDoc.AppendChild(node);
                    XmlElement root = xmlDoc.CreateElement("root");

                    xmlDoc.AppendChild(root);
                    XmlNode xml_Dev = xmlDoc.CreateElement("Dev");
                    root.AppendChild(xml_Dev);
                    XmlAttribute xml_DevName = xmlDoc.CreateAttribute("DevName");
                    if (Dev_Select.SelectedIndex == 0)
                    {
                        xml_DevName.Value = "探测器A";
                    }
                    else
                    {
                        xml_DevName.Value = "探测器B";
                    }
                    xml_Dev.Attributes.Append(xml_DevName);
                    XmlAttribute xml_nongdu = xmlDoc.CreateAttribute("nongdu");
                    xml_nongdu.Value = TextBox_nongdu.Text;
                    xml_Dev.Attributes.Append(xml_nongdu);
                    XmlAttribute xml_num = xmlDoc.CreateAttribute("Num");
                    xml_num.Value = TexBox_Num.Text;
                    xml_Dev.Attributes.Append(xml_num);
                    XmlAttribute xml_Select = xmlDoc.CreateAttribute("Select");
                    switch (ComBox_Select10.SelectedIndex)
                    {
                        case 0:
                            {
                                xml_Select.Value = "LED2";
                                break;
                            }
                        case 1:
                            {
                                xml_Select.Value = "ALED2";
                                break;
                            }
                        case 2:
                            {
                                xml_Select.Value = "LED1";
                                break;
                            }
                        case 3:
                            {
                                xml_Select.Value = "ALED1";
                                break;
                            }
                        case 4:
                            {
                                xml_Select.Value = "LED2-ALED2";
                                break;
                            }
                        case 5:
                            {
                                xml_Select.Value = "LED1-ALED1";
                                break;
                            }
                    }
                    xml_Dev.Attributes.Append(xml_Select);
                    XmlAttribute xml_start = xmlDoc.CreateAttribute("start");
                    xml_start.Value = TextBox_start10.Text;
                    xml_Dev.Attributes.Append(xml_start);
                    XmlAttribute xml_end = xmlDoc.CreateAttribute("end");
                    xml_end.Value = TextBox_end10.Text;
                    xml_Dev.Attributes.Append(xml_end);
                    XmlAttribute xml_Ave = xmlDoc.CreateAttribute("Ave");
                    xml_Ave.Value = TextBox_Aver10.Text;
                    xml_Dev.Attributes.Append(xml_Ave);
                }
                else
                {
                    string DevName="";
                     if (Dev_Select.SelectedIndex == 0)
                    {
                        DevName = "探测器A";
                    }
                    else
                    {
                        DevName = "探测器B";
                    }
                     string SelectValue = "";
                     switch (ComBox_Select10.SelectedIndex)
                     {
                         case 0:
                             {
                                 SelectValue = "LED2";
                                 break;
                             }
                         case 1:
                             {
                                 SelectValue = "ALED2";
                                 break;
                             }
                         case 2:
                             {
                                 SelectValue = "LED1";
                                 break;
                             }
                         case 3:
                             {
                                 SelectValue = "ALED1";
                                 break;
                             }
                         case 4:
                             {
                                 SelectValue = "LED2-ALED2";
                                 break;
                             }
                         case 5:
                             {
                                 SelectValue = "LED1-ALED1";
                                 break;
                             }
                     }

                    xmlDoc.Load("Analysis.xml");
                    XmlNode xml_rootNode = xmlDoc.SelectSingleNode("root");
                    XmlNodeList DevList = xmlDoc.GetElementsByTagName("Dev");
                    bool b = false;
                    foreach (XmlNode node in DevList)
                    {
                        if (((XmlElement)node).GetAttribute("DevName") == DevName && ((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu.Text && ((XmlElement)node).GetAttribute("Select")==SelectValue)
                        {
                           
                            b = true;
                            ((XmlElement)node).SetAttribute("Num", TexBox_Num.Text);
                            ((XmlElement)node).SetAttribute("start",TextBox_start10.Text);
                            ((XmlElement)node).SetAttribute("end", TextBox_end10.Text);
                            ((XmlElement)node).SetAttribute("Ave", TextBox_Aver10.Text);

                            break;
                        }
                    }
                    if (b == false)
                    {

                        XmlNode xml_Dev = xmlDoc.CreateElement("Dev");
                        xml_rootNode.AppendChild(xml_Dev);
                        XmlAttribute xml_DevName = xmlDoc.CreateAttribute("DevName");
                        xml_DevName.Value = DevName;
                        xml_Dev.Attributes.Append(xml_DevName);
                        XmlAttribute xml_nongdu = xmlDoc.CreateAttribute("nongdu");
                        xml_nongdu.Value = TextBox_nongdu.Text;
                        xml_Dev.Attributes.Append(xml_nongdu);
                        XmlAttribute xml_num = xmlDoc.CreateAttribute("Num");
                        xml_num.Value = TexBox_Num.Text;
                        xml_Dev.Attributes.Append(xml_num);
                        XmlAttribute xml_Select = xmlDoc.CreateAttribute("Select");
                        xml_Select.Value = SelectValue;
                        xml_Dev.Attributes.Append(xml_Select);
                        XmlAttribute xml_start = xmlDoc.CreateAttribute("start");
                        xml_start.Value = TextBox_start10.Text;
                        xml_Dev.Attributes.Append(xml_start);
                        XmlAttribute xml_end = xmlDoc.CreateAttribute("end");
                        xml_end.Value = TextBox_end10.Text;
                        xml_Dev.Attributes.Append(xml_end);
                        XmlAttribute xml_Ave = xmlDoc.CreateAttribute("Ave");
                        xml_Ave.Value = TextBox_Aver10.Text;
                        xml_Dev.Attributes.Append(xml_Ave);
                    }
                }
                xmlDoc.Save(path);
                MessageBox.Show("save success");
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

       

        private void Btntest1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED1.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                    case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");

                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu1.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA1.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB1.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA1.Text = (float.Parse(TextBox_DevA1.Text) / float.Parse(TextBox_DevB1.Text)).ToString();
                TextBox_resultB1.Text = (float.Parse(TextBox_DevB1.Text) / float.Parse(TextBox_DevA1.Text)).ToString();
                lnAB1.Text = Math.Log(double.Parse(TextBox_resultA1.Text)).ToString();
                lnBA1.Text = Math.Log(double.Parse(TextBox_resultB1.Text)).ToString();
            }
            catch 
            {
                MessageBox.Show("error");
            }
        }

        private void Btntest2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED2.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                     case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");
                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu2.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA2.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB2.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA2.Text = (float.Parse(TextBox_DevA2.Text) / float.Parse(TextBox_DevB2.Text)).ToString();
                TextBox_resultB2.Text = (float.Parse(TextBox_DevB2.Text) / float.Parse(TextBox_DevA2.Text)).ToString();
                lnAB2.Text = Math.Log(double.Parse(TextBox_resultA2.Text)).ToString();
                lnBA2.Text = Math.Log(double.Parse(TextBox_resultB2.Text)).ToString();
            }
            catch 
            {
                MessageBox.Show("error");
            }
        }

        private void Btntest3_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED3.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                     case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");
                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu3.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA3.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB3.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA3.Text = (float.Parse(TextBox_DevA3.Text) / float.Parse(TextBox_DevB3.Text)).ToString();
                TextBox_resultB3.Text = (float.Parse(TextBox_DevB3.Text) / float.Parse(TextBox_DevA3.Text)).ToString();
                lnAB3.Text = Math.Log(double.Parse(TextBox_resultA3.Text)).ToString();
                lnBA3.Text = Math.Log(double.Parse(TextBox_resultB3.Text)).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btntest4_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED4.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                   case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");
                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu4.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA4.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB4.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA4.Text = (float.Parse(TextBox_DevA4.Text) / float.Parse(TextBox_DevB4.Text)).ToString();
                TextBox_resultB4.Text = (float.Parse(TextBox_DevB4.Text) / float.Parse(TextBox_DevA4.Text)).ToString();
                lnAB4.Text = Math.Log(double.Parse(TextBox_resultA4.Text)).ToString();
                lnBA4.Text = Math.Log(double.Parse(TextBox_resultB4.Text)).ToString();
            }
            catch 
            {
                MessageBox.Show("error");
            }
        }

        private void Btntest5_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED5.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                    case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");
                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu5.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA5.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB5.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA5.Text = (float.Parse(TextBox_DevA5.Text) / float.Parse(TextBox_DevB5.Text)).ToString();
                TextBox_resultB5.Text = (float.Parse(TextBox_DevB5.Text) / float.Parse(TextBox_DevA5.Text)).ToString();
                lnAB5.Text = Math.Log(double.Parse(TextBox_resultA5.Text)).ToString();
                lnBA5.Text = Math.Log(double.Parse(TextBox_resultB5.Text)).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btntest6_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("Analysis.xml");
                int i = ComBox_LED6.SelectedIndex;
                string ComBox_selectvalue = "";
                switch (i)
                {
                    case 0:
                        {
                            ComBox_selectvalue = "LED1";
                            break;
                        }
                    case 1:
                        {
                            ComBox_selectvalue = "LED2";
                            break;
                        }
                    case 2:
                        {
                            ComBox_selectvalue = "LED1-ALED1";
                            break;
                        }
                    case 3:
                        {
                            ComBox_selectvalue = "LED2-ALED2";
                            break;
                        }
                }
                XmlElement rootElem = xmlDoc.DocumentElement;
                XmlNodeList DevList = rootElem.GetElementsByTagName("Dev");
                foreach (XmlNode node in DevList)
                {
                    if (((XmlElement)node).GetAttribute("nongdu") == TextBox_nongdu6.Text)
                    {
                        string Select = ((XmlElement)node).GetAttribute("Select");
                        if (ComBox_selectvalue == Select)
                        {
                            if (((XmlElement)node).GetAttribute("DevName") == "探测器A")
                            {
                                TextBox_DevA6.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                            else
                            {
                                TextBox_DevB6.Text = ((XmlElement)node).GetAttribute("Ave");
                            }
                        }
                    }
                }
                TextBox_resultA6.Text = (float.Parse(TextBox_DevA6.Text) / float.Parse(TextBox_DevB6.Text)).ToString();
                TextBox_resultB6.Text = (float.Parse(TextBox_DevB6.Text) / float.Parse(TextBox_DevA6.Text)).ToString();
                lnAB6.Text = Math.Log(double.Parse(TextBox_resultA6.Text)).ToString();
                lnBA6.Text = Math.Log(double.Parse(TextBox_resultB6.Text)).ToString();
            }
            catch 
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Add_Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(TextBox_k1.Text.Length<5 || TextBox_k2.Text.Length<3 || TextBox_c1.Text.Length<5 || TextBox_c2.Text.Length<3){
                    MessageBox.Show("位数错误");
                    return;
                }
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] Data_save = new Byte[21];
                Data_save[0] = 0x13; Data_save[1] = 0x12; Data_save[19] = 0x15; Data_save[20] = 0x0D;
                if (ComBox_Add_Sel.SelectedIndex == 0)
                {
                    Data_save[18] = 0x30;
                }
                else if (ComBox_Add_Sel.SelectedIndex == 1)
                {
                    Data_save[18] = 0x31;
                }
                else if (ComBox_Add_Sel.SelectedIndex == 2)
                {
                    Data_save[18] = 0x32;
                }
                else if (ComBox_Add_Sel.SelectedIndex == 3)
                {
                    Data_save[18] = 0x33;
                }
                string str = TextBox_k1.Text;
                if (str[0] == '+')
                {
                    Data_save[2] = 0x2B;
                }
                else if (str[0] == '-')
                {
                    Data_save[2] = 0x2D;
                }
                Data_save[3] = (Byte)chan(str[1]); Data_save[4] = (Byte)chan(str[2]); Data_save[5] = (Byte)chan(str[3]); Data_save[6] = (Byte)chan(str[4]);
                str = TextBox_k2.Text;
                Data_save[7] = (Byte)chan(str[0]); Data_save[8] = (Byte)chan(str[1]); Data_save[9] = (Byte)chan(str[2]);
                str = TextBox_c1.Text;
                if (str[0] == '+')
                {
                    Data_save[10] = 0x2B;
                }
                else if (str[0] == '-')
                {
                    Data_save[10] = 0x2D;
                }
                Data_save[11] = (Byte)chan(str[1]); Data_save[12] = (Byte)chan(str[2]); Data_save[13] = (Byte)chan(str[3]); Data_save[14] = (Byte)chan(str[4]);
                str = TextBox_c2.Text;
                Data_save[15] = (Byte)chan(str[0]); Data_save[16] = (Byte)chan(str[1]); Data_save[17] = (Byte)chan(str[2]);
              
                if (!port.IsOpen)
                {
                        port.Open();
                        port.Write(Data_save, 0, 21);
                        Thread.Sleep(30);
                        port.Close();
                        MessageBox.Show("success");
                        string path = "Config.xml";
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(path);
                        XmlNode deviceNode, kNode, cNode, fNode;
                        XmlNodeList projectList;
                        kNode = xmlDoc.CreateElement("k");
                        kNode.InnerText = TextBox_k1.Text + "." + TextBox_k2.Text;
                        cNode = xmlDoc.CreateElement("c");
                        cNode.InnerText = TextBox_c1.Text + "." + TextBox_c2.Text;
                        fNode = xmlDoc.CreateElement("f");
                        if (ComBox_Add_Sel.SelectedIndex == 0)
                        {
                            fNode.InnerText = "A/B";
                        }
                        else
                        {
                            fNode.InnerText = "B/A";
                        }
                        deviceNode = xmlDoc.SelectSingleNode("仪器编号");
                        projectList = xmlDoc.GetElementsByTagName("检测项目");
                        foreach (XmlNode node in projectList)
                        {
                            if (((XmlElement)node).GetAttribute("Name") == ComBox_yuansu.SelectedValue.ToString())
                            {
                                node.AppendChild(kNode);
                                node.AppendChild(cNode);
                                node.AppendChild(fNode);
                            }
                        }
                        xmlDoc.Save(path);
                    }
            }
            catch
            {
                MessageBox.Show("error");

            }
        }
     
        private void btn_EVM_Defaults_Click(object sender, RoutedEventArgs e)
        {
            Byte[] t1 = new Byte[10], t2 = new Byte[10], t3 = new Byte[10],
               t4 = new Byte[10], t5 = new Byte[10], t6 = new Byte[10],
               t7 = new Byte[10], t8 = new Byte[10], t9 = new Byte[10],
               t10 = new Byte[10], t11 = new Byte[10], t12 = new Byte[10],
               t13 = new Byte[10], t14 = new Byte[10], t15 = new Byte[10],
               t16 = new Byte[10], t17 = new Byte[10], t18 = new Byte[10],
               t19 = new Byte[10], t20 = new Byte[10], t21 = new Byte[10],
               t22 = new Byte[10], t23 = new Byte[10], t24 = new Byte[10],
               t25 = new Byte[10], t26 = new Byte[10], t27 = new Byte[10],
               t28 = new Byte[10], t29 = new Byte[10], t30 = new Byte[10],
               t31 = new Byte[10], t32 = new Byte[10], t33 = new Byte[10],
               t34 = new Byte[10], t35 = new Byte[10], t36 = new Byte[10],
               t37 = new Byte[10], t38 = new Byte[10], t39 = new Byte[10],
               t40 = new Byte[10], t41 = new Byte[10], t42 = new Byte[10],
               t43 = new Byte[10], t44 = new Byte[10];
               
            t1[0]=0x02;    t1[1]=0x30;  t1[2]=0x30;     t1[3]=0x30;     t1[4]=0x30;     t1[5]=0x30;     t1[6]=0x30;     t1[7]=0x30;     t1[8]=0x30; t1[9]=0x0D;
            t2[0]=0x02;    t2[1]=0x30;  t2[2]=0x31;     t2[3]=0x30;     t2[4]=0x30;     t2[5]=0x31;     t2[6]=0x37;     t2[7]=0x43;     t2[8]=0x30; t2[9]=0x0D;
            t3[0]=0x02;    t3[1]=0x30;  t3[2]=0x32;     t3[3]=0x30;     t3[4]=0x30;     t3[5]=0x31;     t3[6]=0x46;     t3[7]=0x33;     t3[8]=0x45; t3[9]=0x0D;
            t4[0]=0x02;    t4[1]=0x30;	t4[2]=0x33; 	t4[3]=0x30;  	t4[4]=0x30;	    t4[5]=0x31;  	t4[6]=0x37;	    t4[7]=0x37;  	t4[8]=0x30; t4[9]=0x0D;
            t5[0]=0x02;    t5[1]=0x30;  t5[2]=0x34;  	t5[3]=0x30;	    t5[4]=0x30;	    t5[5]=0x31;	    t5[6]=0x46;	    t5[7]=0x33;   	t5[8]=0x46; t5[9]=0x0D;
            t6[0]=0x02;    t6[1]=0x30;  t6[2]=0x35;	    t6[3]=0x30;	    t6[4]=0x30;	    t6[5]=0x30;  	t6[6]=0x30;	    t6[7]=0x35;	    t6[8]=0x30; t6[9]=0x0D;
            t7[0]=0x02;    t7[1]=0x30;  t7[2]=0x36;  	t7[3]=0x30;	    t7[4]=0x30;	    t7[5]=0x30;	    t7[6]=0x37;	    t7[7]=0x43;	    t7[8]=0x45; t7[9]=0x0D;
            t8[0]=0x02;    t8[1]=0x30;	t8[2]=0x37;   	t8[3]=0x30;	    t8[4]=0x30;	    t8[5]=0x30;	    t8[6]=0x38;	    t8[7]=0x32;	    t8[8]=0x30; t8[9]=0x0D;
            t9[0]=0x02;    t9[1]=0x30;	t9[2]=0x38;	    t9[3]=0x30;	    t9[4]=0x30;	    t9[5]=0x30;	    t9[6]=0x46;	    t9[7]=0x39;	    t9[8]=0x45; t9[9]=0x0D;
            t10[0]=0x02;   t10[1]=0x30; t10[2]=0x39;	t10[3]=0x30;	t10[4]=0x30;	t10[5]=0x30;	t10[6]=0x37;	t10[7]=0x44;	t10[8]=0x30;t10[9]=0x0D;
            t11[0]=0x02;   t11[1]=0x30; t11[2]=0x41;	t11[3]=0x30;	t11[4]=0x30;	t11[5]=0x30;	t11[6]=0x46;	t11[7]=0x39;	t11[8]=0x46;t11[9]=0x0D;
            t12[0]=0x02;   t12[1]=0x30;	t12[2]=0x42;	t12[3]=0x30;	t12[4]=0x30;	t12[5]=0x30;	t12[6]=0x46;	t12[7]=0x46;	t12[8]=0x30;t12[9]=0x0D;
            t13[0]=0x02;   t13[1]=0x30;	t13[2]=0x43;	t13[3]=0x30;	t13[4]=0x30;	t13[5]=0x31;	t13[6]=0x37;	t13[7]=0x36;	t13[8]=0x45;t13[9]=0x0D;
            t14[0]=0x02;   t14[1]=0x30;	t14[2]=0x44;	t14[3]=0x30;	t14[4]=0x30;	t14[5]=0x30;	t14[6]=0x30;	t14[7]=0x30;	t14[8]=0x36;t14[9]=0x0D;
            t15[0]=0x02;   t15[1]=0x30;	t15[2]=0x45;	t15[3]=0x30;	t15[4]=0x30;	t15[5]=0x30;	t15[6]=0x37;	t15[7]=0x43;	t15[8]=0x46;t15[9]=0x0D;
            t16[0]=0x02;   t16[1]=0x30;	t16[2]=0x46;	t16[3]=0x30;	t16[4]=0x30;	t16[5]=0x30;	t16[6]=0x37;	t16[7]=0x44;	t16[8]=0x36;t16[9]=0x0D;
            t17[0]=0x02;   t17[1]=0x31;	t17[2]=0x30;	t17[3]=0x30;	t17[4]=0x30;	t17[5]=0x30;	t17[6]=0x46;	t17[7]=0x39;	t17[8]=0x46;t17[9]=0x0D;
            t18[0]=0x02;   t18[1]=0x31; t18[2]=0x31;	t18[3]=0x30;	t18[4]=0x30;	t18[5]=0x30;	t18[6]=0x46;	t18[7]=0x41;	t18[8]=0x36;t18[9]=0x0D;
            t19[0]=0x02;   t19[1]=0x31;	t19[2]=0x32;	t19[3]=0x30;	t19[4]=0x30;	t19[5]=0x31;	t19[6]=0x37;	t19[7]=0x36;	t19[8]=0x46;t19[9]=0x0D;
            t20[0]=0x02;   t20[1]=0x31;	t20[2]=0x33;	t20[3]=0x30;	t20[4]=0x30;	t20[5]=0x31;	t20[6]=0x37;	t20[7]=0x37;	t20[8]=0x36;t20[9]=0x0D;
            t21[0]=0x02;   t21[1]=0x31;	t21[2]=0x34;	t21[3]=0x30;	t21[4]=0x30;	t21[5]=0x31;	t21[6]=0x46;	t21[7]=0x33;	t21[8]=0x46;t21[9]=0x0D;
            t22[0]=0x02;   t22[1]=0x31;	t22[2]=0x35;	t22[3]=0x30;	t22[4]=0x30;	t22[5]=0x30;	t22[6]=0x30;	t22[7]=0x30;	t22[8]=0x30;t22[9]=0x0D;
            t23[0]=0x02;   t23[1]=0x31;	t23[2]=0x36;	t23[3]=0x30;	t23[4]=0x30;	t23[5]=0x30;	t23[6]=0x30;	t23[7]=0x30;	t23[8]=0x35;t23[9]=0x0D;
            t24[0]=0x02;   t24[1]=0x31;	t24[2]=0x37;	t24[3]=0x30;	t24[4]=0x30;	t24[5]=0x30;	t24[6]=0x37;	t24[7]=0x44;	t24[8]=0x30;t24[9]=0x0D;
            t25[0]=0x02;   t25[1]=0x31;	t25[2]=0x38;	t25[3]=0x30;	t25[4]=0x30;	t25[5]=0x30;	t25[6]=0x37;	t25[7]=0x44;	t25[8]=0x35;t25[9]=0x0D;
            t26[0]=0x02;   t26[1]=0x31;	t26[2]=0x39;	t26[3]=0x30;	t26[4]=0x30;	t26[5]=0x30;	t26[6]=0x46;	t26[7]=0x41;	t26[8]=0x30;t26[9]=0x0D;
            t27[0]=0x02;   t27[1]=0x31;	t27[2]=0x41;	t27[3]=0x30;	t27[4]=0x30;	t27[5]=0x30;	t27[6]=0x46;    t27[7]=0x41;	t27[8]=0x35;t27[9]=0x0D;
            t28[0]=0x02;   t28[1]=0x31;	t28[2]=0x42;	t28[3]=0x30;	t28[4]=0x30;	t28[5]=0x31;	t28[6]=0x37;	t28[7]=0x37;	t28[8]=0x30;t28[9]=0x0D;
            t29[0]=0x02;   t29[1]=0x31;	t29[2]=0x43;	t29[3]=0x30;	t29[4]=0x30;	t29[5]=0x31;	t29[6]=0x37;	t29[7]=0x37;	t29[8]=0x35;t29[9]=0x0D;
            t30[0]=0x02;   t30[1]=0x31; t30[2]=0x44;	t30[3]=0x30;	t30[4]=0x30;	t30[5]=0x31;	t30[6]=0x46;	t30[7]=0x33;	t30[8]=0x46;t30[9]=0x0D;
            t31[0]=0x02;   t31[1]=0x31;	t31[2]=0x45;	t31[3]=0x30;	t31[4]=0x30;	t31[5]=0x30;	t31[6]=0x31;	t31[7]=0x30;	t31[8]=0x31;t31[9]=0x0D;
            t32[0]=0x02;   t32[1]=0x31;	t32[2]=0x46;	t32[3]=0x30;	t32[4]=0x30;	t32[5]=0x30;	t32[6]=0x30;	t32[7]=0x30;	t32[8]=0x30;t32[9]=0x0D;
            t33[0]=0x02;   t33[1]=0x32;	t33[2]=0x30;	t33[3]=0x30;	t33[4]=0x30;	t33[5]=0x30;	t33[6]=0x30;	t33[7]=0x30;	t33[8]=0x30;t33[9]=0x0D;
            t34[0]=0x02;   t34[1]=0x32;	t34[2]=0x31;	t34[3]=0x30;	t34[4]=0x30;	t34[5]=0x30;	t34[6]=0x30;	t34[7]=0x30;	t34[8]=0x32;t34[9]=0x0D;
            t35[0]=0x02;   t35[1]=0x32;	t35[2]=0x32;	t35[3]=0x30;	t35[4]=0x31;	t35[5]=0x31;	t35[6]=0x34;	t35[7]=0x31;	t35[8]=0x34;t35[9]=0x0D;
            t36[0]=0x02;   t36[1]=0x32;	t36[2]=0x33;	t36[3]=0x30;	t36[4]=0x30;	t36[5]=0x30;	t36[6]=0x30;	t36[7]=0x30;	t36[8]=0x30;t36[9]=0x0D;
            t37[0]=0x02;   t37[1]=0x32;	t37[2]=0x34;	t37[3]=0x30;	t37[4]=0x30;	t37[5]=0x30;    t37[6]=0x30;	t37[7]=0x30;	t37[8]=0x30;t37[9]=0x0D;
            t38[0]=0x02;   t38[1]=0x32;	t38[2]=0x35;	t38[3]=0x30;	t38[4]=0x30;	t38[5]=0x30;	t38[6]=0x30;	t38[7]=0x30;	t38[8]=0x30;t38[9]=0x0D;
            t39[0]=0x02;   t39[1]=0x32;	t39[2]=0x36;	t39[3]=0x30;	t39[4]=0x30;	t39[5]=0x30;	t39[6]=0x30;	t39[7]=0x30;	t39[8]=0x30;t39[9]=0x0D;
            t40[0]=0x02;   t40[1]=0x32;	t40[2]=0x37;	t40[3]=0x32;	t40[4]=0x38;	t40[5]=0x30;	t40[6]=0x30;	t40[7]=0x38;	t40[8]=0x30;t40[9]=0x0D;
            t41[0]=0x02;   t41[1]=0x32;	t41[2]=0x38;	t41[3]=0x34;	t41[4]=0x30;	t41[5]=0x30;	t41[6]=0x33;	t41[7]=0x30;	t41[8]=0x30;t41[9]=0x0D;
            t42[0]=0x02;   t42[1]=0x33;	t42[2]=0x31;	t42[3]=0x30;	t42[4]=0x30;	t42[5]=0x30;	t42[6]=0x30;	t42[7]=0x30;	t42[8]=0x30;t42[9]=0x0D;
            t43[0]=0x02;   t43[1]=0x33;	t43[2]=0x32;	t43[3]=0x30;	t43[4]=0x30;	t43[5]=0x30;	t43[6]=0x30;	t43[7]=0x30;	t43[8]=0x30;t43[9]=0x0D;
            t44[0]=0x02;   t44[1]=0x33;	t44[2]=0x33;	t44[3]=0x30;	t44[4]=0x30;	t44[5]=0x30;	t44[6]=0x30;	t44[7]=0x30;	t44[8]=0x30;t44[9]=0x0D;

            try { 
            SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
            port.WriteBufferSize=1000;
            if (!port.IsOpen)
            {
                port.Open();
                port.Write(t1, 0, 10);
                Thread.Sleep(150);
                port.Write(t2, 0, 10);
                Thread.Sleep(150);

                port.Write(t3, 0, 10);
                Thread.Sleep(50);
                port.Write(t4, 0, 10);
                Thread.Sleep(50);
                port.Write(t5, 0, 10);
                Thread.Sleep(50);
                port.Write(t6, 0, 10);
                Thread.Sleep(50);
                port.Write(t7, 0, 10);
                Thread.Sleep(50);
                port.Write(t8, 0, 10);
                Thread.Sleep(50);
                port.Write(t9, 0, 10);
                Thread.Sleep(50);
                port.Write(t10, 0, 10);
                Thread.Sleep(50);

                port.Write(t11, 0, 10);
                Thread.Sleep(50);
                port.Write(t12, 0, 10);
                Thread.Sleep(50);
                port.Write(t13, 0, 10);
                Thread.Sleep(50);
                port.Write(t14, 0, 10);
                Thread.Sleep(50);
                port.Write(t15, 0, 10);
                Thread.Sleep(50);
                port.Write(t16, 0, 10);
                Thread.Sleep(50);
                port.Write(t17, 0, 10);
                Thread.Sleep(50);
                port.Write(t18, 0, 10);
                Thread.Sleep(50);
                port.Write(t19, 0, 10);
                Thread.Sleep(50);
                port.Write(t20, 0, 10);
                Thread.Sleep(50);
                port.Write(t21, 0, 10);
                Thread.Sleep(50);
                port.Write(t22, 0, 10);
                Thread.Sleep(50);
                port.Write(t23, 0, 10);
                Thread.Sleep(50);
                port.Write(t24, 0, 10);
                Thread.Sleep(50);
                port.Write(t25, 0, 10);
                Thread.Sleep(50);
                port.Write(t26, 0, 10);
                Thread.Sleep(50);
                port.Write(t27, 0, 10);
                Thread.Sleep(50);
                port.Write(t28, 0, 10);
                Thread.Sleep(50);
                port.Write(t29, 0, 10);
                Thread.Sleep(50);
                port.Write(t30, 0, 10);
                Thread.Sleep(50);
                port.Write(t31, 0, 10);
                Thread.Sleep(50);
                port.Write(t32, 0, 10);
                Thread.Sleep(50);
                port.Write(t33, 0, 10);
                Thread.Sleep(50);
                port.Write(t34, 0, 10);
                Thread.Sleep(50);
                port.Write(t35, 0, 10);
                Thread.Sleep(50);
                port.Write(t36, 0, 10);
                Thread.Sleep(50);
                port.Write(t37, 0, 10);
                Thread.Sleep(50);
                port.Write(t38, 0, 10);
                Thread.Sleep(50);
                port.Write(t39, 0, 10);
                Thread.Sleep(50);
                port.Write(t40, 0, 10);
                Thread.Sleep(50);
                port.Write(t41, 0, 10);
                Thread.Sleep(50);
                port.Write(t42, 0, 10);
                Thread.Sleep(50);
                port.Write(t43, 0, 10);
                Thread.Sleep(50);
                port.Write(t44, 0, 10);
                Thread.Sleep(50);

                port.Close();

                MessageBox.Show("success");
            }
                
            
            }
            catch
            {
                MessageBox.Show("error");
            }

        }
        private int chan(int s)
        {
            //return Convert.ToInt32(s.ToString("X"));
            if (s > 47 && s < 58)
            {
                return 0x30 + s - 48;
            }
            else if(s>64 && s<91)
            {
                return 0x41 + s - 65;
            }
            else if (s > 96 && s < 123)
            {
                return 0x61 + s - 97;
            }
            else if(s==45)
            {
                return 0x2D;
            }
            else
            {
                return 0;
            }
        }
        private void Btn_writeDev_Click(object sender, RoutedEventArgs e)
        {
            if (Instrument_number.Text.Length < 13)
            {
                MessageBox.Show("仪器编号应等于13");
                return;
            }
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] bt = new Byte[17];
                bt[0] = 0x13; bt[1] = 0x14;
                bt[15] = 0x17; bt[16] = 0x0D;
              string str = Instrument_number.Text;
        
                if (!port.IsOpen)
                {
                        port.Open();
                        for (int i = 0; i < 13; i++)
                        {
                          //  bt[i + 2] = (Byte)chan(str[i]);
                           bt[i+2]=(Byte)chan(str[i]);
                           if (chan(str[i]) == 0)
                           {
                               MessageBox.Show("error");
                               return;
                           }
                        }
                      
                      
                        port.Write(bt, 0, 17);
                        Thread.Sleep(30);
                        port.Close();
                        MessageBox.Show("success");
                }
            }
            catch (Exception E1)
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_readDev_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] bt = new Byte[4],bt_receive=new Byte[17];
                bt[0] = 0x13; bt[1] = 0x16; bt[2] = 0x19; bt[3] = 0x0D;
                char c;
                        port.Open();
                        port.Write(bt, 0, 4);
                        Thread.Sleep(30);
                        port.Read(bt_receive, 0, 17);
                        string str = "";
                      
                            for (int i = 2; i <= 14; i++)
                            {
                                c = (char)bt_receive[i];
                               // c = (char)Convert.ToInt32(bt_receive[i].ToString(), 16);
                                str += c.ToString();
                            }
                     
                     //   MessageBox.Show(str1);
                            port.Close();
                            MessageBox.Show(str);
            }
            catch
            {
                MessageBox.Show("error");

            }
        }

        private void a_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] bt = new Byte[10];
                bt[0] = 0x02; bt[1] = 0x33; bt[2] = 0x31; bt[3] = 0x30;
                bt[4] = 0x30; bt[5] = 0x38; bt[6] = 0x30; bt[7] = 0x30;
                bt[8] = 0x30; bt[9] = 0x0D;
                if (!port.IsOpen)
                {
                   
                        port.Open();
                        port.Write(bt, 0, 10);
                        Thread.Sleep(30);
                        port.Close();
                        MessageBox.Show("success");
                }
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void b_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] bt = new Byte[4];
                bt[0] = 0x13; bt[1] = 0x13; bt[2] = 0x16; bt[3] = 0x0D;
                if (!port.IsOpen)
                {
                        port.Open();
                        port.Write(bt, 0, 4);
                        Thread.Sleep(30);
                        port.Close();
                        MessageBox.Show("success");
                }
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void c_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                Byte[] bt = new Byte[4];
                bt[0] = 0x13; bt[1] = 0x11; bt[2] = 0x14; bt[3] = 0x0D;
                if (!port.IsOpen)
                {
                        port.Open();
                        port.Write(bt, 0, 4);
                        Thread.Sleep(30);
                        port.Close();
                        MessageBox.Show("success");
                }
            }
            catch
            {
                MessageBox.Show("error");

            }
        }

        private void Btn_Reset_Click(object sender, RoutedEventArgs e)
        {
            Byte[] bt = new Byte[10];
            bt[0] = 0x02;
            bt[1] = bt[2] = bt[3] = bt[4] = bt[5] = bt[6] = bt[7] = 0x30;
            bt[9] = 0x0D;
            switch (Q)
            {
                case 0:
                    {
                        bt[8] = 0x32;
                        Q = 2;
                        break;
                    }
                case 1:
                    {
                        bt[8] = 0x36;
                        Q = 3;
                        break;
                    }
                case 2:
                    {
                        bt[8] = 0x32;
                        break;
                    }
                case 3:
                    {
                        bt[8] = 0x36;
                        break;
                    }
            }
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.Open();
                port.Write(bt, 0, 10);
                Thread.Sleep(30);
                port.Close();
                MessageBox.Show("success");
            }
            catch
            {
                MessageBox.Show("error");

            }
        }

        private void Btn_calcute_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double a1 = double.Parse(X1.Text);
                double b1 = double.Parse(Y1.Text);
                double a2 = double.Parse(X2.Text);
                double b2 = double.Parse(Y2.Text);
                double d = (b1 - b2) / (a1 - a2);
                double f = (b1 * a2 - a1 * b2) / (a2 - a1);
                k.Text = d.ToString();
                C.Text = f.ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Dev_Select_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                SerialPort port = new SerialPort(CommNum, IntBdr, Parity.None, 8);
                port.Open();
                Byte[] bt = new Byte[10];
                if (Dev_Select.SelectedIndex == 0)
                {
                    bt[0] = 0x02; bt[1] = 0x33; bt[2] = 0x34;
                    bt[3] = 0x30; bt[4] = 0x30; bt[5] = 0x30;
                    bt[6] = 0x30; bt[7] = 0x30; bt[8] = 0x30;
                    bt[9] = 0x0D;
                }
                else
                {
                    bt[0] = 0x02; bt[1] = 0x33; bt[2] = 0x34;
                    bt[3] = 0x30; bt[4] = 0x30; bt[5] = 0x30;
                    bt[6] = 0x30; bt[7] = 0x30; bt[8] = 0x31;
                    bt[9] = 0x0D;
                }
                port.Write(bt, 0, 10);
                port.Close();
                MessageBox.Show("success");
                
            }
            catch
            {
            }
        }

        private void Btn_Test6_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start6.Text), i;
                int end = Convert.ToInt32(TextBox_end6.Text);

                if (ComBox_Select6.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select6.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select6.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select6.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select6.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select6.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver6.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test7_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start7.Text), i;
                int end = Convert.ToInt32(TextBox_end7.Text);

                if (ComBox_Select7.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select7.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select7.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select7.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select7.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select7.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver7.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test8_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start8.Text), i;
                int end = Convert.ToInt32(TextBox_end8.Text);

                if (ComBox_Select8.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select8.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select8.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select8.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select8.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select8.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver8.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Btn_Test9_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sum = 0, ave;
                int start = Convert.ToInt32(TextBox_start9.Text), i;
                int end = Convert.ToInt32(TextBox_end9.Text);

                if (ComBox_Select9.SelectedIndex == 0)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[0, i];
                    }
                }
                else if (ComBox_Select9.SelectedIndex == 1)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[1, i];
                    }
                }
                else if (ComBox_Select9.SelectedIndex == 2)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[2, i];
                    }
                }
                else if (ComBox_Select9.SelectedIndex == 3)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[3, i];
                    }
                }
                else if (ComBox_Select9.SelectedIndex == 4)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[4, i];
                    }
                }
                else if (ComBox_Select9.SelectedIndex == 5)
                {
                    for (i = start; i <= end; i++)
                    {
                        sum += Datas[5, i];
                    }
                }
                ave = sum / (end - start + 1);
                TextBox_Aver9.Text = ((decimal)ave).ToString();
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

        private void Save_kc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FileStream fs = new FileStream("xishu.ini", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                string str = "";
                str = K1.Text + " " + C1.Text + " " + K2.Text + " " + C2.Text + " " + K3.Text + " " + C3.Text + " " + K4.Text + " " + C4.Text + " ";
                sw.Write(str);
                sw.Close();
                fs.Close();
                xK1 = double.Parse(K1.Text);
                xC1 = double.Parse(C1.Text);
                xK2 = double.Parse(K2.Text);
                xC2 = double.Parse(C2.Text);
                xK3 = double.Parse(K3.Text);
                xC3 = double.Parse(C3.Text);
                xK4 = double.Parse(K4.Text);
                xC4 = double.Parse(C4.Text);
                MessageBox.Show("save success");
            }
            catch
            {
                MessageBox.Show("error");
            }
        }

      
       
     
       
    }
}
