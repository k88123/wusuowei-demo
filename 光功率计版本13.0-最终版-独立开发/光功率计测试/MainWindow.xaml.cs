using HidLibrary;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace 光功率计测试
{
    public partial class MainWindow : Window
    {
        private const int VendorId = 0x0483;
        private const int ProductId = 0x5750;

        private const byte REPORT_ID_SENSOR_DATA = 0x01;
        private const byte REPORT_ID_TABLE_DATA = 0x03;

        private const byte CMD_REQUEST_POWER_TABLE = 0x04;
        private const byte CMD_REQUEST_TEMP_TABLE = 0x05;
        private const byte CMD_SAVE_POWER_TABLE = 0x14;
        private const byte CMD_SAVE_TEMP_TABLE = 0x15;
        private const byte CMD_RESTORE_DEFAULTS = 0x20;

        private const byte DATA_START_POWER_TABLE = 0xF0;
        private const byte DATA_START_TEMP_POINTS = 0xF1;
        private const byte DATA_START_COMP_FACTORS = 0xF2;
        private const byte DATA_TEMP_POINTS = 0xFE;
        private const byte DATA_SEGMENT = 0xFD;
        private const byte DATA_END = 0xFF;

        private const byte ACK_SUCCESS = 0x01;
        private const byte ACK_READY = 0x02;
        private const byte ACK_ERROR = 0xFF;

        public class CompensationMatrixData : INotifyPropertyChanged
        {
            private int _segmentIndex;
            private string _adcRange;
            private ushort _adcMin;
            private ushort _adcMax;

            private float _temp10;
            private float _temp12;
            private float _temp14;
            private float _temp16;
            private float _temp18;
            private float _temp20;
            private float _temp22;
            private float _temp24;
            private float _temp25;
            private float _temp26;
            private float _temp27;
            private float _temp28;
            private float _temp29;
            private float _temp30;
            private float _temp31;
            private float _temp32;
            private float _temp33;
            private float _temp34;
            private float _temp35;
            private float _temp36;
            private float _temp37;
            private float _temp38;
            private float _temp39;
            private float _temp40;

            public float Temp10
            {
                get => _temp10;
                set
                {
                    _temp10 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp12
            {
                get => _temp12;
                set
                {
                    _temp12 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp14
            {
                get => _temp14;
                set
                {
                    _temp14 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp16
            {
                get => _temp16;
                set
                {
                    _temp16 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp18
            {
                get => _temp18;
                set
                {
                    _temp18 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp20
            {
                get => _temp20;
                set
                {
                    _temp20 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp22
            {
                get => _temp22;
                set
                {
                    _temp22 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp24
            {
                get => _temp24;
                set
                {
                    _temp24 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp25
            {
                get => _temp25;
                set
                {
                    _temp25 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp26
            {
                get => _temp26;
                set
                {
                    _temp26 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp27
            {
                get => _temp27;
                set
                {
                    _temp27 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp28
            {
                get => _temp28;
                set
                {
                    _temp28 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp29
            {
                get => _temp29;
                set
                {
                    _temp29 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp30
            {
                get => _temp30;
                set
                {
                    _temp30 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp31
            {
                get => _temp31;
                set
                {
                    _temp31 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp32
            {
                get => _temp32;
                set
                {
                    _temp32 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp33
            {
                get => _temp33;
                set
                {
                    _temp33 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp34
            {
                get => _temp34;
                set
                {
                    _temp34 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp35
            {
                get => _temp35;
                set
                {
                    _temp35 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp36
            {
                get => _temp36;
                set
                {
                    _temp36 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp37
            {
                get => _temp37;
                set
                {
                    _temp37 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp38
            {
                get => _temp38;
                set
                {
                    _temp38 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp39
            {
                get => _temp39;
                set
                {
                    _temp39 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public float Temp40
            {
                get => _temp40;
                set
                {
                    _temp40 = ValidateFactorValue(value);
                    OnPropertyChanged();
                }
            }

            public int SegmentIndex
            {
                get => _segmentIndex;
                set { _segmentIndex = value; OnPropertyChanged(); }
            }

            public string AdcRange
            {
                get => _adcRange;
                set { _adcRange = value; OnPropertyChanged(); }
            }

            public ushort AdcMin
            {
                get => _adcMin;
                set { _adcMin = value; OnPropertyChanged(); }
            }

            public ushort AdcMax
            {
                get => _adcMax;
                set { _adcMax = value; OnPropertyChanged(); }
            }

            private float ValidateFactorValue(float value)
            {
                if (value < 0.0f) return 0.0f;
                if (value > 10.0f) return 10.0f;
                if (float.IsNaN(value) || float.IsInfinity(value)) return 1.0f;
                return value;
            }

            public float GetFactorByTemperatureIndex(int tempIndex)
            {
                switch (tempIndex)
                {
                    case 0: return Temp10;
                    case 1: return Temp12;
                    case 2: return Temp14;
                    case 3: return Temp16;
                    case 4: return Temp18;
                    case 5: return Temp20;
                    case 6: return Temp22;
                    case 7: return Temp24;
                    case 8: return Temp24;
                    case 9: return Temp25;
                    case 10: return Temp26;
                    case 11: return Temp27;
                    case 12: return Temp28;
                    case 13: return Temp29;
                    case 14: return Temp30;
                    case 15: return Temp31;
                    case 16: return Temp32;
                    case 17: return Temp33;
                    case 18: return Temp34;
                    case 19: return Temp35;
                    case 20: return Temp36;
                    case 21: return Temp37;
                    case 22: return Temp38;
                    case 23: return Temp39;
                    case 24: return Temp40;
                    default: return 1.0f;
                }
            }

            public float GetFactorByTemperature(float temperature)
            {
                int tempKey = (int)Math.Round(temperature);

                switch (tempKey)
                {
                    case 10: return Temp10;
                    case 12: return Temp12;
                    case 14: return Temp14;
                    case 16: return Temp16;
                    case 18: return Temp18;
                    case 20: return Temp20;
                    case 22: return Temp22;
                    case 24: return Temp24;
                    case 25: return Temp25;
                    case 26: return Temp26;
                    case 27: return Temp27;
                    case 28: return Temp28;
                    case 29: return Temp29;
                    case 30: return Temp30;
                    case 31: return Temp31;
                    case 32: return Temp32;
                    case 33: return Temp33;
                    case 34: return Temp34;
                    case 35: return Temp35;
                    case 36: return Temp36;
                    case 37: return Temp37;
                    case 38: return Temp38;
                    case 39: return Temp39;
                    case 40: return Temp40;
                    default: return 1.0f;
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class CompensationFactorData : INotifyPropertyChanged
        {
            private int _segmentIndex;
            private int _temperatureIndex;
            private float _temperature;
            private ushort _adcMin;
            private ushort _adcMax;
            private float _factorValue;

            public int SegmentIndex
            {
                get => _segmentIndex;
                set { _segmentIndex = value; OnPropertyChanged(); }
            }

            public int TemperatureIndex
            {
                get => _temperatureIndex;
                set { _temperatureIndex = value; OnPropertyChanged(); }
            }

            public float Temperature
            {
                get => _temperature;
                set { _temperature = value; OnPropertyChanged(); }
            }

            public ushort AdcMin
            {
                get => _adcMin;
                set { _adcMin = value; OnPropertyChanged(); }
            }

            public ushort AdcMax
            {
                get => _adcMax;
                set { _adcMax = value; OnPropertyChanged(); }
            }

            public float FactorValue
            {
                get => _factorValue;
                set { _factorValue = value; OnPropertyChanged(); }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class CalibrationDataPoint : INotifyPropertyChanged
        {
            private byte _index;
            private ushort _adcValue;
            private float _powerValue;
            private float _temperature;
            private float _compensationFactor;

            public byte Index
            {
                get => _index;
                set { _index = value; OnPropertyChanged(); }
            }

            public ushort AdcValue
            {
                get => _adcValue;
                set { _adcValue = value; OnPropertyChanged(); }
            }

            public float PowerValue
            {
                get => _powerValue;
                set { _powerValue = value; OnPropertyChanged(); }
            }

            public float Temperature
            {
                get => _temperature;
                set { _temperature = value; OnPropertyChanged(); }
            }

            public float CompensationFactor
            {
                get => _compensationFactor;
                set { _compensationFactor = value; OnPropertyChanged(); }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class PowerCalibrationData
        {
            public int Index { get; set; }
            public int AdcValue { get; set; }
            public double PowerValue { get; set; }
        }

        public class TempCompensationSegment
        {
            public byte SegmentIndex { get; set; }
            public ushort AdcMin { get; set; }
            public ushort AdcMax { get; set; }
            public Dictionary<byte, float> CompensationFactors { get; set; } = new Dictionary<byte, float>();
        }

        private ObservableCollection<CalibrationDataPoint> powerCalibrationData;
        private ObservableCollection<CalibrationDataPoint> tempPointsData;
        private ObservableCollection<CompensationFactorData> compensationFactorData;
        private ObservableCollection<CompensationMatrixData> compensationMatrixData;

        private ObservableCollection<CalibrationDataPoint> receivingPowerData;
        private ObservableCollection<CalibrationDataPoint> receivingTempPoints;
        private List<TempCompensationSegment> receivingTempSegments = new List<TempCompensationSegment>();


        private bool isReceivingPowerTable = false;
        private bool isReceivingTempTable = false;
        private string lastDataHash = "";

        private HidDevice device;
        private bool isReading = false;

        private Dictionary<int, List<int>> adcDataPerDegree = new Dictionary<int, List<int>>();

        private List<AdcDataPoint> adcDataLog = new List<AdcDataPoint>();
        private bool isLogging = false;
        private DateTime loggingStartTime;

        private int currentAdcValue = 0;
        private int maxAdcValue = int.MinValue;
        private int minAdcValue = int.MaxValue;
        private double avgAdcValue = 0;
        private long sumAdcValue = 0;
        private int dataPointCount = 0;
        private double currentLightValue = 0;
        private double maxLightValue = double.MinValue;
        private double minLightValue = double.MaxValue;
        private double sumLightValue = 0;
        private int currentTempAdc = 0;
        private int maxTempAdc = int.MinValue;
        private int minTempAdc = int.MaxValue;
        private long sumTempAdc = 0;
        private double currentTempValue = 0;
        private double maxTempValue = double.MinValue;
        private double minTempValue = double.MaxValue;
        private double avgTempValue = 0;
        private double sumTempValue = 0;
        private double lastRecordedTemperature = double.MinValue;
        private bool isRising = true;

        private int upperAdcThreshold = 32767;
        private int lowerAdcThreshold = 0;
        private bool isAlertMuted = false;

        private DataFormat selectedDataFormat = DataFormat.LittleEndian16Bit;

        private bool autoLoggingEnabled = false;
        private bool hasAutoLogged = false;
        private const double AutoLoggingTemperature = 29;
        private const double AutoStopTemperature = 31;
        private const double AutoCoolingStopTemperature = 28;
        private bool _hasShownStopDialog = false;


        public MainWindow()
        {
            InitializeComponent();
            dataFormatComboBox.SelectedIndex = 0;
            autoLoggingCheckBox.Checked += AutoLoggingCheckBox_Checked;
            autoLoggingCheckBox.Unchecked += AutoLoggingCheckBox_Unchecked;

            InitializeCalibrationDataGrids();
        }
        private void InitializeCalibrationDataGrids()
        {
            //数据容器
            powerCalibrationData = new ObservableCollection<CalibrationDataPoint>();
            tempPointsData = new ObservableCollection<CalibrationDataPoint>();
            compensationFactorData = new ObservableCollection<CompensationFactorData>();
            compensationMatrixData = new ObservableCollection<CompensationMatrixData>();

            receivingPowerData = new ObservableCollection<CalibrationDataPoint>();
            receivingTempPoints = new ObservableCollection<CalibrationDataPoint>();

            //建立契约：告诉表格，你的数据来源就是这个集合
            powerCalibrationGrid.ItemsSource = powerCalibrationData;
            compensationMatrixGrid.ItemsSource = compensationMatrixData;

            UpdateCompensationGridColumns();
            UpdateCompensationMatrixGridColumns();
        }
        private void UpdateCompensationMatrixGridColumns()
        {
            compensationMatrixGrid.Columns.Clear();

            compensationMatrixGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "分段",
                Binding = new Binding("SegmentIndex"),
                Width = 60,
                IsReadOnly = true
            });

            compensationMatrixGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "ADC范围",
                Binding = new Binding("AdcRange"),
                Width = 100,
                IsReadOnly = true
            });

            var temperatures = new[] { 10, 12, 14, 16, 18, 20, 22, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40 };

            foreach (var temp in temperatures)
            {
                string propertyName = $"Temp{temp}";
                compensationMatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = $"{temp}℃",
                    Binding = new Binding(propertyName) { StringFormat = "F4" },
                    Width = 70,
                    IsReadOnly = false
                });
            }
        }

        private void UpdateCompensationGridColumns()
        {
        }
        private void MergeTemperatureAndCompensationData()
        {
            compensationFactorData.Clear();
            compensationMatrixData.Clear();

            var temperatureMapping = new Dictionary<int, float>();
            foreach (var tempPoint in receivingTempPoints)
            {
                temperatureMapping[tempPoint.Index] = tempPoint.Temperature;
            }

            foreach (var segment in receivingTempSegments)
            {
                var matrixData = new CompensationMatrixData
                {
                    SegmentIndex = segment.SegmentIndex,
                    AdcMin = segment.AdcMin,
                    AdcMax = segment.AdcMax,
                    AdcRange = $"{segment.AdcMin}-{segment.AdcMax}"
                };

                foreach (var factor in segment.CompensationFactors)
                {
                    var compensationData = new CompensationFactorData
                    {
                        SegmentIndex = segment.SegmentIndex,
                        TemperatureIndex = factor.Key,
                        AdcMin = segment.AdcMin,
                        AdcMax = segment.AdcMax,
                        FactorValue = factor.Value
                    };

                    if (temperatureMapping.TryGetValue(factor.Key, out float temperature))
                    {
                        compensationData.Temperature = temperature;

                        SetMatrixTemperatureValue(matrixData, temperature, factor.Value);
                    }

                    compensationFactorData.Add(compensationData);
                }

                compensationMatrixData.Add(matrixData);
            }

            var sortedData = compensationFactorData
                .OrderBy(c => c.SegmentIndex)
                .ThenBy(c => c.TemperatureIndex)
                .ToList();

            compensationFactorData.Clear();
            foreach (var item in sortedData)
            {
                compensationFactorData.Add(item);
            }

            var sortedMatrixData = compensationMatrixData
                .OrderBy(m => m.SegmentIndex)
                .ToList();

            compensationMatrixData.Clear();
            foreach (var item in sortedMatrixData)
            {
                compensationMatrixData.Add(item);
            }
        }

        private void SetMatrixTemperatureValue(CompensationMatrixData matrixData, float temperature, float factorValue)
        {
            int tempKey = (int)Math.Round(temperature);

            switch (tempKey)
            {
                case 10: matrixData.Temp10 = factorValue; break;
                case 12: matrixData.Temp12 = factorValue; break;
                case 14: matrixData.Temp14 = factorValue; break;
                case 16: matrixData.Temp16 = factorValue; break;
                case 18: matrixData.Temp18 = factorValue; break;
                case 20: matrixData.Temp20 = factorValue; break;
                case 22: matrixData.Temp22 = factorValue; break;
                case 24: matrixData.Temp24 = factorValue; break;
                case 25: matrixData.Temp25 = factorValue; break;
                case 26: matrixData.Temp26 = factorValue; break;
                case 27: matrixData.Temp27 = factorValue; break;
                case 28: matrixData.Temp28 = factorValue; break;
                case 29: matrixData.Temp29 = factorValue; break;
                case 30: matrixData.Temp30 = factorValue; break;
                case 31: matrixData.Temp31 = factorValue; break;
                case 32: matrixData.Temp32 = factorValue; break;
                case 33: matrixData.Temp33 = factorValue; break;
                case 34: matrixData.Temp34 = factorValue; break;
                case 35: matrixData.Temp35 = factorValue; break;
                case 36: matrixData.Temp36 = factorValue; break;
                case 37: matrixData.Temp37 = factorValue; break;
                case 38: matrixData.Temp38 = factorValue; break;
                case 39: matrixData.Temp39 = factorValue; break;
                case 40: matrixData.Temp40 = factorValue; break;
            }
        }

        private void RefreshCompensationMatrixGrid()
        {
            Dispatcher.Invoke(() =>
            {
                compensationMatrixGrid.Items.Refresh();
            });
        }

        private void CompensationMatrixGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var column = e.Column as DataGridTextColumn;
                if (column != null && column.Header != null)
                {
                    var matrixData = e.Row.Item as CompensationMatrixData;
                    if (matrixData != null)
                    {
                        try
                        {
                            var textBox = e.EditingElement as TextBox;
                            if (textBox != null)
                            {
                                string newValue = textBox.Text.Trim();

                                if (float.TryParse(newValue, out float factorValue))
                                {
                                    UpdateCompensationFactorFromMatrix(matrixData, column.Header.ToString(), factorValue);

                                }
                                else
                                {
                                    MessageBox.Show("请输入有效的数字", "输入错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                                    e.Cancel = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("更新数据时出错", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            e.Cancel = true;
                        }
                    }
                }
            }
        }

        private void UpdateCompensationFactorFromMatrix(CompensationMatrixData matrixData, string temperatureHeader, float factorValue)
        {
            try
            {
                string tempStr = temperatureHeader.ToString().Replace("℃", "");
                if (int.TryParse(tempStr, out int temperature))
                {
                    var tempPoint = tempPointsData.FirstOrDefault(t => Math.Abs(t.Temperature - temperature) < 0.1);
                    if (tempPoint != null)
                    {
                        byte temperatureIndex = tempPoint.Index;

                        var segment = receivingTempSegments.FirstOrDefault(s => s.SegmentIndex == matrixData.SegmentIndex);
                        if (segment != null)
                        {
                            if (segment.CompensationFactors.ContainsKey(temperatureIndex))
                            {
                                segment.CompensationFactors[temperatureIndex] = factorValue;
                            }
                            else
                            {
                                segment.CompensationFactors.Add(temperatureIndex, factorValue);
                            }

                            var factorData = compensationFactorData.FirstOrDefault(f =>
                                f.SegmentIndex == matrixData.SegmentIndex && f.TemperatureIndex == temperatureIndex);
                            if (factorData != null)
                            {
                                factorData.FactorValue = factorValue;
                            }

                        }
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (device != null && device.IsConnected)
            {
                DisconnectDevice();
            }
            else ConnectToDevice();

        }

        private void ConnectToDevice()
        {
            var devices = HidDevices.Enumerate(VendorId, ProductId).ToList();
            if (devices.Count == 0)
            {
                MessageBox.Show($"未找到VID: 0x{VendorId:X4}, PID: 0x{ProductId:X4}的设备");
                return;
            }
            device = devices.First();
            try
            {
                if (device.IsOpen) device.CloseDevice();
                device.OpenDevice();
                if (device.IsConnected)
                {
                    statusText.Text = "设备已连接";
                    statusText.Foreground = Brushes.Green;
                    connectButton.Content = "断开连接";
                    isReading = true;
                    StartAsyncRead();
                }
                else statusText.Text = "设备连接失败";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"连接设备时出错: {ex.Message}");
            }
        }

        private void StartAsyncRead()
        {
            if (isReading && device != null && device.IsConnected)
            {
                // 关键点在这里：把 OnReportReceived 函数作为“回调(Callback)”塞给了底层的 HidLibrary
                device.ReadReport(OnReportReceived);
            }
        }

        private void OnReportReceived(HidReport report)
        {
            try
            {
                // 因为当前执行这个函数的线程，是 .NET 线程池里随机抓来的后台线程
                // 所以绝对不能直接改 UI，必须立刻用 Dispatcher.Invoke 抛给主线程去处理 ProcessReport
                if (report != null) Dispatcher.Invoke(() => ProcessReport(report));

                // 处理完当前包之后，再次发起异步监听，形成类似 while(true) 的接收环路
                if (isReading && device != null && device.IsConnected) device.ReadReport(OnReportReceived);
            }
            catch (Exception)
            {
                if (isReading && device != null && device.IsConnected)
                {
                    Thread.Sleep(1000);
                    device.ReadReport(OnReportReceived);
                }
            }
        }

        private void ProcessReport(HidReport report)
        {
            if (report?.Data == null || report.Data.Length < 2) return;

            byte reportId = report.ReportId;
            byte[] data = report.Data;


            switch (reportId)
            {
                case REPORT_ID_SENSOR_DATA:
                    ProcessSensorData(report.Data);
                    break;

                case REPORT_ID_TABLE_DATA:
                    ProcessTableData(data);
                    break;

                default:
                    break;
            }
        }

        private void ProcessSensorData(byte[] data)
        {
            try
            {
                int opticalValue1 = ParseAdcValue(data, 0, selectedDataFormat);
                int opticalAdcValue = ParseAdcValue(data, 2, selectedDataFormat);
                int temperature = ParseAdcValue(data, 4, selectedDataFormat);
                int tempAdcValue = ParseAdcValue(data, 6, selectedDataFormat);

                double powerValue = opticalValue1 / 100.0;
                double tempCelsius = temperature / 1000.0;

                currentAdcText.Text = opticalAdcValue.ToString();
                currentPowerText.Text = powerValue.ToString("F2");
                currentTempAdcText.Text = tempAdcValue.ToString();
                currentTempText.Text = $"{tempCelsius:F1}℃";

                currentTemperatureText.Text = tempCelsius.ToString("F1");

                UpdateStatistics(opticalAdcValue, powerValue, tempAdcValue, tempCelsius);

                if (isLogging)
                {
                    var dataPoint = new AdcDataPoint
                    {
                        Timestamp = DateTime.Now,
                        AdcValue = opticalAdcValue,
                        PowerValue = powerValue,
                        TemperatureRaw = tempAdcValue,
                        Temperature = tempCelsius
                    };

                    adcDataLog.Add(dataPoint);

                    RecordDataPoint(opticalAdcValue, tempCelsius);

                    UpdateAdcChart(opticalAdcValue, tempCelsius);
                }

                CheckAutoLoggingConditions(tempCelsius);

                CheckAdcThresholds(opticalAdcValue);

            }
            catch (Exception)
            {
            }
        }
        private int ParseAdcValue(byte[] data, int offset, DataFormat format)
        {
            switch (format)
            {
                case DataFormat.LittleEndian16Bit: return data.Length >= offset + 2 ? (data[offset + 1] << 8) | data[offset] : 0;
                case DataFormat.BigEndian16Bit: return data.Length >= offset + 2 ? (data[offset] << 8) | data[offset + 1] : 0;
                case DataFormat.Adc12Bit: return data.Length >= offset + 2 ? ((data[offset + 1] & 0x0F) << 8) | data[offset] : 0;
                case DataFormat.Adc10Bit: return data.Length >= offset + 2 ? ((data[offset + 1] & 0x03) << 8) | data[offset] : 0;
                case DataFormat.RawBytes: return data.Length >= offset + 1 ? data[offset] : 0;
                default: return 0;
            }
        }

        private void CheckAutoLoggingConditions(double temperature)
        {
            if (isLogging && !_hasShownStopDialog)
            {
                if (isRising && temperature >= AutoStopTemperature)
                {
                    StopAutoLoggingDueToTemperature();
                    loggingStatusText.Text = $"达到最高温度{AutoStopTemperature}℃，开始降温记录...";
                    return;
                }
            }

            if (autoLoggingEnabled && !hasAutoLogged && temperature >= AutoLoggingTemperature)
            {
                StartAutoLogging();
                hasAutoLogged = true;
            }
        }

        private void StopAutoLoggingDueToCooling()
        {
            isLogging = false;
            startLoggingButton.IsEnabled = true;
            stopLoggingButton.IsEnabled = false;

            string stopMessage = $"温度降至{AutoLoggingTemperature - 1}℃以下，自动停止记录。共记录 {adcDataPerDegree.Values.Sum(list => list.Count)} 个数据点";
            loggingStatusText.Text = stopMessage;
            autoLoggingStatusText.Text = $"已自动停止 (降温完成)";
            autoLoggingStatusText.Foreground = Brushes.Red;
            hasAutoLogged = false;

            AutoExportData(stopMessage);

            if (!_hasShownStopDialog)
            {
                _hasShownStopDialog = true;
                MessageBox.Show(stopMessage, "自动停止", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ProcessTableData(byte[] data)
        {
            if (data.Length < 1) return;

            byte command = data[0];
            string dataHash = BitConverter.ToString(data);

            if (dataHash == lastDataHash) return;
            lastDataHash = dataHash;

            switch (command)
            {
                case DATA_START_POWER_TABLE:
                    StartReceivingPowerTable();
                    break;

                case DATA_START_TEMP_POINTS:
                    StartReceivingTempPoints();
                    break;

                case DATA_START_COMP_FACTORS:
                    StartReceivingCompFactors();
                    break;

                case DATA_TEMP_POINTS:
                    ParseTemperaturePointsPacket(data);
                    break;

                case DATA_SEGMENT:
                    ParseSegmentDataPacket(data);
                    break;

                case DATA_END:
                    FinishReceivingTables();
                    break;

                default:
                    if (command <= 0x18 && isReceivingPowerTable)
                    {
                        ParsePowerDataPacket(data);
                    }
                    break;
            }
        }

        private void StartReceivingPowerTable()
        {
            receivingPowerData.Clear();
            isReceivingPowerTable = true;
            isReceivingTempTable = false;

            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = "正在接收光功率表数据...";
                calibrationStatusText.Foreground = Brushes.Orange;
            });
        }

        private void ParsePowerDataPacket(byte[] data)
        {
            try
            {
                if (data.Length < 2) return;

                byte startIndex = data[0];

                for (int j = 0; j < 5; j++)
                {
                    int dataIndex = 1 + j * 6;
                    if (dataIndex + 5 < data.Length)
                    {
                        byte pointIndex = (byte)(startIndex + j);
                        if (pointIndex >= 24) continue;

                        ushort adcValue = (ushort)((data[dataIndex + 1] << 8) | data[dataIndex]);

                        float powerValue = BitConverter.ToSingle(data, dataIndex + 2);

                        UpdatePowerDataPoint(pointIndex, adcValue, powerValue);
                    }
                }

                receivingPowerData = SortObservableCollection(receivingPowerData);
                RefreshPowerCalibrationGrid();
            }
            catch (Exception)
            {
            }
        }

        private void UpdatePowerDataPoint(byte index, ushort adcValue, float powerValue)
        {
            var point = new CalibrationDataPoint
            {
                Index = index,
                AdcValue = adcValue,
                PowerValue = powerValue
            };

            var existingPoint = receivingPowerData.FirstOrDefault(p => p.Index == index);
            if (existingPoint != null)
            {
                receivingPowerData.Remove(existingPoint);
            }
            receivingPowerData.Add(point);

        }

        private void StartReceivingTempPoints()
        {
            receivingTempPoints.Clear();
            isReceivingTempTable = true;

            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = "正在接收温度点数据...";
                calibrationStatusText.Foreground = Brushes.Orange;
            });
        }

        private void StartReceivingCompFactors()
        {
            receivingTempSegments.Clear();
        }

        private void ParseTemperaturePointsPacket(byte[] data)
        {
            try
            {
                if (data.Length < 3) return;

                byte startIndex = data[1];

                for (int i = 0; i < 6; i++)
                {
                    int dataIndex = 2 + i * 4;
                    if (dataIndex + 3 < data.Length)
                    {
                        byte tempIndex = (byte)(startIndex + i);
                        if (tempIndex >= 24) continue;

                        float temperature = BitConverter.ToSingle(data, dataIndex);

                        if (temperature != 0)
                        {
                            UpdateTempDataPoint(tempIndex, temperature);
                        }
                    }
                }

                receivingTempPoints = SortObservableCollection(receivingTempPoints);
            }
            catch (Exception)
            {
            }
        }

        private ObservableCollection<CalibrationDataPoint> SortObservableCollection(ObservableCollection<CalibrationDataPoint> collection)
        {
            var sortedList = collection.OrderBy(p => p.Index).ToList();
            var sortedCollection = new ObservableCollection<CalibrationDataPoint>();

            foreach (var item in sortedList)
            {
                sortedCollection.Add(item);
            }

            return sortedCollection;
        }

        private void ParseSegmentDataPacket(byte[] data)
        {
            try
            {
                if (data.Length < 3) return;

                byte segIndex = data[1];
                byte subPacketIndex = data[2];

                if (subPacketIndex == 0)
                {
                    ParseSegmentFirstSubPacket(data, segIndex);
                }
                else if (subPacketIndex == 1)
                {
                    ParseSegmentSecondSubPacket(data, segIndex);
                }
            }
            catch (Exception)
            {
            }
        }

        private void ParseSegmentFirstSubPacket(byte[] data, byte segIndex)
        {
            if (data.Length >= 7)
            {
                ushort adcMin = (ushort)((data[4] << 8) | data[3]);
                ushort adcMax = (ushort)((data[6] << 8) | data[5]);

                var segment = GetOrCreateSegment(segIndex);
                segment.AdcMin = adcMin;
                segment.AdcMax = adcMax;

                for (int i = 0; i < 12; i++)
                {
                    int dataIndex = 7 + i * 4;
                    if (dataIndex + 3 < data.Length)
                    {
                        float factor = BitConverter.ToSingle(data, dataIndex);
                        segment.CompensationFactors[(byte)i] = factor;
                    }
                }
            }
        }

        private void ParseSegmentSecondSubPacket(byte[] data, byte segIndex)
        {
            for (int i = 0; i < 12; i++)
            {
                int dataIndex = 3 + i * 4;
                if (dataIndex + 3 < data.Length)
                {
                    float factor = BitConverter.ToSingle(data, dataIndex);
                    byte factorIndex = (byte)(12 + i);

                    var segment = GetOrCreateSegment(segIndex);
                    segment.CompensationFactors[factorIndex] = factor;
                }
            }
        }

        private TempCompensationSegment GetOrCreateSegment(byte segIndex)
        {
            var segment = receivingTempSegments.FirstOrDefault(s => s.SegmentIndex == segIndex);
            if (segment == null)
            {
                segment = new TempCompensationSegment { SegmentIndex = segIndex };
                receivingTempSegments.Add(segment);
            }
            return segment;
        }

        private void UpdateTempDataPoint(byte index, float temperature)
        {
            var point = new CalibrationDataPoint
            {
                Index = index,
                Temperature = temperature
            };

            var existingPoint = receivingTempPoints.FirstOrDefault(p => p.Index == index);
            if (existingPoint != null)
            {
                receivingTempPoints.Remove(existingPoint);
            }
            receivingTempPoints.Add(point);

        }

        private void FinishReceivingTables()
        {
            if (isReceivingPowerTable)
            {
                powerCalibrationData.Clear();
                foreach (var item in receivingPowerData.OrderBy(p => p.Index))
                {
                    powerCalibrationData.Add(item);
                }

                RefreshPowerCalibrationGrid();

                Dispatcher.Invoke(() =>
                {
                    ShowSuccess($"光功率表接收完成，共 {powerCalibrationData.Count} 个点");
                });
            }

            if (isReceivingTempTable)
            {
                tempPointsData.Clear();
                foreach (var item in receivingTempPoints.OrderBy(p => p.Index))
                {
                    tempPointsData.Add(item);
                }

                MergeTemperatureAndCompensationData();

                RefreshCompensationMatrixGrid();

                Dispatcher.Invoke(() =>
                {
                    ShowSuccess($"温度表接收完成，温度点: {tempPointsData.Count}, 补偿因子: {compensationFactorData.Count}");
                });
            }

            isReceivingPowerTable = false;
            isReceivingTempTable = false;
        }

        private void LoadCalibrationBtn_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*",
                Title = "选择校准表文件"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    LoadCalibrationFromFile(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载文件失败: {ex.Message}");
                }
            }
        }

        private void ExportCalibrationBtn_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                Title = "导出校准表文件"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExportCalibrationToFile(saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出文件失败: {ex.Message}");
                }
            }
        }

        private void LoadCalibrationFromFile(string filePath)
        {
        }

        private void ExportCalibrationToFile(string filePath)
        {
        }

        private void RecordDataPoint(int adcValue, double temperature)
        {
            if (Math.Abs(temperature - lastRecordedTemperature) >= 0.1)
            {
                if (lastRecordedTemperature != double.MinValue)
                {
                    isRising = temperature > lastRecordedTemperature;
                }

                int tempKey = (int)Math.Floor(temperature);
                if (!adcDataPerDegree.ContainsKey(tempKey))
                {
                    adcDataPerDegree[tempKey] = new List<int>();
                }
                adcDataPerDegree[tempKey].Add(adcValue);
                lastRecordedTemperature = temperature;

                string direction = isRising ? "↑升温" : "↓降温";
                loggingStatusText.Text = $"记录中...{direction} 当前温度 {temperature:F1}°C, 已记录 {adcDataPerDegree[tempKey].Count} 点";

                var dataPoint = new AdcDataPoint
                {
                    Timestamp = DateTime.Now,
                    AdcValue = adcValue,
                    PowerValue = currentLightValue,
                    TemperatureRaw = currentTempAdc,
                    Temperature = temperature,
                    IsRising = isRising
                };
                adcDataLog.Add(dataPoint);

            }
        }


        private void StartAutoLogging()
        {
            isLogging = true;
            loggingStartTime = DateTime.Now;
            adcDataPerDegree.Clear();
            adcDataLog.Clear();
            lastRecordedTemperature = double.MinValue;
            hasAutoLogged = true;

            startLoggingButton.IsEnabled = false;
            stopLoggingButton.IsEnabled = true;
            loggingStatusText.Text = $"自动记录中... 温度范围: {AutoLoggingTemperature}℃ → {AutoStopTemperature}℃";
            autoLoggingStatusText.Text = $"自动记录中 ({AutoLoggingTemperature}℃ → {AutoStopTemperature}℃)";
            autoLoggingStatusText.Foreground = Brushes.Green;
        }

        private void AutoLoggingCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            autoLoggingEnabled = true;
            hasAutoLogged = false;
            startLoggingButton.IsEnabled = false;
            stopLoggingButton.IsEnabled = false;
            autoLoggingStatusText.Text = "等待温度达到28℃...";
            autoLoggingStatusText.Foreground = Brushes.Orange;
        }

        private void AutoLoggingCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            autoLoggingEnabled = false;
            startLoggingButton.IsEnabled = !isLogging;
            stopLoggingButton.IsEnabled = isLogging;
            autoLoggingStatusText.Text = "自动记录已关闭";
            autoLoggingStatusText.Foreground = Brushes.Gray;
        }

        private void StartLoggingButton_Click(object sender, RoutedEventArgs e)
        {
            if (autoLoggingEnabled)
            {
                MessageBox.Show("自动记录已启用，请先关闭自动记录功能", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                autoLoggingCheckBox.IsChecked = false;
                return;
            }

            isLogging = true;
            loggingStartTime = DateTime.Now;
            adcDataPerDegree.Clear();
            adcDataLog.Clear();
            lastRecordedTemperature = double.MinValue;

            startLoggingButton.IsEnabled = false;
            stopLoggingButton.IsEnabled = true;
            loggingStatusText.Text = $"记录中... 开始时间: {loggingStartTime:HH:mm:ss}";
        }

        private void StopLoggingButton_Click(object sender, RoutedEventArgs e)
        {
            isLogging = false;
            hasAutoLogged = false;

            startLoggingButton.IsEnabled = true;
            stopLoggingButton.IsEnabled = false;
            loggingStatusText.Text = $"已停止记录. 共记录了 {adcDataPerDegree.Values.Sum(list => list.Count)} 个数据点";

            if (autoLoggingEnabled)
            {
                autoLoggingStatusText.Text = $"等待温度重新达到 {AutoLoggingTemperature}℃...";
                autoLoggingStatusText.Foreground = Brushes.Orange;
            }
            else
            {
                autoLoggingStatusText.Text = "手动记录已停止";
                autoLoggingStatusText.Foreground = Brushes.Gray;
            }
        }

        private void DisconnectDevice()
        {
            isReading = false;
            isLogging = false;
            autoLoggingEnabled = false;
            hasAutoLogged = false;

            if (device != null)
            {
                device.CloseDevice();
                device = null;
            }
            statusText.Text = "设备未连接";
            statusText.Foreground = Brushes.Red;
            connectButton.Content = "连接设备";
            rawDataText.Text = "原始光强ADC数据: --";
            dataFormatText.Text = "数据格式: --";
            autoLoggingCheckBox.IsChecked = false;
            autoLoggingStatusText.Text = "自动记录状态";
            autoLoggingStatusText.Foreground = Brushes.Gray;
            ResetStatistics();
        }

        private void StopAutoLoggingDueToTemperature()
        {
            isLogging = false;
            startLoggingButton.IsEnabled = true;
            stopLoggingButton.IsEnabled = false;
            string stopMessage = $"温度达到{AutoStopTemperature}℃，自动停止记录。共记录 {adcDataPerDegree.Values.Sum(list => list.Count)} 个数据点";
            loggingStatusText.Text = stopMessage;
            autoLoggingStatusText.Text = $"已自动停止 (温度{AutoStopTemperature}℃)";
            autoLoggingStatusText.Foreground = Brushes.Red;
            hasAutoLogged = false;

            AutoExportData(stopMessage);

            if (!_hasShownStopDialog)
            {
                _hasShownStopDialog = true;
                MessageBox.Show(stopMessage, "自动停止", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private bool CheckDeviceConnection()
        {
            if (device == null || !device.IsConnected)
            {
                MessageBox.Show("设备未连接");
                return false;
            }
            return true;
        }
        private async void RequestPowerTableBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckDeviceConnection()) return;

            powerCalibrationData.Clear();
            receivingPowerData.Clear();
            lastDataHash = "";
            RefreshPowerCalibrationGrid();

            await SendCommandAsync(CMD_REQUEST_POWER_TABLE);

            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = "请求光功率校准表中...";
                calibrationStatusText.Foreground = Brushes.Orange;
            });

            _ = Task.Delay(5000).ContinueWith(t =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (powerCalibrationData.Count == 0)
                    {
                        ShowError("获取光功率表失败，请重试");
                    }
                });
            });
        }

        private async void RequestTempTableBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckDeviceConnection()) return;

            tempPointsData.Clear();
            compensationFactorData.Clear();
            receivingTempPoints.Clear();
            receivingTempSegments.Clear();
            lastDataHash = "";

            await SendCommandAsync(CMD_REQUEST_TEMP_TABLE);

            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = "请求温度补偿表中...";
                calibrationStatusText.Foreground = Brushes.Orange;
            });

            _ = Task.Delay(5000).ContinueWith(t =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (tempPointsData.Count == 0 && compensationFactorData.Count == 0)
                    {
                        ShowError("获取温度表失败，请重试");
                    }
                });
            });
        }


        private void RefreshPowerCalibrationGrid()
        {
            Dispatcher.Invoke(() =>
            {
                powerCalibrationGrid.Items.Refresh();
            });
        }


        private async void SavePowerTableBtn_Click(object sender, RoutedEventArgs e)
        {
            if (device == null || !device.IsConnected)
            {
                MessageBox.Show("设备未连接");
                return;
            }

            if (powerCalibrationData.Count == 0)
            {
                MessageBox.Show("没有光功率校准数据可保存");
                return;
            }

            ShowProgress("光功率表数据发送中...");

            try
            {
                await SendCommandAsync(CMD_SAVE_POWER_TABLE);
                await Task.Delay(100);

                foreach (var point in powerCalibrationData)
                {
                    await SendPowerDataPointAsync(point.Index, point.AdcValue, point.PowerValue);
                    await Task.Delay(50);
                }

                _ = Task.Delay(3000).ContinueWith(t =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        ShowSuccess("光功率表数据发送完成");
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存数据失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async Task SendPowerDataPointAsync(byte index, ushort adcValue, float powerValue)
        {
            await Task.Run(() =>
            {
                byte[] data = new byte[64];
                data[0] = 0x03;
                data[1] = index;

                data[2] = (byte)(adcValue & 0xFF);
                data[3] = (byte)((adcValue >> 8) & 0xFF);

                byte[] powerBytes = BitConverter.GetBytes(powerValue);
                Array.Copy(powerBytes, 0, data, 4, 4);

                var report = new HidReport(64, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));
                bool success = device.WriteReport(report);

            });
        }


        private async void SaveTempTableBtn_Click(object sender, RoutedEventArgs e)
        {
            if (device == null || !device.IsConnected)
            {
                MessageBox.Show("设备未连接");
                return;
            }

            if (tempPointsData.Count == 0)
            {
                MessageBox.Show("没有温度补偿数据可保存");
                return;
            }

            ShowProgress("温度表数据发送中...");

            try
            {
                await SendCommandAsync(CMD_SAVE_TEMP_TABLE);
                await Task.Delay(100);

                await SendAllTemperaturePointsAsync();
                await Task.Delay(50);

                await SendAllCompensationFactorsDirectAsync();
                await Task.Delay(50);

                _ = Task.Delay(5000).ContinueWith(t =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        ShowSuccess("温度表数据发送完成");
                    });
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存数据失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ShowProgress(string message)
        {
            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = $"⏳ {message}";
                calibrationStatusText.Foreground = new SolidColorBrush(Color.FromRgb(52, 152, 219));
            });
        }

        private void ShowSuccess(string message)
        {
            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = $"✅ {message}";
                calibrationStatusText.Foreground = new SolidColorBrush(Color.FromRgb(39, 174, 96));
            });
        }

        private void ShowError(string message)
        {
            Dispatcher.Invoke(() =>
            {
                calibrationStatusText.Text = $"❌ {message}";
                calibrationStatusText.Foreground = new SolidColorBrush(Color.FromRgb(231, 76, 60));
            });
        }

        private async Task SendAllTemperaturePointsAsync()
        {
            if (tempPointsData.Count == 0) return;

            var sortedTempPoints = tempPointsData.OrderBy(p => p.Index).ToList();

            for (byte i = 0; i < sortedTempPoints.Count; i++)
            {
                var point = sortedTempPoints[i];
                await SendTempPointDirectAsync(point.Index, point.Temperature);
                await Task.Delay(30);
            }
        }

        private async Task SendTempPointDirectAsync(byte index, float temperature)
        {
            await Task.Run(() =>
            {
                byte[] data = new byte[64];
                data[0] = REPORT_ID_TABLE_DATA;
                data[1] = DATA_TEMP_POINTS;
                data[2] = index;

                byte[] tempBytes = BitConverter.GetBytes(temperature);
                Array.Copy(tempBytes, 0, data, 3, 4);

                var report = new HidReport(64, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));
                bool success = device.WriteReport(report);

            });
        }


        private async Task SendAllCompensationFactorsDirectAsync()
        {
            if (compensationMatrixData.Count == 0) return;

            var sortedMatrixData = compensationMatrixData.OrderBy(m => m.SegmentIndex).ToList();

            for (byte segIndex = 0; segIndex < sortedMatrixData.Count; segIndex++)
            {
                var matrixData = sortedMatrixData[segIndex];

                await SendTempSegmentDirectAsync((byte)matrixData.SegmentIndex, matrixData.AdcMin, matrixData.AdcMax);
                await Task.Delay(30);

                var factors = GetCompensationFactorsFromMatrix(matrixData);

                await SendCompensationSubPacketDirectAsync((byte)matrixData.SegmentIndex, 0, factors.Take(12).ToArray());
                await Task.Delay(30);

                await SendCompensationSubPacketDirectAsync((byte)matrixData.SegmentIndex, 1, factors.Skip(12).Take(12).ToArray());
                await Task.Delay(30);
            }
        }

        private async Task SendTempSegmentDirectAsync(byte segIndex, ushort adcMin, ushort adcMax)
        {
            await Task.Run(() =>
            {
                byte[] data = new byte[64];
                data[0] = REPORT_ID_TABLE_DATA;
                data[1] = DATA_SEGMENT;
                data[2] = segIndex;

                data[3] = (byte)(adcMin & 0xFF);
                data[4] = (byte)((adcMin >> 8) & 0xFF);

                data[5] = (byte)(adcMax & 0xFF);
                data[6] = (byte)((adcMax >> 8) & 0xFF);

                var report = new HidReport(64, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));
                bool success = device.WriteReport(report);

            });
        }

        private float[] GetCompensationFactorsFromMatrix(CompensationMatrixData matrixData)
        {
            float[] factors = new float[24];

            var temperatureMapping = new Dictionary<int, float>
            {
                {0, 10}, {1, 12}, {2, 14}, {3, 16}, {4, 18}, {5, 20}, {6, 22}, {7, 24},
                {8, 25}, {9, 26}, {10, 27}, {11, 28}, {12, 29}, {13, 30}, {14, 31},
                {15, 32}, {16, 33}, {17, 34}, {18, 35}, {19, 36}, {20, 37}, {21, 38},
                {22, 39}, {23, 40}
            };

            foreach (var mapping in temperatureMapping)
            {
                factors[mapping.Key] = matrixData.GetFactorByTemperatureIndex(mapping.Key);
            }

            return factors;
        }

        private async Task SendCompensationSubPacketDirectAsync(byte segIndex, byte subPacketIndex, float[] factors)
        {
            await Task.Run(() =>
            {
                byte[] data = new byte[64];
                data[0] = REPORT_ID_TABLE_DATA;
                data[1] = DATA_SEGMENT;
                data[2] = segIndex;
                data[3] = subPacketIndex;

                var matrixData = compensationMatrixData.FirstOrDefault(m => m.SegmentIndex == segIndex);
                if (matrixData == null)
                {
                    return;
                }

                ushort adcMin = matrixData.AdcMin;
                ushort adcMax = matrixData.AdcMax;

                if (subPacketIndex == 0)
                {
                    data[4] = (byte)(adcMin & 0xFF);
                    data[5] = (byte)((adcMin >> 8) & 0xFF);

                    data[6] = (byte)(adcMax & 0xFF);
                    data[7] = (byte)((adcMax >> 8) & 0xFF);

                    for (int i = 0; i < 12; i++)
                    {
                        byte[] factorBytes = BitConverter.GetBytes(factors[i]);
                        Array.Copy(factorBytes, 0, data, 8 + i * 4, 4);
                    }
                }
                else
                {
                    for (int i = 0; i < 12; i++)
                    {
                        byte[] factorBytes = BitConverter.GetBytes(factors[i]);
                        Array.Copy(factorBytes, 0, data, 4 + i * 4, 4);
                    }
                }

                var report = new HidReport(64, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));

                bool success = device.WriteReport(report);

            });
        }

        private async Task SendCommandAsync(byte command)
        {
            // Task.Run 出马：向线程池申请一个后台线程，把大括号里的活儿扔给它干
            await Task.Run(() =>
            {
                byte[] data = new byte[64];
                data[0] = REPORT_ID_TABLE_DATA;
                data[1] = command;

                var report = new HidReport(64, new HidDeviceData(data, HidDeviceData.ReadStatus.Success));

                bool success = device.WriteReport(report);
            });
        }

        private void AutoExportData(string statusMessage)
        {
            if (adcDataLog.Count == 0 || adcDataPerDegree.Count == 0)
            {
                MessageBox.Show("自动停止时没有数据可导出。", "自动停止", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string fileNameBase = $"AutoExport_ADC_Data_{DateTime.Now:yyyyMMdd_HHmmss}";
                string directory = @"C:\Users\weiji\Desktop\温度补偿插值法";

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string rawFile = System.IO.Path.Combine(directory, fileNameBase + "_原始数据.csv");
                string avgFile = System.IO.Path.Combine(directory, fileNameBase + "_温度平均.csv");

                var sortedKeys = adcDataPerDegree.Keys.ToList();
                sortedKeys.Sort();

                if (sortedKeys.Count == 0) return;

                bool isRisingProcess = IsRisingProcess(adcDataLog);
                int baseTempKey;
                string processType;

                if (isRisingProcess)
                {
                    baseTempKey = sortedKeys.First();
                    processType = "升温";
                }
                else
                {
                    baseTempKey = sortedKeys.Last();
                    processType = "降温";
                }

                double baseAverageAdc = adcDataPerDegree[baseTempKey].Average();

                StringBuilder rawCsv = new StringBuilder();
                rawCsv.AppendLine("时间戳,光强ADC,光强值(mW),温度ADC,温度值(℃),补偿因子,方向,过程类型");

                double lastExportedTemperature = adcDataLog.First().Temperature - 1.0;
                const double temperatureInterval = 0.1;

                foreach (var dataPoint in adcDataLog)
                {
                    if (Math.Abs(dataPoint.Temperature - lastExportedTemperature) >= temperatureInterval)
                    {
                        double compensation = baseAverageAdc / dataPoint.AdcValue;
                        string direction = dataPoint.IsRising ? "升温" : "降温";

                        rawCsv.AppendLine(
                            $"{dataPoint.Timestamp:yyyy-MM-dd HH:mm:ss}," +
                            $"{dataPoint.AdcValue}," +
                            $"{dataPoint.PowerValue:F2}," +
                            $"{dataPoint.TemperatureRaw}," +
                            $"{dataPoint.Temperature:F1}," +
                            $"{compensation:F4}," +
                            $"{direction}," +
                            $"{processType}"
                        );
                        lastExportedTemperature = dataPoint.Temperature;
                    }
                }

                File.WriteAllText(rawFile, rawCsv.ToString(), Encoding.UTF8);

                StringBuilder avgCsv = new StringBuilder();
                avgCsv.AppendLine("温度(℃),采样点数,平均ADC,归一化值");
                avgCsv.AppendLine($"# 过程类型: {processType}, 基准温度: {baseTempKey}℃, 基准平均ADC: {baseAverageAdc:F2}");

                List<int> displayKeys;
                if (isRisingProcess)
                {
                    displayKeys = sortedKeys.OrderBy(x => x).ToList();
                }
                else
                {
                    displayKeys = sortedKeys.OrderByDescending(x => x).ToList();
                }

                foreach (var tempKey in displayKeys)
                {
                    var adcValuesForDegree = adcDataPerDegree[tempKey];
                    if (adcValuesForDegree != null && adcValuesForDegree.Count > 0)
                    {
                        int dataPointCount = adcValuesForDegree.Count;
                        double currentAverageAdc = adcValuesForDegree.Average();
                        double normalizedValue = baseAverageAdc / currentAverageAdc;

                        avgCsv.AppendLine($"{tempKey},{dataPointCount},{currentAverageAdc:F2},{normalizedValue:F4}");
                    }
                }

                File.WriteAllText(avgFile, avgCsv.ToString(), Encoding.UTF8);

                MessageBox.Show($"自动停止并成功导出数据：\n过程类型: {processType}\n基准温度: {baseTempKey}℃\n文件路径:\n{rawFile}\n{avgFile}",
                    "自动导出成功", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"自动导出数据时出错: {ex.Message}", "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsRisingProcess(List<AdcDataPoint> dataLog)
        {
            if (dataLog.Count < 2) return true;

            int sampleSize = Math.Max(1, dataLog.Count / 10);

            var startPoints = dataLog.Take(sampleSize);
            var endPoints = dataLog.Skip(dataLog.Count - sampleSize).Take(sampleSize);

            double startAvgTemp = startPoints.Average(p => p.Temperature);
            double endAvgTemp = endPoints.Average(p => p.Temperature);

            return endAvgTemp > startAvgTemp;
        }

        private void ExportDataButton_Click(object sender, RoutedEventArgs e)
        {
            if (adcDataLog.Count == 0 || adcDataPerDegree.Count == 0)
            {
                MessageBox.Show("没有数据可导出");
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                FileName = $"ADC_Data_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    StringBuilder rawCsv = new StringBuilder();
                    rawCsv.AppendLine("时间戳,光强ADC,光强值(mW),温度ADC,温度值(℃),补偿因子");

                    var sortedKeys = adcDataPerDegree.Keys.ToList();
                    sortedKeys.Sort();
                    bool isRisingProcess = IsRisingProcess(adcDataLog);
                    int baseTempKey;
                    string processType;
                    if (isRisingProcess)
                    {
                        baseTempKey = sortedKeys.First();
                        processType = "升温";
                    }
                    else
                    {
                        baseTempKey = sortedKeys.Last();
                        processType = "降温";
                    }
                    double baseAverageAdc = adcDataPerDegree[baseTempKey].Average();

                    double lastExportedTemperature = adcDataLog.First().Temperature - 1.0;
                    const double temperatureInterval = 0.1;

                    foreach (var dataPoint in adcDataLog)
                    {
                        if (Math.Abs(dataPoint.Temperature - lastExportedTemperature) >= temperatureInterval)
                        {
                            double compensation = baseAverageAdc / dataPoint.AdcValue;

                            rawCsv.AppendLine(
                                $"{dataPoint.Timestamp:yyyy-MM-dd HH:mm:ss}," +
                                $"{dataPoint.AdcValue}," +
                                $"{dataPoint.PowerValue:F2}," +
                                $"{dataPoint.TemperatureRaw}," +
                                $"{dataPoint.Temperature:F1}," +
                                $"{compensation:F4}"
                            );

                            lastExportedTemperature = dataPoint.Temperature;
                        }
                    }

                    string rawFile = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(saveFileDialog.FileName),
                        System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName) + "_原始数据.csv"
                    );
                    File.WriteAllText(rawFile, rawCsv.ToString(), Encoding.UTF8);

                    StringBuilder avgCsv = new StringBuilder();
                    avgCsv.AppendLine("温度(℃),采样点数,平均ADC,归一化值");
                    avgCsv.AppendLine($"# 基准温度: {baseTempKey}℃, 基准平均ADC: {baseAverageAdc:F2}");

                    List<int> displayKeys;
                    if (isRisingProcess)
                    {
                        displayKeys = sortedKeys.OrderBy(x => x).ToList();
                    }
                    else
                    {
                        displayKeys = sortedKeys.OrderByDescending(x => x).ToList();
                    }

                    foreach (var tempKey in sortedKeys)
                    {
                        var adcValuesForDegree = adcDataPerDegree[tempKey];
                        if (adcValuesForDegree != null && adcValuesForDegree.Count > 0)
                        {
                            int dataPointCount = adcValuesForDegree.Count;
                            double currentAverageAdc = adcValuesForDegree.Average();
                            double normalizedValue = baseAverageAdc / currentAverageAdc;

                            avgCsv.AppendLine($"{tempKey},{dataPointCount},{currentAverageAdc:F2},{normalizedValue:F4}");
                        }
                    }

                    string avgFile = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(saveFileDialog.FileName),
                        System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName) + "_温度平均.csv"
                    );
                    File.WriteAllText(avgFile, avgCsv.ToString(), Encoding.UTF8);

                    MessageBox.Show($"数据已导出：\n{rawFile}\n{avgFile}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出数据时出错: {ex.Message}");
                }
            }
        }

        private string GetDataFormatDescription(DataFormat format)
        {
            switch (format)
            {
                case DataFormat.LittleEndian16Bit: return "16位整数(小端)";
                case DataFormat.BigEndian16Bit: return "16位整数(大端)";
                case DataFormat.Adc12Bit: return "12位ADC值";
                case DataFormat.Adc10Bit: return "10位ADC值";
                case DataFormat.RawBytes: return "原始字节数据";
                default: return "未知格式";
            }
        }

        private void UpdateStatistics(int adcValue, double powerValue, int temperatureRaw, double temperature)
        {
            if (adcValue > maxAdcValue) { maxAdcValue = adcValue; maxAdcText.Text = maxAdcValue.ToString(); }
            if (adcValue < minAdcValue) { minAdcValue = adcValue; minAdcText.Text = minAdcValue.ToString(); }
            sumAdcValue += adcValue;
            dataPointCount++;
            avgAdcValue = (double)sumAdcValue / dataPointCount;
            avgAdcText.Text = avgAdcValue.ToString("F2");
            if (powerValue > maxLightValue) { maxLightValue = powerValue; maxPowerText.Text = maxLightValue.ToString("F2"); }
            if (powerValue < minLightValue) { minLightValue = powerValue; minPowerText.Text = minLightValue.ToString("F2"); }
            sumLightValue += powerValue;
            avgPowerText.Text = (sumLightValue / dataPointCount).ToString("F2");
            if (temperature > maxTempValue) { maxTempValue = temperature; maxTempText.Text = $"{maxTempValue:F1}℃"; }
            if (temperature < minTempValue) { minTempValue = temperature; minTempText.Text = $"{minTempValue:F1}℃"; }
            sumTempValue += temperature;
            avgTempValue = sumTempValue / dataPointCount;
            avgTempText.Text = $"{avgTempValue:F1}℃";
            if (temperatureRaw > maxTempAdc) { maxTempAdc = temperatureRaw; maxTempAdcText.Text = maxTempAdc.ToString(); }
            if (temperatureRaw < minTempAdc) { minTempAdc = temperatureRaw; minTempAdcText.Text = minTempAdc.ToString(); }
            sumTempAdc += temperatureRaw;
            avgTempAdcText.Text = (sumTempAdc / dataPointCount).ToString("F0");
        }

        private void UpdateAdcChart(int adcValue, double temperature)
        {
            adcChart.Children.Clear();
            if (adcDataLog.Count > 1)
            {
                double canvasWidth = adcChart.ActualWidth;
                double canvasHeight = adcChart.ActualHeight;
                if (canvasWidth == 0 || canvasHeight == 0) return;
                double minTemp = adcDataLog.Min(p => p.Temperature);
                double maxTemp = adcDataLog.Max(p => p.Temperature);
                double tempRange = Math.Max(maxTemp - minTemp, 1);
                int minAdc = adcDataLog.Min(p => p.AdcValue);
                int maxAdc = adcDataLog.Max(p => p.AdcValue);
                double adcRange = Math.Max(maxAdc - minAdc, 1);
                Polyline polyline = new Polyline { Stroke = Brushes.Blue, StrokeThickness = 2 };
                foreach (var point in adcDataLog)
                {
                    double x = (point.Temperature - minTemp) / tempRange * canvasWidth;
                    double y = canvasHeight - ((point.AdcValue - minAdc) / adcRange * canvasHeight);
                    polyline.Points.Add(new Point(x, y));
                }
                adcChart.Children.Add(polyline);
            }
        }

        private void CheckAdcThresholds(int adcValue)
        {
            if (isAlertMuted) return;
            if (adcValue > upperAdcThreshold || adcValue < lowerAdcThreshold)
            {
                dataStatusText.Text = "超出范围!";
                dataStatusText.Foreground = Brushes.Red;
                System.Media.SystemSounds.Exclamation.Play();
            }
            else
            {
                dataStatusText.Text = "数据正常";
                dataStatusText.Foreground = Brushes.Green;
            }
        }

        private void ResetStatistics()
        {
            maxAdcValue = int.MinValue;
            minAdcValue = int.MaxValue;
            avgAdcValue = 0;
            sumAdcValue = 0;
            dataPointCount = 0;
            currentAdcText.Text = "--";
            maxAdcText.Text = "--";
            minAdcText.Text = "--";
            avgAdcText.Text = "--";
            maxTempValue = double.MinValue;
            minTempValue = double.MaxValue;
            avgTempValue = 0;
            sumTempValue = 0;
            currentTempText.Text = "--";
            maxTempText.Text = "--";
            minTempText.Text = "--";
            avgTempText.Text = "--";
            adcDataLog.Clear();
            adcChart.Children.Clear();
            adcChart.Children.Add(new TextBlock
            {
                Text = "光强ADC与温度值对应的数据图表区域",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Gray,
                FontStyle = FontStyles.Italic
            });
        }

        private void SetAdcThresholdButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                upperAdcThreshold = int.Parse(upperAdcThresholdTextBox.Text);
                lowerAdcThreshold = int.Parse(lowerAdcThresholdTextBox.Text);
                MessageBox.Show($"ADC范围设置成功: 上限={upperAdcThreshold}, 下限={lowerAdcThreshold}");
                CheckAdcThresholds(currentAdcValue);
            }
            catch (FormatException)
            {
                MessageBox.Show("请输入有效的整数");
            }
        }

        private void MuteAlertButton_Click(object sender, RoutedEventArgs e)
        {
            isAlertMuted = !isAlertMuted;
            muteAlertButton.Content = isAlertMuted ? "取消静音" : "静音提醒";
            if (!isAlertMuted) CheckAdcThresholds(currentAdcValue);
            else
            {
                dataStatusText.Text = "数据正常";
                dataStatusText.Foreground = Brushes.Green;
            }
        }

        private void DataFormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataFormatComboBox.SelectedIndex >= 0)
            {
                selectedDataFormat = (DataFormat)dataFormatComboBox.SelectedIndex;
                dataFormatText.Text = $"数据格式: {GetDataFormatDescription(selectedDataFormat)}";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            DisconnectDevice();
        }

        private async void RestoreDefaultsBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckDeviceConnection()) return;

            var result = MessageBox.Show(
                "确定要恢复出厂默认校准表吗？\n此操作将清除所有自定义校准数据并恢复默认设置。",
                "确认恢复出厂设置",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                await SendRestoreDefaultsCommandAsync();
            }
        }

        private async Task SendRestoreDefaultsCommandAsync()
        {
            try
            {
                ShowProgress("正在恢复出厂设置...");

                await SendCommandAsync(CMD_RESTORE_DEFAULTS);

                //有时候第一段 2 秒是等硬件重启
                await Task.Delay(2000);

                //第二段 1 秒是等硬件重新初始化传感器采样
                await Task.Delay(1000);

                RequestPowerTableBtn_Click(null, null);
                await Task.Delay(1000);
                RequestTempTableBtn_Click(null, null);

                Dispatcher.Invoke(() =>
                {
                    ShowSuccess("恢复出厂设置完成，已重新加载默认校准表");

                    MessageBox.Show(
                        "恢复出厂设置完成！\n设备已恢复默认校准表。",
                        "恢复成功",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                });

            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    ShowError($"恢复出厂设置失败: {ex.Message}");
                    MessageBox.Show($"恢复出厂设置时出错: {ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }
    }

    public class AdcDataPoint
    {
        public int AdcValue { get; set; }
        public double Temperature { get; set; }
        public double PowerValue { get; set; }
        public int TemperatureRaw { get; set; }

        public DateTime Timestamp { get; set; }
        public bool IsRising { get; set; }
    }

    public enum DataFormat
    {
        LittleEndian16Bit,
        BigEndian16Bit,
        Adc12Bit,
        Adc10Bit,
        RawBytes
    }
}