using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;
using System.Net;
using System.Timers;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interactivity;
using System.Windows.Markup;
using System.Windows.Threading;
using System.Xml.Schema;
using AAM.Windows;
//using Basler.Pylon.Controls.Common.Helpers;       //jun
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using OxyPlot;
using NLog;
using OxyPlot.Wpf;
using LineSeries = OxyPlot.Series.LineSeries;


namespace AAM {
    public class MainViewModel : INotifyPropertyChanged {
        #region ENUM
        public enum OutputModeSelection {
            [Description("HSSL(LVTTL)")] HSSL_LVTTL = 0,
            [Description("HSSL(LVDS)")] HSSL_LVDS,
            [Description("AquadB(LVTTL)")] AquadB_LVTTL,
            [Description("AquadB(LVDS)")] AquadB_LVDS,
            [Description("SinCos(LVTTL Error Signal)")] SinCos_LVTTL,
            [Description("SinCos(LVDS Error Signal)")] SinCos_LVDS
        }

        public enum Direction {
            [Description("Bi-Direction")] BiDirection=0,
            [Description("Uni-Direction")] UniDirection
        }

        public enum Location
        {
            [Description("Singapore")] Singapore = 0,
            [Description("Shanghai")] Shanghai
        }

        public enum AxisInfo
        {
            [Description("X")] X = 0,
            [Description("Y")] Y,
            [Description("Z")] Z
        }

        public enum MeasurementMode {
            [Description("Repeatabiltiy (A1)")]
            Repeatability = 0,
            [Description("Straightness (A1&2)")]
            Straightness,
            [Description("Flatness (A1&3)")]
            Flatness,
            [Description("Straightness & Flatness (A1,2,3)")]
            StraightnessFlatness
        }
        #endregion

        #region Declaration
        private ICommand _changeTitle, _capture, _exitCommand, _connect, _applyInterfaceSetting, _discardInfaceSetting,
            _disconnect,_restart, _changeview, _start, _stop, _setupbutton, _inputbutton, _closeinfodialog, _startalignment, _startpilot,
            _changesignage, _changedatum, _changeresolution, _startecu, _startoperate;
        //private string _plotTitle;
        private readonly IList<DataPoint>[] _pointList = {
            new List<DataPoint>(), new List<DataPoint>(), new List<DataPoint>(),
            new List<DataPoint>(), new List<DataPoint>(), new List<DataPoint>()
        };
        private readonly LineSeries[] _series = new LineSeries[6];
        //private DataPoint _tempPoint = new DataPoint();
        private ObservableCollection<DisplayResult> _displaygrid = new ObservableCollection<DisplayResult>();
        private DisplayResult _selectedrow = new DisplayResult();
        private PlotModel _plotModel;

        private bool _connectionstatus,
            _measurementstatus,
            _avgon,
            _piloton,
            _machinerunning,
            _interfacechanged,
            _iswait,
            _ispreparing,
            _alignmenton,
            _measurementon,
            _cancelon,
            _a1enabled,
            _a2enabled,
            _a3enabled,
            _ecuon,
            _lastrow;
        private readonly IDialogCoordinator _dialogCoordinator;
        private int _a1Signal = 100, _a2Signal = 100, _a3Signal = 100, _port = 9090,_lastsettingindex=2, _signage=1, _resolutionindex, _avg/*, _count*/;
        private string _view, _ipaddress = "192.168.1.1", _infotext, _resolution = "0.00000", _sError, _sStrightness, _sFlatness,/*_sTemperatureStart,_sHumidityStart,_sPressureStart,_sRefractiveIndexStart*/_operatorName;
        private readonly string[] _resolutionArray = { "0.0000", "0.00000", "0.000000" };
        private double tempError, tempStrightness, tempFlatness;                                                                   
        // private double _error, _target;                                                                                        
        private CustomDialog _dialog, _infodialog,_waitForConnecting,_waitForStartAlignment,_inputinformation;
       // private OutputModeSelection _outputmode = OutputModeSelection.HSSL_LVTTL;
        private readonly AttocubeMeasurement _attocubeMeasurement;
        private AttocubeMeasurement.InterfaceProperties _localInterface; //, _currentInterface
        private AttocubeMeasurement.InterfaceProperties _currentInterfaceConnection;
        // private AttocubeMeasurement.InterfaceProperties _attocubeproperties;
        private ObservableCollection<string> _errorlist;
        private RunSetting _runSetting = new RunSetting {Interval = 0, Run = 1, RunDirection = Direction.BiDirection, Stroke=0};
        private Location _location;
        private AxisInfo _axisinfo;
        private string _a1displacement, _a2displacement, _a3displacement;
        private AttocubeMeasurement.PassMode _axisMode = AttocubeMeasurement.PassMode.SinglePass;
        private MeasurementMode _measurementMode = MeasurementMode.Repeatability;
        public readonly static ILogger Log = LogManager.GetLogger("History");
        private readonly System.Timers.Timer _timer = new System.Timers.Timer(300); //100ms of interval
        //private AttocubeMeasurement.RunMode currentmode;
        private string _ecutemp;

        // private DataRow _selectedrow;
        #endregion

        #region Init
        public MainViewModel(IDialogCoordinator dialogCoordinator) {
            _dialogCoordinator = dialogCoordinator;
            _localInterface = new AttocubeMeasurement.InterfaceProperties();
            SwitchView = "operation";
            //PlotPointList = new List<DataPoint>();
            InitDataGrid();
            _plotModel = new PlotModel() {TitleColor = OxyColors.Red};
            //_attocubeproperties = _attocubeMeasurement.Interface;
            try {
                _attocubeMeasurement = new AttocubeMeasurement();
            }
            catch {
                MessageBox.Show("Fail to load Attocube dll");
            }
            _errorlist = new ObservableCollection<string>() {"a", "b"};
            SetupLines();
            _timer.Elapsed += _timer_Elapsed;
            Log.Info("DataContext loaded");
            MeasurementOn = false;
            MachineRunning = false;
            CancelOn = false;
        }
        #endregion Init

        #region Property
        #region Setting Page
        #region Interface Setting 
        public int HsslLow {
            get { return LocalInterface.HsslLowResolution; }
            set
             {
                LocalInterface.HsslLowResolution = value;
                OnPropertyChanged();
            }
        }

        public int HsslHigh {
            get { return LocalInterface.HsslHighResolution; }
            set {
                LocalInterface.HsslHighResolution = value;
                OnPropertyChanged();
            }
        }

        public int HsslClock {
            get { return LocalInterface.HsslClockPeriod; }
            set {
                LocalInterface.HsslClockPeriod = value;
                OnPropertyChanged();
            }
        }

        public int HsslGap {
            get { return LocalInterface.HsslPeriodGap; }
            set {
                LocalInterface.HsslPeriodGap = value;
                OnPropertyChanged();
            }
        }

        public int SincosClock {
            get { return LocalInterface.SinCosClockPeriod; }
            set {
                LocalInterface.SinCosClockPeriod = value;
                OnPropertyChanged();
            }
        }

        public int SincosRes {
            get { return LocalInterface.SinCosClockPeriod; }
            set {
                LocalInterface.SinCosClockPeriod = value;
                OnPropertyChanged();
            }
        }

        public OutputModeSelection OutputMode {
            get { return (OutputModeSelection)LocalInterface.RealtimeOutput; }
            set {
                LocalInterface.RealtimeOutput = (int)value;
                OnPropertyChanged();
            }
        }

        public int Average{
            get {
                int tempavg;
                return _avg = (_attocubeMeasurement.GetAveragingN(out tempavg) == 0) ? tempavg : _avg;
            }
            set {
                if (_attocubeMeasurement.SetAveragingN(value) == 0) {
                    _avg = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool InterfaceChanged {
            get { return _interfacechanged; }
            set {
                _interfacechanged = value;
                OnPropertyChanged();
            }
        }
        public int LastSettingIndex{
            get { return _lastsettingindex; }
            set {
                _lastsettingindex = value;
                OnPropertyChanged();
            }
        }

        private AttocubeMeasurement.InterfaceProperties CurrentInterface{
            get { return _attocubeMeasurement.Interface; }
            set {
                _attocubeMeasurement.Interface = value;
                _currentInterfaceConnection = _attocubeMeasurement.Interface;
                OnPropertyChanged();
            }
        }
        


        private AttocubeMeasurement.InterfaceProperties LocalInterface {
            get { return _localInterface; }
            set {
                _localInterface = value;
                InterfaceChanged = !Equals(_localInterface, _currentInterfaceConnection);
                OnPropertyChanged();
            }
        }
        #endregion
        #region Alignment
        public AttocubeMeasurement.PassMode AxisMode {
            get { return _axisMode; }
            set {
                _axisMode = value;
                OnPropertyChanged();
            }
        }
        public bool MeasurementOn {
            get { return _measurementon; }
            set {
                _measurementon = value;
                OnPropertyChanged();
            }
        }
        public bool CancelOn
        {
            get { return _cancelon; }
            set
            {
                _cancelon = value;
                OnPropertyChanged();
            }
        }
        public bool AlignmentOn{
            get { return _alignmenton; }
            set {
                _alignmenton = value;
                OnPropertyChanged();
            }
        }
        public bool PilotOn {
            get { return _piloton; }
            set {
                _piloton = value;
                OnPropertyChanged();
            }
        }
         public int A1SignalStrength {
            get { return _a1Signal; } 
            set {
                _a1Signal = value;
                OnPropertyChanged();
            }
        }

        public int A2SignalStrength {
            get { return _a2Signal; }
            set {
                _a2Signal = value;
                OnPropertyChanged();
            }
        }

        public int A3SignalStrength {
            get { return _a3Signal; }
            set {
                _a3Signal = value;
                OnPropertyChanged();
            }
        }
        public MeasurementMode RunMeasurementMode {
            get { return _measurementMode; }
            set {
                _measurementMode = value;
                A1Enabled = true;
                A2Enabled = (_measurementMode == MeasurementMode.Straightness) ||
                            (_measurementMode == MeasurementMode.StraightnessFlatness);
                A3Enabled = (_measurementMode == MeasurementMode.Flatness) ||
                            (_measurementMode == MeasurementMode.StraightnessFlatness);
                OnPropertyChanged();
            }
        }
        #endregion
        #region ECU
        public bool EcuOn {
            get { return _ecuon; }
            set {
                _ecuon = value;
                OnPropertyChanged();
            }
        }
        public bool EcuConnectionStatus {
            get { return _attocubeMeasurement.Ecu.ConnectionStatus; }
            set {
                _attocubeMeasurement.Ecu.ConnectionStatus = value;
                OnPropertyChanged();
            }
        }
        public string EcuStat {
            get { return _attocubeMeasurement.Ecu.ConnectionStatus ? "Connected" : "Disconnected"; }
            set { OnPropertyChanged(); }
        }
        public Color EcuMedia {
            get { return _attocubeMeasurement.Ecu.ConnectionStatus ? Color.Lime : Color.Red; }
            set { OnPropertyChanged(); }
        }
        public double EcuHumidity {
            get { return _attocubeMeasurement.Ecu.Humidity; }
            set {
                _attocubeMeasurement.Ecu.Humidity = value;
                OnPropertyChanged();
            }
        }
        public double EcuPressure {
            get { return _attocubeMeasurement.Ecu.Pressure; }
            set {
                _attocubeMeasurement.Ecu.Pressure = value;
                OnPropertyChanged();
            }
        }
        public double EcuRefractiveIndex {
            get { return _attocubeMeasurement.Ecu.RefractiveIndex; }
            set {
                _attocubeMeasurement.Ecu.RefractiveIndex = value;
                OnPropertyChanged();
            }
        }
        public double EcuTemperature
        {
            get { return _attocubeMeasurement.Ecu.Temperature; }
            set
            {
                _attocubeMeasurement.Ecu.Temperature = value;
                OnPropertyChanged();
            }
        }
        #endregion

        #region Measurement
        public string A1Displacement {
            get { return _a1displacement; }
            set {
                _a1displacement = value;
                OnPropertyChanged();
            }
        }
        public string A2Displacement {
            get { return _a2displacement; }
            set {
                _a2displacement = value;
                OnPropertyChanged();
            }
        }
        public string A3Displacement {
            get { return _a3displacement; }
            set {
                _a3displacement = value;
                OnPropertyChanged();
            }
        }
        public bool A1Enabled {
            get { return _a1enabled; }
            set {
               _a1enabled = value;
                OnPropertyChanged();
            }
        }
        public bool A2Enabled {
            get { return _a2enabled; }
            set {
                _a2enabled = value;
                OnPropertyChanged();
            }
        }
        public bool A3Enabled {
            get { return _a3enabled; }
            set {
                _a3enabled = value;
                OnPropertyChanged();
            }
        }
        public string DisplayResolution {
            get { return _resolution; }
            set {
                _resolution = value;
                OnPropertyChanged();
            }
        }
        #endregion
        #region Display
        public Direction DirectionMode {
            get { return _runSetting.RunDirection; }
            set {
                _runSetting.RunDirection = value;
                OnPropertyChanged();
            }
        }

        public Location LocationMode
        {
            get { return _location; }
            set
            {
                _location = value;
                OnPropertyChanged();
            }
        }

        public AxisInfo Axis
        {
            get { return _axisinfo; }
            set
            {
                _axisinfo = value;
                OnPropertyChanged();
            }
        }

        public double Stroke {
            get { return _runSetting.Stroke; }
            set {
                _runSetting.Stroke = value;
                OnPropertyChanged();
            }
        }
        public int Runs {
            get { return _runSetting.Run; }
            set {
                _runSetting.Run = value;
                OnPropertyChanged();
            }
        }
        public double Interval {
            get { return _runSetting.Interval; }
            set {
                _runSetting.Interval = value;
                OnPropertyChanged();
            }
        }

        public string Address {
            get { return _ipaddress; }
            set {
                IPAddress tempip;
                if (IPAddress.TryParse(value, out tempip)) {
                    if (value == null || value.Count(c => c == '.') != 3) {
                        return;
                    }
                    _ipaddress = value;
                }
                OnPropertyChanged();
            }
        }
        public int Port {
            get { return _port; }
            set {
                _port = value;
                OnPropertyChanged();
            }
        }

        public string PlotTitle {
            get { return _plotModel.Title; } //_plotTitle; }
            private set {
                _plotModel.Title = value;//_plotTitle = value;
                PlotModel.InvalidatePlot(true);
                OnPropertyChanged();
            }
        }

        public string OperatorName
        {
            get { return _operatorName; } //_plotTitle; }
            private set
            {
                _operatorName = value;//_plotTitle = value;
                OnPropertyChanged();
            }
        }

        public PlotModel PlotModel {
            get { return _plotModel; }
            private set {
                _plotModel = value;
                _plotModel.InvalidatePlot(true);
                OnPropertyChanged();
            }
        }

        //public IList<DataPoint> PlotPointList {
        //    get { return _pointList; }
        //    private set {
        //        _pointList = value;
        //        OnPropertyChanged();
        //    }
        //}

        public ObservableCollection<DisplayResult> DisplayGrid {
            get { return _displaygrid; }
            set {
                _displaygrid = value;
                OnPropertyChanged();
            }
        }

        //public DataRow SelectedGridItem{
        public DisplayResult SelectedGridItem {
            get { return _selectedrow; }
            set {
                _selectedrow = value;
                OnPropertyChanged();
            }
        }

        public string SwitchView {
            get { return _view; }
            private set {
                _view =value;
               // StartMeasureFunction(value != "settings");
                OnPropertyChanged();
            }
        }
        public bool IsConnected {
            get {
                return _connectionstatus;
            }
            private set {
                _connectionstatus = value;
                OnPropertyChanged();
            }
        }
        //public bool IsConnected => _attocubeMeasurement.IsConnected;

        public bool IsMeasurementStarted {
            get { return _measurementstatus; }
            set {
                _measurementstatus = value;
                OnPropertyChanged();
            }
        }

        public bool AverageOn {
            get {
                return _avgon;
            }
            set {
                if (SetAveragingOn(value))
                {
                    _avgon = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool MachineRunning {
            get { return _machinerunning; }
            set {
                _machinerunning = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> ErrorList {
            get { return _errorlist; }
            private set {
                _errorlist = value;
                OnPropertyChanged();
            }
        }
        #endregion Display
        #endregion

        #region ICommand List
        //For Oxyplot
        public ICommand ChangeTitle => _changeTitle ?? (_changeTitle = new RelayCommand(CanExecute, param => Change()));
        public ICommand Capture => _capture ?? (_capture = new RelayCommand(s => UnionCaptureAndOperate()/*CheckIsLastRow()*/, param => AddData()));

        //For Operation User Control
        public ICommand Connect => _connect ?? (_connect = new RelayCommand(CanExecute, param => ConnectDevice()));
        public ICommand Disconnect => _disconnect ?? (_disconnect = new RelayCommand(CanExecute, param => DisconnectDevice()));
        public ICommand ReStart => _restart ?? (_restart = new RelayCommand(CanExecute, param => Restart()));
        public ICommand StartMeasure=> _start ?? (_start = new RelayCommand(s => UnionStartAndOperate(), param => StartMeasurement()));
        public ICommand StopMeasure => _stop ?? (_stop = new RelayCommand(s => UnionEndAndOperate()/*GetMeasurementStartedStatus()*/, param => StopMeasurement()));
        public ICommand CloseInfoDialog => _closeinfodialog ?? (_closeinfodialog = new RelayCommand(CanExecute, param => HideInfoDialog(_infodialog)));
        public ICommand ChangeSignage => _changesignage ?? (_changesignage = new RelayCommand(CanExecute, param => Signage()));
        public ICommand ChangeDatum => _changedatum ?? (_changedatum = new RelayCommand(CanExecute, param => Datum()));
        public ICommand ChangeResolution => _changeresolution ?? (_changeresolution = new RelayCommand(CanExecute, param => Resolution()));

        //For MainWindow
        public ICommand ExitApp => _exitCommand ?? (_exitCommand = new RelayCommand(CanExecute, Exit));
        public ICommand ChangeView => _changeview ?? (_changeview = new RelayCommand(CanExecute, ChangeViewUserControl));

        //For Setup User Control
        public ICommand SetupButton => _setupbutton ?? (_setupbutton = new RelayCommand(CanExecute, param =>SetupPressed(Convert.ToBoolean(param))));

        public ICommand InputButton => _inputbutton ??
                                       (_inputbutton = new RelayCommand(CanExecute, param => SaveDatatoFile()));
        public ICommand SaveInterface => _applyInterfaceSetting?? (_applyInterfaceSetting = new RelayCommand(GetInterfaceChanged, param => SaveOrDiscardInterface((Convert.ToBoolean(param)))));
        public ICommand DiscardInterface => _discardInfaceSetting ?? (_discardInfaceSetting = new RelayCommand(GetInterfaceChanged, param => SaveOrDiscardInterface((Convert.ToBoolean(param)))));

        //For Operation
        public ICommand StartOperate => _startoperate ?? (_startoperate = new RelayCommand(CanExecute, param => StartMeasureFunction((Convert.ToBoolean(param)))));

        //For Alignment
        public ICommand StartAlignment => _startalignment ?? (_startalignment = new RelayCommand(CanExecute, param => StartAlignmentFunction((Convert.ToBoolean(param)))));
        public ICommand StartPilot => _startpilot ?? (_startpilot = new RelayCommand(CanExecute, param => StartPilotFunction((Convert.ToBoolean(param)))));

        //For Interface
        public ICommand StartEcu => _startecu ?? (_startecu = new RelayCommand(CanExecute, param => StartEcuOnFunction((Convert.ToBoolean(param)))));
        #endregion
        #endregion Property

        #region Private Method

        private bool SetAveragingOn(bool toOn) {
            if (_attocubeMeasurement.SetAveragingN(((toOn) ? Average : 0)) == 0) {
                Log.Info($"Averaging is turned {(toOn ? "on" : "off")}, Average value is {Average}");
                return true;
            }
            Log.Error("Fail to configure averaging");
            return false;
        }

        private async void StartMeasureFunction(bool tostart) {
            if (MeasurementOn)
            {
                var mySettings = new MetroDialogSettings()
                {
                    AnimateShow = true,
                    ColorScheme = MetroDialogColorScheme.Theme
                };
                _waitForConnecting = new CustomDialog(mySettings) {Content = new WaitForPreparing()};
                await _dialogCoordinator.ShowMetroDialogAsync(this, _waitForConnecting);

                InfoText = "Preparing...";
                IsPreparing = true;

                DateTime _start = DateTime.Now;

                if (!StopAllRun())
                {
                    MeasurementOn = false;
                    MachineRunning = false;
                    HideInfoDialog(_waitForConnecting);
                    CancelOn = false;
                    return;
                }
                if (tostart)
                {
                    CancelOn =true;
                    int result = 0;
                    await Task.Run(() => (result = _attocubeMeasurement.StartMeasurement()));// == 0);

                    
                    if(result!=0)//if (_attocubeMeasurement.StartMeasurement() != 0)
                    {
                        if (_attocubeMeasurement.ErrorCode == 4099)
                        {
                            Log.Error("Please set alignment first");
                        }
                        MeasurementOn = false;
                        HideInfoDialog(_waitForConnecting);
                        CancelOn = false;
                        Log.Error("Unable to start measurement");
                        return;
                    }
                    
                    MeasurementOn = true;
                    HideInfoDialog(_waitForConnecting);
                    CancelOn = false;
                    MachineRunning = true;
                }
                IsPreparing = false;
            }
            else
            {
                Thread.Sleep(50);
                if (!StopAllRun())
                {
                    return;
                }
            }
            
        }

        private async void StartAlignmentFunction(bool tostart) {
            if (AlignmentOn)
            {
                var mySettings = new MetroDialogSettings()
                {
                    AnimateShow = true,
                    ColorScheme = MetroDialogColorScheme.Theme
                };
                _waitForStartAlignment= new CustomDialog(mySettings) { Content = new WaitForStartAlignment() };
                await _dialogCoordinator.ShowMetroDialogAsync(this, _waitForStartAlignment);

                InfoText = "Preparing...";
                IsPreparing = true;

                if (!StopAllRun())
                {
                    HideInfoDialog(_waitForStartAlignment);
                    return;
                }
                if (tostart)
                {
                    CancelOn = true;
                    int result = 0;
                    await Task.Run(() => (result = _attocubeMeasurement.StartOpticalAlignment()));

                    if (result != 0)
                    {
                        HideInfoDialog(_waitForStartAlignment);
                        CancelOn = false;
                        Log.Error("Unable to start optical alignment");
                        return;
                    }
                    AlignmentOn = true;
                    HideInfoDialog(_waitForStartAlignment);
                    CancelOn = false;
                }
                IsPreparing = false;
            }
            else
            {
                if (!StopAllRun())
                {
                    return;
                }
                if (tostart)
                {
                    if (_attocubeMeasurement.StartOpticalAlignment() != 0)
                    {
                        Log.Error("Unable to start optical alignment");
                        return;
                    }
                    AlignmentOn = true;
                }
            }
        }

        private bool StartPilotFunction(bool tostart) {
            if (!StopAllRun()) {
                return false;
            }
            if (tostart) {
                if (_attocubeMeasurement.SetPilotLaser(true) != 0) {
                    Log.Error("Unable to turn on pilot laser");
                    return false;
                }
                PilotOn = true;
            }
            return true;
        }

        private bool StartEcuOnFunction(bool tostart) {
            if (tostart) {  //jun
                if (_attocubeMeasurement.EnableEcu() != 0) {
                    Log.Error($"Unable to enable ECU");
                    return false;
                }
            } else {
                if (_attocubeMeasurement.DisableEcu() != 0) {
                    Log.Error($"Unable to disable ECU");
                    return false;
                }
            }
            EcuOn = tostart;
            return true;
        }

        private void ChangeViewUserControl(object viewnum) {
            SwitchView = Convert.ToString(viewnum );
            Log.Info("Switch view to " + SwitchView);
        }
        
        private bool StopAllRun() {
            if (_attocubeMeasurement == null) {
                Log.Error("Attocube measurement is not initialized");
                return false;
            }
           AttocubeMeasurement.RunMode currentmode = _attocubeMeasurement.GetCurrentRunMode();//.CurrentRunMode;
            //jun: Check & stop running process
            if ((currentmode == AttocubeMeasurement.RunMode.MeasurementStarting) || 
                (currentmode == AttocubeMeasurement.RunMode.MeasurementRunning)) {
                if (_attocubeMeasurement.StopMeasurement() != 0) {
                    Log.Error("Unable to stop measurement mode");
                    return false;
                }
                MeasurementOn = false;
                MachineRunning = false;
                IsMeasurementStarted = false;
            }
            else if ((currentmode == AttocubeMeasurement.RunMode.AlignmentStarting) || 
                (currentmode == AttocubeMeasurement.RunMode.AlignmentRunning)) {
                if (_attocubeMeasurement.StopOpticalAlignment() != 0) {
                    Log.Error("Unable to stop alignment mode");
                    return false;
                }
                AlignmentOn = false;
            }
            else if (currentmode == AttocubeMeasurement.RunMode.PilotLightOn) {
                if (_attocubeMeasurement.SetPilotLaser(false) != 0) {
                    Log.Error("Unable to turn off pilot laser");
                    return false;
                }
                PilotOn = false;
            }
            return true;
        }

        private void Change() {
            PlotTitle = "Test";
        }

        private bool CanExecute(object o) {
            return true;
        }

        private async void Exit(object o) {
            if (o != null)
            {
                var arg = (CancelEventArgs) o;
                arg.Cancel = true;
            }
            var mySettings = new MetroDialogSettings() {
                AnimateShow = true,
                FirstAuxiliaryButtonText = "Cancel",
                AffirmativeButtonText = "Yes",
                ColorScheme = MetroDialogColorScheme.Accented,
            };
            var result = await _dialogCoordinator.ShowMessageAsync(this, "Exit Application", "Do you want to exit Application?",
                    MessageDialogStyle.AffirmativeAndNegative, mySettings);
            if (result == MessageDialogResult.Negative) {
                return;
            }
            _timer.Stop();
            Application.Current.Shutdown();
        }

        //private void ConnectDevice() {
        //    ShowInfoDialog("Waiting for connection...", true);
        //        Application.Current.Dispatcher.Invoke(DispatcherPriority.Background,
        //            new Action(delegate { }));
        //    if (_attocubeMeasurement.Connect(Address, Port) == 0) {
        //        LocalInterface = CurrentInterface;
        //        IsConnected = true;
        //    }
        //    HideInfoDialog();
        //}

        private async void ConnectDevice() {
            var mySettings = new MetroDialogSettings() {
                AnimateShow = true,
                ColorScheme = MetroDialogColorScheme.Theme
            };

            _infodialog = new CustomDialog(mySettings) { Content = new InfoDialog() };
            await _dialogCoordinator.ShowMetroDialogAsync(this, _infodialog);

            InfoText = "Waiting for connection...";
            IsWaiting = true;
            CancelOn = true;

            int result=0;
            await Task.Run(()=>(result = _attocubeMeasurement.Connect(Address, Port)));// == 0);
            IsConnected = _attocubeMeasurement.IsConnected;
            if (result == 0) {
                LocalInterface = CurrentInterface;
                if (!StopAllRun()) {
                    return;
                }
                _timer.Start();
                HideInfoDialog(_infodialog);
                CancelOn = false;
            }
            else {
                InfoText = "Fail to Connect into IDS!";
                IsWaiting = false;
                CancelOn = false;
            }
        }

        private void DisconnectDevice() {
            if (_attocubeMeasurement.Disconnect() == 0) {
                IsConnected = _attocubeMeasurement.IsConnected;
            }
        }

        private void Restart()
        {
            StopAllRun();
            InitDataGrid();
            ClearPoints();
            A1Displacement = "";
            A2Displacement = "";
            A3Displacement = "";
            DisconnectDevice();
            ConnectDevice();
        }

        private bool CheckIsLastRow() {
            if (!IsMeasurementStarted)
                return false;
            if (_lastrow)
                return false;
            return true;
        }

        private async void  StartMeasurement() {
            DisplayGrid.Clear();
            ClearPoints();
            SelectedGridItem = null;
            _lastrow = false;
            var mySettings = new MetroDialogSettings() {
                AnimateShow = true,
                FirstAuxiliaryButtonText = "Cancel",
                ColorScheme = MetroDialogColorScheme.Inverted
            };
            _dialog = new CustomDialog(mySettings) {Content = new SetupDialog()};
            await _dialogCoordinator.ShowMetroDialogAsync(this, _dialog);
        }

        private async void InputInfomation()//dian, for typing in operator's information
        {
            var mySettings = new MetroDialogSettings()
            {
                AnimateShow = true,
                FirstAuxiliaryButtonText = "Cancel",
                ColorScheme = MetroDialogColorScheme.Inverted
            };
            _inputinformation = new CustomDialog(mySettings) { Content = new InfoInput() };
            await _dialogCoordinator.ShowMetroDialogAsync(this, _inputinformation);
        }

        private void SetupPressed(bool result)
        {
            var mySettings = new MetroDialogSettings()
            {
                AnimateShow = true,
                FirstAuxiliaryButtonText = "Cancel",
                ColorScheme = MetroDialogColorScheme.Accented,
            };
            if ((Runs == 0 || Stroke == 0 || Interval == 0) && result)
            {
                MessageBox.Show("Invalid Stroke or Run or Interval");
                return;
            }
            if ((int) (Stroke % Interval) != 0 && result)
            {
                //jun: double must convert to int first
                MessageBox.Show("Interval is not valid for Stroke");
                return;
            }
            if (_dialog.Visibility == Visibility.Visible)
            {
                _dialogCoordinator.HideMetroDialogAsync(this, _dialog, mySettings);
            }
            IsMeasurementStarted = result;

            if (!result) return; //return if cancel button is pressed

            //calculate setup parameters
            if (DirectionMode == Direction.BiDirection)
            {
                int runindex = (int) (Stroke / Interval) + 1;
                int index = 0, runnum = 1;
                for (int i = 0; i < Runs; i++)
                {
                    for (int j = 0; j < runindex; j++)
                    {
                        DisplayResult indexresult = new DisplayResult
                        {
                            TargetPos = index * Interval,
                            Index = runnum,
                            DirectionPositive = true
                        };
                        DisplayGrid.Add(indexresult);
                        index++;
                    }
                    runnum++;
                    for (int j = 0; j < runindex; j++)
                    {
                        index--;
                        DisplayResult indexresult = new DisplayResult
                        {
                            TargetPos = index * Interval,
                            Index = runnum,
                            DirectionPositive = false
                        };
                        DisplayGrid.Add(indexresult);
                    }
                    runnum++;
                }
            }
            if (DirectionMode == Direction.UniDirection)
            {
                int runindex = (int)(Stroke / Interval) + 1;
                int runnum = 1;
                for (int i = 0; i < Runs; i++)
                {
                    int index = 0;
                    for (int j = 0; j < runindex; j++)
                    {
                        DisplayResult indexresult = new DisplayResult
                        {
                            TargetPos = index * Interval,
                            Index = runnum,
                            DirectionPositive = true
                        };
                        DisplayGrid.Add(indexresult);
                        index++;
                    }
                    runnum++;
                }
            }
            SelectedGridItem = DisplayGrid.First();
        }

        private async void StopMeasurement() {
            var mySettings = new MetroDialogSettings() {
                AnimateShow = true, AffirmativeButtonText = "Save", DefaultButtonFocus = MessageDialogResult.Affirmative, FirstAuxiliaryButtonText = "Cancel",
                NegativeButtonText= "Discard",
                ColorScheme = MetroDialogColorScheme.Accented,
            };
            var result = await _dialogCoordinator.ShowMessageAsync(this,"End Acquisition", "Do you want to save or discard result?",
                    MessageDialogStyle.AffirmativeAndNegativeAndSingleAuxiliary, mySettings);
            if (result == MessageDialogResult.FirstAuxiliary) { //Cancel button pressed
                return;
            }
            if (result == MessageDialogResult.Affirmative) {
                //todo save result function
                //SavePlot();
                //SaveDatatoFile();
                InputInfomation();
            }
        IsMeasurementStarted = false;
        }

        private bool GetMeasurementStartedStatus() {//Measurement status
            return (IsMeasurementStarted);
        }
        private bool GetMeasurementRunStatus() { //Measurement is running
            return (MeasurementOn);
        }
        private bool GetStatusToStartMeasure() {
            return !IsMeasurementStarted;
        }

        private bool UnionStartAndOperate()//dian, binding operate and start
        {
            return ((GetStatusToStartMeasure()) && MeasurementOn);
        }

        private bool UnionEndAndOperate()
        {
            return GetMeasurementStartedStatus() && MeasurementOn;
        }

        private bool UnionCaptureAndOperate()
        {
            return CheckIsLastRow() && MeasurementOn;
        }

        private bool InitDataGrid() {
            //_gridResult.Add(new DisplayGridResult() {TargetPos = 100, Error = 0.001, Index = 1});
            //_displayindex.Add(1);
            _displaygrid.Add(new DisplayResult() {Error = 0, Index = 1, TargetPos = 100});
            _displaygrid.Add(new DisplayResult() { Error = 0, Index = 2, TargetPos = 200});
            return true;
        }

        private bool AddGridRowData() {
            return true;
        }

        private void SaveOrDiscardInterface(bool tosave) {
            if (tosave) {
                _attocubeMeasurement.Interface = _localInterface;
                if (_attocubeMeasurement.Interface != _localInterface) {
                    Log.Error("Unable to save interface setting");
                }
            }
            else {
                _localInterface = _attocubeMeasurement.Interface;
            }
        }

        //private async void ShowInfoDialog(string info, bool showwait) {
        //    var mySettings = new MetroDialogSettings() {
        //        AnimateShow = true,
        //        ColorScheme = MetroDialogColorScheme.Inverted,
        //    };

        //    _infodialog = new CustomDialog(mySettings) { Content = new InfoDialog() };
        //    InfoText = info;
        //    IsWaiting = showwait;
        //    await _dialogCoordinator.ShowMetroDialogAsync(this, _infodialog);
        //}

        public string InfoText {
            get { return _infotext; }
            private set {
                _infotext = value;
                OnPropertyChanged();
            }
        }

        public bool IsWaiting {
            get { return _iswait; }
            private set {
                _iswait = value;
                OnPropertyChanged();
            }
        }

        public bool IsPreparing
        {
            get { return _ispreparing; }
            private set
            {
                _ispreparing = value;
                OnPropertyChanged();
            }
        }

        private void HideInfoDialog(CustomDialog _dialog) {
            if(_dialog.Visibility == Visibility.Visible)
            _dialogCoordinator.HideMetroDialogAsync(this, _dialog);
        }

        private bool GetInterfaceChanged(object o) {
            return InterfaceChanged;
        }

        private void AddPointsToPlot(DataPoint dataPoint, MeasurementMode mode, bool isPositive) {
            if (mode == MeasurementMode.StraightnessFlatness)
                return;
            int direction = (isPositive) ? 0 : 3;
            int index = (int)mode + direction;
            PlotModel plot = _plotModel;
            _pointList[index].Add(dataPoint);
            PlotModel = plot;
        }

        private void SetupLines() {
            bool isPositive = true;
            for (int i = 0; i < 2; i++) {
                foreach (MeasurementMode mode in Enum.GetValues(typeof (MeasurementMode))) {
                    if(mode == MeasurementMode.StraightnessFlatness) continue;
                    OxyColor colors;
                    switch (mode) {
                        case MeasurementMode.Repeatability:
                            colors = (isPositive) ? OxyColors.Red : OxyColors.Goldenrod;
                            break;
                        case MeasurementMode.Straightness:
                            colors = (isPositive) ? OxyColors.LimeGreen : OxyColors.MediumBlue;
                            break;
                        case MeasurementMode.Flatness:
                            colors = (isPositive) ? OxyColors.Cyan : OxyColors.Purple;
                            break;
                        default:
                            colors = OxyColors.Black;
                            break;
                    }
                    int direction = (isPositive) ? 0 : 3;
                    int index = (int)mode + direction;
                    _series[index] = new LineSeries {
                        ItemsSource = _pointList[index],
                        MarkerType = MarkerType.Diamond,
                        MarkerSize = 2,
                        Color = colors,
                        MarkerStroke = OxyColors.Gray,
                        MarkerFill = OxyColors.Gray
                    };
                }
                isPositive = false;
            }
            foreach (var line in _series) {
                PlotModel.Series.Add(line);
            }
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e) {    //jun
            _timer.Stop();
            AttocubeMeasurement.RunMode currentmode = _attocubeMeasurement.GetCurrentRunMode();//.CurrentRunMode;
            //update displacement 
            try
            {
                if (currentmode == AttocubeMeasurement.RunMode.MeasurementRunning)
                {
                    long[] displacement;
                    if (_attocubeMeasurement.GetMeasurementDisplacement(out displacement) == 0)
                    {
                        A1Displacement = (displacement[0] == -1) ? "Error" : 
                        ((Convert.ToDouble(displacement[0]) / 1000000000*_signage).ToString(DisplayResolution));
                        if (RunMeasurementMode == MeasurementMode.Repeatability ||
                            RunMeasurementMode == MeasurementMode.Flatness)
                        {
                            A2Displacement = (displacement[1] == -1)
                                ? "-"
                                : ((Convert.ToDouble(displacement[1]) / 1000000000 * _signage)
                                    .ToString(DisplayResolution));
                        }
                        else
                        {
                            A2Displacement = (displacement[1] == -1)
                                ? "Error"
                                : ((Convert.ToDouble(displacement[1]) / 1000000000 * _signage)
                                    .ToString(DisplayResolution));
                        }
                        if (RunMeasurementMode == MeasurementMode.Repeatability ||
                            RunMeasurementMode == MeasurementMode.Straightness)
                        {
                            A3Displacement = (displacement[2] == -1)
                                ? "-"
                                : ((Convert.ToDouble(displacement[2]) / 1000000000 * _signage)
                                    .ToString(DisplayResolution));
                        }
                        else
                        {
                            A3Displacement = (displacement[2] == -1)
                                ? "Error"
                                : ((Convert.ToDouble(displacement[2]) / 1000000000 * _signage)
                                    .ToString(DisplayResolution));
                        }
                    }
                    else
                    {
                        Log.Error("Operation stopped");
                        StopAllRun();
                    }
                }
                if (currentmode== AttocubeMeasurement.RunMode.AlignmentRunning) {
                    int[] contrast;
                    if (_attocubeMeasurement.GetOpticalAlignmentData(out contrast) == 0) {      //jun, added
                        A1SignalStrength = contrast[0];
                        A2SignalStrength = contrast[1];
                        A3SignalStrength = contrast[2];
                    }
                }
                if (EcuOn) {
                    AttocubeMeasurement.EcuProperties ecu;
                    if (_attocubeMeasurement.GetEcuProperties(out ecu) == 0) {     //jun, added
                        //EcuConnectionStatus = ecu.ConnectionStatus;
                        EcuStat = ecu.ConnectionStatus.ToString();
                        EcuMedia = Color.Black;
                        EcuTemperature = ecu.Temperature;
                        EcuHumidity = ecu.Humidity;
                        EcuPressure = ecu.Pressure;
                        EcuRefractiveIndex = ecu.RefractiveIndex;
                    }
                    
                }
                if (!EcuOn)
                {
                    AttocubeMeasurement.EcuProperties ecu;
                    if (_attocubeMeasurement.GetEcuProperties(out ecu) == 0)
                    {
                        EcuStat = ecu.ConnectionStatus.ToString();
                        }
                }
            }
            catch { }
            _timer.Start();
        }

        private void Signage() {
            _signage *= -1;
            Log.Info($"Signage is inverted. current sign is {_signage}");
        }

        private void Datum() {
            if (_attocubeMeasurement.ResetDisplacement() != 0) { 
                Log.Error("Fail to reset all displacement");
            }
        }

        private void Resolution()
        {
            _resolutionindex++;
            if (_resolutionindex > _resolutionArray.Length-1) 
                _resolutionindex = 0;
            DisplayResolution = _resolutionArray[_resolutionindex];
        }

        private bool SavePlot() {
            _plotModel = PlotModel;
            try {
                using (var stream = new FileStream(Environment.CurrentDirectory + "\\test.png", FileMode.OpenOrCreate)) {
                    var pngExporter = new PngExporter {Width = 600, Height = 400, Background = OxyColors.White};
                    pngExporter.Export(_plotModel, stream);
                }
                Log.Info("Plot image saved sucessfully");
                return true;
            }
            catch (Exception ex) {
                Log.Error($"Fail to save Plot Image: {ex.Message}");
                return false;
            }
        }

        private bool SaveDatatoFile(string path = "")
        {
            var mySettings = new MetroDialogSettings()
            {
                AnimateShow = true,
                FirstAuxiliaryButtonText = "Cancel",
                ColorScheme = MetroDialogColorScheme.Inverted
            };
            _dialogCoordinator.HideMetroDialogAsync(this, _inputinformation, mySettings);
            if (path == string.Empty)
                path = Environment.CurrentDirectory + "\\Result-"+PlotTitle+".rtl";
            try {
                using (StreamWriter stream = new StreamWriter(path)) {
                    stream.WriteLine("HEADER::");
                    stream.WriteLine("File type  : rtl");
                    stream.WriteLine("Owner      : Linear");
                    stream.WriteLine("Version no : V20.02.02");
                    stream.WriteLine("");
                    stream.WriteLine("TARGET DATA::");
                    stream.WriteLine("Filetype  : ISO");
                    stream.WriteLine("Target-count: "+$"{Stroke/Interval+1}");
                    stream.WriteLine("Targets :");
                    int i = 0;
                    while (i < Stroke / Interval + 1)
                    {
                        foreach (var value in DisplayGrid)
                        {
                            if (i % 5 == 0 && i != Stroke / Interval)
                            {
                                stream.Write(value.TargetPos.ToString(DisplayResolution));
                                stream.Write(" ");
                            }
                            else if ((i % 5 == 4 || i == Stroke / Interval)&& i % 5 != 0)//judge if it is the last value of one row
                            {
                                stream.Write(" ");
                                stream.Write(value.TargetPos.ToString(DisplayResolution));
                                stream.Write(" \r\n");
                            }
                            else if (i == Stroke / Interval && i % 5 == 0)
                            {
                                stream.Write(value.TargetPos.ToString(DisplayResolution));
                                stream.Write(" \r\n");
                            }
                            else
                            {
                                stream.Write(" ");
                                stream.Write(value.TargetPos.ToString(DisplayResolution));
                                stream.Write(" ");
                            }
                            i++;
                            if (i >= Stroke / Interval + 1)
                            break;
                        }
                        
                    }
                    if (DirectionMode == Direction.UniDirection)
                    {
                        stream.WriteLine("Flags: 0 1 2 0");
                    }
                    if (DirectionMode == Direction.BiDirection)
                    {
                        stream.WriteLine("Flags: 0 0 2 0");
                    }
                    stream.WriteLine("");
                    stream.WriteLine("USER-TEXT::");
                    stream.WriteLine("Machine:                ");
                    stream.WriteLine("Serial No:              ");
                    DateTime dt = DateTime.Now;
                    string format = "yyyy-MM-dd HH:mm:ss";
                    stream.WriteLine("Date:"+dt.ToString(format));
                    stream.WriteLine("By:"+_operatorName);
                    stream.WriteLine("Axis:"+_axisinfo);
                    stream.WriteLine("Location:"+_location);
                    stream.WriteLine("TITLE:"+PlotTitle);
                    stream.WriteLine("");
                    stream.WriteLine("RUNS::");
                    stream.WriteLine("Run-count:"+$"{Runs*2}");
                    stream.WriteLine("");
                    stream.WriteLine("DEVIATIONS::");
                    stream.WriteLine("Run Target Data:");
                    int j = 1, indexDisplay=1;
                    foreach (var value in DisplayGrid) {
                        if (DirectionMode == Direction.BiDirection)
                        {
                            if (value.Index % 2 == 1)
                            {
                                stream.WriteLine($"{value.Index}" + "   " + $"{j}" + "      " + $"{value.Error}");
                                j++;
                            }
                            else
                            {
                                j--;
                                stream.WriteLine($"{value.Index}" + "   " + $"{j}" + "      " + $"{value.Error}");
                            }
                        }
                        else
                        {
                                stream.WriteLine($"{indexDisplay}" + "   " + $"{j}" + "      " + $"{value.Error}");
                                j++;
                                if (j == Stroke / Interval + 2)
                                {
                                    while (j--!=1)
                                        {
                                        stream.WriteLine($"{indexDisplay + 1}" + "   " + $"{j}" + "      " + $"0");
                                    }
                                    j += 1;
                                    indexDisplay += 2;
                                }
                        }
                    }
                    //stream.WriteLine("");
                    //stream.WriteLine("ENVIRONMENT::");
                    //stream.WriteLine("Air temp   :");
                    //stream.WriteLine("Air Press  :");
                    //stream.WriteLine("Air Humid  :");
                    //stream.WriteLine("Mat temp 1 :");
                    //stream.WriteLine("Mat temp 2 :");
                    //stream.WriteLine("Mat temp 3 :");
                    //stream.WriteLine("Env factor :");
                    //stream.WriteLine("Exp coeff  :");
                    //stream.WriteLine("Date Time  :");
                    //stream.WriteLine("Final Data :");
                    stream.WriteLine("");
                    stream.WriteLine("EOF::");
                    stream.WriteLine("");
                }
                Log.Info("Result saved sucessfully");
            }
            catch (Exception ex) {
                Log.Error($"Fail to save result to file: {ex.Message}");
                return false;
            }
            return true;
        }
        #endregion Private Method

        #region Public Method
        public void AddData() {
            if (SelectedGridItem != null) {
                var i = DisplayGrid.IndexOf(SelectedGridItem);
                var result = new DisplayResult {                //Instantiate a temp displayresult to store index & targetpos
                    Index = _displaygrid[i].Index,
                    TargetPos = _displaygrid[i].TargetPos,
                    DirectionPositive = _displaygrid[i].DirectionPositive
                };
                double a1, a2, a3;
                double targetmicron = result.TargetPos;
                switch (RunMeasurementMode) {
                    case MeasurementMode.Repeatability:
                        if (double.TryParse(A1Displacement, out a1))
                        {
                            _sError = ((a1 - targetmicron) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sError, out tempError);
                            result.Error = tempError;
                        }
                        result.Straightness = 0;
                        result.Flatness = 0;
                        AddPointsToPlot(new DataPoint(result.TargetPos,result.Error), MeasurementMode.Repeatability, result.DirectionPositive);
                        break;
                    case MeasurementMode.Straightness:
                        if (double.TryParse(A1Displacement, out a1)) {
                            _sError = ((a1 - targetmicron) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sError, out tempError);
                            result.Error = tempError;
                        }
                        if (double.TryParse(A2Displacement, out a2)) {
                            _sStrightness= ((a1 - a2) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sStrightness, out tempStrightness);
                            result.Straightness = tempStrightness;
                        }
                        result.Flatness = 0;
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Error), MeasurementMode.Repeatability, result.DirectionPositive);
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Straightness), MeasurementMode.Straightness, result.DirectionPositive);
                        break;
                    case MeasurementMode.Flatness:
                        if (double.TryParse(A1Displacement, out a1)) {
                            _sError = ((a1 - targetmicron) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sError, out tempError);
                            result.Error = tempError;
                        }
                        result.Straightness = 0;
                        if (double.TryParse(A3Displacement, out a3)) {
                            _sFlatness = ((a1 - a3) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sFlatness, out tempFlatness);
                            result.Flatness = tempFlatness;
                        }
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Error), MeasurementMode.Repeatability, result.DirectionPositive);
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Flatness), MeasurementMode.Flatness, result.DirectionPositive);
                        break;
                    case MeasurementMode.StraightnessFlatness:
                        if (double.TryParse(A1Displacement, out a1)) {
                            _sError = ((a1 - targetmicron) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sError, out tempError);
                            result.Error = tempError;
                        }
                        if (double.TryParse(A2Displacement, out a2)) {
                            _sStrightness = ((a1 - a2) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sStrightness, out tempStrightness);
                            result.Straightness = tempStrightness;
                        }
                        if (double.TryParse(A3Displacement, out a3)) {
                            _sFlatness = ((a1 - a3) * 1000).ToString(DisplayResolution);
                            double.TryParse(_sFlatness, out tempFlatness);
                            result.Flatness = tempFlatness;
                        }
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Error), MeasurementMode.Repeatability, result.DirectionPositive);
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Straightness), MeasurementMode.Straightness, result.DirectionPositive);
                        AddPointsToPlot(new DataPoint(result.TargetPos, result.Flatness), MeasurementMode.Flatness, result.DirectionPositive);
                        break;
                }

                DisplayGrid[i] = result;                         //Transfer latest data to grid
                if (i + 1 < DisplayGrid.Count) {                 //Select next row
                    SelectedGridItem = DisplayGrid[i + 1];
                }
                else {
                    _lastrow = true;
                }
            }
        }

        public void ClearPoints() {
            foreach (var pointlist in _pointList) {
                pointlist.Clear();
            }
        }

        #endregion Public Method

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion INotifyPropertyChanged

        public struct RunSetting {
            public Direction RunDirection;
            public int Run;
            public double Interval;
            public double Stroke;
        }
    }

    public class DisplayResult {
        //For DataGrid Display
        public int Index { get; set; }
        public double TargetPos { get; set; }
        public double Error { get; set; }
        public double Straightness { get; set; }
        public double Flatness { get; set; }
        public bool DirectionPositive = true;
    }
}
