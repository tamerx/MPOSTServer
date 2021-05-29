using NetMQ;
using NetMQ.Sockets;
using NLog;
using System;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using MPOST;

namespace CinigazMPostServer
{
    class Program
    {


        private PowerUp PupMode = PowerUp.A;
        private Acceptor BillAcceptor = new Acceptor();
        private bool rejectedFlag=false;

        private int port = 0;
        public int Port
        {
            get
            {
                return port;
            }
            set
            {
                port = value;
            }
        }

        private String comPortName = "COM2";
        public String ComPortName
        {
            get
            {
                return comPortName;
            }
            set
            {
                comPortName = value;
            }
        }


        // Power - Power UP
        private const String POWERUP = "POWERUP";
        private const String NOTPOWERUP = "NOTPOWERUP";


        // Connection
        private const String CONNECTED = "CONNECTED";
        private const String CONNECTIONALREADYOPENED = "CONNECTIONALREADYOPENED";
        private const String CONNECTEDANDMONEYBOXOPEN = "CONNECTEDANDMONEYBOXOPEN";
        private const String DISCONNECTED = "DISCONNECTED";
        private const String CONNECTIONCLOSED = "CONNECTIONCLOSED";
        private const String CONNECTIONALREADYCLOSED = "CONNECTIONALREADYCLOSED";

        // Commands
        private const String COMMANDNOTUNDERSTOOD = "COMMANDNOTUNDERSTOOD";
        private const String READING = "READING";
        private const String CANNOTPROCESS = "CANNOTPROCESS";
        private const String CANNOTREADINGDEVICEANOTHERSTATE = "CANNOTREADINGDEVICEANOTHERSTATE";
        private const String NOMONEYINTHETREASURY = "NOMONEYINTHETREASURY";
        private const String MONEYGIVENBACK = "MONEYGIVENBACK";
        private const String MONEYHASBEENSAFE = "MONEYHASBEENSAFE";
        private const String DEVICEJAMMED = "DEVICEJAMMED";
        private const String CASHBOXATTACHED = "CASHBOXATTACHED";
        private const String CASHBOXFULLANDATTACHED = "CASHBOXFULLANDATTACHED";
        private const String CASHBOXNOTATTACHED = "CASHBOXNOTATTACHED";
        private const String REJECTED = "REJECTED";

        #region Event Delegate Definitions

        private CalibrateFinishEventHandler CalibrateFinishDelegate;
        private CalibrateProgressEventHandler CalibrateProgressDelegate;
        private CalibrateStartEventHandler CalibrateStartDelegate;
        private CashBoxCleanlinessEventHandler CashBoxCleanlinessDelegate;
        private CashBoxAttachedEventHandler CashBoxAttachedDelegate;
        private CashBoxRemovedEventHandler CashBoxRemovedDelegate;
        private CheatedEventHandler CheatedDelegate;
        private ConnectedEventHandler ConnectedDelegate;
        private DisconnectedEventHandler DisconnectedDelegate;
        private DownloadFinishEventHandler DownloadFinishDelegate;
        private DownloadProgressEventHandler DownloadProgressDelegate;
        private DownloadRestartEventHandler DownloadRestartDelegate;
        private DownloadStartEventHandler DownloadStartDelegate;
        private ErrorOnSendMessageEventHandler ErrorOnSendMessageDelegate;
        private EscrowEventHandler EscrowedDelegate;
        private FailureClearedEventHandler FailureClearedDelegate;
        private FailureDetectedEventHandler FailureDetectedDelegate;
        private InvalidCommandEventHandler InvalidCommandDelegate;
        private JamClearedEventHandler JamClearedDelegate;
        private JamDetectedEventHandler JamDetectedDelegate;
        private NoteRetrievedEventHandler NoteRetrievedDelegate;
        private PauseClearedEventHandler PauseClearedDelegate;
        private PauseDetectedEventHandler PauseDetectedDelegate;
        private PowerUpCompleteEventHandler PowerUpCompleteDelegate;
        private PowerUpEventHandler PowerUpDelegate;
        private PUPEscrowEventHandler PUPEscrowDelegate;
        private RejectedEventHandler RejectedDelegate;
        private ReturnedEventHandler ReturnedDelegate;
        private StackedEventHandler StackedDelegate;
        private StackerFullClearedEventHandler StackerFullClearedDelegate;
        private StackerFullEventHandler StackerFullDelegate;
        private StallClearedEventHandler StallClearedDelegate;
        private StallDetectedEventHandler StallDetectedDelegate;

        #endregion

        private void ListEvent(String e)
        {
            Console.WriteLine(e);
        }

        public Program()
        {

            try
            {
                CalibrateFinishDelegate = new CalibrateFinishEventHandler(HandleCalibrateFinishEvent);
                CalibrateProgressDelegate = new CalibrateProgressEventHandler(HandleCalibrateProgressEvent);
                CalibrateStartDelegate = new CalibrateStartEventHandler(HandleCalibrateStartEvent);
                CashBoxCleanlinessDelegate = new CashBoxCleanlinessEventHandler(HandleCashBoxCleanlinessEvent);
                CashBoxAttachedDelegate = new CashBoxAttachedEventHandler(HandleCashBoxAttachedEvent);
                CashBoxRemovedDelegate = new CashBoxRemovedEventHandler(HandleCashBoxRemovedEvent);
                CheatedDelegate = new CheatedEventHandler(HandleCheatedEvent);
                ConnectedDelegate = new ConnectedEventHandler(HandleConnectedEvent);
                DisconnectedDelegate = new DisconnectedEventHandler(HandleDisconnectedEvent);
                DownloadFinishDelegate = new DownloadFinishEventHandler(HandleDownloadFinishEvent);
                DownloadProgressDelegate = new DownloadProgressEventHandler(HandleDownloadProgressEvent);
                DownloadRestartDelegate = new DownloadRestartEventHandler(HandleDownloadRestartEvent);
                DownloadStartDelegate = new DownloadStartEventHandler(HandleDownloadStartEvent);
                ErrorOnSendMessageDelegate = new ErrorOnSendMessageEventHandler(HandleSendMessageErrorEvent);
                EscrowedDelegate = new EscrowEventHandler(HandleEscrowedEvent);
                FailureClearedDelegate = new FailureClearedEventHandler(HandleFailureClearedEvent);
                FailureDetectedDelegate = new FailureDetectedEventHandler(HandleFailureDetectedEvent);
                InvalidCommandDelegate = new InvalidCommandEventHandler(HandleInvalidCommandEvent);
                JamClearedDelegate = new JamClearedEventHandler(HandleJamClearedEvent);
                JamDetectedDelegate = new JamDetectedEventHandler(HandleJamDetectedEvent);
                NoteRetrievedDelegate = new NoteRetrievedEventHandler(HandleNoteRetrievedEvent);
                PauseClearedDelegate = new PauseClearedEventHandler(HandlePauseClearedEvent);
                PauseDetectedDelegate = new PauseDetectedEventHandler(HandlePauseDetectedEvent);
                PowerUpCompleteDelegate = new PowerUpCompleteEventHandler(HandlePowerUpCompleteEvent);
                PowerUpDelegate = new PowerUpEventHandler(HandlePowerUpEvent);
                PUPEscrowDelegate = new PUPEscrowEventHandler(HandlePUPEscrowEvent);
                RejectedDelegate = new RejectedEventHandler(HandleRejectedEvent);
                ReturnedDelegate = new ReturnedEventHandler(HandleReturnedEvent);
                StackedDelegate = new StackedEventHandler(HandleStackedEvent);
                StackerFullClearedDelegate = new StackerFullClearedEventHandler(HandleStackerFullClearedEvent);
                StackerFullDelegate = new StackerFullEventHandler(HandleStackerFullEvent);
                StallClearedDelegate = new StallClearedEventHandler(HandleStallClearedEvent);
                StallDetectedDelegate = new StallDetectedEventHandler(HandleStallDetectedEvent);

                // Connect to the events.
                BillAcceptor.OnCalibrateFinish += CalibrateFinishDelegate;
                BillAcceptor.OnCalibrateProgress += CalibrateProgressDelegate;
                BillAcceptor.OnCalibrateStart += CalibrateStartDelegate;
                BillAcceptor.OnCashBoxCleanlinessDetected += CashBoxCleanlinessDelegate;
                BillAcceptor.OnCashBoxAttached += CashBoxAttachedDelegate;
                BillAcceptor.OnCashBoxRemoved += CashBoxRemovedDelegate;
                BillAcceptor.OnCheated += CheatedDelegate;
                BillAcceptor.OnConnected += ConnectedDelegate;
                BillAcceptor.OnDisconnected += DisconnectedDelegate;
                BillAcceptor.OnDownloadFinish += DownloadFinishDelegate;
                BillAcceptor.OnDownloadProgress += DownloadProgressDelegate;
                BillAcceptor.OnDownloadRestart += DownloadRestartDelegate;
                BillAcceptor.OnDownloadStart += DownloadStartDelegate;
                BillAcceptor.OnSendMessageFailure += ErrorOnSendMessageDelegate;
                BillAcceptor.OnEscrow += EscrowedDelegate;
                BillAcceptor.OnFailureCleared += FailureClearedDelegate;
                BillAcceptor.OnFailureDetected += FailureDetectedDelegate;
                BillAcceptor.OnInvalidCommand += InvalidCommandDelegate;
                BillAcceptor.OnJamCleared += JamClearedDelegate;
                BillAcceptor.OnJamDetected += JamDetectedDelegate;
                BillAcceptor.OnNoteRetrieved += NoteRetrievedDelegate;
                BillAcceptor.OnPauseCleared += PauseClearedDelegate;
                BillAcceptor.OnPauseDetected += PauseDetectedDelegate;
                BillAcceptor.OnPowerUpComplete += PowerUpCompleteDelegate;
                BillAcceptor.OnPowerUp += PowerUpDelegate;
                BillAcceptor.OnPUPEscrow += PUPEscrowDelegate;
                BillAcceptor.OnRejected += RejectedDelegate;
                BillAcceptor.OnReturned += ReturnedDelegate;
                BillAcceptor.OnStacked += StackedDelegate;
                BillAcceptor.OnStackerFullCleared += StackerFullClearedDelegate;
                BillAcceptor.OnStackerFull += StackerFullDelegate;
                BillAcceptor.OnStallCleared += StallClearedDelegate;
                BillAcceptor.OnStallDetected += StallDetectedDelegate;

            }
            catch (Exception ex)
            {
                ListEvent("Constructor Hatasi" + ex.ToString());
            }
        }

        private static void Main(string[] args)
        {

            Program program = new Program();
            Console.Clear();
            Console.WriteLine(" Açılıyor ");

            if (args.Length <= 1)
            {
                Console.WriteLine("Port ve COM port bilgilerini girmelisiniz");
                Console.WriteLine("Çıkılıyor");
                Environment.Exit(0);
            }
            else
            {
                if (int.TryParse(args[0], out program.port))
                {
                    program.ComPortName = args[1];
                    Console.WriteLine("Parametreler Uygun Formatta");
                }
                else
                {
                    Console.WriteLine("Girdiğiniz port bilgisinin 0 - 65536 arasında kullanılmayan bir sayı olması gerekiyor");
                    Environment.Exit(0);
                }
            }

            program.connectToZeroMQServer(program.Port);
            Console.ReadLine();

        }

        public void connectToZeroMQServer(int port)
        {

            Console.WriteLine(" Zero MQ Server {0} :: portunda baglanti kuruyor... ", port);

            String hostName = string.Concat("tcp://localhost:", port);
            ResponseSocket responseSocket = null;
            try
            {
                ResponseSocket responseSocket1 = new ResponseSocket(hostName);
                responseSocket = responseSocket1;
                using (responseSocket1)
                {
                    Console.WriteLine(" Zero MQ baglantisi {0}  portunda  acildi ", port);
                    while (true)
                    {
                        string receivedMessage = responseSocket.ReceiveFrameString();
                        Console.WriteLine("Mesaj alindi :: {0}", receivedMessage);
                        Console.WriteLine(receivedMessage);
                        string statusMessage = getStatus(receivedMessage);
                        responseSocket.SendFrame(statusMessage, false);
                        Console.WriteLine("Cevap verildi :: {0}", statusMessage);
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("islemde hata olustu");
                Console.WriteLine(exception);
            }

        }

        public void connectToKiosk(String portName, PowerUp pupMode)
        {

            try
            {
                ListEvent("Connecting to Kiosk...");
                ListEvent("PortName :: " + portName);
                ListEvent("PupMode :: " + PupMode.ToString());

                BillAcceptor.Open(portName, pupMode);
            }
            catch (Exception err)
            {
                ListEvent(err.ToString());
            }

        }

        private string getStatus(string m1)
        {
            string reply = string.Empty;
            string str = m1.Split(new char[] { ' ' })[0];
            try
            {
                if (str == null || str == "")
                {
                    reply = COMMANDNOTUNDERSTOOD;
                    return reply;
                }

                if (str == "Ac")
                {
                    if(BillAcceptor.Connected == false){

                        try
                        {
                            connectToKiosk(ComPortName, PupMode);
                            Console.WriteLine("Cevap Bekleniyor...");
                            Thread.Sleep(2000); // 2 sn bekleme süresi sonrası cevap veriliyor

                            if (BillAcceptor.Connected == true)
                            {
                                reply = CONNECTED;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(string.Concat(new string[] { "Unable to open the bill acceptor on com port :: ", ComPortName, " :: ", ex.Message, "Open Bill Acceptor Error" }));
                            reply = DISCONNECTED;
                            throw ex;
                        }

                    }
                    else if (BillAcceptor.Connected == true)
                    {
                        reply = CONNECTIONALREADYOPENED;
                    }
                   
                }


                if (str == "Kapat")
                {
                    if (BillAcceptor.Connected)
                    {
                        BillAcceptor.Close();
                        reply = CONNECTIONCLOSED;
                    }
                    else if (BillAcceptor.Connected == false)
                    {
                        reply = CONNECTIONALREADYCLOSED;
                    }
                }

                if (str == "Acikmi")
                {
                    PowerUp powerUp = BillAcceptor.DevicePowerUp;

                    if (powerUp == PowerUp.A)
                    {
                        reply = POWERUP;
                    }
                    else
                    {
                        reply = NOTPOWERUP;
                    }

                }

                if (str == "Baglan" || str == "Baglandimi")
                {
                    if (BillAcceptor.Connected)
                    {
                        BillAcceptor.EnableAcceptance = true;
                        reply = CONNECTEDANDMONEYBOXOPEN;

                    }
                    else if (BillAcceptor.Connected == false)
                    {
                        reply = DISCONNECTED;
                    }
                }


                if (str == "Oku")
                {


                 
                   

                    if (BillAcceptor.Connected)
                    {
                        BillAcceptor.EnableAcceptance = true;
                    }
                    else if (BillAcceptor.Connected == false)
                    {
                        reply = DISCONNECTED;
                    }

                    // bağlantı açık ve para kasası açık konumda ise  okuma işlemi gerçekleşir.
                    if (BillAcceptor.Connected && BillAcceptor.EnableAcceptance)
                    {
                        State deviceState = BillAcceptor.DeviceState;


                        if (deviceState == State.Accepting)
                        {
                            reply = READING;
                            if (rejectedFlag)
                            {
                                reply = REJECTED;
                                return reply;
                            }
                        }
                        else if (deviceState != State.Idling)
                        {
                            if (deviceState != State.Escrow)
                            {
                                reply = CANNOTREADINGDEVICEANOTHERSTATE;
                            }
                            else if (deviceState == State.Escrow)
                            {
                                reply = DocInfoToString();
                            }
                        }
                        else
                        {
                            reply = CANNOTPROCESS;
                        }
                    }
                    else
                    {
                        reply = DISCONNECTED;
                    }

                }

                if (str == "KasaTakilimi")
                {
                    if (BillAcceptor.Connected)
                    {
                        if (BillAcceptor.CashBoxAttached)
                        {
                            if (BillAcceptor.CashBoxFull)
                            {
                                reply = CASHBOXFULLANDATTACHED;
                            }
                            else
                            {
                                reply = CASHBOXATTACHED;
                            }
                        }
                        else
                        {
                            reply = CASHBOXNOTATTACHED;
                        }

                    }
                    else
                    {
                        reply = DISCONNECTED;
                    }
                }

                if (str == "KasayaAl")
                {
                    // bağlantı açık ve para kasası açık konumda ise  okuma işlemi gerçekleşir.
                    if (BillAcceptor.Connected && BillAcceptor.EnableAcceptance)
                    {
                        if (BillAcceptor.DeviceState != State.Escrow)
                        {
                            reply = NOMONEYINTHETREASURY;
                        }
                        else
                        {
                            if (BillAcceptor.CashBoxAttached)
                            {
                                if (BillAcceptor.CashBoxFull == false)
                                {
                                    if (BillAcceptor.DeviceJammed == false && BillAcceptor.DeviceState == State.Escrow)
                                    {
                                        Console.WriteLine("DeviceState::" + BillAcceptor.DeviceState.ToString());
                                        BillAcceptor.EnableAcceptance = true;
                                        BillAcceptor.EscrowStack();
                                        BillAcceptor.EnableAcceptance = false;
                                        reply = MONEYHASBEENSAFE;
                                    }
                                    else if (BillAcceptor.DeviceJammed == true)
                                    {
                                        reply = DEVICEJAMMED;
                                    }
                                }
                                else if (BillAcceptor.CashBoxFull == true)
                                {
                                    reply = CASHBOXFULLANDATTACHED;
                                }
                            }
                            else if (BillAcceptor.CashBoxAttached == false)
                            {
                                reply = CASHBOXNOTATTACHED;
                            }
                        }

                    }
                    else
                    {
                        reply = DISCONNECTED;
                    }
                }



                if (str == "GeriVer")
                {
                    // bağlantı açık ve para kasası açık konumda ise  okuma işlemi gerçekleşir.
                    if (BillAcceptor.Connected)
                    {
                        if (BillAcceptor.DeviceState != State.Escrow)
                        {
                            reply = NOMONEYINTHETREASURY;
                        }
                        else
                        {
                            if (BillAcceptor.DeviceJammed == false &&  BillAcceptor.DeviceState == State.Escrow)
                            {

                                Console.WriteLine("DeviceState::" + BillAcceptor.DeviceState.ToString());
                                BillAcceptor.EnableAcceptance = true;
                                BillAcceptor.EscrowReturn();
                                 // para kasasını kapatır , mavi ışıkları söndürür
                                 BillAcceptor.EnableAcceptance = false;
                                reply = MONEYGIVENBACK;
                            }
                            else if (BillAcceptor.DeviceJammed == true)
                            {
                                reply = DEVICEJAMMED;

                            }
                        }
                    }
                    else
                    {
                        reply = DISCONNECTED;
                    }

                    
                }

            }
            catch (Exception exception2)
            {
                reply = exception2.Message;
            }
            return reply;
        }

        private void PopulateProperties()
        {
            if (BillAcceptor.CapEasitrax)
            {
                ListEvent(BillAcceptor.AssetNumber.ToString());
            }
        }

        private void PopulateInfo()
        {
            try
            {
                if (BillAcceptor.CapDeviceType)
                {
                    ListEvent("BillAcceptor.CapDeviceType" + BillAcceptor.DeviceType);
                }
                else
                {
                    ListEvent("BillAcceptor.CapDeviceType Not supported");
                }
                try
                {
                    String t = BillAcceptor.DeviceSerialNumber;
                    ListEvent("illAcceptor.DeviceSerialNumber" + t);
                }
                catch
                {
                    ListEvent("BillAcceptor.DeviceSerialNumber  Not supported");
                }

                if (BillAcceptor.CapApplicationPN)
                {
                    String t = BillAcceptor.ApplicationPN;

                    ListEvent("BillAcceptor.ApplicationPN" + t);

                    if (BillAcceptor.CapApplicationID)
                    {
                        t += "/" + BillAcceptor.ApplicationID;
                    }
                    else
                    {
                        t += "/Not supported";
                    }
                    ListEvent("AppPN.Text" + t);
                }
                else ListEvent("AppPN.Text" + "not Supported");

                if (BillAcceptor.CapVariantPN)
                {
                    String t = BillAcceptor.VariantPN;

                    if (BillAcceptor.CapVariantID)
                    {
                        t += "/" + BillAcceptor.VariantID;
                    }
                    else
                    {
                        t += "/Not supported";
                    }
                    t = "";

                    foreach (String v in BillAcceptor.VariantNames)
                    {
                        t = t + v + " ";
                    }

                }
                else
                {

                }

                if (BillAcceptor.CapBootPN)
                {
                    ListEvent("BillAcceptor.BootPN" + BillAcceptor.BootPN);
                }
                else
                {
                    ListEvent("BillAcceptor.BootPN" + "Not supported");
                }

                ListEvent("0x" + BillAcceptor.DeviceCRC.ToString("X"));

                if (BillAcceptor.CapBNFStatus)
                {
                    if (BillAcceptor.BNFStatus == BNFStatus.Unknown)
                    {
                        ListEvent("BNFState.Text = Unknown");
                    }
                    else if (BillAcceptor.BNFStatus == BNFStatus.OK)
                    {
                        ListEvent("BNFState.Text = OK");
                    }
                    else if (BillAcceptor.BNFStatus == BNFStatus.Error)
                    {
                        ListEvent("BNFState.Text = Error");
                    }
                    else if (BillAcceptor.BNFStatus == BNFStatus.NotAttached)
                    {
                        ListEvent("BNFState.Text = Not attached");
                    }
                }
                else
                {
                    ListEvent("BNFState.Text = Not supported");
                }

                if (BillAcceptor.CapAudit)
                {

                    ListEvent(BillAcceptor.CapAudit.ToString());
                    Int32[] t = BillAcceptor.AuditLifeTimeTotals;

                    t = BillAcceptor.AuditQP;
                    ListEvent(BillAcceptor.AuditQP.ToString());

                    t = BillAcceptor.AuditPerformance;
                    ListEvent(BillAcceptor.AuditPerformance.ToString());

                }
                else { }

                if (BillAcceptor.CapCashBoxTotal)
                {
                    ListEvent("BillAcceptor.CapCashBoxTotal" + BillAcceptor.CapCashBoxTotal.ToString());

                }
                else
                {

                }

                if (BillAcceptor.CapDeviceResets)
                {
                    ListEvent("BillAcceptor.CapCashBoxTotal" + BillAcceptor.DeviceResets.ToString());
                }
                else
                {
                    ListEvent("BillAcceptor.CapCashBoxTotal" + "Not supported");
                }
                if (BillAcceptor.CashBoxAttached)
                {
                    ListEvent("BillAcceptor.CashBoxAttached ::" + BillAcceptor.CashBoxAttached.ToString());
                    if (BillAcceptor.CashBoxFull)
                    {
                        ListEvent("BillAcceptor.CashBoxFull" + BillAcceptor.CashBoxFull.ToString());
                    }
                    else
                    {
                        ListEvent(" CashBox:: not installed ");
                    }
                }
                else
                {
                    ListEvent(" CashBox:: Removed");
                }

                String billpath = (BillAcceptor.DeviceJammed) ? "Jammed" : "Clear";
                ListEvent("billpath :: " + billpath);

                if (BillAcceptor.DeviceModel < 32)
                {
                    ListEvent("BillAcceptor.DeviceModel :: " + BillAcceptor.DeviceModel.ToString());
                }
                else
                {
                    ListEvent(BillAcceptor.DeviceModel.ToString() + " (" + ((Char)BillAcceptor.DeviceModel).ToString() + ")");
                }

                Int32 rev = BillAcceptor.DeviceRevision;

                ListEvent((rev / 10).ToString() + "." + (rev % 10).ToString());
            }
            catch (Exception err)
            {
                ListEvent("Populate Info Exception");
                ListEvent(err.ToString());
            }
        }

        private String DocInfoToString()
        {
            if (BillAcceptor.DocType == DocumentType.None) return "Doc Type: None";
            else if (BillAcceptor.DocType == DocumentType.NoValue) return "Doc Type: No Value";
            else if (BillAcceptor.DocType == DocumentType.Bill)
            {
                if (BillAcceptor.Bill == null) return "Doc Type Bill = null";
                else if (!BillAcceptor.CapOrientationExt) return "Doc Type Bill = " + BillAcceptor.Bill.ToString();
                else return "Doc Type Bill = " + BillAcceptor.Bill.ToString() + " (" + BillAcceptor.EscrowOrientation.ToString() + ")";
            }
            else if (BillAcceptor.DocType == DocumentType.Barcode)
            {
                if (BillAcceptor.BarCode == null) return "Doc Type Bar Code = null";
                else return "Doc Type Bar Code = " + BillAcceptor.BarCode;
            }
            else if (BillAcceptor.DocType == DocumentType.Coupon)
            {
                if (BillAcceptor.Coupon == null) return "Doc Type Coupon = null";
                else return "Doc Type Coupon = " + BillAcceptor.Coupon.ToString();
            }
            else return "Unknown Doc Type Error";
        }

        #region Event Handlers

        private void HandleCalibrateFinishEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Calibrate Finish.");
        }
        private void HandleCalibrateProgressEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Calibrate Progress.");
        }
        private void HandleCalibrateStartEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Calibrate Start.");

        }
        private void HandleCashBoxCleanlinessEvent(object sender, CashBoxCleanlinessEventArgs e)
        {

            ListEvent(string.Format("Event: Cashbox Cleanliness - {0}", e.Value.ToString()));

        }
        private void HandleCashBoxAttachedEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Cassette Installed.");

        }
        private void HandleCashBoxRemovedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Cassette Removed.");

        }
        private void HandleCheatedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Cheat Detected.");

        }
        private void HandleConnectedEvent(object sender, EventArgs e)
        {

            PopulateCapabilities();
            PopulateBillSet();
            PopulateBillValue();
            PopulateProperties();
            PopulateInfo();
            BillAcceptor.EnableAcceptance = false;
            BillAcceptor.AutoStack = false;
            ListEvent("Event: Connected.");

        }
        private void HandleDisconnectedEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Disconnected.");
        }

        private void PopulateCapabilities()
        {
            string[] row;

            row = new string[] {
        "CapAdvBookmark",
        "False",
        "The advanced bookmark feature is available"
      };
            if (BillAcceptor.CapAdvBookmark) row[1] = "True";

            row = new string[] {
        "CapApplicationID",
        "False",
        "The application part number is available"
      };
            if (BillAcceptor.CapApplicationID) row[1] = "True";

            row = new string[] {
        "CapApplicationPN",
        "False",
        "The application file's part number is available"
      };
            if (BillAcceptor.CapApplicationPN) row[1] = "True";

            row = new string[] {
        "CapAssetNumber",
        "False",
        "The asset number may be set."
      };
            if (BillAcceptor.CapAssetNumber) row[1] = "True";

            row = new string[] {
        "CapAudit",
        "False",
        "Audit data is available"
      };
            if (BillAcceptor.CapAudit) row[1] = "True";

            row = new string[] {
        "CapBarCodes",
        "False",
        "The unit supports bar coded documents"
      };
            if (BillAcceptor.CapBarCodes) row[1] = "True";

            row = new string[] {
        "CapBarCodesExt",
        "False",
        "Extended bar codes are supported"
      };
            if (BillAcceptor.CapBarCodesExt) row[1] = "True";

            row = new string[] {
        "CapBNFStatus",
        "False",
        "The BNFStatus property is supported"
      };
            if (BillAcceptor.CapBNFStatus) row[1] = "True";

            row = new string[] {
        "CapBookmark",
        "False",
        "Bookmark documents are supported"
      };
            if (BillAcceptor.CapBookmark) row[1] = "True";

            row = new string[] {
        "CapBootPN",
        "False",
        "The bootloader part number is available"
      };
            if (BillAcceptor.CapBootPN) row[1] = "True";

            row = new string[] {
        "CapCalibrate",
        "False",
        "The unit may be calibrated"
      };
            if (BillAcceptor.CapCalibrate) row[1] = "True";

            row = new string[] {
        "CapCashBoxTotal",
        "False",
        "The unit supports a cash box total counter"
      };
            if (BillAcceptor.CapCashBoxTotal) row[1] = "True";

            row = new string[] {
        "CapCouponExt",
        "False",
        "The unit supports a extended generic coupons"
      };
            if (BillAcceptor.CapCouponExt) row[1] = "True";

            row = new string[] {
        "CapDevicePaused",
        "False",
        "The unit supports the paused state"
      };
            if (BillAcceptor.CapDevicePaused) row[1] = "True";

            row = new string[] {
        "CapDeviceSoftReset",
        "False",
        "The unit supports the soft reset command"
      };
            if (BillAcceptor.CapDeviceSoftReset) row[1] = "True";

            row = new string[] {
        "CapDeviceType",
        "False",
        "The unit reports its device type"
      };
            if (BillAcceptor.CapDeviceType) row[1] = "True";

            row = new string[] {
        "CapDeviceResets",
        "False",
        "The unit reports its reset counter"
      };
            if (BillAcceptor.CapDeviceResets) row[1] = "True";

            row = new string[] {
        "CapDeviceSerialNumber",
        "False",
        "The unit reports its serial number"
      };
            if (BillAcceptor.CapDeviceSerialNumber) row[1] = "True";

            row = new string[] {
        "CapEasiTrax",
        "False",
        "EasiTrax is supported"
      };
            if (BillAcceptor.CapEasitrax) row[1] = "True";

            row = new string[] {
        "CapEscrowTimeout",
        "False",
        "The unit supports the escrow timeout command"
      };
            if (BillAcceptor.CapEscrowTimeout) row[1] = "True";

            row = new string[] {
        "CapFlashDownload",
        "False",
        "The unit supports flash download"
      };
            if (BillAcceptor.CapFlashDownload) row[1] = "True";

            row = new string[] {
        "CapNoPush",
        "False",
        "The unit supports no_push mode"
      };
            if (BillAcceptor.CapNoPush) row[1] = "True";

            row = new string[] {
        "CapNoteRetrieved",
        "False",
        "The unit supports reporting when user takes rejected/returned notes"
      };
            if (BillAcceptor.CapNoteRetrieved) row[1] = "True";

            row = new string[] {
        "CapOrientationExt",
        "False",
        "The unit supports extended handling of bill orientation"
      };
            if (BillAcceptor.CapOrientationExt) row[1] = "True";

            row = new string[] {
        "CapPupExt",
        "False",
        "The unit supports extended PUP mode"
      };
            if (BillAcceptor.CapPupExt) row[1] = "True";

            row = new string[] {
        "CapSetBezel",
        "False",
        "The bezel may be configured"
      };
            if (BillAcceptor.CapSetBezel) row[1] = "True";

            row = new string[] {
        "CapTestDoc",
        "False",
        "Special Test Documents are supported"
      };
            if (BillAcceptor.CapTestDoc) row[1] = "True";

            row = new string[] {
        "CapVariantID",
        "False",
        "The variant part number is available"
      };
            if (BillAcceptor.CapVariantID) row[1] = "True";

            row = new string[] {
        "CapVariantPN",
        "False",
        "The variant file's part number is available"
      };
            if (BillAcceptor.CapVariantPN) row[1] = "True";

        }

        private void HandleDownloadFinishEvent(object sender, AcceptorDownloadFinishEventArgs e)
        {

            if (e.Success)
            {
                ListEvent("Event: Download Finished: OK");
            }
            else
            {
                ListEvent("Event: Download Finished: FAILED");
            }

        }
        private void HandleDownloadProgressEvent(object sender, AcceptorDownloadEventArgs e)
        {

            if (e.SectorCount % 100 == 0)
            {
                ListEvent("Event: Download Progress:" + e.SectorCount.ToString());
            }

        }
        private void HandleDownloadRestartEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Download Restart.");

        }
        private void HandleDownloadStartEvent(object sender, AcceptorDownloadEventArgs e)
        {

            ListEvent("Event: Download Start:" + e.SectorCount.ToString());

        }

        private void HandleSendMessageErrorEvent(object sender, AcceptorMessageEventArgs e)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append("Event: Error in send message. ");
            sb.Append(e.Msg.Description);
            sb.Append("  ");

            foreach (byte b in e.Msg.Payload)
            {
                sb.Append(b.ToString("X2") + " ");
            }

            ListEvent(sb.ToString());

            if (BillAcceptor.DeviceState == State.Escrow)
            {

            }

        }
        private void HandleEscrowedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Escrowed: " + DocInfoToString());

        }
        private void HandleFailureClearedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Device Failure Cleared. ");

        }
        private void HandleFailureDetectedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Device Failure Detected. ");

        }
        private void HandleInvalidCommandEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Invalid Command.");

        }
        private void HandleJamClearedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Jam Cleared.");

        }
        private void HandleJamDetectedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Jam Detected.");

        }
        private void HandleNoteRetrievedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Note Retreived.");

        }
        private void HandlePauseClearedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Pause Cleared.");

        }
        private void HandlePauseDetectedEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Pause Detected.");

        }
        private void HandlePowerUpCompleteEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Power Up Complete.");

        }
        private void HandlePowerUpEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Power Up.");

        }
        private void HandlePUPEscrowEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Power Up with Escrow: " + DocInfoToString());

        }
        private void HandleRejectedEvent(object sender, EventArgs e)
        {

            rejectedFlag = true;
            ListEvent("Event: Rejected.");

        }
        private void HandleReturnedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Returned.");

        }
        private void HandleStackedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Stacked: " + DocInfoToString());

            if (BillAcceptor.CapCashBoxTotal)
            {

            }

        }
        private void HandleStackerFullClearedEvent(object sender, EventArgs e)
        {

            ListEvent("Event: Cassette Full Cleared.");

        }
        private void HandleStackerFullEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Cassette Full.");
        }

        private void HandleStallClearedEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Stall Cleared.");
        }

        private void HandleStallDetectedEvent(object sender, EventArgs e)
        {
            ListEvent("Event: Stall Detected.");
        }

        #endregion

        private void PopulateBillSet()
        {
            MPOST.Bill[] Bills = BillAcceptor.BillTypes;
            Boolean[] Enables = BillAcceptor.GetBillTypeEnables();

            for (int i = 0; i < Bills.Length; i++)
            {
                String[] row = new String[5];
                row[0] = (i + 1).ToString();
                row[1] = Bills[i].Country;
                row[2] = Bills[i].Value.ToString();
                row[3] = Bills[i].Type.ToString() + " " + Bills[i].Series.ToString() + " " + Bills[i].Compatibility.ToString() + " " + Bills[i].Version.ToString();
                row[4] = Enables[i] ? "True" : "False";

            }
        }

        private void PopulateBillValue()
        {
            MPOST.Bill[] Bills = BillAcceptor.BillValues;
            Boolean[] Enables = BillAcceptor.GetBillValueEnables();

            for (int i = 0; i < Bills.Length; i++)
            {
                String[] row = new String[5];
                row[0] = (i + 1).ToString();
                row[1] = Bills[i].Country;
                row[2] = Bills[i].Value.ToString();
                row[3] = Bills[i].Type.ToString() + " " + Bills[i].Series.ToString() + " " + Bills[i].Compatibility.ToString() + " " + Bills[i].Version.ToString();
                row[4] = Enables[i] ? "True" : "False";

            }
        }

    }
}