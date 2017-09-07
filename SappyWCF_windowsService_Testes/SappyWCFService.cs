using System;
using System.ServiceProcess;
using System.ServiceModel;


namespace SappyWCF_windowsService
{
    public partial class SappyWCFService : ServiceBase
    { 
        private ServiceHost serviceHost = null;

        public SappyWCFService()
        {
            InitializeComponent();
            this.ServiceName = "Sappy WCF Service - Testes";
            this.AutoLog = true;
            this.CanPauseAndContinue = false; 
        }

        protected override void OnStart(string[] args)
        { 
            try
            { 
                Logger.Log.Debug("Starting service...");
                SBOHandler.Init();
                if (this.serviceHost != null) this.serviceHost.Close();
                var serviceHost = new ServiceHost(typeof(SappyWcf));
                Logger.Log.Info("Starting state:"+serviceHost.State.ToString());
                serviceHost.Open();
                Logger.Log.Info("Starting state:" + serviceHost.State.ToString());

            }
            catch (System.Exception ex)
            {
                Logger.Log.Error(ex.Message, ex);
            }  
        }

        protected override void OnStop()
        {
            if (this.serviceHost != null)
            {
                this.serviceHost.Close();
                this.serviceHost = null;
            }
            Logger.Log.Info("Service stopped.");
        }
    }
}
