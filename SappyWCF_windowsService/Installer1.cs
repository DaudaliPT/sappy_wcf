
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;


namespace SappyWCF_windowsService
{

    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        private ServiceProcessInstaller process;
        private ServiceInstaller service;

        public Installer1()
        {
            string parvals = "";
            foreach (string item in this.Context.Parameters.Keys)
            {
                string value = this.Context.Parameters[item];
                parvals += "\n" + item + "=" + value;
            }
            this.Context.LogMessage(parvals);
            InitializeComponent();
            this.process = new ServiceProcessInstaller();
            this.process.Account = ServiceAccount.LocalSystem;
            this.service = new ServiceInstaller();
            this.service.ServiceName = new SappyWCFService().ServiceName;
            this.service.Description = "Sappy WCF: Serviço REST para interação com Sap B1 usando DIAPI e Crystal Reportserviço REST para interação com Sap B1 usando DIAPI e Crystal Reports";
            this.service.StartType = ServiceStartMode.Automatic;
            Installers.Add(this.process);
            Installers.Add(this.service);
        }
    }
}
