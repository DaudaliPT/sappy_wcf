using System; 
using System.ServiceModel;

namespace HostForDebug
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            { 
                Logger.Log.Debug("Starting service...");
                SBOHandler.Init();
                var serviceHost = new ServiceHost(typeof(SappyWcf));
                Logger.Log.Info("Starting state:"+serviceHost.State.ToString());
                serviceHost.Open();
                Logger.Log.Info("Starting state:" + serviceHost.State.ToString());

            }
            catch (System.Exception ex)
            {
                Logger.Log.Error("Exeption", ex);
            }

            System.Console.ReadLine();
        }
    }
}
