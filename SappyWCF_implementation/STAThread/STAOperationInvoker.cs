using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Dispatcher;
using System.Threading;
using System.ServiceModel;

namespace WcfSappy.STAThread
{
    internal class STAOperationInvoker : IOperationInvoker
    {
        // This is to cache threads
        // private static Dictionary<string, Thread> staThreads = new Dictionary<string, Thread>();

        IOperationInvoker _innerInvoker;
        public STAOperationInvoker(IOperationInvoker invoker)
        {
            _innerInvoker = invoker;
        }

        public object[] AllocateInputs()
        {
            return _innerInvoker.AllocateInputs();
        }

        public object Invoke(object instance, object[] inputs, out object[] outputs)
        {
            // Create a new, STA thread, invoke it
            object[] threadOutputs = null;
            Exception threadEx = null;
            object threadRet = null;


            ThreadStart threadDelegate = new ThreadStart(() =>
            {
                try
                {
                    threadRet = _innerInvoker.Invoke(instance, inputs, out threadOutputs);
                }
                catch (Exception ex)
                {
                    // If we dont catch this execption, since is on other thread, father tread wont be notified!!!
                    threadEx = ex;
                }
            });

            Thread thread = new Thread(threadDelegate);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (threadEx != null) //Check if there were errors
            {
                Logger.Log.Error(threadEx.Message, threadEx);
                throw new FaultException(threadEx.Message);
            }

            outputs = threadOutputs;
            return threadRet;

        }

        public IAsyncResult InvokeBegin(object instance, object[] inputs, AsyncCallback callback, object state)
        {
            // We don’t handle async…
            throw new NotImplementedException();
        }

        public object InvokeEnd(object instance, out object[] outputs, IAsyncResult result)
        {
            // We don’t handle async…
            throw new NotImplementedException();
        }

        public bool IsSynchronous
        {
            get { return true; }
        }
    }
}
