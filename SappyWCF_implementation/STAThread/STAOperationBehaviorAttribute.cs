using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;

namespace WcfSappy.STAThread
{

    /// <summary>
    /// Must be defined on methods that implment OperationContract and must call COM STA 
    /// </summary>
    internal class STAOperationBehaviorAttribute : Attribute, IOperationBehavior
    {
        public void AddBindingParameters(OperationDescription operationDescription, BindingParameterCollection bindingParameters)
        {

        }

        public void ApplyClientBehavior(OperationDescription operationDescription, ClientOperation clientOperation)
        {
            // If this is applied on the client, well, it just doesn’t make sense.
            // Don’t throw in case this attribute was applied on the contract
            // instead of the implementation.
        }

        public void ApplyDispatchBehavior(OperationDescription operationDescription, DispatchOperation dispatchOperation)
        {
            // Change the IOperationInvoker for this operation.
            dispatchOperation.Invoker = new STAOperationInvoker(dispatchOperation.Invoker);
        }

        public void Validate(OperationDescription operationDescription)
        {
            if (operationDescription.SyncMethod == null)
            {
                throw new InvalidOperationException("The STAOperationBehaviorAttribute " +
                    "only works for synchronous method invocations.");
            }
        }
    }
}
