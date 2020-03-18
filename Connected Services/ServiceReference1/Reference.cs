﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ServiceReference1
{
    using System.Runtime.Serialization;
    
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.Runtime.Serialization.DataContractAttribute(Name="CompositeType", Namespace="http://schemas.datacontract.org/2004/07/Conductor")]
    public partial class CompositeType : object
    {
        
        private bool BoolValueField;
        
        private string StringValueField;
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public bool BoolValue
        {
            get
            {
                return this.BoolValueField;
            }
            set
            {
                this.BoolValueField = value;
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string StringValue
        {
            get
            {
                return this.StringValueField;
            }
            set
            {
                this.StringValueField = value;
            }
        }
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.ServiceContractAttribute(ConfigurationName="ServiceReference1.IService1")]
    public interface IService1
    {
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/GetData", ReplyAction="http://tempuri.org/IService1/GetDataResponse")]
        ServiceReference1.GetDataResponse GetData(ServiceReference1.GetDataRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/GetData", ReplyAction="http://tempuri.org/IService1/GetDataResponse")]
        System.Threading.Tasks.Task<ServiceReference1.GetDataResponse> GetDataAsync(ServiceReference1.GetDataRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/GetDataUsingDataContract", ReplyAction="http://tempuri.org/IService1/GetDataUsingDataContractResponse")]
        ServiceReference1.GetDataUsingDataContractResponse GetDataUsingDataContract(ServiceReference1.GetDataUsingDataContractRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/GetDataUsingDataContract", ReplyAction="http://tempuri.org/IService1/GetDataUsingDataContractResponse")]
        System.Threading.Tasks.Task<ServiceReference1.GetDataUsingDataContractResponse> GetDataUsingDataContractAsync(ServiceReference1.GetDataUsingDataContractRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/MoveFiles", ReplyAction="http://tempuri.org/IService1/MoveFilesResponse")]
        ServiceReference1.MoveFilesResponse MoveFiles(ServiceReference1.MoveFilesRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/MoveFiles", ReplyAction="http://tempuri.org/IService1/MoveFilesResponse")]
        System.Threading.Tasks.Task<ServiceReference1.MoveFilesResponse> MoveFilesAsync(ServiceReference1.MoveFilesRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/PostFileEventNotification", ReplyAction="http://tempuri.org/IService1/PostFileEventNotificationResponse")]
        ServiceReference1.PostFileEventNotificationResponse PostFileEventNotification(ServiceReference1.PostFileEventNotificationRequest request);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IService1/PostFileEventNotification", ReplyAction="http://tempuri.org/IService1/PostFileEventNotificationResponse")]
        System.Threading.Tasks.Task<ServiceReference1.PostFileEventNotificationResponse> PostFileEventNotificationAsync(ServiceReference1.PostFileEventNotificationRequest request);
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="GetData", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class GetDataRequest
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string value;
        
        public GetDataRequest()
        {
        }
        
        public GetDataRequest(string value)
        {
            this.value = value;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="GetDataResponse", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class GetDataResponse
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string GetDataResult;
        
        public GetDataResponse()
        {
        }
        
        public GetDataResponse(string GetDataResult)
        {
            this.GetDataResult = GetDataResult;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="GetDataUsingDataContract", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class GetDataUsingDataContractRequest
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public ServiceReference1.CompositeType composite;
        
        public GetDataUsingDataContractRequest()
        {
        }
        
        public GetDataUsingDataContractRequest(ServiceReference1.CompositeType composite)
        {
            this.composite = composite;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="GetDataUsingDataContractResponse", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class GetDataUsingDataContractResponse
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public ServiceReference1.CompositeType GetDataUsingDataContractResult;
        
        public GetDataUsingDataContractResponse()
        {
        }
        
        public GetDataUsingDataContractResponse(ServiceReference1.CompositeType GetDataUsingDataContractResult)
        {
            this.GetDataUsingDataContractResult = GetDataUsingDataContractResult;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="MoveFiles", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class MoveFilesRequest
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string frompath;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=1)]
        public string topath;
        
        public MoveFilesRequest()
        {
        }
        
        public MoveFilesRequest(string frompath, string topath)
        {
            this.frompath = frompath;
            this.topath = topath;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="MoveFilesResponse", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class MoveFilesResponse
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string MoveFilesResult;
        
        public MoveFilesResponse()
        {
        }
        
        public MoveFilesResponse(string MoveFilesResult)
        {
            this.MoveFilesResult = MoveFilesResult;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="PostFileEventNotification", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class PostFileEventNotificationRequest
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string notificationType;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=1)]
        public string filepath;
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=2)]
        public string postToUrl;
        
        public PostFileEventNotificationRequest()
        {
        }
        
        public PostFileEventNotificationRequest(string notificationType, string filepath, string postToUrl)
        {
            this.notificationType = notificationType;
            this.filepath = filepath;
            this.postToUrl = postToUrl;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    [System.ServiceModel.MessageContractAttribute(WrapperName="PostFileEventNotificationResponse", WrapperNamespace="http://tempuri.org/", IsWrapped=true)]
    public partial class PostFileEventNotificationResponse
    {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Namespace="http://tempuri.org/", Order=0)]
        public string PostFileEventNotificationResult;
        
        public PostFileEventNotificationResponse()
        {
        }
        
        public PostFileEventNotificationResponse(string PostFileEventNotificationResult)
        {
            this.PostFileEventNotificationResult = PostFileEventNotificationResult;
        }
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    public interface IService1Channel : ServiceReference1.IService1, System.ServiceModel.IClientChannel
    {
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.2")]
    public partial class Service1Client : System.ServiceModel.ClientBase<ServiceReference1.IService1>, ServiceReference1.IService1
    {
        
        /// <summary>
        /// Implement this partial method to configure the service endpoint.
        /// </summary>
        /// <param name="serviceEndpoint">The endpoint to configure</param>
        /// <param name="clientCredentials">The client credentials</param>
        static partial void ConfigureEndpoint(System.ServiceModel.Description.ServiceEndpoint serviceEndpoint, System.ServiceModel.Description.ClientCredentials clientCredentials);
        
        public Service1Client() : 
                base(Service1Client.GetDefaultBinding(), Service1Client.GetDefaultEndpointAddress())
        {
            this.Endpoint.Name = EndpointConfiguration.BasicHttpBinding_IService1.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }
        
        public Service1Client(EndpointConfiguration endpointConfiguration) : 
                base(Service1Client.GetBindingForEndpoint(endpointConfiguration), Service1Client.GetEndpointAddress(endpointConfiguration))
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }
        
        public Service1Client(EndpointConfiguration endpointConfiguration, string remoteAddress) : 
                base(Service1Client.GetBindingForEndpoint(endpointConfiguration), new System.ServiceModel.EndpointAddress(remoteAddress))
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }
        
        public Service1Client(EndpointConfiguration endpointConfiguration, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(Service1Client.GetBindingForEndpoint(endpointConfiguration), remoteAddress)
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }
        
        public Service1Client(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(binding, remoteAddress)
        {
        }
        
        public ServiceReference1.GetDataResponse GetData(ServiceReference1.GetDataRequest request)
        {
            return base.Channel.GetData(request);
        }
        
        public System.Threading.Tasks.Task<ServiceReference1.GetDataResponse> GetDataAsync(ServiceReference1.GetDataRequest request)
        {
            return base.Channel.GetDataAsync(request);
        }
        
        public ServiceReference1.GetDataUsingDataContractResponse GetDataUsingDataContract(ServiceReference1.GetDataUsingDataContractRequest request)
        {
            return base.Channel.GetDataUsingDataContract(request);
        }
        
        public System.Threading.Tasks.Task<ServiceReference1.GetDataUsingDataContractResponse> GetDataUsingDataContractAsync(ServiceReference1.GetDataUsingDataContractRequest request)
        {
            return base.Channel.GetDataUsingDataContractAsync(request);
        }
        
        public ServiceReference1.MoveFilesResponse MoveFiles(ServiceReference1.MoveFilesRequest request)
        {
            return base.Channel.MoveFiles(request);
        }
        
        public System.Threading.Tasks.Task<ServiceReference1.MoveFilesResponse> MoveFilesAsync(ServiceReference1.MoveFilesRequest request)
        {
            return base.Channel.MoveFilesAsync(request);
        }
        
        public ServiceReference1.PostFileEventNotificationResponse PostFileEventNotification(ServiceReference1.PostFileEventNotificationRequest request)
        {
            return base.Channel.PostFileEventNotification(request);
        }
        
        public System.Threading.Tasks.Task<ServiceReference1.PostFileEventNotificationResponse> PostFileEventNotificationAsync(ServiceReference1.PostFileEventNotificationRequest request)
        {
            return base.Channel.PostFileEventNotificationAsync(request);
        }
        
        public virtual System.Threading.Tasks.Task OpenAsync()
        {
            return System.Threading.Tasks.Task.Factory.FromAsync(((System.ServiceModel.ICommunicationObject)(this)).BeginOpen(null, null), new System.Action<System.IAsyncResult>(((System.ServiceModel.ICommunicationObject)(this)).EndOpen));
        }
        
        public virtual System.Threading.Tasks.Task CloseAsync()
        {
            return System.Threading.Tasks.Task.Factory.FromAsync(((System.ServiceModel.ICommunicationObject)(this)).BeginClose(null, null), new System.Action<System.IAsyncResult>(((System.ServiceModel.ICommunicationObject)(this)).EndClose));
        }
        
        private static System.ServiceModel.Channels.Binding GetBindingForEndpoint(EndpointConfiguration endpointConfiguration)
        {
            if ((endpointConfiguration == EndpointConfiguration.BasicHttpBinding_IService1))
            {
                System.ServiceModel.BasicHttpBinding result = new System.ServiceModel.BasicHttpBinding();
                result.MaxBufferSize = int.MaxValue;
                result.ReaderQuotas = System.Xml.XmlDictionaryReaderQuotas.Max;
                result.MaxReceivedMessageSize = int.MaxValue;
                result.AllowCookies = true;
                return result;
            }
            throw new System.InvalidOperationException(string.Format("Could not find endpoint with name \'{0}\'.", endpointConfiguration));
        }
        
        private static System.ServiceModel.EndpointAddress GetEndpointAddress(EndpointConfiguration endpointConfiguration)
        {
            if ((endpointConfiguration == EndpointConfiguration.BasicHttpBinding_IService1))
            {
                return new System.ServiceModel.EndpointAddress("http://localhost:8733/Design_Time_Addresses/Conductor/Service1/");
            }
            throw new System.InvalidOperationException(string.Format("Could not find endpoint with name \'{0}\'.", endpointConfiguration));
        }
        
        private static System.ServiceModel.Channels.Binding GetDefaultBinding()
        {
            return Service1Client.GetBindingForEndpoint(EndpointConfiguration.BasicHttpBinding_IService1);
        }
        
        private static System.ServiceModel.EndpointAddress GetDefaultEndpointAddress()
        {
            return Service1Client.GetEndpointAddress(EndpointConfiguration.BasicHttpBinding_IService1);
        }
        
        public enum EndpointConfiguration
        {
            
            BasicHttpBinding_IService1,
        }
    }
}
