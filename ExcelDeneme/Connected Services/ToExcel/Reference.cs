﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ExcelDeneme.ToExcel {
    using System.Runtime.Serialization;
    
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Name="Tema", Namespace="http://schemas.datacontract.org/2004/07/wcfExel")]
    public enum Tema : int {
        
        [System.Runtime.Serialization.EnumMemberAttribute()]
        Mavi = 1,
        
        [System.Runtime.Serialization.EnumMemberAttribute()]
        Yeşil = 2,
        
        [System.Runtime.Serialization.EnumMemberAttribute()]
        Kırmızı = 3,
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.ServiceContractAttribute(ConfigurationName="ToExcel.IExcel")]
    public interface IExcel {
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IExcel/GetData", ReplyAction="http://tempuri.org/IExcel/GetDataResponse")]
        string GetData(int value);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IExcel/GetData", ReplyAction="http://tempuri.org/IExcel/GetDataResponse")]
        System.Threading.Tasks.Task<string> GetDataAsync(int value);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IExcel/datatableToExcel", ReplyAction="http://tempuri.org/IExcel/datatableToExcelResponse")]
        byte[] datatableToExcel(string[] Baslik, System.Data.DataTable dataTable, ExcelDeneme.ToExcel.Tema t);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IExcel/datatableToExcel", ReplyAction="http://tempuri.org/IExcel/datatableToExcelResponse")]
        System.Threading.Tasks.Task<byte[]> datatableToExcelAsync(string[] Baslik, System.Data.DataTable dataTable, ExcelDeneme.ToExcel.Tema t);
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public interface IExcelChannel : ExcelDeneme.ToExcel.IExcel, System.ServiceModel.IClientChannel {
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public partial class ExcelClient : System.ServiceModel.ClientBase<ExcelDeneme.ToExcel.IExcel>, ExcelDeneme.ToExcel.IExcel {
        
        public ExcelClient() {
        }
        
        public ExcelClient(string endpointConfigurationName) : 
                base(endpointConfigurationName) {
        }
        
        public ExcelClient(string endpointConfigurationName, string remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public ExcelClient(string endpointConfigurationName, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public ExcelClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(binding, remoteAddress) {
        }
        
        public string GetData(int value) {
            return base.Channel.GetData(value);
        }
        
        public System.Threading.Tasks.Task<string> GetDataAsync(int value) {
            return base.Channel.GetDataAsync(value);
        }
        
        public byte[] datatableToExcel(string[] Baslik, System.Data.DataTable dataTable, ExcelDeneme.ToExcel.Tema t) {
            return base.Channel.datatableToExcel(Baslik, dataTable, t);
        }
        
        public System.Threading.Tasks.Task<byte[]> datatableToExcelAsync(string[] Baslik, System.Data.DataTable dataTable, ExcelDeneme.ToExcel.Tema t) {
            return base.Channel.datatableToExcelAsync(Baslik, dataTable, t);
        }
    }
}