using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace NETBackgroundWorkerWrapper
{
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRunWorkerCompletedEventArgsWrapper
    {
        void SetResult(object result);
        object GetResult();
        ErrorInfo Error { get; set; }
        bool Cancelled { get; set; }
    }

    [Guid("27F54F98-0C2F-4C8E-9312-61036DA03DC1")]
    [ClassInterface(ClassInterfaceType.None)]
    public class RunWorkerCompletedEventArgsWrapper : IRunWorkerCompletedEventArgsWrapper
    {
        // A creatable COM class must have a Public Sub New() 
        // with no parameters, otherwise, the class will not be 
        // registered in the COM registry and cannot be created 
        // via CreateObject.
        public RunWorkerCompletedEventArgsWrapper()
        {
            _result = null;
            Error = new ErrorInfo();
            Cancelled = false;
        }

        public RunWorkerCompletedEventArgsWrapper(object result, ErrorInfo error, bool cancelled)
        {
            _result = result;
            Error = error;
            Cancelled = cancelled;
        }

        public RunWorkerCompletedEventArgsWrapper(RunWorkerCompletedEventArgs e)
        {
            _result = e.Result;
            Error.Number = -1;
            Error.Description = e.Error.Message;
            Cancelled = e.Cancelled;
        }

        private object _result;
        public void SetResult(object result)
        {
            _result = result;
        }

        public object GetResult()
        {
            return _result;
        }

        public ErrorInfo Error { get; set; }
        public bool Cancelled { get; set; }
    }
}
