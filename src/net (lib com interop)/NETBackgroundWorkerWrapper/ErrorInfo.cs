using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace NETBackgroundWorkerWrapper
{
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IErrorInfo
    {
        int Number { get; set; }
        string Description { get; set; }
    }

    [Guid("C53C7DA7-A1CD-449D-A4E6-DCF527DF70D3")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ErrorInfo : IErrorInfo
    {
        // A creatable COM class must have a Public Sub New() 
        // with no parameters, otherwise, the class will not be 
        // registered in the COM registry and cannot be created 
        // via CreateObject.
        public ErrorInfo()
        {
        }

        public int Number { get; set; }
        public string Description { get; set; }
    }
}
