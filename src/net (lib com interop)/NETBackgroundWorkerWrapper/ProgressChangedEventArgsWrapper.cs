using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace NETBackgroundWorkerWrapper
{
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IProgressChangedEventArgsWrapper
    {
        int ProgressPercentage { get; set; }
        object UserState { get; set; }
    }

    [Guid("B03B469E-4894-4222-BBBC-8C642A761156")]
    [ClassInterface(ClassInterfaceType.None)]
    //[ProgId("NETBackgroundWorkerWrapper.ProgressChangedEventArgsWrapper")]
    public class ProgressChangedEventArgsWrapper : IProgressChangedEventArgsWrapper
    {
        // A creatable COM class must have a Public Sub New() 
        // with no parameters, otherwise, the class will not be 
        // registered in the COM registry and cannot be created 
        // via CreateObject.
        public ProgressChangedEventArgsWrapper()
        {
        }

        public ProgressChangedEventArgsWrapper(int progressPercentage, object userState)
        {
            ProgressPercentage = progressPercentage;
            UserState = userState;
        }

        public ProgressChangedEventArgsWrapper(ProgressChangedEventArgs e)
        {
            ProgressPercentage = e.ProgressPercentage;
            UserState = e.UserState;
        }

        public int ProgressPercentage { get; set; }
        public object UserState { get; set; }
    }
}
