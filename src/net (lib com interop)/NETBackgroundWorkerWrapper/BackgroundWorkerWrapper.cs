using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace NETBackgroundWorkerWrapper
{
    [ComVisible(false)]
    public delegate void ProgressChangedEventHandler(object sender, ProgressChangedEventArgsWrapper e);

    [ComVisible(false)]
    public delegate void RunWorkerCompletedEventHandler(object sender, RunWorkerCompletedEventArgsWrapper e);

    [Guid("69F1B4B6-D748-467A-902A-2EDE6009870E")]
    public interface IBackgroundWorkerWrapper
    {
        void RunWorkerSync(IntPtr callback, object argument);
        void RunWorkerAsync(IntPtr callback, object argument);
        void ReportProgress(int percentProgress, object userState);
        void CancelAsync();
        bool CancellationPending { get; }
        bool IsBusy { get; }
    }

    [Guid("EB80692E-FC7B-4C38-A159-B050323FDA37")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IBackgroundWorkerWrapperEvents
    {
        void ProgressChanged(object sender, ProgressChangedEventArgsWrapper e);
        void RunWorkerCompleted(object sender, RunWorkerCompletedEventArgsWrapper e);
    }

    [Guid("E246F042-8104-44B7-BC2F-B207E9B6D8DD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IBackgroundWorkerWrapperEvents))]
    public class BackgroundWorkerWrapper : IBackgroundWorkerWrapper
    {
        private delegate void DoWork(ref object argument, ref RunWorkerCompletedEventArgsWrapper e);        

        public event ProgressChangedEventHandler ProgressChanged;
        public event RunWorkerCompletedEventHandler RunWorkerCompleted;

        private BackgroundWorker _backgroundWorker;
        private IntPtr callback;
        private RunWorkerCompletedEventArgsWrapper results;

        // A creatable COM class must have a Public Sub New() 
        // with no parameters, otherwise, the class will not be 
        // registered in the COM registry and cannot be created 
        // via CreateObject.

        public BackgroundWorkerWrapper() : base()
        {
            BackgroundWorker = new BackgroundWorker();
            BackgroundWorker.WorkerReportsProgress = true;
            BackgroundWorker.WorkerSupportsCancellation = true;
        }

        private BackgroundWorker BackgroundWorker
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _backgroundWorker;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_backgroundWorker != null)
                {
                    _backgroundWorker.DoWork -= DoWorkEventHandler;
                    _backgroundWorker.ProgressChanged -= ProgressChangedEventHandler;
                    _backgroundWorker.RunWorkerCompleted -= RunWorkerCompletedEventHandler;
                }

                _backgroundWorker = value;
                if (_backgroundWorker != null)
                {
                    _backgroundWorker.DoWork += DoWorkEventHandler;
                    _backgroundWorker.ProgressChanged += ProgressChangedEventHandler;
                    _backgroundWorker.RunWorkerCompleted += RunWorkerCompletedEventHandler;
                }
            }
        }

        public bool CancellationPending
        {
            get { return BackgroundWorker.CancellationPending; }
        }

        public bool IsBusy
        {
            get { return BackgroundWorker.IsBusy; }
        }
        
        # region event handler

        private void DoWorkEventHandler(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var worker = (DoWork)Marshal.GetDelegateForFunctionPointer(callback, typeof(DoWork));
            var r = new RunWorkerCompletedEventArgsWrapper();

            var argument = e.Argument;
            worker?.Invoke(ref argument, ref r);

            this.results = r;
        }

        private void ProgressChangedEventHandler(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            ProgressChanged?.Invoke(sender, new ProgressChangedEventArgsWrapper(e));
        }

        private void RunWorkerCompletedEventHandler(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            RunWorkerCompleted?.Invoke(sender, results);
        }

        # endregion

        # region public method

        public void RunWorkerSync(IntPtr callback, object argument)
        {            
            var worker = (DoWork)Marshal.GetDelegateForFunctionPointer(callback, typeof(DoWork));
            var r = new RunWorkerCompletedEventArgsWrapper();

            worker?.Invoke(ref argument, ref r);
            RunWorkerCompleted?.Invoke(this, r);
        }

        public void RunWorkerAsync(IntPtr callback, object argument)
        {
            this.callback = callback;
            BackgroundWorker.RunWorkerAsync(argument);
        }

        public void ReportProgress(int percentProgress, object userState)
        {
            BackgroundWorker.ReportProgress(percentProgress, userState);
        }

        public void CancelAsync()
        {
            BackgroundWorker.CancelAsync();
        }

        # endregion                
    }
}
