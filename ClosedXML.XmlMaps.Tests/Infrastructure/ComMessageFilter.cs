using System;
using System.Runtime.InteropServices;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// Retries COM calls Excel rejects because it is busy (<c>RPC_E_SERVERCALL_RETRYLATER</c>), e.g. while it calculates.
    /// Works only on an STA thread.
    /// </summary>
    internal sealed class ComMessageFilter : ComMessageFilter.IOleMessageFilter, IDisposable
    {
        private const int ServerCallRetryLater = 2;
        private const int RetryDelay = 100;
        private const int RetryTimeout = 60_000;
        private const int PendingMessageWaitDefProcess = 2;

        private readonly IOleMessageFilter? _previous;

        private ComMessageFilter()
        {
            CoRegisterMessageFilter(this, out _previous);
        }

        internal static ComMessageFilter Register() => new();

        public void Dispose()
        {
            CoRegisterMessageFilter(_previous, out _);
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // A value between 0 and 99 means: retry immediately, 100 and more: retry after this number of milliseconds, -1: cancel.
            return dwRejectType == ServerCallRetryLater && dwTickCount < RetryTimeout ? RetryDelay : -1;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => PendingMessageWaitDefProcess;

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

        [ComImport]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }
    }
}
