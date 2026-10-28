using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ClosedXML.XmlMaps.Tests.Infrastructure
{
    /// <summary>
    /// A hidden Excel instance driven through late-bound COM. No interop assembly is needed.
    /// </summary>
    internal sealed class ExcelSession : IDisposable
    {
        private const int XlOpenXmlWorkbook = 51;
        private const int MsoAutomationSecurityForceDisable = 3;

        private readonly int _processId;
        private readonly ComMessageFilter? _messageFilter;
        private dynamic? _app;

        private ExcelSession(dynamic app)
        {
            _messageFilter = System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA ? ComMessageFilter.Register() : null;
            _app = app;
            app.DisplayAlerts = false;
            app.Visible = false;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.AutomationSecurity = MsoAutomationSecurityForceDisable;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out _processId);
        }

        internal static bool IsAvailable => Type.GetTypeFromProgID("Excel.Application") is not null;

        internal dynamic App => _app ?? throw new ObjectDisposedException(nameof(ExcelSession));

        internal static ExcelSession Start()
        {
            var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new InvalidOperationException("Excel is not installed.");
            return new ExcelSession(Activator.CreateInstance(type)!);
        }

        /// <summary>
        /// Open a workbook. The caller must close it with <see cref="Close"/>.
        /// </summary>
        internal dynamic Open(string path, bool readOnly = false)
        {
            return App.Workbooks.Open(path, 0, readOnly);
        }

        internal static void SaveAs(dynamic workbook, string path)
        {
            workbook.SaveAs(path, XlOpenXmlWorkbook);
        }

        internal static void Close(dynamic workbook)
        {
            try
            {
                workbook.Close(false);
            }
            finally
            {
                Marshal.FinalReleaseComObject(workbook);
            }
        }

        public void Dispose()
        {
            if (_app is null)
                return;

            try
            {
                foreach (var workbook in _app.Workbooks)
                    workbook.Close(false);

                _app.Quit();
            }
            catch (COMException)
            {
                // Excel is gone already, the process is killed below.
            }
            finally
            {
                Marshal.FinalReleaseComObject(_app);
                _app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            try
            {
                using var process = Process.GetProcessById(_processId);
                if (!process.WaitForExit(5000))
                    process.Kill();
            }
            catch (ArgumentException)
            {
                // The process has exited.
            }

            _messageFilter?.Dispose();
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }
}
