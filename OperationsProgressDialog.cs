using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ANDAI_Launcher
{
    public class OperationsProgressDialog : Component
    {
        const string CLSID_ProgressDialog = "{F8383852-FCD3-11d1-A6B9-006097DF5BD4}";
        const string IDD_ProgressDialog = "0C9FB851-E5C9-43EB-A370-F0677B13874C";
        const string CLSID_IShellItem = "{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}";
        private Type _progressDialogType;
        private Type _IShellItemType;
        private IOperationsProgressDialog _nativeProgressDialog;
        private PROGDLG dialogFlags;
        private SPACTION operationFlags;
        private PDMODE modeFlags;
        private int currentProgress;
        private int totalProgress;
        private int currentSize;
        private int totalSize;
        private int currentItems;
        private int totalItems;
        private bool estimateValue;
        private IShellItem sourceItem;
        private IShellItem destItem;
        private IShellItem currentItem;
        private DIALOGSTATUS dialogStatus;

        public enum DIALOGSTATUS
        {
            DLG_NOTSTARTED = 0,
            DLG_RUNNING = 1,
            DLG_DISPOSED = 2,
            DLG_ERRORED = 3,
        }

        public DIALOGSTATUS IsDialogActive
        {
            get { return dialogStatus; }
        }

        public PROGDLG DialogFlags
        {
            set { dialogFlags = value; }
            get { return dialogFlags; }
        }

        public SPACTION OperationFlags
        {
            set { operationFlags = value; }
            get { return operationFlags; }
        }

        public PDMODE ModeFlags
        {
            set { modeFlags = value; }
            get { return modeFlags; }
        }

        public int ProgressBarValue
        {
            get { return currentProgress; }
            set { currentProgress = value; UpdateProgress(); }
        }

        public int ProgressBarMaxValue
        {
            get { return totalProgress; }
            set { totalProgress = value; UpdateProgress(); }
        }

        public int ProgressDialogSizeValue
        {
            get { return currentSize; }
            set { currentSize = value; UpdateProgress(); }
        }

        public int ProgressDialogSizeMaxValue
        {
            get { return totalSize; }
            set { totalSize = value; UpdateProgress(); }
        }

        public int ProgressDialogItemsValue
        {
            get { return currentItems; }
            set { currentItems = value; UpdateProgress(); }
        }

        public int ProgressDialogItemsMaxValue
        {
            get { return totalItems; }
            set { totalItems = value; UpdateProgress(); }
        }

        /// <summary>
        /// Checks if the dialog has reached 100%.
        /// </summary>
        public bool isFinished
        {
            get { return currentProgress == totalProgress; }
        }

        /// <summary>
        /// Returns a boolean of if the user has cancelled or not.
        /// </summary>
        public bool hasUserCancelled
        {
            get { if (_nativeProgressDialog != null) { return _nativeProgressDialog.GetOperationStatus() == PDOPSTATUS.PDOPS_CANCELLED; } else { return false; } }
        }

        /// <summary>
        /// Estimates the progress based on the item count.
        /// </summary>
        public bool EstimateValue
        {
            set { estimateValue = value; }
        }

        /// <summary>
        /// Closes the dialog.
        /// </summary>
        public async void Close()
        {
            if (_nativeProgressDialog != null && dialogStatus != DIALOGSTATUS.DLG_DISPOSED)
            {
                _nativeProgressDialog.StopProgressDialog();
                dialogStatus = DIALOGSTATUS.DLG_DISPOSED;
                await Task.Delay(900);
                Marshal.FinalReleaseComObject(_nativeProgressDialog);
                _nativeProgressDialog = null;
            }
        }

        private void UpdateProgress()
        {
            if (estimateValue)
            {
                currentProgress = currentItems;
                totalProgress = totalItems;
            }
            if (_nativeProgressDialog != null && dialogStatus != DIALOGSTATUS.DLG_DISPOSED)
            {
                _nativeProgressDialog.UpdateProgress((uint)currentProgress, (uint)totalProgress, (uint)currentSize, (uint)totalSize, (uint)currentItems, (uint)totalItems);
                _nativeProgressDialog.UpdateLocations(sourceItem, destItem, currentItem);
            }
        }

        /// <summary>
        /// Initializes the dialog.
        /// </summary>
        public OperationsProgressDialog()
        {
            dialogStatus = DIALOGSTATUS.DLG_NOTSTARTED;
            _progressDialogType = Type.GetTypeFromCLSID(new Guid(CLSID_ProgressDialog));
            _IShellItemType = Type.GetTypeFromCLSID(new Guid(CLSID_IShellItem));
            dialogFlags = PROGDLG.PROGDLG_NORMAL;
            operationFlags = SPACTION.SPACTION_NONE;
            modeFlags = PDMODE.PDM_DEFAULT;
            currentProgress = 0;
            totalProgress = 100;
            currentSize = 0;
            totalSize = 100;
            currentItems = 0;
            totalItems = 100;
            estimateValue = false;
            SHCreateItemFromParsingName("https://andai.heliohost.org/packages.php", IntPtr.Zero, typeof(IShellItem).GUID, out sourceItem);
            SHCreateItemFromParsingName(Environment.CurrentDirectory, IntPtr.Zero, typeof(IShellItem).GUID, out destItem);
            SHCreateItemFromParsingName(Environment.CurrentDirectory, IntPtr.Zero, typeof(IShellItem).GUID, out currentItem);
        }

        public OperationsProgressDialog(IContainer container)
			: this() {
            container.Add(this);
        }

        /// <summary>
        /// Shows the dialog with the specified parent. If the parent is null, uses the active form.
        /// </summary>
        /// <param name="parent"></param>
        public void Show(IWin32Window parent)
        {
            if (parent == null) parent = Form.ActiveForm;
            IntPtr handle = (parent == null) ? IntPtr.Zero : parent.Handle;
            _nativeProgressDialog = (IOperationsProgressDialog)Activator.CreateInstance(_progressDialogType);
            _nativeProgressDialog.StartProgressDialog(handle, dialogFlags);
            _nativeProgressDialog.SetOperation(operationFlags);
            _nativeProgressDialog.SetMode(modeFlags);
            UpdateProgress();
            dialogStatus = DIALOGSTATUS.DLG_RUNNING;
        }

        [ComImport, Guid(IDD_ProgressDialog), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOperationsProgressDialog
        {
            void StartProgressDialog(IntPtr hwndOwner, PROGDLG flags);
            void StopProgressDialog();
            void SetOperation(SPACTION action);
            void SetMode(PDMODE mode);
            void UpdateProgress(uint ullPointsCurrent, uint ullPointsTotal, uint ullSizeCurrent, uint ullSizeTotal, uint ullItemsCurrent, uint ullItemsTotal);
            void UpdateLocations(IShellItem psiSource, IShellItem psiTarget, IShellItem psiItem);
            void ResetTimer();
            void PauseTimer();
            void ResumeTimer();
            void GetMilliseconds(ulong pullElapsed, ulong pullRemaining);
            PDOPSTATUS GetOperationStatus();
        }

        /// <summary>
        /// Shows the dialog (if no settings were modified, uses default).
        /// </summary>
        internal void Show()
        {
            Show(null);
        }

        /// <summary>
        /// Specifies what the dialog is currently doing.
        /// </summary>
        /// <param name="operation"></param>
        internal void SetOperation(SPACTION operation)
        {
            if(_nativeProgressDialog != null && dialogStatus != DIALOGSTATUS.DLG_DISPOSED)
            {
                _nativeProgressDialog.SetOperation(operation);
            }
        }

        /// <summary>
        /// Specifies the status of the current operation.
        /// </summary>
        /// <param name="mode"></param>
        internal void SetMode(PDMODE mode)
        {
            if (_nativeProgressDialog != null && dialogStatus!=DIALOGSTATUS.DLG_DISPOSED)
            {
                _nativeProgressDialog.SetMode(mode);
            }
        }

        /// <summary>
        /// Specifies the items being used.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        /// <param name="item"></param>
        internal void UpdateLocations(string source, string destination, string item)
        {
            try
            {
                SHCreateItemFromParsingName(source, IntPtr.Zero, typeof(IShellItem).GUID, out sourceItem);
                SHCreateItemFromParsingName(destination, IntPtr.Zero, typeof(IShellItem).GUID, out destItem);
                SHCreateItemFromParsingName(item, IntPtr.Zero, typeof(IShellItem).GUID, out currentItem);
            }
            catch (FileNotFoundException ex)
            {
                throw new IOperationsProgressDialogException("One or more of the locations do not exist.", ex);
            }
            UpdateProgress();
        }

        internal void UpdateLocations(string source, string destination)
        {
            try
            {
                System.Diagnostics.Debug.Print(Environment.CurrentDirectory + " " + source + " " + destination);
                SHCreateItemFromParsingName(source, IntPtr.Zero, typeof(IShellItem).GUID, out sourceItem);
                SHCreateItemFromParsingName(destination, IntPtr.Zero, typeof(IShellItem).GUID, out destItem);
            }
            catch (FileNotFoundException ex)
            {
                throw new IOperationsProgressDialogException("One or more of the locations do not exist.", ex);
            }
            UpdateProgress();
        }

        internal void UpdateLocations(string item)
        {
            try
            {
                SHCreateItemFromParsingName(item.Replace("/", "\\"), IntPtr.Zero, typeof(IShellItem).GUID, out currentItem);
            }
            catch (FileNotFoundException ex)
            {
                try
                {
                    SHCreateItemFromParsingName(Environment.CurrentDirectory + "\\" + item.Replace("/", "\\"), IntPtr.Zero, typeof(IShellItem).GUID, out currentItem);
                }
                catch (FileNotFoundException)
                {
                    throw new IOperationsProgressDialogException("One or more of the locations do not exist.", ex);
                }                
            }
            UpdateProgress();
        }

        public enum PROGDLG : uint
        {
            /// <summary>
            /// Default, normal progress dialog behavior.
            /// </summary>
            PROGDLG_NORMAL = 0x00000000,
            /// <summary>
            /// The dialog is modal to its hwndOwner. The default setting is modeless.
            /// </summary>
            PROGDLG_MODAL = 0x00000001,
            /// <summary>
            /// Update "Line3" text with the time remaining. This flag does not need to be implicitly set because progress dialogs started by IOperationsProgressDialog::StartProgressDialog automatically display the time remaining.
            /// </summary>
            PROGDLG_AUTOTIME = 0x00000002,
            /// <summary>
            /// Do not show the time remaining. We do not recommend setting this flag through IOperationsProgressDialog because it goes against the purpose of the dialog.
            /// </summary>
            PROGDLG_NOTIME = 0x00000004,
            /// <summary>
            /// Do not display the minimize button.
            /// </summary>
            PROGDLG_NOMINIMIZE = 0x00000008,
            /// <summary>
            /// Do not display the progress bar.
            /// </summary>
            PROGDLG_NOPROGRESSBAR = 0x00000010,
            /// <summary>
            /// This flag is invalid in this method. To set the progress bar to marquee mode, use the flags in IOperationsProgressDialog::SetMode.
            /// </summary>
            PROGDLG_MARQUEEPROGRESS = 0x00000020,
            /// <summary>
            /// Do not display a cancel button because the operation cannot be canceled. Use this value only when absolutely necessary.
            /// </summary>
            PROGDLG_NOCANCEL = 0x00000040,
            /// <summary>
            /// Windows 7 and later. Indicates default, normal operation progress dialog behavior.
            /// </summary>
            OPPROGDLG_DEFAULT = 0x00000000,
            /// <summary>
            /// Display a pause button. Use this only in situations where the operation can be paused.
            /// </summary>
            OPPROGDLG_ENABLEPAUSE = 0x00000080,
            /// <summary>
            /// The operation can be undone through the dialog. The Stop button becomes Undo. If pressed, the Undo button then reverts to Stop.
            /// </summary>
            OPPROGDLG_ALLOWUNDO = 0x00000100,
            /// <summary>
            /// Do not display the path of source file in the progress dialog.
            /// </summary>
            OPPROGDLG_DONTDISPLAYSOURCEPATH = 0x00000200,
            /// <summary>
            /// Do not display the path of the destination file in the progress dialog.
            /// </summary>
            OPPROGDLG_DONTDISPLAYDESTPATH = 0x00000400,
            /// <summary>
            /// Windows 7 and later. If the estimated time to completion is greater than one day, do not display the time.
            /// </summary>
            OPPROGDLG_NOMULTIDAYESTIMATES = 0x00000800,
            /// <summary>
            /// Windows 7 and later. Do not display the location line in the progress dialog.
            /// </summary>
            OPPROGDLG_DONTDISPLAYLOCATIONS = 0x00001000
        }

        public enum PDMODE : uint
        {
            /// <summary>
            /// Use the default progress dialog operations mode.
            /// </summary>
            PDM_DEFAULT = 0x00000000,
            /// <summary>
            /// The operation is running.
            /// </summary>
            PDM_RUN = 0x00000001,
            /// <summary>
            /// The operation is gathering data before it begins, such as calculating the predicted operation time.
            /// </summary>
            PDM_PREFLIGHT = 0x00000002,
            /// <summary>
            /// The operation is rolling back due to an Undo command from the user.
            /// </summary>
            PDM_UNDOING = 0x00000004,
            /// <summary>
            /// Error dialogs are blocking progress from continuing.
            /// </summary>
            /// <remarks>
            /// Appears to only work when progress has already begun.
            /// </remarks>
            PDM_ERRORSBLOCKING = 0x00000008,
            /// <summary>
            /// The length of the operation is indeterminate. Do not show a timer and display the progress bar in marquee mode.
            /// </summary>
            PDM_INDETERMINATE = 0x00000010
        }

        public enum SPACTION : uint
        {
            /// <summary>
            /// No action is being performed.
            /// </summary>
            SPACTION_NONE = 0,
            /// <summary>
            /// Files are being moved.
            /// </summary>
            SPACTION_MOVING,
            /// <summary>
            /// Files are being copied.
            /// </summary>
            SPACTION_COPYING,
            /// <summary>
            /// Files are being deleted.
            /// </summary>
            SPACTION_RECYCLING,
            /// <summary>
            /// A set of attributes are being applied to files.
            /// </summary>
            SPACTION_APPLYINGATTRIBS,
            /// <summary>
            /// A file is being downloaded from a remote source.
            /// </summary>
            SPACTION_DOWNLOADING,
            /// <summary>
            /// An Internet search is being performed.
            /// </summary>
            SPACTION_SEARCHING_INTERNET,
            /// <summary>
            /// A calculation is being performed.
            /// </summary>
            SPACTION_CALCULATING,
            /// <summary>
            /// A file is being uploaded to a remote source.
            /// </summary>
            SPACTION_UPLOADING,
            /// <summary>
            /// A local search is being performed.
            /// </summary>
            SPACTION_SEARCHING_FILES,
            /// <summary>
            /// Windows Vista and later. A deletion is being performed.
            /// </summary>
            SPACTION_DELETING,
            /// <summary>
            /// Windows Vista and later. A renaming action is being performed.
            /// </summary>
            SPACTION_RENAMING,
            /// <summary>
            /// Windows Vista and later. A formatting action is being performed.
            /// </summary>
            SPACTION_FORMATTING,
            /// <summary>
            /// Windows 7 and later. A copy or move action is being performed.
            /// </summary>
            SPACTION_COPY_MOVING
        }

        public enum PDOPSTATUS : uint
        {
            /// <summary>
            /// Operation is running, no user intervention.
            /// </summary>
            PDOPS_RUNNING = 1,
            /// <summary>
            /// Operation has been paused by the user.
            /// </summary>
            PDOPS_PAUSED = 2,
            /// <summary>
            /// Operation has been canceled by the user - now go undo.
            /// </summary>
            PDOPS_CANCELLED = 3,
            /// <summary>
            /// Operation has been stopped by the user - terminate completely.
            /// </summary>
            PDOPS_STOPPED = 4,
            /// <summary>
            /// Operation has gone as far as it can go without throwing error dialogs.
            /// </summary>
            PDOPS_ERRORS = 5
        }
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        public interface IShellItem
        {
            void BindToHandler(IntPtr pbc,
                [MarshalAs(UnmanagedType.LPStruct)]Guid bhid,
                [MarshalAs(UnmanagedType.LPStruct)]Guid riid,
                out IntPtr ppv);

            void GetParent(out IShellItem ppsi);

            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);

            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);

            void Compare(IShellItem psi, uint hint, out int piOrder);
        };

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void SHCreateItemFromParsingName([In][MarshalAs(UnmanagedType.LPWStr)] string pszPath, [In] IntPtr pbc, [In][MarshalAs(UnmanagedType.LPStruct)]Guid riid, [Out][MarshalAs(UnmanagedType.Interface, IidParameterIndex = 2)] out IShellItem ppv);

        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0,
            PARENTRELATIVEPARSING = 0x80018001,
            PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
            DESKTOPABSOLUTEPARSING = 0x80028000,
            PARENTRELATIVEEDITING = 0x80031001,
            DESKTOPABSOLUTEEDITING = 0x8004c000,
            FILESYSPATH = 0x80058000,
            URL = 0x80068000
        }
    }

    [Serializable]
    internal class IOperationsProgressDialogException : Exception
    {
        public IOperationsProgressDialogException()
        {
        }

        public IOperationsProgressDialogException(string message) : base(message)
        {
        }

        public IOperationsProgressDialogException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected IOperationsProgressDialogException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}