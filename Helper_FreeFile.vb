Imports System.Runtime.InteropServices
Imports System.Threading
'Imports srm = System.Reflection.MethodBase
Public Class NativeMethods

End Class
Public Class Helper_FreeFile
    Public Shared Sub FreeFileByHandle(ByVal Filename As String)
        Dim LockedFileProcesses As List(Of Process) = WhoIsLocking(Filename)
        If LockedFileProcesses.Count > 0 Then
            For Each proc As Process In LockedFileProcesses
                'a = proc.MainWindowTitle
                Dim pid As Integer = proc.Id
                Dim hInfo As New List(Of HandleInfo)
                hInfo = GetAllHandles(pid)
                For Each h As HandleInfo In hInfo
                    CloseHandle(CType(h.Handle, IntPtr))
                Next
            Next
            Helper.wait(500)
        End If

    End Sub
    Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As Boolean
    <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function RmRegisterResources(pSessionHandle As UInteger, nFiles As UInt32, rgsFilenames As String(), nApplications As UInt32, <[In]> rgApplications As RM_UNIQUE_PROCESS(), nServices As UInt32,
        rgsServiceNames As String()) As Integer
    End Function

    <DllImport("rstrtmgr.dll", CharSet:=CharSet.Auto)>
    Private Shared Function RmStartSession(ByRef pSessionHandle As UInteger, dwSessionFlags As Integer, strSessionKey As String) As Integer
    End Function

    <DllImport("rstrtmgr.dll")>
    Private Shared Function RmEndSession(pSessionHandle As UInteger) As Integer
    End Function

    <DllImport("rstrtmgr.dll")>
    Private Shared Function RmGetList(dwSessionHandle As UInteger, ByRef pnProcInfoNeeded As UInteger, ByRef pnProcInfo As UInteger, <[In], Out> rgAffectedApps As RM_PROCESS_INFO(), ByRef lpdwRebootReasons As UInteger) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RM_UNIQUE_PROCESS
        Public dwProcessId As Integer
        Public ProcessStartTime As System.Runtime.InteropServices.ComTypes.FILETIME
    End Structure

    Const RmRebootReasonNone As Integer = 0
    Const CCH_RM_MAX_APP_NAME As Integer = 255
    Const CCH_RM_MAX_SVC_NAME As Integer = 63

    Private Enum RM_APP_TYPE
        RmUnknownApp = 0
        RmMainWindow = 1
        RmOtherWindow = 2
        RmService = 3
        RmExplorer = 4
        RmConsole = 5
        RmCritical = 1000
    End Enum

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure RM_PROCESS_INFO
        Public Process As RM_UNIQUE_PROCESS

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_APP_NAME + 1)>
        Public strAppName As String

        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_SVC_NAME + 1)>
        Public strServiceShortName As String

        Public ApplicationType As RM_APP_TYPE
        Public AppStatus As UInteger
        Public TSSessionId As UInteger
        <MarshalAs(UnmanagedType.Bool)>
        Public bRestartable As Boolean
    End Structure
    Private Shared Function WhoIsLocking(path As String) As List(Of Process)
        Dim handle As UInteger
        Dim key As String = Guid.NewGuid().ToString()
        Dim processes As New List(Of Process)()

        Dim res As Integer = RmStartSession(handle, 0, key)
        If res <> 0 Then
            Throw New Exception("Could not begin restart session.  Unable to determine file locker.")
        End If

        Try
            Const ERROR_MORE_DATA As Integer = 234
            Dim pnProcInfoNeeded As UInteger = 0, pnProcInfo As UInteger = 0, lpdwRebootReasons As UInteger = RmRebootReasonNone

            Dim resources As String() = New String() {path}
            ' Just checking on one resource.
            res = RmRegisterResources(handle, CUInt(resources.Length), resources, 0, Nothing, 0, Nothing)

            If res <> 0 Then
                Throw New Exception("Could not register resource.")
            End If

            'Note: there's a race condition here -- the first call to RmGetList() returns
            '      the total number of process. However, when we call RmGetList() again to get
            '      the actual processes this number may have increased.
            res = RmGetList(handle, pnProcInfoNeeded, pnProcInfo, Nothing, lpdwRebootReasons)

            If res = ERROR_MORE_DATA Then
                ' Create an array to store the process results
                Dim processInfo As RM_PROCESS_INFO() = New RM_PROCESS_INFO(Helper_VarConvert.ConvertToInteger(pnProcInfoNeeded - 1, 0)) {}
                pnProcInfo = pnProcInfoNeeded

                ' Get the list
                res = RmGetList(handle, pnProcInfoNeeded, pnProcInfo, processInfo, lpdwRebootReasons)
                If res = 0 Then
                    processes = New List(Of Process)(Helper_VarConvert.ConvertToInteger(pnProcInfo, 0))

                    ' Enumerate all of the results and add them to the 
                    ' list to be returned
                    For i As Integer = 0 To Helper_VarConvert.ConvertToInteger(pnProcInfo - 1, 0) 
                        Try
                            processes.Add(Process.GetProcessById(processInfo(i).Process.dwProcessId))
                            ' catch the error -- in case the process is no longer running
                        Catch generatedExceptionName As ArgumentException
                        End Try
                    Next
                Else
                    Throw New Exception("Could not list processes locking resource.")
                End If
            ElseIf res <> 0 Then
                Throw New Exception("Could not list processes locking resource. Failed to get size of result.")
            End If
        Finally
            RmEndSession(handle)
        End Try

        Return processes
    End Function
    '----------------------------------
    Public Enum SYSTEM_INFORMATION_CLASS
        SystemBasicInformation
        SystemProcessorInformation
        SystemPerformanceInformation
        SystemTimeOfDayInformation
        SystemPathInformation
        SystemProcessInformation
        SystemCallCountInformation
        SystemDeviceInformation
        SystemProcessorPerformanceInformation
        SystemFlagsInformation
        SystemCallTimeInformation
        SystemModuleInformation
        SystemLocksInformation
        SystemStackTraceInformation
        SystemPagedPoolInformation
        SystemNonPagedPoolInformation
        SystemHandleInformation
        SystemObjectInformation
        SystemPageFileInformation
        SystemVdmInstemulInformation
        SystemVdmBopInformation
        SystemFileCacheInformation
        SystemPoolTagInformation
        SystemInterruptInformation
        SystemDpcBehaviorInformation
        SystemFullMemoryInformation
        SystemLoadGdiDriverInformation
        SystemUnloadGdiDriverInformation
        SystemTimeAdjustmentInformation
        SystemSummaryMemoryInformation
        SystemMirrorMemoryInformation
        SystemPerformanceTraceInformation
        SystemObsolete0
        SystemExceptionInformation
        SystemCrashDumpStateInformation
        SystemKernelDebuggerInformation
        SystemContextSwitchInformation
        SystemRegistryQuotaInformation
        SystemExtendServiceTableInformation
        SystemPrioritySeperation
        SystemVerifierAddDriverInformation
        SystemVerifierRemoveDriverInformation
        SystemProcessorIdleInformation
        SystemLegacyDriverInformation
        SystemCurrentTimeZoneInformation
        SystemLookasideInformation
        SystemTimeSlipNotification
        SystemSessionCreate
        SystemSessionDetach
        SystemSessionInformation
        SystemRangeStartInformation
        SystemVerifierInformation
        SystemVerifierThunkExtend
        SystemSessionProcessInformation
        SystemLoadGdiDriverInSystemSpace
        SystemNumaProcessorMap
        SystemPrefetcherInformation
        SystemExtendedProcessInformation
        SystemRecommendedSharedDataAlignment
        SystemComPlusPackage
        SystemNumaAvailableMemory
        SystemProcessorPowerInformation
        SystemEmulationBasicInformation
        SystemEmulationProcessorInformation
        SystemExtendedHandleInformation
        SystemLostDelayedWriteInformation
        SystemBigPoolInformation
        SystemSessionPoolTagInformation
        SystemSessionMappedViewInformation
        SystemHotpatchInformation
        SystemObjectSecurityMode
        SystemWatchdogTimerHandler
        SystemWatchdogTimerInformation
        SystemLogicalProcessorInformation
        SystemWow64SharedInformationObsolete
        SystemRegisterFirmwareTableInformationHandler
        SystemFirmwareTableInformation
        SystemModuleInformationEx
        SystemVerifierTriageInformation
        SystemSuperfetchInformation
        SystemMemoryListInformation
        SystemFileCacheInformationEx
        SystemThreadPriorityClientIdInformation
        SystemProcessorIdleCycleTimeInformation
        SystemVerifierCancellationInformation
        SystemProcessorPowerInformationEx
        SystemRefTraceInformation
        SystemSpecialPoolInformation
        SystemProcessIdInformation
        SystemErrorPortInformation
        SystemBootEnvironmentInformation
        SystemHypervisorInformation
        SystemVerifierInformationEx
        SystemTimeZoneInformation
        SystemImageFileExecutionOptionsInformation
        SystemCoverageInformation
        SystemPrefetchPatchInformation
        SystemVerifierFaultsInformation
        SystemSystemPartitionInformation
        SystemSystemDiskInformation
        SystemProcessorPerformanceDistribution
        SystemNumaProximityNodeInformation
        SystemDynamicTimeZoneInformation
        SystemCodeIntegrityInformation
        SystemProcessorMicrocodeUpdateInformation
        MaxSystemInfoClass
    End Enum


    <StructLayout(LayoutKind.Explicit)>
    Public Structure LARGE_INTEGER
        <FieldOffset(0)>
        Public LowPart As Integer

        <FieldOffset(4)>
        Public HighPart As Integer

        <FieldOffset(0)>
        Public QuadPart As Long
    End Structure


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure UNICODE_STRING
        Public Length As UShort
        Public MaximumLength As UShort
        <MarshalAs(UnmanagedType.LPWStr)>
        Public Buffer As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure SYSTEM_PROCESS_INFORMATION
        Public NextEntryOffset As UInteger
        Public NumberOfThreads As UInteger
        Public WorkingSetPrivateSize As LARGE_INTEGER
        Public SpareLi2 As LARGE_INTEGER
        Public SpareLi3 As LARGE_INTEGER
        Public CreateTime As LARGE_INTEGER
        Public UserTime As LARGE_INTEGER
        Public KernelTime As LARGE_INTEGER
        Public ImageName As UNICODE_STRING
        Public BasePriority As Integer
        Public UniqueProcessId As Integer
        Public InheritedFromUniqueProcessId As Integer
        Public HandleCount As UInteger
        Public SessionId As UInteger
        Public UniqueProcessKey As UInteger
        Public PeakVirtualSize As UInteger
        Public VirtualSize As UInteger
        Public PageFaultCount As UInteger
        Public PeakWorkingSetSize As UInteger
        Public WorkingSetSize As UInteger
        Public QuotaPeakPagedPoolUsage As UInteger
        Public QuotaPagedPoolUsage As UInteger
        Public QuotaPeakNonPagedPoolUsage As UInteger
        Public QuotaNonPagedPoolUsage As UInteger
        Public PagefileUsage As UInteger
        Public PeakPagefileUsage As UInteger
        Public PrivatePageCount As UInteger
        Public ReadOperationCount As LARGE_INTEGER
        Public WriteOperationCount As LARGE_INTEGER
        Public OtherOperationCount As LARGE_INTEGER
        Public ReadTransferCount As LARGE_INTEGER
        Public WriteTransferCount As LARGE_INTEGER
        Public OtherTransferCount As LARGE_INTEGER
    End Structure

    '<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    'Public Structure SYSTEM_HANDLE_INFORMATION
    '    Public ProcessId As UInteger
    '    Public ObjectTypeNumber As Byte
    '    Public Flags As Byte
    '    Public Handle As UShort
    '    Public pObject As IntPtr
    '    Public GrantedAccess As Integer
    'End Structure

    <DllImport("NtDll", SetLastError:=True, CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Auto)>
    Public Shared Function NtQuerySystemInformation(SystemInformationClass As SYSTEM_INFORMATION_CLASS, SystemInformation As IntPtr, SystemInformationLength As Integer, ByRef ReturnLength As Integer) As Integer
    End Function
    <StructLayout(LayoutKind.Sequential)> Structure SYSTEM_HANDLE_INFORMATION
        Public ProcessID As Integer
        Public ObjectTypeNumber As Byte
        Public Flags As Byte
        Public Handle As UShort
        Public Object_Pointer As Integer
        Public GrantedAccess As Integer


    End Structure
    'Public Shared Widening Operator CType(v As SYSTEM_HANDLE_INFORMATION) As SYSTEM_HANDLE_INFORMATION
    '    Throw New NotImplementedException()
    'End Operator
    Public Structure HandleInfo
        Public ProcessID As Integer
        Public ObjectTypeNumber As Byte
        Public Flags As Byte
        Public Handle As UShort
        Public Object_Pointer As Integer
        Public GrantedAccess As Integer
    End Structure
    Public Shared Function GetAllHandles(Optional ByVal PID As Integer = 0) As List(Of HandleInfo)
        Dim nCurrentLength As Integer = 0
        Dim nHandleInfoSize As Integer = &H10000
        Dim lpHandle As IntPtr = IntPtr.Zero
        Dim lpBufferHandles As IntPtr = Marshal.AllocHGlobal(nHandleInfoSize)

        Try
            While NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemHandleInformation, lpBufferHandles, nHandleInfoSize, nCurrentLength) > 0
                'STATUS_INFO_LENGTH_MISMATCH
                nHandleInfoSize = nCurrentLength
                Marshal.FreeHGlobal(lpBufferHandles)
                lpBufferHandles = Marshal.AllocHGlobal(nCurrentLength)
            End While

            Dim lHandleCount As Long = 0
            Dim architecture As String = ""
            Dim ri As integer = Helper_VarConvert.ConvertToInteger(IntPtr.Size, 0) 
            If ri > 4 Then
                architecture = "64 Bit"
            Else
                architecture = "32 Bit"
            End If
            If architecture = "64 Bit" Then
                lHandleCount = Marshal.ReadInt64(lpBufferHandles)
                lpHandle = New IntPtr(lpBufferHandles.ToInt64() + 8)
            Else
                lHandleCount = Marshal.ReadInt32(lpBufferHandles)
                lpHandle = New IntPtr(lpBufferHandles.ToInt32() + 4)
            End If

            Dim shHandle As SYSTEM_HANDLE_INFORMATION
            Dim HandlesInfos As New List(Of HandleInfo)

            For lIndex As Long = 0 To lHandleCount - 1
                shHandle = New SYSTEM_HANDLE_INFORMATION()
                If architecture = "64 Bit" Then
                    lpHandle = New IntPtr(lpHandle.ToInt64() + Marshal.SizeOf(shHandle) + 8)
                    shHandle = DirectCast(Marshal.PtrToStructure(lpHandle, shHandle.[GetType]()), SYSTEM_HANDLE_INFORMATION)
                Else
                    lpHandle = New IntPtr(lpHandle.ToInt32() + Marshal.SizeOf(shHandle))
                    shHandle = DirectCast(Marshal.PtrToStructure(lpHandle, shHandle.[GetType]()), SYSTEM_HANDLE_INFORMATION)
                End If
                If PID = 0 OrElse PID = shHandle.ProcessID Then
                    Dim hi As New HandleInfo
                    hi.ProcessID = shHandle.ProcessID
                    hi.ObjectTypeNumber = shHandle.ObjectTypeNumber
                    hi.Flags = shHandle.Flags
                    hi.Handle = shHandle.Handle
                    hi.Object_Pointer = shHandle.Object_Pointer
                    hi.GrantedAccess = shHandle.GrantedAccess
                    HandlesInfos.Add(hi)
                End If
            Next

            ' CloseProcessForHandle()
            Marshal.FreeHGlobal(lpBufferHandles)
            Return HandlesInfos

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return Nothing
        End Try
    End Function
End Class
