<#-------------------------------------------------------------------------------

    Arcane :: Server

    .Developer
        Jean-Pierre LESUEUR (@DarkCoderSc)
        https://www.twitter.com/darkcodersc
        https://www.github.com/PhrozenIO
        https://github.com/DarkCoderSc
        www.phrozen.io
        jplesueur@phrozen.io
        PHROZEN

    .License
        Apache License
        Version 2.0, January 2004
        http://www.apache.org/licenses/

    .Disclaimer
        This script is provided "as is", without warranty of any kind, express or
        implied, including but not limited to the warranties of merchantability,
        fitness for a particular purpose and noninfringement. In no event shall the
        authors or copyright holders be liable for any claim, damages or other
        liability, whether in an action of contract, tort or otherwise, arising
        from, out of or in connection with the software or the use or other dealings
        in the software.

    .Notice
        Writing the entire code in a single PowerShell script is wished,
        allowing it to function both as a module or a standalone script.

-------------------------------------------------------------------------------#>

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Global Variables                                                             #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

$global:ArcaneVersion = "1.0.5"
$global:ArcaneProtocolVersion = "5.0.2"

$global:HostSyncHash = [HashTable]::Synchronized(@{
    host = $host
    ClipboardText = (Get-Clipboard -Raw)
})

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Enums Definitions                                                            #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

enum ClipboardMode {
    Disabled = 1
    Receive = 2
    Send = 3
    Both = 4
}

enum ProtocolCommand {
    Success = 1
    Fail = 2
    RequestSession = 3
    AttachToSession = 4
    BadRequest = 5
    ResourceFound = 6
    ResourceNotFound = 7
}

enum WorkerKind {
    Desktop = 1
    Events = 2
}

enum LogKind {
    Information
    Warning
    Success
    Error
}

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Windows API Definitions                                                      #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

Add-Type -Assembly System.Windows.Forms

Add-Type @"
    using System;
    using System.Security;
    using System.Runtime.InteropServices;

    public static class User32
    {
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EmptyClipboard();

        [DllImport("User32.dll", SetLastError=false)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetProcessDPIAware();

        [DllImport("User32.dll", SetLastError=false)]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern uint LoadCursorA(int hInstance, int lpCursorName);

        [DllImport("User32.dll", SetLastError=false)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorInfo(IntPtr pci);

        [DllImport("user32.dll", SetLastError=false)]
        public static extern void mouse_event(int flags, int dx, int dy, int cButtons, int info);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern int GetSystemMetrics(int nIndex);

        [DllImport("User32.dll", SetLastError=false)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll", SetLastError=false)]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr OpenInputDesktop(
            uint dwFlags,
            bool fInherit,
            uint dwDesiredAccess
        );

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr OpenDesktop(
            string lpszDesktop,
            uint dwFlags,
            bool fInherit,
            uint dwDesiredAccess
        );

        [DllImport("user32.dll", SetLastError=true, CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetUserObjectInformation(
            IntPtr hObj,
            int nIndex,
            IntPtr pvInfo,
            uint nLength,
            ref uint lpnLengthNeeded
        );

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseDesktop(
            IntPtr hDesktop
        );

        [DllImport("user32.dll", SetLastError=true)]
        public static extern IntPtr GetThreadDesktop(uint dwThreadId);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetThreadDesktop(
            IntPtr hDesktop
        );

        [DllImport("user32.dll", SetLastError = true)]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    }

    public static class Kernel32
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("Kernel32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern uint SetThreadExecutionState(uint esFlags);

        [DllImport("kernel32.dll", SetLastError=false, EntryPoint="RtlMoveMemory"), SuppressUnmanagedCodeSecurity]
        public static extern void CopyMemory(
            IntPtr dest,
            IntPtr src,
            IntPtr count
        );

        [DllImport("kernel32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll", SetLastError=true, CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", SetLastError=true, CharSet = CharSet.Ansi)]
        public static extern IntPtr GetProcAddress(
            IntPtr hModule,
            string procName
        );
    }

    public static class MSVCRT
    {
        [DllImport("msvcrt.dll", SetLastError=false, CallingConvention=CallingConvention.Cdecl), SuppressUnmanagedCodeSecurity]
        public static extern IntPtr memcmp(
            IntPtr p1,
            IntPtr p2,
            IntPtr count
        );
    }

    public static class GDI32
    {
        [DllImport("gdi32.dll")]
        public static extern IntPtr DeleteDC(IntPtr hDc);

        [DllImport("gdi32.dll")]
        public static extern IntPtr DeleteObject(IntPtr hDc);

        [DllImport("gdi32.dll", SetLastError=false), SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BitBlt(
            IntPtr hdcDest,
            int xDest,
            int yDest,
            int wDest,
            int hDest,
            IntPtr hdcSource,
            int xSrc,
            int ySrc,
            int RasterOp
        );

        [DllImport("gdi32.dll", SetLastError=false)]
        public static extern IntPtr CreateDIBSection(
            IntPtr hdc,
            IntPtr pbmi,
            uint usage,
            out IntPtr ppvBits,
            IntPtr hSection,
            uint offset
        );

        [DllImport ("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(
            IntPtr hdc,
            int nWidth,
            int nHeight
        );

        [DllImport ("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport ("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr bmp);

        [DllImport ("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
    }

    public static class Shcore {
        [DllImport("Shcore.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.U4)]
        public static extern uint SetProcessDpiAwareness(uint value);
    }
"@

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Script Blocks                                                                #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

$global:WinAPI_Const_ScriptBlock = {
    $GENERIC_ALL = 0x10000000

    $VK_LWIN = 0x5B;
    $KEYEVENTF_KEYDOWN = 0x0;
    $KEYEVENTF_KEYUP = 0x2;
}

# -------------------------------------------------------------------------------

$global:WinAPIException_Class_ScriptBlock = {
    class WinAPIException: System.Exception
    {
        WinAPIException([string] $ApiName) : base (
            [string]::Format(
                "WinApi Exception -> {0}, LastError: {1}",
                $ApiName,
                [System.Runtime.InteropServices.Marshal]::GetLastWin32Error().ToString()
            )
        )
        {}
    }
}

# -------------------------------------------------------------------------------

$global:GetUserObjectInformation_Func_ScriptBlock = {
    function Get-UserObjectInformationName
    {
        <#
            .SYNOPSIS
                Retrieves the name of the specified object.

            .PARAMETER hObj
                A handle to the object.
        #>
        param (
            [Parameter(Mandatory = $true)]
            [IntPtr]$hObj
        )

        $pvInfo = [IntPtr]::Zero
        try
        {
            $lpnLengthNeeded = [UInt32]0

            $UOI_NAME = 0x2
            $null = [User32]::GetUserObjectInformation(
                $hObj,
                $UOI_NAME,
                [IntPtr]::Zero,
                0,
                [ref]$lpnLengthNeeded
            )

            if ($lpnLengthNeeded -eq 0)
            {
                throw [WinAPIException]::New("GetUserObjectInformation(1)")
            }

            $pvInfo = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($lpnLengthNeeded)

            $b = [User32]::GetUserObjectInformation(
                $desktop,
                $UOI_NAME,
                $pvInfo,
                $lpnLengthNeeded,
                [ref]$lpnLengthNeeded
            )

            if ($b -eq $false)
            {
                throw [WinAPIException]::New("GetUserObjectInformation(2)")
            }

            return [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pvInfo)
        }
        finally
        {
            if ($pvInfo -ne [IntPtr]::Zero)
            {
                [System.Runtime.InteropServices.Marshal]::FreeHGlobal($pvInfo)
            }
        }

        return $objectName
    }
}

# -------------------------------------------------------------------------------

$global:GetInputDesktopName_Func_ScriptBlock = {
    function Get-InputDesktopName
    {
        <#
            .SYNOPSIS
                Retrieves the name of the input desktop (Desktop that receive input).
        #>
        $desktop = [IntPtr]::Zero
        try
        {
            $desktop = [User32]::OpenInputDesktop(0, $false, $GENERIC_ALL)
            if ($desktop -eq [IntPtr]::Zero)
            {
                throw [WinAPIException]::New("OpenInputDesktop")
            }

            return Get-UserObjectInformationName -hObj $desktop
        }
        finally
        {
            if ($desktop -ne [IntPtr]::Zero)
            {
                $null = [User32]::CloseDesktop($desktop)
            }
        }
    }
}

# -------------------------------------------------------------------------------

$global:GetCurrentThreadDesktopName_Func_ScriptBlock = {
    function Get-CurrentThreadDesktopName
    {
        <#
            .SYNOPSIS
                Retrieves the name of the desktop associated with the current thread.
        #>
        $desktop = [User32]::GetThreadDesktop([Kernel32]::GetCurrentThreadId())
        if ($desktop -eq [IntPtr]::Zero)
        {
            throw [WinAPIException]::New("GetThreadDesktop")
        }
        try
        {
            return Get-UserObjectInformationName -hObj $desktop
        }
        finally
        {
            $null = [User32]::CloseDesktop($desktop)
        }
    }
}

# -------------------------------------------------------------------------------

$global:UpdateCurrentThreadDesktop_Func_ScriptBlock = {
    function Update-CurrentThreadDesktop
    {
        <#
            .SYNOPSIS
                Updates the desktop associated with the current thread.

            .PARAMETER DesktopName
                The name of the desktop to be associated with the current thread.
        #>
        param(
            [Parameter(Mandatory=$True)]
            [string] $DesktopName
        )

        $desktop = [User32]::OpenDesktop($DesktopName, 0, $true, $GENERIC_ALL)
        if ($desktop -eq [IntPtr]::Zero)
        {
            throw [WinAPIException]::New("OpenDesktop")
        }
        try
        {
            if (-not [User32]::SetThreadDesktop($desktop))
            {
                throw [WinAPIException]::New("SetThreadDesktop")
            }
        }
        finally
        {
            $null = [User32]::CloseDesktop($desktop)
        }
    }
}

# -------------------------------------------------------------------------------

$global:UpdateCurrentThreadDesktopWithInputDesktop_Func_ScriptBlock = {
    function Update-CurrentThreadDesktopWidthInputDesktop()
    {
        <#
            .SYNOPSIS
                Updates the desktop associated with the current thread if input desktop changed.

            .DESCRIPTION
                Exceptions are catched and ignored.
        #>
        try
        {
            $currentThreadDesktopName = Get-CurrentThreadDesktopName
            $inputDesktopName = Get-InputDesktopName

            if ($currentThreadDesktopName -ne "" -and $inputDesktopName -ne "" -and $currentThreadDesktopName -ne $inputDesktopName)
            {
                $desktop = [User32]::OpenInputDesktop(0, $true,  $GENERIC_ALL)
                if ($desktop -eq [IntPtr]::Zero)
                {
                    throw [WinAPIException]::New("OpenInputDesktop")
                }
                try
                {
                    return [User32]::SetThreadDesktop($desktop)
                }
                finally {
                    $null = [User32]::CloseDesktop($desktop)
                }
            }
        }
        catch {}
        return $false
    }
}

# -------------------------------------------------------------------------------

$global:NewRunSpace_Func_ScriptBlock = {
    function New-RunSpace
    {
        <#
            .SYNOPSIS
                Create a new PowerShell Runspace.

            .DESCRIPTION
                Notice: the $host variable is used for debugging purpose to write on caller PowerShell
                Terminal.

            .PARAMETER ScriptBlocks
                Type: ScriptBlock[]
                Default: None
                Description: Instructions to execute in new runspace. Runspace can be composed of one or
                multiple script blocks.

            .PARAMETER Params
                Type: Hashtable
                Default: None
                Description: Hashtable containing parameters to pass to the runspace.

            .PARAMETER RunspaceApartmentState
                Type: String
                Default: "STA"
                Description: The apartment state of the runspace (Single, Multi).
        #>
        param(
            [Parameter(Mandatory=$True)]
            [ScriptBlock[]] $ScriptBlocks,

            [Hashtable]$Params = @{},

            [ValidateSet("STA", "MTA")]
            [string]$RunspaceApartmentState = "STA"
        )

        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.ThreadOptions = "UseNewThread"
        $runspace.ApartmentState = $RunspaceApartmentState
        $runspace.Open()

        if ($Params.Count -gt 0)
        {
            foreach ($key in $Params.Keys)
            {
                $runspace.SessionStateProxy.SetVariable(
                    $key,
                    $Params[$key]
                )
            }
        }

        $powershell = [PowerShell]::Create()

        foreach ($scriptBlock in $ScriptBlocks)
        {
            $null = $powershell.AddScript($scriptBlock)
        }

        $powershell.Runspace = $runspace

        $asyncResult = $powershell.BeginInvoke()

        return New-Object PSCustomObject -Property @{
            Runspace = $runspace
            PowerShell = $powershell
            AsyncResult = $asyncResult
        }
    }
}
. $NewRunSpace_Func_ScriptBlock

# -------------------------------------------------------------------------------

$global:DesktopStreamScriptBlock = {
    function Update-ScreensInformation
    {
        <#
            .SYNOPSIS
                Update screens information (force update).

            .DESCRIPTION
                (
                    PowerShell <= 5.1 -> Confirmed
                    PowerShell >= 6.0 <= 7.0 -> Untested
                    PowerShell >= 7.0 -> Seems to be fixed
                )
                It appears that a long-standing bug (PS < 7.0) affects the display resolution and screen count updates.
                Specifically, display information seems to be cached and is not refreshed until a new PowerShell session
                is started. This issue can be quite inconvenient. One potential solution would be to reimplement the
                display management logic using the Windows API directly.

                Given that the goal is to leverage PowerShell as much as possible, I've opted for a workaround that
                involves patching the internal state of the Screen class. By setting the screens field to null in memory,
                the next access to display information will force a refresh and provide updated screen data. This method,
                while somewhat "hacky", allows us to avoid extensive changes and keep the solution within the PowerShell
                environment.
        #>
        try
        {
            # PowerShell >= 7.0, issue seems to be fixed
            if ($PSVersionTable.PSVersion.Major -ge 7)
            {
                return
            }

            # Patch memory to force screen information to update (cache killer)
            ([System.Windows.Forms.Screen].GetField(
                    "screens",
                    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic
            )).SetValue($null, $null)
        } catch {}
    }

    function Get-Screens
    {
        Update-ScreensInformation

        return ([System.Windows.Forms.Screen]::AllScreens | Sort-Object -Property Primary -Descending)
    }

    function ConvertTo-ScreenObject
    {
        <#
            .SYNOPSIS
                Take a .NET Screen and convert to Arcane Screen object.

            .PARAMETER ScreenName
                Type: String
                Default: None
                Description: .NET Screen object to be converted.
        #>
        param (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Screen]$screen
        )

        return New-Object -TypeName PSCustomObject -Property @{
            Id = $i
            Name = $screen.DeviceName
            Primary = $screen.Primary
            Width = $screen.Bounds.Width
            Height = $screen.Bounds.Height
            X = $screen.Bounds.X
            Y = $screen.Bounds.Y
        }
    }

    function Get-ScreenObjectList
    {
        <#
            .SYNOPSIS
                Return an array of screen objects.

            .DESCRIPTION
                A screen refer to physical or virtual screen (monitor).

        #>
        $result = @()

        $i = 0
        foreach ($screen in (Get-Screens))
        {
            $i++

            $result += (ConvertTo-ScreenObject -screen $screen)
        }

        return ,$result
    }

    function Compare-ScreenInformation
    {
        <#
            .SYNOPSIS
                Compare two screen objects.

            .DESCRIPTION
                Compare two screen objects and return true if they are different.

            .PARAMETER screenToCompare
                Type: System.Windows.Forms.Screen
                Default: None
                Description: Screen object to compare with updated and matching screen object.
        #>
        param (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Screen]$screenToCompare
        )

        $screens = (Get-Screens | Sort-Object -Property Primary -Descending)
        $screen = $screens | Where-Object { $_.DeviceName -eq $screenToCompare.DeviceName }

        if (-not $screen)
        {
            return $true
        }

        return (
            $screen.Bounds.Width -ne $screenToCompare.Bounds.Width -or
            $screen.Bounds.Height -ne $screenToCompare.Bounds.Height -or
            $screen.Bounds.X -ne $screenToCompare.Bounds.X -or
            $screen.Bounds.Y -ne $screenToCompare.Bounds.Y
        )
    }

    $mirrorDesktop_DC = [IntPtr]::Zero
    $desktop_DC = [IntPtr]::Zero
    $mirrorDesktop_hBmp = [IntPtr]::Zero
    $spaceBlock_DC = [IntPtr]::Zero
    $spaceBlock_hBmp = [IntPtr]::Zero
    $dirtyRect_DC = [IntPtr]::Zero
    $pBitmapInfoHeader = [IntPtr]::Zero

    $SRCCOPY = 0x00CC0020
    $DIB_RGB_COLORS = 0x0
    try
    {
        $screens = New-Object PSCustomObject -Property @{
            List = (Get-ScreenObjectList)
        }
        $Client.WriteJson($screens)

        $screen = $null

        $viewerExpectation = $Client.ReadLine() | ConvertFrom-Json

        if ($viewerExpectation.PSobject.Properties.name -contains "ScreenName")
        {
            $screen = (Get-Screens) | Where-Object -FilterScript {
                $_.DeviceName -eq $viewerExpectation.ScreenName
            }

            # Add other parameters if needed
        }

        if (-not $screen)
        {
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen
            if (-not $screen)
            {
                return
            }
        }

        # Default
        $blockSize = 64
        $packetSize = 4096
        $compressionQuality = 100

        # User-defined (Optional)
        if ($viewerExpectation.PSobject.Properties.name -contains "BlockSize")
        {
            $blockSize = $viewerExpectation.BlockSize
        }

        if ($viewerExpectation.PSobject.Properties.name -contains "PacketSize")
        {
            $packetSize = $viewerExpectation.PacketSize
        }

        if ($viewerExpectation.PSobject.Properties.name -contains "ImageCompressionQuality")
        {
            $compressionQuality = $viewerExpectation.ImageCompressionQuality
        }

        $encoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
        $encoderParameters.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
            [System.Drawing.Imaging.Encoder]::Quality,
            $compressionQuality
        )

        $encoder = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' };

        $collapsed = $false

        while ($SafeHash.SessionActive)
        {
            if ($collapsed)
            {
                break
            }

            $SpaceGrid = $null
            try
            {
                $firstIteration = $true

                # Create our desktop mirror (For speeding up BitBlt calls)
                $screenBounds = $screen.Bounds

                $horzBlockCount = [math]::ceiling($screenBounds.Width / $blockSize)
                $vertBlockCount = [math]::ceiling($screenBounds.Height / $blockSize)

                $SpaceGrid = New-Object IntPtr[][] $vertBlockCount, $horzBlockCount

                [IntPtr] $desktop_DC = [User32]::GetDC([IntPtr]::Zero)
                [IntPtr] $mirrorDesktop_DC = [GDI32]::CreateCompatibleDC($desktop_DC)

                [IntPtr] $mirrorDesktop_hBmp = [GDI32]::CreateCompatibleBitmap(
                    $desktop_DC,
                    $screenBounds.Width,
                    $screenBounds.Height
                )

                $null = [GDI32]::SelectObject($mirrorDesktop_DC, $mirrorDesktop_hBmp)

                # Create our block of space for change detection

                <#
                    typedef struct tagBITMAPINFOHEADER {
                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x0
                        DWORD biSize;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x4
                        LONG  biWidth;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x8
                        LONG  biHeight;

                        // x86-32|64: 0x2 Bytes | Padding = 0x0 | Offset: 0xc
                        WORD  biPlanes;

                        // x86-32|64: 0x2 Bytes | Padding = 0x0 | Offset: 0xe
                        WORD  biBitCount;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x10
                        DWORD biCompression;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x14
                        DWORD biSizeImage;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x18
                        LONG  biXPelsPerMeter;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x1c
                        LONG  biYPelsPerMeter;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x20
                        DWORD biClrUsed;

                        // x86-32|64: 0x4 Bytes | Padding = 0x0 | Offset: 0x24
                        DWORD biClrImportant;
                    } BITMAPINFOHEADER, *LPBITMAPINFOHEADER, *PBITMAPINFOHEADER;

                    // x86-32|64 Struct Size: 0x28 (40 Bytes)
                    // BITMAPINFO = BITMAPINFOHEADER (0x28) + RGBQUAD (0x4) = 0x2c
                #>

                $bitmapInfoHeaderSize = 0x28
                $bitmapInfoSize = $bitmapInfoHeaderSize + 0x4

                $pBitmapInfoHeader = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($bitmapInfoSize)

                # ZeroMemory
                for ($i = 0; $i -lt $bitmapInfoSize; $i++)
                {
                    [System.Runtime.InteropServices.Marshal]::WriteByte($pBitmapInfoHeader, $i, 0x0)
                }

                $BITSPIXEL = 12
                $PLANES = 14
                $biBitCount = [GDI32]::GetDeviceCaps($mirrorDesktop_DC, $BITSPIXEL)
                $biPlanes = [GDI32]::GetDeviceCaps($mirrorDesktop_DC, $PLANES)

                [System.Runtime.InteropServices.Marshal]::WriteInt32($pBitmapInfoHeader, 0x0, $bitmapInfoHeaderSize) # biSize
                [System.Runtime.InteropServices.Marshal]::WriteInt32($pBitmapInfoHeader, 0x4, $blockSize) # biWidth
                [System.Runtime.InteropServices.Marshal]::WriteInt32($pBitmapInfoHeader, 0x8, $blockSize) # biHeight
                [System.Runtime.InteropServices.Marshal]::WriteInt16($pBitmapInfoHeader, 0xc, $biPlanes) # biPlanes
                [System.Runtime.InteropServices.Marshal]::WriteInt16($pBitmapInfoHeader, 0xe, $biBitCount) # biBitCount

                [IntPtr] $spaceBlock_DC = [GDI32]::CreateCompatibleDC(0)
                [IntPtr] $spaceBlock_Ptr = [IntPtr]::Zero

                [IntPtr] $spaceBlock_hBmp = [GDI32]::CreateDIBSection(
                    $spaceBlock_DC,
                    $pBitmapInfoHeader,
                    $DIB_RGB_COLORS,
                    [ref] $spaceBlock_Ptr,
                    [IntPtr]::Zero,
                    0
                )

                $null = [GDI32]::SelectObject($spaceBlock_DC, $spaceBlock_hBmp)

                # Create our dirty rect DC
                $dirtyRect_DC = [GDI32]::CreateCompatibleDC(0)

                # Field      | Type  | Size | Offset
                # ----------------------------------
                # Chunk Size | DWORD | 0x4  | 0x0
                # Left       | DWORD | 0x4  | 0x4
                # Top        | DWORD | 0x4  | 0x8
                # ScreenUpd  | BYTE  | 0x1  | 0xc
                # ----------------------------------
                # Total Size : 0xd (13 Bytes)
                $struct = New-Object -TypeName byte[] -ArgumentList 13

                $topLeftBlock = [System.Drawing.Point]::Empty
                $bottomRightBlock = [System.Drawing.Point]::Empty

                $blockMemSize = ((($blockSize * $biBitCount) + $biBitCount) -band -bnot $biBitCount) / 8
                $blockMemSize *= $blockSize
                $ptrBlockMemSize = [IntPtr]::New($blockMemSize)

                $dirtyRect = New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $screenBounds.Width, $screenBounds.Height

                <#
                $fps = 0
                $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
                #>

                while ($SafeHash.SessionActive)
                {
                    if ($logonUIAccess)
                    {
                        $updated = Update-CurrentThreadDesktopWidthInputDesktop
                        if ($updated)
                        {
                            # Respawn a new desktop mirror (Winlogon or Default)

                            break
                        }
                    }

                    if (Compare-ScreenInformation -screenToCompare $screen)
                    {
                        $screen = Get-Screens | Where-Object { $_.DeviceName -eq $screen.DeviceName }
                        if (-not $screen)
                        {
                            # If we cannot find the screen, we fallback to primary screen
                            $screen = [System.Windows.Forms.Screen]::PrimaryScreen
                            if (-not $screen)
                            {
                                return
                            }
                        }

                        # Only set offset `0xc` to true (ScreenUpdated). Other existing structure members will be ignored by
                        # the viewer.
                        # This is a way to tell the viewer that the screen has been updated and receive new screen information.
                        [System.Runtime.InteropServices.Marshal]::WriteByte($struct, 0xc, 0x1)

                        $Client.SSLStream.Write($struct , 0, $struct.Length)

                        # Send new screen information to the viewer
                        $Client.WriteJson((ConvertTo-ScreenObject -screen $screen))

                        # Respawn a new desktop mirror
                        break
                    }

                    $CAPTUREBLT = 0x40000000

                    # Refresh our desktop mirror (Overhead is located here)
                    # It might seems confusing, but in some scenarios, mirroring the desktop is faster than capturing the desktop directly
                    # for each screen block. In modern Windows, this does not seems to be the case anymore but for retro-compatibility, I
                    # decided to keep this method until I can confirm that it is no longer necessary (or offer it as a default option).
                    # Notice that getting rid of this BitBlt call, would considerably improve performance (almost twice)
                    $result = [GDI32]::BitBlt(
                        $mirrorDesktop_DC,
                        0,
                        0,
                        $screenBounds.Width,
                        $screenBounds.Height,
                        $desktop_DC,
                        $screenBounds.Location.X,
                        $screenBounds.Location.Y,
                        $SRCCOPY -bor $CAPTUREBLT
                    )

                    if (-not $result)
                    {
                        continue
                    }

                    $updated = $false

                    for ($y = 0; $y -lt $vertBlockCount; $y++)
                    {
                        for ($x = 0; $x -lt $horzBlockCount; $x++)
                        {
                            $null = [GDI32]::BitBlt(
                                $spaceBlock_DC,
                                0,
                                0,
                                $blockSize,
                                $blockSize,
                                $mirrorDesktop_DC,
                                ($x * $blockSize),
                                ($y * $blockSize),
                                $SRCCOPY
                            );

                            if ($firstIteration)
                            {
                                # Big bang occurs, tangent univers is getting created, where is Donnie?
                                $SpaceGrid[$y][$x] = [Runtime.InteropServices.Marshal]::AllocHGlobal($blockMemSize)

                                [Kernel32]::CopyMemory($SpaceGrid[$y][$x], $spaceBlock_Ptr, $ptrBlockMemSize)
                            }
                            else
                            {
                                if ([MSVCRT]::memcmp($spaceBlock_Ptr, $SpaceGrid[$y][$x], $ptrBlockMemSize) -ne [IntPtr]::Zero)
                                {
                                    [Kernel32]::CopyMemory($SpaceGrid[$y][$x], $spaceBlock_Ptr, $ptrBlockMemSize)

                                    if (-not $updated)
                                    {
                                        # Initialize with the first dirty block coordinates
                                        $topLeftBlock.X = $x
                                        $topLeftBlock.Y = $y

                                        $bottomRightBlock = $topLeftBlock

                                        $updated = $true
                                    }
                                    else
                                    {
                                        if ($x -lt $topLeftBlock.X)
                                        {
                                            $topLeftBlock.X = $x
                                        }

                                        if ($y -lt $topLeftBlock.Y)
                                        {
                                            $topLeftBlock.Y = $y
                                        }

                                        if ($x -gt $bottomRightBlock.X)
                                        {
                                            $bottomRightBlock.X = $x
                                        }

                                        if ($y -gt $bottomRightBlock.Y)
                                        {
                                            $bottomRightBlock.Y = $y
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if ($updated)
                    {
                        # Create new updated rectangle pointing to the dirty region (since last snapshot)
                        $dirtyRect.X = $topLeftBlock.X * $blockSize
                        $dirtyRect.Y = $topLeftBlock.Y * $blockSize

                        $dirtyRect.Width = (($bottomRightBlock.X * $blockSize) + $blockSize) - $dirtyRect.Left
                        $dirtyRect.Height = (($bottomRightBlock.Y * $blockSize) + $blockSize) - $dirtyRect.Top
                    }

                    if ($updated -or $firstIteration)
                    {
                        try
                        {
                            $dirtyRect_hBmp = [GDI32]::CreateCompatibleBitmap(
                                $mirrorDesktop_DC,
                                $dirtyRect.Width,
                                $dirtyRect.Height
                            )

                            $null = [GDI32]::SelectObject($dirtyRect_DC, $dirtyRect_hBmp)

                            $null = [GDI32]::BitBlt(
                                $dirtyRect_DC,
                                0,
                                0,
                                $dirtyRect.Width,
                                $dirtyRect.Height,
                                $mirrorDesktop_DC,
                                $dirtyRect.X,
                                $dirtyRect.Y,
                                $SRCCOPY
                            )

                            [System.Drawing.Bitmap] $updatedDesktop = [System.Drawing.Image]::FromHBitmap($dirtyRect_hBmp)

                            $desktopStream = New-Object System.IO.MemoryStream

                            $updatedDesktop.Save($desktopStream, $encoder, $encoderParameters)

                            $desktopStream.Position = 0

                            try
                            {
                                # One call please
                                [System.Runtime.InteropServices.Marshal]::WriteInt32($struct, 0x0, $desktopStream.Length)
                                [System.Runtime.InteropServices.Marshal]::WriteInt32($struct, 0x4, $dirtyRect.Left)
                                [System.Runtime.InteropServices.Marshal]::WriteInt32($struct, 0x8, $dirtyRect.Top)
                                [System.Runtime.InteropServices.Marshal]::WriteByte($struct, 0xc, 0x0)

                                $Client.SSLStream.Write($struct , 0, $struct.Length)

                                $binaryReader = New-Object System.IO.BinaryReader($desktopStream)
                                do
                                {
                                    $bufferSize = ($desktopStream.Length - $desktopStream.Position)
                                    if ($bufferSize -gt $packetSize)
                                    {
                                        $bufferSize = $packetSize
                                    }

                                    $Client.SSLStream.Write($binaryReader.ReadBytes($bufferSize), 0, $bufferSize)
                                } until ($desktopStream.Position -eq $desktopStream.Length)
                            }
                            catch
                            {
                                $collapsed = $true
                                break
                            }
                        }
                        finally
                        {
                            if ($dirtyRect_hBmp -ne [IntPtr]::Zero)
                            {
                                $null = [GDI32]::DeleteObject($dirtyRect_hBmp)
                            }

                            if ($desktopStream)
                            {
                                $desktopStream.Dispose()
                            }

                            if ($updatedDesktop)
                            {
                                $updatedDesktop.Dispose()
                            }
                        }
                    }

                    if ($firstIteration)
                    {
                        $firstIteration = $false
                    }

                    <#
                    $fps++
                    if ($Stopwatch.ElapsedMilliseconds -ge 1000)
                    {
                        $HostSyncHash.host.ui.WriteLine($fps)
                        $fps = 0

                        $Stopwatch.Restart()
                    }
                    #>
                }
            }
            finally
            {
                # Free allocated resources
                if ($mirrorDesktop_DC -ne [IntPtr]::Zero)
                {
                    $null = [GDI32]::DeleteDC($mirrorDesktop_DC)
                }

                if ($mirrorDesktop_hBmp -ne [IntPtr]::Zero)
                {
                    $null = [GDI32]::DeleteObject($mirrorDesktop_hBmp)
                }

                if ($spaceBlock_DC -ne [IntPtr]::Zero)
                {
                    $null = [GDI32]::DeleteDC($spaceBlock_DC)
                }

                if ($spaceBlock_hBmp -ne [IntPtr]::Zero)
                {
                    $null = [GDI32]::DeleteObject($spaceBlock_hBmp)
                }

                if ($dirtyRect_DC -ne [IntPtr]::Zero)
                {
                    $null = [GDI32]::DeleteDC($dirtyRect_DC)
                }

                if ($pBitmapInfoHeader -ne [IntPtr]::Zero)
                {
                    [System.Runtime.InteropServices.Marshal]::FreeHGlobal($pBitmapInfoHeader)
                }

                if ($desktop_DC -ne [IntPtr]::Zero)
                {
                    $null = [User32]::ReleaseDC([IntPtr]::Zero, $desktop_DC)
                }

                # Tangent univers big crunch
                for ($y = 0; $y -lt $vertBlockCount; $y++)
                {
                    for ($x = 0; $x -lt $horzBlockCount; $x++)
                    {
                        [Runtime.InteropServices.Marshal]::FreeHGlobal($SpaceGrid[$y][$x])
                    }
                }
            }
        }
    }
    catch { }
}

# -------------------------------------------------------------------------------

$global:HandleInputEvent_ScriptBlock = {
    enum MouseFlags {
        MOUSEEVENTF_ABSOLUTE = 0x8000
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP = 0x0004
        MOUSEEVENTF_MIDDLEDOWN = 0x0020
        MOUSEEVENTF_MIDDLEUP = 0x0040
        MOUSEEVENTF_MOVE = 0x0001
        MOUSEEVENTF_RIGHTDOWN = 0x0008
        MOUSEEVENTF_RIGHTUP = 0x0010
        MOUSEEVENTF_WHEEL = 0x0800
        MOUSEEVENTF_XDOWN = 0x0080
        MOUSEEVENTF_XUP = 0x0100
        MOUSEEVENTF_HWHEEL = 0x01000
    }

    enum InputEvent {
        Keyboard = 0x1
        MouseClickMove = 0x2
        MouseWheel = 0x3
        KeepAlive = 0x4
        ClipboardUpdated = 0x5
    }

    enum MouseState {
        Up = 0x1
        Down = 0x2
        Move = 0x3
    }

    enum ClipboardMode {
        Disabled = 1
        Receive = 2
        Send = 3
        Both = 4
    }

    $SM_CXSCREEN = 0
    $SM_CYSCREEN = 1

    function Set-MouseCursorPos
    {
        param(
            [int] $X = 0,
            [int] $Y = 0
        )

        $x_screen = [User32]::GetSystemMetrics($SM_CXSCREEN)
        $y_screen = [User32]::GetSystemMetrics($SM_CYSCREEN)

        [User32]::mouse_event(
            [int][MouseFlags]::MOUSEEVENTF_MOVE -bor [int][MouseFlags]::MOUSEEVENTF_ABSOLUTE,
            (65535 * $X) / $x_screen,
            (65535 * $Y) / $y_screen,
            0,
            0
        );
    }

    function Invoke-SetClipboardData {
        <#
            .SYNOPSIS
                Set text to clipboard using Windows API only.
    
            .DESCRIPTION
                Using Windows API to set text to clipboard is required to support MTA Runspaces.
    
            .PARAMETER Text
                Text to set to clipboard.
        #>
        param(
            [Parameter(Mandatory=$true)]
            [string]$Text
        )

        if (-not [User32]::OpenClipboard([IntPtr]::Zero)) {
            throw [WinAPIException]::new("OpenClipboard")
        }
        try
        {
            if (-not [User32]::EmptyClipboard()) {
                throw [WinAPIException]::new("EmptyClipboard")
            }
    
            $hGlobal = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni($Text)
            if ($hGlobal -eq [IntPtr]::Zero) {
                return
            }

            $CF_UNICODETEXT = 13
            if ([User32]::SetClipboardData($CF_UNICODETEXT, $hGlobal) -eq [IntPtr]::Zero) {
                [System.Runtime.InteropServices.Marshal]::FreeHGlobal($hGlobal)

                throw [WinAPIException]::new("SetClipboardData")
            }
        }
        finally
        {
            $null = [User32]::CloseClipboard()
        }
    }

    function Invoke-InputEvent
    {
        param([PSCustomObject] $aEvent = $null)

        try
        {
            if (-not $aEvent)
            { return }

            switch ([InputEvent] $aEvent.Id)
            {
                # Keyboard Input Simulation
                ([InputEvent]::Keyboard)
                {
                    if (-not ($aEvent.PSobject.Properties.name -match "Keys"))
                    { break }

                    if ($aEvent.Keys -eq "^{ESC}")
                    {
                        # When running as interactive SYSTEM process, `^{ESC}` is not working from `SendWait` method
                        # Instead we rely on `keybd_event` method to simulate the key combination
                        [User32]::keybd_event($VK_LWIN, 0, $KEYEVENTF_KEYDOWN, [UIntPtr]::Zero)
                        [User32]::keybd_event($VK_LWIN, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)
                    }
                    else
                    {
                        [System.Windows.Forms.SendKeys]::SendWait($aEvent.Keys)
                    }

                    break
                }

                # Mouse Move & Click Simulation
                ([InputEvent]::MouseClickMove)
                {
                    if (-not ($aEvent.PSobject.Properties.name -match "Type"))
                    { break }

                    switch ([MouseState] $aEvent.Type)
                    {
                        # Mouse Down/Up
                        {($_ -eq ([MouseState]::Down)) -or ($_ -eq ([MouseState]::Up))}
                        {
                            #[User32]::SetCursorPos($aEvent.X, $aEvent.Y)
                            Set-MouseCursorPos -X $aEvent.X -Y $aEvent.Y

                            $down = ($_ -eq ([MouseState]::Down))

                            $mouseCode = [int][MouseFlags]::MOUSEEVENTF_LEFTDOWN
                            if (-not $down)
                            {
                                $mouseCode = [int][MouseFlags]::MOUSEEVENTF_LEFTUP
                            }

                            switch($aEvent.Button)
                            {
                                "Right"
                                {
                                    if ($down)
                                    {
                                        $mouseCode = [int][MouseFlags]::MOUSEEVENTF_RIGHTDOWN
                                    }
                                    else
                                    {
                                        $mouseCode = [int][MouseFlags]::MOUSEEVENTF_RIGHTUP
                                    }

                                    break
                                }

                                "Middle"
                                {
                                    if ($down)
                                    {
                                        $mouseCode = [int][MouseFlags]::MOUSEEVENTF_MIDDLEDOWN
                                    }
                                    else
                                    {
                                        $mouseCode = [int][MouseFlags]::MOUSEEVENTF_MIDDLEUP
                                    }
                                }
                            }
                            [User32]::mouse_event($mouseCode, 0, 0, 0, 0);

                            break
                        }

                        # Mouse Move
                        ([MouseState]::Move)
                        {
                            #[User32]::SetCursorPos($aEvent.X, $aEvent.Y)
                            Set-MouseCursorPos -X $aEvent.X -Y $aEvent.Y

                            break
                        }
                    }

                    break
                }

                # Mouse Wheel Simulation
                ([InputEvent]::MouseWheel) {
                    [User32]::mouse_event([int][MouseFlags]::MOUSEEVENTF_WHEEL, 0, 0, $aEvent.Delta, 0);

                    break
                }

                # Clipboard Update
                ([InputEvent]::ClipboardUpdated)
                {
                    if ($Clipboard -eq ([ClipboardMode]::Disabled) -or $Clipboard -eq ([ClipboardMode]::Send))
                    { break }

                    if (-not ($aEvent.PSobject.Properties.name -match "Text"))
                    { break }

                    $HostSyncHash.ClipboardText = $aEvent.Text

                    Invoke-SetClipboardData -Text $aEvent.Text
                }
            }
        }
        catch {}
    }
}

# -------------------------------------------------------------------------------

$global:IngressEventScriptBlock = {
    if ($LogonUIAccess)
    {
        $BouncedInputControl_ScriptBlock = {
            # Important Notice: Capturing both the regular and secure desktop (e.g., LogonUI, UAC prompts)
            # in a single process is typically not feasible. My approach for the remote desktop thread involves
            # dynamically switching desktops, as the thread does not directly interact with the desktop and can
            # update its desktop context. However, this method does not work with the Input Thread due to its
            # direct interaction with the desktop, which prevents updating the thread's desktop context.
            # To resolve this, I created an additional thread(s) (Runspace(s)) and switched it to the detected desktop.
            # This allows interaction with the desktop without issues, with each Input thread managing its own desktop.
            # While this method is working, it may not be optimal, future versions may explore alternatives that do
            # not require multiple threads by avoiding methods that lock the current thread's desktop and using WinAPI's
            # Extensive testing will be necessary to validate these potential solutions.
            Update-CurrentThreadDesktop -DesktopName $desktopName

            while ($SafeHash.SessionActive)
            {
                $inputEvent.WaitOne()

                # Check if even concern current input desktop before popping items from the queue
                if (Get-CurrentThreadDesktopName -eq $desktopName)
                {
                    $aEvent = $null
                    while ($inputQueue.TryDequeue([ref]$aEvent)) {
                        Invoke-InputEvent -aEvent $aEvent
                    }

                    $inputEvent.Reset()
                }
            }
        }

        $desktop_runspaces = @{}

        $defaultDesktopName = Get-CurrentThreadDesktopName
        $inputQueue = [System.Collections.Concurrent.ConcurrentQueue[Object]]::new()
        $inputEvent = [System.Threading.ManualResetEvent]::new($false)
    }

    try
    {
        while ($SafeHash.SessionActive)
        {
            try
            {
                $jsonEvent = $Reader.ReadLine()
            }
            catch
            {
                # ($_ | Out-File "c:\temp\debug.txt")

                break
            }

            try
            {
                $aEvent = $jsonEvent | ConvertFrom-Json
            }
            catch { continue }

            if (-not ($aEvent.PSobject.Properties.name -match "Id"))
            { continue }

            if ($LogonUIAccess)
            {
                $currentInputDesktopName = Get-InputDesktopName
                if ($currentInputDesktopName -ne $defaultDesktopName)
                {
                    if (-not $desktop_runspaces.ContainsKey($currentInputDesktopName))
                    {
                        $params = @{
                            "inputQueue" = $inputQueue
                            "inputEvent" = $inputEvent
                            "desktopName" = $currentInputDesktopName
                            "SafeHash" = $SafeHash
                            "HostSyncHash" = $HostSyncHash
                            "Clipboard" = $Clipboard
                        }

                        $desktop_runspaces[$currentInputDesktopName] = (
                            New-RunSpace -RunspaceApartmentState "MTA" -ScriptBlocks @(
                                # Runspace Required Functions
                                $global:WinAPI_Const_ScriptBlock,
                                $global:WinAPIException_Class_ScriptBlock,
                                $global:UpdateCurrentThreadDesktop_Func_ScriptBlock,
                                $global:GetCurrentThreadDesktopName_Func_ScriptBlock,
                                $global:GetUserObjectInformation_Func_ScriptBlock,
                                $global:HandleInputEvent_ScriptBlock,

                                # Runspace Entrypoint
                                $BouncedInputControl_ScriptBlock
                        ) -Params $params)
                    }

                    # Bounce the input event to the current input desktop (Probably Winlogon)
                    $inputQueue.Enqueue($aEvent)
                    $null = $inputEvent.Set()

                    continue
                }
            }

            # Handle Event in default desktop
            Invoke-InputEvent -aEvent $aEvent
        }
    }
    finally
    {
        if ($LogonUIAccess)
        {
            foreach ($desktop_runspaces in $runspaces.Values)
            {
                $null = $desktop_runspaces.PowerShell.EndInvoke($desktop_runspaces.AsyncResult)
                $desktop_runspaces.PowerShell.Runspace.Dispose()
                $desktop_runspaces.PowerShell.Dispose()
            }
        }
    }
}

# -------------------------------------------------------------------------------

$global:EgressEventScriptBlock = {

    enum CursorType {
        IDC_APPSTARTING = 32650
        IDC_ARROW = 32512
        IDC_CROSS = 32515
        IDC_HAND = 32649
        IDC_HELP = 32651
        IDC_IBEAM = 32513
        IDC_ICON = 32641
        IDC_NO = 32648
        IDC_SIZE = 32640
        IDC_SIZEALL = 32646
        IDC_SIZENESW = 32643
        IDC_SIZENS = 32645
        IDC_SIZENWSE = 32642
        IDC_SIZEWE = 32644
        IDC_UPARROW = 32516
        IDC_WAIT = 32514
    }

    enum OutputEvent {
        KeepAlive = 0x1
        MouseCursorUpdated = 0x2
        ClipboardUpdated = 0x3
        DesktopActive = 0x4
        DesktopInactive = 0x5
    }

    enum ClipboardMode {
        Disabled = 1
        Receive = 2
        Send = 3
        Both = 4
    }

    function Initialize-Cursors
    {
        <#
            .SYNOPSIS
                Initialize different Windows supported mouse cursors.

            .DESCRIPTION
                Unfortunately, there is not WinAPI to get current mouse cursor icon state (Ex: as a flag)
                but only current mouse cursor icon (via its handle).

                One solution, is to resolve each supported mouse cursor handles (HCURSOR) with corresponding name
                in a hashtable and then compare with GetCursorInfo() HCURSOR result.
        #>
        $cursors = @{}

        foreach ($cursorType in [CursorType].GetEnumValues()) {
            $result = [User32]::LoadCursorA(0, [int]$cursorType)

            if ($result -gt 0)
            {
                $cursors[[string] $cursorType] = $result
            }
        }

        return $cursors
    }

    function Get-GlobalMouseCursorIconHandle
    {
        <#
            .SYNOPSIS
                Return global mouse cursor handle.
            .DESCRIPTION
                For this project I really want to avoid using "inline c#" but only pure PowerShell Code.
                I'm using a Hackish method to retrieve the global Windows cursor info by playing by hand
                with memory to prepare and read CURSORINFO structure.
                ---
                typedef struct tagCURSORINFO {
                    DWORD   cbSize;       // Size: 0x4
                    DWORD   flags;        // Size: 0x4
                    HCURSOR hCursor;      // Size: 0x4 (32bit) , 0x8 (64bit)
                    POINT   ptScreenPos;  // Size: 0x8
                } CURSORINFO, *PCURSORINFO, *LPCURSORINFO;
                Total Size of Structure:
                    - [32bit] 20 Bytes
                    - [64bit] 24 Bytes
        #>

        # sizeof(cbSize) + sizeof(flags) + sizeof(ptScreenPos) = 16
        $structSize = [IntPtr]::Size + 16

        $cursorInfo = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($structSize)
        try
        {
            # ZeroMemory(@cursorInfo, SizeOf(tagCURSORINFO))
            for ($i = 0; $i -lt $structSize; $i++)
            {
                [System.Runtime.InteropServices.Marshal]::WriteByte($cursorInfo, $i, 0x0)
            }

            [System.Runtime.InteropServices.Marshal]::WriteInt32($cursorInfo, 0x0, $structSize)

            if ([User32]::GetCursorInfo($cursorInfo))
            {
                $hCursor = [System.Runtime.InteropServices.Marshal]::ReadInt64($cursorInfo, 0x8)

                return $hCursor
            }

            <#for ($i = 0; $i -lt $structSize; $i++)
            {
                $offsetValue = [System.Runtime.InteropServices.Marshal]::ReadByte($cursorInfo, $i)
                Write-Host "Offset: ${i} -> " -NoNewLine
                Write-Host $offsetValue -ForegroundColor Green -NoNewLine
                Write-Host ' (' -NoNewLine
                Write-Host ('0x{0:x}' -f $offsetValue) -ForegroundColor Cyan -NoNewLine
                Write-Host ')'
            }#>
        }
        finally
        {
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($cursorInfo)
        }
    }

    function Send-Event
    {
        <#
            .SYNOPSIS
                Send an event to remote peer.

            .PARAMETER AEvent
                Type: Enum
                Default: None
                Description: The event to send to remote viewer.

            .PARAMETER Data
                Type: PSCustomObject
                Default: None
                Description: Additional information about the event.
        #>
        param (
            [Parameter(Mandatory=$True)]
            [OutputEvent] $AEvent,

            [PSCustomObject] $Data = $null
        )

        try
        {
            if (-not $Data)
            {
                $Data = New-Object -TypeName PSCustomObject -Property @{
                    Id = $AEvent
                }
            }
            else
            {
                $Data | Add-Member -MemberType NoteProperty -Name "Id" -Value $AEvent
            }

            $Writer.WriteLine(($Data | ConvertTo-Json -Compress))

            return $true
        }
        catch
        {
            return $false
        }
    }

    $cursors = Initialize-Cursors

    $oldCursor = 0

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($SafeHash.SessionActive)
    {
        # Events that occurs every seconds needs to be placed bellow.
        # If no event has occured during this second we send a Keep-Alive signal to
        # remote peer and detect a potential socket disconnection.
        if ($stopWatch.ElapsedMilliseconds -ge 1000)
        {
            try
            {
                $eventTriggered = $false

                # Clipboard Update Detection
                if (
                    ($Clipboard -eq ([ClipboardMode]::Both) -or $Clipboard -eq ([ClipboardMode]::Send))
                )
                {
                    # IDEA: Check for existing clipboard change event or implement a custom clipboard
                    # change detector using "WM_CLIPBOARDUPDATE" for example (WITHOUT INLINE C#)
                    # It is not very important but it would avoid calling "Get-Clipboard" every seconds.
                    $currentClipboard = (Get-Clipboard -Raw)

                    if ($currentClipboard -and $currentClipboard -cne $HostSyncHash.ClipboardText)
                    {
                        $data = New-Object -TypeName PSCustomObject -Property @{
                            Text = $currentClipboard
                        }

                        if (-not (Send-Event -AEvent ([OutputEvent]::ClipboardUpdated) -Data $data))
                        { break }

                        $HostSyncHash.ClipboardText = $currentClipboard

                        $eventTriggered = $true
                    }
                }

                # Send a Keep-Alive if during this second iteration nothing happened.
                if (-not $eventTriggered)
                {
                    if (-not (Send-Event -AEvent ([OutputEvent]::KeepAlive)))
                    { break }
                }
            }
            finally
            {
                $stopWatch.Restart()
            }
        }

        # Monitor for global mouse cursor change
        # Update Frequently (Maximum probe time to be efficient: 50ms)
        $currentCursor = Get-GlobalMouseCursorIconHandle
        if ($currentCursor -ne 0 -and $currentCursor -ne $oldCursor)
        {
            $cursorTypeName = ($cursors.GetEnumerator() | Where-Object { $_.Value -eq $currentCursor }).Key

            $data = New-Object -TypeName PSCustomObject -Property @{
                Cursor = $cursorTypeName
            }

            if (-not (Send-Event -AEvent ([OutputEvent]::MouseCursorUpdated) -Data $data))
            { break }

            $oldCursor = $currentCursor
        }

        Start-Sleep -Milliseconds 50
    }

    $stopWatch.Stop()
}

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Local Functions                                                              #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

function Write-Banner
{
    <#
        .SYNOPSIS
            Output cool information about current PowerShell module to terminal.
    #>

    Write-Host ""
    Write-Host "Arcane Server " -NoNewLine
    Write-Host $global:ArcaneVersion -ForegroundColor Cyan
    Write-Host "Jean-Pierre LESUEUR (" -NoNewLine
    Write-Host "@DarkCoderSc" -NoNewLine -ForegroundColor Green
    Write-Host ") " -NoNewLine
    Write-Host ""
    Write-Host "License: Apache License (Version 2.0, January 2004)"
    Write-Host ""
}

function Write-Log
{
    <#
        .SYNOPSIS
            Output a log message to terminal with associated "icon".

        .PARAMETER Message
            Type: String
            Default: None

            Description: The message to write to terminal.

        .PARAMETER LogKind
            Type: LogKind Enum
            Default: Information

            Description: Define the logger "icon" kind.
    #>
    param(
        [Parameter(Mandatory=$True)]
        [string] $Message,

        [LogKind] $LogKind = [LogKind]::Information
    )

    switch ($LogKind)
    {
        ([LogKind]::Warning)
        {
            $icon = "!!"
            $color = [System.ConsoleColor]::Yellow

            break
        }

        ([LogKind]::Success)
        {
            $icon = "OK"
            $color = [System.ConsoleColor]::Green

            break
        }

        ([LogKind]::Error)
        {
            $icon = "KO"
            $color = [System.ConsoleColor]::Red

            break
        }

        default
        {
            $color = [System.ConsoleColor]::Cyan
            $icon = "i"
        }
    }

    Write-Host "[ " -NoNewLine
    Write-Host $icon -ForegroundColor $color -NoNewLine
    Write-Host " ] $Message"
}

# -------------------------------------------------------------------------------

function Write-OperationSuccessState
{
    param(
        [Parameter(Mandatory=$True)]
        $Result,

        [Parameter(Mandatory=$True)]
        $Message
    )

    if ($Result)
    {
        $kind = [LogKind]::Success
    }
    else
    {
        $kind = [LogKind]::Error
    }

    Write-Log -Message $Message -LogKind $kind
}

# -------------------------------------------------------------------------------

function Invoke-PreventSleepMode
{
    <#
        .SYNOPSIS
            Prevent computer to enter sleep mode while server is running.

        .DESCRIPTION
            Function returns thread execution state old flags value. You can use this old flags
            to restore thread execution to its original state.
    #>

    $ES_AWAYMODE_REQUIRED = [uint32]"0x00000040"
    $ES_CONTINUOUS = [uint32]"0x80000000"
    $ES_SYSTEM_REQUIRED = [uint32]"0x00000001"

    return [Kernel32]::SetThreadExecutionState(
        $ES_CONTINUOUS -bor
        $ES_SYSTEM_REQUIRED -bor
        $ES_AWAYMODE_REQUIRED
    )
}

# -------------------------------------------------------------------------------

function Update-ThreadExecutionState
{
    <#
        .SYNOPSIS
            Update current thread execution state flags.

        .PARAMETER Flags
            Execution state flags.
    #>
    param(
        [Parameter(Mandatory=$True)]
        $Flags
    )

    return [Kernel32]::SetThreadExecutionState($Flags) -ne 0
}

# -------------------------------------------------------------------------------

function Get-PlainTextPassword
{
    <#
        .SYNOPSIS
            Retrieve the plain-text version of a secure string.

        .PARAMETER SecurePassword
            The SecureString object to be reversed.

    #>
    param(
        [Parameter(Mandatory=$True)]
        [SecureString] $SecurePassword
    )

    $BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    try
    {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
    }
    finally
    {
        [Runtime.InteropServices.Marshal]::FreeBSTR($BSTR)
    }
}

# -------------------------------------------------------------------------------

function Test-PasswordComplexity
{
    <#
        .SYNOPSIS
            Check if password is sufficiently complex.

        .DESCRIPTION
            To return True, Password must follow bellow complexity rules:
                * Minimum 12 Characters.
                * One of following symbols: "!@#%^&*_".
                * At least of lower case character.
                * At least of upper case character.

        .PARAMETER SecurePasswordCandidate
            Type: SecureString
            Default: None
            Description: Secure String object containing the password to test.
    #>
    param (
        [Parameter(Mandatory=$True)]
        [SecureString] $SecurePasswordCandidate
    )

    $complexityRules = "(?=^.{12,}$)(?=.*[!@#%^&*_]+)(?=.*[a-z])(?=.*[A-Z]).*$"

    return (Get-PlainTextPassword -SecurePassword $SecurePasswordCandidate) -match $complexityRules
}

# -------------------------------------------------------------------------------

function New-RandomPassword
{
    <#
        .SYNOPSIS
            Generate a new secure password.

        .DESCRIPTION
            Generate new password candidates until one candidate match complexity rules.
            Generally only one iteration is enough but in some rare case it could be one or two more.
    #>
    do
    {
        $authorizedChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#%^&*_"

        $candidate = -join ((1..18) | ForEach-Object { Get-Random -Input $authorizedChars.ToCharArray() })

        $secureCandidate = ConvertTo-SecureString -String $candidate -AsPlainText -Force
    } until (Test-PasswordComplexity -SecurePasswordCandidate $secureCandidate)

    $candidate = $null

    return $secureCandidate
}

# -------------------------------------------------------------------------------

function Get-DefaultCertificateOrCreate
{
    <#
        .SYNOPSIS
            Get default certificate from user store or create a new one.
    #>
    param (
        [string] $SubjectName = "Arcane.Server",
        [string] $StorePath = "cert:\CurrentUser\My",
        [int] $CertExpirationInDays = 365
    )

    $certificates = Get-ChildItem -Path $StorePath | Where-Object { $_.Subject -eq "CN=" + $SubjectName }

    if (-not $certificates)
    {
        return New-SelfSignedCertificate -CertStoreLocation $StorePath `
                -NotAfter (Get-Date).AddDays($CertExpirationInDays) `
                -Subject $SubjectName
    }
    else
    {
        return $certificates[0]
    }
}

# -------------------------------------------------------------------------------

function Get-SHA512FromString
{
    <#
        .SYNOPSIS
            Return the SHA512 value from string.

        .PARAMETER String
            Type: String
            Default : None
            Description: A String to hash.

        .EXAMPLE
            Get-SHA512FromString -String "Hello, World"
    #>
    param (
        [Parameter(Mandatory=$True)]
        [string] $String
    )

    $buffer = [IO.MemoryStream]::new([byte[]][char[]]$String)

    return (Get-FileHash -InputStream $buffer -Algorithm SHA512).Hash
}

# -------------------------------------------------------------------------------

function Resolve-AuthenticationChallenge
{
    <#
        .SYNOPSIS
            Algorithm to solve the server challenge during password authentication.

        .DESCRIPTION
            Server needs to resolve the challenge and keep the solution in memory before sending
            the candidate to remote peer.

        .PARAMETER Password
            Type: SecureString
            Default: None
            Description: Secure String object containing the password for resolving challenge.

        .PARAMETER Candidate
            Type: String
            Default: None
            Description:
                Random string used to solve the challenge. This string is public and is set across network by server.
                Each time a new connection is requested to server, a new candidate is generated.

        .EXAMPLE
            Resolve-AuthenticationChallenge -Password "s3cr3t!" -Candidate "rKcjdh154@]=Ldc"
    #>
    param (
       [Parameter(Mandatory=$True)]
       [SecureString] $SecurePassword,

       [Parameter(Mandatory=$True)]
       [string] $Candidate
    )

    $pbkdf2 = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
        (Get-PlainTextPassword -SecurePassword $SecurePassword),
        [Text.Encoding]::UTF8.GetBytes($Candidate),
        1000,
        [System.Security.Cryptography.HashAlgorithmName]::SHA512
    )
    try
    {
        return -join ($pbkdf2.GetBytes(64) | ForEach-Object { "{0:X2}" -f $_ })
    }
    finally {
        $pbkdf2.Dispose()
    }
}

# -------------------------------------------------------------------------------

function Test-WinAPI
{
    <#
        .SYNOPSIS
            Check if Windows API is available on current system.

        .DESCRIPTION
            Bellow is another technique to check if a Win32 API function is available on current system.
            But I prefer to limit crashing code to validate something.
            ```
                try
                {
                    $null = # CALL TO WIN32 API FUNCTION

                    return $true
                }
                catch
                {
                    return $false
                }
            ```
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string] $LibraryName,

        [Parameter(Mandatory = $true)]
        [string] $ApiName
    )

    $hModule = [Kernel32]::LoadLibrary($LibraryName)
    try
    {
        if ($hModule -eq [IntPtr]::Zero)
        {
            return $false
        }

        $proc = [Kernel32]::GetProcAddress($hModule, $ApiName)

        return $proc -ne [IntPtr]::Zero
    }
    finally
    {
        $null = [Kernel32]::FreeLibrary($hModule)
    }
}

# -------------------------------------------------------------------------------

function Test-Administrator
{
    <#
        .SYNOPSIS
            Check if current user is administrator.
    #>
    $windowsPrincipal = New-Object Security.Principal.WindowsPrincipal(
        [Security.Principal.WindowsIdentity]::GetCurrent()
    )

    return $windowsPrincipal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Local Classes                                                                #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

class ClientIO {
    [System.Net.Sockets.TcpClient] $Client = $null
    [System.IO.StreamWriter] $Writer = $null
    [System.IO.StreamReader] $Reader = $null
    [System.Net.Security.SslStream] $SSLStream = $null


    ClientIO(
        [System.Net.Sockets.TcpClient] $Client,
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
        [bool] $UseTLSv1_3
    ) {
        if ((-not $Client) -or (-not $Certificate))
        {
            throw "ClientIO Class requires both a valid TcpClient and X509Certificate2."
        }

        $this.Client = $Client

        Write-Verbose "Create new SSL Stream..."

        $this.SSLStream = New-Object System.Net.Security.SslStream($this.Client.GetStream(), $false)

        if ($UseTLSv1_3)
        {
            $TLSVersion = [System.Security.Authentication.SslProtocols]::TLS13
        }
        else {
            $TLSVersion = [System.Security.Authentication.SslProtocols]::TLS12
        }

        Write-Verbose "Authenticate as server using ${TLSVersion}..."

        $this.SSLStream.AuthenticateAsServer(
            $Certificate,
            $false,
            $TLSVersion,
            $false
        )

        if (-not $this.SSLStream.IsEncrypted)
        {
            throw "Could not established an encrypted tunnel with remote peer."
        }

        $this.SSLStream.WriteTimeout = 5000
        $this.SSLStream.ReadTimeout = [System.Threading.Timeout]::Infinite # Default

        Write-Verbose "Open communication channels..."

        $this.Writer = New-Object System.IO.StreamWriter($this.SSLStream)
        $this.Writer.AutoFlush = $true

        $this.Reader = New-Object System.IO.StreamReader($this.SSLStream)

        Write-Verbose "Connection ready for use."
    }

    [bool] Authentify([SecureString] $SecurePassword) {
        <#
            .SYNOPSIS
                Handle authentication process with remote peer.

            .PARAMETER Password
                Type: SecureString
                Default: None
                Description: Secure String object containing the password.

            .EXAMPLE
                .Authentify((ConvertTo-SecureString -String "urCompl3xP@ssw0rd" -AsPlainText -Force))
        #>
        try
        {
            if (-not $SecurePassword) {
                throw "During client authentication, a password cannot be blank."
            }

            Write-Verbose "New authentication challenge..."

            $candidate = -join ((1..128) | ForEach-Object {Get-Random -input ([char[]](33..126))})
            $candidate = Get-SHA512FromString -String $candidate

            $challengeSolution = Resolve-AuthenticationChallenge -Candidate $candidate -SecurePassword $SecurePassword

            Write-Verbose "@Challenge:"
            Write-Verbose "Candidate: ""${candidate}"""
            Write-Verbose "Solution: ""${challengeSolution}"""
            Write-Verbose "---"

            $this.Writer.WriteLine($candidate)

            Write-Verbose "Candidate sent to client, waiting for answer..."

            $challengeReply = $this.ReadLine(5 * 1000)

            Write-Verbose "Replied solution: ""${challengeReply}"""

            # Challenge solution is a Sha512 Hash so comparison doesn't need to be sensitive (-ceq or -cne)
            if ($challengeReply -ne $challengeSolution)
            {
                $this.Writer.WriteLine(([ProtocolCommand]::Fail))

                throw "Client challenge solution does not match our solution."
            }
            else
            {
                $this.Writer.WriteLine(([ProtocolCommand]::Success))

                Write-Verbose "Password Authentication Success"

                return 280121 # True
            }
        }
        catch
        {
            throw "Password Authentication Failed. Reason: `r`n $($_)"
        }
    }

    [string] RemoteAddress() {
        return $this.Client.Client.RemoteEndPoint.Address
    }

    [int] RemotePort() {
        return $this.Client.Client.RemoteEndPoint.Port
    }

    [string] LocalAddress() {
        return $this.Client.Client.LocalEndPoint.Address
    }

    [int] LocalPort() {
        return $this.Client.Client.LocalEndPoint.Port
    }

    [string] ReadLine([int] $Timeout)
    {
        <#
            .SYNOPSIS
                Read string message from remote peer with timeout support.

            .PARAMETER Timeout
                Type: Integer
                Description: Maximum period of time to wait for incomming data.
        #>
        $defautTimeout = $this.SSLStream.ReadTimeout
        try
        {
            $this.SSLStream.ReadTimeout = $Timeout

            return $this.Reader.ReadLine()
        }
        finally
        {
            $this.SSLStream.ReadTimeout = $defautTimeout
        }
    }

    [string] ReadLine()
    {
        <#
            .SYNOPSIS
                Shortcut to Reader ReadLine method. No timeout support.
        #>
        return $this.Reader.ReadLine()
    }

    [void] WriteJson([PSCustomObject] $Object)
    {
        <#
            .SYNOPSIS
                Transform a PowerShell Object as a JSON Representation then send to remote
                peer.

            .PARAMETER Object
                Type: PSCustomObject
                Description: Object to be serialized in JSON.
        #>

        $this.Writer.WriteLine(($Object | ConvertTo-Json -Compress))
    }

    [void] WriteLine([string] $Value)
    {
        $this.Writer.WriteLine($Value)
    }

    [void] Close() {
        <#
            .SYNOPSIS
                Release streams and client.
        #>

        if ($this.Writer)
        {
            $this.Writer.Close()
        }

        if ($this.Reader)
        {
            $this.Reader.Close()
        }

        if ($this.Stream)
        {
            $this.Stream.Close()
        }

        if ($this.Client)
        {
            $this.Client.Close()
        }
    }
}

# -------------------------------------------------------------------------------

class TcpListenerEx : System.Net.Sockets.TcpListener
{
    TcpListenerEx([string] $ListenAddress, [int] $ListenPort) : base($ListenAddress, $ListenPort)
    { }

    [bool] Active()
    {
        return $this.Active
    }
}

# -------------------------------------------------------------------------------

class ServerIO {
    [TcpListenerEx] $Server = $null
    [System.IO.StreamWriter] $Writer = $null
    [System.IO.StreamReader] $Reader = $null

    ServerIO()
    { }

    [void] Listen(
        [string] $ListenAddress,
        [int] $ListenPort
    )
    {
        if ($this.Server)
        {
            $this.Close()
        }

        $this.Server = New-Object TcpListenerEx(
            $ListenAddress,
            $ListenPort
        )

        $this.Server.Start()

        Write-Verbose "Listening on ""$($ListenAddress):$($ListenPort)""..."
    }

    [ClientIO] PullClient(
        [SecureString] $SecurePassword,

        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,

        [bool] $UseTLSv13,
        [int] $Timeout
    ) {
        <#
            .SYNOPSIS
                Accept new client and associate this client with a new ClientIO Object.

            .PARAMETER Timeout
                Type: Integer
                Description:
                    By default AcceptTcpClient() will block current thread until a client connects.

                    Using Timeout and a cool technique, you can stop waiting for client after a certain amount
                    of time (In Milliseconds)

                    If Timeout is greater than 0 (Milliseconds) then connection timeout is enabled.

                    Other method: AsyncWaitHandle.WaitOne([timespan])'h:m:s') -eq $true|$false with BeginAcceptTcpClient(...)
        #>

        if (-not (Test-PasswordComplexity -SecurePasswordCandidate $SecurePassword))
        {
            throw "Client socket pull request requires a complex password to be set."
        }

        if ($Timeout -gt 0)
        {
            $socketReadList = [System.Collections.ArrayList]@($this.Server.Server)

            [System.Net.Sockets.Socket]::Select($socketReadList, $null, $null, $Timeout * 1000)

            if (-not $socketReadList.Contains($this.Server.Server))
            {
                throw "Pull timeout."
            }
        }

        $socket = $this.Server.AcceptTcpClient()

        $client = [ClientIO]::New(
            $socket,
            $Certificate,
            $UseTLSv13
        )
        try
        {
            Write-Verbose "New client socket connected from: ""$($client.RemoteAddress())""."

            $authenticated = ($client.Authentify($SecurePassword) -eq 280121)
            if (-not $authenticated)
            {
                throw "Access Denied."
            }
        }
        catch
        {
            $client.Close()

            throw $_
        }

        return $client
    }

    [bool] Active()
    {
        if ($this.Server)
        {
            return $this.Server.Active()
        }
        else
        {
            return $false
        }
    }

    [void] Close()
    {
        <#
            .SYNOPSIS
                Stop listening and release TcpListener object.
        #>
        if ($this.Server)
        {
            if ($this.Server.Active)
            {
                $this.Server.Stop()
            }

            $this.Server = $null

            Write-Verbose "Server is now released."
        }
    }
}

# -------------------------------------------------------------------------------

class ServerSession {
    [string] $Id = ""
    [bool] $ViewOnly = $false
    [ClipboardMode] $Clipboard = [ClipboardMode]::Both
    [string] $ViewerLocation = ""
    [bool] $logonUIAccess = $false

    [System.Collections.Generic.List[PSCustomObject]]
    $WorkerThreads = @()

    [System.Collections.Generic.List[ClientIO]]
    $Clients = @()

    $SafeHash = [HashTable]::Synchronized(@{
        SessionActive = $true
    })

    ServerSession(
        [bool] $ViewOnly,
        [ClipboardMode] $Clipboard,
        [string] $ViewerLocation
    )
    {
        $this.Id = (SHA512FromString -String (-join ((1..128) | ForEach-Object {Get-Random -input ([char[]](33..126))})))

        $this.ViewOnly = $ViewOnly
        $this.Clipboard = $Clipboard
        $this.ViewerLocation = $ViewerLocation

        # Check if current arcane server is running under NT AUTHORITY\SYSTEM
        # This is required to capture secure desktop (Winlogon)
        $this.logonUIAccess = [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
    }

    [bool] CompareSession([string] $Id)
    {
        <#
            .SYNOPSIS
                Compare two session object. In this case just compare session id string.

            .PARAMETER Id
                Type: String
                Description: A session id to compare with current session object.
        #>
        return ($this.Id -ceq $Id)
    }

    [void] NewDesktopWorker([ClientIO] $Client)
    {
        <#
            .SYNOPSIS
                Create a new desktop streaming worker (Runspace/Thread).

            .PARAMETER Client
                Type: ClientIO
                Description: Established connection with a remote peer.
        #>

        $this.WorkerThreads.Add(
            (
                New-RunSpace -RunspaceApartmentState "MTA" -ScriptBlocks @(
                    # Runspace Required Functions
                    $global:WinAPI_Const_ScriptBlock,
                    $global:WinAPIException_Class_ScriptBlock,
                    $global:GetUserObjectInformation_Func_ScriptBlock,
                    $global:GetInputDesktopName_Func_ScriptBlock,
                    $global:GetCurrentThreadDesktopName_Func_ScriptBlock,
                    $global:UpdateCurrentThreadDesktopWithInputDesktop_Func_ScriptBlock,

                    # Runspace Entrypoint
                    $global:DesktopStreamScriptBlock
                ) -Params @{
                    # Required Runspace Variables
                    "HostSyncHash" = $global:HostSyncHash
                    "Client" = $Client
                    "SafeHash" = $this.SafeHash
                    "LogonUIAccess" = $this.logonUIAccess
                }
            )
        )

        ###

        $this.Clients.Add($Client)
    }

    [void] NewEventWorker([ClientIO] $Client)
    {
        <#
            .SYNOPSIS
                Create a new egress / ingress worker (Runspace/Thread) to process outgoing / incomming events.

            .PARAMETER Client
                Type: ClientIO
                Description: Established connection with a remote peer.
        #>

        if ($this.ViewOnly)
        {
            # Ignore demand for event worker if session is view only.
            return
        }

        $this.WorkerThreads.Add(
            (
                New-RunSpace -ScriptBlocks @(
                    # Runspace Entrypoint
                    $global:EgressEventScriptBlock
                ) -Params @{
                    # Required Runspace Variables
                    "HostSyncHash" = $global:HostSyncHash
                    "Writer" = $Client.Writer
                    "Clipboard" = $this.Clipboard
                    "SafeHash" = $this.SafeHash
                    "LogonUIAccess" = $this.logonUIAccess
                }
            )
        )

        ###

        if ($this.LogonUIAccess)
        {
            $runspaceApartmentState = "MTA"
        }
        else
        {
            $runspaceApartmentState = "STA"
        }

        $this.WorkerThreads.Add(
            (
                New-RunSpace -RunspaceApartmentState $runspaceApartmentState -ScriptBlocks @(
                    # Runspace Required Functions
                    $global:WinAPI_Const_ScriptBlock,
                    $global:WinAPIException_Class_ScriptBlock,
                    $global:GetCurrentThreadDesktopName_Func_ScriptBlock,
                    $global:GetInputDesktopName_Func_ScriptBlock,
                    $global:GetUserObjectInformation_Func_ScriptBlock,
                    $global:NewRunSpace_Func_ScriptBlock,
                    $global:HandleInputEvent_ScriptBlock,

                    # Runspace Entrypoint
                    $global:IngressEventScriptBlock
                ) -Params @{
                    # Required Runspace Variables
                    "HostSyncHash" = $global:HostSyncHash
                    "Reader" = $Client.Reader
                    "Clipboard" = $this.Clipboard
                    "SafeHash" = $this.SafeHash
                    "LogonUIAccess" = $this.logonUIAccess

                    # Required Script Blocks (LogonUIAccess)
                    "WinAPI_Const_ScriptBlock" = $global:WinAPI_Const_ScriptBlock
                    "WinAPIException_Class_ScriptBlock" = $global:WinAPIException_Class_ScriptBlock
                    "UpdateCurrentThreadDesktop_Func_ScriptBlock" = $global:UpdateCurrentThreadDesktop_Func_ScriptBlock
                    "GetCurrentThreadDesktopName_Func_ScriptBlock" = $global:GetCurrentThreadDesktopName_Func_ScriptBlock
                    "GetInputDesktopName_Func_ScriptBlock" = $global:GetInputDesktopName_Func_ScriptBlock
                    "GetUserObjectInformation_Func_ScriptBlock" = $global:GetUserObjectInformation_Func_ScriptBlock
                    "HandleInputEvent_ScriptBlock" = $global:HandleInputEvent_ScriptBlock
                }
            )
        )

        ###

        $this.Clients.Add($Client)
    }

    [void] CheckSessionIntegrity()
    {
        <#
            .SYNOPSIS
                Check if session integrity is still respected.

            .DESCRIPTION
                We consider that a dead session, is a session with at least one worker that has completed his
                tasks.

                This will notify other workers that something happened (disconnection, fatal exception).
        #>

        foreach ($worker in $this.WorkerThreads)
        {
            if ($worker.AsyncResult.IsCompleted)
            {
                $this.Close()

                break
            }
        }
    }

    [void] Close()
    {
        <#
            .SYNOPSIS
                Close components associated with current session (Ex: runspaces, sockets etc..)
        #>

        Write-Verbose "Closing session..."

        $this.SafeHash.SessionActive = $false

        Write-Verbose "Close associated peers..."

        # Close connection with remote peers associated with this session
        foreach ($client in $this.Clients)
        {
            $client.Close()
        }

        $this.Clients.Clear()

        Write-Verbose "Wait for associated threads to finish their tasks..."

        while ($true)
        {
            $completed = $true

            foreach ($worker in $this.WorkerThreads)
            {
                if (-not $worker.AsyncResult.IsCompleted)
                {
                    $completed = $false

                    break
                }
            }

            if ($completed)
            { break }

            Start-Sleep -Seconds 1
        }

        Write-Verbose "Dispose threads (runspaces)..."

        # Terminate runspaces associated with this session
        foreach ($worker in $this.WorkerThreads)
        {
            $null = $worker.PowerShell.EndInvoke($worker.AsyncResult)
            $worker.PowerShell.Runspace.Dispose()
            $worker.PowerShell.Dispose()
        }
        $this.WorkerThreads.Clear()

        Write-Host "Session terminated with viewer: $($this.ViewerLocation)"

        Write-Verbose "Session closed."
    }
}

# -------------------------------------------------------------------------------

class SessionManager {
    [ServerIO] $Server = $null

    [System.Collections.Generic.List[ServerSession]]
    $Sessions = @()

    [SecureString] $SecurePassword = $null

    [System.Security.Cryptography.X509Certificates.X509Certificate2]
    $Certificate = $null

    [bool] $ViewOnly = $false
    [bool] $UseTLSv13 = $false

    [ClipboardMode] $Clipboard = [ClipboardMode]::Both

    SessionManager(
        [SecureString] $SecurePassword,

        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $Certificate,

        [bool] $ViewOnly,
        [bool] $UseTLSv13,
        [ClipboardMode] $Clipboard
    )
    {
        Write-Verbose "Initialize new session manager..."

        $this.SecurePassword = $SecurePassword
        $this.ViewOnly = $ViewOnly
        $this.UseTLSv13 = $UseTLSv13
        $this.Clipboard = $Clipboard

        if (-not $Certificate)
        {
            Write-Verbose "No custom certificate specified, using default X509 Certificate (Not Recommended)."

            $this.Certificate = Get-DefaultCertificateOrCreate
        }
        else
        {
            $this.Certificate = $Certificate
        }

        Write-Verbose "@Certificate:"
        Write-Verbose $this.Certificate
        Write-Verbose "---"

        Write-Verbose "Session manager initialized."
    }

    [void] OpenServer(
        [string] $ListenAddress,
        [int] $ListenPort
    )
    {
        <#
            .SYNOPSIS
                Create a new server object then start listening on desired interface / port.

            .PARAMETER ListenAddress
                Desired interface to listen for new peers.
                "127.0.0.1" = Only listen for localhost peers.
                "0.0.0.0" = Listen on all interfaces for peers.

            .PARAMETER ListenPort
                TCP Port to listen for new peers (0-65535)
        #>

        $this.CloseServer()
        try
        {
            $this.Server = [ServerIO]::New()

            $this.Server.Listen(
                $ListenAddress,
                $ListenPort
            )
        }
        catch
        {
            $this.CloseServer()

            throw $_
        }
    }

    [ServerSession] GetSession([string] $SessionId)
    {
        <#
            .SYNOPSIS
                Find a session by its id on current session pool.

            .PARAMETER SessionId
                Type: String
                Description: SessionId to retrieve from session pool.
        #>
        foreach ($session in $this.Sessions)
        {
            if ($session.CompareSession($SessionId))
            {
                return $session
            }
        }

        return $null
    }

    [void] ProceedNewSessionRequest([ClientIO] $Client)
    {
        <#
            .SYNOPSIS
                Attempt a new session request with remote peer.

            .DESCRIPTION
                Session creation is now requested from a dedicated client instead of using
                same client as for desktop streaming.

                I prefer to use a dedicated client to have a more cleaner session establishement
                process.

                Session request will basically generate a new session object, send some information
                about current server marchine state then wait for viewer acknowledgement with desired
                configuration (Ex: desired screen to capture, quality and local size constraints).

                When session creation is done, client is then closed.
        #>
        try
        {
            Write-Verbose "Remote peer as requested a new session..."

            $session = [ServerSession]::New(
                $this.ViewOnly,
                $this.Clipboard,
                $client.RemoteAddress()
            )

            Write-Verbose "@ServerSession"
            Write-Verbose "Id: ""$($session.Id)"""
            Write-Verbose "---"

            $serverInformation = New-Object PSCustomObject -Property @{
                # Session information and configuration
                SessionId = $session.Id
                Version = $global:ArcaneProtocolVersion
                ViewOnly = $this.ViewOnly
                Clipboard = $this.Clipboard

                # Local machine information
                MachineName = [Environment]::MachineName
                Username = [Environment]::UserName
                WindowsVersion = [Environment]::OSVersion.VersionString
            }

            Write-Verbose "Sending server information to remote peer..."

            Write-Verbose "@ServerInformation:"
            Write-Verbose $serverInformation
            Write-Verbose "---"

            $client.WriteJson($serverInformation)

            Write-Verbose "New session successfully created."

            $this.Sessions.Add($session)

            $client.WriteLine(([ProtocolCommand]::Success))
        }
        catch
        {
            $session = $null

            throw $_
        }
        finally
        {
            if ($client)
            {
                $client.Close()
            }
        }
    }

    [void] ProceedAttachRequest([ClientIO] $Client)
    {
        <#
            .SYNOPSIS
                Attach a new peer to an existing session then dispatch this new peer as a
                new stateful worker.

            .PARAMETER Client
                An established connection with remote peer as a ClientIO Object.
        #>
        Write-Verbose "Proceed new session attach request..."

        $session = $this.GetSession($Client.ReadLine(5 * 1000))
        if (-not $session)
        {
            $Client.WriteLine(([ProtocolCommand]::ResourceNotFound))

            throw "Could not locate session."
        }

        Write-Verbose "Client successfully attached to session: ""$($session.id)"""

        $Client.WriteLine(([ProtocolCommand]::ResourceFound))

        $workerKind = $Client.ReadLine(5 * 1000)

        switch ([WorkerKind] $workerKind)
        {
            (([WorkerKind]::Desktop))
            {
                $session.NewDesktopWorker($Client)

                break
            }

            (([WorkerKind]::Events))
            {
                $session.NewEventWorker($Client) # I/O

                break
            }
        }
    }

    [void] ListenForWorkers()
    {
        <#
            .SYNOPSIS
                Process server client queue and dispatch accordingly.
        #>
        while ($true)
        {
            if (-not $this.Server -or -not $this.Server.Active())
            {
                throw "A server must be active to listen for new workers."
            }

            try
            {
                $this.CheckSessionsIntegrity()
            }
            catch
            {  }

            $client = $null
            try
            {
                $client = $this.Server.PullClient(
                    $this.SecurePassword,
                    $this.Certificate,
                    $this.UseTLSv13,
                    5 * 1000
                )

                $requestMode = $client.ReadLine(5 * 1000)

                switch ([ProtocolCommand] $requestMode)
                {
                    ([ProtocolCommand]::RequestSession)
                    {
                        $remoteAddress = $client.RemoteAddress()

                        $this.ProceedNewSessionRequest($client)

                        Write-Host "New remote desktop session established with: $($remoteAddress)"

                        break
                    }

                    ([ProtocolCommand]::AttachToSession)
                    {
                        $this.ProceedAttachRequest($client)

                        break
                    }

                    default:
                    {
                        $client.WriteLine(([ProtocolCommand]::BadRequest))

                        throw "Bad request."
                    }
                }
            }
            catch
            {
                if ($client)
                {
                    $client.Close()

                    $client = $null
                }
            }
            finally
            { }
        }
    }

    [void] CheckSessionsIntegrity()
    {
        <#
            .SYNOPSIS
                Check if existing server sessions integrity is respected.
                Use this method to free dead/half-dead sessions.
        #>
        foreach ($session in $this.Sessions)
        {
            $session.CheckSessionIntegrity()
        }
    }

    [void] CloseSessions()
    {
        <#
            .SYNOPSIS
                Terminate existing server sessions.
        #>

        foreach ($session in $this.Sessions)
        {
            $session.Close()
        }

        $this.Sessions.Clear()
    }

    [void] CloseServer()
    {
        <#
            .SYNOPSIS
                Terminate existing server sessions then release server.
        #>

        $this.CloseSessions()

        if ($this.Server)
        {
            $this.Server.Close()

            $this.Server = $null
        }
    }
}

# -------------------------------------------------------------------------------

class ValidateFileAttribute : System.Management.Automation.ValidateArgumentsAttribute
{
    <#
        .SYNOPSIS
            Check if file argument exists on disk.
    #>

    [void]Validate([System.Object] $arguments, [System.Management.Automation.EngineIntrinsics] $engineIntrinsics)
    {
        if(-not (Test-Path -Path $arguments))
        {
            throw [System.IO.FileNotFoundException]::new()
        }
    }
}

# -------------------------------------------------------------------------------

class ValidateBase64StringAttribute : System.Management.Automation.ValidateArgumentsAttribute
{
    <#
        .SYNOPSIS
            Check if string argument is a valid Base64 String.
    #>

    [void]Validate([System.Object] $arguments, [System.Management.Automation.EngineIntrinsics] $engineIntrinsics)
    {
        [Convert]::FromBase64String($arguments)
    }
}

# ----------------------------------------------------------------------------- #
#                                                                               #
#                                                                               #
#                                                                               #
#  Arcane Entry Point                                                           #
#                                                                               #
#                                                                               #
#                                                                               #
# ----------------------------------------------------------------------------- #

function Invoke-ArcaneServer
{
    <#
        .SYNOPSIS
            Create and start a new Arcane Server.

        .DESCRIPTION
            Notices:

                1- Prefer using SecurePassword over plain-text password even if a plain-text password is getting converted to SecureString anyway.

                2- Not specifying a custom certificate using CertificateFile or EncodedCertificate result in generating a default
                self-signed certificate (if not already generated) that will get installed on local machine thus requiring administrator privilege.
                If you want to run the server as a non-privileged account, specify your own certificate location.

                3- If you don't specify a SecurePassword or Password, a random complex password will be generated and displayed on terminal
                (this password is temporary)

        .PARAMETER ListenAddress
            Type: String
            Default: 0.0.0.0
            Description: IP Address that represents the local IP address.

        .PARAMETER ListenPort
            Type: Integer
            Default: 2801 (0 - 65535)
            Description: The port on which to listen for incoming connection.

        .PARAMETER SecurePassword
            Type: SecureString
            Default: None
            Description: SecureString object containing password used to authenticate remote viewer (Recommended)

        .PARAMETER Password
            Type: String
            Default: None
            Description: Plain-Text Password used to authenticate remote viewer (Not recommended, use SecurePassword instead)

        .PARAMETER CertificateFile
            Type: String
            Default: None
            Description: A file containing valid certificate information (x509), must include the private key.

        .PARAMETER EncodedCertificate
            Type: String (Base64 Encoded)
            Default: None
            Description: A base64 representation of the whole certificate file, must include the private key.

        .PARAMETER UseTLSv1_3
            Type: Switch
            Default: False
            Description: If present, TLS v1.3 will be used instead of TLS v1.2 (Recommended if applicable to both systems)

        .PARAMETER DisableVerbosity
            Type: Switch
            Default: False
            Description: If present, program wont show verbosity messages.

        .PARAMETER Clipboard
            Type: Enum
            Default: Both
            Description:
                Define clipboard synchronization mode (Both, Disabled, Send, Receive) see bellow for more detail.

                * Disabled -> Clipboard synchronization is disabled in both side
                * Receive  -> Only incomming clipboard is allowed
                * Send     -> Only outgoing clipboard is allowed
                * Both     -> Clipboard synchronization is allowed on both side

        .PARAMETER ViewOnly (Default: None)
            Type: Swtich
            Default: False
            Description: If present, remote viewer is only allowed to view the desktop (Mouse and Keyboard are not authorized)

        .PARAMETER PreventComputerToSleep
            Type: Switch
            Default: False
            Description: If present, this option will prevent computer to enter in sleep mode while server is active and waiting for new connections.

        .PARAMETER CertificatePassword
            Type: SecureString
            Default: None
            Description: Specify the password used to open a password-protected x509 Certificate provided by user.

        .EXAMPLE
            Invoke-ArcaneServer -ListenAddress "0.0.0.0" -ListenPort 2801 -SecurePassword (ConvertTo-SecureString -String "urCompl3xP@ssw0rd" -AsPlainText -Force)
            Invoke-ArcaneServer -ListenAddress "0.0.0.0" -ListenPort 2801 -SecurePassword (ConvertTo-SecureString -String "urCompl3xP@ssw0rd" -AsPlainText -Force) -CertificateFile "c:\certs\phrozen.p12"
    #>

    param (
        [string] $ListenAddress = "0.0.0.0",

        [ValidateRange(0, 65535)]
        [int] $ListenPort = 2801,

        [SecureString] $SecurePassword = $null,
        [string] $Password = "",
        [String] $CertificateFile = $null,
        [string] $EncodedCertificate = "",
        [switch] $UseTLSv1_3,
        [switch] $DisableVerbosity,
        [ClipboardMode] $Clipboard = [ClipboardMode]::Both,
        [switch] $ViewOnly,
        [switch] $PreventComputerToSleep,
        [SecureString] $CertificatePassword = $null
    )

    $oldErrorActionPreference = $ErrorActionPreference
    $oldVerbosePreference = $VerbosePreference
    try
    {
        $ErrorActionPreference = "stop"

        if (-not $DisableVerbosity)
        {
            $VerbosePreference = "continue"
        }
        else
        {
            $VerbosePreference = "SilentlyContinue"
        }

        Write-Banner

        if ((Test-WinAPI -LibraryName "Shcore.dll" -ApiName "SetProcessDpiAwareness"))
        {
            # Windows >= 8.1
            $PROCESS_PER_MONITOR_DPI_AWARE = 2
            $null = [Shcore]::SetProcessDpiAwareness($PROCESS_PER_MONITOR_DPI_AWARE)
        }
        elseif ((Test-WinAPI -LibraryName "User32.dll" -ApiName "SetProcessDPIAware"))
        {
            # Windows >= Vista
            $null = [User32]::SetProcessDPIAware()
        }

        $Certificate = $null

        if ($CertificateFile -or $EncodedCertificate)
        {
            $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            try
            {
                if ($CertificateFile)
                {
                    if(-not (Test-Path -Path $CertificateFile))
                    {
                        throw [System.IO.FileNotFoundException]::new()
                    }

                    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $CertificateFile, $CertificatePassword
                }
                else
                {
                    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 ([Convert]::FromBase64String($EncodedCertificate)), $CertificatePassword
                }
            }
            catch
            {
                $message =  "Could not open provided x509 Certificate. Possible Reasons:`r`n" +
                            "* Provided certificate is not a valid x509 Certificate.`r`n" +
                            "* Certificate is corrupted.`r`n"

                if (-not $CertificatePassword)
                {
                    $message += "* Certificate is protected by a password.`r`n"
                }
                else
                {
                    $message += "* Provided certificate password is not valid.`r`n"
                }

                $message += "More detail: $($_)"

                throw $message
            }

            if (-not $Certificate.HasPrivateKey)
            {
                throw "Provided Certificate must have private-key included."
            }
        }

        # If plain-text password is set, we convert this password to a secured representation.
        if ($Password -and -not $SecurePassword)
        {
            $SecurePassword = (ConvertTo-SecureString -String $Password -AsPlainText -Force)
        }

        if (-not $SecurePassword)
        {
            $SecurePassword = New-RandomPassword

            Write-Host -NoNewLine "Server password: """
            Write-Host -NoNewLine $(Get-PlainTextPassword -SecurePassword $SecurePassword) -ForegroundColor green
            Write-Host """."
        }
        else
        {
            if (-not (Test-PasswordComplexity -SecurePasswordCandidate $SecurePassword))
            {
                throw "Password complexity is too weak. Please choose a password following following rules:`r`n`
                * Minimum 12 Characters`r`n`
                * One of following symbols: ""!@#%^&*_""`r`n`
                * At least of lower case character`r`n`
                * At least of upper case character`r`n"
            }
        }

        Remove-Variable -Name "Password" -ErrorAction SilentlyContinue

        try
        {
            $oldExecutionStateFlags = $null
            if ($PreventComputerToSleep)
            {
                $oldExecutionStateFlags = Invoke-PreventSleepMode

                Write-OperationSuccessState -Message "Preventing computer to entering sleep mode." -Result ($oldExecutionStateFlags -gt 0)
            }

            Write-Host "Loading remote desktop server components..."

            $sessionManager = [SessionManager]::New(
                $SecurePassword,
                $Certificate,
                $ViewOnly,
                $UseTLSv1_3,
                $Clipboard
            )

            $sessionManager.OpenServer(
                $ListenAddress,
                $ListenPort
            )

            Write-Host "Server is ready to receive new connections..."

            $sessionManager.ListenForWorkers()
        }
        finally
        {
            if ($sessionManager)
            {
                $sessionManager.CloseServer()

                $sessionManager = $null
            }

            if ($oldExecutionStateFlags)
            {
                Write-OperationSuccessState -Message "Stop preventing computer to enter sleep mode. Restore thread execution state." -Result (Update-ThreadExecutionState -Flags $oldExecutionStateFlags)
            }

            Write-Host "Remote desktop was closed."
        }
    }
    finally
    {
        $ErrorActionPreference = $oldErrorActionPreference
        $VerbosePreference = $oldVerbosePreference
    }
}

# -------------------------------------------------------------------------------

try {
    Export-ModuleMember -Function Invoke-ArcaneServer
} catch {}