function Log($msg) {
  $date = "[{0:HH:mm:ss}]" -f (Get-Date)
  Write-Host "$date $msg"
}

# Create data directory for
# - Microsoft Edge user data folder (see below why we need this)
# - video playback history of today (to avoid playing duplicate videos)
# - stdout/stderr of last run (for diagnostic purposes)
$dataDir = "$PSScriptRoot\.data"
if (-not (Test-Path $dataDir)) {
  Log "Creating $dataDir"
  New-Item $dataDir -ItemType "directory" | Out-Null
  (Get-Item $dataDir).Attributes += "Hidden"
}

# Start logging to file
Log "Starting logging to $PSScriptRoot\.data\lastrun.txt"
Start-Transcript -path "$PSScriptRoot\.data\lastrun.txt" | Out-Null
try { # see end of file

# Read the configuration
Log "Reading configuration from $PSScriptRoot\config.ps1"
. "$PSScriptRoot\config.ps1"

# Check if today's day passes the day of week filter
if ($activeDays -notcontains (Get-Date).DayOfWeek) {
  Log "Skipping, week day filter"
  exit
}

Add-Type -AssemblyName PresentationFramework

# Ask for permission to continue
if ($ask) {
  Log "Showing confirmation box"
  $answer = [System.Windows.MessageBox]::Show($askText, $askTitle, 'YesNo', 'Question')
  if ($answer -ne 'Yes') {
    Log "Exiting, answered No"
    exit
  }
  Log "Continuing, answered Yes"
}

# Check which videos were already played today
$historyPath = "$dataDir\history.txt"
$midnight = Get-Date -Hour 0 -Minute 0 -Second 0
if (Test-Path $historyPath -NewerThan $midnight) {
  Log "Appending to existing history from today: $historyPath"
  $history = @(Get-Content -Path $historyPath)
} else {
  Log "Starting new history for today: $historyPath"
  $history = @()
}

# Filter to videos not played yet
if ($history) {
  $filteredVideos = @($videos | Where-Object { $history -notcontains $_.url })
  if ($filteredVideos) {
    Log "$($filteredVideos.Count) videos left after filtering videos watched today (total: $($videos.Count))"
    $videos = $filteredVideos
  } else {
    Log "No videos left after filtering videos watched today, skipping filter"
  }
}

# Select a random video
$idx = Get-Random -Maximum $videos.Length
$video = $videos[$idx]
$videoUrl = [System.Uri]$video.url
if ($videoUrl.Host -ne "youtu.be") {
  [void][System.Windows.MessageBox]::Show(
    "Invalid url: " + $video.url + "`nMust be https://youtu.be/...`nUse the 'Share' button in YouTube!",
    "PSExercise: Configuration issue",
    "OK", "Error"
    )
  exit
}
$videoId = $videoUrl.Segments[1]
$videoStart = [int]$videoUrl.Query.Replace("?t=", "") # defaults to 0 if missing
$videoEnd = $video.t2 # optional (if missing, no auto-exit)
Log "Video: $($video.url)"

# Update history
$history += $video.url
Set-Content -Path $historyPath -Value $history
Log "Updating history"

# Low-level Win32 APIs for interacting with the browser window
$source = @"
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;

namespace Custom
{
  public class Window
  {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    internal enum AccentState
    {
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    }
  
    [StructLayout(LayoutKind.Sequential)]
    internal struct AccentPolicy
    {
        public AccentState AccentState;
        public uint AccentFlags;
        public uint GradientColor;
        public uint AnimationId;
    }
  
    internal enum WindowCompositionAttribute
    {
        WCA_ACCENT_POLICY = 19
    }
  
    [StructLayout(LayoutKind.Sequential)]
    internal struct WindowCompositionAttributeData
    {
        public WindowCompositionAttribute Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    [DllImport("user32.dll")]
    internal static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    public static void EnableBlur(IntPtr hwnd, uint rgbColor, float opacity)
    {
      // https://github.com/riverar/sample-win32-acrylicblur

      // BGR color format
      uint blurBackgroundColor = (rgbColor & 0xFF0000) >> 16 | (rgbColor & 0x00FF00) | (rgbColor & 0x0000FF) << 16;
      uint blurOpacity = (uint)(opacity * 255);

      var accent = new AccentPolicy();
      accent.AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND;
      accent.GradientColor = (blurOpacity << 24) | (blurBackgroundColor & 0xFFFFFF);

      var accentStructSize = Marshal.SizeOf(accent);

      var accentPtr = Marshal.AllocHGlobal(accentStructSize);
      Marshal.StructureToPtr(accent, accentPtr, false);

      var data = new WindowCompositionAttributeData();
      data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY;
      data.SizeOfData = accentStructSize;
      data.Data = accentPtr;
      
      SetWindowCompositionAttribute(hwnd, ref data);

      Marshal.FreeHGlobal(accentPtr);
    }
  }

	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
	public struct DISPLAY_DEVICE 
	{
    [MarshalAs(UnmanagedType.U4)]
    public int cb;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)]
    public string DeviceName;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
    public string DeviceString;
    [MarshalAs(UnmanagedType.U4)]
    public int StateFlags;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
    public string DeviceID;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
    public string DeviceKey;
	}

  public class Displays
  {
    [DllImport("user32.dll")]
		public static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);
		public static string GetDeviceID(string deviceName)
		{     
      DISPLAY_DEVICE outVar = new DISPLAY_DEVICE();
      outVar.cb = (short)Marshal.SizeOf(outVar);
      if (EnumDisplayDevices(deviceName, 0, ref outVar, 1U)) {
        return outVar.DeviceID;
      } else {
        return null;
      }
		}
  }
}
"@

Add-Type -TypeDefinition $source -ReferencedAssemblies System.Windows.Forms, presentationframework

# Query display information:
#  Index: screen number (as seen in display settings dialog)
#  Bounds: position of display in virtual desktop (X, Y, Width, Height)
#  PhysicalSizeInches: screen size, for example 23.9
$screens = @()
foreach ($screen in [System.Windows.Forms.Screen]::AllScreens) {
  # DeviceName = "\\.\DISPLAY1"
  $data = @{
    Index = [int]$screen.DeviceName.replace("\\.\DISPLAY", "")
    X = $screen.Bounds.X
    Y = $screen.Bounds.Y
    Width = $screen.Bounds.Width
    Height = $screen.Bounds.Height
  }
  # deviceId = \\?\DISPLAY#DELA0C5#5&18cf046e&0&UID260#{e6f07b5f-ee97-4a90-b076-33f57bf4eaa7}
  $deviceId = [Custom.Displays]::GetDeviceID($screen.DeviceName)
  # match on model number (DELA0C5), since we only care about getting the physical size
  if ($deviceId -match ".+?#(.+?)#.+") {
    $model = $matches[1]
    # InstanceName = "DISPLAY\DELA0C5\5&18cf046e&0&UID260_0"
    $params = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBasicDisplayParams `
      | Where-Object { $_.InstanceName -match $model } `
      | Select-Object -First 1
    if ($params) {
      $data.PhysicalSizeInches = `
        [System.Math]::Sqrt(
            [System.Math]::Pow($params.MaxHorizontalImageSize, 2) + 
            [System.Math]::Pow($params.MaxVerticalImageSize, 2)
        ) / 2.54
    }
  }
  if (!$data.PhysicalSizeInches) {
    $data.PhysicalSizeInches = 0
  }
  $screens += $data
}

Log "Found $($screens.Count) screen$(if ($screens.Count -gt 1) { 's' })"

if ($videoMonitor -eq "largest") {
  # Find the largest screen, preferring left-most screens if equal sizes
  $browserScreen = $screens `
    | Sort-Object -Property @{Expression = "PhysicalSizeInches"; Descending = $True}, @{Expression = "X"; Ascending = $True} `
    | Select-Object -First 1
} else {
  # Find screen by number
  $browserScreen = $screens | Where-Object { $_.Index -eq $videoMonitor }
  if (!$browserScreen) {
    [void][System.Windows.MessageBox]::Show(
      "Monitor $videoMonitor not found.",
      "PSExercise: Configuration issue",
      "OK", "Error")
    exit
  }
}
Log "Using screen $($browserScreen.Index) ($([Math]::Round($browserScreen.PhysicalSizeInches))`")"

$otherScreens = @($screens | Where-Object { $_.Index -ne $browserScreen.Index })

# Open browser in kiosk and inprivate mode and auto-play video
# --user-data-dir is used to force launching a separate process,
# otherwise kiosk mode will not work if the browser was open already
# See https://superuser.com/a/1281222
$url = "https://www.youtube.com/embed/${videoId}?start=${videoStart}&end=${videoEnd}&autoplay=1"
$appPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$windowX = $browserScreen.X
$windowY = $browserScreen.Y
$windowWidth = $browserScreen.Width
$windowHeight = $browserScreen.Height
$appArgs = "--user-data-dir=`"$dataDir\Edge\User Data`" --inprivate --kiosk --window-position=$windowX,$windowY --window-size=$windowWidth,$windowHeight $url"
Log "Starting browser: $appPath $appArgs"
$process = Start-Process $appPath -ArgumentList $appArgs -PassThru

# Show color/image overlay on other screens to avoid distraction
$overlays = @()
# Note: This uses old-style Windows Forms which are much more performant than WPF Windows
#       With WPF the browser would stutter occasionally on my system
#       The only caveat with Forms is that background images take a second or so to render
#       and they typically render twice due to Zoom mode
if ($otherScreens -and $otherMonitorsOverlay -ne "none") {
  Add-Type -AssemblyName System.Drawing
  
  if ($otherMonitorsOverlay -eq "glass") {
    Log "Using glass overlay for other screens"
    $bgColor = "#000000"
  } elseif ($otherMonitorsOverlay.StartsWith("#")) {
    Log "Using single color overlay for other screens: $otherMonitorsOverlay"
    $bgColor = $otherMonitorsOverlay
  } else {
    Log "Using image overlay for other screens: $otherMonitorsOverlay"
    $bgColor = "#000000"
    $image = $null
    if ($otherMonitorsOverlay.StartsWith("http")) {
      try {
        Log "Downloading image..."
        $request = [System.Net.WebRequest]::create($otherMonitorsOverlay)
        # Note that the DNS resolution timeout cannot be set and may take up to 15s
        $request.ReadWriteTimeout = 3000
        # Note: getResponse() automatically throws an exception for bad HTTP status codes
        $response = $request.getResponse()
        $image = [System.Drawing.Image]::FromStream($response.getResponseStream())
        Log "Downloading image...done"
      } catch {
        [void][System.Windows.MessageBox]::Show(
          "Could not load image from URL: " + $Error[0],
          "PSExercise: Network issue",
          "OK", "Error")
      }
    } else {
      try {
        Log "Loading image from file"
        Set-Location $PSScriptRoot
        $otherMonitorsOverlay = Resolve-Path $otherMonitorsOverlay -ErrorAction Stop
        $image = [System.Drawing.Image]::FromFile($otherMonitorsOverlay)
      } catch {
        [void][System.Windows.MessageBox]::Show(
          "Invalid image path: " + $Error[0],
          "PSExercise: Configuration issue",
          "OK", "Error")
      }
    }
  }
  
  foreach ($screen in $otherScreens) {
    Log "Showing overlay on screen $($screen.Index)"
    $overlay = New-Object Windows.Forms.Form
    $overlay.StartPosition = 'Manual'
    $overlay.Location = New-Object System.Drawing.Point($screen.X, $screen.Y)
    $overlay.TopMost = $true
    $overlay.WindowState = "Maximized"
    $overlay.FormBorderStyle = "None"
    $overlay.BackColor = $bgColor
    if ($image) {
      $overlay.BackgroundImage = $image
      $overlay.BackgroundImageLayout = "Zoom"
    }
    
    $overlay.Show()
  
    if (!$image -and $otherMonitorsOverlay -eq "glass") {
      [Custom.Window]::EnableBlur($overlay.Handle, [int32]$bgColor, 0.5)
    }
  
    $overlays += $overlay
  }
}

if (!$process.HasExited) {
  Log "Waiting for browser window to be ready"
  while (!$process.MainWindowHandle) {
    Start-Sleep -Milliseconds 100
  }

  if ($overlays) {
    # The overlays created above have taken away the focus from the browser window
    # Restore focus to allow easy closing via Alt+F4
    Log "Restoring focus on browser window"
    [void][Custom.Window]::SetForegroundWindow($process.MainWindowHandle)
  }
} else {
  Log "Browser has exited"
}

if ($videoEnd) {
  # Wait for video to finish playing
  $loadTime = 3
  $videoDuration = $videoEnd - $videoStart + $loadTime
  Log "Video has end time, waiting $videoDuration seconds until automatically closing browser"
  $startTime = $(get-date)
  $lastMessageAtElapsed = 0
  do {
    if ($overlays) {
      # Avoid overlays from showing a Wait cursor when hovering over them
      [System.Windows.Forms.Application]::DoEvents()
    }
    Start-Sleep -Seconds 1
    $elapsedTime = ($(get-date) - $startTime).TotalSeconds
    $remainingTime = $videoDuration - $elapsedTime
    if ($elapsedTime - $lastMessageAtElapsed -gt 10) {
      Log "Remaining: $remainingTime seconds, elapsed: $elapsedTime seconds"
      $lastMessageAtElapsed = $elapsedTime
    }
    if ($process.HasExited) {
      Log "Browser was closed manually"
      break
    }
  } while ($remainingTime -gt 0)
} else {
  # Wait until user presses Alt+F4
  Log "Video has no end time, waiting until browser is closed manually"
  $process.WaitForExit()
}

# Remove overlays from other screens
foreach ($overlay in $overlays) {
  Log "Closing overlay"
  $overlay.Close()
}

# Close browser again
if (!$process.HasExited) {
  Log "Closing browser automatically"
  $process.Kill()
}

# See start of file
} finally {
  Log "Stopping logging to file"
  Stop-Transcript | out-null
}
