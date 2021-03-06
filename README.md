# PSExercise

A tool for Windows that shows fullscreen YouTube videos on a configurable time schedule.

It's called PSExercise because I use it to interrupt my desk life with exercise videos and it is based on [PowerShell](https://docs.microsoft.com/en-us/powershell/scripting/overview).

| ![Screenshot](screenshot.jpg) |
|:--:| 
| *Two monitors using the blur option* |

## Features

- Configuration through simple PowerShell file (see [`config.sample.ps1`](config.sample.ps1))
- Add videos by copy-pasting "Share" link from YouTube
- Automatically closes video if end time is given in addition to "Share" link
- Select screen on which to show video (default: the largest)
- Blur out extra screens or show an image or color
- Option to ask for confirmation before starting video
- Delayed confirmation popup when fullscreen app is in foreground
- Set start/end time of day and interval for showing videos
- Set weekdays for showing videos
- Uses Microsoft Edge in private-browsing mode (with a fresh profile)
- Uses Windows Task Scheduler (no background processes!)

## Getting started

First, clone or download this repository.

**All the following commands have to be run in a PowerShell terminal.**
Try [Windows Terminal](https://aka.ms/terminal) for a modern terminal app.

If you downloaded and extracted this repository from the **ZIP archive**, you need to unblock the scripts first:
```sh
ls *.ps1 | Unblock-File
```

Before running the scripts, you may have to change your execution policy (affects only current terminal session):
```sh
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
```

NOTE: Never run scripts that you haven't inspected yourself first!

### Configuration

Make a copy of `config.sample.ps1` and save it as `config.ps1`.
You can configure all options and your video collection in this file.
Open it with any text or code editor like [VS Code](https://code.visualstudio.com/).
When you're done, continue with the next section.

### Run manually

To test your configuration options (at least the ones unrelated to task scheduling), run the following:
```sh
.\run.ps1
```

**Note:** To stop a video early or if you haven't specified `t2`, press <kbd>Alt</kbd>+<kbd>F4</kbd>.

### Register as scheduled task

```sh
.\register.ps1
```
When opening the [Windows Task Scheduler](https://en.wikipedia.org/wiki/Windows_Task_Scheduler) you can now find the task inside the `letmaik` folder.

Note: Any time you make a change to the task registration options in `config.ps1` or if you move or rename the repository folder, simply re-run the above command.

### Unregister the scheduled task

To stop the automated schedule, you can remove the task again from the Windows Task Scheduler by running:
```sh
.\register.ps1 -remove
```

## FAQ

### How do I stop the video (early)?

Press <kbd>Alt</kbd>+<kbd>F4</kbd>.

### Why is <kbd>Esc</kbd> and <kbd>F11</kbd> not working?

The browser is started in kiosk mode which makes it fullscreen and blocks access to most controls. The idea is that it shouldn't be too easy to skip a video.

### Why is the window closed too early when the video is paused for a while?

When the end time `t2` is given for a video in `config.ps1` then the video duration is automatically computed and the window is closed when the video is supposed to have ended (by keeping a timer running in the background), meaning the tool does not actually check the playback status of YouTube, since it doesn't have easy access to it.

This essentially means that videos shouldn't be paused if `t2` is given, otherwise the window will close too early.
If `t2` is not given, then the window is not closed automatically, instead <kbd>Alt</kbd>+<kbd>F4</kbd> has to be pressed.

### I use an image for my extra screens. Why does it take 1-2s to display?

This is a known technical limitation and happens especially for big images.
If you're curious, check out the comments in `run.ps1` starting around `$overlays = @()`.

### Why is this tool written in PowerShell?

Mostly because it doesn't require any extra setup/installation steps.
The tool also interacts a fair bit with .NET and Win32 APIs, both of which is easy with PowerShell.

### Why is this tool relying on Microsoft Edge?

Edge is available on all up-to-date Windows 10 PCs and this means the tool will work out-of-the-box.

### What is stored in the hidden `.data` folder?

`history.txt`: Playback history of the current day to avoid showing the same video twice

`lastrun.txt`: Terminal output from the last run to aid in diagnosing issues

`Edge/`: A fresh browser profile to force starting Edge in a separate process and avoid interfering with any other open browser windows
