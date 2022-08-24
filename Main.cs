using System.Diagnostics;
using System.ServiceProcess;

namespace OVRSwitch;

public class Main : ApplicationContext
{
    private const string TaskPath = "OVRSwitch";
    private readonly ToolStripMenuItem toggle;
    private readonly ToolStripMenuItem auto;
    private readonly ToolStripMenuItem startup;
    private readonly NotifyIcon notifyIcon;
    private readonly Microsoft.Win32.TaskScheduler.TaskService taskService = new();
    private readonly ServiceController serviceController = new("OVRService");
    private bool running = true;
    private bool autoMode = Properties.App.Default.AutoMode;
    public Main()
    {

        toggle = new("Toggle", null, async (s, e) => await Toggle());
        auto = new("AutoMode", null, (s, e) =>
        {
            Properties.App.Default.AutoMode = auto!.Checked = autoMode = !autoMode;
            Properties.App.Default.Save();
        })
        {
            Checked = autoMode
        };
        var task = taskService.GetTask(TaskPath);
        var exePath = Environment.ProcessPath;
        startup = new("Run at Startup", null, (s, e) =>
        {
            if (startup!.Checked)
            {
                taskService.RootFolder.DeleteTask(TaskPath, false);
                startup.Checked = false;
            }
            else
            {
                var task = taskService.NewTask();
                task.Triggers.Add(new Microsoft.Win32.TaskScheduler.LogonTrigger
                {
                    Delay = TimeSpan.FromMinutes(1),
                    Enabled = true
                });
                task.Actions.Add(new Microsoft.Win32.TaskScheduler.ExecAction
                {
                    Path = "cmd.exe",
                    Arguments = $"/C start /B {Path.GetFileName(exePath)}",
                    WorkingDirectory = Path.GetDirectoryName(exePath)
                });
                task.Principal.RunLevel = Microsoft.Win32.TaskScheduler.TaskRunLevel.Highest;
                taskService.RootFolder.RegisterTaskDefinition(TaskPath, task);
                startup.Checked = true;

            }
        })
        {
            Checked = task?.Enabled == true
        };
        notifyIcon = new NotifyIcon
        {
            Icon = new Icon("icon.ico"),
            ContextMenuStrip = new ContextMenuStrip
            {
                Items =
                {
                    toggle,
                    auto,
                    startup,
                    new ToolStripMenuItem("Exit", null, (s, e) =>
                    {
                        Application.Exit();
                    })
                }
            },
            Visible = true
        };
        notifyIcon.ContextMenuStrip.Opening += (s, e) =>
        {
            UpdateMenu();
        };
        CreateCheckTask();
    }

    private void CreateCheckTask()
    {
        var sc = SynchronizationContext.Current;
        Task.Factory.StartNew(async () =>
        {
            while (running)
            {
                if (autoMode && !Process.GetProcessesByName("OculusClient").Any() && !Process.GetProcessesByName("OVRServiceLauncher").Any())
                {
                    serviceController.Refresh();
                    if (serviceController.Status == ServiceControllerStatus.Running)
                    {
                        sc?.Send(async (state) => await Toggle(false), null);
                    }
                }
                await Task.Delay(TimeSpan.FromMinutes(1));
            }
        }, TaskCreationOptions.LongRunning);
    }

    private void UpdateMenu()
    {
        if (autoMode)
        {
            toggle.Enabled = false;
            toggle.Text = "Auto Mode Running";
        }
        else
        {
            serviceController.Refresh();
            switch (serviceController.Status)
            {
                case ServiceControllerStatus.Running:
                    toggle.Text = "Stop OVRService";
                    toggle.Enabled = true;
                    break;
                case ServiceControllerStatus.Stopped:
                    toggle.Text = "Start OVRService";
                    toggle.Enabled = true;
                    break;
                default:
                    toggle.Text = "Buzy..";
                    toggle.Enabled = false;
                    break;
            }
        }
    }

    protected override void Dispose(bool disposing)
    {
        taskService.Dispose();
        running = false;
        notifyIcon.Visible = false;
        notifyIcon.Dispose();
        serviceController.Dispose();
        base.Dispose(disposing);
    }

    private readonly SemaphoreSlim semaphoreSlim = new(1);
    private async Task Toggle(bool? force = null)
    {
        try
        {
            await semaphoreSlim.WaitAsync();
            toggle.Enabled = false;

            if (serviceController.Status == ServiceControllerStatus.Running && !force.HasValue || force == false)
            {
                await Task.Run(() => serviceController.Stop());
            }
            else if (serviceController.Status == ServiceControllerStatus.Stopped && !force.HasValue || force == true)
            {
                await Task.Run(() => serviceController.Start());
            }
            UpdateMenu();
        }
        finally
        {
            toggle.Enabled = true;
            semaphoreSlim.Release();

        }
    }
}