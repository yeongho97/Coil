using System;
using System.Runtime.InteropServices;

public class MultimediaTimer : IDisposable
{
    private int _timerId;
    private delegate void TimerEventDelegate(uint id, uint msg, UIntPtr user, UIntPtr param1, UIntPtr param2);
    private TimerEventDelegate _callback;

    public event EventHandler Tick;

    public int Period { get; set; } = 10;         // 주기(ms)
    public int Resolution { get; set; } = 1;       // 해상도(ms)

    [DllImport("winmm.dll", SetLastError = true)]
    private static extern int timeSetEvent(uint delay, uint resolution, TimerEventDelegate callback, UIntPtr user, uint mode);

    [DllImport("winmm.dll", SetLastError = true)]
    private static extern int timeKillEvent(int id);

    [DllImport("winmm.dll")]
    private static extern uint timeBeginPeriod(uint uMilliseconds);

    [DllImport("winmm.dll")]
    private static extern uint timeEndPeriod(uint uMilliseconds);

    public void Start()
    {
        Stop(); // 중복 방지

        timeBeginPeriod((uint)Resolution);
        _callback = new TimerEventDelegate(TimerCallback);
        _timerId = timeSetEvent((uint)Period, (uint)Resolution, _callback, UIntPtr.Zero, 1); // TIME_PERIODIC = 1
    }

    public void Stop()
    {
        if (_timerId != 0)
        {
            timeKillEvent(_timerId);
            timeEndPeriod((uint)Resolution);
            _timerId = 0;
        }
    }

    private void TimerCallback(uint id, uint msg, UIntPtr user, UIntPtr param1, UIntPtr param2)
    {
        Tick?.Invoke(this, EventArgs.Empty);
    }

    public void Dispose()
    {
        Stop();
    }
}
