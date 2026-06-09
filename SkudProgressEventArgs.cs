namespace Tabel;

public sealed class SkudProgressEventArgs : EventArgs
{
    public string Stage { get; }
    public string Message { get; }
    public int Current { get; }
    public int Total { get; }
    public int Percent { get; }
    public bool IsFinished { get; }

    public SkudProgressEventArgs(
        string stage,
        string message,
        int current,
        int total,
        int percent,
        bool isFinished = false)
    {
        Stage = stage;
        Message = message;
        Current = current;
        Total = total;
        Percent = percent;
        IsFinished = isFinished;
    }
}