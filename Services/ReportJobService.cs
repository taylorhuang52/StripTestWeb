using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using StripTestWeb.Models;

namespace StripTestWeb.Services
{
    // ── Message sent to client via SSE ────────────────────────────
    public record JobMessage(
        string Type,          // "progress" | "done" | "error" | "status"
        string? FileName  = null,
        string? Status    = null,   // "processing" | "done" | "error"
        int?    Index     = null,
        int?    Total     = null,
        object? Row       = null,   // serialised LogResult summary
        string? Error     = null,
        byte[]? XlsxBytes = null
    );

    // ── Per-job state ──────────────────────────────────────────────
    public class ReportJob
    {
        public string Id { get; } = Guid.NewGuid().ToString("N");
        public List<(string name, byte[] data, string ext)> Files { get; set; } = new();

        // SSE message queue
        public ConcurrentQueue<string> Messages { get; } = new();
        public SemaphoreSlim           Signal   { get; } = new(0, int.MaxValue);

        // Pause / Stop controls
        TaskCompletionSource<bool> _pauseTcs =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        bool _paused  = false;
        bool _stopped = false;

        public bool IsRunning  { get; private set; } = false;
        public bool IsFinished { get; private set; } = false;

        public void Start()
        {
            IsRunning = true;
            _pauseTcs.TrySetResult(true);
            _ = RunAsync();
        }

        public void Pause()
        {
            if (_paused) return;
            _paused  = true;
            _pauseTcs = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            Enqueue("status", null, "⏸ Paused");
        }

        public void Resume()
        {
            if (!_paused) return;
            _paused = false;
            _pauseTcs.TrySetResult(true);
            Enqueue("status", null, "▶ Resuming…");
        }

        public void Stop()
        {
            _stopped = true;
            _pauseTcs.TrySetResult(true);
            Enqueue("status", null, "⏹ Stopped by user.");
        }

        // ── Core processing loop ───────────────────────────────────
        async Task RunAsync()
        {
            var results = new List<LogResult>();
            int total   = Files.Count;

            for (int i = 0; i < total; i++)
            {
                if (_stopped) break;

                // Async-friendly pause
                await _pauseTcs.Task.ConfigureAwait(false);
                if (_stopped) break;

                var (name, data, ext) = Files[i];
                Enqueue("progress", name, "processing", i, total);

                LogResult result;
                try
                {
                    await using var ms = new MemoryStream(data);
                    result = await Task.Run(() => LogParser.ParseStream(ms, name, ext));
                }
                catch (Exception ex)
                {
                    result = new LogResult { FileName = name, Error = ex.Message };
                }

                results.Add(result);

                // Send row summary to client
                var rowSummary = new
                {
                    index    = i,
                    fileName = name,
                    status   = result.Error == null ? "done" : "error",
                    error    = result.Error,
                    input    = result.Input,
                    output   = result.Output,
                    yield    = result.Yield,
                    vthLt15  = result.VthLt15,
                    vthGt6   = result.VthGt6,
                };
                Enqueue("row", name, result.Error == null ? "done" : "error",
                        i, total, rowSummary);
            }

            // Generate Excel if not stopped
            if (!_stopped && results.Any(r => r.Error == null))
            {
                Enqueue("status", null, "Writing report…");
                try
                {
                    var bytes    = await Task.Run(() =>
                        ReportGenerator.Generate(results.Where(r => r.Error == null), DateTime.Today));
                    string b64   = Convert.ToBase64String(bytes);
                    string fname = $"Strip_Test_Report_{DateTime.Today:yyyyMMdd}.xlsx";

                    // Send download payload
                    var msg = JsonSerializer.Serialize(new
                    {
                        type     = "done",
                        fileName = fname,
                        data     = b64,
                        count    = results.Count(r => r.Error == null)
                    });
                    Enqueue(msg);
                }
                catch (Exception ex)
                {
                    Enqueue("error", null, ex.Message);
                }
            }
            else if (_stopped)
            {
                Enqueue("status", null, "⏹ Stopped. No report saved.");
            }

            IsFinished = true;
            IsRunning  = false;
            Enqueue("end");
        }

        // ── SSE helpers ───────────────────────────────────────────
        void Enqueue(string rawJson) { Messages.Enqueue(rawJson); Signal.Release(); }

        void Enqueue(string type, string? fileName = null, string? status = null,
                     int? index = null, int? total = null, object? row = null)
        {
            var obj = new Dictionary<string, object?>();
            obj["type"] = type;
            if (fileName != null) obj["fileName"] = fileName;
            if (status   != null) obj["status"]   = status;
            if (index    != null) obj["index"]     = index;
            if (total    != null) obj["total"]     = total;
            if (row      != null) obj["row"]       = row;
            Enqueue(JsonSerializer.Serialize(obj));
        }
    }

    // ── Job registry (singleton) ───────────────────────────────────
    public class ReportJobService
    {
        readonly ConcurrentDictionary<string, ReportJob> _jobs = new();

        public ReportJob Create()
        {
            var job = new ReportJob();
            _jobs[job.Id] = job;
            return job;
        }

        public ReportJob? Get(string id) => _jobs.TryGetValue(id, out var j) ? j : null;

        public void Remove(string id) => _jobs.TryRemove(id, out _);
    }
}
