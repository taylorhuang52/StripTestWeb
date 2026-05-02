using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using StripTestWeb.Services;

namespace StripTestWeb.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportController : ControllerBase
    {
        readonly ReportJobService _jobs;
        public ReportController(ReportJobService jobs) => _jobs = jobs;

        // ── POST /api/report/upload ────────────────────────────────
        // Accepts multiple XLS/XLSX files, creates a job, returns jobId
        [HttpPost("upload")]
        [RequestSizeLimit(500 * 1024 * 1024)]   // 500 MB
        public async Task<IActionResult> Upload()
        {
            var files = Request.Form.Files;
            if (files == null || files.Count == 0)
                return BadRequest(new { error = "No files uploaded." });

            var job = _jobs.Create();
            foreach (var f in files)
            {
                string ext = Path.GetExtension(f.FileName).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx") continue;
                if (f.FileName.StartsWith("~$")) continue;   // skip Excel lock files

                using var ms = new MemoryStream();
                await f.CopyToAsync(ms);
                job.Files.Add((f.FileName, ms.ToArray(), ext));
            }

            if (job.Files.Count == 0)
                return BadRequest(new { error = "No valid XLS/XLSX files found." });

            job.Start();
            return Ok(new { jobId = job.Id, count = job.Files.Count });
        }

        // ── GET /api/report/stream/{jobId} ────────────────────────
        // Server-Sent Events stream: progress, rows, done/error
        [HttpGet("stream/{jobId}")]
        public async Task Stream(string jobId, CancellationToken ct)
        {
            var job = _jobs.Get(jobId);
            if (job == null) { Response.StatusCode = 404; return; }

            Response.Headers["Content-Type"]  = "text/event-stream";
            Response.Headers["Cache-Control"] = "no-cache";
            Response.Headers["X-Accel-Buffering"] = "no";

            await Response.Body.FlushAsync(ct);

            while (!ct.IsCancellationRequested)
            {
                // Wait for a message or timeout (keepalive)
                bool got = await job.Signal.WaitAsync(15_000, ct).ConfigureAwait(false);

                if (!got)
                {
                    // Keepalive ping
                    await WriteEvent(": ping\n\n", ct);
                    continue;
                }

                // Drain all queued messages
                while (job.Messages.TryDequeue(out var msg))
                {
                    await WriteEvent($"data: {msg}\n\n", ct);

                    if (msg.Contains("\"type\":\"end\""))
                    {
                        await Response.Body.FlushAsync(ct);
                        _jobs.Remove(jobId);
                        return;
                    }
                }

                await Response.Body.FlushAsync(ct);
            }
        }

        async Task WriteEvent(string text, CancellationToken ct)
        {
            var bytes = Encoding.UTF8.GetBytes(text);
            await Response.Body.WriteAsync(bytes, ct);
        }

        // ── POST /api/report/pause/{jobId} ────────────────────────
        [HttpPost("pause/{jobId}")]
        public IActionResult Pause(string jobId)
        {
            var job = _jobs.Get(jobId);
            if (job == null) return NotFound();
            job.Pause();
            return Ok();
        }

        // ── POST /api/report/resume/{jobId} ───────────────────────
        [HttpPost("resume/{jobId}")]
        public IActionResult Resume(string jobId)
        {
            var job = _jobs.Get(jobId);
            if (job == null) return NotFound();
            job.Resume();
            return Ok();
        }

        // ── POST /api/report/stop/{jobId} ─────────────────────────
        [HttpPost("stop/{jobId}")]
        public IActionResult Stop(string jobId)
        {
            var job = _jobs.Get(jobId);
            if (job == null) return NotFound();
            job.Stop();
            return Ok();
        }
    }
}
