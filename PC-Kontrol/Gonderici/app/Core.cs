using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PcCheck {
    public static class Json {
        public static string Quote(string value) {
            if (value == null) return "null";
            StringBuilder b = new StringBuilder("\"");
            foreach (char c in value) {
                switch (c) {
                    case '"': b.Append("\\\""); break;
                    case '\\': b.Append("\\\\"); break;
                    case '\n': b.Append("\\n"); break;
                    case '\r': b.Append("\\r"); break;
                    case '\t': b.Append("\\t"); break;
                    default: if (c < 32) b.Append("\\u" + ((int)c).ToString("x4")); else b.Append(c); break;
                }
            }
            return b.Append('"').ToString();
        }
    }

    public sealed class CpuRun {
        private readonly object sync = new object();
        private string status = "pending", error = null;
        private long iterations, elapsed;
        private int workers;
        private volatile bool cancelled;
        public bool Active { get { lock (sync) return status == "running"; } }
        public void Cancel() { cancelled = true; }
        public bool Start() {
            lock (sync) {
                if (status == "running") return false;
                status = "running"; error = null; iterations = 0; elapsed = 0; cancelled = false;
                workers = Math.Min(2, Math.Max(1, Environment.ProcessorCount));
            }
            ThreadPool.QueueUserWorkItem(delegate(object unused) { Run(); });
            return true;
        }
        private void Run() {
            Stopwatch total = Stopwatch.StartNew();
            try {
                Parallel.For(0, workers, new ParallelOptions { MaxDegreeOfParallelism = workers }, delegate(int worker) {
                    byte[] input = Encoding.ASCII.GetBytes("abc");
                    byte[] expected = {0xba,0x78,0x16,0xbf,0x8f,0x01,0xcf,0xea,0x41,0x41,0x40,0xde,0x5d,0xae,0x22,0x23,0xb0,0x03,0x61,0xa3,0x96,0x17,0x7a,0x9c,0xb4,0x10,0xff,0x61,0xf2,0x00,0x15,0xad};
                    long count = 0;
                    using (SHA256 sha = SHA256.Create()) {
                        while (!cancelled && total.ElapsedMilliseconds < 3000) {
                            long sum = 0;
                            for (int i = 1; i <= 1000; i++) sum += (long)i * i;
                            if (sum != 333833500L) throw new InvalidOperationException("Integer calculation mismatch");
                            byte[] actual = sha.ComputeHash(input);
                            for (int i = 0; i < expected.Length; i++)
                                if (actual[i] != expected[i]) throw new InvalidOperationException("SHA-256 mismatch");
                            count++;
                        }
                    }
                    Interlocked.Add(ref iterations, count);
                });
                lock (sync) { elapsed = total.ElapsedMilliseconds; status = cancelled ? "cancelled" : (iterations > 0 ? "pass" : "unknown"); }
            } catch (Exception ex) {
                lock (sync) { elapsed = total.ElapsedMilliseconds; status = "fail"; error = ex.GetBaseException().Message; }
            }
        }
        public string ToJson() {
            lock (sync) {
                return "{\"status\":" + Json.Quote(status) + ",\"durationMs\":" + elapsed +
                    ",\"iterations\":" + Interlocked.Read(ref iterations) + ",\"workers\":" + workers +
                    ",\"error\":" + Json.Quote(error) + ",\"scope\":\"3-second integer and SHA-256 check; at most 2 workers; no temperature or full stability assessment\"}";
            }
        }
    }

}
