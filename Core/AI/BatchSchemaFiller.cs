using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp.Core.AI
{
    /// <summary>
    /// Orchestrates schema-fill operations across multiple batches for large datasets.
    /// Stage 2 of the structured batch fill feature.
    /// </summary>
    public sealed class BatchSchemaFiller
    {
        private readonly IChatPlanner _planner;
        private readonly int _batchSize;
        private readonly int _maxConcurrent;

        public BatchSchemaFiller(IChatPlanner planner, int batchSize = 30, int maxConcurrent = 3)
        {
            _planner = planner;
            _batchSize = Math.Max(5, Math.Min(50, batchSize));
            _maxConcurrent = Math.Max(1, Math.Min(5, maxConcurrent));
        }

        public sealed class BatchResult
        {
            public int BatchIndex { get; set; }
            public int StartRow { get; set; }  // 0-based
            public int RowCount { get; set; }
            public AIPlan? Plan { get; set; }
            public string? Error { get; set; }
            public bool Success => Plan != null && string.IsNullOrEmpty(Error);
        }

        public sealed class FillProgress
        {
            public int TotalBatches { get; set; }
            public int CompletedBatches { get; set; }
            public int TotalRows { get; set; }
            public int CompletedRows { get; set; }
            public int FailedBatches { get; set; }
        }

        /// <summary>
        /// Estimates the number of batches needed for a given row count.
        /// </summary>
        public int EstimateBatchCount(int totalRows)
        {
            return (int)Math.Ceiling((double)totalRows / _batchSize);
        }

        /// <summary>
        /// Splits a large schema fill into batches and executes them with progress reporting.
        /// </summary>
        /// <param name="baseContext">The AI context template (headers, schema, policy). StartRow and Rows will be adjusted per batch.</param>
        /// <param name="basePrompt">The fill prompt template.</param>
        /// <param name="inputValues">The input column values, one per row.</param>
        /// <param name="startRow">0-based starting row in the sheet.</param>
        /// <param name="onProgress">Called after each batch completes.</param>
        /// <param name="ct">Cancellation token. Cancelling finishes the current batch but stops further batches.</param>
        /// <returns>Results per batch, in order.</returns>
        public async Task<List<BatchResult>> FillAsync(
            AIContext baseContext,
            string basePrompt,
            string[] inputValues,
            int startRow,
            Action<FillProgress>? onProgress = null,
            CancellationToken ct = default)
        {
            int totalRows = inputValues.Length;
            int batchCount = EstimateBatchCount(totalRows);
            var results = new List<BatchResult>(batchCount);
            var progress = new FillProgress
            {
                TotalBatches = batchCount,
                TotalRows = totalRows
            };

            // Build batch specs
            var batches = new List<(int index, int startOffset, int count)>();
            for (int i = 0; i < batchCount; i++)
            {
                int offset = i * _batchSize;
                int count = Math.Min(_batchSize, totalRows - offset);
                batches.Add((i, offset, count));
            }

            // Execute batches with limited concurrency
            using var semaphore = new SemaphoreSlim(_maxConcurrent);
            var tasks = new List<Task<BatchResult>>();

            foreach (var (index, offset, count) in batches)
            {
                if (ct.IsCancellationRequested) break;

                await semaphore.WaitAsync(ct).ConfigureAwait(false);

                var batchTask = Task.Run(async () =>
                {
                    try
                    {
                        var batchInputs = inputValues.Skip(offset).Take(count).ToArray();
                        var batchCtx = CloneContextForBatch(baseContext, startRow + offset, count, batchInputs);

                        using var batchCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                        batchCts.CancelAfter(TimeSpan.FromSeconds(60)); // per-batch timeout

                        int retries = 0;
                        const int maxRetries = 2;
                        AIPlan? plan = null;
                        string? lastError = null;

                        while (retries <= maxRetries)
                        {
                            try
                            {
                                plan = await _planner.PlanAsync(batchCtx, basePrompt, batchCts.Token).ConfigureAwait(false);
                                if (plan.Commands.Count > 0) break;
                                lastError = "Empty plan returned";
                            }
                            catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                            {
                                lastError = "Batch timed out";
                            }
                            catch (Exception ex)
                            {
                                lastError = ex.Message;
                            }
                            retries++;
                        }

                        return new BatchResult
                        {
                            BatchIndex = index,
                            StartRow = startRow + offset,
                            RowCount = count,
                            Plan = plan?.Commands.Count > 0 ? plan : null,
                            Error = plan?.Commands.Count > 0 ? null : lastError
                        };
                    }
                    catch (Exception ex)
                    {
                        return new BatchResult
                        {
                            BatchIndex = index,
                            StartRow = startRow + offset,
                            RowCount = count,
                            Error = ex.Message
                        };
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }, ct);

                tasks.Add(batchTask);
            }

            // Collect results in order
            foreach (var task in tasks)
            {
                try
                {
                    var result = await task.ConfigureAwait(false);
                    results.Add(result);
                    progress.CompletedBatches++;
                    progress.CompletedRows += result.RowCount;
                    if (!result.Success) progress.FailedBatches++;
                    try { onProgress?.Invoke(progress); } catch { }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }

            return results.OrderBy(r => r.BatchIndex).ToList();
        }

        private static AIContext CloneContextForBatch(AIContext template, int batchStartRow, int batchRowCount, string[] batchInputs)
        {
            var ctx = new AIContext
            {
                SheetName = template.SheetName,
                StartRow = batchStartRow,
                StartCol = template.StartCol,
                Rows = batchRowCount,
                Cols = template.Cols,
                Title = template.Title,
                Workbook = template.Workbook,
                AllowedCommands = template.AllowedCommands ?? new[] { "set_values" },
                WritePolicy = template.WritePolicy,
                Schema = template.Schema
            };

            // Build NearbyValues with header row + batch inputs
            if (template.NearbyValues != null && template.NearbyValues.Length > 0)
            {
                var nearby = new List<string[]>();
                nearby.Add(template.NearbyValues[0]); // header row
                foreach (var input in batchInputs)
                {
                    var row = new string[template.NearbyValues[0].Length];
                    row[0] = input;
                    nearby.Add(row);
                }
                ctx.NearbyValues = nearby.ToArray();
            }

            return ctx;
        }
    }
}
