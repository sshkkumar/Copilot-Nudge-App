
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;              // ApiException (Graph SDK v5)
using System.Net.Http.Headers;



namespace Copilot_Nudge_App
{
    // An utility class to wrap Graph calls with robust retry.

    public static class GraphRetry
    {
        /// <summary>
        /// Executes a Graph call with retries for throttling (429) and transient 5xx.
        /// Works with Graph SDK v4 (ServiceException) and v5 (ApiException).
        /// Honors Retry-After / x-ms-retry-after-ms when present.
        /// </summary>
        public static async Task<T> ExecuteWithGraphRetryAsync<T>(
            Func<Task<T>> operation,
            CancellationToken ct,
            int maxAttempts = 6,
            TimeSpan? baseDelay = null,
            TimeSpan? maxDelay = null)
        {
            if (operation is null) throw new ArgumentNullException(nameof(operation));

            var initialDelay = baseDelay ?? TimeSpan.FromMilliseconds(500);
            var capDelay = maxDelay ?? TimeSpan.FromSeconds(20);
            var rng = new Random();

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    return await operation().ConfigureAwait(false);
                }
                catch (Exception ex) when (IsTransientGraphError(ex, out var statusCode, out var getHeaderValues))
                {
                    // Last attempt: rethrow
                    if (attempt == maxAttempts) throw;

                    // Prefer server-suggested delay headers; else exponential backoff with jitter
                    var serverDelay = GetServerSuggestedDelay(getHeaderValues);
                    var backoff = serverDelay ?? ComputeExponentialBackoff(initialDelay, attempt, capDelay, rng);
                                        
                    LocalLogger.LogException($"Graph transient {statusCode}: waiting {backoff.TotalMilliseconds:N0} ms (attempt {attempt}/{maxAttempts}).", ex);

                    await Task.Delay(backoff, ct).ConfigureAwait(false);
                    continue;
                }
            }

            throw new InvalidOperationException("Retry loop terminated unexpectedly.");
        }

        /// <summary>
        /// Non-generic overload for operations returning Task.
        /// </summary>
        public static Task ExecuteWithGraphRetryAsync(
            Func<Task> operation,
            CancellationToken ct,
            int maxAttempts = 6,
            TimeSpan? baseDelay = null,
            TimeSpan? maxDelay = null)
        {
            return ExecuteWithGraphRetryAsync(async () =>
            {
                await operation().ConfigureAwait(false);
                return true;
            }, ct, maxAttempts, baseDelay, maxDelay);
        }

        /// <summary>
        /// Determines if error is transient and returns status code + a header accessor delegate.
        /// </summary>
        private static bool IsTransientGraphError(
            Exception ex,
            out int statusCode,
            out Func<string, IEnumerable<string>?> getHeaderValues)
        {
            statusCode = 0;
            getHeaderValues = _ => null;

            switch (ex)
            {
                // Graph SDK v4
                case ServiceException svc:
                    // Prefer ResponseStatusCode; some builds don't expose StatusCode
                    if (svc.ResponseStatusCode != 0)
                        statusCode = (int)svc.ResponseStatusCode;
                    

                    var httpHeaders = svc.ResponseHeaders as HttpResponseHeaders;
                    getHeaderValues = (name) =>
                    {
                        if (httpHeaders == null) return null;
                        return httpHeaders.TryGetValues(name, out var vals) ? vals : null;
                    };
                    break;

                // Graph SDK v5 (Kiota)
                case ApiException api:
                    statusCode = api.ResponseStatusCode;
                    var dictHeaders = api.ResponseHeaders; // IDictionary<string, IEnumerable<string>>
                    getHeaderValues = (name) =>
                    {
                        if (dictHeaders != null && dictHeaders.TryGetValue(name, out var vals)) return vals;
                        return null;
                    };
                    break;

                

                default:
                    // Not a Graph exception: no retry
                    return false;
            }

            // Transient set: 429 throttling, 500/502/503/504 service/gateway errors
            return statusCode == 429 || statusCode == 500 || statusCode == 502 || statusCode == 503 || statusCode == 504;
        }

        /// <summary>
        /// Parses server-provided delay headers using the accessor delegate.
        /// </summary>
        private static TimeSpan? GetServerSuggestedDelay(Func<string, IEnumerable<string>?> getHeaderValues)
        {
            // Retry-After: seconds
            var ra = getHeaderValues("Retry-After")?.FirstOrDefault();
            if (int.TryParse(ra, out var seconds) && seconds > 0)
                return TimeSpan.FromSeconds(seconds);

            // x-ms-retry-after-ms: milliseconds
            var raMs = getHeaderValues("x-ms-retry-after-ms")?.FirstOrDefault();
            if (int.TryParse(raMs, out var ms) && ms > 0)
                return TimeSpan.FromMilliseconds(ms);

            return null;
        }

        /// <summary>
        /// Exponential backoff capped at maxDelay + small jitter.
        /// </summary>
        private static TimeSpan ComputeExponentialBackoff(
            TimeSpan baseDelay, int attempt, TimeSpan capDelay, Random rng)
        {
            var backoffMs = Math.Min(baseDelay.TotalMilliseconds * Math.Pow(2, attempt - 1),
                                     capDelay.TotalMilliseconds);
            var jitterMs = rng.Next(100, 251);
            return TimeSpan.FromMilliseconds(backoffMs + jitterMs);
        }
    }





}
