
They do still make a lot of requests — just not one request per keystroke in the naive sense.

The usual pattern is:

1. **Wait for a useful pause, not every keypress.**
   VS Code’s inline completion flow is designed around “when the user stopped typing,” and editors also expose a minimum show delay for inline suggestions. So these tools usually debounce a bit and only fire when the cursor state looks stable enough to matter. ([Visual Studio Code][1])

2. **Cancel stale requests immediately.**
   Editor APIs pass a cancellation token specifically because a completion request often becomes obsolete as soon as the user types again. Good autocomplete clients aggressively abort in-flight work so the backend does not keep finishing junk predictions for old cursor states. ([Visual Studio Code][1])

3. **Only trigger on “high-value” moments.**
   They do not treat every character equally. A newline, indentation, `.`, `(`, `,`, `=`, finishing an identifier, or a short pause after a comment is much more predictive than the middle of a fast typing burst. The editor API also distinguishes automatic triggers from explicit user requests, which lets the client be selective. ([Visual Studio Code][1])

4. **Send tiny, structured context snapshots.**
   They usually avoid re-sending the whole project on every request. Instead they send the active prefix/suffix around the cursor plus a compact context pack: current file chunk, nearby symbols, imports, maybe a few relevant open files. Copilot’s docs explicitly say suggestions use current and open files as context, which implies context assembly rather than blind full-repo uploads each time. ([Visual Studio Code][2])

5. **Cache aggressively.**
   If you typed `pri` then `prin` then `print`, a lot of the prompt is unchanged. Systems commonly cache by prefix, cursor region, file hash, AST region, or retrieved-context bundle, so many “new” requests reuse prior work. This part is an inference from how low-latency completion systems are usually built; vendors do not always publish the exact cache keys.

6. **Use cheap local filtering before expensive model calls.**
   A lot of requests die locally: wrong language, tiny file, cursor inside a string where suggestions are low-value, user typing too fast, same request as 80 ms ago, etc. Often there is also local ranking or post-filtering before showing anything. That is an inference, but it is standard for keeping latency and cost down.

7. **Keep models specialized for completion.**
   Inline code completion is usually handled by a smaller/faster completion model rather than the same heavyweight model used for chat. GitHub documents that inline suggestions use a dedicated completion model, and its plans meter “inline suggestions” separately from premium requests, which strongly suggests a separate serving path optimized for this workload. ([GitHub Docs][3])

8. **Optimize serving hard.**
   Vendors shave milliseconds with batching, prefix KV-cache reuse, speculative decoding, GPU-friendly serving, and regional low-latency infra. Supermaven publicly marketed much lower latency than several competitors and a very large context window, and speculative decoding is now a common inference optimization for fast generation. ([supermaven.com][4])

So the real answer is:

**They do not avoid lots of “events.” They avoid lots of *expensive completed inferences*.**
They debounce, cancel, cache, trim context, gate triggers, and use a fast completion-specific backend.

If you were building one yourself, the minimal sane strategy would be:

* debounce around 100–250 ms
* cancel on every new edit
* only trigger on pauses + syntax boundaries
* send current file window + nearby symbols + a few open-file snippets
* cache by `(file hash, cursor region, prefix)`
* drop responses that return for an outdated document version

That gets you surprisingly far.

[1]: https://code.visualstudio.com/api/references/vscode-api?utm_source=chatgpt.com "VS Code API | Visual Studio Code Extension API"
[2]: https://code.visualstudio.com/docs/copilot/ai-powered-suggestions?utm_source=chatgpt.com "Inline suggestions from GitHub Copilot in VS Code"
[3]: https://docs.github.com/en/copilot/concepts/completions/code-suggestions?utm_source=chatgpt.com "GitHub Copilot code suggestions in your IDE"
[4]: https://supermaven.com/blog/introducing-supermaven?utm_source=chatgpt.com "Introducing Supermaven, the first code completion tool with ..."
