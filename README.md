# Blade Runner

A pattern for running long-running tasks in Google Apps Script beyond the 5-minute execution limit.

## The Problem

Google Apps Script enforces a ~6-minute execution limit per run. Any function that exceeds it is killed — making long-running tasks (bulk data processing, external API polling, etc.) seemingly impossible.

## The Solution

Split the task into resumable chunks using **Script Properties** as persistent state, and use `google.script.run` from a modal dialog to automatically restart the function whenever it gets killed.

```
Client (modal dialog)
  └─ calls heavyTask()
       └─ on failure (timeout) → calls heavyTask() again  ← the key trick
```

Since `google.script.run.withFailureHandler()` fires on execution timeout, the client silently relaunches the task. The server-side function reads its last saved progress from Script Properties and picks up where it left off.

## How It Works

1. **Open the dialog** via the "Oleksii 🛸" menu → "Open Launcher"
2. **Click Run** — calls `heavyTask()` on the server
3. `heavyTask()` initialises a `PROGRESS` counter in Script Properties on first run
4. Each iteration sleeps 1 minute and increments the counter
5. If the execution limit kills the function, `withFailureHandler` fires and restarts it
6. The task completes after 7 iterations regardless of how many restarts were needed

## Key APIs Used

- `PropertiesService` — persists progress state across executions
- `HtmlService` — serves the modal dialog
- `google.script.run` with `withFailureHandler` — client-side restart on timeout
