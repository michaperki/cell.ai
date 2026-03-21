
This is working much better, but I found the next important UX issue:

There **is** a persistent conversation context/thread now, but the UI does **not display it as a thread**.

What I observed:

* I confirmed the AI remembers prior questions
* For example, I asked:
  **“What was the last question I asked you? And also the question I asked two questions ago?”**
* The AI clearly had access to prior turns and responded based on them
* But the sidebar only shows the **latest** prompt/answer pair
* Each new message seems to overwrite the previous one visually instead of accumulating into a visible conversation history

So the current UX is misleading:

* behavior = **threaded conversation**
* UI = **single-turn panel**

That mismatch is confusing.

What I want changed:

1. **Show an actual visible conversation thread in Ask mode**

* Prior prompts and answers should accumulate in the sidebar
* Do not overwrite the previous exchange when a new one is sent
* Make it read like a real session history

2. **Separate thread history from write previews / execution history**

* In **Ask** mode, the main pane should behave like a lightweight conversation thread
* In write-oriented modes (Fill / Append / Transform), preview/apply remains the core surface
* Don’t force all modes into the same rendering model

3. **Make user vs AI turns visually distinct**

* User question
* AI answer
* Then next user question
* Then next AI answer
  This does not need to look like a consumer chat app, but it should clearly preserve turns.

4. **Keep session visibility explicit**

* The status line is good
* Since there is a real ongoing thread, the visible history now needs to match that reality

My product opinion:

* **Ask mode should feel conversational**
* **Fill / Append / Transform should feel operational**
* Right now Ask mode has the backend behavior of a conversation, but not the frontend presentation

Implementation direction I’d like:

* In Ask mode, render a scrollable thread of prior Q/A turns in the main response area
* Keep newest turn at the bottom
* Preserve the current session until Reset / New session
* Make sure the answer area is not replaced wholesale on each new Ask submission

This is now the main issue I want fixed, because the app already has thread memory, and the UI should stop hiding that.
