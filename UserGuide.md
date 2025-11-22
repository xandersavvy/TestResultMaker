# Test Evidence Helper – User Guide

Welcome to **Test Evidence Helper**, the fastest and most painless way for Manual QAs to capture screenshots, write notes, organise test steps, and export everything into a polished Word document.

This guide will walk you through how to use the app end‑to‑end.

---

## 1. What the App Does

**Test Evidence Helper** helps you:

* Capture full/partial screenshots instantly using a hotkey.
* Add notes for each screenshot.
* Automatically organise data as “Test Steps”.
* Maintain a running list of all evidence.
* Export everything into a clean Microsoft Word report.

Perfect for manual testers tired of copy‑paste gymnastics.

---

## 2. Starting the Application

1. Double‑click the application executable or run `python main.py` if using source.
2. The main UI will appear, showing:

   * **Step List Panel** on left
   * **Preview Panel** on right
   * **Buttons** at the top for major actions

---

## 3. Setting Up Your Run

Before you begin capturing evidence:

1. Click **New Session**.
2. Choose a folder where the session data will be saved.
3. The tool will create a session directory and store all screenshots + metadata inside it.

This keeps each test run neatly isolated.

---

## 4. Capturing Screenshots

### **Hotkey Capture (Recommended)**

* Press **Ctrl + Alt + S** (or the hotkey you configured).
* A screenshot is instantly captured.
* The screenshot automatically appears as a new step in the left panel.

The UI does *not block* your work during capture.

### **Manual Capture**

If you prefer clicking:

* Hit the **Capture Screenshot** button.

---

## 5. Adding Notes to a Step

Each step contains:

* Screenshot
* Timestamp
* Notes text box

You can:

* Click on a step in the left panel.
* Type your notes in the "Notes" field.
* Notes are saved automatically.

Examples:

* “Clicked Login button; error message appeared.”
* “Validated dashboard loads within 2 seconds.”

---

## 6. Editing Steps

For each step you can:

* **Rename the step title** (right-click → rename).
* **Reorder steps** (drag and drop).
* **Delete a step** if captured accidentally.

---

## 7. Exporting to Word

When you're done:

1. Click **Export to Word**.
2. Choose a location.
3. The tool generates a `.docx` file with:

   * Step number
   * Screenshot
   * Notes
   * Timestamp

The output is clean, structured, and ready to submit.

---

## 8. Session Auto-Save

Everything in a session (screenshots, notes, structure) is saved inside the session folder.
You can:

* Close the app anytime.
* Reopen and load the same session later.

---

## 9. Clipboard Helper

The tool supports reading text from clipboard when adding notes.
Useful when copying error messages from UI.

---

## 10. Excel Export (If Enabled)

The app can export steps into an Excel file containing:

* Step ID
* Notes
* Screenshot paths
* Timestamps

This is optional.

---

## 11. Troubleshooting

**Screenshots not captured?**

* Ensure no other software is using the same hotkey.

**Export fails?**

* Check if Word file is already open.
* Make sure session folder isn't read‑only.

**App freezes?**

* Heavy screenshots? Just restart; auto-save prevents data loss.

---

## 12. Best Practices

* Start a fresh session for each test case.
* Add short, crisp notes.
* Use hotkey capture—it’s way faster.
* Export after completion to avoid multi-file clutter.

---

## 13. Future Add‑Ons

Suggested roadmap:

* Auto-detect UI changes and capture screenshots.
* Preview Screenshots realtime
* Jira integration directly from the tool

---

## 14. About the Tool

Built for QAs who deserve efficiency.
This tool eliminates hours of repetitive work and produces clean evidence instantly.

Happy testing! 🚀
