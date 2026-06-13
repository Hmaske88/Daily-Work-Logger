# Daily Work Logger — Setup Guide

## Files You Received
| File | Purpose |
|------|---------|
| `Daily_Work_Log.xlsx` | Your Excel work log (with sample row) |
| `daily_work_logger.py` | Python popup script (runs at 6 PM) |
| `DailyWorkLogger_Task.xml` | Windows Task Scheduler config |

---

## Step 1 — Install Python (if not already installed)
1. Download from https://python.org/downloads
2. During install, check ✅ **"Add Python to PATH"**
3. Open CMD and run: `pip install openpyxl`

---

## Step 2 — Set Up Your Files
1. Copy `Daily_Work_Log.xlsx` to a folder, e.g.:
   `C:\Users\YourName\Documents\Daily_Work_Log.xlsx`

2. Copy `daily_work_logger.py` to same folder

3. Open `daily_work_logger.py` in Notepad and **edit line 23**:
   ```
   EXCEL_PATH = r"C:\Users\YourName\Documents\Daily_Work_Log.xlsx"
   ```
   Replace `YourName` with your actual Windows username.

4. You can also change ASSOCIATE name, TASK, SUBTASKS at the top of the script.

---

## Step 3 — Set Up Task Scheduler (Auto-run at 6 PM)

### Option A: Import XML (Easiest)
1. Open **Task Scheduler** (search in Start Menu)
2. Click **Action → Import Task...**
3. Select `DailyWorkLogger_Task.xml`
4. Edit the two paths inside it to match your system:
   - `<Command>` → path to `pythonw.exe`
     (Find it: open CMD, type `where python`, replace `python.exe` with `pythonw.exe`)
   - `<Arguments>` → path to `daily_work_logger.py`
5. Click OK → Done!

### Option B: Manual Setup
1. Open Task Scheduler → **Create Basic Task**
2. Name: `Daily Work Logger`
3. Trigger: **Daily** at **6:00 PM**
4. Action: **Start a program**
   - Program: `pythonw.exe` (full path)
   - Arguments: full path to `daily_work_logger.py`
5. Finish

---

## How It Works Every Day

| Day | What Happens |
|-----|-------------|
| **Mon–Fri** | Popup appears at 6 PM → You type "Other Task" → Auto-fills all columns |
| **Sat–Sun** | Popup appears → Click one button → Logs "Weekend" automatically |

**Auto-filled columns (no typing needed):**
- Sr No → Auto-increments
- Date → Today's date
- Duration → Full Day
- Task → Auditing
- Subtasks → Financial Statement Audit
- Associate Name → Raj Kumar

**You only type:** Other Task (e.g., "Excel", "Report", "Meeting")

---

## Testing
To test without waiting for 6 PM:
1. Double-click `daily_work_logger.py`  
   *(or right-click → Open With → Python)*
2. The popup should appear immediately

---

## Troubleshooting
| Problem | Fix |
|---------|-----|
| Popup doesn't appear | Check Task Scheduler → right-click task → Run |
| "File not found" error | Double-check EXCEL_PATH in the script |
| openpyxl not found | Run `pip install openpyxl` in CMD |
| Black CMD window flashes | Use `pythonw.exe` instead of `python.exe` in Task Scheduler |
