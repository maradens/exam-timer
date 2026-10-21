# Exam Timer for PowerPoint

A free countdown timer for exams that shows students the time left on a slide. There is no trial or time limit.

It comes in two versions. Pick the one that suits you:

| | **A. Add-in** (no macro) | **B. Macro** (VBA) |
|---|---|---|
| Save as | normal `.pptx` | `.pptm` (macro-enabled) |
| Works in | every presentation, after one-time setup | only the file you imported it into |
| Internet | needed to insert/open the timer | not needed |
| "Enable Content" warning | no | yes |
| Windows / Mac / PowerPoint on the web | yes / yes / yes | Windows (Mac mostly) / no |
| Folder | [`addin/`](addin) | [`macro/`](macro) |

## Features (both versions)

- Large countdown digits: `MM:SS`, or `H:MM:SS` for exams of an hour or more
- Start / Pause / Resume, Reset, and **+1 / +5 / −1 / −5 minutes** while running
- Turns **amber** at 10 minutes left and **red** at 5 minutes left
- Shows **"TIME'S UP – please stop writing"** and flashes at zero
- **"Ends at 11:30 AM"** line and a progress bar
- Follows the real clock, so it stays accurate when you change slides

The add-in also has:

- Presets (15 min to 3 h)
- An exam title, such as "Mathematics – Paper 1"
- A chime when time is up
- A dark background option
- Buttons that hide while the timer is running (move the mouse to show them)
- Keyboard shortcuts: **Space** start/pause, **R** reset, **↑/↓** ±1 min

---

## A. Add-in (no macro)

The timer page is hosted on GitHub Pages at
`https://maradens.github.io/exam-timer/addin/index.html`.
PowerPoint loads it using [`addin/manifest.xml`](addin/manifest.xml).

### One-time setup on each computer

**Option 1: Upload My Add-in** (try this first)

1. Download [`addin/manifest.xml`](addin/manifest.xml) to your computer.
2. In PowerPoint: **Home → Add-ins → Advanced…** (or **Insert → Get Add-ins**).
3. Open the **MY ADD-INS** tab, then **Upload My Add-in** → choose `manifest.xml` → **Upload**.

**Option 2: Shared folder** (if Upload My Add-in is missing or blocked)

1. Make a folder, e.g. `C:\ExamTimerAddin`, and put `manifest.xml` inside.
2. Right-click the folder → **Properties → Sharing → Share…** → **Share**.
   Copy the network path it shows, e.g. `\\MY-PC\ExamTimerAddin`.
3. In PowerPoint: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
   Paste the network path into **Catalog Url** → **Add catalog** → tick **Show in Menu** → **OK**.
4. Restart PowerPoint.
5. **Home → Add-ins → Advanced… → SHARED FOLDER** → **Exam Timer** → **Add**.

### Using it

1. Go to a slide and choose **Home → Add-ins → My Add-ins → Exam Timer**.
2. Resize the timer box to fill the slide.
3. Click **⚙** to set the exam length, title and options, then **Save**.
4. Start the slide show and click **Start**.

Each timer remembers its own settings when you save the presentation.
School or work accounts may block custom add-ins. If neither option works, ask IT or use the macro version.

---

## B. Macro (VBA)

Full steps are in [`macro/HOW-TO-INSTALL.txt`](macro/HOW-TO-INSTALL.txt). In short:

1. Open your presentation. Go to **Developer → Visual Basic** (or press Alt+F11), then **File → Import File… → `macro/ExamTimer.bas`**.
2. Go to a blank slide and choose **View → Macros → InsertExamTimer → Run**. Type the minutes.
3. Save as **PowerPoint Macro-Enabled Presentation (.pptm)**. If PowerPoint asks about saving as macro-free, click **No**.
4. In the slide show, click **Start**.

To remove the timer: **View → Macros → RemoveExamTimer → Run**.
When you reopen the file, click **Enable Content** if PowerPoint asks.
To reuse it, keep one "Exam Timer.pptm" and copy it for each exam.

---

## Project layout

```
addin/
  index.html      timer page (HTML/CSS/JS + Office.js), served by GitHub Pages
  manifest.xml    tells PowerPoint where the page is
  assets/         icons
macro/
  ExamTimer.bas   VBA module
  HOW-TO-INSTALL.txt
```

### Changing the add-in

Edit `addin/index.html`, then commit and push. GitHub Pages updates within a minute or two, and PowerPoint picks up the change the next time the timer loads.
If you fork this project, replace every `maradens.github.io` URL in `manifest.xml` with your own GitHub Pages address.
