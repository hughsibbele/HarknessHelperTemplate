# Harkness Helper

**Automates Harkness discussion workflows.** Upload a recording, get a transcript with speaker identification, AI-generated feedback, and distribute results via email or Canvas.

Built on Google Sheets + Google Apps Script. Free to run (you provide your own API keys).

---

## What You Need Before Starting

You'll need two API keys (takes ~5 minutes to get both):

1. **ElevenLabs API key** (for transcription) — sign up at [elevenlabs.io](https://elevenlabs.io)
   - Go to Profile (bottom-left) > API Keys
   - Transcription costs ~$0.37/hour of audio

2. **Google Gemini API key** (for speaker ID and feedback) — get one free at [aistudio.google.com](https://aistudio.google.com)
   - The free tier is sufficient for typical classroom use

---

## Setup Guide (5 minutes)

### Step 1: Copy the Spreadsheet

1. Click this link: **[Make a copy of Harkness Helper](https://docs.google.com/spreadsheets/d/1eO3PXxVS_e9TF6e83tVshmbCcWrBqN-nMwAr1WdKRQA/copy)**
2. Click **Make a copy** when prompted

You now have your own spreadsheet with 8 tabs: Settings, Discussions, Students, Transcripts, SpeakerMap, StudentReports, Prompts, and Courses.

### Step 2: Add the Code

1. In your spreadsheet, go to **Extensions > Apps Script**
2. This opens the script editor in a new tab
3. Delete any code already in the editor
4. For each `.gs` file in the `src/` folder of this repository (`Code.gs`, `Config.gs`, `Sheets.gs`, `Prompts.gs`, `ElevenLabs.gs`, `Gemini.gs`, `Canvas.gs`, `DriveMonitor.gs`, `Email.gs`, `Webapp.gs`):
   > There are 10 script files total, plus 1 HTML file.
   - Click the **+** next to "Files" in the left sidebar, choose **Script**, and name it to match (e.g., `Config`)
   - Copy the file contents and paste them in
   - For `RecorderApp.html`: click **+** > **HTML**, name it `RecorderApp`, and paste the contents
5. In the script editor, click the gear icon (**Project Settings**) on the left sidebar
   - Check **"Show 'appsscript.json' manifest file in editor"**
   - Go back to the Editor, open `appsscript.json`, and replace its contents with the `appsscript.json` from this repository
6. Click **Save** (or Ctrl+S)

### Step 3: Run the Setup Wizard

1. Go back to your **spreadsheet** tab (not the script editor)
2. Refresh the page — wait a few seconds for the menu to appear
3. Click **Harkness Helper > Setup Wizard (start here)**
4. Google will ask you to authorize the script — click through the permissions prompts
   - You may see a "This app isn't verified" warning. Click **Advanced > Go to [project name]** to continue. This is your own script running on your own account — it's safe.
5. Enter your ElevenLabs API key and Gemini API key
6. Click **Run Setup**

The wizard automatically creates your Drive folders (`Harkness Helper / Upload` and `Harkness Helper / Processing`) and initializes all sheets.

---

## How to Use

### 1. Upload Audio

Upload your discussion recording (mp3, m4a, wav, etc.) to the **Harkness Helper / Upload** folder in your Google Drive.

**Tip:** Name files like `Section 3 - 2025-01-15.m4a` and the section and date will be auto-detected.

### 2. Start Processing

Click **Harkness Helper > Start Processing**. This:
- Immediately checks for new files
- Installs a 10-minute background trigger to keep checking
- Automatically stops after 60 minutes (no runaway triggers)

### 3. Generate Feedback

After transcription completes (check the Discussions sheet for status):
1. Enter a grade in the **grade** column on the Discussions sheet
2. Click **Harkness Helper > Generate Feedback**
3. Gemini generates feedback based on the transcript and grade

### 4. Review and Approve

1. Read the generated feedback in the **group_feedback** column
2. Edit if needed
3. Check the **approved** checkbox

### 5. Send

Click **Harkness Helper > Send Approved Feedback** to distribute via email and/or Canvas.

---

## Customization

### Settings Sheet

Key settings you can change:

| Setting | Default | Description |
|---------|---------|-------------|
| `mode` | `group` | `group` = one grade for the class; `individual` = per-student grades |
| `distribute_email` | `true` | Send feedback via email |
| `distribute_canvas` | `false` | Post grades to Canvas (configure first) |
| `grade_scale` | `0-100` | Grade scale description |
| `teacher_email` | | Your email for reply-to |
| `email_subject_template` | `Harkness Discussion Report - {date}` | Email subject line |

### Prompts Sheet

All AI prompts are editable in the Prompts sheet. You can customize:
- **SPEAKER_IDENTIFICATION** — How Gemini identifies speakers from introductions
- **GROUP_FEEDBACK** — The group feedback format and tone
- **INDIVIDUAL_FEEDBACK** — Per-student feedback format

### Canvas Integration

Click **Harkness Helper > Configure Canvas Course** to set up Canvas grade posting:
1. Enter your Canvas base URL
2. Enter a Canvas API token
3. Enter the course ID

Rosters are synced automatically. To switch courses or re-sync, run the dialog again.

### Multi-Course Support

If you teach multiple courses that use Harkness discussions, you can track them all in one spreadsheet:

1. Click **Harkness Helper > Enable Multi-Course** (or add rows to the **Courses** sheet manually)
2. Add each course's name, Canvas course ID, and Canvas base URL
3. When uploading audio, prefix the filename with the course name: `AP Lit - Section 2 - 2025-01-15.m4a`
4. Use **Sync Canvas Roster** to sync students from all configured courses at once
5. Each discussion row has a `course` column to track which course it belongs to

### Mobile Recorder (Optional)

Record discussions directly from your phone using the built-in recorder app. The recorder is hosted on **GitHub Pages** (not served from Apps Script) because Apps Script's sandboxed iframe blocks microphone access.

**Backend setup:**
1. In the script editor, click **Deploy > New deployment**
2. Choose **Web app**, set "Execute as" to **Me**, and "Who has access" to **Anyone** (not "Anyone with a Google account" — the GitHub Pages frontend needs unauthenticated API access)
3. Click **Deploy** and copy the `/exec` URL

**Frontend setup:**
4. In `src/RecorderApp.html`, replace `DEPLOYMENT_ID` in the `APPS_SCRIPT_URL` variable with the deployment ID from the URL above
5. Host the repository on GitHub Pages (Settings > Pages > Source: main branch, `/` root) or serve `src/RecorderApp.html` from any static host
6. Bookmark the GitHub Pages URL on your phone or add it to your home screen

The recorder lets you pick a section, date, and (in multi-course mode) course, then uploads the recording directly to your Upload folder.

---

## Costs

- **Transcription**: ~$0.37/hour via ElevenLabs Scribe v2
- **AI Feedback**: Free via Gemini's free tier (gemini-2.0-flash)
- **Everything else**: Free (Google Sheets, Apps Script, Gmail)

---

## Troubleshooting

**"Setup Wizard" doesn't appear in the menu**
- Refresh the spreadsheet page and wait 5-10 seconds. The menu loads when the sheet opens.

**Processing doesn't find my files**
- Make sure files are in the `Harkness Helper / Upload` folder (not Processing or Completed)
- Check that the file format is supported (mp3, m4a, wav, ogg, webm, aac, flac, mp4, mov)

**Transcription shows "error" status**
- Check the `error_message` column on the Discussions sheet
- Very large files (>1 hour) may time out — try splitting the audio

**Canvas roster sync failed**
- Double-check your Canvas API token, base URL, and course ID
- Make sure your token has permission to access the course
